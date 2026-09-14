# -*- coding: utf-8 -*-
"""
Los dos procesos completos, de punta a punta.
"""

import pandas as pd

from .externos import desempeno_fd, ofertas_adj
from .diccionarios import (
    construir_dic_cmg, construir_dic_resumen_capacidad,
    construir_dic_resumen_eficiencia, construir_dic_resumen_factor,
    construir_mapa_barra,
)
from .ecostos import (
    completar_calculo_e_costos_grupos, construir_calculo_e_costos,
)
from .ecostos_prorratas import construir_dic_mapeo_diccionario
from .escritura import escribir_pagos_bess, escribir_salida
from .estructura import SECCIONES_CONSOLIDADO, SECCIONES_PAGOS
from .fma import (
    TITULO_BLOQUE_FD, TITULO_BLOQUE_FMA_CPF, cargar_tablas_fma,
    construir_dic_bloque_diccionario,
)
from .hojas_entrada import construir_fd, construir_subastas, leer_cmg
from .lectura import (
    construir_homologacion, leer_centrales, leer_medidas_sae,
)
from .medidores import construir_medidores
from .ofertas_sscc import construir_resumen_ventana_oferta
from .parametros import (
    ARCHIVO_CENTRALES, ARCHIVO_CMG, ARCHIVO_MEDIDAS_SAE,
    CARPETA_DB_SUBASTAS, CARPETA_FD_FMA, HOJA_DICCIONARIO,
)
from .re545 import (
    completar_calculo_re545, completar_checks_resumen_re545,
    construir_calculo_re545, renombrar_calculo_re545,
)
from .re545_componentes import calcular_componentes_re545
from .re545_resumen import construir_resumen_ventanas_re545
from .rutas import (
    buscar_archivo_ofertas, buscar_archivo_sscc_desempeno,
    buscar_archivo_subastas, buscar_soc, periodo_desde_aamm,
    resolver_rutas, validar_aamm,
)
from .soc import extraer_soc
from .subastas_accdb import (
    construir_mapa_propietario, construir_subastas_desde_accdb,
)
from .utiles import ErrorEntrada


def generar_consolidado(
    carpeta_base, aamm, secciones_activas, registrar=print, progreso=None
):
    """
    Genera/actualiza Consolidado_entradas.xlsx, recalculando solo las
    hojas de las secciones pedidas (ids de SECCIONES_CONSOLIDADO) y
    preservando el resto tal cual estaba en el archivo existente (ver
    escribir_salida). Si el archivo no existe, se crea. La usan los
    botones "Actualizar" de las filas-hoja del diagrama (y el
    "Actualizar todo" de la fila del archivo, que manda todas).

    secciones_activas: iterable de ids de SECCIONES_CONSOLIDADO
    ("medidores", "ofertas_sscc", "cmg", "fd", "subastas") a
    recalcular esta vez. "medidores" y "ofertas_sscc" comparten una
    unica lectura/calculo (construir_medidores() arma las dos hojas
    de una, porque Medidores!R:S:T depende de Ofertas SSCC) --
    actualizar cualquiera de las dos dispara esa lectura; lo que cada
    id decide por separado es solo que hoja se reescribe.
    """

    def avanzar(valor):
        if progreso:
            progreso(valor)

    secciones_activas = set(secciones_activas)
    ids_validos = {seccion[0] for seccion in SECCIONES_CONSOLIDADO}
    desconocidas = secciones_activas - ids_validos

    if desconocidas:
        raise ErrorEntrada(
            f"Seccion(es) desconocida(s): {sorted(desconocidas)}"
        )

    if not secciones_activas:
        raise ErrorEntrada(
            "No se tildo ninguna entrada para generar/actualizar."
        )

    hojas_regenerar = set()
    for id_seccion, _, _, hojas in SECCIONES_CONSOLIDADO:
        if id_seccion in secciones_activas:
            hojas_regenerar.update(hojas)

    rutas = resolver_rutas(carpeta_base)

    df_medidores = None
    avisos, incidencias = [], []
    df_wxy = df_resumen_ventana = None
    df_cmg = df_fd_csf = df_fd_cpf = df_subastas = None

    avanzar(5)

    if "medidores" in secciones_activas or "ofertas_sscc" in secciones_activas:

        aamm_val = validar_aamm(aamm)

        if not rutas["medidas_sae"].is_file():
            raise ErrorEntrada(
                f"No se encontro {rutas['medidas_sae']}"
            )

        if not rutas["centrales"].is_file():
            raise ErrorEntrada(
                f"No se encontro {rutas['centrales']}"
            )

        archivo_ofertas = buscar_archivo_ofertas(rutas["ofertas_dir"])
        if not archivo_ofertas:
            raise ErrorEntrada(
                f"No se encontro ningun archivo *OfertasSSCC* en "
                f"{rutas['ofertas_dir']}"
            )

        archivo_soc = buscar_soc(rutas["medidas_dir"], aamm_val)
        anio, mes = periodo_desde_aamm(aamm_val)

        registrar(f"Periodo indicado: {anio}-{mes:02d} ({aamm_val})")

        registrar("Leyendo Centrales.xlsx...")
        _, diccionario = leer_centrales(rutas["centrales"])
        mapa = construir_homologacion(diccionario)
        registrar(f"  homologaciones cargadas: {len(mapa):,}")
        avanzar(20)

        registrar(f"Leyendo {archivo_soc.name}...")
        df_soc, incidencias = extraer_soc(archivo_soc, mapa)
        registrar(
            f"  bloques leidos: "
            f"{df_soc['central'].nunique()}   "
            f"registros: {len(df_soc):,}"
        )

        for incidencia in incidencias:
            registrar(f"  [SOC] {incidencia}")

        avanzar(35)

        registrar(f"Leyendo {ARCHIVO_MEDIDAS_SAE}...")
        df_sae = leer_medidas_sae(rutas["medidas_sae"])
        registrar(f"  filas: {len(df_sae):,}")
        avanzar(50)

        registrar("Construyendo Medidores...")
        (
            df_medidores,
            avisos,
            df_wxy,
            df_resumen_ventana,
        ) = construir_medidores(
            df_sae,
            df_soc,
            anio,
            mes,
            archivo_ofertas,
            diccionario,
            registrar=registrar,
        )

        for aviso in avisos:
            registrar(f"  [AVISO] {aviso}")

    avanzar(60)

    if "cmg" in secciones_activas:
        if not rutas["cmg"].is_file():
            raise ErrorEntrada(f"No se encontro {rutas['cmg']}")
        registrar(f"Leyendo {ARCHIVO_CMG}...")
        df_cmg = leer_cmg(rutas["cmg"], registrar=registrar)

    avanzar(72)

    if "fd" in secciones_activas:
        archivo_sscc = buscar_archivo_sscc_desempeno(
            rutas["sscc_desempeno_dir"]
        )
        if not archivo_sscc:
            raise ErrorEntrada(
                f"No se encontro ningun archivo SSCC_Desempeño_* en "
                f"{rutas['sscc_desempeno_dir']}"
            )
        registrar(f"Leyendo {archivo_sscc.name}...")
        df_fd_csf, df_fd_cpf = construir_fd(
            archivo_sscc, registrar=registrar
        )

    avanzar(84)

    if "subastas" in secciones_activas:

        aamm_val = validar_aamm(aamm)
        hay_accdb = bool(
            ofertas_adj.accdb_presentes(rutas["subastas_dir"], aamm_val)
        )

        if hay_accdb:
            # Camino normal desde ahora: el origen real de las
            # subastas son los Access, no la planilla 3.
            mapa_propietarios = {}
            dic_cpf = {}
            diccionario_centrales = None

            if rutas["centrales"].is_file():
                resumen_bess, diccionario_centrales = leer_centrales(
                    rutas["centrales"]
                )
                mapa_propietarios = construir_mapa_propietario(resumen_bess)
                dic_cpf = construir_dic_bloque_diccionario(
                    diccionario_centrales, TITULO_BLOQUE_FMA_CPF
                )
            else:
                registrar(
                    f"  [AVISO] no se encontro {rutas['centrales']}: la "
                    f"columna Propietario de Subastas queda vacia."
                )

            registrar(
                f"Leyendo las salidas de FMA de {CARPETA_FD_FMA}/..."
            )
            tablas_fma = cargar_tablas_fma(
                rutas["sscc_desempeno_dir"], aamm_val, registrar=registrar
            )

            # El FD (Subastas!P) y el Vector de Participacion CSF que
            # necesita el FMA salen los dos del archivo SSCC_Desempeño_*.
            tablas_fd = None
            dic_unidad_fd = {}

            archivo_sscc_fd = buscar_archivo_sscc_desempeno(
                rutas["sscc_desempeno_dir"]
            )

            if archivo_sscc_fd:
                registrar(f"Leyendo el FD de {archivo_sscc_fd.name}...")
                try:
                    tablas_fd = desempeno_fd.construir_tablas_fd(
                        archivo_sscc_fd, registrar=registrar
                    )
                except desempeno_fd.ErrorDesempeno as error:
                    raise ErrorEntrada(str(error)) from error

                if rutas["centrales"].is_file():
                    dic_unidad_fd = construir_dic_bloque_diccionario(
                        diccionario_centrales, TITULO_BLOQUE_FD
                    ) or construir_dic_mapeo_diccionario(diccionario_centrales)
            else:
                registrar(
                    f"  [AVISO] no hay ningun SSCC_Desempeño_* en "
                    f"{CARPETA_FD_FMA}/ (se baja con el boton 'Traer "
                    f"FD'): las columnas FD y el Vector de Participacion "
                    f"CSF quedan sin dato."
                )

            if not dic_cpf:
                registrar(
                    f"  [AVISO] {ARCHIVO_CENTRALES} no tiene el bloque "
                    f"'{TITULO_BLOQUE_FMA_CPF}' en la hoja "
                    f"{HOJA_DICCIONARIO} (Configuración -> central como "
                    f"la nombra fma_cpf): el FMA de CPF se busca con el "
                    f"nombre tal cual."
                )

            registrar(
                f"Leyendo las subastas de {CARPETA_DB_SUBASTAS}/ "
                f"(periodo {aamm_val})..."
            )
            df_subastas = construir_subastas_desde_accdb(
                rutas["subastas_dir"],
                aamm_val,
                mapa_propietarios=mapa_propietarios,
                tablas_fma=tablas_fma,
                dic_cpf=dic_cpf,
                tablas_fd=tablas_fd,
                dic_unidad_fd=dic_unidad_fd,
                registrar=registrar,
            )

        else:
            # Respaldo: los casos que todavia no tienen los Access
            # copiados siguen andando con la planilla 3 de siempre.
            archivo_subastas = buscar_archivo_subastas(rutas["subastas_dir"])

            if not archivo_subastas:
                raise ErrorEntrada(
                    f"No hay de donde sacar las subastas del periodo "
                    f"{aamm_val}:\n\n"
                    f"  - {rutas['db_subastas_dir']} no tiene ningun "
                    f"Access del periodo (traelos con el boton 'Traer "
                    f"subastas'), y\n"
                    f"  - tampoco hay un archivo "
                    f"3_REMUNERACIÓN_SUBASTAS_E_ID_* de respaldo en "
                    f"{rutas['subastas_dir']}."
                )

            registrar(
                f"  [AVISO] no hay Access del periodo en "
                f"{CARPETA_DB_SUBASTAS}/: se usa el respaldo "
                f"{archivo_subastas.name} (planilla 3)."
            )
            df_subastas = construir_subastas(
                archivo_subastas, registrar=registrar
            )

    avanzar(92)

    registrar(f"Escribiendo {rutas['salida'].name}...")
    escribir_salida(
        df_medidores,
        rutas["salida"],
        avisos,
        incidencias,
        df_wxy,
        df_resumen_ventana,
        df_cmg,
        df_fd_csf,
        df_fd_cpf,
        df_subastas,
        ruta_existente=rutas["salida"],
        hojas_regenerar=hojas_regenerar,
        registrar=registrar,
    )

    avanzar(100)
    registrar(f"Listo: {rutas['salida']}")

    return rutas["salida"]


def generar_pagos_bess(
    carpeta_base, secciones_activas, registrar=print, progreso=None
):
    """
    Genera/actualiza Pagos_BESS.xlsx, recalculando solo las hojas de
    las secciones pedidas (ids de SECCIONES_PAGOS: "ecostos",
    "re545") y preservando el resto tal cual estaba en el archivo
    existente (ver escribir_pagos_bess/hojas_regenerar) -- mismo
    criterio que generar_consolidado()/SECCIONES_CONSOLIDADO. Si el
    archivo no existe, se crea. La usan los botones "Actualizar" de
    las filas-hoja del diagrama.

    No recalcula Medidores ni Subastas: los lee tal cual estan en
    Consolidado_entradas.xlsx, que debe generarse primero con su
    propio boton "Actualizar". Centrales.xlsx y cmg.xlsx si se leen/
    recalculan frescos. El archivo SSCC_Desempeño_* solo se exige si
    se pide "ecostos" -- "re545" no usa FD.
    """

    def avanzar(valor):
        if progreso:
            progreso(valor)

    secciones_activas = set(secciones_activas)
    ids_validos = {seccion[0] for seccion in SECCIONES_PAGOS}
    desconocidas = secciones_activas - ids_validos

    if desconocidas:
        raise ErrorEntrada(
            f"Seccion(es) desconocida(s): {sorted(desconocidas)}"
        )

    if not secciones_activas:
        raise ErrorEntrada(
            "No se tildo ninguna seccion para generar/actualizar."
        )

    quiere_ecostos = "ecostos" in secciones_activas
    quiere_re545 = "re545" in secciones_activas

    hojas_regenerar = set()
    for id_seccion, _, _, hojas in SECCIONES_PAGOS:
        if id_seccion in secciones_activas:
            hojas_regenerar.update(hojas)

    rutas = resolver_rutas(carpeta_base)

    if not rutas["salida"].is_file():
        raise ErrorEntrada(
            f"No se encontro {rutas['salida']}. Primero hay que "
            f"generar Consolidado_entradas.xlsx (boton 'Generar' de "
            f"esa fila)."
        )

    registrar(f"Leyendo hoja 'Medidores' de {rutas['salida'].name}...")

    try:
        df_medidores = pd.read_excel(rutas["salida"], sheet_name="Medidores")
    except ValueError as error:
        raise ErrorEntrada(
            f"{rutas['salida'].name} no tiene la hoja 'Medidores' "
            f"todavia. Genera Consolidado_entradas.xlsx primero "
            f"(tildando 'Medidores + Ofertas SSCC')."
        ) from error

    if df_medidores.empty:
        raise ErrorEntrada(
            f"La hoja 'Medidores' de {rutas['salida'].name} esta "
            f"vacia. Genera Consolidado_entradas.xlsx primero "
            f"(tildando 'Medidores + Ofertas SSCC')."
        )

    registrar(f"  filas: {len(df_medidores):,}")
    avanzar(10)

    if not rutas["centrales"].is_file():
        raise ErrorEntrada(f"No se encontro {rutas['centrales']}")

    registrar("Leyendo Centrales.xlsx...")
    resumen, diccionario = leer_centrales(rutas["centrales"])
    mapa_barra = construir_mapa_barra(resumen)
    dic_factor, umbral_soc_minimo = construir_dic_resumen_factor(resumen)

    # Eficiencia y Capacidad solo las usa RE545 (V y U/BC/BN).
    dic_eficiencia = dic_capacidad = None
    if quiere_re545:
        dic_eficiencia = construir_dic_resumen_eficiencia(resumen)
        dic_capacidad = construir_dic_resumen_capacidad(resumen)

    avanzar(20)

    if not rutas["cmg"].is_file():
        raise ErrorEntrada(f"No se encontro {rutas['cmg']}")

    registrar(f"Leyendo {ARCHIVO_CMG}...")
    df_cmg = leer_cmg(rutas["cmg"], registrar=registrar)
    dic_cmg = construir_dic_cmg(df_cmg)
    avanzar(30)

    registrar(f"Leyendo hoja 'Subastas' de {rutas['salida'].name}...")

    try:
        df_subastas = pd.read_excel(rutas["salida"], sheet_name="Subastas")
    except ValueError as error:
        raise ErrorEntrada(
            f"{rutas['salida'].name} no tiene la hoja 'Subastas' "
            f"todavia. Genera Consolidado_entradas.xlsx primero "
            f"(tildando 'Subastas')."
        ) from error

    if df_subastas.empty:
        raise ErrorEntrada(
            f"La hoja 'Subastas' de {rutas['salida'].name} esta "
            f"vacia. Genera Consolidado_entradas.xlsx primero "
            f"(tildando 'Subastas')."
        )

    avanzar(40)

    df_ecostos = None
    df_re545 = None
    df_resumen_re545 = None

    if quiere_ecostos:

        registrar(
            "Construyendo Calculo E Costos (etapa base: H + CMg + "
            "traspaso de Medidores)..."
        )
        df_ecostos = construir_calculo_e_costos(
            df_medidores, mapa_barra, dic_cmg, registrar=registrar
        )
        avanzar(50)

        archivo_sscc = buscar_archivo_sscc_desempeno(
            rutas["sscc_desempeno_dir"]
        )
        if not archivo_sscc:
            raise ErrorEntrada(
                f"No se encontro ningun archivo SSCC_Desempeño_* en "
                f"{rutas['sscc_desempeno_dir']} (hace falta para AM:AR "
                f"de Calculo E Costos)."
            )

        registrar(f"Leyendo {archivo_sscc.name}...")
        df_fd_csf, df_fd_cpf = construir_fd(archivo_sscc, registrar=registrar)
        avanzar(60)

        registrar(
            "Completando L, M, N, O, R, S, T, U, W, X, Y, AB, AC, AD, AE, "
            "AF, AG, AH, AI, AJ, AK, AL, AM, AN, AO, AP, AQ, AR, AS, AT, "
            "AU, AV, AW, AX, AZ..."
        )
        df_ecostos = completar_calculo_e_costos_grupos(
            df_ecostos, df_subastas, dic_factor, umbral_soc_minimo,
            diccionario, df_fd_csf, df_fd_cpf,
            registrar=registrar,
        )

    avanzar(70)

    if quiere_re545:

        registrar(
            "Construyendo Calculo RE545 (etapa base: traspaso de "
            "Medidores + L, M, N, O, S, U, V)..."
        )
        df_re545 = construir_calculo_re545(
            df_medidores, mapa_barra, dic_cmg, registrar=registrar
        )
        df_re545_base = completar_calculo_re545(
            df_re545, df_subastas, umbral_soc_minimo, dic_capacidad,
            dic_eficiencia, registrar=registrar,
        )

        # AY ("Oferta Completa") del resumen por central+ventana sale
        # de la misma tabla que ya alimenta Medidores!T, reconstruida
        # aca a partir de la hoja Medidores ya generada.
        resumen_ventana_oferta = construir_resumen_ventana_oferta(
            df_medidores["clave"],
            df_medidores["Ventana"],
            df_medidores["Oferta_Completa_Dia"],
            registrar=registrar,
        )
        df_resumen_re545 = construir_resumen_ventanas_re545(
            df_re545_base, resumen_ventana_oferta, dic_capacidad,
            registrar=registrar,
        )

        registrar("  Calculo RE545: Componente 1 y Componente 2 (BI:CE)...")
        for interno, serie in calcular_componentes_re545(
            df_re545_base, df_resumen_re545, dic_factor,
            registrar=registrar,
        ).items():
            df_re545_base[interno] = serie

        df_resumen_re545 = completar_checks_resumen_re545(
            df_resumen_re545, df_re545_base
        )

        df_re545 = renombrar_calculo_re545(df_re545_base)

    avanzar(90)

    registrar(f"Escribiendo {rutas['salida_pagos'].name}...")
    escribir_pagos_bess(
        rutas["salida_pagos"],
        df_ecostos,
        df_re545,
        df_resumen_re545,
        ruta_existente=rutas["salida_pagos"],
        hojas_regenerar=hojas_regenerar,
        registrar=registrar,
    )

    avanzar(100)
    registrar(f"Listo: {rutas['salida_pagos']}")

    return rutas["salida_pagos"]
