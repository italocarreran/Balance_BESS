# -*- coding: utf-8 -*-
"""
Nucleo de calculo de la etapa Medidores.

Sin dependencias de interfaz: todo lo de aca se puede correr y
testear sin abrir la ventana.
"""

import calendar
import math
import re
import unicodedata
from pathlib import Path

import openpyxl
import pandas as pd


# ============================================================
# PARAMETROS FIJOS
# ============================================================

# Equivale a Medidores!S1 en la planilla 11.
INICIO_VENTANA = 10

# Umbral de la columna O. La planilla usa 6%.
UMBRAL_SOC = 0.06


# ============================================================
# NOMBRES DE ARCHIVO Y CARPETA
# ============================================================

CARPETA_MEDIDAS = "Medidas"
CARPETA_AUXILIARES = "Auxiliares"
CARPETA_OFERTAS = "Ofertas"
CARPETA_CMG = "Cmg"
CARPETA_SSCC_DESEMPENO = "SSCC_Desempeño"
CARPETA_SUBASTAS = "Subastas"

ARCHIVO_MEDIDAS_SAE = "Medidas_SAE.xlsx"
HOJA_MEDIDAS_SAE = "Medidas"

ARCHIVO_CENTRALES = "Centrales.xlsx"
HOJA_RESUMEN_BESS = "Resumen BESS"
HOJA_DICCIONARIO = "Diccionario"

# Nombre literal y fijo (a diferencia de SoC/Ofertas/SSCC_Desempeño/
# Subastas): asi lo exige Cargar_CMg_Desde_Archivo.
ARCHIVO_CMG = "cmg.xlsx"
HOJA_CMG_ORIGEN = "CMg"

HOJA_CPF_HORARIO = "CPF Horario"
HOJA_CSF_HORARIO = "CSF Horario"

HOJA_SUBASTAS_ORIGEN = "DB"

ARCHIVO_SALIDA = "Consolidado_entradas.xlsx"

# Etapa siguiente (Calculo E Costos / "Ecostos"): el usuario pidio que
# viva en una planilla aparte de Consolidado_entradas.xlsx. Nombre
# provisorio, puede cambiar.
ARCHIVO_SALIDA_PAGOS = "Pagos_BESS.xlsx"
HOJA_CALCULO_ECOSTOS = "Calculo E Costos"
HOJA_CALCULO_RE545 = "Calculo RE545"

# El periodo AAMM (ej. "2607") ya no se infiere del nombre del archivo:
# lo ingresa el usuario en la ventana. El archivo de SoC solo debe
# contener "SOC" y el AAMM en su nombre (plan, seccion 19.1) - no existe
# un nombre de archivo literal fijo.
PATRON_AAMM = re.compile(r"^\d{4}$")

# Extensiones de Excel aceptadas para los archivos que se buscan por
# patron de nombre (OfertasSSCC, SSCC_Desempeño_*, 3_REMUNERACIÓN_
# SUBASTAS_E_ID_*) - no para SoC (siempre .xlsx) ni para cmg.xlsx
# (nombre literal fijo).
EXTENSIONES_EXCEL = {".xlsx", ".xlsm", ".xlsb", ".xls"}

# Se derivan con .lower() en vez de transcribir el literal a mano: con
# letras dobles/triples seguidas ("Ofertas"+"SSCC") es facil perder una
# al tipear (ya paso una vez, ver METODOLOGIA.md #7).
PATRON_NOMBRE_OFERTAS = "OfertasSSCC".lower()
PATRON_NOMBRE_SSCC_DESEMPENO = "SSCC_Desempeño_".lower()
PATRON_NOMBRE_SUBASTAS = "3_REMUNERACIÓN_SUBASTAS_E_ID_".lower()


# ============================================================
# MAPEO DE COLUMNAS A:I
# ============================================================

# Orden de Medidas_SAE.xlsx, que alimenta Medidores!A:I.
COLUMNAS_AI = [
    "Mes",
    "Dia",
    "Hora",
    "Minutos",
    "Hora Mes",
    "Cuarto de Hora",
    "clave",
    "intervalo",
    "Gen_Unidad",
]

# Nombre logico de cada letra de Excel, para poder comparar contra la
# hoja original columna por columna. El orden de insercion de este
# dict ES el orden final de columnas de Medidores.
#
# V, W, X, Y, AB, AC, AD, AE NO estan aca: en la planilla original no
# son una columna por fila de Medidores, son tablas auxiliares de otro
# largo (una fila por central x dia, o por central x ventana) que solo
# viven en esas letras de columna porque ahi habia espacio libre. En
# Python se escriben como hoja propia de Consolidado_entradas.xlsx en
# vez de forzarlas a columnas del mismo largo que A:U (ver ejecutar()).
LETRA_A_CAMPO = {
    "A": "Mes",
    "B": "Dia",
    "C": "Hora",
    "D": "Minutos",
    "E": "Hora Mes",
    "F": "Cuarto de Hora",
    "G": "clave",
    "H": "intervalo",
    "I": "Gen_Unidad",
    "J": "SoC",
    "K": "Copia_Ventana",
    "L": "Ventana",
    "M": "M_VACIA",
    "N": "Clave_Dia_HoraMes",
    "O": "Indicador_SoC",
    "P": "P_VACIA",
    "Q": "Q_VACIA",
    "R": "Oferta_Completa_Dia",
    "S": "Indicador_Ventana_Oferta",
    "T": "Ventana_No_Completa",
    "U": "U_VACIA",
}

# Columnas que el plan define como deliberadamente vacias (plan
# seccion 16.3): no son trabajo pendiente, es el diseño confirmado.
COLUMNAS_VACIAS = [
    "M_VACIA",
    "P_VACIA",
    "Q_VACIA",
    "U_VACIA",
]


# ============================================================
# UTILIDADES
# ============================================================

def normalizar(texto):
    """Minusculas, sin tildes, espacios colapsados."""

    if texto is None:
        return ""

    # Una celda vacia llega como NaN y str(NaN) es 'nan', que
    # se confundiria con un nombre real.
    try:
        if pd.isna(texto):
            return ""
    except (TypeError, ValueError):
        pass

    nfkd = unicodedata.normalize("NFKD", str(texto))

    sin_tildes = "".join(
        c for c in nfkd
        if not unicodedata.combining(c)
    )

    return re.sub(r"\s+", " ", sin_tildes).strip().lower()


class ErrorEntrada(Exception):
    """Problema en las entradas que impide continuar."""


# ============================================================
# RESOLUCION DE RUTAS
# ============================================================

def resolver_rutas(carpeta_base):
    """
    Deriva todas las rutas desde la carpeta base elegida por el
    usuario. Nada depende de donde vive el .py.
    """

    base = Path(carpeta_base)

    medidas_dir = base / CARPETA_MEDIDAS
    auxiliares_dir = base / CARPETA_AUXILIARES
    ofertas_dir = base / CARPETA_OFERTAS
    cmg_dir = base / CARPETA_CMG
    sscc_desempeno_dir = base / CARPETA_SSCC_DESEMPENO
    subastas_dir = base / CARPETA_SUBASTAS

    return {
        "base": base,
        "medidas_dir": medidas_dir,
        "auxiliares_dir": auxiliares_dir,
        "ofertas_dir": ofertas_dir,
        "cmg_dir": cmg_dir,
        "sscc_desempeno_dir": sscc_desempeno_dir,
        "subastas_dir": subastas_dir,
        "medidas_sae": medidas_dir / ARCHIVO_MEDIDAS_SAE,
        "centrales": auxiliares_dir / ARCHIVO_CENTRALES,
        "cmg": cmg_dir / ARCHIVO_CMG,
        "salida": base / ARCHIVO_SALIDA,
        "salida_pagos": base / ARCHIVO_SALIDA_PAGOS,
    }


def validar_aamm(aamm):
    """Levanta ErrorEntrada si aamm no son 4 digitos (ej. '2607')."""

    if not aamm or not PATRON_AAMM.match(str(aamm).strip()):
        raise ErrorEntrada(
            "Ingresa el periodo AAMM en la ventana (4 digitos, "
            "por ejemplo 2607 para julio de 2026)."
        )
    return str(aamm).strip()


def _es_archivo_de_soc(nombre_archivo, aamm):
    """
    El nombre del archivo de SoC no sigue un patron fijo (no es
    literalmente 'SOC_AAMM.xlsx'): basta con que contenga 'SOC' y el
    AAMM del periodo, en cualquier posicion y con cualquier separador
    (plan de migracion, seccion 19.1). El archivo en si siempre es
    .xlsx (ver buscar_soc) - no confundir con otros archivos del caso
    que puedan compartir "SOC"+AAMM en el nombre sin serlo.
    """

    nombre = normalizar(Path(nombre_archivo).stem)
    return "soc" in nombre and aamm in nombre


def buscar_soc(medidas_dir, aamm):
    """
    Busca dentro de Medidas/ el archivo de SoC del periodo AAMM
    indicado por el usuario.

    Ninguno   -> error
    Uno       -> se usa
    Varios    -> error, NO se elige por fecha de modificacion
    """

    medidas_dir = Path(medidas_dir)
    aamm = validar_aamm(aamm)

    if not medidas_dir.is_dir():
        raise ErrorEntrada(
            f"No existe la carpeta {medidas_dir}"
        )

    candidatos = [
        archivo
        for archivo in medidas_dir.iterdir()
        if archivo.is_file()
        and not archivo.name.startswith("~$")
        and archivo.suffix.lower() == ".xlsx"
        and _es_archivo_de_soc(archivo.name, aamm)
    ]

    if not candidatos:
        raise ErrorEntrada(
            f"No se encontro ningun archivo de SoC del periodo {aamm} "
            f"en {medidas_dir}\n"
            f"El nombre debe contener 'SOC' y '{aamm}', por ejemplo "
            f"SOC_{aamm}.xlsx"
        )

    if len(candidatos) > 1:
        nombres = ", ".join(sorted(c.name for c in candidatos))
        raise ErrorEntrada(
            f"Hay {len(candidatos)} archivos de SoC del periodo {aamm} "
            f"en {medidas_dir}:\n"
            f"  {nombres}\n"
            f"Deja solo el del periodo que vas a procesar. "
            f"No se elige automaticamente para no tomar en "
            f"silencio el archivo equivocado."
        )

    return candidatos[0]


def _buscar_archivo_excel_mas_reciente(carpeta, patron, desde_inicio=False):
    """
    Busca en 'carpeta' el archivo Excel (EXTENSIONES_EXCEL) mas
    reciente por fecha de modificacion cuyo nombre (en minusculas)
    contenga 'patron' (o empiece con el, si desde_inicio=True).
    Ignora archivos temporales (~$). None si la carpeta no existe o no
    hay ningun candidato.

    Comun a *OfertasSSCC*, SSCC_Desempeño_* y 3_REMUNERACIÓN_SUBASTAS_
    E_ID_*: las tres macros originales, ante varios candidatos, SI
    eligen automaticamente el mas reciente (a diferencia del SoC).
    """

    carpeta = Path(carpeta)

    if not carpeta.is_dir():
        return None

    patron = patron.lower()

    def coincide(nombre):
        nombre = nombre.lower()
        return nombre.startswith(patron) if desde_inicio else patron in nombre

    candidatos = [
        archivo
        for archivo in carpeta.iterdir()
        if archivo.is_file()
        and not archivo.name.startswith("~$")
        and archivo.suffix.lower() in EXTENSIONES_EXCEL
        and coincide(archivo.stem)
    ]

    if not candidatos:
        return None

    return max(candidatos, key=lambda a: a.stat().st_mtime)


def buscar_archivo_ofertas(ofertas_dir):
    """
    Replica OSSCC_BuscarArchivoOfertas: busca dentro de Ofertas/
    cualquier archivo Excel cuyo nombre contenga "OfertasSSCC".
    """

    return _buscar_archivo_excel_mas_reciente(
        ofertas_dir, PATRON_NOMBRE_OFERTAS
    )


def buscar_archivo_sscc_desempeno(carpeta):
    """
    Replica BuscarArchivoSSCCMasReciente: busca dentro de
    SSCC_Desempeño/ el archivo Excel mas reciente cuyo nombre empiece
    con "SSCC_Desempeño_".
    """

    return _buscar_archivo_excel_mas_reciente(
        carpeta, PATRON_NOMBRE_SSCC_DESEMPENO, desde_inicio=True
    )


def buscar_archivo_subastas(carpeta):
    """
    Replica BuscarArchivoSubastasMasRecienteRapido: busca dentro de
    Subastas/ el archivo Excel mas reciente cuyo nombre empiece con
    "3_REMUNERACIÓN_SUBASTAS_E_ID_".

    Nota de arquitectura: la macro original buscaba este archivo
    directamente en la carpeta del .xlsm (sin subcarpeta). Aca se le
    da su propia carpeta (Subastas/) para ser consistente con el resto
    de las entradas externas (Medidas/, Auxiliares/, Ofertas/, Cmg/,
    SSCC_Desempeño/), cada una con su propia carpeta bajo la carpeta
    base del caso.
    """

    return _buscar_archivo_excel_mas_reciente(
        carpeta, PATRON_NOMBRE_SUBASTAS, desde_inicio=True
    )


def periodo_desde_aamm(aamm):
    """'2607' -> (2026, 7)"""

    anio = 2000 + int(aamm[:2])
    mes = int(aamm[2:])

    if not 1 <= mes <= 12:
        raise ErrorEntrada(
            f"El periodo '{aamm}' no tiene un mes valido."
        )

    return anio, mes


# ============================================================
# VALIDACION DE ESTRUCTURA
# ============================================================

def revisar_estructura(carpeta_base, aamm=None):
    """
    Revisa la carpeta base y devuelve una lista de
    (etiqueta, estado, detalle) para pintar en la ventana.

    aamm: periodo ingresado por el usuario en la ventana (4 digitos,
    ej. '2607'). Sin un AAMM valido no se puede buscar el SoC.

    estado: 'ok' | 'falta' | 'pendiente'
    """

    rutas = resolver_rutas(carpeta_base)
    filas = []

    def agregar(etiqueta, existe, detalle=""):
        filas.append(
            (
                etiqueta,
                "ok" if existe else "falta",
                detalle,
            )
        )

    agregar(
        "Carpeta base",
        rutas["base"].is_dir(),
        str(rutas["base"]),
    )

    agregar(
        f"{CARPETA_MEDIDAS}/",
        rutas["medidas_dir"].is_dir(),
    )

    agregar(
        ARCHIVO_MEDIDAS_SAE,
        rutas["medidas_sae"].is_file(),
    )

    # Periodo AAMM: lo ingresa el usuario, no se infiere de un nombre
    # de archivo (ver PATRON_AAMM / validar_aamm).
    try:
        aamm_valido = validar_aamm(aamm)
        filas.append(("Periodo (AAMM)", "ok", aamm_valido))
    except ErrorEntrada as error:
        aamm_valido = None
        filas.append(
            ("Periodo (AAMM)", "falta", str(error).split("\n")[0])
        )

    rutas["aamm"] = aamm_valido

    # SOC del periodo indicado
    if aamm_valido:
        try:
            archivo_soc = buscar_soc(rutas["medidas_dir"], aamm_valido)
            filas.append(
                (archivo_soc.name, "ok", f"periodo {aamm_valido}")
            )
            rutas["soc"] = archivo_soc
        except ErrorEntrada as error:
            filas.append(
                (
                    f"SoC periodo {aamm_valido}",
                    "falta",
                    str(error).split("\n")[0],
                )
            )
            rutas["soc"] = None
    else:
        filas.append(
            (
                "SoC del periodo",
                "falta",
                "ingresa el AAMM para poder buscarlo",
            )
        )
        rutas["soc"] = None

    agregar(
        f"{CARPETA_AUXILIARES}/",
        rutas["auxiliares_dir"].is_dir(),
    )

    agregar(
        ARCHIVO_CENTRALES,
        rutas["centrales"].is_file(),
    )

    # Hojas del maestro
    if rutas["centrales"].is_file():
        try:
            hojas = pd.ExcelFile(rutas["centrales"]).sheet_names
            hojas_norm = {normalizar(h) for h in hojas}

            for hoja in (HOJA_RESUMEN_BESS, HOJA_DICCIONARIO):
                filas.append(
                    (
                        f"  hoja '{hoja}'",
                        (
                            "ok"
                            if normalizar(hoja) in hojas_norm
                            else "falta"
                        ),
                        "",
                    )
                )
        except Exception as error:
            filas.append(
                ("  hojas de Centrales.xlsx", "falta", str(error))
            )

    # Ofertas SSCC: ubicacion definida en <CARPETA_BASE>/Ofertas/ (plan
    # seccion 16.2). Las macros de Ofertas SSCC son obligatorias (plan
    # seccion 17-18), asi que esto SI bloquea Ejecutar si falta.
    agregar(
        f"{CARPETA_OFERTAS}/",
        rutas["ofertas_dir"].is_dir(),
    )

    archivo_ofertas = buscar_archivo_ofertas(rutas["ofertas_dir"])
    rutas["ofertas"] = archivo_ofertas

    if archivo_ofertas:
        filas.append(
            (archivo_ofertas.name, "ok", f"en {CARPETA_OFERTAS}/")
        )
    else:
        filas.append(
            (
                "Archivo *OfertasSSCC*",
                "falta",
                f"ningun archivo en {CARPETA_OFERTAS}/ contiene "
                f"'OfertasSSCC' en el nombre",
            )
        )

    # CMg: nombre de archivo literal fijo (Cargar_CMg_Desde_Archivo).
    agregar(
        f"{CARPETA_CMG}/",
        rutas["cmg_dir"].is_dir(),
    )
    agregar(
        ARCHIVO_CMG,
        rutas["cmg"].is_file(),
    )

    # SSCC_Desempeño (alimenta la hoja FD): obligatorio, se toma el mas
    # reciente si hay varios (Cargar_SSCC_Desempeno_En_FD).
    agregar(
        f"{CARPETA_SSCC_DESEMPENO}/",
        rutas["sscc_desempeno_dir"].is_dir(),
    )

    archivo_sscc = buscar_archivo_sscc_desempeno(
        rutas["sscc_desempeno_dir"]
    )
    rutas["sscc_desempeno"] = archivo_sscc

    if archivo_sscc:
        filas.append(
            (archivo_sscc.name, "ok", f"en {CARPETA_SSCC_DESEMPENO}/")
        )
    else:
        filas.append(
            (
                "Archivo SSCC_Desempeño_*",
                "falta",
                f"ningun archivo en {CARPETA_SSCC_DESEMPENO}/ empieza "
                f"con 'SSCC_Desempeño_'",
            )
        )

    # Subastas (alimenta la hoja Subastas): obligatorio, se toma el
    # mas reciente si hay varios (Cargar_Remuneracion_Subastas_Rapido).
    agregar(
        f"{CARPETA_SUBASTAS}/",
        rutas["subastas_dir"].is_dir(),
    )

    archivo_subastas = buscar_archivo_subastas(rutas["subastas_dir"])
    rutas["subastas"] = archivo_subastas

    if archivo_subastas:
        filas.append(
            (archivo_subastas.name, "ok", f"en {CARPETA_SUBASTAS}/")
        )
    else:
        filas.append(
            (
                "Archivo 3_REMUNERACIÓN_SUBASTAS_E_ID_*",
                "falta",
                f"ningun archivo en {CARPETA_SUBASTAS}/ empieza con "
                f"'3_REMUNERACIÓN_SUBASTAS_E_ID_'",
            )
        )

    return rutas, filas


# ============================================================
# LECTURA DE MEDIDAS SAE (A:I)
# ============================================================

def leer_medidas_sae(ruta):
    """Lee Medidas_SAE.xlsx y valida que traiga las 9 columnas."""

    df = pd.read_excel(ruta, sheet_name=HOJA_MEDIDAS_SAE)

    faltantes = [
        columna for columna in COLUMNAS_AI
        if columna not in df.columns
    ]

    if faltantes:
        raise ErrorEntrada(
            f"A {ARCHIVO_MEDIDAS_SAE} le faltan columnas:\n"
            f"  {faltantes}\n"
            f"Columnas encontradas:\n"
            f"  {list(df.columns)}"
        )

    df = df[COLUMNAS_AI].copy()

    df["intervalo"] = pd.to_datetime(df["intervalo"])

    return df


# ============================================================
# LECTURA DEL MAESTRO
# ============================================================

def _leer_resumen_bess(ruta, nombre_hoja, filas_a_revisar=15):
    """
    Lee "Resumen BESS" detectando la fila de encabezados en vez de
    asumir que es la primera fila: el archivo real trae un titulo
    ("Cuadro N° 1: Resumen BESS") arriba de los encabezados reales.
    Mismo criterio que detectar_fila_nombres() para el SoC: nunca una
    posicion fija, se busca la fila que contiene los textos
    esperados ('Nombre activo' / 'Barra inyección').
    """

    crudo = pd.read_excel(
        ruta, sheet_name=nombre_hoja, header=None, nrows=filas_a_revisar
    )

    fila_encabezado = None

    for indice in range(len(crudo)):

        textos = [normalizar(v) for v in crudo.iloc[indice].tolist()]

        tiene_nombre = any(
            "nombre" in t and "activ" in t for t in textos
        )
        tiene_barra = any("barra" in t for t in textos)

        if tiene_nombre and tiene_barra:
            fila_encabezado = indice
            break

    if fila_encabezado is None:
        raise ErrorEntrada(
            f"No se encontro, en las primeras {filas_a_revisar} filas de "
            f"la hoja '{nombre_hoja}' de {ARCHIVO_CENTRALES}, una fila de "
            f"encabezados con 'Nombre activo' y 'Barra inyección'."
        )

    return pd.read_excel(ruta, sheet_name=nombre_hoja, header=fila_encabezado)


def leer_centrales(ruta):
    """Devuelve (resumen_bess, diccionario) como DataFrames."""

    excel = pd.ExcelFile(ruta)

    def buscar_hoja(nombre_buscado):
        for hoja in excel.sheet_names:
            if normalizar(hoja) == normalizar(nombre_buscado):
                return hoja
        raise ErrorEntrada(
            f"{ARCHIVO_CENTRALES} no tiene la hoja "
            f"'{nombre_buscado}'. Hojas: {excel.sheet_names}"
        )

    resumen = _leer_resumen_bess(ruta, buscar_hoja(HOJA_RESUMEN_BESS))

    diccionario = pd.read_excel(
        ruta,
        sheet_name=buscar_hoja(HOJA_DICCIONARIO),
        header=None,
    )

    return resumen, diccionario


def construir_homologacion(diccionario):
    """
    Arma un mapa nombre_origen -> nombre_canonico a partir de la
    hoja Diccionario, que viene en bloques sin encabezado fijo.

    Estrategia deliberadamente conservadora: cada fila aporta
    equivalencias entre todos sus textos no vacios. No se
    corrige ni reinterpreta nada, tal como pide el plan.
    """

    mapa = {}

    for _, fila in diccionario.iterrows():

        valores = [
            str(v).strip()
            for v in fila.tolist()
            if v is not None
            and str(v).strip() != ""
            and str(v).strip().lower() != "nan"
        ]

        if len(valores) < 2:
            continue

        canonico = valores[0]

        for valor in valores:
            mapa.setdefault(normalizar(valor), canonico)

    return mapa


# ============================================================
# EXTRACCION DEL SoC POR BLOQUES
# ============================================================

def detectar_fila_nombres(df_crudo, maximo_filas=30):
    """
    Busca la fila que contiene los nombres de centrales.

    Criterio: la fila inmediatamente anterior a aquella donde
    aparecen los encabezados 'Time Stamp' / 'Value', mirando
    hacia arriba hasta encontrar una fila con varios textos.
    """

    fila_encabezados = None

    for indice in range(min(maximo_filas, len(df_crudo))):

        textos = {
            normalizar(v)
            for v in df_crudo.iloc[indice].tolist()
        }

        if "time stamp" in textos and "value" in textos:
            fila_encabezados = indice
            break

    if fila_encabezados is None:
        raise ErrorEntrada(
            "No se encontro ninguna fila con los encabezados "
            "'Time Stamp' y 'Value' en el archivo SOC."
        )

    # Hacia arriba, la primera fila con al menos un texto que
    # no sea un encabezado del bloque.
    encabezados_bloque = {
        "status",
        "questionable",
        "time stamp",
        "value",
        "",
    }

    for indice in range(fila_encabezados - 1, -1, -1):

        textos = [
            normalizar(v)
            for v in df_crudo.iloc[indice].tolist()
        ]

        utiles = [
            t for t in textos
            if t not in encabezados_bloque
        ]

        if utiles:
            return indice, fila_encabezados

    raise ErrorEntrada(
        "Se encontraron los encabezados 'Time Stamp'/'Value' "
        "pero no una fila de nombres de centrales arriba."
    )


def detectar_bloques(df_crudo, fila_nombres, fila_encabezados):
    """
    Para cada nombre de la fila de nombres, define su bloque
    horizontal y ubica dentro sus columnas Time Stamp y Value.

    No se usan offsets fijos: los encabezados se buscan por
    texto dentro del rango de cada central.
    """

    encabezados_bloque = {
        "status",
        "questionable",
        "time stamp",
        "value",
        "",
    }

    fila_n = df_crudo.iloc[fila_nombres].tolist()

    posiciones = []

    for indice, valor in enumerate(fila_n):
        texto = normalizar(valor)
        if texto and texto not in encabezados_bloque:
            posiciones.append((indice, str(valor).strip()))

    if not posiciones:
        raise ErrorEntrada(
            "La fila de nombres no contiene ninguna central."
        )

    fila_e = [
        normalizar(v)
        for v in df_crudo.iloc[fila_encabezados].tolist()
    ]

    bloques = []
    incidencias = []

    for orden, (columna_inicio, nombre) in enumerate(posiciones):

        if orden + 1 < len(posiciones):
            columna_fin = posiciones[orden + 1][0] - 1
        else:
            columna_fin = len(fila_e) - 1

        columnas_ts = [
            c
            for c in range(columna_inicio, columna_fin + 1)
            if c < len(fila_e) and fila_e[c] == "time stamp"
        ]

        columnas_val = [
            c
            for c in range(columna_inicio, columna_fin + 1)
            if c < len(fila_e) and fila_e[c] == "value"
        ]

        if len(columnas_ts) != 1 or len(columnas_val) != 1:
            incidencias.append(
                f"{nombre}: se esperaba exactamente un "
                f"'Time Stamp' y un 'Value' entre las columnas "
                f"{columna_inicio + 1} y {columna_fin + 1}; "
                f"se encontraron {len(columnas_ts)} y "
                f"{len(columnas_val)}. Bloque excluido."
            )
            continue

        bloques.append(
            {
                "nombre_bess_origen": nombre,
                "columna_inicio_bloque": columna_inicio,
                "columna_fin_bloque": columna_fin,
                "columna_timestamp": columnas_ts[0],
                "columna_value": columnas_val[0],
            }
        )

    return bloques, incidencias


def _extraer_nombre_desde_ruta_scada(texto):
    """
    Si el nombre de un bloque de SoC viene como una ruta SCADA (visto
    con datos reales: exportaciones tipo PI traen el nombre de la
    central como
        \\SERVIDOR\SEN\Generación\SEN\<region>\<central>|<sufijo>
    en vez de solo "<central>"), devuelve unicamente "<central>": el
    ultimo tramo de la ruta (separado por "\\"), sin el sufijo
    despues de "|".

    No es una reinterpretacion de datos: es separar una ESTRUCTURA
    conocida (ruta + sufijo) que ya viene asi en el archivo, no una
    suposicion sobre a que central corresponde. Si el texto no tiene
    ese formato (no contiene "\\"), se devuelve tal cual -- no se
    inventa nada quitando texto de un nombre que no es una ruta.
    """

    texto = str(texto).strip()

    if "\\" not in texto:
        return texto

    return texto.split("\\")[-1].split("|")[0].strip()


def extraer_soc(ruta_soc, mapa_homologacion=None):
    """
    Lee el archivo SOC y devuelve (df_soc, incidencias).

    df_soc: central | timestamp | soc | nombre_scada_original
    """

    df_crudo = pd.read_excel(
        ruta_soc,
        sheet_name=0,
        header=None,
    )

    fila_nombres, fila_encabezados = detectar_fila_nombres(
        df_crudo
    )

    bloques, incidencias = detectar_bloques(
        df_crudo,
        fila_nombres,
        fila_encabezados,
    )

    mapa_homologacion = mapa_homologacion or {}

    partes = []

    for bloque in bloques:

        nombre_origen = bloque["nombre_bess_origen"]

        # Primero se prueba el texto literal (compatibilidad con
        # cualquier archivo de SoC "limpio", sin ruta SCADA). Si no
        # hay match, se prueba de nuevo con el nombre extraido de la
        # ruta (ver _extraer_nombre_desde_ruta_scada) -- el
        # Diccionario puede tener registrada cualquiera de las dos
        # formas. Si ninguna tiene match, se usa igual el nombre
        # extraido (no la ruta completa) como "canonico": aunque no
        # homologue, es mucho mas legible en avisos/incidencias que
        # la ruta cruda, y no cambia el comportamiento (sigue sin
        # cruzar contra Medidores).
        clave_directa = normalizar(nombre_origen)
        nombre_limpio = _extraer_nombre_desde_ruta_scada(nombre_origen)

        if clave_directa in mapa_homologacion:
            canonico = mapa_homologacion[clave_directa]
        else:
            canonico = mapa_homologacion.get(
                normalizar(nombre_limpio),
                nombre_limpio,
            )

        sub = df_crudo.iloc[
            fila_encabezados + 1:,
            [
                bloque["columna_timestamp"],
                bloque["columna_value"],
            ],
        ].copy()

        sub.columns = ["timestamp", "soc"]

        sub["timestamp"] = pd.to_datetime(
            sub["timestamp"],
            errors="coerce",
        )

        sub["soc"] = pd.to_numeric(
            sub["soc"],
            errors="coerce",
        )

        antes = len(sub)

        sub = sub.dropna(subset=["timestamp"])

        descartadas = antes - len(sub)

        if sub.empty:
            incidencias.append(
                f"{nombre_origen}: bloque sin registros validos."
            )
            continue

        if sub["soc"].isna().any():
            incidencias.append(
                f"{nombre_origen}: "
                f"{int(sub['soc'].isna().sum())} valores de SoC "
                f"no numericos."
            )

        duplicados = sub["timestamp"].duplicated().sum()

        if duplicados:
            incidencias.append(
                f"{nombre_origen}: {duplicados} timestamps "
                f"duplicados."
            )

        sub["central"] = canonico
        sub["nombre_scada_original"] = nombre_origen

        partes.append(sub)

    if not partes:
        raise ErrorEntrada(
            "No se pudo extraer SoC de ningun bloque.\n"
            + "\n".join(incidencias)
        )

    df_soc = pd.concat(partes, ignore_index=True)

    df_soc = df_soc[
        [
            "central",
            "timestamp",
            "soc",
            "nombre_scada_original",
        ]
    ]

    return df_soc, incidencias


# ============================================================
# OFERTAS SSCC
#
# Replica, en este orden, las macros y formulas de la planilla 11:
#   1. Generar_Resumen_Ofertas_SSCC        -> construir_resumen_ofertas_sscc
#   2. OSSCC_CargarResumenEnMedidores      -> cargar_resumen_en_medidores
#   3. Formula de Medidores!V y R          -> calcular_r
#   4. Formula de Medidores!S              -> calcular_s
#   5. Resumir_Medidores_Central_Ventana_
#      Oferta_Completa                     -> construir_resumen_ventana_oferta
#   6. Formula de Medidores!T              -> calcular_t
# ============================================================

def _texto_seguro(valor):
    """Replica OSSCC_TextoSeguro / RESOF: '' para vacio/NaN, si no str().strip()."""

    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass

    return str(valor).strip()


def _tiene_valor(valor):
    """Replica OSSCC_TieneValor / RESOF_TieneValor."""

    return len(_texto_seguro(valor)) > 0


def _es_numero(valor):
    if not _tiene_valor(valor):
        return False

    try:
        float(valor)
        return True
    except (TypeError, ValueError):
        return False


def _valor_clave(valor):
    """
    Replica OSSCC_ValorClave / RESOF_ValorClave: formatea un numero
    sin ceros/decimales sobrantes, o el texto tal cual si no es numero.
    """

    if not _es_numero(valor):
        return _texto_seguro(valor)

    numero = float(valor)

    if numero == int(numero):
        return str(int(numero))

    return repr(numero)


def _limpiar_nombre_mostrar(valor):
    """Replica OSSCC_LimpiarNombreMostrar: NBSP->espacio, trim, espacios colapsados."""

    if not _tiene_valor(valor):
        return ""

    texto = str(valor).replace("\xa0", " ").strip()

    while "  " in texto:
        texto = texto.replace("  ", " ")

    return texto


def _normalizar_nombre_clave(valor):
    """Replica OSSCC_NormalizarNombreClave: LimpiarNombreMostrar + mayusculas."""

    return _limpiar_nombre_mostrar(valor).upper()


def _contiene_bess_o_sae(texto):
    """Replica OSSCC_ContieneBESSoSAE."""

    texto = texto.upper()
    return "BESS" in texto or "SAE" in texto or "BAT" in texto


def _servicio_termina_en_rs(servicio):
    """Replica OSSCC_ServicioTerminaEnRS."""

    return servicio.strip().upper().endswith("_RS")


def _es_respuesta_si(valor):
    """Replica OSSCC_EsRespuestaSi: normaliza y compara contra 'SI'."""

    if not _tiene_valor(valor):
        return False

    texto = str(valor).strip().upper()
    texto = texto.replace("\xa0", "").replace(" ", "").replace("\t", "")

    for con_tilde, sin_tilde in (
        ("Í", "I"), ("Ì", "I"), ("Ï", "I"), ("Î", "I"),
    ):
        texto = texto.replace(con_tilde, sin_tilde)

    for signo in ".,;:":
        texto = texto.replace(signo, "")

    return texto == "SI"


def _normalizar_periodo(valor):
    """
    Replica OSSCC_NormalizarPeriodo: interpreta 'valor' como un
    periodo horario 1..24. Acepta enteros 1..24, fracciones de dia
    estilo Excel, '24:00' y texto de hora/fecha reconocible. None si
    no se puede interpretar.
    """

    if not _tiene_valor(valor):
        return None

    if _es_numero(valor):
        numero = float(valor)

        if numero == int(numero) and 1 <= numero <= 24:
            return int(numero)

        if 0 <= numero < 1:
            total_segundos = round(numero * 86400)
            hora = (total_segundos // 3600) % 24
            return int(hora) + 1

        return None

    texto = str(valor).strip()

    if texto in ("24:00", "24:00:00"):
        return 24

    try:
        marca = pd.to_datetime(texto)
        return int(marca.hour) + 1
    except (ValueError, TypeError):
        return None


def _valor_oferta_binario(valor):
    """Replica OSSCC_ValorOfertaBinario: 1 solo si el valor es exactamente 1."""

    if not _tiene_valor(valor):
        return 0

    if _es_numero(valor):
        return 1 if float(valor) == 1 else 0

    return 1 if str(valor).strip() == "1" else 0


def construir_resumen_ofertas_sscc(ruta_ofertas, registrar=print):
    """
    Replica Generar_Resumen_Ofertas_SSCC.

    Lee TODAS las hojas de ruta_ofertas (columnas A:I, desde la fila
    2), filtra filas cuyo nombre (A) contenga BESS/SAE/BAT y cuyo
    servicio (H) termine en "_RS", y agrupa por (Nombre, Año, Mes,
    Día). Para cada grupo, una columna por cada servicio _RS
    encontrado en TODO el archivo (1 si en ese grupo el servicio tiene
    ofertado ("Sí") las 24 horas, 0 si no) y una columna final "Oferta
    completa" (1 solo si todos los servicios del grupo estan
    completos).

    Devuelve un DataFrame ordenado por Nombre, Año, Mes, Día.
    """

    hojas = pd.read_excel(ruta_ofertas, sheet_name=None, header=None)

    grupos = {}
    datos_grupo = {}
    servicios_globales = set()

    for df_hoja in hojas.values():

        if df_hoja.shape[0] < 2 or df_hoja.shape[1] < 9:
            continue

        for _, fila in df_hoja.iloc[1:, 0:9].iterrows():

            nombre = _texto_seguro(fila.iloc[0])
            anio, mes, dia = fila.iloc[1], fila.iloc[2], fila.iloc[3]
            servicio = _texto_seguro(fila.iloc[7])
            indicador = fila.iloc[8]

            if not _contiene_bess_o_sae(nombre):
                continue

            if not _servicio_termina_en_rs(servicio):
                continue

            if not (
                _tiene_valor(anio)
                and _tiene_valor(mes)
                and _tiene_valor(dia)
            ):
                continue

            clave_grupo = (
                nombre,
                _valor_clave(anio),
                _valor_clave(mes),
                _valor_clave(dia),
            )

            if clave_grupo not in grupos:
                grupos[clave_grupo] = {}
                datos_grupo[clave_grupo] = (nombre, anio, mes, dia)

            servicios_globales.add(servicio)

            horas_servicio = grupos[clave_grupo].setdefault(
                servicio, set()
            )

            if _es_respuesta_si(indicador):
                periodo = _normalizar_periodo(fila.iloc[4])
                if periodo is not None:
                    horas_servicio.add(periodo)

    if not grupos:
        raise ErrorEntrada(
            "No se encontraron registros que cumplan las condiciones "
            "BESS/SAE/BAT y servicio terminado en _RS en "
            f"{Path(ruta_ofertas).name}"
        )

    lista_servicios = sorted(servicios_globales, key=str.upper)

    filas_salida = []

    for clave_grupo, servicios_grupo in grupos.items():

        nombre, anio, mes, dia = datos_grupo[clave_grupo]
        fila_salida = {
            "Nombre": nombre,
            "Año": anio,
            "Mes": mes,
            "Día": dia,
        }

        todos_completos = True

        for servicio in lista_servicios:
            completo = len(servicios_grupo.get(servicio, set())) == 24
            fila_salida[servicio] = 1 if completo else 0
            if not completo:
                todos_completos = False

        fila_salida["Oferta completa"] = (
            1 if todos_completos and lista_servicios else 0
        )

        filas_salida.append(fila_salida)

    columnas = (
        ["Nombre", "Año", "Mes", "Día"]
        + lista_servicios
        + ["Oferta completa"]
    )

    df_resumen = pd.DataFrame(filas_salida, columns=columnas)

    df_resumen["_orden_nombre"] = df_resumen["Nombre"].str.upper()
    df_resumen = (
        df_resumen
        .sort_values(by=["_orden_nombre", "Año", "Mes", "Día"])
        .drop(columns="_orden_nombre")
        .reset_index(drop=True)
    )

    registrar(
        f"  Resumen Ofertas SSCC: {len(df_resumen):,} registros, "
        f"{len(lista_servicios)} servicio(s) _RS: "
        f"{', '.join(lista_servicios)}"
    )

    return df_resumen


def cargar_resumen_en_medidores(
    df_resumen, claves_medidores, diccionario, registrar=print
):
    """
    Replica OSSCC_CargarResumenEnMedidores.

    Arma la tabla equivalente a Medidores!W:Y (Nombre, Dia, Oferta
    completa) con una fila por central x dia del mes: toma los
    nombres y ofertas de df_resumen, y agrega ademas -con oferta 0
    para todos sus dias- cualquier central de Medidores!clave que no
    este representada ahi ni mediante una equivalencia de
    Diccionario!E:F:G.

    claves_medidores: valores unicos de la columna 'clave' de Medidores.

    Devuelve (df_wxy, (anio, mes), avisos).
    """

    if "Oferta completa" not in df_resumen.columns:
        raise ErrorEntrada(
            "El resumen de Ofertas SSCC no tiene la columna "
            "'Oferta completa'."
        )

    dic_nombres = {}
    dic_ofertas = {}
    periodos = set()

    for _, fila in df_resumen.iterrows():

        nombre_mostrar = _limpiar_nombre_mostrar(fila["Nombre"])
        if not nombre_mostrar:
            continue

        anio, mes, dia = fila["Año"], fila["Mes"], fila["Día"]

        if not (_es_numero(anio) and _es_numero(mes) and _es_numero(dia)):
            continue

        anio_i, mes_i, dia_i = int(anio), int(mes), int(dia)

        if not 1 <= mes_i <= 12:
            continue

        periodos.add((anio_i, mes_i))

        clave_nombre = _normalizar_nombre_clave(nombre_mostrar)
        dic_nombres.setdefault(clave_nombre, nombre_mostrar)

        oferta = _valor_oferta_binario(fila["Oferta completa"])
        dic_ofertas[(clave_nombre, dia_i)] = oferta

    if len(periodos) == 0:
        raise ErrorEntrada(
            "No fue posible determinar el año y mes del resumen de "
            "Ofertas SSCC."
        )

    if len(periodos) > 1:
        raise ErrorEntrada(
            "El resumen de Ofertas SSCC contiene mas de un año o "
            f"mes: {sorted(periodos)}. No es posible construir "
            "Medidores!W:Y porque el destino solo admite un periodo "
            "por vez."
        )

    anio, mes = next(iter(periodos))
    dias_del_mes = calendar.monthrange(anio, mes)[1]

    # Equivalencias Diccionario!E:F:G (columnas 5,6,7 -> indices 4,5,6)
    dic_equivalencias = {}

    if diccionario.shape[1] > 6:
        for _, fila in diccionario.iterrows():
            alias = []
            for indice in (4, 5, 6):
                valor = _normalizar_nombre_clave(fila.iloc[indice])
                if valor and valor not in alias:
                    alias.append(valor)
            for valor in alias:
                dic_equivalencias[valor] = set(alias)

    dic_no_encontrados = {}

    for nombre_crudo in claves_medidores:

        nombre_mostrar = _limpiar_nombre_mostrar(nombre_crudo)
        if not nombre_mostrar:
            continue

        clave_nombre = _normalizar_nombre_clave(nombre_mostrar)

        if clave_nombre in dic_nombres:
            continue

        equivalentes = dic_equivalencias.get(clave_nombre)
        ya_representado = False

        if equivalentes:
            ya_representado = any(
                alias in dic_nombres for alias in equivalentes
            )
        else:
            dic_no_encontrados.setdefault(clave_nombre, nombre_mostrar)

        if not ya_representado:
            dic_nombres.setdefault(clave_nombre, nombre_mostrar)

    if not dic_nombres:
        raise ErrorEntrada(
            "No se encontraron nombres en el resumen de Ofertas SSCC "
            "ni en Medidores (columna clave)."
        )

    avisos = []

    if dic_no_encontrados:
        nombres_avisados = sorted(
            dic_no_encontrados.values(), key=str.upper
        )
        cola = (
            f" ... y {len(nombres_avisados) - 30} mas."
            if len(nombres_avisados) > 30 else ""
        )
        avisos.append(
            f"{len(nombres_avisados)} nombre(s) de Medidores!clave no "
            "se encontraron en Diccionario!E:F:G. Se incorporaron con "
            "oferta 0: " + ", ".join(nombres_avisados[:30]) + cola
        )

    nombres_ordenados = sorted(
        dic_nombres.items(), key=lambda kv: kv[1].upper()
    )

    filas_salida = []
    for clave_nombre, nombre in nombres_ordenados:
        for dia in range(1, dias_del_mes + 1):
            oferta = dic_ofertas.get((clave_nombre, dia), 0)
            filas_salida.append((nombre, dia, oferta))

    df_wxy = pd.DataFrame(
        filas_salida, columns=["Nombre", "Dia", "Oferta completa"]
    )

    registrar(
        f"  Ofertas SSCC por dia (equivalente a Medidores!W:Y): "
        f"{len(df_wxy):,} filas ({len(nombres_ordenados)} "
        f"central(es) x {dias_del_mes} dias)"
    )

    return df_wxy, (anio, mes), avisos


def _mapas_homologacion_fge(diccionario):
    """
    Precalcula, desde Diccionario!F y Diccionario!G (columnas 6 y 7,
    indices 5 y 6), el mapa hacia Diccionario!E (columna 5, indice 4)
    que usa la formula de Medidores!V. Ante nombres repetidos se
    conserva el primero, igual que XLOOKUP con la primera coincidencia.
    """

    mapa_f = {}
    mapa_g = {}

    if diccionario.shape[1] > 6:
        for _, fila in diccionario.iterrows():
            valor_e = _texto_seguro(fila.iloc[4])
            clave_f = _normalizar_nombre_clave(fila.iloc[5])
            clave_g = _normalizar_nombre_clave(fila.iloc[6])

            if clave_f and clave_f not in mapa_f:
                mapa_f[clave_f] = valor_e

            if clave_g and clave_g not in mapa_g:
                mapa_g[clave_g] = valor_e

    return mapa_f, mapa_g


def _homologar_fge(nombre, mapa_f, mapa_g):
    """
    Replica el homologador de Medidores!V:
    XLOOKUP(nombre, Diccionario!F, Diccionario!E,
        XLOOKUP(nombre, Diccionario!G, Diccionario!E, 0, 0), 0)
    """

    clave = _normalizar_nombre_clave(nombre)

    if clave in mapa_f:
        return mapa_f[clave]

    if clave in mapa_g:
        return mapa_g[clave]

    return "0"


def calcular_r(df_medidores, df_wxy, diccionario, registrar=print):
    """
    Replica Medidores!R = VLOOKUP(B&G, V:Y, 4, FALSE).

    V (la tabla auxiliar, aca solo en memoria) es Dia & homologado(W)
    via Diccionario!F/G->E, construida a partir de df_wxy (equivalente
    a Medidores!W:Y). El cruce se hace por Dia + nombre homologado en
    mayusculas, separados por '|' (Excel concatena sin separador; acá
    se agrega uno para no confundir, p.ej., dia=1+'0ABC' con
    dia=10+'ABC' - un caso limite que en Excel tampoco se distingue).

    Devuelve (serie_r, avisos).
    """

    mapa_f, mapa_g = _mapas_homologacion_fge(diccionario)

    tabla_v = {}
    for _, fila in df_wxy.iterrows():
        homologado = _homologar_fge(fila["Nombre"], mapa_f, mapa_g)
        clave_v = f"{int(fila['Dia'])}|{homologado.upper()}"
        tabla_v.setdefault(clave_v, fila["Oferta completa"])

    claves_medidores = (
        df_medidores["Dia"].astype("Int64").astype(str)
        + "|"
        + df_medidores["clave"].astype(str).str.strip().str.upper()
    )

    r = claves_medidores.map(tabla_v).astype("Int64")

    avisos = []
    no_encontrados = int(r.isna().sum())

    if no_encontrados:
        avisos.append(
            f"{no_encontrados:,} fila(s) de Medidores no encontraron "
            "coincidencia (Dia + central homologada) en la tabla de "
            "Ofertas SSCC al calcular la columna R. Revisar que todas "
            "las centrales de Medidores!clave esten homologadas en "
            "Diccionario!E:F:G."
        )

    registrar(f"  Columna R (Oferta_Completa_Dia): {no_encontrados:,} sin match")

    return r, avisos


def calcular_s(ventana, r_valor):
    """
    Replica Medidores!S:
        =IF(L3=L2, S2, IF(R3=1, 1, 2))

    S se mantiene igual a la fila anterior mientras la Ventana (L) no
    cambie; al cambiar, se recalcula segun si R vale 1. No se reinicia
    aparte por central: en la planilla original tampoco lo hace, se
    apoya en que L ya cambia al cambiar de central.
    """

    ventana = pd.Series(ventana).reset_index(drop=True)
    r_valor = pd.Series(r_valor).reset_index(drop=True)

    cambia_ventana = ventana.ne(ventana.shift())

    # R sin match (ver calcular_r) se trata como "no es 1", igual que
    # cualquier valor de R distinto de 1 en la formula de Excel.
    es_uno = r_valor.fillna(-1).eq(1)
    valor_si_cambia = es_uno.map({True: 1, False: 2})

    s = valor_si_cambia.where(cambia_ventana)

    return s.ffill().astype("Int64")


def construir_resumen_ventana_oferta(
    clave, ventana, oferta_r, inicio_ventana=INICIO_VENTANA, registrar=print
):
    """
    Replica Resumir_Medidores_Central_Ventana_Oferta_Completa.

    Agrupa Medidores por (clave=Central, Ventana=L), suma R, y marca
    "Completa" segun la oferta esperada para esa ventana: 96 en las
    ventanas normales, (inicio_ventana-1)*4 en la ventana 0, y
    (25-inicio_ventana)*4 en la ultima ventana del mes (la de mayor
    valor numerico encontrado en L).

    Devuelve un DataFrame con columnas Central, Ventana T, Oferta,
    Completa.
    """

    if not 1 <= inicio_ventana <= 24:
        raise ErrorEntrada(
            "INICIO_VENTANA debe ser una hora entre 1 y 24."
        )

    oferta_esperada_inicial = (inicio_ventana - 1) * 4
    oferta_esperada_final = (25 - inicio_ventana) * 4

    ventana_numerica = pd.to_numeric(pd.Series(ventana), errors="coerce")
    ultima_ventana = ventana_numerica.max()

    if pd.isna(ultima_ventana):
        raise ErrorEntrada(
            "No se encontraron valores numericos en la columna "
            "Ventana (L) para resumir Ofertas SSCC."
        )

    df = pd.DataFrame(
        {
            "Central": pd.Series(clave).reset_index(drop=True),
            "Ventana T": ventana_numerica.reset_index(drop=True),
            "Oferta": pd.to_numeric(
                pd.Series(oferta_r), errors="coerce"
            ).reset_index(drop=True),
        }
    )
    df = df.dropna(subset=["Central", "Ventana T"])

    resumen = (
        df.groupby(["Central", "Ventana T"], as_index=False)["Oferta"]
        .sum()
    )

    def oferta_esperada(v):
        if abs(v) < 1e-6:
            return oferta_esperada_inicial
        if abs(v - ultima_ventana) < 1e-6:
            return oferta_esperada_final
        return 96

    esperada = resumen["Ventana T"].map(oferta_esperada)
    resumen["Completa"] = (
        (resumen["Oferta"] - esperada).abs() < 1e-6
    ).astype("int64")

    resumen = (
        resumen
        .sort_values(by=["Central", "Ventana T"])
        .reset_index(drop=True)
    )

    registrar(
        f"  Resumen Ventana Oferta: {len(resumen):,} grupo(s) "
        f"central x ventana; ultima ventana detectada "
        f"{ultima_ventana:g}"
    )

    return resumen


def calcular_t(clave, ventana, resumen_ventana_oferta):
    """
    Replica Medidores!T = 1 - Completa(central=G, ventana=L), buscando
    en el resumen central+ventana+oferta completa (AB:AE).

    Devuelve (serie_t, cantidad_sin_match).
    """

    df = pd.DataFrame(
        {
            "Central": pd.Series(clave).reset_index(drop=True),
            "Ventana T": pd.Series(ventana).reset_index(drop=True),
        }
    )

    cruzado = df.merge(
        resumen_ventana_oferta[["Central", "Ventana T", "Completa"]],
        on=["Central", "Ventana T"],
        how="left",
    )

    t = (1 - cruzado["Completa"]).astype("Int64")

    return t, int(t.isna().sum())


# ============================================================
# CMg, FD, SUBASTAS
#
# Replican las macros de carga de esas tres hojas (no las macros que
# las consumen despues, como Asignar_CMg_a_Calculos_Turbo o
# Actualizar_Calculos_Columnas, que pertenecen a una etapa posterior
# todavia no implementada):
#   - Cargar_CMg_Desde_Archivo               -> leer_cmg
#   - Cargar_SSCC_Desempeno_En_FD            -> construir_fd
#   - Cargar_Remuneracion_Subastas_Rapido    -> construir_subastas
#
# Ninguna de las tres tiene una hoja de referencia de dominio tan
# detallada como la de Medidores (Plan_Traspaso...): no se conocen
# nombres de negocio para casi ninguna columna mas alla de lo que las
# formulas de Excel revelan. Por eso las columnas que solo se copian
# (no se calculan) se nombran con su letra de Excel tal cual, en vez
# de inventarles un nombre que no esta documentado en ningun lado.
# ============================================================

def _contiene_bess_o_sae_sin_bat(texto):
    """
    Replica EsBESSoSAE (macros F_Leer_FD y G_Lee_Subastas): contiene
    "BESS" o "SAE". A diferencia de OSSCC_ContieneBESSoSAE (Ofertas
    SSCC), esta NO incluye "BAT" - son dos filtros distintos aunque se
    parezcan, no simplificar a una sola funcion.
    """

    texto = texto.upper()
    return "BESS" in texto or "SAE" in texto


def _dia_hora_mes_fd(columna_fecha, columna_hora):
    """
    Replica la formula compartida de FD (B, M, R y AE):
        =(DAY(fecha)-1)*24 + hora + 1 + IF(DAY(fecha)>100,1,0)

    El termino IF(DAY(fecha)>100,...) nunca es verdadero para un dia
    de calendario real (DAY() da 1..31): se conserva tal cual, tal
    como esta en la formula original, en vez de "limpiarla".
    """

    fecha = pd.to_datetime(columna_fecha, errors="coerce")
    dia = fecha.dt.day
    hora = pd.to_numeric(columna_hora, errors="coerce")
    ajuste_dia_mayor_100 = (dia > 100).astype("Int64")

    return (dia - 1) * 24 + hora + 1 + ajuste_dia_mayor_100


# --------------------------------------------------------------
# CMg
# --------------------------------------------------------------

def leer_cmg(ruta_cmg, registrar=print):
    """
    Replica Cargar_CMg_Desde_Archivo: lee cmg.xlsx (hoja "CMg" si
    existe, si no la primera hoja), columnas A:I desde la fila 2, y
    las ordena por columna D ascendente y luego H ascendente - el
    mismo orden que la macro aplica sobre el origen antes de pegarlo.

    No se renombran las columnas: se preserva el encabezado real del
    archivo (fila 1), igual que hace la macro al no tocarlo.
    """

    excel = pd.ExcelFile(ruta_cmg)

    if not excel.sheet_names:
        raise ErrorEntrada(
            f"{Path(ruta_cmg).name} no contiene hojas."
        )

    nombre_hoja = (
        HOJA_CMG_ORIGEN
        if HOJA_CMG_ORIGEN in excel.sheet_names
        else excel.sheet_names[0]
    )

    df = pd.read_excel(ruta_cmg, sheet_name=nombre_hoja)

    if df.shape[1] < 9:
        raise ErrorEntrada(
            f"{Path(ruta_cmg).name} debe tener al menos 9 columnas "
            f"(A:I) en la hoja '{nombre_hoja}'; tiene {df.shape[1]}."
        )

    df = df.iloc[:, :9].copy()

    columna_d = df.columns[3]
    columna_h = df.columns[7]

    df = (
        df
        .sort_values(
            by=[columna_d, columna_h],
            kind="mergesort",
            na_position="last",
        )
        .reset_index(drop=True)
    )

    registrar(
        f"  CMg: {len(df):,} filas leidas de {Path(ruta_cmg).name} "
        f"(hoja '{nombre_hoja}')"
    )

    return df


# --------------------------------------------------------------
# FD
# --------------------------------------------------------------

# Encabezados reales de FD (confirmados por el usuario contra un caso
# real, no inventados). Se aplican al final de cada bloque, despues de
# calcular todo con los nombres de letra (asi se evitan columnas
# duplicadas en el DataFrame mientras se opera con el, ya que "Hora
# Mes" se repite dos veces en cada bloque real - B y M en el CSF, R y
# AE en el CPF - igual que en el archivo original).
NOMBRES_FD_CSF = {
    "A": "id",
    "B": "Hora Mes",
    "C": "Dia",
    "D": "Fecha",
    "E": "Hora",
    "F": "Unidad",
    "G": "Respuesta CSF\n (Fact_CSF)",
    "H": "Disponibilidad\n(Fdis_CSF)",
    "I": "Desempeño\n(DCSF)",
    "J": "Factor de Desempeño\n (Fd_CSF)",
    "K": "CSF(+)",
    "L": "CSF(-)",
    "M": "Hora Mes",
}

NOMBRES_FD_CPF = {
    "Q": "id",
    "R": "Hora Mes",
    "S": "Dia",
    "T": "Fecha",
    "U": "Hora",
    "V": "Unidad",
    "W": "Respuesta CPF+\n(Fact_CPF+)",
    "X": "Respuesta CPF-\n(Fact_CPF-)",
    "Y": "Disponibilidad\n(Fdis_CPF)",
    "Z": "Desempeño\n(DCPF)",
    "AA": "Factor de Desempeño\n(Fd_CPF)",
    "AB": "Cuenta con equipo\nregistrador validado",
    "AC": "CPF(+)",
    "AD": "CPF(-)",
    "AE": "Hora Mes",
}

def _filtrar_bess_sae_posicional(df_bloque, indice_columna_filtro):
    """
    Replica FiltrarFilasBESSoSAE: conserva las filas donde la columna
    dada (posicion 0-indexada dentro de df_bloque) contiene "BESS" o
    "SAE".
    """

    textos = df_bloque.iloc[:, indice_columna_filtro].map(_texto_seguro)
    mascara = textos.map(_contiene_bess_o_sae_sin_bat)

    return df_bloque[mascara].reset_index(drop=True)


def _construir_bloque_fd_csf(df_filtrado):
    """
    A partir del bloque ya filtrado (7 columnas, origen B:H de "CSF
    Horario" en ese orden), arma las columnas A:M de FD tal como las
    escribe Cargar_SSCC_Desempeno_En_FD + sus formulas (plan de
    migracion, seccion de formulas de FD).
    """

    df = df_filtrado.iloc[:, :7].copy()
    df.columns = ["D", "E", "F", "G", "H", "I", "J"]
    df = df.reset_index(drop=True)

    df["B"] = _dia_hora_mes_fd(df["D"], df["E"])
    df["C"] = pd.to_datetime(df["D"], errors="coerce").dt.day
    df["A"] = (
        df["B"].astype("Int64").astype(str)
        + df["F"].map(_texto_seguro)
    )
    df["K"] = df["J"]
    df["L"] = df["K"]
    df["M"] = df["B"]

    df = df[list("ABCDEFGHIJKLM")]

    # Los nombres reales duplican "Hora Mes" (B y M): se renombra al
    # final, ya con las columnas en su posicion definitiva.
    return df.set_axis(
        [NOMBRES_FD_CSF[letra] for letra in df.columns], axis=1
    )


def _construir_bloque_fd_cpf(df_filtrado):
    """
    A partir del bloque ya filtrado (9 columnas, origen B:J de "CPF
    Horario" en ese orden), arma las columnas Q:AE de FD tal como las
    escribe Cargar_SSCC_Desempeno_En_FD + sus formulas.
    """

    df = df_filtrado.iloc[:, :9].copy()
    df.columns = ["T", "U", "V", "W", "X", "Y", "Z", "AA", "AB"]
    df = df.reset_index(drop=True)

    df["R"] = _dia_hora_mes_fd(df["T"], df["U"])
    df["S"] = pd.to_datetime(df["T"], errors="coerce").dt.day
    df["Q"] = (
        df["R"].astype("Int64").astype(str)
        + df["V"].map(_texto_seguro)
    )
    df["AC"] = df["AA"]
    df["AD"] = df["AC"]
    df["AE"] = df["R"]

    df = df[list("QRSTUVWXYZ") + ["AA", "AB", "AC", "AD", "AE"]]

    # Los nombres reales duplican "Hora Mes" (R y AE): se renombra al
    # final, ya con las columnas en su posicion definitiva.
    return df.set_axis(
        [NOMBRES_FD_CPF[letra] for letra in df.columns], axis=1
    )


def construir_fd(ruta_sscc, registrar=print):
    """
    Replica Cargar_SSCC_Desempeno_En_FD.

    Lee, del archivo SSCC_Desempeño_*, las hojas "CPF Horario" y "CSF
    Horario" desde la fila 12, filtra por BESS/SAE en la columna D de
    cada una, y arma dos bloques independientes (distinto largo cada
    uno, igual que en la planilla): A:M (desde CSF) y Q:AE (desde
    CPF), con sus nombres de columna reales (NOMBRES_FD_CSF/
    NOMBRES_FD_CPF, confirmados por el usuario). N:P quedan fuera de
    alcance (la macro no las toca).

    Devuelve (df_csf, df_cpf).
    """

    ruta_sscc = Path(ruta_sscc)

    excel = pd.ExcelFile(ruta_sscc)

    for hoja in (HOJA_CPF_HORARIO, HOJA_CSF_HORARIO):
        if hoja not in excel.sheet_names:
            raise ErrorEntrada(
                f"No existe la hoja '{hoja}' en {ruta_sscc.name}."
            )

    df_cpf_crudo = pd.read_excel(
        ruta_sscc, sheet_name=HOJA_CPF_HORARIO, header=None
    )
    df_csf_crudo = pd.read_excel(
        ruta_sscc, sheet_name=HOJA_CSF_HORARIO, header=None
    )

    # Fila 12 de Excel (1-indexada) = indice 11 (0-indexado).
    # CPF: columnas B:J (9); CSF: columnas B:H (7).
    bloque_cpf = df_cpf_crudo.iloc[11:, 1:10]
    bloque_csf = df_csf_crudo.iloc[11:, 1:8]

    # D es la 3ra columna de cada bloque (B, C, D -> indice 2).
    filtrado_cpf = _filtrar_bess_sae_posicional(bloque_cpf, 2)
    filtrado_csf = _filtrar_bess_sae_posicional(bloque_csf, 2)

    df_csf = _construir_bloque_fd_csf(filtrado_csf)
    df_cpf = _construir_bloque_fd_cpf(filtrado_cpf)

    registrar(
        f"  FD: {len(df_csf):,} fila(s) CSF Horario, "
        f"{len(df_cpf):,} fila(s) CPF Horario (filtro BESS/SAE)"
    )

    return df_csf, df_cpf


# --------------------------------------------------------------
# SUBASTAS
# --------------------------------------------------------------

# Encabezados reales de Subastas!B:Q (confirmados por el usuario). "A"
# (Concepto) no esta: la macro Cargar_Remuneracion_Subastas_Rapido no
# la toca. "Q" no tiene encabezado en el archivo real (queda como
# columna sin nombre, no se le inventa uno).
NOMBRES_SUBASTAS = {
    "B": "Control",
    "C": "Sub_Baj",
    "D": "Fecha",
    "E": "Año",
    "F": "Mes",
    "G": "Dia",
    "H": "Hora_dia",
    "I": "Hora_mes",
    "J": "Configuración",
    "K": "Propietario",
    "L": "Clave horaria",
    "M": "Ciclo",
    "N": "Energía SSCC",
    "O": "FD",
    "P": "FMA",
    "Q": "",
}


def construir_subastas(ruta_subastas, registrar=print):
    """
    Replica Cargar_Remuneracion_Subastas_Rapido.

    La macro original consulta la hoja "DB" del archivo de origen por
    ADO/SQL (equivalente a filtrar y seleccionar columnas de una
    tabla); aca se lee directamente con pandas y se aplica el mismo
    filtro y la misma seleccion de columnas.

    Arma las columnas B:Q de Subastas (nombres reales en
    NOMBRES_SUBASTAS, confirmados por el usuario):
      - Control:Clave horaria (B:L): copia directa de DB!B:L
        (filtrado por Propietario/DB!K contiene BESS/SAE).
      - Ciclo (M, formula): = Propietario & Hora_dia & Hora_mes
        (K&H&I).
      - Energía SSCC (N): se deja vacia EN ESTA HOJA. Ya no es un
        calculo desconocido (ver calcular_subastas_energia_sscc: es
        el "Ciclo de Carga del mes" de Calculo E Costos homologado
        por Hora_mes + Configuración), pero depende de Calculo E
        Costos, que se arma despues y en el otro archivo
        (Pagos_BESS.xlsx). Se calcula ahi, donde se usa.
      - FD, FMA (O, P) y la columna sin nombre (Q): copias de DB!P,
        DB!Y, DB!V respectivamente (asi lo indica la macro original).
    """

    ruta_subastas = Path(ruta_subastas)

    excel = pd.ExcelFile(ruta_subastas)

    if HOJA_SUBASTAS_ORIGEN not in excel.sheet_names:
        raise ErrorEntrada(
            f"No existe la hoja '{HOJA_SUBASTAS_ORIGEN}' en "
            f"{ruta_subastas.name}."
        )

    df_crudo = pd.read_excel(
        ruta_subastas, sheet_name=HOJA_SUBASTAS_ORIGEN, header=None
    )

    # Fila 3 de Excel (1-indexada) = indice 2. Columnas B:Y (24).
    bloque = df_crudo.iloc[2:, 1:25]

    # K es la 10ma columna del bloque B:Y (B=0 ... K=9).
    filtrado = _filtrar_bess_sae_posicional(bloque, 9)

    df = filtrado.iloc[:, 0:11].copy()
    df.columns = list("BCDEFGHIJKL")
    df = df.reset_index(drop=True)

    df["M"] = (
        df["K"].map(_texto_seguro)
        + df["H"].map(_texto_seguro)
        + df["I"].map(_texto_seguro)
    )

    df["N"] = pd.NA

    # P, Y, V del bloque original (indices 14, 23, 20) -> O, P, Q.
    df["O"] = filtrado.iloc[:, 14].reset_index(drop=True)
    df["P"] = filtrado.iloc[:, 23].reset_index(drop=True)
    df["Q"] = filtrado.iloc[:, 20].reset_index(drop=True)

    df = df[list("BCDEFGHIJKLMNOPQ")]
    df = df.rename(columns=NOMBRES_SUBASTAS)

    registrar(
        f"  Subastas: {len(df):,} fila(s) (filtro Propietario "
        f"contiene BESS/SAE)"
    )

    return df


# ============================================================
# CALCULO E COSTOS (etapa base)
#
# El usuario pidio avanzar "por etapas: primero H + CMg + traspaso
# de Medidores". Lo que sigue replica solo esa parte de dos macros:
#
#   - Traspasar_Medidores_A_Calculos_Rapido (modulo
#     B_medidores_a_calculos): A:G (con D<->E invertidas), I/J
#     (energia de Medidores!I separada por signo, solo si
#     Ventana_No_Completa==1; si no, es de "Calculo RE545", fuera de
#     alcance), Medidores!J -> K, Medidores!K -> P. H NO la toca esta
#     macro (es formula, ver mas abajo).
#   - Asignar_CMg_a_Calculos_Turbo (modulo A_Carga_Cmg_a_Destino):
#     arma un diccionario CMg!D (Barra) + "|" + CMg!H (Cuarto de
#     Hora, normalizado con NormalizaCuarto) -> CMg!F, y lo vuelca en
#     la columna Q. Para "Calculo E Costos" la macro NO escribe R
#     (escribirR=False): eso solo aplica a "Calculo RE545".
#
# H (Barra), en la planilla original, es formula:
#     =VLOOKUP(G4, Resumen!B:G, 6, FALSE)
# Se homologa por NOMBRE de columna ("Nombre activo" / "Barra
# inyeccion" de Centrales.xlsx!Resumen BESS) en vez de por posicion,
# porque Centrales.xlsx no reproduce el layout Resumen!B:G del libro
# original.
#
# El resto de columnas de Actualizar_Calculos_Columnas (L, M, N, O,
# R, S, T, U, W, X, Y, AB:AF, AG:AX, AZ) queda para una etapa
# posterior (decision explicita del usuario).
#
# Nombres de columna: son PLACEHOLDERS derivados de los comentarios
# de la macro. Todavia no se pudo confirmar contra un archivo real
# con los encabezados de "Calculo E Costos" (las dos veces que el
# usuario adjunto un archivo para esto, solo traia las hojas FD y
# Subastas) -- se corrigen apenas se reciba ese archivo.
# ============================================================

def _normaliza_cuarto(valor):
    """
    Replica NormalizaCuarto:
        Error       -> ""
        Numerico    -> CStr(CLng(valor))  (texto del entero redondeado)
        Otro        -> Trim(CStr(valor))
    """

    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(valor, str):
        texto = valor.strip()
        if texto == "":
            return ""
        try:
            numero = float(texto.replace(",", "."))
        except ValueError:
            return texto
        return str(round(numero))

    if isinstance(valor, (int, float)):
        return str(round(float(valor)))

    return str(valor).strip()


def construir_dic_cmg(df_cmg):
    """
    Replica el paso 1) de Asignar_CMg_a_Calculos_Turbo: arma un
    diccionario clave -> (valor para Q, valor para R), a partir de la
    hoja CMg (columnas por posicion, sin renombrar - ver leer_cmg):
        D (indice 3) = Barra
        F (indice 5) = valor a asignar en Q
        H (indice 7) = Cuarto de Hora
        I (indice 8) = valor a asignar en R ("CMg Promedio")

    Los dos valores son los dos elementos del Array() que guarda el
    diccionario del VBA: dictCMg(clave)(0) va a Q en las dos hojas y
    dictCMg(clave)(1) va a R, pero SOLO en "Calculo RE545"
    (CompletarDestinoTurbo se llama con escribirR:=False para
    "Calculo E Costos" y escribirR:=True para "Calculo RE545").

    Si una clave se repite, gana la primera fila (igual que
    "If Not dictCMg.Exists(clave) Then Add" en VBA).
    """

    columna_d = df_cmg.columns[3]
    columna_f = df_cmg.columns[5]
    columna_h = df_cmg.columns[7]
    columna_i = df_cmg.columns[8]

    diccionario = {}

    for _, fila in df_cmg.iterrows():

        barra = fila[columna_d]
        barra = "" if pd.isna(barra) else str(barra).strip()

        cuarto_hora = _normaliza_cuarto(fila[columna_h])

        if barra == "" or cuarto_hora == "":
            continue

        clave = barra.upper() + "|" + cuarto_hora

        if clave not in diccionario:
            diccionario[clave] = (fila[columna_f], fila[columna_i])

    return diccionario


def construir_mapa_barra(resumen_bess):
    """
    Arma nombre_central -> barra de inyeccion, a partir de la hoja
    "Resumen BESS" de Centrales.xlsx (columnas "Nombre activo" y
    "Barra inyeccion", confirmadas en el plan de traspaso, seccion
    4.1). Se busca por nombre de columna normalizado, no por
    posicion.
    """

    columna_nombre = None
    columna_barra = None

    for columna in resumen_bess.columns:
        clave = normalizar(columna)
        if columna_nombre is None and "nombre" in clave and "activ" in clave:
            columna_nombre = columna
        if columna_barra is None and "barra" in clave:
            columna_barra = columna

    if columna_nombre is None or columna_barra is None:
        raise ErrorEntrada(
            f"La hoja '{HOJA_RESUMEN_BESS}' de {ARCHIVO_CENTRALES} debe "
            f"tener una columna de nombre de central ('Nombre activo') y "
            f"una de barra de inyeccion ('Barra inyección'). Columnas "
            f"encontradas: {list(resumen_bess.columns)}"
        )

    mapa = {}

    for _, fila in resumen_bess.iterrows():

        nombre = fila[columna_nombre]
        if pd.isna(nombre):
            continue

        barra = fila[columna_barra]
        mapa[normalizar(nombre)] = "" if pd.isna(barra) else str(barra).strip()

    return mapa


def construir_dic_resumen_factor(resumen_bess):
    """
    Arma nombre_central -> factor (columna "Pmax (MW)") y el umbral
    global de SoC minimo, a partir de la MISMA hoja "Resumen BESS" de
    Centrales.xlsx que ya usa construir_mapa_barra().

    Replica Resumen!B:C (factor, usado en AE/AF) y Resumen!H8
    (umbral, usado en M) de Actualizar_Calculos_Columnas. El usuario
    confirmo con un archivo real que la hoja "Resumen" del libro
    original es la MISMA tabla que "Resumen BESS" (los mismos 9
    encabezados: Nombre activo...Eficiencia) -- no hace falta una
    hoja nueva ni un archivo aparte.

    El umbral (celda fija H8 en el original) es, en la practica, el
    valor de "% Energia sobre minima" de la PRIMERA fila de datos de
    la tabla -- aca se toma igual (primera fila con nombre de
    central, no una fila fija: el encabezado de Centrales.xlsx no
    esta siempre en la misma posicion, ver _leer_resumen_bess()).
    """

    columna_nombre = None
    columna_factor = None
    columna_umbral = None

    for columna in resumen_bess.columns:
        clave = normalizar(columna)
        if columna_nombre is None and "nombre" in clave and "activ" in clave:
            columna_nombre = columna
        if columna_factor is None and "pmax" in clave:
            columna_factor = columna
        if columna_umbral is None and "energia sobre" in clave:
            columna_umbral = columna

    if columna_nombre is None or columna_factor is None or columna_umbral is None:
        raise ErrorEntrada(
            f"La hoja '{HOJA_RESUMEN_BESS}' de {ARCHIVO_CENTRALES} debe "
            f"tener columnas de nombre de central ('Nombre activo'), "
            f"factor ('Pmax (MW)') y umbral ('% Energía sobre mínima "
            f"(indicador nuevo ciclo)'). Columnas encontradas: "
            f"{list(resumen_bess.columns)}"
        )

    dic_factor = {}

    for _, fila in resumen_bess.iterrows():

        nombre = fila[columna_nombre]
        if pd.isna(nombre):
            continue

        factor = fila[columna_factor]
        dic_factor[normalizar(nombre)] = (
            pd.NA if pd.isna(factor) else float(factor)
        )

    filas_con_nombre = resumen_bess[columna_nombre].notna()

    if not filas_con_nombre.any():
        raise ErrorEntrada(
            f"La hoja '{HOJA_RESUMEN_BESS}' de {ARCHIVO_CENTRALES} no "
            f"tiene filas de datos para sacar el umbral de SoC minimo."
        )

    primer_indice = resumen_bess.index[filas_con_nombre][0]
    umbral_soc_minimo = float(resumen_bess.loc[primer_indice, columna_umbral])

    return dic_factor, umbral_soc_minimo


def _buscar_cmg(dic_cmg, barra, cuarto_hora):
    """
    Replica la busqueda de CompletarDestinoTurbo: clave
    UCase(Barra)+"|"+NormalizaCuarto(Cuarto de Hora). Devuelve
    siempre un par (valor para Q, valor para R); sin match, los dos
    en blanco (el VBA escribe vbNullString en las dos).
    """

    barra = "" if not barra else str(barra).strip()
    cuarto = _normaliza_cuarto(cuarto_hora)

    if barra == "" or cuarto == "":
        return (pd.NA, pd.NA)

    return dic_cmg.get(barra.upper() + "|" + cuarto, (pd.NA, pd.NA))


def construir_calculo_e_costos(
    df_medidores, mapa_barra, dic_cmg, registrar=print
):
    """
    Etapa base de "Calculo E Costos": traspaso desde Medidores (A:G
    con D<->E invertidas, I/J, K, P) + H (Barra, homologada por
    nombre) + Q (CMg, homologado por Barra+Cuarto de Hora). Ver el
    comentario de seccion mas arriba para el detalle de cada macro
    replicada.

    Solo cubre "Calculo E Costos" (Ventana_No_Completa == 1);
    "Calculo RE545" (Ventana_No_Completa <> 1) queda fuera de esta
    etapa.
    """

    n = len(df_medidores)
    df_medidores = df_medidores.reset_index(drop=True)

    df = pd.DataFrame(index=range(n))

    df["Mes"] = df_medidores["Mes"]
    df["Dia"] = df_medidores["Dia"]
    df["Hora"] = df_medidores["Hora"]

    # D <-> E invertidas: Destino D = Medidores E, Destino E = Medidores D.
    df["Hora Mes"] = df_medidores["Hora Mes"]
    df["Minutos"] = df_medidores["Minutos"]

    df["Cuarto de Hora"] = df_medidores["Cuarto de Hora"]
    df["clave"] = df_medidores["clave"]

    df["Barra"] = df["clave"].map(
        lambda valor: mapa_barra.get(normalizar(valor), "")
    )

    energia = pd.to_numeric(
        df_medidores["Gen_Unidad"], errors="coerce"
    ).fillna(0.0)

    va_a_ecostos = pd.to_numeric(
        df_medidores["Ventana_No_Completa"], errors="coerce"
    ).eq(1)

    energia_positiva = energia.where(energia > 0, 0.0).where(va_a_ecostos, 0.0)
    energia_negativa = energia.where(energia < 0, 0.0).where(va_a_ecostos, 0.0)

    df["Energia_Positiva"] = energia_positiva
    df["Energia_Negativa"] = energia_negativa

    # Medidores J (SoC) -> Destino K; Medidores K (Copia_Ventana) -> Destino P.
    df["SoC"] = df_medidores["SoC"]
    df["Copia_Ventana"] = df_medidores["Copia_Ventana"]

    df["CMg"] = [
        _buscar_cmg(dic_cmg, barra, cuarto_hora)[0]
        for barra, cuarto_hora in zip(df["Barra"], df["Cuarto de Hora"])
    ]

    sin_barra = int((df["Barra"] == "").sum())
    sin_cmg = int(df["CMg"].isna().sum())

    if sin_barra:
        registrar(
            f"  Calculo E Costos: {sin_barra:,} fila(s) sin barra de "
            f"inyeccion (central no encontrada en '{HOJA_RESUMEN_BESS}')."
        )

    if sin_cmg:
        registrar(
            f"  Calculo E Costos: {sin_cmg:,} fila(s) sin CMg (sin match "
            f"Barra+Cuarto de Hora en {ARCHIVO_CMG})."
        )

    registrar(
        f"  Calculo E Costos: {n:,} fila(s) traspasadas desde Medidores."
    )

    return df


_HOJAS_PAGOS = (HOJA_CALCULO_ECOSTOS, HOJA_CALCULO_RE545)


def escribir_pagos_bess(
    ruta_salida,
    df_ecostos=None,
    df_re545=None,
    df_resumen_re545=None,
    ruta_existente=None,
    hojas_regenerar=None,
    registrar=print,
):
    """
    Escribe Pagos_BESS.xlsx: la hoja "Calculo E Costos" (si se pasa
    df_ecostos) y la hoja "Calculo RE545" (si se pasa df_re545, con
    la tabla resumen df_resumen_re545 al lado si tambien se pasa). El
    usuario pidio explicitamente que esto viva en un archivo separado
    de Consolidado_entradas.xlsx ("pagos_bess o algo asi por ahora")
    -- el nombre es provisorio.

    hojas_regenerar: None (por defecto) escribe cada hoja para la que
    se paso su DataFrame, sin mas (asi funcionaba antes de que
    generar_pagos_bess() tuviera casillas por seccion). Si es un set
    con alguno de los nombres de _HOJAS_PAGOS ("Calculo E Costos",
    "Calculo RE545"), la(s) que NO esten en el set se copian tal cual
    desde ruta_existente en vez de escribirse desde el DataFrame --
    mismo criterio que escribir_salida()/hojas_regenerar para
    Consolidado_entradas.xlsx (una hoja destildada en la ventana
    "Generar" se preserva, no se recalcula). Si una hoja a preservar
    no existe en ruta_existente, queda vacia y se registra un aviso.
    """

    ruta_salida = Path(ruta_salida)

    regenerar = (
        set(_HOJAS_PAGOS) if hojas_regenerar is None else set(hojas_regenerar)
    )

    wb_existente = None
    if hojas_regenerar is not None and ruta_existente is not None:
        ruta_existente = Path(ruta_existente)
        if ruta_existente.is_file():
            wb_existente = openpyxl.load_workbook(
                ruta_existente, data_only=True
            )

    avisos_preservacion = []

    def _preservar_o_avisar(writer, nombre_hoja):
        if _copiar_hoja_existente(wb_existente, nombre_hoja, writer.book):
            return
        pd.DataFrame().to_excel(writer, sheet_name=nombre_hoja, index=False)
        mensaje = (
            f"No se regenero la hoja '{nombre_hoja}' (seccion no "
            f"tildada) y no se encontro una version anterior para "
            f"preservarla; quedo vacia."
        )
        avisos_preservacion.append(mensaje)

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:

        if HOJA_CALCULO_ECOSTOS in regenerar:
            if df_ecostos is not None:
                df_ecostos.to_excel(
                    writer,
                    sheet_name=HOJA_CALCULO_ECOSTOS,
                    index=False,
                )
        else:
            _preservar_o_avisar(writer, HOJA_CALCULO_ECOSTOS)

        if HOJA_CALCULO_RE545 in regenerar:
            if df_re545 is not None:
                df_re545.to_excel(
                    writer,
                    sheet_name=HOJA_CALCULO_RE545,
                    index=False,
                )

                if df_resumen_re545 is not None:
                    # Tabla de otro largo, al lado del bloque
                    # principal con una columna en blanco de
                    # separacion (mismo criterio que CSF/CPF de FD).
                    df_resumen_re545.to_excel(
                        writer,
                        sheet_name=HOJA_CALCULO_RE545,
                        index=False,
                        startcol=len(df_re545.columns) + 1,
                    )
        else:
            _preservar_o_avisar(writer, HOJA_CALCULO_RE545)

    for mensaje in avisos_preservacion:
        registrar(f"  [AVISO] {mensaje}")

    registrar(f"Archivo generado: {ruta_salida}")

    return ruta_salida


# ============================================================
# CALCULO RE545 (etapa base): A:V
#
# La otra hoja de calculo del libro, hermana de "Calculo E Costos".
# Las dos se alimentan de la MISMA macro de traspaso
# (Traspasar_Medidores_A_Calculos_Rapido), que reparte cada fila de
# Medidores a una o a la otra segun Medidores!L (Ventana_No_Completa):
#   Ventana_No_Completa = 1  -> la energia va a "Calculo E Costos"
#   cualquier otro valor, o vacio -> va a "Calculo RE545"
# Las columnas A:G, K y P se escriben IGUALES en las dos hojas (no se
# reparten): lo unico que cambia es I/J, que quedan en 0 en la hoja
# que no corresponde.
#
# Diferencias propias de RE545 respecto de E Costos:
#   - T ("Ventana de valorizacion") = Medidores!L. En E Costos no
#     existe: la macro solo escribe T en RE545.
#   - R ("CMg Promedio") = CMg!I, via el mismo diccionario de
#     Asignar_CMg_a_Calculos_Turbo, que para esta hoja se llama con
#     escribirR:=True.
#   - Las columnas calculadas son OTRAS y, cuando comparten letra con
#     E Costos, casi nunca significan lo mismo (R, S, T, U, V son el
#     ejemplo claro). Los nombres reales estan en NOMBRES_CALCULO_RE545,
#     confirmados contra el archivo "Calculo_RE545_reducido_para_IA.xlsx"
#     que entrego el usuario (fila 3 del original = nombres, filas 1-2
#     = titulos de grupo).
#
# Igual que en E Costos, las columnas vacias del original (W:AB, AV,
# BH, BP, BS...) no se escriben: la hoja de salida no reproduce la
# letra de Excel, solo el orden y el contenido.
#
# Ver plan seccion 26.
# ============================================================

# Nombres reales de columna de "Calculo RE545" (fila 3 del archivo
# real). Mismo criterio que NOMBRES_CALCULO_E_COSTOS: se calcula todo
# con las letras/nombres internos y se renombra recien al final.
NOMBRES_CALCULO_RE545 = {
    "Mes": "Mes",
    "Dia": "Dia",
    "Hora": "Hora",
    "Hora Mes": "Hora mes",
    "Minutos": "Minuto",
    "Cuarto de Hora": "Bloque horario",
    "clave": "Configuracion",
    "Barra": "Barra",
    "Energia_Positiva": "Descarga kWh",
    "Energia_Negativa": "Carga kWh",
    "SoC": "SoC %",
    "L": "Adj SSCC",
    "M": "SoC sobre el minimo",
    "N": "Energía SSCC (-) por remunerar",
    "O": "Energía SSCC (+) por remunerar",
    "Copia_Ventana": "Ciclo de Carga del mes",
    "CMg": "CMg",
    "R": "CMg Promedio",
    "S": "ranking cmg",
    "T": "Ventana de valorizacion",
    "U": "EiniT",
    "V": "EalmT",
    "AC": "CPF(-)",
    "AD": "CSF(-)",
    "AE": "CTF(-)",
    "AF": "CPF(+)",
    "AG": "CSF(+)",
    "AH": "CTF(+)",
    "AI": "CPF(-)",
    "AJ": "CSF(-)",
    "AK": "CTF(-)",
    "AL": "CPF(+)",
    "AM": "CSF(+)",
    "AN": "CTF(+)",
    "AO": "CPF(-)",
    "AP": "CSF(-)",
    "AQ": "CTF(-)",
    "AR": "CPF(+)",
    "AS": "CSF(+)",
    "AT": "CTF(+)",
    "AU": "SUMA Reservas*FMA*FD",
    "BI": "Orden",
    "BJ": "Periodo",
    "BK": "Curva Cmg Decendente promedio horario",
    "BL": "",
    "BM": "Curva Cmg Decendente",
    "BN": "Edisp_Asig",
    "BO": "Total  C1_545",
    "BQ": "Energia Total",
    "BR": "Ventana de Valorizacion",
    "BS": "inyeccion en el periodo del Cmg Descendente",
    "BT": "Energía ya Asignada",
    "BU": "Asignacion Edisponible",
    "BV": "SSCC ultima hora",
    "BW": "inyeccion orden cronologico",
    "BX": "Energia Ultima hora",
    "BY": "Energía ya Asignada ultima hora",
    "BZ": "Energia Asignada Ultima hora",
    "CA": "Asignacion Edisponible+SSCC ultima hora",
    "CC": "Total C2_545",
    "CE": "Monto a compensar",
}


def construir_dic_resumen_capacidad(resumen_bess):
    """
    Arma nombre_central -> "Capacidad (MWh)", de la hoja
    "Resumen BESS" de Centrales.xlsx.

    Es el VLOOKUP(G, Resumen!$B$8:$J$26, 4, 0) que aparece en U y BC
    de RE545 (y en BC de Calculo E Costos). OJO: la 4ta columna del
    rango B:J NO es "Pmax (MW)" sino "Capacidad (MWh)" -- el orden
    real de la tabla es Nombre activo, Pmax (MW), Horas para descarga
    forzada, Capacidad (MWh), Energia minima, Barra inyeccion, %
    Energia sobre minima, Ciclos max diarios, Eficiencia (confirmado
    en el plan seccion 25.8, y consistente con que H use el indice 6
    para la barra de inyeccion).

    El "factor" de AE/AF de Calculo E Costos es otra cosa (Resumen!
    B:C, o sea el indice 2 = "Pmax (MW)"): eso lo da
    construir_dic_resumen_factor(), y es el mismo indice 2 que usa BN
    de RE545. No confundirlas.
    """

    return _mapa_resumen_bess_por_nombre(
        resumen_bess, "capacidad", "Capacidad (MWh)",
    )


def _mapa_resumen_bess_por_nombre(resumen_bess, texto_buscado, etiqueta):
    """
    Helper comun: central -> valor de la columna de "Resumen BESS"
    cuyo nombre normalizado contiene texto_buscado.
    """

    columna_nombre = None
    columna_valor = None

    for columna in resumen_bess.columns:
        clave = normalizar(columna)
        if columna_nombre is None and "nombre" in clave and "activ" in clave:
            columna_nombre = columna
        if columna_valor is None and texto_buscado in clave:
            columna_valor = columna

    if columna_nombre is None or columna_valor is None:
        raise ErrorEntrada(
            f"La hoja '{HOJA_RESUMEN_BESS}' de {ARCHIVO_CENTRALES} debe "
            f"tener columnas de nombre de central ('Nombre activo') y "
            f"'{etiqueta}' (hace falta para Calculo RE545). Columnas "
            f"encontradas: {list(resumen_bess.columns)}"
        )

    mapa = {}

    for _, fila in resumen_bess.iterrows():

        nombre = fila[columna_nombre]
        if pd.isna(nombre):
            continue

        valor = fila[columna_valor]
        mapa[normalizar(nombre)] = pd.NA if pd.isna(valor) else float(valor)

    return mapa


def construir_dic_resumen_eficiencia(resumen_bess):
    """
    Arma nombre_central -> "Eficiencia", de la misma hoja
    "Resumen BESS" de Centrales.xlsx.

    Es el VLOOKUP(G, Resumen!$B$8:$J$26, 9, 0) de la columna V de
    RE545: la 9na columna del rango B:J es la ultima de las 9 de esa
    tabla, "Eficiencia".
    """

    return _mapa_resumen_bess_por_nombre(
        resumen_bess, "eficiencia", "Eficiencia",
    )


def construir_calculo_re545(
    df_medidores, mapa_barra, dic_cmg, registrar=print
):
    """
    Etapa base de "Calculo RE545": traspaso desde Medidores (A:G, I/J,
    K, P, T) + H (Barra) + Q/R (CMg y CMg Promedio).

    Es la hoja espejo de construir_calculo_e_costos(): misma macro de
    traspaso, misma A:G, misma K/P, y la energia repartida al reves
    (aca entra la de las filas con Ventana_No_Completa <> 1, incluido
    el caso "vacio o no numerico", que el VBA manda explicitamente a
    RE545).
    """

    n = len(df_medidores)
    df_medidores = df_medidores.reset_index(drop=True)

    df = pd.DataFrame(index=range(n))

    df["Mes"] = df_medidores["Mes"]
    df["Dia"] = df_medidores["Dia"]
    df["Hora"] = df_medidores["Hora"]
    df["Hora Mes"] = df_medidores["Hora Mes"]
    df["Minutos"] = df_medidores["Minutos"]
    df["Cuarto de Hora"] = df_medidores["Cuarto de Hora"]
    df["clave"] = df_medidores["clave"]

    df["Barra"] = df["clave"].map(
        lambda valor: mapa_barra.get(normalizar(valor), "")
    )

    energia = pd.to_numeric(
        df_medidores["Gen_Unidad"], errors="coerce"
    ).fillna(0.0)

    # El VBA manda a RE545 todo lo que NO tiene T = 1 numerico:
    # distinto de 1, vacio, no numerico o error.
    va_a_ecostos = pd.to_numeric(
        df_medidores["Ventana_No_Completa"], errors="coerce"
    ).eq(1)
    va_a_re545 = ~va_a_ecostos

    df["Energia_Positiva"] = energia.where(energia > 0, 0.0).where(
        va_a_re545, 0.0
    )
    df["Energia_Negativa"] = energia.where(energia < 0, 0.0).where(
        va_a_re545, 0.0
    )

    df["SoC"] = df_medidores["SoC"]
    df["Copia_Ventana"] = df_medidores["Copia_Ventana"]

    # Medidores L -> RE545 T (solo esta hoja lo recibe).
    df["T"] = df_medidores["Ventana"]

    pares = [
        _buscar_cmg(dic_cmg, barra, cuarto_hora)
        for barra, cuarto_hora in zip(df["Barra"], df["Cuarto de Hora"])
    ]
    df["CMg"] = [par[0] for par in pares]
    df["R"] = [par[1] for par in pares]

    filas_con_energia = int(va_a_re545.sum())

    registrar(
        f"  Calculo RE545: {n:,} fila(s) traspasadas desde Medidores "
        f"({filas_con_energia:,} con energia; el resto la tiene "
        f"'Calculo E Costos')."
    )

    return df


def calcular_s_re545(df_re545):
    """
    Replica S ("ranking cmg") de RE545:

        =(COUNTIFS($T:$T,T4,$R:$R,">"&R4,G:G,G4)
        + COUNTIFS($T:$T,T4,$R:$R,R4,$C:$C,">"&C4,G:G,G4))/4 + 1

    O sea, dentro del grupo central (G) + ventana de valorizacion (T):
    cuantas filas tienen CMg Promedio (R) mayor, mas cuantas lo tienen
    igual pero con Hora (C) mayor; todo eso dividido por 4 y +1.

    Ojo: NO es la misma columna que el "ranking cmg" de Calculo E
    Costos (esa es la R de esa hoja, agrupa por P y ordena por CMg, no
    por CMg Promedio).
    """

    df = df_re545.reset_index(drop=True)

    grupos = {}

    for posicion, (central, ventana) in enumerate(
        zip(df["clave"], df["T"])
    ):
        clave = (
            _normaliza_valor_vba(central),
            _normaliza_valor_vba(ventana),
        )
        grupos.setdefault(clave, []).append(posicion)

    r = pd.to_numeric(df["R"], errors="coerce")
    c = pd.to_numeric(df["Hora"], errors="coerce")

    resultado = [pd.NA] * len(df)

    for posiciones in grupos.values():

        for posicion in posiciones:

            r_fila = r.iloc[posicion]
            c_fila = c.iloc[posicion]

            if pd.isna(r_fila):
                # COUNTIFS contra un blanco no cuenta nada: queda el +1.
                resultado[posicion] = 1.0
                continue

            mayores = 0
            empates = 0

            for otra in posiciones:

                r_otra = r.iloc[otra]

                if pd.isna(r_otra):
                    continue

                if r_otra > r_fila:
                    mayores += 1
                elif r_otra == r_fila:
                    c_otra = c.iloc[otra]
                    if pd.notna(c_otra) and pd.notna(c_fila) and c_otra > c_fila:
                        empates += 1

            resultado[posicion] = (mayores + empates) / 4.0 + 1.0

    return pd.Series(resultado, index=df_re545.index)


def calcular_u_v_re545(df_re545, dic_capacidad, dic_eficiencia):
    """
    Replica U ("EiniT") y V ("EalmT") de RE545:

        U = K * VLOOKUP(G, Resumen!B:J, 4, 0) * 1000
            (SoC % x "Capacidad (MWh)" x 1000 -- la 4ta columna del
             rango B:J es la Capacidad, no la Pmax; ver
             construir_dic_resumen_capacidad)

        V = -SUMIFS(J:J, T:T, T4, G:G, G4)
            * VLOOKUP(G, Resumen!B:J, 9, 0)
            (la carga total del grupo central+ventana, cambiada de
             signo, por la "Eficiencia" de esa central)

    Una central que no esta en "Resumen BESS" deja las dos en blanco
    (equivale al #N/A del VLOOKUP original).
    """

    df = df_re545

    capacidad = df["clave"].map(
        lambda valor: dic_capacidad.get(normalizar(valor), pd.NA)
    )
    eficiencia = df["clave"].map(
        lambda valor: dic_eficiencia.get(normalizar(valor), pd.NA)
    )

    soc = pd.to_numeric(df["SoC"], errors="coerce")
    capacidad_num = pd.to_numeric(capacidad, errors="coerce")
    eficiencia_num = pd.to_numeric(eficiencia, errors="coerce")

    u = soc * capacidad_num * 1000.0

    carga = pd.to_numeric(df["Energia_Negativa"], errors="coerce").fillna(0.0)
    suma_carga = carga.groupby([df["clave"], df["T"]]).transform("sum")

    v = -suma_carga * eficiencia_num

    return u, v


# ------------------------------------------------------------
# CALCULO RE545 (etapa 2): AC:AU -- reservas por subasta
#
# Tres bloques de 6 columnas con los MISMOS 6 encabezados
# (CPF(-), CSF(-), CTF(-), CPF(+), CSF(+), CTF(+)), que se
# distinguen por el titulo de grupo de la fila 2: "Subastas"
# (AC:AH), "FD" (AI:AN) y "FMA" (AO:AT). Los tres son el mismo
# SUMIFS contra Subastas, cambiando la columna que se suma:
#
#   AC4 = SUMIFS(Subastas!$O:$O, Subastas!$K:$K, $G4,
#                Subastas!$J:$J, $D4, Subastas!$B:$B, AC$3)
#   AI4 = idem sobre Subastas!$P:$P
#   AO4 = idem sobre Subastas!$Q:$Q
#
# Criterios (homologados por NOMBRE contra nuestra hoja Subastas,
# igual que en calcular_l y en los umbrales de E Costos):
#   central       -> Configuración
#   hora del mes  -> Hora_mes
#   tipo          -> Control (la columna con los CPF/CSF, la misma
#                    que ya usa construir_prorrata_sscc)
#
# PENDIENTE DE CONFIRMAR (ver BITACORA): las tres columnas que se
# suman se toman por POSICION (O, P, Q de nuestra hoja Subastas,
# que es como las escribe la macro de carga), no por nombre. Los
# nombres reales que trajo el archivo de encabezados llaman "FD" a
# O y "FMA" a P, corridos una columna respecto de los titulos de
# grupo de RE545 (que dicen Subastas/FD/FMA para O/P/Q). Es el
# mismo corrimiento de una columna que el usuario ya describio para
# el archivo de Subastas. Se eligio seguir la formula (posicion),
# no el nombre, porque la formula es la fuente primaria.
# ------------------------------------------------------------

# Los 6 encabezados que la formula usa como criterio (AC$3 y sus
# equivalentes). Van en este orden en los tres bloques.
TIPOS_RESERVA_RE545 = (
    "CPF(-)", "CSF(-)", "CTF(-)", "CPF(+)", "CSF(+)", "CTF(+)",
)

# AC:AH, AI:AN, AO:AT -- las tres claves internas de cada bloque.
_BLOQUES_RESERVA_RE545 = (
    ("AC", "AD", "AE", "AF", "AG", "AH"),
    ("AI", "AJ", "AK", "AL", "AM", "AN"),
    ("AO", "AP", "AQ", "AR", "AS", "AT"),
)


def construir_dic_reservas_subastas(df_subastas):
    """
    Arma los tres diccionarios (central, hora del mes, tipo) -> suma,
    uno por cada columna de Subastas que suman los tres bloques de
    RE545 (O, P y Q por posicion; ver el comentario de seccion).

    Un SUMIFS sin coincidencias da 0, asi que el valor por defecto de
    los tres diccionarios es 0, no blanco.
    """

    columnas = list(df_subastas.columns)

    if len(columnas) < 16:
        raise ErrorEntrada(
            f"La hoja Subastas tiene {len(columnas)} columna(s); hacen "
            f"falta al menos 16 (B:Q) para las reservas de Calculo "
            f"RE545."
        )

    # B=0 ... N=12, O=13, P=14, Q=15.
    columnas_suma = (columnas[13], columnas[14], columnas[15])

    claves = [
        (
            _normaliza_valor_vba(central),
            _normaliza_valor_vba(hora_mes),
            _normaliza_valor_vba(tipo),
        )
        for central, hora_mes, tipo in zip(
            df_subastas["Configuración"],
            df_subastas["Hora_mes"],
            df_subastas["Control"],
        )
    ]

    diccionarios = []

    for columna in columnas_suma:

        valores = pd.to_numeric(df_subastas[columna], errors="coerce")

        acumulado = {}

        for clave, valor in zip(claves, valores):
            if pd.isna(valor):
                continue
            acumulado[clave] = acumulado.get(clave, 0.0) + float(valor)

        diccionarios.append(acumulado)

    return tuple(diccionarios)


def calcular_reservas_re545(df_re545, dics_reservas):
    """
    Replica AC:AT (los tres bloques de 6 columnas) y AU:

        AU = SUMPRODUCT(AC:AH, AI:AN, AO:AT) / 4 * 1000

    o sea, la suma de los 6 productos "reserva x FD x FMA", dividida
    por 4 y por mil.
    """

    df = df_re545

    centrales = df["clave"].map(_normaliza_valor_vba)
    horas_mes = df["Hora Mes"].map(_normaliza_valor_vba)

    columnas = {}

    for bloque, dic in zip(_BLOQUES_RESERVA_RE545, dics_reservas):

        for interno, tipo in zip(bloque, TIPOS_RESERVA_RE545):

            tipo_normalizado = _normaliza_valor_vba(tipo)

            columnas[interno] = pd.Series(
                [
                    dic.get((central, hora_mes, tipo_normalizado), 0.0)
                    for central, hora_mes in zip(centrales, horas_mes)
                ],
                index=df.index,
            )

    au = pd.Series(0.0, index=df.index)

    for posicion in range(6):
        au = au + (
            columnas[_BLOQUES_RESERVA_RE545[0][posicion]]
            * columnas[_BLOQUES_RESERVA_RE545[1][posicion]]
            * columnas[_BLOQUES_RESERVA_RE545[2][posicion]]
        )

    columnas["AU"] = au / 4.0 * 1000.0

    return columnas


# ------------------------------------------------------------
# CALCULO RE545 (etapa 3): AW:BG -- resumen por central + ventana
#
# OTRA TABLA, no mas columnas de la misma: en el original tiene 288
# filas (9 centrales x 32 ventanas) contra las 26.787 del bloque
# principal, compartiendo hoja de la fila 4 para abajo. Es el mismo
# patron de "dos tablas de distinto largo en una hoja" que ya
# aparecio en FD (bloques CSF/CPF) y en Ofertas SSCC.
#
# AW (central), AX (Ventana) y AY (Oferta Completa) NO son formulas
# ni las escribe ninguna macro: en el .xlsm son constantes. AW/AX son
# el cruce central x ventana, y AY es la marca de "oferta completa"
# de esa ventana -- que en esta migracion NO hay que inventar: es la
# columna "Completa" de construir_resumen_ventana_oferta(), la misma
# que ya alimenta Medidores!T (T = 1 - Completa). Coincide el nombre,
# la clave (central+ventana), el dominio (0/1) y el sentido: RE545 se
# queda con las filas de ventana NO completa y la marca en AY.
#
# Formulas replicadas:
#   AZ = primer U del grupo (G=AW, T=AX)   [INDEX/AGGREGATE(15,6,...,1)]
#   BA = primer V del mismo grupo
#   BB = SUMIFS(AU:AU, T:T,AX, G:G,AW)
#   BC = MIN(MAX(MIN(AZ+BA, Capacidad*1000), BA), BB) * AY * (AX<>31)
#   BF = SUMIFS(BV:BV, BR:BR,AX, G:G,AW)   (BV, ver calcular_bv_re545)
#   BG = AND(AZ+BA - BF(ventana anterior) > Capacidad*1000, BF_previa<>0)*1
#
# BD ("check 1") y BE ("check 2") quedan para la etapa siguiente:
# dependen de BN y BU, que son columnas del bloque principal todavia
# sin implementar. Son columnas de CONTROL, no entran en ningun
# calculo posterior.
# ------------------------------------------------------------

NOMBRES_RESUMEN_RE545 = {
    "AW": "",
    "AX": "Ventana",
    "AY": "Oferta Completa",
    "AZ": "EiniT",
    "BA": "EalmT",
    "BB": "Total Reservas* FD *FMA",
    "BC": "Edisp_T",
    "BD": "check 1",
    "BE": "check 2",
    "BF": "Margen ultima hora",
    "BG": "flag ultima hora",
}


def calcular_bv_re545(df_re545, inicio_ventana=INICIO_VENTANA):
    """
    Replica BV ("SSCC ultima hora"):

        =IF(AND(C4 = Medidores!$S$1 - 1, AU4 <> 0), AU4, 0)

    Medidores!S1 es la hora en que arranca la ventana (el mismo dato
    que la constante INICIO_VENTANA: la formula de Medidores!L
    incrementa la ventana justo cuando la hora es igual a S1), asi
    que la condicion es "la hora anterior al inicio de la ventana",
    o sea la ultima hora de la ventana que termina.

    Se devuelve como Serie suelta, sin guardarla todavia en el
    DataFrame: BV vive en el bloque BQ:CE, que es de una etapa
    posterior, y meterla antes desordenaria las columnas de salida.
    """

    au = pd.to_numeric(df_re545["AU"], errors="coerce").fillna(0.0)
    hora = pd.to_numeric(df_re545["Hora"], errors="coerce")

    return au.where((hora == inicio_ventana - 1) & (au != 0), 0.0)


def construir_resumen_ventanas_re545(
    df_re545, resumen_ventana_oferta, dic_capacidad, registrar=print
):
    """
    Arma la tabla AW:BG (una fila por central + ventana de
    valorizacion). Ver el comentario de seccion para el detalle de
    cada columna.

    resumen_ventana_oferta: el DataFrame de
    construir_resumen_ventana_oferta() (columnas Central, Ventana T,
    Oferta, Completa) -- de ahi salen AW, AX y AY.
    dic_capacidad: central -> "Capacidad (MWh)"
    (construir_dic_resumen_capacidad) -- es el VLOOKUP con indice 4
    de BC, que NO es la Pmax.
    """

    base = (
        resumen_ventana_oferta[["Central", "Ventana T", "Completa"]]
        .copy()
        .sort_values(["Central", "Ventana T"], kind="mergesort")
        .reset_index(drop=True)
    )

    df = df_re545.reset_index(drop=True)

    u = pd.to_numeric(df["U"], errors="coerce")
    v = pd.to_numeric(df["V"], errors="coerce")
    au = pd.to_numeric(df["AU"], errors="coerce").fillna(0.0)
    bv = calcular_bv_re545(df)

    primer_u = {}
    primer_v = {}
    suma_au = {}
    suma_bv = {}

    for posicion, (central, ventana) in enumerate(zip(df["clave"], df["T"])):

        clave = (
            _normaliza_valor_vba(central),
            _normaliza_valor_vba(ventana),
        )

        # INDEX + AGGREGATE(15,6,...,1): gana la PRIMERA fila del grupo.
        if clave not in primer_u:
            primer_u[clave] = u.iloc[posicion]
            primer_v[clave] = v.iloc[posicion]

        suma_au[clave] = suma_au.get(clave, 0.0) + float(au.iloc[posicion])
        suma_bv[clave] = suma_bv.get(clave, 0.0) + float(bv.iloc[posicion])

    filas = []

    for _, fila in base.iterrows():

        central = fila["Central"]
        ventana = fila["Ventana T"]

        clave = (
            _normaliza_valor_vba(central),
            _normaliza_valor_vba(ventana),
        )

        az = primer_u.get(clave, pd.NA)
        ba = primer_v.get(clave, pd.NA)
        bb = suma_au.get(clave, 0.0)
        bf = suma_bv.get(clave, 0.0)

        completa = fila["Completa"]
        completa = 0.0 if pd.isna(completa) else float(completa)

        factor = dic_capacidad.get(normalizar(central), pd.NA)

        filas.append(
            {
                "AW": central,
                "AX": ventana,
                "AY": completa,
                "AZ": az,
                "BA": ba,
                "BB": bb,
                "BF": bf,
                "_factor": factor,
            }
        )

    tabla = pd.DataFrame(
        filas,
        columns=["AW", "AX", "AY", "AZ", "BA", "BB", "BF", "_factor"],
    )

    if tabla.empty:
        tabla = tabla.assign(BC=[], BD=[], BE=[], BG=[])
        return tabla[list(NOMBRES_RESUMEN_RE545)].rename(
            columns=NOMBRES_RESUMEN_RE545
        )

    az_num = pd.to_numeric(tabla["AZ"], errors="coerce")
    ba_num = pd.to_numeric(tabla["BA"], errors="coerce")
    factor_num = pd.to_numeric(tabla["_factor"], errors="coerce")
    ventana_num = pd.to_numeric(tabla["AX"], errors="coerce")

    capacidad = factor_num * 1000.0

    # BC = MIN(MAX(MIN(AZ+BA, Pmax*1000), BA), BB) * AY * (AX<>31)
    tabla["BC"] = (
        pd.concat(
            [
                pd.concat(
                    [
                        pd.concat([az_num + ba_num, capacidad], axis=1).min(axis=1),
                        ba_num,
                    ],
                    axis=1,
                ).max(axis=1),
                pd.to_numeric(tabla["BB"], errors="coerce"),
            ],
            axis=1,
        ).min(axis=1)
        * tabla["AY"]
        * ventana_num.ne(31).astype(float)
    )

    # BG: la resta usa el BF de la MISMA central en la ventana anterior,
    # y ademas exige que el BF de la FILA ANTERIOR de esta tabla sea
    # distinto de 0. En el original la primera fila apunta a la fila de
    # encabezados (texto), que en Excel tambien cumple "<> 0".
    bf_por_clave = {
        (_normaliza_valor_vba(c), _normaliza_valor_vba(x)): float(b)
        for c, x, b in zip(tabla["AW"], tabla["AX"], tabla["BF"])
    }

    bg = []

    for posicion in range(len(tabla)):

        central = tabla["AW"].iloc[posicion]
        ventana = ventana_num.iloc[posicion]

        bf_anterior_ventana = bf_por_clave.get(
            (
                _normaliza_valor_vba(central),
                _normaliza_valor_vba(
                    ventana - 1 if pd.notna(ventana) else ventana
                ),
            ),
            0.0,
        )

        if posicion == 0:
            bf_fila_previa_no_cero = True
        else:
            bf_fila_previa_no_cero = float(tabla["BF"].iloc[posicion - 1]) != 0.0

        suma = az_num.iloc[posicion] + ba_num.iloc[posicion]
        limite = capacidad.iloc[posicion]

        if pd.isna(suma) or pd.isna(limite):
            bg.append(0)
            continue

        bg.append(
            int(
                (suma - bf_anterior_ventana > limite)
                and bf_fila_previa_no_cero
            )
        )

    tabla["BG"] = bg

    # Dependen de BN/BU, del bloque principal (etapa siguiente).
    tabla["BD"] = pd.NA
    tabla["BE"] = pd.NA

    registrar(
        f"  Calculo RE545: {len(tabla):,} fila(s) en el resumen por "
        f"central+ventana (AW:BG); BD/BE quedan pendientes."
    )

    tabla = tabla[list(NOMBRES_RESUMEN_RE545)]

    return tabla.rename(columns=NOMBRES_RESUMEN_RE545)


# ------------------------------------------------------------
# CALCULO RE545 (etapa 4): BI:CE -- Componente 1 y Componente 2
#
# Ultimo bloque del bloque principal. Formulas (fila 4 del original):
#
#   BI "Orden"    = 1 en las 4 primeras filas; despues
#                   IF(T(i)=T(i-4), BI(i-4)+1, 1)  <- salto de 4 filas
#   BJ "Periodo"  = 0,15,30,45 en las 4 primeras; despues BJ(i-4)
#   BK            = SUMIFS(R, G=G, S=BI, T=T, E=BJ)
#   BL            = SUMIFS(Q, S=BI, E=BJ, G=G, T=T)
#   BM "Curva Cmg Decendente"
#                 = LARGE(IF(BK_todas = BK(i), BL_todas), BJ(i)/15 + 1)
#   BN "Edisp_Asig"
#                 = MAX(0, MIN(MAX(0, Pmax*1000/4 - BS),
#                              BC(central,ventana) - suma de los BN
#                              ANTERIORES del mismo (T,G)))
#   BO "Total C1_545" = BN * BM
#   BQ "Energia Total" = BS + BN
#   BR "Ventana de Valorizacion" = T
#   BS "inyeccion en el periodo del Cmg Descendente"
#                 = I de la primera fila con BR=BR(i), S=BI(i),
#                   G=G(i), E=BJ(i)
#   BT "Energía ya Asignada" = suma de los BU POSTERIORES del mismo
#                   (BR, G)   <- mira hacia adelante
#   BU "Asignacion Edisponible"
#                 = IF(BS=0, 0, MAX(0, MIN(BQ, BC(central,ventana)
#                       - BT - suma de BV del grupo)))
#   BV "SSCC ultima hora" (ver calcular_bv_re545, ya usada por BF)
#   BW "inyeccion orden cronologico" = I + J
#   BX "Energia Ultima hora" = BF del resumen (central, ventana)
#   BY "Energía ya Asignada ultima hora"
#                 = IF(BW<0, BX, suma de los BZ ANTERIORES del grupo)
#   BZ "Energia Asignada Ultima hora"
#                 = IF(BW<0, 0, 8) * BG del resumen (central, ventana)
#   CA            = BU + BZ
#   CC "Total C2_545" = CA * BM
#   CE "Monto a compensar"
#                 = MAX(suma(BO del grupo) - suma(CC del grupo), 0)
#                   * AU / suma(AU del grupo)
#
# Dos recursiones que hay que resolver en orden, no vectorizables de
# una: BN necesita los BN anteriores de su grupo (se recorre de
# arriba hacia abajo) y BU necesita los BU POSTERIORES del suyo (se
# recorre de abajo hacia arriba). BY necesita los BZ anteriores, pero
# BZ no depende de BY, asi que ahi alcanza con calcular BZ primero.
#
# OJO con el VLOOKUP de BN: usa el indice 2 del rango B:J, o sea
# "Pmax (MW)" (dic_factor) -- distinto del indice 4 ("Capacidad
# (MWh)", dic_capacidad) que usan U y BC.
# ------------------------------------------------------------

def _clave_grupo_re545(central, ventana):
    return (_normaliza_valor_vba(central), _normaliza_valor_vba(ventana))


def calcular_bi_bj_re545(df_re545):
    """
    Replica BI ("Orden") y BJ ("Periodo"), las dos con el mismo salto
    de 4 filas del original (4 bloques de 15 minutos por hora):
    BI arranca en 1 y suma 1 cada vez que la fila de 4 mas arriba
    tiene la misma ventana; BJ repite el periodo de esa misma fila.

    Las 4 primeras filas son constantes en el .xlsm (BI = 1 y BJ = 0,
    15, 30, 45). Aca se toman de la propia columna "Minutos" para no
    hardcodear una grilla de 15 minutos que el resto del codigo no
    asume en ningun lado.
    """

    df = df_re545.reset_index(drop=True)
    n = len(df)

    ventana = list(df["T"])
    minutos = list(df["Minutos"])

    bi = [1.0] * n
    bj = [None] * n

    for i in range(n):

        if i < 4:
            bi[i] = 1.0
            bj[i] = minutos[i]
            continue

        bj[i] = bj[i - 4]

        if _normaliza_valor_vba(ventana[i]) == _normaliza_valor_vba(
            ventana[i - 4]
        ):
            bi[i] = bi[i - 4] + 1.0
        else:
            bi[i] = 1.0

    return (
        pd.Series(bi, index=df_re545.index),
        pd.Series(bj, index=df_re545.index),
    )


def calcular_bk_bl_bm_bs_re545(df_re545):
    """
    Replica BK, BL, BM y BS, que comparten la misma clave
    (central + Orden + ventana + Periodo):

      BK = suma de R (CMg Promedio) de las filas con esa clave
      BL = suma de Q (CMg) de las filas con esa clave
      BM = el k-esimo valor mas grande de BL entre TODAS las filas
           (de toda la hoja, no del grupo) cuyo BK es igual al de la
           fila, con k = Periodo/15 + 1
      BS = el I (Descarga kWh) de la PRIMERA fila con esa clave
    """

    df = df_re545.reset_index(drop=True)

    r = pd.to_numeric(df["R"], errors="coerce").fillna(0.0)
    q = pd.to_numeric(df["CMg"], errors="coerce").fillna(0.0)
    i_energia = pd.to_numeric(df["Energia_Positiva"], errors="coerce")

    claves = [
        (
            _normaliza_valor_vba(central),
            _normaliza_valor_vba(orden),
            _normaliza_valor_vba(ventana),
            _normaliza_valor_vba(periodo),
        )
        for central, orden, ventana, periodo in zip(
            df["clave"], df["BI"], df["T"], df["BJ"]
        )
    ]

    suma_r = {}
    suma_q = {}
    primer_i = {}

    for posicion, clave in enumerate(claves):
        suma_r[clave] = suma_r.get(clave, 0.0) + float(r.iloc[posicion])
        suma_q[clave] = suma_q.get(clave, 0.0) + float(q.iloc[posicion])
        if clave not in primer_i:
            primer_i[clave] = i_energia.iloc[posicion]

    bk = [suma_r[clave] for clave in claves]
    bl = [suma_q[clave] for clave in claves]
    bs = [primer_i[clave] for clave in claves]

    # BM: LARGE(IF(BK = BK(i), BL), Periodo/15 + 1). El IF recorre
    # TODA la columna, no el grupo: se indexa por valor de BK.
    por_bk = {}

    for valor_bk, valor_bl in zip(bk, bl):
        por_bk.setdefault(round(float(valor_bk), 9), []).append(float(valor_bl))

    for lista in por_bk.values():
        lista.sort(reverse=True)

    bm = []

    for posicion, valor_bk in enumerate(bk):

        periodo = pd.to_numeric(
            pd.Series([df["BJ"].iloc[posicion]]), errors="coerce"
        ).iloc[0]

        if pd.isna(periodo):
            bm.append(pd.NA)
            continue

        k = int(periodo / 15) + 1
        lista = por_bk.get(round(float(valor_bk), 9), [])

        # LARGE con k fuera de rango da #NUM! -> el IFERROR lo deja "".
        bm.append(lista[k - 1] if 1 <= k <= len(lista) else pd.NA)

    indice = df_re545.index

    return (
        pd.Series(bk, index=indice),
        pd.Series(bl, index=indice),
        pd.Series(bm, index=indice),
        pd.Series(bs, index=indice),
    )


def _mapa_resumen_por_grupo(df_resumen_re545, columna):
    """
    (central, ventana) -> valor de una columna de la tabla resumen
    AW:BG, que ya viene con los nombres reales. Se toma por posicion
    porque AW no tiene encabezado en el archivo real.
    """

    centrales = df_resumen_re545.iloc[:, 0]
    ventanas = df_resumen_re545.iloc[:, 1]
    valores = pd.to_numeric(df_resumen_re545[columna], errors="coerce")

    mapa = {}

    for central, ventana, valor in zip(centrales, ventanas, valores):
        mapa[_clave_grupo_re545(central, ventana)] = (
            0.0 if pd.isna(valor) else float(valor)
        )

    return mapa


def calcular_componentes_re545(
    df_re545, df_resumen_re545, dic_factor
):
    """
    Replica BI:CE (menos BV, que ya calcula calcular_bv_re545) y
    devuelve un diccionario columna interna -> Serie. Ver el
    comentario de seccion para la formula de cada una.

    dic_factor: central -> "Pmax (MW)" (indice 2 del VLOOKUP de BN).
    df_resumen_re545: la tabla AW:BG ya construida (de ahi salen BC,
    BF y BG por central+ventana).
    """

    df = df_re545.reset_index(drop=True).copy()
    n = len(df)

    bi, bj = calcular_bi_bj_re545(df)
    df["BI"] = bi
    df["BJ"] = bj

    bk, bl, bm, bs = calcular_bk_bl_bm_bs_re545(df)
    df["BK"], df["BL"], df["BM"], df["BS"] = bk, bl, bm, bs

    br = df["T"]
    claves = [
        _clave_grupo_re545(central, ventana)
        for central, ventana in zip(df["clave"], br)
    ]

    dic_bc = _mapa_resumen_por_grupo(df_resumen_re545, "Edisp_T")
    dic_bf = _mapa_resumen_por_grupo(df_resumen_re545, "Margen ultima hora")
    dic_bg = _mapa_resumen_por_grupo(df_resumen_re545, "flag ultima hora")

    bv = calcular_bv_re545(df)

    suma_bv = {}
    for clave, valor in zip(claves, bv):
        suma_bv[clave] = suma_bv.get(clave, 0.0) + float(valor)

    pmax = [
        dic_factor.get(normalizar(central), pd.NA) for central in df["clave"]
    ]

    bs_num = pd.to_numeric(df["BS"], errors="coerce")

    # --- BN: recursion hacia ABAJO (suma de los BN anteriores) ---
    bn = [0.0] * n
    acumulado_bn = {}

    for posicion in range(n):

        clave = claves[posicion]
        capacidad_pmax = pmax[posicion]
        valor_bs = bs_num.iloc[posicion]

        if pd.isna(capacidad_pmax) or pd.isna(valor_bs):
            bn[posicion] = pd.NA
            continue

        disponible = max(
            0.0, float(capacidad_pmax) * 1000.0 / 4.0 - float(valor_bs)
        )
        techo = dic_bc.get(clave, 0.0) - acumulado_bn.get(clave, 0.0)

        valor = max(0.0, min(disponible, techo))

        bn[posicion] = valor
        acumulado_bn[clave] = acumulado_bn.get(clave, 0.0) + valor

    bn = pd.Series(bn, index=df.index)

    bm_num = pd.to_numeric(df["BM"], errors="coerce")
    bn_num = pd.to_numeric(bn, errors="coerce")

    bo = bn_num * bm_num
    bq = bs_num.fillna(0.0) + bn_num.fillna(0.0)

    # --- BU: recursion hacia ARRIBA (BT mira las filas siguientes) ---
    bt = [0.0] * n
    bu = [0.0] * n
    acumulado_bu = {}

    for posicion in range(n - 1, -1, -1):

        clave = claves[posicion]

        # BT = suma de los BU de las filas POSTERIORES del grupo.
        bt[posicion] = acumulado_bu.get(clave, 0.0)

        valor_bs = bs_num.iloc[posicion]

        if pd.isna(valor_bs) or float(valor_bs) == 0.0:
            bu[posicion] = 0.0
        else:
            techo = (
                dic_bc.get(clave, 0.0)
                - bt[posicion]
                - suma_bv.get(clave, 0.0)
            )
            bu[posicion] = max(0.0, min(float(bq.iloc[posicion]), techo))

        acumulado_bu[clave] = acumulado_bu.get(clave, 0.0) + bu[posicion]

    bt = pd.Series(bt, index=df.index)
    bu = pd.Series(bu, index=df.index)

    bw = (
        pd.to_numeric(df["Energia_Positiva"], errors="coerce").fillna(0.0)
        + pd.to_numeric(df["Energia_Negativa"], errors="coerce").fillna(0.0)
    )

    bx = pd.Series([dic_bf.get(clave, 0.0) for clave in claves], index=df.index)

    bz = pd.Series(
        [
            (0.0 if bw.iloc[posicion] < 0 else 8.0)
            * dic_bg.get(claves[posicion], 0.0)
            for posicion in range(n)
        ],
        index=df.index,
    )

    # --- BY: suma de los BZ ANTERIORES del grupo (BZ ya esta) ---
    by = [0.0] * n
    acumulado_bz = {}

    for posicion in range(n):

        clave = claves[posicion]

        if bw.iloc[posicion] < 0:
            by[posicion] = float(bx.iloc[posicion])
        else:
            by[posicion] = acumulado_bz.get(clave, 0.0)

        acumulado_bz[clave] = acumulado_bz.get(clave, 0.0) + float(
            bz.iloc[posicion]
        )

    by = pd.Series(by, index=df.index)

    ca = bu + bz
    cc = ca * bm_num

    # --- CE: por grupo ---
    au = pd.to_numeric(df["AU"], errors="coerce").fillna(0.0)

    suma_bo = {}
    suma_cc = {}
    suma_au = {}

    for posicion, clave in enumerate(claves):
        valor_bo = bo.iloc[posicion]
        valor_cc = cc.iloc[posicion]
        suma_bo[clave] = suma_bo.get(clave, 0.0) + (
            0.0 if pd.isna(valor_bo) else float(valor_bo)
        )
        suma_cc[clave] = suma_cc.get(clave, 0.0) + (
            0.0 if pd.isna(valor_cc) else float(valor_cc)
        )
        suma_au[clave] = suma_au.get(clave, 0.0) + float(au.iloc[posicion])

    ce = []

    for posicion, clave in enumerate(claves):

        total_au = suma_au.get(clave, 0.0)

        if total_au == 0.0:
            # division por cero -> el IFERROR original devuelve 0
            ce.append(0.0)
            continue

        ce.append(
            max(suma_bo.get(clave, 0.0) - suma_cc.get(clave, 0.0), 0.0)
            * float(au.iloc[posicion])
            / total_au
        )

    indice = df_re545.index

    def _serie(valores):
        return pd.Series(list(valores), index=indice)

    return {
        "BI": _serie(df["BI"]),
        "BJ": _serie(df["BJ"]),
        "BK": _serie(df["BK"]),
        "BL": _serie(df["BL"]),
        "BM": _serie(df["BM"]),
        "BN": _serie(bn),
        "BO": _serie(bo),
        "BQ": _serie(bq),
        "BR": _serie(br),
        "BS": _serie(df["BS"]),
        "BT": _serie(bt),
        "BU": _serie(bu),
        "BV": _serie(bv),
        "BW": _serie(bw),
        "BX": _serie(bx),
        "BY": _serie(by),
        "BZ": _serie(bz),
        "CA": _serie(ca),
        "CC": _serie(cc),
        "CE": _serie(ce),
    }


def completar_checks_resumen_re545(df_resumen_re545, df_re545):
    """
    Completa BD ("check 1") y BE ("check 2") de la tabla resumen, que
    dependen de BN y BU del bloque principal:

        BD = SUMIFS(BN, BR=AX, G=AW) - BC
        BE = SUMIFS(BU, BR=AX, G=AW) - BC + BF

    Son columnas de control: no alimentan ningun calculo posterior.
    """

    tabla = df_resumen_re545.copy()

    suma_bn = {}
    suma_bu = {}

    bn = pd.to_numeric(df_re545["BN"], errors="coerce").fillna(0.0)
    bu = pd.to_numeric(df_re545["BU"], errors="coerce").fillna(0.0)

    for posicion, (central, ventana) in enumerate(
        zip(df_re545["clave"], df_re545["BR"])
    ):
        clave = _clave_grupo_re545(central, ventana)
        suma_bn[clave] = suma_bn.get(clave, 0.0) + float(bn.iloc[posicion])
        suma_bu[clave] = suma_bu.get(clave, 0.0) + float(bu.iloc[posicion])

    claves = [
        _clave_grupo_re545(central, ventana)
        for central, ventana in zip(tabla.iloc[:, 0], tabla.iloc[:, 1])
    ]

    bc = pd.to_numeric(tabla["Edisp_T"], errors="coerce").fillna(0.0)
    bf = pd.to_numeric(tabla["Margen ultima hora"], errors="coerce").fillna(0.0)

    tabla["check 1"] = [
        suma_bn.get(clave, 0.0) - float(bc.iloc[posicion])
        for posicion, clave in enumerate(claves)
    ]
    tabla["check 2"] = [
        suma_bu.get(clave, 0.0) - float(bc.iloc[posicion])
        + float(bf.iloc[posicion])
        for posicion, clave in enumerate(claves)
    ]

    return tabla


def completar_calculo_re545(
    df_re545, df_subastas, umbral_soc_minimo, dic_capacidad,
    dic_eficiencia, registrar=print,
):
    """
    Agrega a la etapa base de RE545 las columnas calculadas L, M, N,
    O, S, U, V y AC:AU, en el orden final de la hoja pero todavia
    con los nombres internos (el renombre a los nombres reales lo
    hace renombrar_calculo_re545, despues de la tabla resumen).

    L y M son literalmente las mismas formulas que en Calculo E Costos
    (mismo COUNTIFS contra Subastas, mismo 1*(SoC > umbral)), asi que
    se reusan calcular_l() y calcular_m(). N y O tambien: la formula
    de RE545 (SUMIFS del grupo menos SUMIFS de los bloques anteriores)
    es la version en formula de lo mismo que calcula calcular_n_o()
    para E Costos -- suma, dentro del grupo central+ciclo y contando
    solo filas con L=1, la energia de los bloques horarios >= al de la
    fila. El resto (S, U, V) es propio de esta hoja.

    Agrega tambien AC:AU (los tres bloques de reservas por subasta y
    el SUMPRODUCT que los combina, ver la seccion de mas arriba).

    Todavia FUERA de esta etapa (ver plan seccion 26.3): AW:BG (el
    resumen por central+ventana, que es una tabla de otro largo) y
    BI:CE (Componentes 1 y 2).
    """

    df = df_re545.reset_index(drop=True).copy()

    df["L"] = calcular_l(df, df_subastas)
    df["M"] = calcular_m(df, umbral_soc_minimo)

    n, o = calcular_n_o(df)
    df["N"] = n.reset_index(drop=True)
    df["O"] = o.reset_index(drop=True)

    df["S"] = calcular_s_re545(df)

    u, v = calcular_u_v_re545(df, dic_capacidad, dic_eficiencia)
    df["U"] = u
    df["V"] = v

    dics_reservas = construir_dic_reservas_subastas(df_subastas)

    for interno, serie in calcular_reservas_re545(df, dics_reservas).items():
        df[interno] = serie

    participa = int(df["L"].sum())
    registrar(
        f"  Calculo RE545: {participa:,} de {len(df):,} fila(s) "
        f"marcadas como 'participa en subasta' (L=1)."
    )

    return df


def renombrar_calculo_re545(df_re545):
    """
    Deja las columnas en el orden final de la hoja y las pasa de los
    nombres internos (letras) a los nombres reales. Se hace al final
    de todo y aparte, porque la tabla resumen AW:BG y las columnas
    BI:CE necesitan el DataFrame con los nombres internos.
    """

    return (
        df_re545[list(NOMBRES_CALCULO_RE545)]
        .rename(columns=NOMBRES_CALCULO_RE545)
    )


# ============================================================
# CALCULO E COSTOS (etapa 2): L, M, N, O, R, S, T, U, W, X, Y, AB,
# AC, AD, AE, AF
#
# Replica esa parte de Actualizar_Calculos_Columnas (modulo
# J_Calculo_Ecostos, ver plan seccion 25.6/25.7). El usuario confirmo
# con un archivo real que la hoja "Resumen" del libro original (que
# M usa para el umbral, y AE/AF para el factor por central) es la
# MISMA tabla que Centrales.xlsx!Resumen BESS -- no hacia falta una
# hoja nueva. Ver construir_dic_resumen_factor().
#
# Quedan FUERA de esta etapa (bloqueados): AG:AZ -- dependen de la
# tabla dinamica "Prorrata SSCC" (todavia no se construye en Python,
# aunque el usuario ya confirmo su estructura: Filas: Configuración,
# Hora_mes / Columnas: Control / Valores: Cuenta de Sub_Baj) y de un
# umbral de subida/bajada por central+ventana (en el .xlsm original
# vive en Subastas!R:V o U:W segun la fuente -- la posicion exacta
# todavia no esta clara ni siquiera con el archivo de encabezados
# real, ver BITACORA) y de una categoria "CTF" en FD que no existe en
# nuestra hoja FD (que solo tiene CSF/CPF).
#
# ADVERTENCIA sobre L (parcialmente resuelta, ver BITACORA): el VBA
# original arma la clave de match contra Subastas usando columnas por
# posicion que, en el archivo de trazabilidad, no coincidian con los
# encabezados reales. El usuario confirmo que el "tipo" (BAJADA/
# SUBIDA) esta en Subastas!Sub_Baj. La central equivalente se uso
# como Subastas!Configuración, y un archivo real posterior confirmo
# que "Calculo E Costos"!G se llama literalmente "Configuracion" --
# el mismo campo en ambas hojas, lo que da bastante mas confianza en
# esta homologacion (aunque no es una confirmacion letra por letra
# del match, solo de que el NOMBRE del campo coincide en las dos
# hojas). Si al correr esto la cantidad de filas con L=1 sale
# sospechosamente baja o en cero, sigue siendo la primera sospechosa
# a revisar.
# ============================================================

def _normaliza_valor_vba(valor):
    """
    Replica NormalizarValor: texto en mayusculas y recortado. Los
    numeros se renderizan sin decimales de mas (igual que CStr en
    VBA: 7 -> "7", no "7.0"), para poder armar claves compuestas
    comparables entre hojas.
    """

    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(valor, bool):
        texto = str(valor)
    elif isinstance(valor, (int, float)):
        numero = float(valor)
        texto = str(int(numero)) if numero.is_integer() else str(numero)
    else:
        texto = str(valor)

    return texto.strip().upper()


def _columna_clave_vba(serie):
    """
    serie.map(_normaliza_valor_vba), pero forzando el resultado a
    dtype string SIEMPRE, incluso cuando serie esta vacia (0 filas).

    Trampa real (encontrada con datos reales, no en los sinteticos):
    pandas.Series.map() sobre una Series vacia es un no-op que NO
    llama a la funcion -- devuelve una Series vacia con el MISMO
    dtype que tenia antes de mapear. Si esa columna original era
    numerica (ej. "Mes"/"Hora Mes" leida como int64 desde Excel) y
    el resultado se concatena con "+" contra una Series de texto (u
    otro separador), la suma falla: numpy no sabe sumar int64 con
    texto, aunque las dos esten vacias.

    Pasa cuando el filtro previo (ej. Subastas!Sub_Baj en {BAJADA,
    SUBIDA}) no encuentra ninguna fila -- un caso real y valido (no
    hay ninguna subasta en el periodo), no un error de datos. Forzar
    .astype(str) despues del .map() corrige el dtype en los dos
    casos (vacio o no), sin cambiar ningun valor.
    """

    return serie.map(_normaliza_valor_vba).astype(str)


def _construir_set_subastas_tipo(df_subastas):
    """
    Conjunto de claves "central¦mes¦dia¦hora" que SI participaron en
    una subasta de subida o bajada (Subastas!Sub_Baj en {BAJADA,
    SUBIDA}), para la columna L.
    """

    tipo = df_subastas["Sub_Baj"].map(_normaliza_valor_vba)
    filtro = tipo.isin(["BAJADA", "SUBIDA"])

    sub = df_subastas.loc[filtro]

    claves = (
        _columna_clave_vba(sub["Configuración"])
        + "¦" + _columna_clave_vba(sub["Mes"])
        + "¦" + _columna_clave_vba(sub["Dia"])
        + "¦" + _columna_clave_vba(sub["Hora_dia"])
    )

    return set(claves)


def calcular_l(df_ecostos, df_subastas):
    """
    Replica la columna L: 1 si la central+mes+dia+hora de la fila
    existe en Subastas como registro BAJADA o SUBIDA, si no 0.
    """

    claves_subasta = _construir_set_subastas_tipo(df_subastas)

    clave_fila = (
        _columna_clave_vba(df_ecostos["clave"])
        + "¦" + _columna_clave_vba(df_ecostos["Mes"])
        + "¦" + _columna_clave_vba(df_ecostos["Dia"])
        + "¦" + _columna_clave_vba(df_ecostos["Hora"])
    )

    return clave_fila.isin(claves_subasta).astype("int64")


def calcular_n_o(df_ecostos):
    """
    Replica N y O: por grupo (central=clave, ventana=Copia_Ventana),
    ordenando por 'Cuarto de Hora' descendente, suma acumulada de I
    (N) y de -J (O), solo contando filas con L=1, y repartida a TODAS
    las filas que comparten el mismo 'Cuarto de Hora' dentro del
    grupo (no solo a las que tienen L=1).
    """

    def _por_grupo(grupo):

        valido = grupo["L"] == 1

        i_valido = grupo["Energia_Positiva"].where(valido, 0.0)
        j_valido = grupo["Energia_Negativa"].where(valido, 0.0)

        suma_i_por_f = i_valido.groupby(grupo["Cuarto de Hora"]).sum()
        suma_j_por_f = j_valido.groupby(grupo["Cuarto de Hora"]).sum()

        acumulado_i = suma_i_por_f.sort_index(ascending=False).cumsum()
        acumulado_j = suma_j_por_f.sort_index(ascending=False).cumsum()

        n = grupo["Cuarto de Hora"].map(acumulado_i)
        o = -grupo["Cuarto de Hora"].map(acumulado_j)

        return pd.DataFrame({"N": n, "O": o}, index=grupo.index)

    resultado = (
        df_ecostos
        .groupby(["clave", "Copia_Ventana"], sort=False, group_keys=False)
        .apply(_por_grupo)
    )

    return resultado["N"], resultado["O"]


def calcular_r_ecostos(df_ecostos):
    """
    Replica "Calculo E Costos"!R: ranking por grupo (central+
    ventana), ordenando por CMg descendente y 'Cuarto de Hora'
    descendente; las filas empatadas en ambos comparten el mismo
    ranking (la posicion donde empieza el empate), igual que un
    RANK() de Excel con "competition ranking" (no denso).

    Nombre con sufijo _ecostos a proposito: Medidores ya tiene su
    propia calcular_r() (Oferta_Completa_Dia, logica no relacionada)
    -- no renombrarla ni fusionarlas, son dos columnas "R" de hojas
    distintas con formulas distintas.
    """

    def _por_grupo(grupo):

        orden = grupo.sort_values(
            ["CMg", "Cuarto de Hora"],
            ascending=[False, False],
            kind="mergesort",
        )

        posiciones = pd.Series(range(1, len(orden) + 1), index=orden.index)

        posicion_min = posiciones.groupby(
            [orden["CMg"], orden["Cuarto de Hora"]]
        ).transform("min")

        return pd.DataFrame(
            {"R": posicion_min.reindex(grupo.index)}, index=grupo.index
        )

    resultado = (
        df_ecostos
        .groupby(["clave", "Copia_Ventana"], sort=False, group_keys=False)
        .apply(_por_grupo)
    )

    return resultado["R"]


def calcular_s_t_u(df_ecostos):
    """
    Replica S (=I*Q), T (=J*Q) y U (=L*(S+T)). No dependen de
    agrupar por central/ventana.
    """

    cmg_num = pd.to_numeric(df_ecostos["CMg"], errors="coerce").fillna(0.0)

    s = df_ecostos["Energia_Positiva"].fillna(0.0) * cmg_num
    t = df_ecostos["Energia_Negativa"].fillna(0.0) * cmg_num
    u = df_ecostos["L"].astype(float) * (s + t)

    return s, t, u


def calcular_w_x(df_ecostos):
    """
    Replica X (=P, copia de Copia_Ventana) y W (contador que se
    reinicia a 1 cada vez que cambia X respecto de la fila anterior,
    GLOBAL -- no por grupo). La primera fila es una excepcion fiel al
    original: W toma el valor de 'Hora' en vez de 1, y esa diferencia
    se arrastra en el resto de su bloque (W sigue siendo "contador
    que suma 1", solo que ese primer bloque no arranca en 1).
    """

    x = df_ecostos["Copia_Ventana"].reset_index(drop=True)

    cambia = x.ne(x.shift())
    bloque = cambia.cumsum()

    w = (
        x.groupby(bloque)
        .cumcount()
        .add(1)
        .astype(float)
    )

    if len(w):
        primera_hora = pd.to_numeric(
            df_ecostos["Hora"].iloc[:1], errors="coerce"
        ).fillna(0.0).iloc[0]
        offset = primera_hora - 1.0
        primer_bloque = bloque.iloc[0]
        w = w.mask(bloque == primer_bloque, w + offset)

    w.index = df_ecostos.index
    x.index = df_ecostos.index

    return w, x


def calcular_y_ab_ac_ad(df_ecostos):
    """
    Replica Y, AB (a partir de las filas con L=1 e I!=0, ordenadas
    por CMg descendente) y AC, AD (analogo con J!=0, ordenadas por
    CMg ASCENDENTE), por grupo (central+ventana). Cada fila del grupo
    (en su orden original) recibe los valores de la fila en esa
    posicion dentro del orden calificado; si el grupo tiene menos
    filas calificadas que filas totales, las posiciones sobrantes
    toman los valores de las filas NO calificadas en su orden
    original (sin ordenar).
    """

    def _por_grupo(grupo):

        l_uno = grupo["L"] == 1
        calif_i = l_uno & (grupo["Energia_Positiva"] != 0)
        calif_j = l_uno & (grupo["Energia_Negativa"] != 0)

        orden_i = list(
            grupo.loc[calif_i]
            .sort_values("CMg", ascending=False, kind="mergesort")
            .index
        )
        fuente_i = orden_i + list(grupo.index[~calif_i])
        cantidad_calif_i = len(orden_i)

        orden_j = list(
            grupo.loc[calif_j]
            .sort_values("CMg", ascending=True, kind="mergesort")
            .index
        )
        fuente_j = orden_j + list(grupo.index[~calif_j])
        cantidad_calif_j = len(orden_j)

        destino = list(grupo.index)

        y, ab, ac, ad = [], [], [], []

        for posicion, idx_origen in enumerate(fuente_i):
            y.append(grupo.at[idx_origen, "Cuarto de Hora"])
            ab.append(
                grupo.at[idx_origen, "CMg"]
                if posicion < cantidad_calif_i
                else pd.NA
            )

        for posicion, idx_origen in enumerate(fuente_j):
            ac.append(grupo.at[idx_origen, "Cuarto de Hora"])
            ad.append(
                grupo.at[idx_origen, "CMg"]
                if posicion < cantidad_calif_j
                else pd.NA
            )

        return pd.DataFrame(
            {"Y": y, "AB": ab, "AC": ac, "AD": ad}, index=destino
        )

    resultado = (
        df_ecostos
        .groupby(["clave", "Copia_Ventana"], sort=False, group_keys=False)
        .apply(_por_grupo)
    )

    return resultado["Y"], resultado["AB"], resultado["AC"], resultado["AD"]


def calcular_m(df_ecostos, umbral_soc_minimo):
    """
    Replica M ("SoC sobre el minimo"): 1 si SoC > umbral_soc_minimo
    (ver construir_dic_resumen_factor), si no 0.
    """

    soc = pd.to_numeric(df_ecostos["SoC"], errors="coerce").fillna(0.0)

    return (soc > umbral_soc_minimo).astype("int64")


def _calcular_asignacion_energia(bloque, energia_maxima, factor):
    """
    Replica CalcularAsignacionEnergia (AE/AF): distribuye
    energia_maxima en bloques de 15 minutos segun 'factor' (Pmax de
    la central) -- el bloque asigna 1 (completo) si cae dentro de la
    cantidad de bloques llenos, una fraccion al siguiente bloque si
    sobra un resto, y 0 al resto. Int() de VBA redondea hacia abajo
    incluso con numeros negativos, igual que math.floor.
    """

    cantidad_bloques = 4.0 * energia_maxima / factor / 1000.0
    parte_entera = math.floor(cantidad_bloques)
    fraccion = cantidad_bloques - parte_entera

    if bloque <= cantidad_bloques:
        proporcion = 1.0
    elif fraccion != 0 and bloque == parte_entera + 1.0:
        proporcion = fraccion
    else:
        proporcion = 0.0

    return proporcion * factor / 4.0 * 1000.0


def calcular_ae_af(df_ecostos, dic_factor):
    """
    Replica AE ("Energía descargada") y AF ("Energía cargada"):
    asigna, dentro de cada grupo (central+ventana), la energia
    maxima acumulada (N para AE, O para AF) en bloques segun el
    orden W de cada fila y un 'factor' por central (Resumen BESS!
    Pmax (MW), ver construir_dic_resumen_factor). Requiere que N, O
    y W ya esten calculados en df_ecostos.

    Si no hay factor para la central (no encontrada) o el factor es
    0 o no numerico, AE/AF quedan en blanco (pd.NA) -- equivalente a
    los #N/A / #VALOR! / #DIV/0! del original, sin fabricar un tipo
    de error de Excel en Python.
    """

    maximo_n = (
        df_ecostos.groupby(["clave", "Copia_Ventana"])["N"].transform("max")
    )
    maximo_o = (
        df_ecostos.groupby(["clave", "Copia_Ventana"])["O"].transform("max")
    )

    factor = df_ecostos["clave"].map(
        lambda valor: dic_factor.get(normalizar(valor), pd.NA)
    )

    ae, af = [], []

    for w, mn, mo, f in zip(df_ecostos["W"], maximo_n, maximo_o, factor):

        if pd.isna(f) or not isinstance(f, (int, float)) or f == 0:
            ae.append(pd.NA)
            af.append(pd.NA)
            continue

        ae.append(_calcular_asignacion_energia(w, mn, f))
        af.append(-_calcular_asignacion_energia(w, mo, f))

    return (
        pd.Series(ae, index=df_ecostos.index),
        pd.Series(af, index=df_ecostos.index),
    )


# ============================================================
# CALCULO E COSTOS (etapa 3): AG:AL (prorratas), AM:AR (FD), AS, AT,
# AU, AV
#
# El usuario confirmo que la categoria "CTF" (AI, AL, AO, AR) no
# existe en nuestros datos de FD y que en el archivo real esas
# columnas salen en 0 -- coincide exactamente con el VBA original,
# que las deja hardcodeadas en 0 (no dependen de ningun diccionario).
#
# Tambien confirmo que el archivo de Subastas que uso para armar nuestra
# hoja esta corrido una columna respecto del original (el original
# tiene una columna vacia al principio que el nuestro no tiene). Eso
# explica por que el VBA original documentaba "D"/"G,H,I,K" para L:
# aplicando ese corrimiento, esas letras coinciden exactamente con
# Sub_Baj/Configuración+Mes+Dia+Hora_dia -- lo mismo que ya se venia
# usando, ahora con una explicacion clara de la discrepancia.
#
# AW, AX y AZ quedaron fuera de ESTA etapa (dependian de una tabla
# de umbrales de subida/bajada que en su momento no se sabia donde
# vivia) y se resolvieron en la etapa 4, mas abajo.
# ============================================================

def construir_prorrata_sscc(df_subastas):
    """
    Construye la tabla dinamica "Prorrata SSCC" a partir de Subastas
    (confirmado por el usuario, NO es un archivo externo):
        Filas: Configuración, Hora_mes
        Columnas: Control
        Valores: Cuenta de Sub_Baj
    """

    tabla = (
        df_subastas
        .pivot_table(
            index=["Configuración", "Hora_mes"],
            columns="Control",
            values="Sub_Baj",
            aggfunc="count",
            fill_value=0,
        )
        .reset_index()
    )
    tabla.columns.name = None

    return tabla


def construir_dic_prorrata(tabla_prorrata, registrar=print):
    """
    Arma central+hora_mes -> (CPF, CSF) a partir de la tabla dinamica
    Prorrata SSCC (construir_prorrata_sscc()). Busca, entre las
    columnas que dejo el pivot (una por cada valor distinto de
    Subastas!Control), la que contenga "cpf" y la que contenga "csf"
    en el nombre -- INFERIDO (no confirmado letra por letra que
    Control tenga exactamente esos dos valores); si no las encuentra,
    avisa y usa 0 en vez de fallar.
    """

    columnas_valor = [
        c for c in tabla_prorrata.columns
        if c not in ("Configuración", "Hora_mes")
    ]

    columna_cpf = next(
        (c for c in columnas_valor if "cpf" in normalizar(str(c))), None
    )
    columna_csf = next(
        (c for c in columnas_valor if "csf" in normalizar(str(c))), None
    )

    if columna_cpf is None or columna_csf is None:
        registrar(
            f"  [AVISO] La columna 'Control' de Subastas no tiene valores "
            f"identificables como CPF/CSF (encontrados: {columnas_valor}); "
            f"CPF(-)/CSF(-)/CPF(+)/CSF(+) de Prorrata SSCC quedaran en 0."
        )

    dic = {}

    for _, fila in tabla_prorrata.iterrows():

        clave = (
            _normaliza_valor_vba(fila["Configuración"])
            + "¦" + _normaliza_valor_vba(fila["Hora_mes"])
        )

        cpf = float(fila[columna_cpf]) if columna_cpf is not None else 0.0
        csf = float(fila[columna_csf]) if columna_csf is not None else 0.0

        dic[clave] = (cpf, csf)

    return dic


def calcular_prorratas(df_ecostos, dic_prorrata):
    """
    Replica AG/AH (y sus duplicados AJ/AK, ver comentario de
    completar_calculo_e_costos_grupos): busca central+"Hora Mes" en
    el diccionario de construir_dic_prorrata(); si no hay match,
    ambas quedan en 0 (igual que el original, que no distingue
    "sin match" de "cero").
    """

    clave = (
        _columna_clave_vba(df_ecostos["clave"])
        + "¦" + _columna_clave_vba(df_ecostos["Hora Mes"])
    )

    valores = clave.map(lambda k: dic_prorrata.get(k, (0.0, 0.0)))

    ag = valores.map(lambda v: v[0])
    ah = valores.map(lambda v: v[1])

    return ag, ah


def construir_dic_mapeo_diccionario(diccionario):
    """
    Replica CrearDiccionarioPrimerValor(datosDiccionario, 1, 2):
    columna A -> columna B de la hoja Diccionario (header=None), la
    primera coincidencia gana. Es un mapeo DISTINTO de
    construir_homologacion() (usa toda la fila) y de
    _mapas_homologacion_fge() (usa E/F/G) -- tres lecturas distintas
    de la misma hoja Diccionario, no fusionar.
    """

    dic = {}

    for _, fila in diccionario.iterrows():

        clave = fila[0]
        if pd.isna(clave):
            continue

        clave_norm = normalizar(clave)

        if clave_norm and clave_norm not in dic:
            dic[clave_norm] = fila[1]

    return dic


def _calcular_bloque(valor):
    """Replica CalcularBloque: Int((valor-1)/4)+1."""

    return math.floor((valor - 1.0) / 4.0) + 1.0


def construir_dic_fd_bloque(df_fd, columna_id, columna_mas, columna_menos):
    """
    Arma id->(valor_mas, valor_menos) a partir de un bloque de FD
    (CSF o CPF, ya con nombres reales de columna), clave = columna_id
    normalizada, primera coincidencia gana. Replica
    CrearDiccionarioFDAL/CrearDiccionarioFDQAD.
    """

    dic = {}

    for _, fila in df_fd.iterrows():

        clave = normalizar(fila[columna_id])

        if clave and clave not in dic:
            dic[clave] = (fila[columna_mas], fila[columna_menos])

    return dic


def calcular_fd_prorrateado(df_ecostos, dic_mapeo, dic_fd_csf, dic_fd_cpf):
    """
    Replica AM, AN, AP, AQ: homologa la central ("clave") contra
    Diccionario!A->B; si no esta en el Diccionario, las 4 quedan en
    blanco (pd.NA, equivalente al #N/A del original). Si esta,
    arma una clave "bloque+central homologada" (bloque = CalcularBloque
    de Y para descarga, de AC para carga) y busca esa clave en los
    diccionarios de FD (CPF da AM/AP, CSF da AN/AQ); si no hay match
    en FD, queda en 0 (fiel al original: ahi solo se registra un
    aviso, no se propaga un error).
    """

    y_num = pd.to_numeric(df_ecostos["Y"], errors="coerce").fillna(0.0)
    ac_num = pd.to_numeric(df_ecostos["AC"], errors="coerce").fillna(0.0)

    bloque_y = y_num.map(_calcular_bloque)
    bloque_ac = ac_num.map(_calcular_bloque)

    am, an, ap, aq = [], [], [], []

    for central, by, bac in zip(df_ecostos["clave"], bloque_y, bloque_ac):

        mapeo = dic_mapeo.get(normalizar(central))

        if mapeo is None:
            am.append(pd.NA)
            an.append(pd.NA)
            ap.append(pd.NA)
            aq.append(pd.NA)
            continue

        mapeo_texto = _normaliza_valor_vba(mapeo)

        clave_y = normalizar(f"{int(by)}{mapeo_texto}")
        clave_ac = normalizar(f"{int(bac)}{mapeo_texto}")

        cpf_y = dic_fd_cpf.get(clave_y)
        csf_y = dic_fd_csf.get(clave_y)
        csf_ac = dic_fd_csf.get(clave_ac)

        ap.append(cpf_y[0] if cpf_y is not None else 0.0)
        am.append(cpf_y[1] if cpf_y is not None else 0.0)
        an.append(csf_y[1] if csf_y is not None else 0.0)
        aq.append(csf_ac[0] if csf_ac is not None else 0.0)

    return (
        pd.Series(am, index=df_ecostos.index),
        pd.Series(an, index=df_ecostos.index),
        pd.Series(ap, index=df_ecostos.index),
        pd.Series(aq, index=df_ecostos.index),
    )


def _calcular_costo_ponderado(cantidad1, cantidad2, precio1, precio2, energia):
    """Replica CalcularCostoPonderado (AS/AT)."""

    if pd.isna(energia):
        return pd.NA

    if (cantidad1 + cantidad2) > 0:

        if pd.isna(precio1) or pd.isna(precio2):
            return pd.NA

        factor = float(precio1) * cantidad1 + float(precio2) * cantidad2

    else:
        factor = 1.0

    return factor * float(energia)


def calcular_as_at(df_ecostos):
    """
    Replica AS (energia descarga con FD) y AT (energia carga con
    FD): CalcularCostoPonderado combinando las prorratas (AG/AH), el
    FD homologado (AM/AN para AS, AP/AQ para AT) y la energia
    asignada (AE para AS, AF para AT).
    """

    as_ = [
        _calcular_costo_ponderado(ag, ah, am, an, ae)
        for ag, ah, am, an, ae in zip(
            df_ecostos["AG"], df_ecostos["AH"],
            df_ecostos["AM"], df_ecostos["AN"], df_ecostos["AE"],
        )
    ]

    at = [
        _calcular_costo_ponderado(ag, ah, ap, aq, af)
        for ag, ah, ap, aq, af in zip(
            df_ecostos["AG"], df_ecostos["AH"],
            df_ecostos["AP"], df_ecostos["AQ"], df_ecostos["AF"],
        )
    ]

    return (
        pd.Series(as_, index=df_ecostos.index),
        pd.Series(at, index=df_ecostos.index),
    )


def calcular_au_av(df_ecostos):
    """
    Replica AU (ingreso descarga) y AV (costo carga): promedio de AB
    (donde AE es valido y distinto de 0, dentro del grupo central+
    ventana) multiplicado por AE, solo si la suma GLOBAL de energia
    de descarga por ventana (todas las centrales que comparten esa
    misma "Copia_Ventana"/P) supera 10; analogo para AV con AF/AD,
    umbral de suma global menor a -10.
    """

    ae = df_ecostos["AE"]
    af = df_ecostos["AF"]
    ab = df_ecostos["AB"]
    ad = df_ecostos["AD"]

    grupo = [df_ecostos["clave"], df_ecostos["Copia_Ventana"]]

    ae_valido = ae.notna() & (pd.to_numeric(ae, errors="coerce") != 0)
    af_valido = af.notna() & (pd.to_numeric(af, errors="coerce") != 0)

    ab_contable = ab.where(ae_valido & ab.notna())
    ad_contable = ad.where(af_valido & ad.notna())

    suma_ab = ab_contable.groupby(grupo).transform("sum")
    cantidad_ab = ab_contable.groupby(grupo).transform("count")
    hay_error_ae = ae.isna().groupby(grupo).transform("any")

    suma_ad = ad_contable.groupby(grupo).transform("sum")
    cantidad_ad = ad_contable.groupby(grupo).transform("count")
    hay_error_af = af.isna().groupby(grupo).transform("any")

    promedio_ab = (suma_ab / cantidad_ab.replace(0, pd.NA)).fillna(0.0)
    promedio_ad = (suma_ad / cantidad_ad.replace(0, pd.NA)).fillna(0.0)

    # sumaIP/sumaJP: suma GLOBAL por ventana (P), agrupando TODAS las
    # centrales que comparten esa Copia_Ventana -- no por central+P.
    suma_i_global = (
        df_ecostos["Energia_Positiva"]
        .groupby(df_ecostos["Copia_Ventana"])
        .transform("sum")
    )
    suma_j_global = (
        df_ecostos["Energia_Negativa"]
        .groupby(df_ecostos["Copia_Ventana"])
        .transform("sum")
    )

    condicion_au = (
        (cantidad_ab > 0) & (~hay_error_ae) & (suma_i_global > 10) & ae.notna()
    )
    condicion_av = (
        (cantidad_ad > 0) & (~hay_error_af) & (suma_j_global < -10) & af.notna()
    )

    au = (
        promedio_ab * pd.to_numeric(ae, errors="coerce").fillna(0.0)
    ).where(condicion_au, 0.0)
    av = (
        promedio_ad * pd.to_numeric(af, errors="coerce").fillna(0.0)
    ).where(condicion_av, 0.0)

    return au, av


# ============================================================
# CALCULO E COSTOS (etapa 4): Subastas!N, umbrales, AW, AX, AZ
#
# El usuario entrego el documento de trazabilidad completo
# (docs/Trazabilidad_11_PAGOS_BESS_2607_Definitivo.md), que trae las
# tres piezas que faltaban y que estaban anotadas como bloqueantes:
#
# 1. Subastas!N ("Energía SSCC"). Formula del libro original
#    (seccion 5.3 del documento):
#      =IFERROR(XLOOKUP(1,
#          ('Calculo E Costos'!$D$2:$D$50000=J3)*
#          ('Calculo E Costos'!$G$2:$G$50000=K3),
#          'Calculo E Costos'!$P$2:$P$50000,""),"")
#    O sea: NO es una energia. Es el "Ciclo de Carga del mes"
#    (Calculo E Costos!P = Copia_Ventana) de la primera fila de
#    Calculo E Costos que coincide en "Hora mes" (D) y
#    "Configuracion" (G). El nombre de la columna enga~na.
#
# 2. La tabla de umbrales de subida/bajada, que era EL bloqueante.
#    Vive en Subastas!U:W del libro original (seccion 5.3):
#      U3 = S3&"&"&T3                              (clave)
#      V3 = COUNTIFS($N:$N,$T3,$D:$D,V$2,K:K,S3)   (V$2 = "SUBIDA")
#      W3 = COUNTIFS($N:$N,$T3,$D:$D,W$2,K:K,S3)   (W$2 = "BAJADA")
#    No es un archivo externo ni una hoja aparte: se deriva de
#    Subastas + Subastas!N, igual que la Prorrata SSCC. Y la
#    "dependencia circular" que se habia anotado NO existe: N
#    depende de Calculo E Costos!P (Copia_Ventana), que viene de
#    Medidores y ya esta disponible desde la etapa base, antes de
#    cualquier columna calculada.
#
# 3. El bloque "AU, AV, AW, AX Y AZ" de Actualizar_Calculos_Columnas
#    (modulo J_Calculo_Ecostos) + CrearDiccionarioUmbralesSubastas.
#
# OJO CON LAS LETRAS: en el libro original la hoja Subastas esta
# corrida una columna respecto de la nuestra (confirmado por el
# usuario; el original tiene una primera columna vacia). Por eso las
# letras de arriba se leen asi contra NUESTRA hoja:
#     original D (tipo)      -> nuestra Sub_Baj
#     original J (hora mes)  -> nuestra Hora_mes
#     original K (central)   -> nuestra Configuración
# Es el mismo corrimiento ya aplicado en calcular_l(). N, en cambio,
# es una columna de formula a letra fija: original N = nuestra N.
# Igual que en todo el resto del proyecto, aca se homologa por
# NOMBRE de columna, no por posicion.
#
# Sigue fuera de alcance: toda la hoja "Calculo RE545". La columna
# AY del archivo real tampoco se calcula aca: no la escribe la macro
# J (no hay region de formulas para AY4:AY26787 en el documento, ver
# seccion 5.4) -- la macro salta de AX a AZ, y este codigo tambien.
# ============================================================

def calcular_subastas_energia_sscc(df_subastas, df_ecostos):
    """
    Resuelve Subastas!N ("Energía SSCC"), pendiente desde que se creo
    la hoja Subastas. Replica el XLOOKUP de arriba: para cada fila de
    Subastas, el Copia_Ventana ("Ciclo de Carga del mes") de la
    PRIMERA fila de Calculo E Costos que coincide en Hora Mes +
    central. Sin coincidencia -> "" (vacio), como el IFERROR original.

    df_ecostos: solo necesita las columnas de la etapa base ("Hora
    Mes", "clave", "Copia_Ventana") -- no depende de ninguna columna
    calculada, por eso no hay circularidad.

    La comparacion se hace con _normaliza_valor_vba en los dos lados
    (mismo criterio que calcular_l): el "=" de Excel compara valores,
    no texto crudo, asi que 7 y "7" tienen que cruzar igual.
    """

    indice = {}

    for hora_mes, central, ventana in zip(
        df_ecostos["Hora Mes"],
        df_ecostos["clave"],
        df_ecostos["Copia_Ventana"],
    ):
        clave = (
            _normaliza_valor_vba(hora_mes)
            + "¦" + _normaliza_valor_vba(central)
        )

        # XLOOKUP sin modo de busqueda: gana la primera coincidencia.
        if clave not in indice:
            indice[clave] = ventana

    claves = (
        _columna_clave_vba(df_subastas["Hora_mes"])
        + "¦" + _columna_clave_vba(df_subastas["Configuración"])
    )

    return claves.map(lambda clave: indice.get(clave, ""))


def _clave_central_ciclo(central, ciclo):
    """
    Clave "central&ciclo" de la tabla de umbrales. Replica las dos
    puntas de la homologacion, que en el original se arman distinto
    pero tienen que dar lo mismo:
      - Subastas!U = S & "&" & T (concatenacion de Excel)
      - Calculo E Costos = NormalizarValor(
            TextoSeguro(G) & "&" & TextoSeguro(P))
    _valor_clave() es lo que hace que un 3 y un 3.0 den "3" en las
    dos, igual que CStr/la concatenacion de Excel.
    """

    return _normaliza_valor_vba(
        _valor_clave(central) + "&" + _valor_clave(ciclo)
    )


def construir_dic_umbrales_subastas(df_subastas, energia_sscc):
    """
    Replica la tabla auxiliar Subastas!U:W (clave, SUBIDA, BAJADA) y
    CrearDiccionarioUmbralesSubastas: devuelve
    {"CENTRAL&CICLO": (umbral_subida, umbral_bajada)}.

    Cada umbral es el COUNTIFS del original: cuantas filas de
    Subastas tienen ese ciclo (columna "Energía SSCC"), esa central
    (Configuración) y ese tipo (Sub_Baj = SUBIDA / BAJADA).

    La tabla del libro original es una lista fija de central x ciclo
    escrita a mano; aca las combinaciones salen de los datos. Es
    equivalente: una combinacion que la lista tiene pero los datos no
    daria 0/0, y con umbral 0 ninguna fila pasa el filtro
    "W <= umbral*4" (W arranca en 1), o sea AW = 0 igual.
    """

    dic = {}

    tipos = df_subastas["Sub_Baj"].map(_normaliza_valor_vba)

    for central, ciclo, tipo in zip(
        df_subastas["Configuración"], energia_sscc, tipos
    ):
        if not _tiene_valor(ciclo):
            continue

        clave = _clave_central_ciclo(central, ciclo)

        if not clave:
            continue

        conteo = dic.setdefault(clave, [0, 0])

        if tipo == "SUBIDA":
            conteo[0] += 1
        elif tipo == "BAJADA":
            conteo[1] += 1

    return {clave: (valores[0], valores[1]) for clave, valores in dic.items()}


def calcular_aw_ax(df_ecostos, dic_umbrales):
    """
    Replica AW ("Descuento FD") y AX ("Total") del bloque
    "AU, AV, AW, AX Y AZ" de Actualizar_Calculos_Columnas.

    Por grupo (central=clave + ventana=Copia_Ventana):
      - clave de umbrales = central & "&" & ventana. Si no esta en el
        diccionario, el grupo no tiene umbrales validos y AW = 0.
      - promedioABW = promedio de AB entre las filas del grupo con
        W (Bloque ordenado) <= umbralBajada * 4.
      - promedioADW = promedio de AD entre las filas del grupo con
        W <= umbralSubida * 4.
        (Si, van cruzados: AB con el umbral de BAJADA y AD con el de
        SUBIDA. Asi esta en el VBA original, no es un tipeo.)
      - AW = (AE - AS) * promedioABW - (AF - AT) * promedioADW, y solo
        si hay umbrales validos, las dos cantidades promediadas son > 0
        y AE, AF, AS y AT son numeros validos. Si no, AW = 0.

    AX = AU + AV - AW, siempre (sin condiciones).
    """

    df = df_ecostos

    grupo = [df["clave"], df["Copia_Ventana"]]

    clave_umbral = pd.Series(
        [
            _clave_central_ciclo(central, ventana)
            for central, ventana in zip(df["clave"], df["Copia_Ventana"])
        ],
        index=df.index,
    )

    umbral_subida = clave_umbral.map(
        lambda clave: dic_umbrales.get(clave, (None, None))[0]
    )
    umbral_bajada = clave_umbral.map(
        lambda clave: dic_umbrales.get(clave, (None, None))[1]
    )

    umbrales_validos = umbral_subida.notna() & umbral_bajada.notna()

    w = pd.to_numeric(df["W"], errors="coerce")
    ab = pd.to_numeric(df["AB"], errors="coerce")
    ad = pd.to_numeric(df["AD"], errors="coerce")

    ab_contable = ab.where(
        umbrales_validos
        & (w <= pd.to_numeric(umbral_bajada, errors="coerce") * 4.0)
        & ab.notna()
    )
    ad_contable = ad.where(
        umbrales_validos
        & (w <= pd.to_numeric(umbral_subida, errors="coerce") * 4.0)
        & ad.notna()
    )

    cantidad_ab = ab_contable.groupby(grupo).transform("count")
    cantidad_ad = ad_contable.groupby(grupo).transform("count")

    promedio_ab = (
        ab_contable.groupby(grupo).transform("sum")
        / cantidad_ab.replace(0, pd.NA)
    ).fillna(0.0)
    promedio_ad = (
        ad_contable.groupby(grupo).transform("sum")
        / cantidad_ad.replace(0, pd.NA)
    ).fillna(0.0)

    ae = pd.to_numeric(df["AE"], errors="coerce")
    af = pd.to_numeric(df["AF"], errors="coerce")
    as_ = pd.to_numeric(df["AS"], errors="coerce")
    at = pd.to_numeric(df["AT"], errors="coerce")

    condicion_aw = (
        umbrales_validos
        & (cantidad_ab > 0)
        & (cantidad_ad > 0)
        & ae.notna()
        & af.notna()
        & as_.notna()
        & at.notna()
    )

    aw = (
        (ae - as_) * promedio_ab - (af - at) * promedio_ad
    ).where(condicion_aw, 0.0).fillna(0.0)

    ax = (
        pd.to_numeric(df["AU"], errors="coerce").fillna(0.0)
        + pd.to_numeric(df["AV"], errors="coerce").fillna(0.0)
        - aw
    )

    return aw, ax


def calcular_az(df_ecostos):
    """
    Replica AZ ("Monto a compensar"): por grupo (central + ventana),
    (suma de AX - suma de U) / cantidad de filas del grupo, nunca
    negativo (max(0, ...)), y el MISMO valor en todas las filas del
    grupo. U no numerica cuenta como 0 (NumeroSeguro del original).
    """

    grupo = [df_ecostos["clave"], df_ecostos["Copia_Ventana"]]

    ax = pd.to_numeric(df_ecostos["AX"], errors="coerce").fillna(0.0)
    u = pd.to_numeric(df_ecostos["U"], errors="coerce").fillna(0.0)

    suma_ax = ax.groupby(grupo).transform("sum")
    suma_u = u.groupby(grupo).transform("sum")
    cantidad = ax.groupby(grupo).transform("size")

    return ((suma_ax - suma_u) / cantidad).clip(lower=0.0)


# Nombres reales de columna de "Calculo E Costos" (confirmados por el
# usuario contra un archivo real, hoja "E COSTOS"). Se calcula todo
# con los nombres/letras internos usados hasta aca y se renombra
# recien al final, mismo criterio que NOMBRES_FD_CSF/NOMBRES_SUBASTAS.
#
# OJO: AG:AL ("Prorratas") y AM:AR ("FD") comparten los mismos 6
# nombres cortos (CPF(-), CSF(-), CTF(-), CPF(+), CSF(+), CTF(+)) --
# asi esta en el archivo real (se distinguen por un encabezado de
# grupo en las filas 1-2 que no se replica en este esquema de una
# sola fila de encabezado, igual que Subastas!Q sin nombre o el
# "Hora Mes" duplicado de FD). No es un error de tipeo. Lo mismo pasa
# con "Total": es el nombre real de U y tambien de AX (el total del
# grupo "Componente 1"). Para llegar sin ambiguedad a una de esas
# columnas hay que ir por posicion, no por nombre.
NOMBRES_CALCULO_E_COSTOS = {
    "Mes": "Mes",
    "Dia": "Dia",
    "Hora": "Hora",
    "Hora Mes": "Hora mes",
    "Minutos": "Minuto",
    "Cuarto de Hora": "Bloque horario",
    "clave": "Configuracion",
    "Barra": "Barra",
    "Energia_Positiva": "Descarga kWh",
    "Energia_Negativa": "Carga kWh",
    "SoC": "SoC %",
    "Copia_Ventana": "Ciclo de Carga del mes",
    "CMg": "CMg",
    "L": "Adj SSCC",
    "M": "SoC sobre el minimo",
    "N": "Energía SSCC (-) por remunerar",
    "O": "Energía SSCC (+) por remunerar",
    "R": "ranking cmg",
    "S": "Valorizacion Descarga",
    "T": "Valorizacion Carga",
    "U": "Total",
    "W": "Bloque ordenado",
    "X": "Ciclo",
    "Y": "Bloque Mes Descarga",
    "AB": "Curva monotona CMg Descarga",
    "AC": "Bloque Mes  Carga",
    "AD": "Curva monotona CMg Carga",
    "AE": "Energía descargada",
    "AF": "Energía cargada",
    "AG": "CPF(-)",
    "AH": "CSF(-)",
    "AI": "CTF(-)",
    "AJ": "CPF(+)",
    "AK": "CSF(+)",
    "AL": "CTF(+)",
    "AM": "CPF(-)",
    "AN": "CSF(-)",
    "AO": "CTF(-)",
    "AP": "CPF(+)",
    "AQ": "CSF(+)",
    "AR": "CTF(+)",
    "AS": "Energía descarga con FD",
    "AT": "Energía carga con FD",
    "AU": "Ingreso descarga",
    "AV": "Costo carga",
    "AW": "Descuento FD",
    "AX": "Total",
    "AZ": "Monto a compensar",
}


def completar_calculo_e_costos_grupos(
    df_ecostos,
    df_subastas,
    dic_factor,
    umbral_soc_minimo,
    diccionario,
    df_fd_csf,
    df_fd_cpf,
    registrar=print,
):
    """
    Etapas 2, 3 y 4 de "Calculo E Costos": agrega L, M, N, O, R, S,
    T, U, W, X, Y, AB, AC, AD, AE, AF, AG, AH, AI, AJ, AK, AL, AM,
    AN, AO, AP, AQ, AR, AS, AT, AU, AV, AW, AX y AZ a df_ecostos
    (que ya viene con la etapa base
    de construir_calculo_e_costos), y renombra todas las columnas a
    sus nombres reales (NOMBRES_CALCULO_E_COSTOS) antes de devolver.
    Ver los comentarios de seccion mas arriba para el detalle y las
    advertencias de cada columna.

    dic_factor, umbral_soc_minimo: de construir_dic_resumen_factor().
    diccionario: hoja Diccionario de Centrales.xlsx (header=None).
    df_fd_csf, df_fd_cpf: de construir_fd() (nombres reales ya
    aplicados).

    Fuera de esta funcion: la columna AY del archivo real (que la
    macro original tampoco escribe) y toda la hoja Calculo RE545.
    """

    df = df_ecostos.reset_index(drop=True).copy()

    df["L"] = calcular_l(df, df_subastas)
    df["M"] = calcular_m(df, umbral_soc_minimo)

    n, o = calcular_n_o(df)
    df["N"] = n.reset_index(drop=True)
    df["O"] = o.reset_index(drop=True)

    df["R"] = calcular_r_ecostos(df).reset_index(drop=True)

    s, t, u = calcular_s_t_u(df)
    df["S"] = s
    df["T"] = t
    df["U"] = u

    w, x = calcular_w_x(df)
    df["W"] = w
    df["X"] = x

    y, ab, ac, ad = calcular_y_ab_ac_ad(df)
    df["Y"] = y.reset_index(drop=True)
    df["AB"] = ab.reset_index(drop=True)
    df["AC"] = ac.reset_index(drop=True)
    df["AD"] = ad.reset_index(drop=True)

    ae, af = calcular_ae_af(df, dic_factor)
    df["AE"] = ae
    df["AF"] = af

    tabla_prorrata = construir_prorrata_sscc(df_subastas)
    dic_prorrata = construir_dic_prorrata(tabla_prorrata, registrar=registrar)

    ag, ah = calcular_prorratas(df, dic_prorrata)
    df["AG"] = ag
    df["AH"] = ah
    df["AI"] = 0.0
    df["AJ"] = df["AG"]
    df["AK"] = df["AH"]
    df["AL"] = 0.0

    dic_mapeo = construir_dic_mapeo_diccionario(diccionario)
    dic_fd_csf = construir_dic_fd_bloque(df_fd_csf, "id", "CSF(+)", "CSF(-)")
    dic_fd_cpf = construir_dic_fd_bloque(df_fd_cpf, "id", "CPF(+)", "CPF(-)")

    am, an, ap, aq = calcular_fd_prorrateado(
        df, dic_mapeo, dic_fd_csf, dic_fd_cpf
    )
    df["AM"] = am
    df["AN"] = an
    df["AO"] = 0.0
    df["AP"] = ap
    df["AQ"] = aq
    df["AR"] = 0.0

    as_, at = calcular_as_at(df)
    df["AS"] = as_
    df["AT"] = at

    au, av = calcular_au_av(df)
    df["AU"] = au
    df["AV"] = av

    energia_sscc = calcular_subastas_energia_sscc(df_subastas, df)
    dic_umbrales = construir_dic_umbrales_subastas(df_subastas, energia_sscc)

    con_ciclo = int(energia_sscc.map(_tiene_valor).sum())
    registrar(
        f"  Subastas: {con_ciclo:,} de {len(df_subastas):,} fila(s) "
        f"cruzaron con un ciclo de Calculo E Costos ('Energía SSCC'); "
        f"{len(dic_umbrales):,} par(es) central+ciclo con umbral "
        f"SUBIDA/BAJADA."
    )

    aw, ax = calcular_aw_ax(df, dic_umbrales)
    df["AW"] = aw
    df["AX"] = ax

    df["AZ"] = calcular_az(df)

    participa = int(df["L"].sum())
    registrar(
        f"  Calculo E Costos: {participa:,} de {len(df):,} fila(s) "
        f"marcadas como 'participa en subasta' (L=1)."
    )

    return df.rename(columns=NOMBRES_CALCULO_E_COSTOS)


# ============================================================
# COLUMNAS CALCULADAS
# ============================================================

def calcular_ventana(clave, hora, inicio_ventana=INICIO_VENTANA):
    """
    Replica Medidores!L:

        =IF(G4<>G3, 0,
            IF(C4=C3, 0, IF(C4=$S$1, 1, 0)) + L3)

    Es un contador acumulado que:
      - suma 1 cuando la hora cambia y pasa a ser la de inicio;
      - se reinicia a 0 cuando cambia la central.

    El reinicio es por BLOQUE de filas consecutivas con la misma
    clave, no por groupby, para ser fiel a la comparacion fila a
    fila que hace Excel.
    """

    clave = pd.Series(clave).reset_index(drop=True)
    hora = pd.Series(hora).reset_index(drop=True)

    cambia_clave = clave.ne(clave.shift())
    cambia_hora = hora.ne(hora.shift())

    incremento = (
        cambia_hora & hora.eq(inicio_ventana)
    ).astype(int)

    # En la primera fila de cada bloque L vale 0 sin importar
    # el incremento, porque la formula corta antes de sumarlo.
    incremento = incremento.where(~cambia_clave, 0)

    bloque = cambia_clave.cumsum()

    return (
        incremento
        .groupby(bloque)
        .cumsum()
        .astype("int64")
    )


def calcular_indicador_soc(soc, umbral=UMBRAL_SOC):
    """
    Replica Medidores!O:  =1*(J3>6%)

    Un SoC vacio da 0, igual que en Excel donde una celda vacia
    no supera el umbral.
    """

    return (
        pd.Series(soc)
        .fillna(0)
        .gt(umbral)
        .astype("int64")
        .reset_index(drop=True)
    )


def calcular_clave_auxiliar(dia, hora_mes):
    """
    Replica Medidores!N:  =B3&"&"&E3

    Excel concatena como texto. Se reproduce el formato entero
    para que no aparezcan '.0' que el Excel no tiene.
    """

    def texto(serie):
        serie = pd.Series(serie).reset_index(drop=True)
        numerica = pd.to_numeric(serie, errors="coerce")
        return numerica.map(
            lambda v: (
                ""
                if pd.isna(v)
                else (
                    str(int(v))
                    if float(v).is_integer()
                    else str(v)
                )
            )
        )

    return texto(dia) + "&" + texto(hora_mes)


# ============================================================
# CONSTRUCCION DE LA HOJA
# ============================================================

def construir_medidores(
    df_sae,
    df_soc,
    anio,
    mes,
    ruta_ofertas,
    diccionario,
    registrar=print,
):
    """
    Arma la tabla equivalente a Medidores, incluyendo las columnas R,
    S y T (dependientes de Ofertas SSCC).

    ruta_ofertas: archivo *OfertasSSCC* encontrado en Ofertas/.
    diccionario:  hoja Diccionario de Centrales.xlsx (header=None), la
                  misma que se usa para homologar el SoC.

    Devuelve (df_medidores, avisos, df_wxy, df_resumen_ventana). Estas
    ultimas dos son las tablas auxiliares equivalentes a Medidores!W:Y
    y a Medidores!AB:AE respectivamente (ver comentario de
    LETRA_A_CAMPO). El resumen intermedio equivalente a la hoja
    "Resumen Ofertas SSCC" (Nombre/Año/Mes/Día/servicios/Oferta
    completa) es puramente auxiliar para calcular df_wxy: no se
    devuelve ni se persiste, solo sirve como paso intermedio.
    """

    avisos = []

    df = df_sae.copy()

    # --------------------------------------------------------
    # ORDEN
    # --------------------------------------------------------

    # Excel trae los datos ordenados por central y tiempo, y las
    # columnas recursivas dependen de ese orden.
    df = (
        df.sort_values(["clave", "intervalo"])
        .reset_index(drop=True)
    )

    # --------------------------------------------------------
    # VALIDACION DE PERIODO
    # --------------------------------------------------------

    meses = sorted(df["Mes"].dropna().unique().tolist())

    if meses != [mes]:
        avisos.append(
            f"El periodo del SOC indica mes {mes} pero "
            f"{ARCHIVO_MEDIDAS_SAE} contiene {meses}."
        )

    # --------------------------------------------------------
    # J: CRUCE DEL SoC
    # --------------------------------------------------------

    soc = df_soc.copy()

    soc["timestamp"] = pd.to_datetime(soc["timestamp"])

    soc = (
        soc
        .dropna(subset=["timestamp"])
        .drop_duplicates(
            subset=["central", "timestamp"],
            keep="last",
        )
    )

    df = df.merge(
        soc[["central", "timestamp", "soc"]].rename(
            columns={
                "central": "clave",
                "timestamp": "intervalo",
                "soc": "SoC",
            }
        ),
        on=["clave", "intervalo"],
        how="left",
    )

    # Reporte de cobertura
    centrales_sae = set(df["clave"].unique())
    centrales_soc = set(soc["central"].unique())

    sin_soc = sorted(centrales_sae - centrales_soc)
    sobrantes = sorted(centrales_soc - centrales_sae)

    if sin_soc:
        avisos.append(
            f"Centrales en {ARCHIVO_MEDIDAS_SAE} sin bloque de "
            f"SoC: {sin_soc}"
        )

    if sobrantes:
        avisos.append(
            f"Centrales en el SOC que no estan en "
            f"{ARCHIVO_MEDIDAS_SAE}: {sobrantes}"
        )

    for central in sorted(centrales_sae & centrales_soc):
        sub = df[df["clave"] == central]
        faltan = int(sub["SoC"].isna().sum())
        if faltan:
            avisos.append(
                f"{central}: {faltan:,} de {len(sub):,} "
                f"intervalos sin SoC."
            )

    # --------------------------------------------------------
    # COLUMNAS CALCULADAS
    # --------------------------------------------------------

    df["Ventana"] = calcular_ventana(
        df["clave"],
        df["Hora"],
    )

    # K es copia de L fila a fila (plan, seccion 16.3).
    df["Copia_Ventana"] = df["Ventana"]

    df["Clave_Dia_HoraMes"] = calcular_clave_auxiliar(
        df["Dia"],
        df["Hora Mes"],
    )

    df["Indicador_SoC"] = calcular_indicador_soc(df["SoC"])

    # --------------------------------------------------------
    # COLUMNAS DELIBERADAMENTE VACIAS (diseño confirmado, no pendiente)
    # --------------------------------------------------------

    for columna in COLUMNAS_VACIAS:
        df[columna] = pd.NA

    # --------------------------------------------------------
    # OFERTAS SSCC: R, S, T (plan seccion 17-18, obligatorias)
    # --------------------------------------------------------

    registrar(f"  Leyendo {Path(ruta_ofertas).name}...")
    df_resumen_ofertas = construir_resumen_ofertas_sscc(
        ruta_ofertas, registrar=registrar
    )

    df_wxy, periodo_ofertas, avisos_wxy = cargar_resumen_en_medidores(
        df_resumen_ofertas,
        df["clave"].unique(),
        diccionario,
        registrar=registrar,
    )
    avisos.extend(avisos_wxy)

    if periodo_ofertas != (anio, mes):
        avisos.append(
            f"El resumen de Ofertas SSCC indica el periodo "
            f"{periodo_ofertas[0]}-{periodo_ofertas[1]:02d}, pero el "
            f"caso corresponde a {anio}-{mes:02d}."
        )

    r_valor, avisos_r = calcular_r(
        df, df_wxy, diccionario, registrar=registrar
    )
    avisos.extend(avisos_r)
    df["Oferta_Completa_Dia"] = r_valor

    df["Indicador_Ventana_Oferta"] = calcular_s(
        df["Ventana"], df["Oferta_Completa_Dia"]
    )

    df_resumen_ventana = construir_resumen_ventana_oferta(
        df["clave"],
        df["Ventana"],
        df["Oferta_Completa_Dia"],
        registrar=registrar,
    )

    t_valor, sin_match_t = calcular_t(
        df["clave"], df["Ventana"], df_resumen_ventana
    )
    df["Ventana_No_Completa"] = t_valor

    if sin_match_t:
        avisos.append(
            f"{sin_match_t:,} fila(s) de Medidores no encontraron su "
            "grupo central+ventana en el resumen de Ofertas SSCC al "
            "calcular la columna T."
        )

    # --------------------------------------------------------
    # ORDEN FINAL DE COLUMNAS (A -> U, en el orden de insercion de
    # LETRA_A_CAMPO)
    # --------------------------------------------------------

    df = df[list(LETRA_A_CAMPO.values())]

    registrar(
        f"Medidores construido: {len(df):,} filas x "
        f"{len(df.columns)} columnas"
    )

    return df, avisos, df_wxy, df_resumen_ventana


HOJA_OFERTAS_SSCC = "Ofertas SSCC"


def _escribir_tabla_con_titulo(
    writer, hoja, df, titulo, fila_inicio=0, columna_inicio=0
):
    """
    Escribe 'titulo' en una fila y 'df' (con su propio encabezado)
    justo debajo, dentro de la hoja 'hoja', empezando en fila_inicio y
    columna_inicio.

    Devuelve (fila_siguiente, columna_siguiente): donde empezaria el
    proximo bloque si se apila debajo, o si se pone al lado,
    respectivamente (cada llamada solo usa el que necesite).
    """

    pd.DataFrame([[titulo]]).to_excel(
        writer,
        sheet_name=hoja,
        index=False,
        header=False,
        startrow=fila_inicio,
        startcol=columna_inicio,
    )

    df.to_excel(
        writer,
        sheet_name=hoja,
        index=False,
        startrow=fila_inicio + 1,
        startcol=columna_inicio,
    )

    celda_titulo = writer.sheets[hoja].cell(
        row=fila_inicio + 1, column=columna_inicio + 1
    )
    celda_titulo.font = celda_titulo.font.copy(bold=True)

    # Debajo: +1 titulo, +1 encabezado de df, +len(df) filas, +2 de separacion.
    fila_siguiente = fila_inicio + len(df) + 4

    # Al lado: +len(df.columns) del bloque, +2 columnas de separacion.
    columna_siguiente = columna_inicio + len(df.columns) + 2

    return fila_siguiente, columna_siguiente


# Columna Q en indice 0 (A=0): usada para ubicar el bloque CPF de FD
# a la derecha del bloque CSF, en la misma hoja.
_COLUMNA_Q_INDICE = 16


def _copiar_hoja_existente(wb_origen, nombre_hoja, wb_destino):
    """
    Copia una hoja (solo valores, sin formulas ni formato) de un
    workbook openpyxl a otro. La usa escribir_salida() para preservar
    una hoja que el usuario decidio NO regenerar en la ventana
    "Generar" (ver hojas_regenerar). Devuelve False si wb_origen es
    None o no tiene esa hoja (no hay nada que preservar).
    """

    if wb_origen is None or nombre_hoja not in wb_origen.sheetnames:
        return False

    hoja_o = wb_origen[nombre_hoja]
    hoja_d = wb_destino.create_sheet(title=nombre_hoja)

    for fila in hoja_o.iter_rows():
        for celda in fila:
            hoja_d.cell(
                row=celda.row, column=celda.column, value=celda.value
            )

    for letra, dim in hoja_o.column_dimensions.items():
        if dim.width:
            hoja_d.column_dimensions[letra].width = dim.width

    return True


_HOJAS_CONSOLIDADO = ("Medidores", "Ofertas SSCC", "CMg", "FD", "Subastas")


def escribir_salida(
    df,
    ruta_salida,
    avisos,
    incidencias,
    df_wxy=None,
    df_resumen_ventana=None,
    df_cmg=None,
    df_fd_csf=None,
    df_fd_cpf=None,
    df_subastas=None,
    ruta_existente=None,
    hojas_regenerar=None,
    registrar=print,
):
    """
    Escribe Consolidado_entradas.xlsx: la tabla Medidores (A:U, una
    fila por registro), las tablas auxiliares de Ofertas SSCC -de
    otro largo, ver comentario de LETRA_A_CAMPO- juntas en una misma
    hoja (HOJA_OFERTAS_SSCC, una debajo de la otra), CMg, FD (el
    bloque CSF y el bloque CPF lado a lado, de distinto largo cada
    uno - ver construir_fd), Subastas, y un Log.

    hojas_regenerar: None (por defecto) regenera las 5 hojas de datos
    con lo que se haya pasado. Si es un set con algunos nombres de
    _HOJAS_CONSOLIDADO, las que NO esten en el set se copian tal cual
    desde ruta_existente en vez de recalcularse -- lo usa
    generar_consolidado() cuando el usuario destilda una entrada en
    la ventana "Generar". Si una hoja a preservar no existe en
    ruta_existente, queda vacia y se registra un aviso (en el log de
    esta corrida y como fila del Log).
    """

    ruta_salida = Path(ruta_salida)

    regenerar = (
        set(_HOJAS_CONSOLIDADO)
        if hojas_regenerar is None
        else set(hojas_regenerar)
    )

    wb_existente = None
    if hojas_regenerar is not None and ruta_existente is not None:
        ruta_existente = Path(ruta_existente)
        if ruta_existente.is_file():
            wb_existente = openpyxl.load_workbook(
                ruta_existente, data_only=True
            )

    avisos_preservacion = []

    def _preservar_o_avisar(writer, nombre_hoja):
        if _copiar_hoja_existente(wb_existente, nombre_hoja, writer.book):
            return
        pd.DataFrame().to_excel(writer, sheet_name=nombre_hoja, index=False)
        mensaje = (
            f"No se regenero la hoja '{nombre_hoja}' (entrada no "
            f"tildada) y no se encontro una version anterior para "
            f"preservarla; quedo vacia."
        )
        avisos_preservacion.append(mensaje)
        registrar(f"  [AVISO] {mensaje}")

    registros = (
        [("aviso", a) for a in avisos]
        + [("aviso", a) for a in avisos_preservacion]
        + [("incidencia_soc", i) for i in incidencias]
    )

    df_log = pd.DataFrame(
        registros or [("ok", "Sin observaciones.")],
        columns=["tipo", "detalle"],
    )

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:

        if "Medidores" in regenerar:
            df.to_excel(
                writer,
                sheet_name="Medidores",
                index=False,
            )
        else:
            _preservar_o_avisar(writer, "Medidores")

        if "Ofertas SSCC" in regenerar:

            columna = 0

            if df_wxy is not None:
                _, columna = _escribir_tabla_con_titulo(
                    writer,
                    HOJA_OFERTAS_SSCC,
                    df_wxy,
                    "Ofertas SSCC por dia (equivalente a Medidores!W:Y)",
                    columna_inicio=columna,
                )

            if df_resumen_ventana is not None:
                _escribir_tabla_con_titulo(
                    writer,
                    HOJA_OFERTAS_SSCC,
                    df_resumen_ventana,
                    "Resumen ventana oferta (equivalente a Medidores!AB:AE)",
                    columna_inicio=columna,
                )
        else:
            _preservar_o_avisar(writer, HOJA_OFERTAS_SSCC)

        if "CMg" in regenerar:
            if df_cmg is not None:
                df_cmg.to_excel(
                    writer,
                    sheet_name="CMg",
                    index=False,
                )
        else:
            _preservar_o_avisar(writer, "CMg")

        if "FD" in regenerar:

            if df_fd_csf is not None:
                df_fd_csf.to_excel(
                    writer,
                    sheet_name="FD",
                    index=False,
                    startcol=0,
                )

            if df_fd_cpf is not None:
                df_fd_cpf.to_excel(
                    writer,
                    sheet_name="FD",
                    index=False,
                    startcol=_COLUMNA_Q_INDICE,
                )
        else:
            _preservar_o_avisar(writer, "FD")

        if "Subastas" in regenerar:
            if df_subastas is not None:
                df_subastas.to_excel(
                    writer,
                    sheet_name="Subastas",
                    index=False,
                )
        else:
            _preservar_o_avisar(writer, "Subastas")

        df_log.to_excel(
            writer,
            sheet_name="Log",
            index=False,
        )

    return ruta_salida


# ============================================================
# PROCESO COMPLETO
#
# Dos salidas independientes, cada una con su ventana "Generar" en
# Balance_BESS.py:
#
#   - generar_consolidado(): Consolidado_entradas.xlsx. El usuario
#     tilda que "secciones" quiere recalcular esta vez; el resto se
#     preserva tal cual estaba (ver escribir_salida/hojas_regenerar).
#   - generar_pagos_bess(): Pagos_BESS.xlsx. Por ahora sin checkboxes
#     (una sola hoja) -- lee Medidores de Consolidado_entradas.xlsx
#     ya generado, no lo recalcula.
#
# SECCIONES_CONSOLIDADO agrupa los 4 checkboxes de esa ventana con
# las hojas que produce cada uno. "medidores" junta Medidas_SAE, SoC,
# Centrales (Diccionario) y OfertasSSCC porque construir_medidores()
# necesita los 4 juntos: no se pueden tildar por separado a ese nivel
# de detalle sin recalcular con datos parcialmente viejos.
# ============================================================

SECCIONES_CONSOLIDADO = (
    (
        "medidores",
        "Medidores + Ofertas SSCC",
        f"Usa {ARCHIVO_MEDIDAS_SAE}, el SoC del periodo, "
        f"{ARCHIVO_CENTRALES} (hoja Diccionario) y el archivo "
        f"OfertasSSCC -- los 4 se leen juntos para armar estas dos "
        f"hojas, no se pueden actualizar por separado.",
        ("Medidores", "Ofertas SSCC"),
    ),
    (
        "cmg",
        "CMg",
        f"Usa {ARCHIVO_CMG}.",
        ("CMg",),
    ),
    (
        "fd",
        "FD",
        f"Usa el archivo {CARPETA_SSCC_DESEMPENO}/ (hojas CPF/CSF "
        f"Horario).",
        ("FD",),
    ),
    (
        "subastas",
        "Subastas",
        f"Usa el archivo {CARPETA_SUBASTAS}/.",
        ("Subastas",),
    ),
)


SECCIONES_PAGOS = (
    (
        "ecostos",
        "Calculo E Costos",
        f"Usa las hojas 'Medidores' y 'Subastas' de {ARCHIVO_SALIDA}, "
        f"{ARCHIVO_CENTRALES}, {ARCHIVO_CMG} y el archivo "
        f"{CARPETA_SSCC_DESEMPENO}/ (para el FD homologado de AM:AR).",
        (HOJA_CALCULO_ECOSTOS,),
    ),
    (
        "re545",
        "Calculo RE545",
        f"Usa las hojas 'Medidores' y 'Subastas' de {ARCHIVO_SALIDA}, "
        f"{ARCHIVO_CENTRALES} y {ARCHIVO_CMG} -- no necesita el "
        f"archivo {CARPETA_SSCC_DESEMPENO}/.",
        (HOJA_CALCULO_RE545,),
    ),
)


def generar_consolidado(
    carpeta_base, aamm, secciones_activas, registrar=print, progreso=None
):
    """
    Genera/actualiza Consolidado_entradas.xlsx, recalculando solo las
    hojas de las secciones tildadas (ids de SECCIONES_CONSOLIDADO) y
    preservando el resto tal cual estaba en el archivo existente (ver
    escribir_salida). La usa la ventana "Generar" de esa fila.

    secciones_activas: iterable de ids de SECCIONES_CONSOLIDADO
    ("medidores", "cmg", "fd", "subastas") a recalcular esta vez.
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

    if "medidores" in secciones_activas:

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
        archivo_subastas = buscar_archivo_subastas(rutas["subastas_dir"])
        if not archivo_subastas:
            raise ErrorEntrada(
                f"No se encontro ningun archivo "
                f"3_REMUNERACIÓN_SUBASTAS_E_ID_* en "
                f"{rutas['subastas_dir']}"
            )
        registrar(f"Leyendo {archivo_subastas.name}...")
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
    las secciones tildadas (ids de SECCIONES_PAGOS: "ecostos",
    "re545") y preservando el resto tal cual estaba en el archivo
    existente (ver escribir_pagos_bess/hojas_regenerar) -- mismo
    criterio que generar_consolidado()/SECCIONES_CONSOLIDADO. La usa
    la ventana "Generar" de esa fila.

    No recalcula Medidores ni Subastas: los lee tal cual estan en
    Consolidado_entradas.xlsx, que debe generarse primero con su
    propia ventana "Generar". Centrales.xlsx y cmg.xlsx si se leen/
    recalculan frescos. El archivo SSCC_Desempeño_* solo se exige si
    "ecostos" esta tildada -- "re545" no usa FD.
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
            df_re545_base, df_resumen_re545, dic_factor
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
