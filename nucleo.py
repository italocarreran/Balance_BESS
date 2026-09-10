# -*- coding: utf-8 -*-
"""
Nucleo de calculo de la etapa Medidores.

Sin dependencias de interfaz: todo lo de aca se puede correr y
testear sin abrir la ventana.
"""

import calendar
import re
import unicodedata
from pathlib import Path

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

    resumen = pd.read_excel(
        ruta,
        sheet_name=buscar_hoja(HOJA_RESUMEN_BESS),
    )

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

        canonico = mapa_homologacion.get(
            normalizar(nombre_origen),
            nombre_origen,
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

    return df[list("ABCDEFGHIJKLM")]


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

    return df[list("QRSTUVWXYZ") + ["AA", "AB", "AC", "AD", "AE"]]


def construir_fd(ruta_sscc, registrar=print):
    """
    Replica Cargar_SSCC_Desempeno_En_FD.

    Lee, del archivo SSCC_Desempeño_*, las hojas "CPF Horario" y "CSF
    Horario" desde la fila 12, filtra por BESS/SAE en la columna D de
    cada una, y arma dos bloques independientes (distinto largo cada
    uno, igual que en la planilla): A:M (desde CSF) y Q:AE (desde
    CPF). N:P quedan fuera de alcance (la macro no las toca).

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

def construir_subastas(ruta_subastas, registrar=print):
    """
    Replica Cargar_Remuneracion_Subastas_Rapido.

    La macro original consulta la hoja "DB" del archivo de origen por
    ADO/SQL (equivalente a filtrar y seleccionar columnas de una
    tabla); aca se lee directamente con pandas y se aplica el mismo
    filtro y la misma seleccion de columnas.

    Arma las columnas B:Q de Subastas:
      - B:L: copia directa de DB!B:L (filtrado por DB!K contiene
        BESS/SAE).
      - M: formula (=K&H&I).
      - N: PENDIENTE - depende de la hoja "Calculo E Costos", una
        etapa posterior que todavia no se implementa (no se adivina).
      - O, P, Q: copias de DB!P, DB!Y, DB!V respectivamente (asi lo
        indica la macro original).
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

    registrar(
        f"  Subastas: {len(df):,} fila(s) (filtro DB!K contiene "
        f"BESS/SAE)"
    )

    return df


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


def _escribir_tabla_con_titulo(writer, hoja, df, titulo, fila_inicio):
    """
    Escribe 'titulo' en una fila y 'df' (con su propio encabezado)
    justo debajo, dentro de la hoja 'hoja', empezando en fila_inicio.
    Devuelve la fila donde debería empezar el siguiente bloque.
    """

    pd.DataFrame([[titulo]]).to_excel(
        writer,
        sheet_name=hoja,
        index=False,
        header=False,
        startrow=fila_inicio,
    )

    df.to_excel(
        writer,
        sheet_name=hoja,
        index=False,
        startrow=fila_inicio + 1,
    )

    celda_titulo = writer.sheets[hoja].cell(
        row=fila_inicio + 1, column=1
    )
    celda_titulo.font = celda_titulo.font.copy(bold=True)

    # +1 titulo, +1 encabezado de df, +len(df) filas, +2 de separacion
    return fila_inicio + len(df) + 4


# Columna Q en indice 0 (A=0): usada para ubicar el bloque CPF de FD
# a la derecha del bloque CSF, en la misma hoja.
_COLUMNA_Q_INDICE = 16


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
):
    """
    Escribe Consolidado_entradas.xlsx: la tabla Medidores (A:U, una
    fila por registro), las tablas auxiliares de Ofertas SSCC -de
    otro largo, ver comentario de LETRA_A_CAMPO- juntas en una misma
    hoja (HOJA_OFERTAS_SSCC, una debajo de la otra), CMg, FD (el
    bloque CSF y el bloque CPF lado a lado, de distinto largo cada
    uno - ver construir_fd), Subastas, y un Log.
    """

    ruta_salida = Path(ruta_salida)

    registros = (
        [("aviso", a) for a in avisos]
        + [("incidencia_soc", i) for i in incidencias]
    )

    df_log = pd.DataFrame(
        registros or [("ok", "Sin observaciones.")],
        columns=["tipo", "detalle"],
    )

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:

        df.to_excel(
            writer,
            sheet_name="Medidores",
            index=False,
        )

        fila = 0

        if df_wxy is not None:
            fila = _escribir_tabla_con_titulo(
                writer,
                HOJA_OFERTAS_SSCC,
                df_wxy,
                "Ofertas SSCC por dia (equivalente a Medidores!W:Y)",
                fila,
            )

        if df_resumen_ventana is not None:
            _escribir_tabla_con_titulo(
                writer,
                HOJA_OFERTAS_SSCC,
                df_resumen_ventana,
                "Resumen ventana oferta (equivalente a Medidores!AB:AE)",
                fila,
            )

        if df_cmg is not None:
            df_cmg.to_excel(
                writer,
                sheet_name="CMg",
                index=False,
            )

        if df_fd_csf is not None or df_fd_cpf is not None:

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

        if df_subastas is not None:
            df_subastas.to_excel(
                writer,
                sheet_name="Subastas",
                index=False,
            )

        df_log.to_excel(
            writer,
            sheet_name="Log",
            index=False,
        )

    return ruta_salida


# ============================================================
# PROCESO COMPLETO
# ============================================================

def ejecutar(carpeta_base, aamm, registrar=print, progreso=None):
    """
    Corre la etapa Medidores de punta a punta.

    aamm:      periodo ingresado por el usuario en la ventana (4
               digitos, ej. '2607').
    registrar: funcion para mensajes.
    progreso:  funcion que recibe 0..100.
    """

    def avanzar(valor):
        if progreso:
            progreso(valor)

    aamm = validar_aamm(aamm)

    rutas, _ = revisar_estructura(carpeta_base, aamm)

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

    if not rutas["cmg"].is_file():
        raise ErrorEntrada(
            f"No se encontro {rutas['cmg']}"
        )

    archivo_sscc = buscar_archivo_sscc_desempeno(
        rutas["sscc_desempeno_dir"]
    )
    if not archivo_sscc:
        raise ErrorEntrada(
            f"No se encontro ningun archivo SSCC_Desempeño_* en "
            f"{rutas['sscc_desempeno_dir']}"
        )

    archivo_subastas = buscar_archivo_subastas(rutas["subastas_dir"])
    if not archivo_subastas:
        raise ErrorEntrada(
            f"No se encontro ningun archivo "
            f"3_REMUNERACIÓN_SUBASTAS_E_ID_* en "
            f"{rutas['subastas_dir']}"
        )

    archivo_soc = buscar_soc(rutas["medidas_dir"], aamm)
    anio, mes = periodo_desde_aamm(aamm)

    registrar(f"Periodo indicado: {anio}-{mes:02d} ({aamm})")
    avanzar(5)

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

    avanzar(50)

    registrar(f"Leyendo {ARCHIVO_MEDIDAS_SAE}...")
    df_sae = leer_medidas_sae(rutas["medidas_sae"])
    registrar(f"  filas: {len(df_sae):,}")
    avanzar(70)

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

    avanzar(80)

    registrar(f"Leyendo {ARCHIVO_CMG}...")
    df_cmg = leer_cmg(rutas["cmg"], registrar=registrar)
    avanzar(85)

    registrar(f"Leyendo {archivo_sscc.name}...")
    df_fd_csf, df_fd_cpf = construir_fd(archivo_sscc, registrar=registrar)
    avanzar(90)

    registrar(f"Leyendo {archivo_subastas.name}...")
    df_subastas = construir_subastas(archivo_subastas, registrar=registrar)
    avanzar(95)

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
    )

    avanzar(100)
    registrar(f"Listo: {rutas['salida']}")

    return df_medidores, avisos, incidencias
