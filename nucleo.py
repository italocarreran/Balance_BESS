# -*- coding: utf-8 -*-
"""
Nucleo de calculo de la etapa Medidores.

Sin dependencias de interfaz: todo lo de aca se puede correr y
testear sin abrir la ventana.
"""

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

ARCHIVO_MEDIDAS_SAE = "Medidas_SAE.xlsx"
HOJA_MEDIDAS_SAE = "Medidas"

ARCHIVO_CENTRALES = "Centrales.xlsx"
HOJA_RESUMEN_BESS = "Resumen BESS"
HOJA_DICCIONARIO = "Diccionario"

ARCHIVO_SALIDA = "Hoja_Medidas.xlsx"

PATRON_SOC = re.compile(r"^SOC[_\-\s]?(\d{4})\.xlsx$", re.IGNORECASE)


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

# Nombre logico de cada letra de Excel, para poder comparar
# contra la hoja original columna por columna.
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
    "K": "K_PENDIENTE",
    "L": "Ventana",
    "M": "M_PENDIENTE",
    "N": "Clave_Dia_HoraMes",
    "O": "Indicador_SoC",
    "P": "P_PENDIENTE",
    "Q": "Q_PENDIENTE",
    "R": "R_PENDIENTE_OFERTAS",
    "S": "S_PENDIENTE_OFERTAS",
    "T": "T_PENDIENTE_OFERTAS",
}

# Columnas que todavia no se pueden calcular porque su regla
# no esta documentada o dependen de Ofertas SSCC.
COLUMNAS_PENDIENTES = [
    "K_PENDIENTE",
    "M_PENDIENTE",
    "P_PENDIENTE",
    "Q_PENDIENTE",
    "R_PENDIENTE_OFERTAS",
    "S_PENDIENTE_OFERTAS",
    "T_PENDIENTE_OFERTAS",
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

    return {
        "base": base,
        "medidas_dir": medidas_dir,
        "auxiliares_dir": auxiliares_dir,
        "medidas_sae": medidas_dir / ARCHIVO_MEDIDAS_SAE,
        "centrales": auxiliares_dir / ARCHIVO_CENTRALES,
        "salida": base / ARCHIVO_SALIDA,
    }


def buscar_soc(medidas_dir):
    """
    Busca SOC_AAMM.xlsx dentro de Medidas/.

    Ninguno   -> error
    Uno       -> se usa
    Varios    -> error, NO se elige por fecha de modificacion
    """

    medidas_dir = Path(medidas_dir)

    if not medidas_dir.is_dir():
        raise ErrorEntrada(
            f"No existe la carpeta {medidas_dir}"
        )

    candidatos = [
        archivo
        for archivo in medidas_dir.iterdir()
        if archivo.is_file()
        and not archivo.name.startswith("~$")
        and PATRON_SOC.match(archivo.name)
    ]

    if not candidatos:
        raise ErrorEntrada(
            f"No se encontro ningun archivo SOC_AAMM.xlsx en "
            f"{medidas_dir}\n"
            f"Ejemplos validos: SOC_2607.xlsx, SOC_2608.xlsx"
        )

    if len(candidatos) > 1:
        nombres = ", ".join(sorted(c.name for c in candidatos))
        raise ErrorEntrada(
            f"Hay {len(candidatos)} archivos SOC en "
            f"{medidas_dir}:\n"
            f"  {nombres}\n"
            f"Deja solo el del periodo que vas a procesar. "
            f"No se elige automaticamente para no tomar en "
            f"silencio el mes equivocado."
        )

    archivo = candidatos[0]
    aamm = PATRON_SOC.match(archivo.name).group(1)

    return archivo, aamm


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

def revisar_estructura(carpeta_base):
    """
    Revisa la carpeta base y devuelve una lista de
    (etiqueta, estado, detalle) para pintar en la ventana.

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

    # SOC
    try:
        archivo_soc, aamm = buscar_soc(rutas["medidas_dir"])
        filas.append(
            (
                archivo_soc.name,
                "ok",
                f"periodo {aamm}",
            )
        )
        rutas["soc"] = archivo_soc
        rutas["aamm"] = aamm
    except ErrorEntrada as error:
        filas.append(
            ("SOC_AAMM.xlsx", "falta", str(error).split("\n")[0])
        )
        rutas["soc"] = None
        rutas["aamm"] = None

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

    # Ofertas SSCC: ubicacion aun no definida en el plan
    filas.append(
        (
            "OfertasSSCC",
            "pendiente",
            "ubicacion por definir en el plan",
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
    registrar=print,
):
    """
    Arma la tabla equivalente a Medidores.

    Devuelve (df_medidores, avisos).
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

    df["Clave_Dia_HoraMes"] = calcular_clave_auxiliar(
        df["Dia"],
        df["Hora Mes"],
    )

    df["Indicador_SoC"] = calcular_indicador_soc(df["SoC"])

    # --------------------------------------------------------
    # COLUMNAS TODAVIA NO DEFINIDAS
    # --------------------------------------------------------

    for columna in COLUMNAS_PENDIENTES:
        df[columna] = pd.NA

    # --------------------------------------------------------
    # ORDEN FINAL DE COLUMNAS
    # --------------------------------------------------------

    orden = [
        LETRA_A_CAMPO[letra]
        for letra in sorted(LETRA_A_CAMPO)
    ]

    df = df[orden]

    registrar(
        f"Medidores construido: {len(df):,} filas x "
        f"{len(df.columns)} columnas"
    )

    return df, avisos


def escribir_salida(df, ruta_salida, avisos, incidencias):
    """Escribe Hoja_Medidas.xlsx con la tabla y un log."""

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

        df_log.to_excel(
            writer,
            sheet_name="Log",
            index=False,
        )

    return ruta_salida


# ============================================================
# PROCESO COMPLETO
# ============================================================

def ejecutar(carpeta_base, registrar=print, progreso=None):
    """
    Corre la etapa Medidores de punta a punta.

    registrar: funcion para mensajes.
    progreso:  funcion que recibe 0..100.
    """

    def avanzar(valor):
        if progreso:
            progreso(valor)

    rutas, _ = revisar_estructura(carpeta_base)

    if not rutas["medidas_sae"].is_file():
        raise ErrorEntrada(
            f"No se encontro {rutas['medidas_sae']}"
        )

    if not rutas["centrales"].is_file():
        raise ErrorEntrada(
            f"No se encontro {rutas['centrales']}"
        )

    archivo_soc, aamm = buscar_soc(rutas["medidas_dir"])
    anio, mes = periodo_desde_aamm(aamm)

    registrar(f"Periodo detectado: {anio}-{mes:02d} ({aamm})")
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
    df_medidores, avisos = construir_medidores(
        df_sae,
        df_soc,
        anio,
        mes,
        registrar=registrar,
    )

    for aviso in avisos:
        registrar(f"  [AVISO] {aviso}")

    avanzar(90)

    registrar(f"Escribiendo {rutas['salida'].name}...")
    escribir_salida(
        df_medidores,
        rutas["salida"],
        avisos,
        incidencias,
    )

    avanzar(100)
    registrar(f"Listo: {rutas['salida']}")

    return df_medidores, avisos, incidencias
