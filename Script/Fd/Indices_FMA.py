# -*- coding: utf-8 -*-
"""
Indices_FMA — arma las tres salidas de FMA del periodo.

Viene de las rutinas `calc_fmacpf`, `calc_fmacsf` y `calc_fmactf` del
script suelto entradas_sscc.py (autor original: Gerardo.Vieyra).

OJO con la palabra "traer": el FMA **no se copia de ningun lado ya
hecho**, se construye. Cada una de las tres sale de un origen distinto:

  - **CPF**: de los reportes diarios que publica el DCO, adentro del
    mismo arbol de indicadores del que sale el FD:
        <version>/01 Respuesta/01 Indices CPF/20AA.MM_Respuesta_CPF/
            Reporte diario <D>-<M>-<AAAA>/tabla_resumen_<D>_<M>_<AAAA>.xlsx
    Se lee **cada hoja menos "Resumen"** (una por central), se le pegan
    Año/Mes/Dia/Hora/Central y se agrupan las columnas de horas.
  - **CSF**: de los archivos diarios `csf_20AAMMDD.xlsx` (los del
    agc_face). Es una concatenacion, sin transformacion.
  - **CTF**: del `CTF_20AAMM.csv`, al que solo se le saca la zona
    horaria de t0/tfin.

Las tres se escriben en <CARPETA_BASE>/FD y FMA/ con los nombres que
usa entradas_sscc.py (`fma_cpf_AAMM.xlsx`, `fma_csf_AAMM.xlsx`,
`fma_cft_AAMM.xlsx` -- "cft" esta asi en el original), que son
exactamente los que despues busca Script/Subastas/Fma.py para armar
Subastas!FMA.

Diferencias a proposito respecto del script original:

  1) el periodo sale del AAMM de la ventana, no de un yaml;
  2) un dia sin archivo se saltea con un aviso en vez de cortar todo
     (el original revienta con la excepcion de pandas);
  3) si falta el origen de alguna de las tres, las otras dos se arman
     igual y se avisa cual falto;
  4) no se usa chardet para adivinar la codificacion del CSV de CTF:
     se prueban las codificaciones que de verdad aparecen, en orden.

No importa nada de nucleo.py (solo pandas). Los errores previsibles
salen como ErrorIndicesFma.
"""

import re
from pathlib import Path

import pandas as pd

try:
    from . import Indicadores_DCO as dco
except ImportError:  # pragma: no cover - depende de como se importe
    import Indicadores_DCO as dco


# ============================================================
# NOMBRES
# ============================================================

# Los nombres de salida de entradas_sscc.py, tal cual ("cft" incluido).
PLANTILLAS_SALIDA = {
    "cpf": "fma_cpf_{aamm}.xlsx",
    "csf": "fma_csf_{aamm}.xlsx",
    "ctf": "fma_cft_{aamm}.xlsx",
}

# Donde cuelgan los reportes diarios de CPF dentro de la carpeta de
# version del DCO. Se busca por nombre, no por ruta literal.
PATRON_CARPETA_RESPUESTA_CPF = "respuesta_cpf"
PLANTILLA_TABLA_RESUMEN = "tabla_resumen_{dia}_{mes}_{anio}"

# Subcarpeta (dentro de "FD y FMA/") donde se esperan los csf_ diarios.
# En el script original era `input/agc_face`.
CARPETA_AGC_FACE = "agc_face"
PLANTILLA_CSF_DIARIO = "csf_{anio}{mes:02d}{dia:02d}"
PLANTILLA_CTF = "CTF_{anio}{mes:02d}"

EXTENSIONES_EXCEL = (".xlsx", ".xlsm", ".xlsb", ".xls")

# Los 34 nombres que entradas_sscc.py le pone al bloque de CPF: 29
# columnas del archivo del DCO mas las 5 que agrega el script.
NOMBRES_CPF = [
    "Hora Local",
    "(a1)", "(a2-)", "(a2+)", "(a3-)", "(a3+)", "(b)",
    "(c)",
    "(d)",
    "(e1)", "(e2-)", "(e2+)", "(e3-)", "(e3+)", "(e4-)", "(e4+)",
    "(f1)", "(f2)", "(f3)",
    "(g1)", "(g2)", "(g3)", "(g4)",
    "NC [Horas]",
    "indic CPF+", "indic CPF-",
    "Tiempo f<49.975 [%]", "Tiempo f>50.025 [%]",
    "Registro incluye una contingencia",
    "Año", "Mes", "Día", "Hora", "Central",
]

GRUPOS_CPF = {
    "Hace CPF [hrs]": ["(a1)", "(a2-)", "(a2+)", "(a3-)", "(a3+)"],
    "No hace CPF [hrs]": [
        "(e1)", "(e2-)", "(e2+)", "(e3-)", "(e3+)", "(e4-)", "(e4+)",
    ],
    "Indef. [hrs]": ["(f1)", "(f2)", "(f3)"],
    "Otro [hrs]": ["(g1)", "(g2)", "(g3)", "(g4)"],
}

COPIAS_CPF = {
    "Paráms. Fuera de rango": "(b)",
    "Hace CPF con lím superior >Pmax": "(c)",
    "Hace CPF con lím inferior <MT": "(d)",
}

COLUMNAS_SALIDA_CPF = [
    "Año", "Mes", "Día", "Hora", "Central",
    "Hace CPF [hrs]",
    "Paráms. Fuera de rango",
    "Hace CPF con lím superior >Pmax",
    "Hace CPF con lím inferior <MT",
    "No hace CPF [hrs]",
    "Indef. [hrs]",
    "Otro [hrs]",
    "Tiempo f<49.975 [%]",
    "Tiempo f>50.025 [%]",
]

HOJA_A_SALTEAR_CPF = "resumen"

# El CSV de CTF viene con separador ";" y coma decimal. El original usa
# chardet para adivinar la codificacion; aca se prueban en orden las que
# de verdad aparecen en estos archivos.
CODIFICACIONES_CTF = ("utf-8-sig", "utf-8", "latin1")


class ErrorIndicesFma(Exception):
    """Error previsible al armar las salidas de FMA."""


def nombre_salida(tipo, aamm):
    """('cpf', '2603') -> 'fma_cpf_2603.xlsx'"""

    return PLANTILLAS_SALIDA[tipo].format(aamm=aamm)


# ============================================================
# CPF
# ============================================================

def buscar_carpeta_respuesta_cpf(carpeta_version):
    """
    Ubica, dentro de la carpeta de version del DCO, la carpeta
    '20AA.MM_Respuesta_CPF' (la que tiene adentro un "Reporte diario"
    por dia). Se busca recursivamente y por nombre normalizado, porque
    el DCO le cambia el nivel de anidamiento de un mes a otro.
    """

    carpeta_version = Path(carpeta_version)

    try:
        for ruta in carpeta_version.rglob("*"):
            if (
                ruta.is_dir()
                and PATRON_CARPETA_RESPUESTA_CPF in dco._normalizar(
                    ruta.name
                ).replace(" ", "_")
            ):
                return ruta
    except OSError as error:
        raise ErrorIndicesFma(
            f"No se pudo recorrer {carpeta_version}: {error}"
        ) from error

    return None


def buscar_tabla_resumen(carpeta_respuesta, anio, mes, dia):
    """
    El `tabla_resumen_<D>_<M>_<AAAA>.xlsx` del dia, buscado
    recursivamente (cuelga de una carpeta "Reporte diario <D>-<M>-<AAAA>",
    pero el nombre de esa carpeta cambia). El dia y el mes van SIN cero
    a la izquierda, como los escribe el DCO.

    Se acepta cualquier sufijo despues del año porque existen variantes
    con la zona horaria en el nombre (`..._UTC-4.xlsx`, `..._utc-3.xlsx`).
    """

    prefijo = dco._normalizar(
        PLANTILLA_TABLA_RESUMEN.format(dia=dia, mes=mes, anio=anio)
    )

    candidatos = []

    try:
        for ruta in Path(carpeta_respuesta).rglob("*"):
            if (
                ruta.is_file()
                and ruta.suffix.lower() in EXTENSIONES_EXCEL
                and dco._normalizar(ruta.stem).startswith(prefijo)
            ):
                candidatos.append(ruta)
    except OSError:
        return None

    if not candidatos:
        return None

    # El nombre "pelado" (sin sufijo de zona horaria) primero; si no,
    # el mas reciente.
    exactos = [
        r for r in candidatos if dco._normalizar(r.stem) == prefijo
    ]

    if exactos:
        return exactos[0]

    return max(candidatos, key=lambda r: r.stat().st_mtime)


def construir_fma_cpf(carpeta_respuesta, aamm, registrar=print):
    """
    Arma el equivalente de `fma_cpf_AAMM.xlsx` recorriendo los reportes
    diarios. Replica calc_fmacpf de entradas_sscc.py.
    """

    anio, mes = dco.periodo_desde_aamm(aamm)
    dias_del_mes = _dias_del_mes(anio, mes)

    partes = []
    dias_sin_archivo = []

    for dia in range(1, dias_del_mes + 1):

        ruta = buscar_tabla_resumen(carpeta_respuesta, anio, mes, dia)

        if ruta is None:
            dias_sin_archivo.append(dia)
            continue

        try:
            excel = pd.ExcelFile(ruta)
        except Exception as error:
            raise ErrorIndicesFma(
                f"No se pudo abrir {ruta.name}: {error}"
            ) from error

        for hoja in excel.sheet_names:

            if dco._normalizar(hoja) == HOJA_A_SALTEAR_CPF:
                continue

            df_hoja = pd.read_excel(excel, header=4, sheet_name=hoja)

            if df_hoja.empty:
                continue

            df_hoja = df_hoja.copy()
            df_hoja["Año"] = anio
            df_hoja["Mes"] = mes
            df_hoja["Día"] = dia
            # La hora sale del indice de fila del reporte (0..23) + 1.
            df_hoja["Hora"] = pd.to_numeric(
                df_hoja.iloc[:, 0], errors="coerce"
            ) + 1
            df_hoja["Central"] = hoja

            if df_hoja.shape[1] != len(NOMBRES_CPF):
                raise ErrorIndicesFma(
                    f"La hoja '{hoja}' de {ruta.name} trae "
                    f"{df_hoja.shape[1]} columnas y se esperaban "
                    f"{len(NOMBRES_CPF)}.\n\n"
                    f"Si el DCO le cambio el formato al reporte, hay que "
                    f"actualizar NOMBRES_CPF en "
                    f"Script/Fd/Indices_FMA.py."
                )

            df_hoja.columns = NOMBRES_CPF
            partes.append(df_hoja)

        excel.close()

    if not partes:
        raise ErrorIndicesFma(
            f"No se encontro ningun 'tabla_resumen_<dia>_{mes}_{anio}' en "
            f"{carpeta_respuesta}."
        )

    df = pd.concat(partes, ignore_index=True)

    for nombre, columnas in GRUPOS_CPF.items():
        df[nombre] = df[columnas].apply(
            pd.to_numeric, errors="coerce"
        ).sum(axis=1)

    for nombre, origen in COPIAS_CPF.items():
        df[nombre] = df[origen]

    df = df[COLUMNAS_SALIDA_CPF]

    # El reporte escribe "-" donde no hubo medicion.
    for columna in ("Tiempo f<49.975 [%]", "Tiempo f>50.025 [%]"):
        df[columna] = df[columna].replace("-", 0)

    if dias_sin_archivo:
        registrar(
            f"  [AVISO] sin reporte diario de CPF para el/los dia(s): "
            f"{', '.join(str(d) for d in dias_sin_archivo)}"
        )

    registrar(
        f"  FMA CPF: {len(df):,} fila(s), "
        f"{dias_del_mes - len(dias_sin_archivo)} dia(s)"
    )

    return df


# ============================================================
# CSF
# ============================================================

def construir_fma_csf(carpeta_csf, aamm, registrar=print):
    """
    Concatena los `csf_20AAMMDD.xlsx` del mes. Replica calc_fmacsf de
    entradas_sscc.py, que no les hace ninguna transformacion.
    """

    anio, mes = dco.periodo_desde_aamm(aamm)
    dias_del_mes = _dias_del_mes(anio, mes)
    carpeta_csf = Path(carpeta_csf)

    partes = []
    dias_sin_archivo = []

    for dia in range(1, dias_del_mes + 1):

        prefijo = dco._normalizar(
            PLANTILLA_CSF_DIARIO.format(anio=anio, mes=mes, dia=dia)
        )

        candidatos = [
            r for r in carpeta_csf.iterdir()
            if r.is_file()
            and r.suffix.lower() in EXTENSIONES_EXCEL
            and dco._normalizar(r.stem).startswith(prefijo)
        ] if carpeta_csf.is_dir() else []

        if not candidatos:
            dias_sin_archivo.append(dia)
            continue

        ruta = max(candidatos, key=lambda r: r.stat().st_mtime)

        try:
            partes.append(pd.read_excel(ruta))
        except Exception as error:
            raise ErrorIndicesFma(
                f"No se pudo leer {ruta.name}: {error}"
            ) from error

    if not partes:
        raise ErrorIndicesFma(
            f"No se encontro ningun 'csf_{anio}{mes:02d}<dia>.xlsx' en "
            f"{carpeta_csf}."
        )

    df = pd.concat(partes, ignore_index=True)

    if dias_sin_archivo:
        registrar(
            f"  [AVISO] sin csf_ diario para el/los dia(s): "
            f"{', '.join(str(d) for d in dias_sin_archivo)}"
        )

    registrar(
        f"  FMA CSF: {len(df):,} fila(s), "
        f"{dias_del_mes - len(dias_sin_archivo)} dia(s)"
    )

    return df


# ============================================================
# CTF
# ============================================================

def construir_fma_ctf(ruta_csv, registrar=print):
    """
    Lee el `CTF_20AAMM.csv` y le saca la zona horaria a t0/tfin.
    Replica calc_fmactf de entradas_sscc.py.
    """

    ruta_csv = Path(ruta_csv)
    df = None
    ultimo_error = None

    for codificacion in CODIFICACIONES_CTF:
        try:
            df = pd.read_csv(
                ruta_csv, sep=";", decimal=",", encoding=codificacion
            )
            break
        except (UnicodeDecodeError, LookupError) as error:
            ultimo_error = error
        except Exception as error:
            raise ErrorIndicesFma(
                f"No se pudo leer {ruta_csv.name}: {error}"
            ) from error

    if df is None:
        raise ErrorIndicesFma(
            f"No se pudo leer {ruta_csv.name} con ninguna de las "
            f"codificaciones probadas "
            f"({', '.join(CODIFICACIONES_CTF)}): {ultimo_error}"
        )

    for columna in ("t0", "tfin"):
        if columna not in df.columns:
            raise ErrorIndicesFma(
                f"A {ruta_csv.name} le falta la columna '{columna}'.\n\n"
                f"Columnas encontradas: {list(df.columns)}"
            )

        df[columna] = _sacar_zona_horaria(df[columna])

    registrar(f"  FMA CTF: {len(df):,} activacion(es) de {ruta_csv.name}")

    return df


# ============================================================
# TRAER LAS TRES
# ============================================================

def _sacar_zona_horaria(serie):
    """
    Replica `pd.to_datetime(col).dt.tz_localize(None)` del script
    original: deja la hora TAL COMO ESTA ESCRITA en el archivo y le
    saca el ofset.

    NO es lo mismo que convertir a UTC: '2026-03-01 04:00:00-03:00'
    tiene que quedar en las 04:00, no en las 07:00. Esa diferencia
    corre todas las horas del CTF (se detecto en la prueba).

    Se contempla el caso de ofsets mezclados en la misma columna -el
    dia del cambio de hora-, donde pandas ya no devuelve una columna
    tz-aware sino objetos sueltos.
    """

    try:
        fechas = pd.to_datetime(serie, errors="coerce", format="mixed")
    except (TypeError, ValueError):
        # format="mixed" existe recien desde pandas 2.0.
        fechas = pd.to_datetime(serie, errors="coerce")

    if getattr(fechas.dtype, "tz", None) is not None:
        return fechas.dt.tz_localize(None)

    def sin_zona(valor):
        if valor is None or pd.isna(valor):
            return pd.NaT
        if getattr(valor, "tzinfo", None) is not None:
            return valor.tz_localize(None)
        return valor

    return fechas.map(sin_zona)


def _dias_del_mes(anio, mes):
    import calendar

    return calendar.monthrange(anio, mes)[1]


def _buscar_en(carpetas, prefijo, extensiones, recursivo=False):
    """
    Primer archivo cuyo nombre empieza con el prefijo, buscando en las
    carpetas dadas en orden (la primera que tenga algo gana).
    """

    prefijo = dco._normalizar(prefijo)

    for carpeta in carpetas:

        carpeta = Path(carpeta)

        if not carpeta.is_dir():
            continue

        try:
            rutas = carpeta.rglob("*") if recursivo else carpeta.iterdir()
            candidatos = [
                r for r in rutas
                if r.is_file()
                and r.suffix.lower() in extensiones
                and dco._normalizar(r.stem).startswith(prefijo)
            ]
        except OSError:
            continue

        if candidatos:
            return max(candidatos, key=lambda r: r.stat().st_mtime)

    return None


def traer_fma(
    carpeta_destino, aamm, version=None, raiz=None, registrar=print
):
    """
    Arma las tres salidas de FMA del periodo y las escribe en
    <CARPETA_BASE>/FD y FMA/ (boton "Traer FMA").

    De donde sale cada una:
      - CPF: de los reportes diarios del DCO (carpeta de version del
        periodo, la misma de la que sale el FD);
      - CSF: de los `csf_20AAMMDD.xlsx`, que se buscan en
        'FD y FMA/agc_face/', en 'FD y FMA/' y, si no, en la carpeta de
        version del DCO;
      - CTF: del `CTF_20AAMM.csv`, con el mismo orden de busqueda.

    Si falta el origen de alguna, las otras se arman igual. Devuelve
    (escritos, faltantes, carpeta_version).
    """

    anio, mes = dco.periodo_desde_aamm(aamm)
    carpeta_destino = Path(carpeta_destino)

    if not carpeta_destino.is_dir():
        raise ErrorIndicesFma(f"No se encontro la carpeta {carpeta_destino}")

    carpeta_publicacion = dco.carpeta_del_periodo(aamm, raiz=raiz)
    carpeta_version = dco.elegir_version(carpeta_publicacion, version)

    if version is None:
        registrar(
            f"  version publicada mas alta: {carpeta_version.name} "
            f"(V1 = Preliminar, V2 = Definitivo)"
        )

    # Donde se buscan los insumos de CSF y CTF, en orden.
    carpetas_locales = [
        carpeta_destino / CARPETA_AGC_FACE,
        carpeta_destino,
    ]

    escritos, faltantes = {}, []

    # ---- CPF ----
    carpeta_respuesta = buscar_carpeta_respuesta_cpf(carpeta_version)

    if carpeta_respuesta is None:
        faltantes.append(
            f"CPF (no hay carpeta '*Respuesta_CPF*' en "
            f"{carpeta_version})"
        )
    else:
        registrar(f"  reportes diarios de CPF: {carpeta_respuesta}")
        df_cpf = construir_fma_cpf(carpeta_respuesta, aamm, registrar)
        escritos["cpf"] = _escribir(
            df_cpf, carpeta_destino / nombre_salida("cpf", aamm)
        )

    # ---- CSF ----
    carpeta_csf = None

    for carpeta in carpetas_locales:
        if _buscar_en(
            [carpeta],
            PLANTILLA_CSF_DIARIO.format(anio=anio, mes=mes, dia=1)[:-2],
            EXTENSIONES_EXCEL,
        ):
            carpeta_csf = carpeta
            break

    if carpeta_csf is None:
        carpeta_csf = _carpeta_con_csf_en_dco(carpeta_version, anio, mes)

    if carpeta_csf is None:
        faltantes.append(
            f"CSF (no hay ningun 'csf_{anio}{mes:02d}*' en "
            f"{carpeta_destino / CARPETA_AGC_FACE}, en {carpeta_destino} "
            f"ni en {carpeta_version})"
        )
    else:
        registrar(f"  csf_ diarios: {carpeta_csf}")
        df_csf = construir_fma_csf(carpeta_csf, aamm, registrar)
        escritos["csf"] = _escribir(
            df_csf, carpeta_destino / nombre_salida("csf", aamm)
        )

    # ---- CTF ----
    prefijo_ctf = PLANTILLA_CTF.format(anio=anio, mes=mes)

    ruta_ctf = _buscar_en(carpetas_locales, prefijo_ctf, (".csv",))

    if ruta_ctf is None:
        ruta_ctf = _buscar_en(
            [carpeta_version], prefijo_ctf, (".csv",), recursivo=True
        )

    if ruta_ctf is None:
        faltantes.append(
            f"CTF (no hay ningun '{prefijo_ctf}.csv' en "
            f"{carpeta_destino} ni en {carpeta_version})"
        )
    else:
        df_ctf = construir_fma_ctf(ruta_ctf, registrar)
        escritos["ctf"] = _escribir(
            df_ctf, carpeta_destino / nombre_salida("ctf", aamm)
        )
        # El script original tambien deja el csv al lado.
        df_ctf.to_csv(
            carpeta_destino / (nombre_salida("ctf", aamm)[:-5] + ".csv"),
            sep=",",
            index=False,
        )

    if not escritos:
        raise ErrorIndicesFma(
            "No se pudo armar ninguna de las tres salidas de FMA:\n  - "
            + "\n  - ".join(faltantes)
        )

    return escritos, faltantes, carpeta_version


def _carpeta_con_csf_en_dco(carpeta_version, anio, mes):
    """
    Busca, dentro de la carpeta de version del DCO, alguna subcarpeta
    que tenga los `csf_` del mes, y devuelve esa subcarpeta.
    """

    prefijo = dco._normalizar(f"csf_{anio}{mes:02d}")

    try:
        for ruta in Path(carpeta_version).rglob("*"):
            if (
                ruta.is_file()
                and ruta.suffix.lower() in EXTENSIONES_EXCEL
                and dco._normalizar(ruta.stem).startswith(prefijo)
            ):
                return ruta.parent
    except OSError:
        return None

    return None


def _escribir(df, ruta):
    """Escribe la salida y devuelve el nombre del archivo."""

    try:
        with pd.ExcelWriter(ruta) as writer:
            df.to_excel(writer, index=False)
    except OSError as error:
        raise ErrorIndicesFma(
            f"No se pudo escribir {Path(ruta).name}: {error}"
        ) from error

    return Path(ruta).name
