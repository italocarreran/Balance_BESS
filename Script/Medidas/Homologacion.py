# -*- coding: utf-8 -*-
"""
Homologacion — el Excel que dice que punto de medida + canal es cada
clave del balance.

Viene de "0_diccionario_prmte_a_claves_balance.py", que leia
'Homologacion ClavesTF y PRMTE.xlsx' del directorio actual y lo
guardaba como homol.parquet para que el paso siguiente lo levantara.
Ese parquet intermedio ya no existe: el archivo se lee una vez, desde
Auxiliares/ (al lado de Centrales.xlsx), y queda en memoria.

El mismo archivo trae ademas la hoja "Gen real" (ver HOJA_GEN_REAL):
las centrales que se miden por la API de operacion real. Estaba
pensada como una hoja de Centrales.xlsx, pero el usuario pidio que
viva aca -- es homologacion, igual que 'homol', y se mantiene con el
mismo archivo.
"""

from pathlib import Path

import pandas as pd

from .comun import (
    ErrorMedidas, EXTENSIONES_EXCEL, columna_que_contenga, normalizar
)


# El nombre real es 'Homologacion ClavesTF y PRMTE.xlsx', pero se busca
# por patron (como *OfertasSSCC*) para no depender de tildes, guiones o
# de que alguien le agregue el periodo al nombre.
PATRON_NOMBRE = "homologacion"
HOJA_HOMOL = "homol"

# Segunda hoja del mismo archivo: las centrales cuya medida NO sale de
# la API por punto de medida sino de la API de operacion real. Vive
# aca, y no en Centrales.xlsx, a pedido del usuario: es homologacion,
# igual que 'homol', y se mantiene con el mismo archivo.
#
# Mismas cuatro columnas que 'homol' (clave / Punto de Medida / Canal /
# Flujo), con una lectura propia de cada una:
#
#   clave           la clave del balance, igual que en 'homol'
#   Punto de Medida el 'topologyName' EXACTO de la API de operacion
#                   real (ej. "SAE PFV Andes Solar III (Inyección)"):
#                   es lo que identifica a la central en esa API, que
#                   no tiene el concepto de punto de medida
#   Canal           no se usa: esa API no expone canales. Se acepta
#                   para que la hoja tenga la misma forma que 'homol'
#   Flujo           +1 / -1, igual que en 'homol' (retiros en -1)
#
# La hoja es opcional: un caso sin centrales de este tipo es valido.
HOJA_GEN_REAL = "Gen real"

COLUMNA_PUNTO = "Punto de Medida"
COLUMNA_CANAL = "Canal"
COLUMNA_CLAVE = "clave"
COLUMNA_FLUJO = "Flujo"


def buscar_archivo_homologacion(carpeta_auxiliares):
    """
    Archivo de homologacion dentro de Auxiliares/: cualquier Excel
    cuyo nombre contenga "homologacion" (sin tildes). Si hay varios,
    el mas reciente por fecha de modificacion. None si no hay ninguno.
    """

    carpeta = Path(carpeta_auxiliares)

    if not carpeta.is_dir():
        return None

    candidatos = [
        archivo for archivo in carpeta.iterdir()
        if archivo.is_file()
        and not archivo.name.startswith("~$")
        and archivo.suffix.lower() in EXTENSIONES_EXCEL
        and PATRON_NOMBRE in normalizar(archivo.stem)
    ]

    if not candidatos:
        return None

    return max(candidatos, key=lambda a: a.stat().st_mtime)


def leer_homologacion(ruta):
    """
    Lee la hoja 'homol' y devuelve el DataFrame con las cuatro
    columnas que usa el cruce posterior:

        Punto de Medida | Canal | clave | Flujo

    'clave' se fuerza a texto (asi lo hacia el script original: hay
    claves que son solo numeros y pandas las leeria como int, y
    despues no cruzarian contra el resto del balance).
    """

    ruta = Path(ruta)

    try:
        excel = pd.ExcelFile(ruta)
    except Exception as error:
        raise ErrorMedidas(
            f"No se pudo abrir {ruta.name}: {error}"
        ) from error

    hoja = _buscar_hoja(excel, HOJA_HOMOL)

    if hoja is None:
        raise ErrorMedidas(
            f"{ruta.name} no tiene la hoja '{HOJA_HOMOL}'. "
            f"Hojas encontradas: {excel.sheet_names}"
        )

    df = pd.read_excel(ruta, sheet_name=hoja)

    columnas = {}
    for interno, fragmentos in (
        (COLUMNA_PUNTO, ("punto", "medida")),
        (COLUMNA_CANAL, ("canal",)),
        (COLUMNA_CLAVE, ("clave",)),
        (COLUMNA_FLUJO, ("flujo",)),
    ):
        real = columna_que_contenga(df, *fragmentos)
        if real is None:
            raise ErrorMedidas(
                f"La hoja '{hoja}' de {ruta.name} no tiene una columna "
                f"'{interno}'. Columnas encontradas: {list(df.columns)}"
            )
        columnas[interno] = real

    df = df.rename(columns={real: interno for interno, real in columnas.items()})
    df[COLUMNA_CLAVE] = df[COLUMNA_CLAVE].astype(str)

    df = df.dropna(subset=[COLUMNA_PUNTO, COLUMNA_CANAL])

    if df.empty:
        raise ErrorMedidas(
            f"La hoja '{hoja}' de {ruta.name} no tiene ninguna fila con "
            f"'Punto de Medida' y 'Canal' cargados."
        )

    return df[[COLUMNA_PUNTO, COLUMNA_CANAL, COLUMNA_CLAVE, COLUMNA_FLUJO]]


def _buscar_hoja(excel, nombre):
    for hoja in excel.sheet_names:
        if normalizar(hoja) == normalizar(nombre):
            return hoja
    return None


def leer_gen_real(ruta):
    """
    Lee la hoja "Gen real" y devuelve una lista de dicts con
    'topologyName', 'clave' y 'factor' -- la forma que espera
    Generacion_Real.

    Si la hoja no existe, devuelve lista vacia y el proceso sigue: un
    caso sin centrales de operacion real es valido.
    """

    ruta = Path(ruta)
    excel = pd.ExcelFile(ruta)

    hoja = _buscar_hoja(excel, HOJA_GEN_REAL)

    if hoja is None:
        return []

    df = pd.read_excel(ruta, sheet_name=hoja)

    columna_clave = columna_que_contenga(df, "clave")
    columna_punto = columna_que_contenga(df, "punto", "medida")
    columna_flujo = columna_que_contenga(df, "flujo")

    if columna_clave is None or columna_punto is None:
        raise ErrorMedidas(
            f"La hoja '{hoja}' de {ruta.name} tiene que tener las mismas "
            f"columnas que '{HOJA_HOMOL}' ('clave', 'Punto de Medida', "
            f"'Canal', 'Flujo'). Columnas encontradas: "
            f"{list(df.columns)}"
        )

    centrales = []

    for _, fila in df.iterrows():

        clave = fila[columna_clave]
        topology = fila[columna_punto]

        if pd.isna(clave) or pd.isna(topology):
            continue

        factor = 1.0

        if columna_flujo is not None and not pd.isna(fila[columna_flujo]):
            try:
                factor = float(fila[columna_flujo])
            except (TypeError, ValueError):
                raise ErrorMedidas(
                    f"En la hoja '{hoja}' de {ruta.name}, la central "
                    f"'{topology}' tiene un Flujo no numerico "
                    f"({fila[columna_flujo]!r}). Usa 1 o -1."
                )

        centrales.append(
            {
                "topologyName": str(topology).strip(),
                "clave": str(clave).strip(),
                "factor": factor,
            }
        )

    return centrales


def puntos_de_medida(df_homol):
    """Los puntos de medida distintos a consultar en la API."""

    return list(pd.unique(df_homol[COLUMNA_PUNTO].dropna()))
