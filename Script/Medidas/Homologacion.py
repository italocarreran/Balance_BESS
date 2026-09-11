# -*- coding: utf-8 -*-
"""
Homologacion — el Excel que dice que punto de medida + canal es cada
clave del balance.

Viene de "0_diccionario_prmte_a_claves_balance.py", que leia
'Homologacion ClavesTF y PRMTE.xlsx' del directorio actual y lo
guardaba como homol.parquet para que el paso siguiente lo levantara.
Ese parquet intermedio ya no existe: el archivo se lee una vez, desde
Auxiliares/ (al lado de Centrales.xlsx), y queda en memoria.
"""

from pathlib import Path

import pandas as pd

from .comun import ErrorMedidas, EXTENSIONES_EXCEL, columna_que_contenga


# El nombre real es 'Homologacion ClavesTF y PRMTE.xlsx', pero se busca
# por patron (como *OfertasSSCC*) para no depender de tildes, guiones o
# de que alguien le agregue el periodo al nombre.
PATRON_NOMBRE = "homologacion"
HOJA_HOMOL = "homol"

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
        and PATRON_NOMBRE in _sin_tildes(archivo.stem)
    ]

    if not candidatos:
        return None

    return max(candidatos, key=lambda a: a.stat().st_mtime)


def _sin_tildes(texto):
    from .comun import normalizar
    return normalizar(texto)


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

    hoja = None
    for nombre in excel.sheet_names:
        if str(nombre).strip().lower() == HOJA_HOMOL:
            hoja = nombre
            break

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


def puntos_de_medida(df_homol):
    """Los puntos de medida distintos a consultar en la API."""

    return list(pd.unique(df_homol[COLUMNA_PUNTO].dropna()))
