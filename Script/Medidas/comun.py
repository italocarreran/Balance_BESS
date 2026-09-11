# -*- coding: utf-8 -*-
"""
Lo minimo compartido por los modulos de Medidas.

A proposito NO se importa nada de nucleo.py (ver Script/__init__.py):
`normalizar` esta duplicada -son diez lineas- para no crear un ciclo
de imports entre el nucleo y sus modulos de etapa.
"""

import unicodedata


class ErrorMedidas(Exception):
    """Error previsible al armar Medidas_SAE.xlsx."""


# ============================================================
# CLAVE DE LA API DEL COORDINADOR
#
# Las dos APIs que usa Medidas (medidas.* y operacion.*) piden la
# misma user_key. Va aca, en el codigo, a pedido explicito del
# usuario: los scripts originales la traian escrita adentro y asi se
# queda. Un solo lugar para las dos (antes estaba repetida en dos
# archivos, y con valores distintos).
#
# Ojo con dos cosas:
#   - queda versionada: quien tenga acceso al repositorio la tiene;
#   - si el Coordinador la cambia, se cambia aca y nada mas.
# ============================================================

USER_KEY = ""


EXTENSIONES_EXCEL = (".xlsx", ".xlsm", ".xlsb", ".xls")


def normalizar(texto):
    """Minusculas, sin tildes, espacios colapsados."""

    if texto is None:
        return ""

    try:
        import pandas as pd
        if pd.isna(texto):
            return ""
    except (ImportError, TypeError, ValueError):
        pass

    texto = str(texto)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))

    return " ".join(texto.lower().split())


def columna_que_contenga(df, *fragmentos):
    """
    Primera columna de df cuyo nombre normalizado contenga TODOS los
    fragmentos dados. None si no hay ninguna. Se busca por nombre y no
    por posicion, igual que en el resto del proyecto.
    """

    for columna in df.columns:
        clave = normalizar(columna)
        if all(fragmento in clave for fragmento in fragmentos):
            return columna

    return None
