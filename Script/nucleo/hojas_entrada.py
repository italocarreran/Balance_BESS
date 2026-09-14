# -*- coding: utf-8 -*-
"""
Hojas CMg, FD y Subastas del consolidado.
"""

import pandas as pd
from pathlib import Path

from .parametros import (
    HOJA_CMG_ORIGEN, HOJA_CPF_HORARIO, HOJA_CSF_HORARIO,
    HOJA_SUBASTAS_ORIGEN,
)
from .utiles import ErrorEntrada, _texto_seguro


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

# Encabezados reales de Subastas!B:Q -- CORREGIDOS con el archivo
# Libro1.xlsx que trae la macro Cargar_Remuneracion_Subastas_Rapido y
# formulas reales cruzadas contra encabezados reales (sesion de
# correccion, ver BITACORA). La version anterior de este diccionario
# tenia TODO corrido una posicion: le faltaba la columna "Concepto"
# (B), la primera de las 11 que copia DB!B:L, que hasta esta sesion
# se asumia (mal) que era "A" y que la macro no tocaba.
#
# La prueba definitiva: Subastas!B1 tiene la formula real
#   =IF(AND(C1="CSF",D1="SUBIDA"),"CSF(+)",IF(AND(C1="CSF",D1="BAJADA"),
#     "CSF(-)",IF(AND(C1="CTF",D1="SUBIDA"),"CTF(+)",
#     IF(AND(C1="CTF",D1="BAJADA"),"CTF(-)","REVISAR"))))
# que arma "Concepto" (B) a partir de DOS insumos: "Control" (C, el
# tipo SIN direccion: CSF/CTF/CPF) y "Sub_Baj" (D, la direccion:
# SUBIDA/BAJADA). La version anterior solo tenia UNA columna ahi
# (llamada "Control", en la posicion de lo que en realidad es
# "Concepto") -- faltaba la columna "Control" real.
#
# "A" SIGUE sin usarse (confirmado: ninguna celda con datos en esa
# columna en el archivo real) -- no era un error de ubicacion, era
# que faltaba contar una columna mas dentro del bloque B:L.
NOMBRES_SUBASTAS = {
    "B": "Concepto",
    "C": "Control",
    "D": "Sub_Baj",
    "E": "Fecha",
    "F": "Año",
    "G": "Mes",
    "H": "Dia",
    "I": "Hora_dia",
    "J": "Hora_mes",
    "K": "Configuración",
    "L": "Propietario",
    "M": "Clave horaria",
    "N": "Ciclo",
    "O": "Energía SSCC",
    "P": "FD",
    "Q": "FMA",
}


def _ordenar_subastas_por_hora_mes(df):
    """
    Deja la hoja Subastas ordenada por Hora_mes (pedido del usuario).
    Se desempata por Configuración y Concepto para que dos corridas
    sobre los mismos datos den exactamente el mismo archivo.

    Ninguna columna de mas abajo depende del orden de las filas (todo
    lo que consume Subastas lo hace por clave: calcular_l,
    construir_prorrata_sscc, construir_dic_umbrales_subastas), asi que
    ordenar es puramente de presentacion.
    """

    columnas_orden = [
        NOMBRES_SUBASTAS["J"],   # Hora_mes
        NOMBRES_SUBASTAS["K"],   # Configuración
        NOMBRES_SUBASTAS["B"],   # Concepto
    ]

    presentes = [c for c in columnas_orden if c in df.columns]

    if not presentes:
        return df

    # Hora_mes puede venir como texto desde la planilla 3: se ordena
    # por su valor numerico, no alfabeticamente ("10" antes que "9").
    auxiliar = df.copy()
    clave_numerica = "__orden_hora_mes__"
    auxiliar[clave_numerica] = pd.to_numeric(
        auxiliar[presentes[0]], errors="coerce"
    )

    auxiliar = auxiliar.sort_values(
        by=[clave_numerica] + presentes[1:],
        kind="stable",
        na_position="last",
    )

    return auxiliar.drop(columns=[clave_numerica]).reset_index(drop=True)


def construir_subastas(ruta_subastas, registrar=print):
    """
    Replica Cargar_Remuneracion_Subastas_Rapido.

    La macro original consulta la hoja "DB" del archivo de origen por
    ADO/SQL (equivalente a filtrar y seleccionar columnas de una
    tabla); aca se lee directamente con pandas y se aplica el mismo
    filtro y la misma seleccion de columnas.

    Arma las columnas B:Q de Subastas (nombres reales en
    NOMBRES_SUBASTAS, corregidos con el archivo real -- ver el
    comentario de esa constante):
      - Concepto:Propietario (B:L, 11 columnas): copia directa de
        DB!B:L (filtrado por DB!K -- que en el archivo real resulta
        ser "Configuración", no "Propietario"; el filtro sigue
        siendo textualmente correcto porque los nombres de central
        BESS/SAE empiezan con "SAE-", asi que "contiene BESS o SAE"
        encuentra las mismas filas de cualquier forma).
      - Clave horaria (M, formula): = Configuración & Dia & Hora_dia
        (K&H&I con las letras REALES de Subastas; coincide con la
        formula real M3=K3&H3&I3 del archivo).
      - Ciclo (N): se deja vacia EN ESTA HOJA. Ya no es un calculo
        desconocido (ver calcular_subastas_ciclo: es el "Ciclo de
        Carga del mes" de Calculo E Costos homologado por Hora_mes +
        Configuración), pero depende de Calculo E Costos, que se
        arma despues y en el otro archivo (Pagos_BESS.xlsx). Se
        calcula ahi, donde se usa. (Antes de esta sesion esta
        columna se llamaba "Energía SSCC" -- nombre equivocado: lo
        que la formula real trae es un numero de ciclo, no energia.)
      - Energía SSCC, FD, FMA (O, P, Q): copias de DB!P, DB!Y, DB!V
        respectivamente (asi lo indica la macro original) -- simples
        copias, sin formula ni calculo dentro de Subastas.
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
    df = _ordenar_subastas_por_hora_mes(df)

    registrar(
        f"  Subastas: {len(df):,} fila(s) (filtro Propietario "
        f"contiene BESS/SAE), ordenadas por Hora_mes"
    )

    return df
