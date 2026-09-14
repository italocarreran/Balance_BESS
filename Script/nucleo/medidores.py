# -*- coding: utf-8 -*-
"""
La hoja Medidores: columnas calculadas y armado.
"""

import pandas as pd
from pathlib import Path

from .ofertas_sscc import (
    calcular_r, calcular_s, calcular_t, cargar_resumen_en_medidores,
    construir_resumen_ofertas_sscc, construir_resumen_ventana_oferta,
)
from .parametros import (
    ARCHIVO_MEDIDAS_SAE, COLUMNAS_VACIAS, INICIO_VENTANA, LETRA_A_CAMPO,
    UMBRAL_SOC,
)
from .utiles import normalizar


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
        # Para cada central sin bloque, se busca si algun nombre CRUDO
        # (antes de homologar, "nombre_scada_original") normaliza igual
        # a esa central -- si lo encuentra, es una pista fuerte de que
        # el bloque SI esta en el archivo de SoC pero la homologacion
        # (Centrales.xlsx!Diccionario) lo esta mandando a otro nombre.
        candidatos_por_normalizado = {}
        for origen in soc["nombre_scada_original"].unique():
            candidatos_por_normalizado.setdefault(
                normalizar(origen), []
            ).append(origen)

        detalle_sin_soc = []
        for central in sin_soc:
            candidatos = candidatos_por_normalizado.get(
                normalizar(central), []
            )
            if candidatos:
                detalle_sin_soc.append(
                    f"{central} (el SoC SI trae un bloque con nombre "
                    f"crudo {candidatos!r} -- revisar si "
                    f"Centrales.xlsx!Diccionario lo esta homologando "
                    f"a otro nombre distinto de '{central}')"
                )
            else:
                detalle_sin_soc.append(central)

        avisos.append(
            f"Centrales en {ARCHIVO_MEDIDAS_SAE} sin bloque de "
            f"SoC: {detalle_sin_soc}"
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
