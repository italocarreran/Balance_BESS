# -*- coding: utf-8 -*-
"""Asignacion de la compensacion mediante la prorrata de retiros."""

from pathlib import Path

import numpy as np
import pandas as pd

from .subastas_accdb import construir_mapa_propietario
from .utiles import ErrorEntrada, normalizar


HOJA_ORIGEN = "Prorrata 15min"
COLUMNAS_ORIGEN = ("Cuarto de Hora", "Suministrador", "Prorrata")
TOLERANCIA_PRORRATA = 1e-6


def buscar_archivo_prorrata(carpeta, aamm=None):
    """Encuentra un unico Prorrata_Retiros_AAMM_pre/def en su carpeta."""

    carpeta = Path(carpeta)
    if not carpeta.is_dir():
        return None
    candidatos = []
    for archivo in carpeta.iterdir():
        nombre = normalizar(archivo.stem).replace(" ", "_")
        if (
            archivo.is_file()
            and not archivo.name.startswith("~$")
            and archivo.suffix.lower() in {".xlsx", ".xlsm", ".xlsb", ".xls"}
            and nombre.startswith("prorrata_retiros_")
            and (nombre.endswith("_pre") or nombre.endswith("_def"))
            and (not aamm or f"_{aamm}_" in nombre)
        ):
            candidatos.append(archivo)
    if len(candidatos) > 1:
        nombres = "\n".join(f"- {p.name}" for p in sorted(candidatos))
        raise ErrorEntrada(
            "Hay mas de un archivo de prorrata para el periodo. Deja solo "
            f"el que corresponda en {carpeta}:\n{nombres}"
        )
    return candidatos[0] if candidatos else None


def leer_prorrata_retiros(ruta):
    """Lee y valida la fuente oficial, siempre desde ``Prorrata 15min``."""

    ruta = Path(ruta)
    try:
        excel = pd.ExcelFile(ruta)
    except Exception as error:
        raise ErrorEntrada(f"No se pudo abrir la prorrata {ruta}: {error}") from error
    if HOJA_ORIGEN not in excel.sheet_names:
        raise ErrorEntrada(
            f"{ruta.name} no tiene la hoja obligatoria '{HOJA_ORIGEN}'."
        )
    df = pd.read_excel(ruta, sheet_name=HOJA_ORIGEN)
    faltantes = [c for c in COLUMNAS_ORIGEN if c not in df.columns]
    if faltantes:
        raise ErrorEntrada(
            f"{ruta.name}/{HOJA_ORIGEN}: faltan columnas {faltantes}. "
            f"Se esperan {list(COLUMNAS_ORIGEN)}."
        )
    df = df.loc[:, COLUMNAS_ORIGEN].copy()
    cuarto = pd.to_numeric(df["Cuarto de Hora"], errors="coerce")
    prorrata = pd.to_numeric(df["Prorrata"], errors="coerce")
    suministrador = df["Suministrador"].fillna("").astype(str).str.strip()
    problemas = []
    if cuarto.isna().any():
        problemas.append(f"Cuarto de Hora no numerico/vacio: {int(cuarto.isna().sum())}")
    if prorrata.isna().any():
        problemas.append(f"Prorrata no numerica/vacia: {int(prorrata.isna().sum())}")
    if suministrador.eq("").any():
        problemas.append(f"Suministrador vacio: {int(suministrador.eq('').sum())}")
    if prorrata.lt(0).any():
        problemas.append(f"Prorrata negativa: {int(prorrata.lt(0).sum())}")
    if problemas:
        raise ErrorEntrada("RET-004/005/006: " + "; ".join(problemas))
    cuarto = cuarto.astype(int)
    duplicados = pd.DataFrame({"q": cuarto, "s": suministrador}).duplicated()
    if duplicados.any():
        ejemplo = df.loc[duplicados, list(COLUMNAS_ORIGEN)].head(5).to_dict("records")
        raise ErrorEntrada(f"RET-007: cuarto+suministrador duplicado. Ejemplos: {ejemplo}")
    sumas = prorrata.groupby(cuarto).sum()
    malas = sumas[~np.isclose(sumas, 1.0, atol=TOLERANCIA_PRORRATA, rtol=0)]
    if not malas.empty:
        detalle = ", ".join(f"{q}={v:.12g}" for q, v in malas.head(10).items())
        raise ErrorEntrada(
            f"RET-001: la prorrata no suma 1 en {len(malas)} cuarto(s): {detalle}"
        )
    return pd.DataFrame({
        "Cuarto de Hora": cuarto,
        "Suministrador": suministrador,
        "Prorrata": prorrata.astype(float),
    })


def construir_prorrata_retiros(df_prorrata, df_ecostos, df_re545):
    """Arma detalle, compensacion por cuarto y consolidado mensual."""

    def monto_por_cuarto(df, columna_cuarto, columna_monto):
        q = pd.to_numeric(df[columna_cuarto], errors="coerce")
        monto = pd.to_numeric(df[columna_monto], errors="coerce").fillna(0.0)
        return monto.groupby(q).sum()

    ec = monto_por_cuarto(df_ecostos, "Bloque Mes Descarga", "Monto a compensar")
    re = monto_por_cuarto(df_re545, "Bloque horario", "Monto a compensar")
    compensacion = ec.add(re, fill_value=0.0).sort_index()
    cuartos_prorrata = set(df_prorrata["Cuarto de Hora"])
    faltantes = [int(q) for q, monto in compensacion.items() if monto != 0 and q not in cuartos_prorrata]
    if faltantes:
        raise ErrorEntrada(
            f"RET-002: {len(faltantes)} cuarto(s) con compensacion no tienen prorrata: {faltantes[:10]}"
        )
    detalle = df_prorrata.copy()
    detalle["Compensacion a repartir"] = detalle["Cuarto de Hora"].map(compensacion)
    sin_compensacion = detalle["Compensacion a repartir"].isna()
    if sin_compensacion.any():
        cuartos = sorted(detalle.loc[sin_compensacion, "Cuarto de Hora"].unique())
        raise ErrorEntrada(
            f"RET-003: {len(cuartos)} cuarto(s) de prorrata no aparecen en los calculos: {cuartos[:10]}"
        )
    detalle["Pago"] = detalle["Compensacion a repartir"] * detalle["Prorrata"]
    asignado = detalle.groupby("Cuarto de Hora")["Pago"].sum()
    esperado = compensacion.reindex(asignado.index).fillna(0.0)
    if not np.allclose(asignado, esperado, atol=1e-4, rtol=1e-9):
        raise ErrorEntrada("RET-008: no se conserva la compensacion por cuarto de hora.")
    pagos = (
        detalle.groupby("Suministrador", as_index=False)["Pago"].sum()
        .sort_values("Suministrador", kind="stable").reset_index(drop=True)
    )
    if not np.isclose(detalle["Pago"].sum(), pagos["Pago"].sum(), atol=1e-4, rtol=1e-9):
        raise ErrorEntrada("RET-009: no se conserva la compensacion mensual.")
    cuartos = compensacion.index.astype(int)
    por_cuarto = pd.DataFrame({
        "DIA": ((cuartos - 1) // 96) + 1,
        "CUARTO HORA": cuartos,
        "Minuto": ((cuartos - 1) % 4) * 15,
        "Total Compensacion [$]": compensacion.to_numpy(),
    })
    return detalle, por_cuarto.reset_index(drop=True), pagos


def construir_compensacion_total(df_ecostos, df_re545, resumen_bess):
    """Consolida lo que recibe cada propietario desde ambos metodos."""

    propietarios = construir_mapa_propietario(resumen_bess)
    partes = []
    for df in (df_ecostos, df_re545):
        parte = pd.DataFrame({
            "Configuracion": df["Configuracion"].fillna("").astype(str).str.strip(),
            "Compensación Total [$]": pd.to_numeric(
                df["Monto a compensar"], errors="coerce"
            ).fillna(0.0),
        })
        parte["Empresa"] = parte["Configuracion"].map(
            lambda c: propietarios.get(normalizar(c), c)
        )
        partes.append(parte)
    return (
        pd.concat(partes, ignore_index=True)
        .groupby("Empresa", as_index=False)["Compensación Total [$]"].sum()
        .sort_values("Empresa", kind="stable").reset_index(drop=True)
    )


def construir_resumen(compensacion_total, pagos):
    """Une acreedores y deudores: RECIBE, PAGA y NETO por empresa."""

    recibe = compensacion_total.rename(columns={"Empresa": "NOMBRE", "Compensación Total [$]": "RECIBE"})
    paga = pagos.rename(columns={"Suministrador": "NOMBRE", "Pago": "PAGA"})
    df = recibe.merge(paga, on="NOMBRE", how="outer").fillna({"RECIBE": 0.0, "PAGA": 0.0})
    df["NETO"] = df["RECIBE"] - df["PAGA"]
    return df.sort_values("NOMBRE", kind="stable").reset_index(drop=True)
