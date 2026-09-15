# -*- coding: utf-8 -*-
"""
Rutas del caso y busqueda de los archivos de entrada.
"""

import re
from pathlib import Path

from .parametros import (
    ARCHIVO_CENTRALES, ARCHIVO_CMG, ARCHIVO_CONTROL, ARCHIVO_MEDIDAS_SAE,
    ARCHIVO_SALIDA, CARPETA_AUXILIARES, CARPETA_CMG,
    CARPETA_DB_SUBASTAS, CARPETA_FD_FMA, CARPETA_FD_FMA_ANTIGUA,
    CARPETA_MEDIDAS, CARPETA_OFERTAS, CARPETA_PRORRATA_RETIROS,
    CARPETA_SUBASTAS,
    CARPETA_TRABAJO_MEDIDAS, EXTENSIONES_EXCEL, PATRON_AAMM,
    PATRON_NOMBRE_OFERTAS, PATRON_NOMBRE_SSCC_DESEMPENO,
)
from .utiles import ErrorEntrada, normalizar


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
    # La carpeta se llama "FD y FMA"; si un caso viejo todavia tiene la
    # de nombre anterior y no la nueva, se usa esa.
    fd_fma_dir = base / CARPETA_FD_FMA

    if not fd_fma_dir.is_dir() and (base / CARPETA_FD_FMA_ANTIGUA).is_dir():
        fd_fma_dir = base / CARPETA_FD_FMA_ANTIGUA
    subastas_dir = base / CARPETA_SUBASTAS
    prorrata_retiros_dir = base / CARPETA_PRORRATA_RETIROS

    return {
        "base": base,
        "medidas_dir": medidas_dir,
        "auxiliares_dir": auxiliares_dir,
        "ofertas_dir": ofertas_dir,
        "cmg_dir": cmg_dir,
        "sscc_desempeno_dir": fd_fma_dir,
        "subastas_dir": subastas_dir,
        "prorrata_retiros_dir": prorrata_retiros_dir,
        "db_subastas_dir": subastas_dir / CARPETA_DB_SUBASTAS,
        "medidas_sae": medidas_dir / ARCHIVO_MEDIDAS_SAE,
        "trabajo_medidas": medidas_dir / CARPETA_TRABAJO_MEDIDAS,
        "centrales": auxiliares_dir / ARCHIVO_CENTRALES,
        "cmg": cmg_dir / ARCHIVO_CMG,
        # Una sola planilla (entradas + calculo) y, al lado, el
        # archivo de control de la corrida.
        "salida": base / ARCHIVO_SALIDA,
        "control": base / ARCHIVO_CONTROL,
    }


# ============================================================
# CREAR EL CASO DE UN PERIODO NUEVO
# ============================================================

# Las carpetas que tiene que tener un caso, en el orden en que se usan.
# Es la misma estructura que revisa revisar_estructura() y que dibuja
# la ventana: si aca falta una, el caso nace incompleto.
#
# "Subastas/DB subastas" va con su carpeta adentro porque la ventana ya
# la creaba sola al revisar (ofertas_adj.asegurar_carpeta_db); crearla
# aca deja el caso entero armado de una vez.
#
# NO se crea Medidas/_trabajo: son los andamios de la descarga, los
# arma generar_medidas_sae cuando corre.
SUBCARPETAS_CASO = (
    CARPETA_MEDIDAS,
    CARPETA_AUXILIARES,
    CARPETA_OFERTAS,
    CARPETA_CMG,
    CARPETA_FD_FMA,
    CARPETA_SUBASTAS,
    f"{CARPETA_SUBASTAS}/{CARPETA_DB_SUBASTAS}",
    CARPETA_PRORRATA_RETIROS,
)


def crear_estructura_caso(carpeta_base, registrar=print):
    """
    Crea la carpeta base de un caso y todas sus subcarpetas
    (SUBCARPETAS_CASO). Devuelve la lista de las que se crearon (las
    que ya estaban no se tocan ni se cuentan).

    Pedido del usuario: *"cuando yo elija un mes que no exista, me
    permita elegir y crear la carpeta con las carpetas dentro"*.

    Es idempotente a proposito: sirve tanto para un periodo nuevo
    (crea todo) como para completar un caso al que le falta una
    carpeta (crea solo esa). Nunca borra ni mueve nada.

    Un caso VIEJO que tiene la carpeta con el nombre anterior
    ("SSCC_Desempeño") no recibe una "FD y FMA" vacia al lado: se
    respeta la que ya esta usando (ver resolver_rutas).
    """

    base = Path(carpeta_base)

    if base.exists() and not base.is_dir():
        raise ErrorEntrada(
            f"{base} existe y no es una carpeta: elegi otro nombre."
        )

    usa_nombre_viejo = (
        (base / CARPETA_FD_FMA_ANTIGUA).is_dir()
        and not (base / CARPETA_FD_FMA).is_dir()
    )

    creadas = []

    try:
        if not base.is_dir():
            base.mkdir(parents=True)
            creadas.append(base)

        for relativa in SUBCARPETAS_CASO:

            if relativa == CARPETA_FD_FMA and usa_nombre_viejo:
                continue

            carpeta = base / relativa

            if not carpeta.is_dir():
                carpeta.mkdir(parents=True)
                creadas.append(carpeta)

    except OSError as error:
        raise ErrorEntrada(
            f"No se pudo crear la estructura del caso en {base}: {error}"
        ) from error

    for carpeta in creadas:
        registrar(f"  creada: {carpeta}")

    if not creadas:
        registrar(f"  {base} ya tenia todas sus carpetas.")

    return creadas


def nombre_caso_sugerido(aamm, carpeta_modelo=None):
    """
    Como llamar a la carpeta del periodo 'aamm', a partir de como se
    llama la del periodo que el usuario tenia abierto.

    Si esa carpeta trae un AAMM en el nombre, se cambia SOLO ese
    pedazo y se conserva el resto ("Balance BESS 2607" -> "Balance
    BESS 2608"): cada equipo nombra sus casos a su manera y no hay una
    convencion que imponer. Si no trae ninguno (o no hay carpeta
    previa), se propone el AAMM pelado.

    Es solo una SUGERENCIA: la ventana la muestra en un campo que el
    usuario puede cambiar antes de crear nada.
    """

    aamm = validar_aamm(aamm)

    if carpeta_modelo:
        nombre = Path(carpeta_modelo).name
        cambiado, cuantos = re.subn(r"(?<!\d)\d{4}(?!\d)", aamm, nombre, count=1)
        if cuantos and cambiado != nombre:
            return cambiado

    return aamm


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


def periodo_desde_aamm(aamm):
    """'2607' -> (2026, 7)"""

    anio = 2000 + int(aamm[:2])
    mes = int(aamm[2:])

    if not 1 <= mes <= 12:
        raise ErrorEntrada(
            f"El periodo '{aamm}' no tiene un mes valido."
        )

    return anio, mes
