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
  - **CSF**: de los reportes diarios del AGC (`csf_<AAAA><MM><DD>`),
    que viven TODOS JUNTOS (todos los meses) en
        \\\\nas-cen1\\D. Transferencias\\SCADA\\reporte_agc_face_NM10
    Se eligen los del mes, se copian a <CARPETA_BASE>/FD y FMA/agcface/
    y se concatenan, sin transformacion.
  - **CTF**: del `CTF_20AAMM.csv`, que esta en el mismo arbol del DCO
    que el CPF pero en otra rama:
        <version>/01 Respuesta/06 Indices CTF/
    Se le saca la zona horaria de t0/tfin.

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
import shutil
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

# Las dos ramas del arbol del DCO (la raiz y la carpeta de version las
# resuelve Indicadores_DCO, que es la misma del FD):
#
#   <version>/01 Respuesta/01 Indices CPF/20AA.MM_Respuesta_CPF/
#       Reporte diario <D>-<M>-<AAAA>/tabla_resumen_<D>_<M>_<AAAA>.xlsx
#   <version>/01 Respuesta/06 Indices CTF/CTF_<AAAA><MM>.csv
#
# Son las rutas que dio el usuario. Igual que con las carpetas de año y
# mes, se recorren comparando por nombre normalizado y, si alguna no
# aparece, se cae a una busqueda recursiva desde la carpeta de version:
# el DCO cambia de anidamiento cada tanto y no vale la pena que eso
# rompa el boton.
SUBCARPETAS_CPF = ("01 Respuesta", "01 Indices CPF")
SUBCARPETAS_CTF = ("01 Respuesta", "06 Indices CTF")

PATRON_CARPETA_RESPUESTA_CPF = "respuesta_cpf"
PLANTILLA_TABLA_RESUMEN = "tabla_resumen_{dia}_{mes}_{anio}"

# Los reportes del AGC no estan en el arbol del DCO: viven todos juntos
# (todos los meses) en esta carpeta, y de ahi hay que elegir los del
# periodo. Cuarta y ultima ruta de red del programa.
RAIZ_AGC_FACE = r"\\nas-cen1\D. Transferencias\SCADA\reporte_agc_face_NM10"

# Las entradas del FMA se copian a <CARPETA_BASE>/FD y FMA/inputs/ con
# el boton "Traer inputs": los reportes diarios de CPF y el CTF sueltos
# ahi, y los del AGC en su propia subcarpeta. Asi el "Generar" de cada
# FMA trabaja contra el disco local y no contra la red -que es lo que
# lo hacia lento-, y ademas queda registrado con que entradas se armo
# cada salida.
CARPETA_INPUTS = "inputs"

# Subcarpeta (dentro de "inputs/") donde se copian los del mes. El
# nombre lo eligio el usuario, sin guion bajo.
CARPETA_AGC_FACE = "agcface"
# Confirmado por el usuario: los reportes del AGC se llaman asi
# (ej. csf_20260301.xlsx), el mismo nombre que espera el script original.
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

# Vive en Indicadores_DCO (lo usan los dos modulos); se reexporta aca
# para no cambiar los llamados de este archivo.
bajar_por_subcarpetas = dco.bajar_por_subcarpetas


def _carpetas_hasta(carpeta, profundidad):
    """
    Las subcarpetas hasta N niveles, sin bajar mas. Se usa en vez de
    rglob() donde el arbol de abajo es enorme (los reportes diarios).
    """

    actuales = [Path(carpeta)]

    for _ in range(profundidad):

        siguientes = []

        for padre in actuales:
            for hijo in _listar(padre):
                if hijo.is_dir():
                    yield hijo
                    siguientes.append(hijo)

        actuales = siguientes


def buscar_carpeta_respuesta_cpf(carpeta_version):
    """
    Ubica la carpeta '20AA.MM_Respuesta_CPF' (la que tiene adentro un
    "Reporte diario" por dia). Primero por la ruta que dio el usuario
    ('01 Respuesta/01 Indices CPF') y, si ahi no esta, buscando
    recursivamente desde la carpeta de version.
    """

    carpeta_version = Path(carpeta_version)

    indices_cpf = bajar_por_subcarpetas(carpeta_version, SUBCARPETAS_CPF)

    if indices_cpf is not None:
        for hijo in sorted(indices_cpf.iterdir()):
            if hijo.is_dir() and PATRON_CARPETA_RESPUESTA_CPF in (
                dco._normalizar(hijo.name).replace(" ", "_")
            ):
                return hijo

    # Respaldo: si el DCO la movio de rama, se busca por nombre pero
    # SIN recorrer el arbol entero -adentro de la carpeta de version
    # estan los reportes diarios, o sea cientos de carpetas, y eso sobre
    # red es carisimo-. Tres niveles alcanzan de sobra.
    for ruta in _carpetas_hasta(carpeta_version, profundidad=3):
        if PATRON_CARPETA_RESPUESTA_CPF in (
            dco._normalizar(ruta.name).replace(" ", "_")
        ):
            return ruta

    return None


def _listar(carpeta):
    """iterdir() que no revienta si la carpeta no se puede leer."""

    try:
        return list(Path(carpeta).iterdir())
    except OSError:
        return []


def _dia_de_carpeta_reporte(nombre, mes, anio):
    """
    'Reporte diario 7-3-2026' -> 7. None si el nombre no es de una
    carpeta de reporte diario de ese mes.
    """

    texto = dco._normalizar(nombre)

    if not texto.startswith("reporte diario"):
        return None

    numeros = re.findall(r"\d+", texto)

    if len(numeros) < 3:
        return None

    dia, mes_nombre, anio_nombre = (int(n) for n in numeros[:3])

    if mes_nombre != mes or anio_nombre not in (anio, anio % 100):
        return None

    return dia


def indexar_reportes_cpf(carpeta_respuesta, anio, mes, dias=None):
    """
    dia -> ruta del `tabla_resumen` de ese dia, para todo el mes.

    **Esta funcion es la que decide si el boton tarda 20 segundos o
    varios minutos.** La version anterior hacia un `rglob` recursivo
    por CADA dia, o sea 31 recorridos completos del arbol de reportes
    sobre una carpeta de red: eso es lo que lo hacia eterno (el script
    original arma la ruta y abre el archivo, sin buscar nada).

    Ahora:

      1) se prueba la ruta EXACTA de cada dia
         (`Reporte diario <D>-<M>-<AAAA>/tabla_resumen_<D>_<M>_<AAAA>.xlsx`),
         que es una sola consulta por dia y resuelve el caso normal;
      2) solo para los dias que fallaron se lista la carpeta de
         reportes UNA vez, se ubica la carpeta de cada dia por su
         nombre y se lista esa sola carpeta, para tolerar las
         variantes (cero a la izquierda, sufijo `_UTC-4`, etc.).

    `dias` permite pedir solo algunos (lo usa la busqueda de version,
    que con encontrar uno ya sabe que esa version sirve).
    """

    carpeta_respuesta = Path(carpeta_respuesta)
    dias = list(dias if dias is not None else range(1, _dias_del_mes(anio, mes) + 1))

    encontrados = {}
    faltantes = []

    # (1) La ruta exacta, tal como la arma entradas_sscc.py.
    for dia in dias:

        directa = (
            carpeta_respuesta
            / f"Reporte diario {dia}-{mes}-{anio}"
            / (PLANTILLA_TABLA_RESUMEN.format(dia=dia, mes=mes, anio=anio)
               + ".xlsx")
        )

        if directa.is_file():
            encontrados[dia] = directa
        else:
            faltantes.append(dia)

    if not faltantes:
        return encontrados

    # (2) Un solo listado de la carpeta para los dias que faltan. De ese
    # listado salen las dos formas posibles: las carpetas "Reporte
    # diario ..." (el arbol del DCO) y los `tabla_resumen` sueltos (la
    # carpeta inputs/, donde se copian todos juntos).
    contenido = _listar(carpeta_respuesta)

    carpetas_por_dia = {}
    sueltos = []

    for hijo in contenido:
        if hijo.is_dir():
            dia = _dia_de_carpeta_reporte(hijo.name, mes, anio)
            if dia is not None:
                carpetas_por_dia.setdefault(dia, hijo)
        elif hijo.suffix.lower() in EXTENSIONES_EXCEL:
            sueltos.append(hijo)

    for dia in faltantes:

        carpeta_dia = carpetas_por_dia.get(dia)

        ruta = (
            _tabla_resumen_en(carpeta_dia, anio, mes, dia)
            if carpeta_dia is not None else None
        )

        if ruta is None:
            ruta = _elegir_tabla_resumen(sueltos, anio, mes, dia)

        if ruta is not None:
            encontrados[dia] = ruta

    return encontrados


def _tabla_resumen_en(carpeta_dia, anio, mes, dia):
    """
    El `tabla_resumen` del dia dentro de SU carpeta (un solo listado,
    sin recorrer nada mas). Se acepta cualquier sufijo despues del año,
    porque hay variantes con la zona horaria en el nombre
    (`..._UTC-4.xlsx`, `..._utc-3.xlsx`); gana el nombre "pelado" y, si
    no esta, el mas reciente.
    """

    return _elegir_tabla_resumen(
        [r for r in _listar(carpeta_dia) if r.is_file()], anio, mes, dia
    )


def _elegir_tabla_resumen(archivos, anio, mes, dia):
    """El `tabla_resumen` del dia entre una lista de archivos ya leida."""

    prefijo = dco._normalizar(
        PLANTILLA_TABLA_RESUMEN.format(dia=dia, mes=mes, anio=anio)
    )

    candidatos = [
        r for r in archivos
        if r.suffix.lower() in EXTENSIONES_EXCEL
        and dco._normalizar(r.stem).startswith(prefijo)
    ]

    if not candidatos:
        return None

    exactos = [r for r in candidatos if dco._normalizar(r.stem) == prefijo]

    if exactos:
        return exactos[0]

    return max(candidatos, key=lambda r: r.stat().st_mtime)


def buscar_tabla_resumen(carpeta_respuesta, anio, mes, dia):
    """El `tabla_resumen` de un dia suelto (ver indexar_reportes_cpf)."""

    return indexar_reportes_cpf(
        carpeta_respuesta, anio, mes, dias=[dia]
    ).get(dia)


def buscar_reportes_cpf(carpeta_version, aamm):
    """
    La carpeta de reportes de CPF de una version, PERO solo si de
    verdad tiene adentro al menos un `tabla_resumen` del periodo.

    Ese "solo si" es el punto: puede existir la carpeta V2 y estar
    vacia (recien creada, a medio subir), y en ese caso hay que seguir
    buscando en V1. Devuelve None si no sirve.

    Alcanza con encontrar UN dia, asi que no se indexa el mes entero.
    """

    carpeta_respuesta = buscar_carpeta_respuesta_cpf(carpeta_version)

    if carpeta_respuesta is None:
        return None

    anio, mes = dco.periodo_desde_aamm(aamm)

    if indexar_reportes_cpf(carpeta_respuesta, anio, mes):
        return carpeta_respuesta

    return None


def buscar_ctf(carpeta_version, aamm):
    """
    El `CTF_<AAAA><MM>.csv` de una version: primero por la ruta que dio
    el usuario ('01 Respuesta/06 Indices CTF') y, si ahi no esta,
    buscando recursivamente desde la carpeta de version. Devuelve None
    si esa version no lo tiene.
    """

    anio, mes = dco.periodo_desde_aamm(aamm)
    prefijo = PLANTILLA_CTF.format(anio=anio, mes=mes)

    carpeta_ctf = bajar_por_subcarpetas(carpeta_version, SUBCARPETAS_CTF)

    if carpeta_ctf is not None:
        encontrado = _buscar_en(
            [carpeta_ctf], prefijo, (".csv",), recursivo=True
        )
        if encontrado:
            return encontrado

    return _buscar_en([carpeta_version], prefijo, (".csv",), recursivo=True)


def construir_fma_cpf(carpeta_respuesta, aamm, registrar=print):
    """
    Arma el equivalente de `fma_cpf_AAMM.xlsx` recorriendo los reportes
    diarios. Replica calc_fmacpf de entradas_sscc.py.
    """

    anio, mes = dco.periodo_desde_aamm(aamm)
    dias_del_mes = _dias_del_mes(anio, mes)

    # Un solo indice para todo el mes (ver indexar_reportes_cpf): antes
    # esto buscaba dia por dia recorriendo el arbol entero cada vez.
    reportes = indexar_reportes_cpf(carpeta_respuesta, anio, mes)

    partes = []
    dias_sin_archivo = []

    for dia in range(1, dias_del_mes + 1):

        ruta = reportes.get(dia)

        if ruta is None:
            dias_sin_archivo.append(dia)
            continue

        registrar(f"  leyendo {ruta.name}...")

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

def ruta_agc_face(raiz=None):
    """Carpeta de red donde estan TODOS los reportes del AGC."""

    return Path(raiz or RAIZ_AGC_FACE)


def carpeta_inputs(carpeta_destino):
    """<CARPETA_BASE>/FD y FMA/ -> <...>/FD y FMA/inputs/"""

    return Path(carpeta_destino) / CARPETA_INPUTS


def carpeta_agcface(carpeta_destino):
    """
    <CARPETA_BASE>/FD y FMA/ -> <...>/FD y FMA/inputs/agcface/

    Si un caso viejo todavia tiene la de antes (colgando directo de
    "FD y FMA/") y la nueva no existe, se usa esa.
    """

    nueva = carpeta_inputs(carpeta_destino) / CARPETA_AGC_FACE

    if not nueva.is_dir():
        vieja = Path(carpeta_destino) / CARPETA_AGC_FACE
        if vieja.is_dir():
            return vieja

    return nueva


def traer_agc_face(carpeta_destino, aamm, raiz=None, registrar=print):
    """
    Copia a <CARPETA_BASE>/FD y FMA/agcface/ los reportes del AGC del
    periodo. En la carpeta de red estan TODOS los meses juntos, asi que
    la gracia es elegir los del mes que corresponde.

    El nombre es `csf_<AAAA><MM><DD>` (confirmado por el usuario: p.ej.
    `csf_20260301`), el mismo que espera el script original.

    Devuelve (copiados, salteados, carpeta_destino_agcface).
    """

    anio, mes = dco.periodo_desde_aamm(aamm)
    origen = ruta_agc_face(raiz)

    if not origen.is_dir():
        raise ErrorIndicesFma(
            f"No se puede leer la carpeta de los reportes del AGC:\n"
            f"  {origen}\n\n"
            f"Revisa que tengas conexion y permiso sobre ese servidor. "
            f"Si la ruta cambio, se cambia en RAIZ_AGC_FACE "
            f"(Script/Fd/Indices_FMA.py)."
        )

    destino = carpeta_agcface(carpeta_destino)

    try:
        destino.mkdir(parents=True, exist_ok=True)
    except OSError as error:
        raise ErrorIndicesFma(
            f"No se pudo crear {destino}: {error}"
        ) from error

    try:
        archivos = [a for a in origen.iterdir() if a.is_file()]
    except OSError as error:
        raise ErrorIndicesFma(
            f"No se pudo leer {origen}: {error}"
        ) from error

    dias_del_mes = _dias_del_mes(anio, mes)

    prefijos = {
        dco._normalizar(
            PLANTILLA_CSF_DIARIO.format(anio=anio, mes=mes, dia=dia)
        )
        for dia in range(1, dias_del_mes + 1)
    }

    del_mes = [
        a for a in archivos
        if a.suffix.lower() in EXTENSIONES_EXCEL
        and any(dco._normalizar(a.stem).startswith(p) for p in prefijos)
    ]

    if not del_mes:
        raise ErrorIndicesFma(
            f"No se encontro ningun reporte del AGC de {anio}-{mes:02d} "
            f"en {origen}.\n\n"
            f"Se buscan los que se llaman "
            f"'{PLANTILLA_CSF_DIARIO.format(anio=anio, mes=mes, dia=1)}' "
            f"y siguientes."
        )

    registrar(f"  reportes del AGC del periodo: {len(del_mes)}")

    copiados, salteados = [], []

    for ruta in sorted(del_mes, key=lambda r: r.name):

        ruta_destino = destino / ruta.name

        if (
            ruta_destino.is_file()
            and ruta_destino.stat().st_size == ruta.stat().st_size
            and int(ruta_destino.stat().st_mtime) == int(ruta.stat().st_mtime)
        ):
            salteados.append(ruta.name)
            continue

        try:
            shutil.copy2(ruta, ruta_destino)
        except OSError as error:
            raise ErrorIndicesFma(
                f"No se pudo copiar {ruta.name}: {error}"
            ) from error

        copiados.append(ruta.name)

    registrar(
        f"  {len(copiados)} copiado(s), {len(salteados)} ya estaban al dia "
        f"en {destino}"
    )

    return copiados, salteados, destino


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


TIPOS_FMA = ("cpf", "csf", "ctf")


def traer_inputs(
    carpeta_destino, aamm, version=None, raiz=None, raiz_agc=None,
    registrar=print,
):
    """
    Copia a <CARPETA_BASE>/FD y FMA/inputs/ TODAS las entradas del FMA
    del periodo (boton "Traer inputs"):

        inputs/
            tabla_resumen_<D>_<M>_<AAAA>.xlsx   <- los reportes de CPF
            CTF_<AAAA><MM>.csv                  <- la entrada del CTF
            agcface/
                csf_<AAAA><MM><DD>.xlsx         <- los reportes del AGC

    Sirve para dos cosas: deja registrado con que entradas se armo cada
    salida, y hace que el "Generar" de cada FMA trabaje contra el disco
    local en vez de la carpeta de red.

    Devuelve (resumen, faltantes), donde resumen es un dict por tipo con
    la cuenta de archivos que quedaron.
    """

    anio, mes = dco.periodo_desde_aamm(aamm)
    carpeta_destino = Path(carpeta_destino)

    if not carpeta_destino.is_dir():
        raise ErrorIndicesFma(f"No se encontro la carpeta {carpeta_destino}")

    destino = carpeta_inputs(carpeta_destino)

    try:
        destino.mkdir(parents=True, exist_ok=True)
    except OSError as error:
        raise ErrorIndicesFma(
            f"No se pudo crear {destino}: {error}"
        ) from error

    carpeta_publicacion = dco.carpeta_del_periodo(aamm, raiz=raiz)

    resumen, faltantes = {}, []

    # ---- reportes diarios de CPF ----
    carpeta_version, carpeta_respuesta, revisadas = dco.buscar_en_versiones(
        carpeta_publicacion,
        lambda cv: buscar_reportes_cpf(cv, aamm),
        version=version,
        registrar=registrar,
    )

    if carpeta_respuesta is None:
        faltantes.append(
            f"CPF (ninguna de las versiones revisadas tiene reportes; "
            f"revisada(s): {', '.join(revisadas)})"
        )
    else:
        registrar(f"  reportes de CPF: {carpeta_respuesta}")
        reportes = indexar_reportes_cpf(carpeta_respuesta, anio, mes)
        resumen["cpf"] = _copiar_todos(
            sorted(reportes.values(), key=lambda r: r.name), destino, registrar
        )

    # ---- CTF ----
    carpeta_version_ctf, ruta_ctf, revisadas_ctf = dco.buscar_en_versiones(
        carpeta_publicacion,
        lambda cv: buscar_ctf(cv, aamm),
        version=version,
        registrar=registrar,
    )

    if ruta_ctf is None:
        faltantes.append(
            f"CTF (ninguna de las versiones revisadas tiene "
            f"'{PLANTILLA_CTF.format(anio=anio, mes=mes)}*.csv'; "
            f"revisada(s): {', '.join(revisadas_ctf)})"
        )
    else:
        registrar(f"  CTF: {ruta_ctf}")
        resumen["ctf"] = _copiar_todos([ruta_ctf], destino, registrar)

    # ---- reportes del AGC ----
    try:
        copiados, salteados, _ = traer_agc_face(
            carpeta_destino, aamm, raiz=raiz_agc, registrar=registrar
        )
        resumen["csf"] = len(copiados) + len(salteados)
    except ErrorIndicesFma as error:
        faltantes.append(f"CSF ({error})")

    if not resumen:
        raise ErrorIndicesFma(
            "No se pudo traer ninguna entrada de FMA:\n  - "
            + "\n  - ".join(faltantes)
        )

    return resumen, faltantes


def _hay_csf_del_periodo(carpeta, anio, mes):
    """True si la carpeta ya tiene algun csf_ del periodo."""

    prefijo = dco._normalizar(f"csf_{anio}{mes:02d}")

    return any(
        r.is_file()
        and r.suffix.lower() in EXTENSIONES_EXCEL
        and dco._normalizar(r.stem).startswith(prefijo)
        for r in _listar(carpeta)
    )


def _copiar_todos(rutas, destino, registrar=print):
    """
    Copia una lista de archivos a la carpeta destino, salteando los que
    ya estan al dia. Devuelve cuantos quedaron.
    """

    copiados = salteados = 0

    for ruta in rutas:

        ruta_destino = Path(destino) / ruta.name

        if (
            ruta_destino.is_file()
            and ruta_destino.stat().st_size == ruta.stat().st_size
            and int(ruta_destino.stat().st_mtime) == int(ruta.stat().st_mtime)
        ):
            salteados += 1
            continue

        try:
            shutil.copy2(ruta, ruta_destino)
        except OSError as error:
            raise ErrorIndicesFma(
                f"No se pudo copiar {ruta.name}: {error}"
            ) from error

        copiados += 1

    registrar(
        f"  {copiados} copiado(s), {salteados} ya estaban al dia"
    )

    return copiados + salteados


def generar_fma(
    carpeta_destino,
    aamm,
    tipos=None,
    version=None,
    raiz=None,
    raiz_agc=None,
    registrar=print,
):
    """
    Arma las salidas de FMA pedidas y las escribe en
    <CARPETA_BASE>/FD y FMA/. Cada una tiene su propio boton "Generar"
    en la ventana, asi que lo normal es que venga un solo tipo.

    tipos: subconjunto de TIPOS_FMA; None = las tres.

    De donde sale cada una (rutas confirmadas por el usuario):
      - cpf: <version>/01 Respuesta/01 Indices CPF/
                 20AA.MM_Respuesta_CPF/Reporte diario .../tabla_resumen_...
      - csf: los reportes del AGC de RAIZ_AGC_FACE, copiados primero a
             'FD y FMA/agcface/'
      - ctf: <version>/01 Respuesta/06 Indices CTF/CTF_<AAAA><MM>.csv

    El CPF y el CTF necesitan la carpeta de version del DCO; el CSF no
    (su origen es otro servidor), asi que pedir SOLO csf no obliga a
    que el DCO este publicado.

    Devuelve (escritos, faltantes, versiones), donde `versiones` dice
    que version del DCO termino usando cada tipo (pueden ser distintas).
    """

    anio, mes = dco.periodo_desde_aamm(aamm)
    carpeta_destino = Path(carpeta_destino)

    if not carpeta_destino.is_dir():
        raise ErrorIndicesFma(f"No se encontro la carpeta {carpeta_destino}")

    tipos = set(tipos or TIPOS_FMA)
    desconocidos = tipos - set(TIPOS_FMA)

    if desconocidos:
        raise ErrorIndicesFma(f"Tipo(s) de FMA desconocido(s): {desconocidos}")

    # La carpeta de publicacion se resuelve una sola vez, pero la
    # VERSION la elige cada tipo por su cuenta: puede estar publicada V2
    # y tener el CPF pero no el CTF (o al reves), y cada uno tiene que
    # usar la version mas alta que tenga SU archivo -pedido explicito
    # del usuario-. Por eso `versiones` es un dict por tipo.
    # La carpeta del DCO se resuelve SOLO si de verdad hay que ir a
    # buscar algo alla: con las entradas ya traidas a inputs/, el
    # "Generar" no toca la red (ni avisa de que no la alcanza).
    publicacion = {"resuelta": False, "carpeta": None}

    def carpeta_publicacion_del_dco():

        if not publicacion["resuelta"]:
            publicacion["resuelta"] = True
            try:
                publicacion["carpeta"] = dco.carpeta_del_periodo(
                    aamm, raiz=raiz
                )
            except dco.ErrorFd as error:
                registrar(f"  [AVISO] {error}")

        return publicacion["carpeta"]

    versiones = {}

    escritos, faltantes = {}, []

    # ---- CPF ----
    if "cpf" in tipos:

        # Primero lo que ya se trajo a inputs/ (disco local); si ahi no
        # hay nada, se va al DCO.
        carpeta_respuesta = None
        carpeta_version = None
        revisadas = []

        locales = carpeta_inputs(carpeta_destino)

        if indexar_reportes_cpf(locales, anio, mes):
            carpeta_respuesta = locales
            registrar(f"  reportes de CPF: {locales} (ya traidos)")
        else:
            carpeta_publicacion = carpeta_publicacion_del_dco()

            if carpeta_publicacion is not None:
                carpeta_version, carpeta_respuesta, revisadas = (
                    dco.buscar_en_versiones(
                        carpeta_publicacion,
                        lambda cv: buscar_reportes_cpf(cv, aamm),
                        version=version,
                        registrar=registrar,
                    )
                )

        if carpeta_respuesta is None:
            faltantes.append(
                f"CPF (no hay reportes 'tabla_resumen' del periodo ni en "
                f"{carpeta_inputs(carpeta_destino)} ni en el DCO bajo "
                f"{'/'.join(SUBCARPETAS_CPF)}"
                + (f"; version(es) revisada(s): {', '.join(revisadas)}"
                   if revisadas else "")
                + ")"
            )
        else:
            if carpeta_version is not None:
                versiones["cpf"] = carpeta_version.name
                registrar(f"  reportes diarios de CPF: {carpeta_respuesta}")
            df_cpf = construir_fma_cpf(carpeta_respuesta, aamm, registrar)
            escritos["cpf"] = _escribir(
                df_cpf, carpeta_destino / nombre_salida("cpf", aamm)
            )

    # ---- CSF ----
    if "csf" in tipos:

        try:
            carpeta_ya_traida = carpeta_agcface(carpeta_destino)

            if _hay_csf_del_periodo(carpeta_ya_traida, anio, mes):
                carpeta_csf = carpeta_ya_traida
                registrar(
                    f"  reportes del AGC: {carpeta_csf} (ya traidos)"
                )
            else:
                _, _, carpeta_csf = traer_agc_face(
                    carpeta_destino, aamm, raiz=raiz_agc, registrar=registrar
                )
        except ErrorIndicesFma as error:
            # Si no se pudo traer de la red, se intenta igual con lo que
            # ya haya copiado en agcface/: puede ser una corrida
            # anterior, o archivos puestos a mano.
            carpeta_csf = carpeta_agcface(carpeta_destino)

            if not carpeta_csf.is_dir():
                faltantes.append(f"CSF ({error})")
                carpeta_csf = None
            else:
                registrar(
                    f"  [AVISO] no se pudieron traer los reportes del "
                    f"AGC ({error}); se usa lo que ya hay en "
                    f"{carpeta_csf}"
                )

        if carpeta_csf is not None:
            df_csf = construir_fma_csf(carpeta_csf, aamm, registrar)
            escritos["csf"] = _escribir(
                df_csf, carpeta_destino / nombre_salida("csf", aamm)
            )

    # ---- CTF ----
    if "ctf" in tipos:

        prefijo_ctf = PLANTILLA_CTF.format(anio=anio, mes=mes)

        carpeta_version = None
        revisadas = []

        # Primero lo que ya se trajo a inputs/ (o lo que el usuario haya
        # dejado a mano en la carpeta del caso).
        ruta_ctf = _buscar_en(
            [carpeta_inputs(carpeta_destino), carpeta_destino],
            prefijo_ctf, (".csv",),
        )

        if ruta_ctf is None:
            carpeta_publicacion = carpeta_publicacion_del_dco()

            if carpeta_publicacion is not None:
                carpeta_version, ruta_ctf, revisadas = dco.buscar_en_versiones(
                    carpeta_publicacion,
                    lambda cv: buscar_ctf(cv, aamm),
                    version=version,
                    registrar=registrar,
                )

        if ruta_ctf is None:
            faltantes.append(
                f"CTF (ninguna de las versiones revisadas tiene "
                f"'{prefijo_ctf}*.csv' bajo "
                f"{'/'.join(SUBCARPETAS_CTF)}, ni esta en "
                f"{carpeta_destino}; revisada(s): {', '.join(revisadas)})"
            )
        else:
            if carpeta_version is not None:
                versiones["ctf"] = carpeta_version.name

            registrar(f"  CTF: {ruta_ctf}")

            # Se copia a la carpeta del caso antes de usarlo, para que
            # quede registrado con que archivo se armo la salida.
            if ruta_ctf.parent != carpeta_destino:
                try:
                    shutil.copy2(ruta_ctf, carpeta_destino / ruta_ctf.name)
                except OSError as error:
                    registrar(
                        f"  [AVISO] no se pudo copiar {ruta_ctf.name} a la "
                        f"carpeta del caso: {error}"
                    )

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
            "No se pudo armar ninguna de las salidas de FMA pedidas:\n  - "
            + "\n  - ".join(faltantes)
        )

    return escritos, faltantes, versiones


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
