# -*- coding: utf-8 -*-
"""
Balance BESS / SSCC.

Ventana unica: se elige la carpeta base del caso, se ingresa el
periodo (AAMM) y debajo se dibuja el diagrama de la estructura del
caso con el estado de cada entrada (OK/FALTA/PENDIENTE).

Todo lo que el programa puede hacer sale de un boton en la fila que
corresponde -no hay ventanas intermedias ni un boton "Ejecutar"
unico-:

    <CARPETA_BASE>/
        Medidas/
            Medidas_SAE.xlsx               [Actualizar]
            <algo>SOC<algo>AAMM<algo>.xlsx
        Auxiliares/
            Centrales.xlsx
                hoja 'Resumen BESS'
                hoja 'Diccionario'
            <algo>Homologacion<algo>.xlsx
                hoja 'homol'
                hoja 'Gen real'
        Ofertas/
            <algo>OfertasSSCC<algo>.xlsx (o .xlsm/.xlsb/.xls)
        Cmg/
            cmg<AAMM>_def_15minutal.csv    [Traer cmg_15min]
            cmg.xlsx                       [Generar]
        FD y FMA/                          [Traer FD]
            SSCC_Desempeño_<algo>.xlsx (o .xlsm/.xlsb/.xls)
            fma_cpf_<AAMM>.xlsx
            fma_csf_<AAMM>.xlsx
            fma_cft_<AAMM>.xlsx (o .csv)
        Subastas/
            DB subastas/                   [Traer subastas]
            3_REMUNERACIÓN_SUBASTAS_E_ID_<algo>.xlsx (respaldo)
        Consolidado_entradas.xlsx          [Actualizar todo]
            hoja 'Medidores'               [Actualizar]
            hoja 'Ofertas SSCC'            [Actualizar]
            hoja 'CMg'                     [Actualizar]
            hoja 'FD'                      [Actualizar]
            hoja 'Subastas'                [Actualizar]
        Pagos_BESS.xlsx                    [Calcular todo]
            hoja 'Calculo E Costos'        [Calcular]
            hoja 'Calculo RE545'           [Calcular]

Las dos salidas se desglosan por hoja igual que Centrales.xlsx: cada
hoja se actualiza sola, y lo que no se toca se conserva tal cual
estaba en el archivo. Si el archivo todavia no existe, se crea al
actualizar la primera hoja.

Medidas/_trabajo/ (los lotes que baja la API, la marca de
reanudacion) NO aparece en el diagrama a pedido del usuario: no es una
entrada ni una salida del caso, son andamios del proceso.

El calculo vive en Script/ (ver Script/__init__.py), incluida la clave
de las dos APIs del Coordinador que usa Medidas (Script/Medidas/
comun.py, USER_KEY). La ubicacion de este .py no influye en nada salvo
en donde se guarda config.json.
"""

import json
import os
import socket
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from Script import nucleo


CONFIG_PATH = Path(__file__).parent / "config.json"

COLOR_OK = "#1a6b1a"
COLOR_FALTA = "#b00020"
COLOR_PENDIENTE = "#a06000"
COLOR_NEUTRO = "#555555"
COLOR_BASE_OK = "#1a4fb0"

SIMBOLO = {
    "ok": "OK",
    "falta": "FALTA",
    "pendiente": "PENDIENTE",
}

COLOR_ESTADO = {
    "ok": COLOR_OK,
    "falta": COLOR_FALTA,
    "pendiente": COLOR_PENDIENTE,
}

# Anchos de las columnas del diagrama. La de acciones (los botones)
# va ANTES del detalle, a pedido del usuario, y por eso necesita un
# ancho fijo en pixeles: si no, cada fila correria el detalle segun el
# largo de su boton.
ANCHO_ESTRUCTURA = 52      # caracteres (Consolas 9)
ANCHO_ESTADO = 11          # caracteres
ANCHO_ACCION = 150         # pixeles
ALTO_ACCION = 26           # pixeles


# ============================================================
# CONFIG POR PC/USUARIO
# ============================================================

def get_usuario():
    usuario = (
        os.environ.get("USERNAME")
        or os.environ.get("USER")
        or "desconocido"
    )
    return f"{socket.gethostname()}_{usuario}"


def leer_config():
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f).get(get_usuario(), {})
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def guardar_config(data):
    todo = {}
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                todo = json.load(f)
        except (json.JSONDecodeError, OSError):
            todo = {}
    todo.setdefault(get_usuario(), {}).update(data)
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(todo, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def abrir_en_explorador(ruta, es_archivo=False):
    p = Path(ruta)
    if not p.exists():
        return
    carpeta = p.parent if es_archivo else p
    if sys.platform == "win32":
        subprocess.Popen(["explorer", str(carpeta)])
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(carpeta)])
    else:
        subprocess.Popen(["xdg-open", str(carpeta)])


def formato_tiempo(segundos):
    minutos, seg = divmod(int(segundos), 60)
    horas, minutos = divmod(minutos, 60)
    return f"{horas:02d}:{minutos:02d}:{seg:02d}"


# ============================================================
# DIAGRAMA DE CARPETAS (arbol de texto tipo consola)
#
# revisar_estructura() devuelve una lista plana de filas con su
# NIVEL (0 = raiz del caso, 1 = adentro de una carpeta/archivo, 2 =
# un nivel mas). Aca solo se traduce esa lista a prefijos de consola
# (├── / └── / │): nucleo.py sabe de estructura, no de dibujo.
# ============================================================

def _es_ultimo_en_su_nivel(niveles, i):
    """
    True si, mirando hacia adelante desde i, se sube de nivel antes
    de encontrar otra fila con el MISMO nivel (o se llega al final de
    la lista): o sea, si i es la ultima de su grupo.
    """
    nivel = niveles[i]
    for j in range(i + 1, len(niveles)):
        if niveles[j] < nivel:
            return True
        if niveles[j] == nivel:
            return False
    return True


def _prefijos_arbol(niveles):
    """
    Prefijos tipo consola (├── / └── / │) para una lista plana de
    niveles (0 = raiz), calculando el relleno de cada ancestro segun
    si ESE ancestro es o no el ultimo de su propio grupo.
    """
    n = len(niveles)
    ultimos = [_es_ultimo_en_su_nivel(niveles, i) for i in range(n)]
    prefijos = []

    for i, nivel in enumerate(niveles):

        if nivel == 0:
            prefijos.append("")
            continue

        relleno = ""
        for ancestro_nivel in range(1, nivel):
            indice_ancestro = next(
                (k for k in range(i - 1, -1, -1) if niveles[k] == ancestro_nivel),
                None,
            )
            relleno += (
                "    "
                if (indice_ancestro is None or ultimos[indice_ancestro])
                else "│   "
            )

        prefijos.append(relleno + ("└── " if ultimos[i] else "├── "))

    return prefijos


# ============================================================
# VENTANA
# ============================================================

def main():

    cfg = leer_config()

    root = tk.Tk()
    root.title("Balance BESS / SSCC")
    root.geometry("1180x820")

    var_base = tk.StringVar(value=cfg.get("carpeta_base", ""))
    var_aamm = tk.StringVar(value=cfg.get("aamm", ""))
    var_estado = tk.StringVar(value="Listo")
    var_tiempo = tk.StringVar(value="00:00:00")

    # Los botones viven DENTRO de las filas del arbol, que se repinta
    # entero en cada revisar(): se guardan por id de fila para poder
    # deshabilitarlos mientras corre algo, y se renuevan en cada
    # pintado.
    botones_arbol = {}
    corriendo = {"activo": False}

    # --------------------------------------------------------
    # BOTONES FIJOS ABAJO (primero, para que no los tape nada)
    # --------------------------------------------------------

    frame_botones = tk.Frame(root)
    frame_botones.pack(side="bottom", fill="x", pady=10)

    # --------------------------------------------------------
    # CANVAS CON SCROLL
    # --------------------------------------------------------

    canvas = tk.Canvas(root, borderwidth=0, highlightthickness=0)
    scroll = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scroll.set)
    scroll.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    contenedor = tk.Frame(canvas)
    ventana_canvas = canvas.create_window((0, 0), window=contenedor, anchor="nw")

    def ajustar(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(ventana_canvas, width=canvas.winfo_width())

    contenedor.bind("<Configure>", ajustar)
    canvas.bind("<Configure>", ajustar)
    canvas.bind_all(
        "<MouseWheel>",
        lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"),
    )

    # --------------------------------------------------------
    # SELECTOR DE CARPETA BASE
    # --------------------------------------------------------

    frame_carpeta = tk.LabelFrame(
        contenedor, text="Carpeta base del caso", padx=10, pady=8
    )
    frame_carpeta.pack(fill="x", padx=20, pady=(14, 6))

    lbl_base = tk.Label(
        frame_carpeta,
        textvariable=var_base,
        wraplength=940,
        justify="center",
        cursor="hand2",
        font=("Segoe UI", 9),
    )
    lbl_base.pack()

    lbl_base.bind(
        "<Button-1>",
        lambda e: abrir_en_explorador(var_base.get()) if var_base.get() else None,
    )

    # --------------------------------------------------------
    # PERIODO DEL CASO (AAMM)
    #
    # Se ingresa aca arriba y ya no aparece como fila del diagrama:
    # no es parte de la estructura de carpetas. De el dependen dos
    # filas (el SoC dentro de Medidas/ y el CSV dentro de Cmg/), que
    # lo dicen en su propio detalle cuando falta.
    # --------------------------------------------------------

    frame_periodo = tk.LabelFrame(
        contenedor, text="Periodo del caso (AAMM)", padx=10, pady=8
    )
    frame_periodo.pack(fill="x", padx=20, pady=6)

    validador_aamm = (
        root.register(lambda s: s == "" or (s.isdigit() and len(s) <= 4)),
        "%P",
    )

    entry_aamm = tk.Entry(
        frame_periodo,
        textvariable=var_aamm,
        width=10,
        justify="center",
        font=("Segoe UI", 10, "bold"),
        validate="key",
        validatecommand=validador_aamm,
    )
    entry_aamm.pack(side="left")

    tk.Label(
        frame_periodo,
        text=(
            "4 digitos: año+mes simplificado. Ej. 2607 para julio de "
            "2026. Con el se buscan el archivo de SoC dentro de "
            "Medidas/ y el CSV 15-minutal dentro de Cmg/."
        ),
        fg=COLOR_NEUTRO,
        font=("Segoe UI", 8),
        wraplength=820,
        justify="left",
    ).pack(side="left", padx=(10, 0))

    # --------------------------------------------------------
    # DIAGRAMA DE LA ESTRUCTURA DEL CASO
    # --------------------------------------------------------

    frame_arbol = tk.LabelFrame(
        contenedor, text="Estructura del caso", padx=8, pady=8
    )
    frame_arbol.pack(fill="both", expand=True, padx=20, pady=6)

    enc_arbol = tk.Frame(frame_arbol)
    enc_arbol.pack(fill="x")

    tk.Label(
        enc_arbol, text="Estructura", font=("Segoe UI", 8, "bold"),
        width=ANCHO_ESTRUCTURA, anchor="w",
    ).pack(side="left")
    tk.Label(
        enc_arbol, text="Estado", font=("Segoe UI", 8, "bold"),
        width=ANCHO_ESTADO, anchor="w",
    ).pack(side="left")

    celda_enc_accion = tk.Frame(enc_arbol, width=ANCHO_ACCION, height=16)
    celda_enc_accion.pack(side="left")
    celda_enc_accion.pack_propagate(False)
    tk.Label(
        celda_enc_accion, text="Accion", font=("Segoe UI", 8, "bold"),
        anchor="w",
    ).pack(side="left")

    tk.Label(
        enc_arbol, text="Detalle", font=("Segoe UI", 8, "bold"), anchor="w",
    ).pack(side="left")

    filas_arbol = tk.Frame(frame_arbol)
    filas_arbol.pack(fill="x")

    def _fila_arbol(parent, prefijo, texto, estado=None, detalle="",
                    negrita=False, boton=None, id_fila=None):
        """
        Una fila del diagrama. Columnas, en orden: estructura, estado,
        accion (el boton, si la fila tiene uno) y detalle -- el boton
        va a la IZQUIERDA del detalle, y por eso su celda tiene ancho
        fijo: asi el detalle arranca siempre en la misma columna,
        tenga o no boton esa fila.
        """

        fila = tk.Frame(parent)
        fila.pack(fill="x", pady=1)

        tk.Label(
            fila,
            text=prefijo + texto,
            font=("Consolas", 9, "bold" if negrita else "normal"),
            width=ANCHO_ESTRUCTURA,
            anchor="w",
        ).pack(side="left")

        tk.Label(
            fila,
            text=SIMBOLO.get(estado, estado) if estado is not None else "",
            width=ANCHO_ESTADO,
            anchor="w",
            fg=COLOR_ESTADO.get(estado, COLOR_NEUTRO),
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left")

        celda_accion = tk.Frame(fila, width=ANCHO_ACCION, height=ALTO_ACCION)
        celda_accion.pack(side="left")
        celda_accion.pack_propagate(False)

        if boton is not None:
            texto_boton, comando = boton
            widget = tk.Button(
                celda_accion, text=texto_boton, font=("Segoe UI", 8, "bold"),
                bg="#fdf0d5", command=comando,
            )
            widget.pack(side="left", padx=(0, 6))
            if id_fila is not None:
                botones_arbol[id_fila] = widget

        tk.Label(
            fila,
            text=detalle,
            anchor="w",
            fg=COLOR_NEUTRO,
            font=("Segoe UI", 8),
            wraplength=380,
            justify="left",
        ).pack(side="left", fill="x", expand=True)

        return fila

    def _boton_de_fila(id_fila):
        """
        Que boton lleva cada fila del diagrama. El id lo pone
        nucleo.revisar_estructura(); la ventana decide que accion le
        cuelga, para que nucleo.py no sepa nada de botones.
        """

        if id_fila == "medidas_sae":
            return ("Actualizar", actualizar_medidas_sae)

        if id_fila == "cmg_csv":
            return ("Traer cmg_15min", traer_cmg_15min)

        if id_fila == "cmg_xlsx":
            return ("Generar", generar_cmg)

        if id_fila == "sscc_dir":
            return ("Traer FD", traer_fd)

        if id_fila == "db_subastas":
            return ("Traer subastas", traer_subastas)

        if id_fila == "consolidado":
            return ("Actualizar todo", lambda: actualizar_consolidado(None))

        if id_fila == "pagos":
            return ("Calcular todo", lambda: actualizar_pagos(None))

        if id_fila.startswith("consolidado:"):
            seccion = id_fila.split(":", 1)[1]
            return ("Actualizar", lambda s=seccion: actualizar_consolidado({s}))

        if id_fila.startswith("pagos:"):
            seccion = id_fila.split(":", 1)[1]
            return ("Calcular", lambda s=seccion: actualizar_pagos({s}))

        return None

    def pintar_arbol(filas):

        for hijo in filas_arbol.winfo_children():
            hijo.destroy()

        botones_arbol.clear()

        if not filas:
            tk.Label(
                filas_arbol,
                text="Selecciona una carpeta base.",
                fg=COLOR_NEUTRO,
                anchor="w",
            ).pack(fill="x")
            return

        prefijos = _prefijos_arbol([fila["nivel"] for fila in filas])

        for fila, prefijo in zip(filas, prefijos):
            _fila_arbol(
                filas_arbol,
                prefijo,
                fila["etiqueta"],
                estado=fila["estado"],
                detalle=fila["detalle"],
                negrita=(fila["nivel"] == 0),
                boton=_boton_de_fila(fila["id"]),
                id_fila=fila["id"],
            )

        if corriendo["activo"]:
            habilitar_botones(False)

    def habilitar_botones(activos):
        for widget in botones_arbol.values():
            try:
                widget.config(state="normal" if activos else "disabled")
            except tk.TclError:
                pass

    # --------------------------------------------------------
    # PROGRESO Y LOG
    # --------------------------------------------------------

    frame_estado = tk.Frame(contenedor)
    frame_estado.pack(fill="x", padx=20, pady=(10, 2))

    tk.Label(
        frame_estado, textvariable=var_estado, anchor="w", font=("Segoe UI", 9),
    ).pack(side="left")

    tk.Label(
        frame_estado, textvariable=var_tiempo,
        font=("Consolas", 10, "bold"), fg="#2d7a2d",
    ).pack(side="right")

    barra = ttk.Progressbar(contenedor, mode="determinate", maximum=100)
    barra.pack(fill="x", padx=20, pady=(0, 8))

    frame_log = tk.LabelFrame(contenedor, text="Registro", padx=6, pady=6)
    frame_log.pack(fill="both", expand=True, padx=20, pady=6)

    txt_log = tk.Text(frame_log, height=14, font=("Consolas", 9), wrap="word")
    scroll_log = tk.Scrollbar(frame_log, command=txt_log.yview)
    txt_log.configure(yscrollcommand=scroll_log.set)
    scroll_log.pack(side="right", fill="y")
    txt_log.pack(side="left", fill="both", expand=True)

    def log(mensaje):
        txt_log.insert("end", str(mensaje) + "\n")
        txt_log.see("end")
        root.update_idletasks()

    # --------------------------------------------------------
    # TIMER
    # --------------------------------------------------------

    timer = {"corriendo": False, "inicio": 0.0}

    def tick():
        if timer["corriendo"]:
            var_tiempo.set(formato_tiempo(time.time() - timer["inicio"]))
            root.after(500, tick)

    # --------------------------------------------------------
    # REVISION DE LA CARPETA
    # --------------------------------------------------------

    def revisar(*_):

        ruta = var_base.get()

        if not ruta or not Path(ruta).is_dir():
            lbl_base.config(fg=COLOR_FALTA)
            pintar_arbol([])
            btn_abrir_salida.config(state="disabled")
            return

        lbl_base.config(fg=COLOR_BASE_OK)

        try:
            _, filas = nucleo.revisar_estructura(ruta, var_aamm.get().strip())
        except Exception as error:
            pintar_arbol(
                [
                    {
                        "id": "error",
                        "etiqueta": "Error al revisar",
                        "nivel": 0,
                        "estado": "falta",
                        "detalle": str(error),
                    }
                ]
            )
            btn_abrir_salida.config(state="disabled")
            return

        pintar_arbol(filas)

        faltan = [f for f in filas if f["estado"] == "falta"]

        if faltan:
            var_estado.set(f"Faltan {len(faltan)} entradas requeridas.")
        else:
            var_estado.set("Entradas completas.")

        btn_abrir_salida.config(state="normal")

    def seleccionar():
        inicial = var_base.get()
        if not inicial or not Path(inicial).is_dir():
            inicial = ""
        ruta = filedialog.askdirectory(
            title="Selecciona la carpeta base del caso",
            initialdir=inicial,
        )
        if ruta:
            var_base.set(ruta)
            guardar_config({"carpeta_base": ruta})
            log(f"Carpeta base: {ruta}")
            revisar()

    tk.Button(frame_carpeta, text="Examinar", command=seleccionar).pack(pady=(6, 0))

    def aamm_cambiado(*_):
        guardar_config({"aamm": var_aamm.get().strip()})
        revisar()

    entry_aamm.bind("<FocusOut>", aamm_cambiado)
    entry_aamm.bind("<Return>", aamm_cambiado)

    # --------------------------------------------------------
    # LANZADOR GENERICO (hilo aparte; log, barra y timer de la
    # ventana principal)
    # --------------------------------------------------------

    def caso_listo():
        """Valida lo minimo comun a todos los botones del arbol."""

        ruta = var_base.get()

        if not ruta or not Path(ruta).is_dir():
            messagebox.showwarning(
                "Falta la carpeta base",
                "Elegi primero la carpeta base del caso.",
            )
            return None

        return ruta

    def lanzar(funcion, kwargs, etiqueta_salida):
        """
        Corre funcion(**kwargs, registrar=..., progreso=...) en un
        hilo aparte. registrar/progreso escriben en el log y la barra
        de la ventana (via root.after, para no tocar tkinter desde el
        hilo). Mientras corre, todos los botones del arbol quedan
        deshabilitados.
        """

        if corriendo["activo"]:
            return

        corriendo["activo"] = True
        habilitar_botones(False)
        barra["value"] = 0
        txt_log.delete("1.0", "end")

        timer["corriendo"] = True
        timer["inicio"] = time.time()
        tick()

        resultado = {"ok": False}

        def trabajo():
            try:
                var_estado.set("Procesando...")
                funcion(
                    **kwargs,
                    registrar=lambda m: root.after(0, log, m),
                    progreso=lambda v: root.after(
                        0, barra.configure, {"value": v}
                    ),
                )
                resultado["ok"] = True

            except nucleo.ErrorEntrada as error:
                root.after(0, log, f"\nERROR DE ENTRADA:\n{error}")
                resultado["error"] = str(error)

            except Exception as error:
                root.after(
                    0, log,
                    f"\nERROR INESPERADO:\n{error}\n\n{traceback.format_exc()}",
                )
                resultado["error"] = str(error)

            finally:
                root.after(0, terminar, resultado)

        def terminar(res):
            timer["corriendo"] = False
            corriendo["activo"] = False

            # revisar() repinta el arbol entero (y con el, todos los
            # botones): no hace falta rehabilitarlos a mano.
            revisar()

            if res["ok"]:
                var_estado.set("Terminado correctamente.")
                barra["value"] = 100
                messagebox.showinfo(
                    "Listo", f"Se genero/actualizo:\n{etiqueta_salida}"
                )
            else:
                var_estado.set("Termino con errores.")
                messagebox.showerror(
                    "Error", res.get("error", "Revisa el registro.")
                )

        threading.Thread(target=trabajo, daemon=True).start()

    # --------------------------------------------------------
    # ACCIONES DE LAS FILAS
    # --------------------------------------------------------

    def traer_cmg_15min():
        """Baja el CSV 15-minutal del periodo a <CARPETA_BASE>/Cmg/."""

        ruta = caso_listo()
        if ruta is None:
            return

        aamm = var_aamm.get().strip()

        try:
            nucleo.validar_aamm(aamm)
        except nucleo.ErrorEntrada as error:
            messagebox.showwarning("Falta el periodo", str(error))
            return

        lanzar(
            nucleo.traer_csv_cmg,
            dict(carpeta_base=ruta, aamm=aamm),
            nucleo.extrae_cmg.nombre_csv_15min(aamm),
        )

    def traer_fd():
        """
        Baja el FD del periodo (SSCC_Disponibilidad_CSF_* /
        SSCC_Desempeño_*) del arbol de indicadores del DCO a
        <CARPETA_BASE>/FD y FMA/.
        """

        ruta = caso_listo()
        if ruta is None:
            return

        aamm = var_aamm.get().strip()

        try:
            nucleo.validar_aamm(aamm)
        except nucleo.ErrorEntrada as error:
            messagebox.showwarning("Falta el periodo", str(error))
            return

        lanzar(
            nucleo.traer_fd,
            dict(carpeta_base=ruta, aamm=aamm),
            f"{nucleo.CARPETA_FD_FMA}/",
        )

    def traer_subastas():
        """
        Copia los Access de subastas del periodo
        (OfertasSSCCAdj*.accdb) de la unidad de red a
        <CARPETA_BASE>/Subastas/DB subastas/.
        """

        ruta = caso_listo()
        if ruta is None:
            return

        aamm = var_aamm.get().strip()

        try:
            nucleo.validar_aamm(aamm)
        except nucleo.ErrorEntrada as error:
            messagebox.showwarning("Falta el periodo", str(error))
            return

        lanzar(
            nucleo.traer_subastas,
            dict(carpeta_base=ruta, aamm=aamm),
            f"{nucleo.CARPETA_DB_SUBASTAS}/",
        )

    def generar_cmg():
        """Arma cmg.xlsx con el CSV que ya esta en Cmg/."""

        ruta = caso_listo()
        if ruta is None:
            return

        aamm = var_aamm.get().strip()

        try:
            nucleo.validar_aamm(aamm)
        except nucleo.ErrorEntrada as error:
            messagebox.showwarning("Falta el periodo", str(error))
            return

        lanzar(
            nucleo.generar_cmg,
            dict(carpeta_base=ruta, aamm=aamm),
            nucleo.ARCHIVO_CMG,
        )

    def actualizar_medidas_sae():
        """
        Corre los cuatro pasos de Medidas_SAE.xlsx de un viaje
        (homologacion -> descarga -> claves -> API de operacion real).
        Es lo mas lento del programa: baja el mes completo de las dos
        APIs del Coordinador. Si se corta, la corrida siguiente retoma
        donde quedo.
        """

        ruta = caso_listo()
        if ruta is None:
            return

        aamm = var_aamm.get().strip()

        try:
            nucleo.validar_aamm(aamm)
        except nucleo.ErrorEntrada as error:
            messagebox.showwarning("Falta el periodo", str(error))
            return

        if not messagebox.askyesno(
            f"Generar {nucleo.ARCHIVO_MEDIDAS_SAE}",
            f"Se van a bajar las medidas del periodo {aamm} de las dos "
            f"APIs del Coordinador, punto de medida por punto de "
            f"medida.\n\nEs el proceso mas lento del programa (puede "
            f"tardar bastante). Si se corta, la proxima vez retoma "
            f"donde quedo.\n\n¿Seguir?",
        ):
            return

        lanzar(
            nucleo.generar_medidas_sae,
            dict(carpeta_base=ruta, aamm=aamm),
            nucleo.ARCHIVO_MEDIDAS_SAE,
        )

    def actualizar_consolidado(secciones):
        """
        secciones: set con UN id de SECCIONES_CONSOLIDADO (el boton
        de esa hoja) o None para todas (el boton del archivo). Lo que
        no entra se conserva tal cual esta hoy en el archivo.
        """

        ruta = caso_listo()
        if ruta is None:
            return

        if secciones is None:
            secciones = {s[0] for s in nucleo.SECCIONES_CONSOLIDADO}

        lanzar(
            nucleo.generar_consolidado,
            dict(
                carpeta_base=ruta,
                aamm=var_aamm.get().strip(),
                secciones_activas=secciones,
            ),
            nucleo.ARCHIVO_SALIDA,
        )

    def actualizar_pagos(secciones):
        """Idem para Pagos_BESS.xlsx / SECCIONES_PAGOS."""

        ruta = caso_listo()
        if ruta is None:
            return

        if secciones is None:
            secciones = {s[0] for s in nucleo.SECCIONES_PAGOS}

        lanzar(
            nucleo.generar_pagos_bess,
            dict(carpeta_base=ruta, secciones_activas=secciones),
            nucleo.ARCHIVO_SALIDA_PAGOS,
        )

    # --------------------------------------------------------
    # BOTONES FIJOS ABAJO
    # --------------------------------------------------------

    def abrir_salida():
        abrir_en_explorador(var_base.get())

    btn_abrir_salida = tk.Button(
        frame_botones,
        text="Abrir carpeta del caso",
        state="disabled",
        command=abrir_salida,
    )
    btn_abrir_salida.pack(side="left", padx=10, expand=True)

    tk.Button(frame_botones, text="Salir", command=root.destroy).pack(
        side="left", padx=10, expand=True
    )

    # --------------------------------------------------------
    # ARRANQUE
    # --------------------------------------------------------

    pintar_arbol([])
    log(
        "Selecciona la carpeta base del caso e ingresa el periodo "
        "(AAMM). Cada fila del diagrama de abajo trae su propio boton: "
        "'Traer cmg_15min' y 'Generar' en Cmg/, 'Traer FD' en "
        "'FD y FMA/', 'Traer subastas' en "
        "Subastas/DB subastas/, 'Actualizar' en cada hoja de "
        "Consolidado_entradas.xlsx y 'Calcular' en cada hoja de "
        "Pagos_BESS.xlsx."
    )

    if var_aamm.get():
        log(f"Periodo recordado: {var_aamm.get()}")

    if var_base.get():
        log(f"Carpeta recordada: {var_base.get()}")
        revisar()

    root.mainloop()


if __name__ == "__main__":
    main()
