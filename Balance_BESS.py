# -*- coding: utf-8 -*-
"""
Balance BESS / SSCC.

Ventana unica: se elige la carpeta base del caso y el programa
resuelve solas todas las entradas por ruta relativa. Debajo del
selector de carpeta y del periodo (AAMM) se muestra un diagrama de la
estructura esperada, con el estado de cada entrada (OK/FALTA/
PENDIENTE). Consolidado_entradas.xlsx y Pagos_BESS.xlsx tienen cada
uno su boton "Generar", que abre una ventana aparte para elegir que
partes recalcular.

    <CARPETA_BASE>/
        Medidas/
            Medidas_SAE.xlsx
            <algo>SOC<algo>AAMM<algo>.xlsx
        Auxiliares/
            Centrales.xlsx
        Ofertas/
            <algo>OfertasSSCC<algo>.xlsx (o .xlsm/.xlsb/.xls)
        Cmg/
            cmg.xlsx
        SSCC_Desempeño/
            SSCC_Desempeño_<algo>.xlsx (o .xlsm/.xlsb/.xls)
        Subastas/
            3_REMUNERACIÓN_SUBASTAS_E_ID_<algo>.xlsx (idem)
        Consolidado_entradas.xlsx      <- salida (boton "Generar")
        Pagos_BESS.xlsx                <- salida (boton "Generar")

La ubicacion de este .py no influye en nada salvo en donde se
guarda config.json.
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

import nucleo


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
# revisar_estructura() de nucleo.py devuelve una lista plana de
# (etiqueta, estado, detalle). Para dibujarla como arbol (igual que
# el patron de Revisor_Reliquidacion.py que pidio el usuario) hace
# falta saber la PROFUNDIDAD de cada fila; se deduce del texto de la
# etiqueta en vez de pedirle a nucleo.py que conozca conceptos de
# interfaz (arbol/prefijos), que no le corresponden.
# ============================================================

def _profundidad_fila(etiqueta):
    if etiqueta in ("Carpeta base", "Periodo (AAMM)"):
        return 0
    if etiqueta.startswith("  hoja "):
        return 2
    if etiqueta.endswith("/"):
        return 0
    return 1


def _es_ultimo_en_su_nivel(profundidades, i):
    """
    True si, mirando hacia adelante desde i, se sube de nivel antes
    de encontrar otra fila con la MISMA profundidad (o se llega al
    final de la lista): o sea, si i es la ultima de su grupo.
    """
    nivel = profundidades[i]
    for j in range(i + 1, len(profundidades)):
        if profundidades[j] < nivel:
            return True
        if profundidades[j] == nivel:
            return False
    return True


def _prefijos_arbol(profundidades):
    """
    Prefijos tipo consola (├── / └── / │) para una lista plana de
    profundidades (0 = raiz), calculando el relleno de cada ancestro
    segun si ESE ancestro es o no el ultimo de su propio grupo.
    """
    n = len(profundidades)
    ultimos = [_es_ultimo_en_su_nivel(profundidades, i) for i in range(n)]
    prefijos = []

    for i, nivel in enumerate(profundidades):

        if nivel == 0:
            prefijos.append("")
            continue

        relleno = ""
        for ancestro_nivel in range(1, nivel):
            indice_ancestro = next(
                (
                    k for k in range(i - 1, -1, -1)
                    if profundidades[k] == ancestro_nivel
                ),
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
    root.geometry("1100x780")

    var_base = tk.StringVar(value=cfg.get("carpeta_base", ""))
    var_aamm = tk.StringVar(value=cfg.get("aamm", ""))
    var_estado = tk.StringVar(value="Listo")
    var_tiempo = tk.StringVar(value="00:00:00")

    estado_listo = {"consolidado": False}
    ventanas_generar = {}

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
    ventana_canvas = canvas.create_window(
        (0, 0), window=contenedor, anchor="nw"
    )

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
        lambda e: (
            abrir_en_explorador(var_base.get()) if var_base.get() else None
        ),
    )

    # --------------------------------------------------------
    # PERIODO DEL CASO (AAMM)
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
            "4 digitos: año+mes simplificado. Ej. 2607 para julio "
            "de 2026. Solo hace falta para (re)generar Medidores + "
            "Ofertas SSCC."
        ),
        fg=COLOR_NEUTRO,
        font=("Segoe UI", 8),
        wraplength=780,
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
        width=52, anchor="w",
    ).pack(side="left")
    tk.Label(
        enc_arbol, text="Estado", font=("Segoe UI", 8, "bold"),
        width=11, anchor="w",
    ).pack(side="left")
    tk.Label(
        enc_arbol, text="Detalle", font=("Segoe UI", 8, "bold"),
        anchor="w",
    ).pack(side="left")

    filas_arbol = tk.Frame(frame_arbol)
    filas_arbol.pack(fill="x")

    def _fila_arbol(parent, prefijo, texto, estado=None, detalle="",
                     negrita=False, boton=None):
        fila = tk.Frame(parent)
        fila.pack(fill="x", pady=1)

        tk.Label(
            fila,
            text=prefijo + texto,
            font=("Consolas", 9, "bold" if negrita else "normal"),
            width=52,
            anchor="w",
        ).pack(side="left")

        if estado is not None:
            tk.Label(
                fila,
                text=SIMBOLO.get(estado, estado),
                width=11,
                anchor="w",
                fg=COLOR_ESTADO.get(estado, COLOR_NEUTRO),
                font=("Segoe UI", 9, "bold"),
            ).pack(side="left")
        else:
            tk.Label(fila, text="", width=11).pack(side="left")

        tk.Label(
            fila,
            text=detalle,
            anchor="w",
            fg=COLOR_NEUTRO,
            font=("Segoe UI", 8),
            wraplength=420,
            justify="left",
        ).pack(side="left", fill="x", expand=True)

        if boton is not None:
            texto_boton, comando = boton
            tk.Button(
                fila, text=texto_boton, font=("Segoe UI", 8, "bold"),
                bg="#fdf0d5", command=comando,
            ).pack(side="right", padx=6)

        return fila

    def pintar_arbol(filas):

        for hijo in filas_arbol.winfo_children():
            hijo.destroy()

        if not filas:
            tk.Label(
                filas_arbol,
                text="Selecciona una carpeta base.",
                fg=COLOR_NEUTRO,
                anchor="w",
            ).pack(fill="x")
            return

        profundidades = [_profundidad_fila(etiqueta) for etiqueta, _, _ in filas]
        prefijos = _prefijos_arbol(profundidades)

        for (etiqueta, estado, detalle), profundidad, prefijo in zip(
            filas, profundidades, prefijos
        ):
            _fila_arbol(
                filas_arbol,
                prefijo,
                etiqueta,
                estado=estado,
                detalle=detalle,
                negrita=(profundidad == 0),
            )

        # Las dos salidas van al final del mismo diagrama, como filas
        # de primer nivel (mismo criterio visual que Medidas/,
        # Auxiliares/, etc.), cada una con su boton "Generar".
        _fila_arbol(
            filas_arbol, "", nucleo.ARCHIVO_SALIDA,
            detalle="se genera/actualiza con el boton Generar ->",
            negrita=True,
            boton=("Generar...", abrir_ventana_generar_consolidado),
        )
        _fila_arbol(
            filas_arbol, "", nucleo.ARCHIVO_SALIDA_PAGOS,
            detalle="se genera/actualiza con el boton Generar ->",
            negrita=True,
            boton=("Generar...", abrir_ventana_generar_pagos),
        )

    # --------------------------------------------------------
    # PROGRESO Y LOG
    # --------------------------------------------------------

    frame_estado = tk.Frame(contenedor)
    frame_estado.pack(fill="x", padx=20, pady=(10, 2))

    tk.Label(
        frame_estado, textvariable=var_estado, anchor="w",
        font=("Segoe UI", 9),
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
            estado_listo["consolidado"] = False
            btn_abrir_salida.config(state="disabled")
            return

        lbl_base.config(fg=COLOR_BASE_OK)

        try:
            _, filas = nucleo.revisar_estructura(ruta, var_aamm.get().strip())
        except Exception as error:
            pintar_arbol([("Error al revisar", "falta", str(error))])
            estado_listo["consolidado"] = False
            btn_abrir_salida.config(state="disabled")
            return

        pintar_arbol(filas)

        faltan = [f for f in filas if f[1] == "falta"]

        if faltan:
            var_estado.set(f"Faltan {len(faltan)} entradas requeridas.")
        else:
            var_estado.set("Entradas completas.")

        rutas = nucleo.resolver_rutas(ruta)
        estado_listo["consolidado"] = rutas["salida"].is_file()
        btn_abrir_salida.config(
            state="normal" if rutas["salida"].is_file() else "disabled"
        )

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

    tk.Button(
        frame_carpeta, text="Examinar", command=seleccionar
    ).pack(pady=(6, 0))

    def aamm_cambiado(*_):
        guardar_config({"aamm": var_aamm.get().strip()})
        revisar()

    entry_aamm.bind("<FocusOut>", aamm_cambiado)
    entry_aamm.bind("<Return>", aamm_cambiado)

    # --------------------------------------------------------
    # LANZADOR GENERICO DE GENERACION (hilo aparte, log/progreso
    # compartidos con la ventana principal)
    # --------------------------------------------------------

    def lanzar_generacion(funcion, kwargs, ventana, boton_actualizar,
                           etiqueta_salida):
        """
        Corre funcion(**kwargs, registrar=..., progreso=...) en un
        hilo aparte. registrar/progreso escriben en el log y la
        barra de progreso de la ventana PRINCIPAL (via root.after,
        para no tocar tkinter desde el hilo).
        """

        boton_actualizar.config(state="disabled")
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
                    f"\nERROR INESPERADO:\n{error}\n\n"
                    f"{traceback.format_exc()}",
                )
                resultado["error"] = str(error)

            finally:
                root.after(0, terminar, resultado)

        def terminar(res):
            timer["corriendo"] = False
            boton_actualizar.config(state="normal")
            revisar()

            if res["ok"]:
                var_estado.set("Terminado correctamente.")
                barra["value"] = 100
                messagebox.showinfo(
                    "Listo", f"Se genero/actualizo:\n{etiqueta_salida}"
                )
                if ventana is not None and ventana.winfo_exists():
                    ventana.destroy()
            else:
                var_estado.set("Termino con errores.")
                messagebox.showerror(
                    "Error", res.get("error", "Revisa el registro.")
                )

        threading.Thread(target=trabajo, daemon=True).start()

    # --------------------------------------------------------
    # VENTANA "GENERAR" — Consolidado_entradas.xlsx
    # --------------------------------------------------------

    def abrir_ventana_generar_consolidado():

        if "consolidado" in ventanas_generar and ventanas_generar["consolidado"].winfo_exists():
            ventanas_generar["consolidado"].lift()
            return

        top = tk.Toplevel(root)
        ventanas_generar["consolidado"] = top
        top.title(f"Generar {nucleo.ARCHIVO_SALIDA}")
        top.geometry("620x560")

        tk.Label(
            top,
            text=(
                "Elegi que entradas recalcular esta vez. Lo que dejes "
                "destildado se conserva tal cual esta hoy en "
                f"{nucleo.ARCHIVO_SALIDA} (si ya existe)."
            ),
            wraplength=580, justify="left", anchor="w",
            font=("Segoe UI", 9),
        ).pack(fill="x", padx=14, pady=(14, 8))

        variables = {}

        for id_seccion, etiqueta, descripcion, _ in nucleo.SECCIONES_CONSOLIDADO:
            var = tk.BooleanVar(value=True)
            variables[id_seccion] = var

            fila = tk.LabelFrame(top, text=etiqueta, padx=8, pady=4)
            fila.pack(fill="x", padx=14, pady=4)

            tk.Checkbutton(
                fila, text="Recalcular esta vez", variable=var,
                font=("Segoe UI", 9, "bold"),
            ).pack(anchor="w")

            tk.Label(
                fila, text=descripcion, fg=COLOR_NEUTRO,
                font=("Segoe UI", 8), wraplength=560, justify="left",
                anchor="w",
            ).pack(anchor="w")

        pie = tk.Frame(top)
        pie.pack(fill="x", side="bottom", pady=10)

        def actualizar():
            activas = {
                id_seccion for id_seccion, var in variables.items()
                if var.get()
            }
            if not activas:
                messagebox.showwarning(
                    "Nada tildado",
                    "Tilda al menos una entrada para generar/actualizar.",
                )
                return

            top.attributes("-topmost", False)
            lanzar_generacion(
                nucleo.generar_consolidado,
                dict(
                    carpeta_base=var_base.get(),
                    aamm=var_aamm.get().strip(),
                    secciones_activas=activas,
                ),
                top,
                btn_actualizar,
                nucleo.ARCHIVO_SALIDA,
            )

        btn_actualizar = tk.Button(
            pie, text="Actualizar", bg="#2d7a2d", fg="white",
            font=("Segoe UI", 10, "bold"), command=actualizar,
        )
        btn_actualizar.pack(side="left", padx=14)

        tk.Button(
            pie, text="Cancelar", command=top.destroy
        ).pack(side="left", padx=6)

    # --------------------------------------------------------
    # VENTANA "GENERAR" — Pagos_BESS.xlsx
    # --------------------------------------------------------

    def abrir_ventana_generar_pagos():

        if "pagos" in ventanas_generar and ventanas_generar["pagos"].winfo_exists():
            ventanas_generar["pagos"].lift()
            return

        top = tk.Toplevel(root)
        ventanas_generar["pagos"] = top
        top.title(f"Generar {nucleo.ARCHIVO_SALIDA_PAGOS}")
        top.geometry("620x380")

        tk.Label(
            top,
            text=(
                "Elegi que hojas recalcular esta vez. Lo que dejes "
                "destildado se conserva tal cual esta hoy en "
                f"{nucleo.ARCHIVO_SALIDA_PAGOS} (si ya existe).\n\n"
                f"Usa la hoja 'Medidores' (y 'Subastas') ya generadas "
                f"en {nucleo.ARCHIVO_SALIDA} (no las recalcula: "
                f"primero hay que generar esa con su propio boton), "
                f"mas Centrales.xlsx y cmg.xlsx frescos."
            ),
            wraplength=580, justify="left", anchor="w",
            font=("Segoe UI", 9),
        ).pack(fill="x", padx=14, pady=(14, 8))

        variables = {}

        for id_seccion, etiqueta, descripcion, _ in nucleo.SECCIONES_PAGOS:
            var = tk.BooleanVar(value=True)
            variables[id_seccion] = var

            fila = tk.LabelFrame(top, text=etiqueta, padx=8, pady=4)
            fila.pack(fill="x", padx=14, pady=4)

            tk.Checkbutton(
                fila, text="Recalcular esta vez", variable=var,
                font=("Segoe UI", 9, "bold"),
            ).pack(anchor="w")

            tk.Label(
                fila, text=descripcion, fg=COLOR_NEUTRO,
                font=("Segoe UI", 8), wraplength=560, justify="left",
                anchor="w",
            ).pack(anchor="w")

        pie = tk.Frame(top)
        pie.pack(fill="x", side="bottom", pady=10)

        def actualizar():
            activas = {
                id_seccion for id_seccion, var in variables.items()
                if var.get()
            }
            if not activas:
                messagebox.showwarning(
                    "Nada tildado",
                    "Tilda al menos una hoja para generar/actualizar.",
                )
                return

            top.attributes("-topmost", False)
            lanzar_generacion(
                nucleo.generar_pagos_bess,
                dict(
                    carpeta_base=var_base.get(),
                    secciones_activas=activas,
                ),
                top,
                btn_actualizar,
                nucleo.ARCHIVO_SALIDA_PAGOS,
            )

        btn_actualizar = tk.Button(
            pie, text="Actualizar", bg="#2d7a2d", fg="white",
            font=("Segoe UI", 10, "bold"), command=actualizar,
        )
        btn_actualizar.pack(side="left", padx=14)

        tk.Button(
            pie, text="Cancelar", command=top.destroy
        ).pack(side="left", padx=6)

    # --------------------------------------------------------
    # BOTONES FIJOS ABAJO
    # --------------------------------------------------------

    def abrir_salida():
        rutas = nucleo.resolver_rutas(var_base.get())
        abrir_en_explorador(rutas["salida"], es_archivo=True)

    btn_abrir_salida = tk.Button(
        frame_botones,
        text="Abrir carpeta de salida",
        state="disabled",
        command=abrir_salida,
    )
    btn_abrir_salida.pack(side="left", padx=10, expand=True)

    tk.Button(
        frame_botones, text="Salir", command=root.destroy
    ).pack(side="left", padx=10, expand=True)

    # --------------------------------------------------------
    # ARRANQUE
    # --------------------------------------------------------

    pintar_arbol([])
    log(
        "Selecciona la carpeta base del caso e ingresa el periodo "
        "(AAMM). Los botones 'Generar' de Consolidado_entradas.xlsx "
        "y Pagos_BESS.xlsx aparecen al final del diagrama de abajo."
    )

    if var_aamm.get():
        log(f"Periodo recordado: {var_aamm.get()}")

    if var_base.get():
        log(f"Carpeta recordada: {var_base.get()}")
        revisar()

    root.mainloop()


if __name__ == "__main__":
    main()
