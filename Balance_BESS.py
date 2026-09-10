# -*- coding: utf-8 -*-
"""
Balance BESS / SSCC — etapa Medidores.

Ventana unica: se elige la carpeta base del caso y el programa
resuelve solas todas las entradas por ruta relativa.

    <CARPETA_BASE>/
        Medidas/
            Medidas_SAE.xlsx
            <algo>SOC<algo>AAMM<algo>.xlsx
        Auxiliares/
            Centrales.xlsx
        Ofertas/
            <algo>OfertasSSCC<algo>.xlsx (o .xlsm/.xlsb/.xls)
        Hoja_Medidas.xlsx      <- salida

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
# VENTANA
# ============================================================

def main():

    cfg = leer_config()

    root = tk.Tk()
    root.title("Balance BESS / SSCC — etapa Medidores")
    root.geometry("980x720")

    var_base = tk.StringVar(
        value=cfg.get("carpeta_base", "")
    )
    var_aamm = tk.StringVar(
        value=cfg.get("aamm", "")
    )
    var_estado = tk.StringVar(value="Listo")
    var_tiempo = tk.StringVar(value="00:00:00")

    # --------------------------------------------------------
    # BOTONES FIJOS ABAJO (primero, para que no los tape nada)
    # --------------------------------------------------------

    frame_botones = tk.Frame(root)
    frame_botones.pack(side="bottom", fill="x", pady=10)

    # --------------------------------------------------------
    # CANVAS CON SCROLL
    # --------------------------------------------------------

    canvas = tk.Canvas(root, borderwidth=0, highlightthickness=0)
    scroll = tk.Scrollbar(
        root, orient="vertical", command=canvas.yview
    )
    canvas.configure(yscrollcommand=scroll.set)
    scroll.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    contenedor = tk.Frame(canvas)
    ventana_canvas = canvas.create_window(
        (0, 0), window=contenedor, anchor="nw"
    )

    def ajustar(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(
            ventana_canvas, width=canvas.winfo_width()
        )

    contenedor.bind("<Configure>", ajustar)
    canvas.bind("<Configure>", ajustar)
    canvas.bind_all(
        "<MouseWheel>",
        lambda e: canvas.yview_scroll(
            int(-e.delta / 120), "units"
        ),
    )

    # --------------------------------------------------------
    # SELECTOR DE CARPETA BASE
    # --------------------------------------------------------

    frame_carpeta = tk.LabelFrame(
        contenedor,
        text="Carpeta base del caso",
        padx=10,
        pady=8,
    )
    frame_carpeta.pack(fill="x", padx=20, pady=(14, 6))

    lbl_base = tk.Label(
        frame_carpeta,
        textvariable=var_base,
        wraplength=820,
        justify="center",
        cursor="hand2",
        font=("Segoe UI", 9),
    )
    lbl_base.pack()

    lbl_base.bind(
        "<Button-1>",
        lambda e: (
            abrir_en_explorador(var_base.get())
            if var_base.get()
            else None
        ),
    )

    # --------------------------------------------------------
    # PERIODO DEL CASO (AAMM)
    # --------------------------------------------------------

    frame_periodo = tk.LabelFrame(
        contenedor,
        text="Periodo del caso (AAMM)",
        padx=10,
        pady=8,
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
            "de 2026. Se usa para ubicar el SoC del periodo."
        ),
        fg=COLOR_NEUTRO,
        font=("Segoe UI", 8),
        wraplength=700,
        justify="left",
    ).pack(side="left", padx=(10, 0))

    # --------------------------------------------------------
    # PANEL DE VALIDACION
    # --------------------------------------------------------

    frame_check = tk.LabelFrame(
        contenedor,
        text="Entradas detectadas",
        padx=10,
        pady=8,
    )
    frame_check.pack(fill="x", padx=20, pady=6)

    filas_check = tk.Frame(frame_check)
    filas_check.pack(fill="x")

    def pintar_checklist(filas):

        for hijo in filas_check.winfo_children():
            hijo.destroy()

        if not filas:
            tk.Label(
                filas_check,
                text="Selecciona una carpeta base.",
                fg=COLOR_NEUTRO,
                anchor="w",
            ).pack(fill="x")
            return

        for etiqueta, estado, detalle in filas:

            fila = tk.Frame(filas_check)
            fila.pack(fill="x", pady=1)

            tk.Label(
                fila,
                text=etiqueta,
                width=28,
                anchor="w",
                font=("Segoe UI", 9),
            ).pack(side="left")

            tk.Label(
                fila,
                text=SIMBOLO.get(estado, estado),
                width=11,
                anchor="w",
                fg=COLOR_ESTADO.get(estado, COLOR_NEUTRO),
                font=("Segoe UI", 9, "bold"),
            ).pack(side="left")

            tk.Label(
                fila,
                text=detalle,
                anchor="w",
                fg=COLOR_NEUTRO,
                font=("Segoe UI", 8),
                wraplength=520,
                justify="left",
            ).pack(side="left", fill="x", expand=True)

    # --------------------------------------------------------
    # PROGRESO Y LOG
    # --------------------------------------------------------

    frame_estado = tk.Frame(contenedor)
    frame_estado.pack(fill="x", padx=20, pady=(10, 2))

    tk.Label(
        frame_estado,
        textvariable=var_estado,
        anchor="w",
        font=("Segoe UI", 9),
    ).pack(side="left")

    tk.Label(
        frame_estado,
        textvariable=var_tiempo,
        font=("Consolas", 10, "bold"),
        fg="#2d7a2d",
    ).pack(side="right")

    barra = ttk.Progressbar(
        contenedor, mode="determinate", maximum=100
    )
    barra.pack(fill="x", padx=20, pady=(0, 8))

    frame_log = tk.LabelFrame(
        contenedor, text="Registro", padx=6, pady=6
    )
    frame_log.pack(fill="both", expand=True, padx=20, pady=6)

    txt_log = tk.Text(
        frame_log, height=16, font=("Consolas", 9), wrap="word"
    )
    scroll_log = tk.Scrollbar(
        frame_log, command=txt_log.yview
    )
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
            var_tiempo.set(
                formato_tiempo(time.time() - timer["inicio"])
            )
            root.after(500, tick)

    # --------------------------------------------------------
    # REVISION DE LA CARPETA
    # --------------------------------------------------------

    estado_listo = {"ok": False}

    def revisar(*_):

        ruta = var_base.get()

        if not ruta or not Path(ruta).is_dir():
            lbl_base.config(fg=COLOR_FALTA)
            pintar_checklist([])
            estado_listo["ok"] = False
            btn_ejecutar.config(state="disabled")
            return

        lbl_base.config(fg="#1a4fb0")

        try:
            _, filas = nucleo.revisar_estructura(
                ruta, var_aamm.get().strip()
            )
        except Exception as error:
            pintar_checklist(
                [("Error al revisar", "falta", str(error))]
            )
            estado_listo["ok"] = False
            btn_ejecutar.config(state="disabled")
            return

        pintar_checklist(filas)

        # Solo bloquea lo que falta; lo pendiente no impide correr
        faltan = [f for f in filas if f[1] == "falta"]

        estado_listo["ok"] = not faltan

        btn_ejecutar.config(
            state="normal" if not faltan else "disabled"
        )

        if faltan:
            var_estado.set(
                f"Faltan {len(faltan)} entradas requeridas."
            )
        else:
            var_estado.set("Entradas completas. Puedes ejecutar.")

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
    # EJECUCION
    # --------------------------------------------------------

    def ejecutar():

        if not estado_listo["ok"]:
            messagebox.showwarning(
                "Entradas incompletas",
                "Faltan entradas requeridas. Revisa el panel.",
            )
            return

        btn_ejecutar.config(state="disabled", bg="#aaaaaa")
        btn_examinar_salida.config(state="disabled")
        barra["value"] = 0
        txt_log.delete("1.0", "end")

        timer["corriendo"] = True
        timer["inicio"] = time.time()
        tick()

        resultado = {"ok": False, "salida": None}

        def trabajo():
            try:
                var_estado.set("Procesando...")

                nucleo.ejecutar(
                    var_base.get(),
                    var_aamm.get().strip(),
                    registrar=lambda m: root.after(0, log, m),
                    progreso=lambda v: root.after(
                        0, barra.configure, {"value": v}
                    ),
                )

                rutas = nucleo.resolver_rutas(var_base.get())
                resultado["salida"] = rutas["salida"]
                resultado["ok"] = True

            except nucleo.ErrorEntrada as error:
                root.after(0, log, f"\nERROR DE ENTRADA:\n{error}")
                resultado["error"] = str(error)

            except Exception as error:
                root.after(
                    0,
                    log,
                    f"\nERROR INESPERADO:\n{error}\n\n"
                    f"{traceback.format_exc()}",
                )
                resultado["error"] = str(error)

            finally:
                root.after(0, terminar, resultado)

        def terminar(res):
            timer["corriendo"] = False
            btn_ejecutar.config(state="normal", bg="#2d7a2d")

            if res["ok"]:
                var_estado.set("Terminado correctamente.")
                barra["value"] = 100
                btn_examinar_salida.config(state="normal")
                messagebox.showinfo(
                    "Listo",
                    f"Se genero:\n{res['salida']}",
                )
            else:
                var_estado.set("Termino con errores.")
                messagebox.showerror(
                    "Error",
                    res.get("error", "Revisa el registro."),
                )

        threading.Thread(target=trabajo, daemon=True).start()

    def abrir_salida():
        rutas = nucleo.resolver_rutas(var_base.get())
        abrir_en_explorador(rutas["salida"], es_archivo=True)

    btn_ejecutar = tk.Button(
        frame_botones,
        text="Ejecutar",
        bg="#2d7a2d",
        fg="white",
        font=("Segoe UI", 10, "bold"),
        state="disabled",
        command=ejecutar,
    )
    btn_ejecutar.pack(side="left", padx=10, expand=True)

    btn_examinar_salida = tk.Button(
        frame_botones,
        text="Abrir carpeta de salida",
        state="disabled",
        command=abrir_salida,
    )
    btn_examinar_salida.pack(side="left", padx=10, expand=True)

    tk.Button(
        frame_botones, text="Salir", command=root.destroy
    ).pack(side="left", padx=10, expand=True)

    # --------------------------------------------------------
    # ARRANQUE
    # --------------------------------------------------------

    pintar_checklist([])
    log("Selecciona la carpeta base del caso e ingresa el periodo (AAMM).")

    if var_aamm.get():
        log(f"Periodo recordado: {var_aamm.get()}")

    if var_base.get():
        log(f"Carpeta recordada: {var_base.get()}")
        revisar()

    root.mainloop()


if __name__ == "__main__":
    main()
