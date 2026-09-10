# MAPA.md — qué hace cada script

Primera lectura obligatoria de cualquier sesión (ver `METODOLOGIA.md` §2).
Un bloque corto por script: qué hace · consume · produce · expone · depende
de.

---

## `Balance_BESS.py`

- **Qué hace:** ventana tkinter única de la etapa Medidores. Deja elegir la
  carpeta base del caso, valida automáticamente su estructura pintando un
  checklist (OK/FALTA/PENDIENTE), ejecuta el cálculo en un hilo aparte con
  log y barra de progreso, y ofrece abrir la carpeta de salida al terminar.
- **Consume:** `nucleo` (`revisar_estructura`, `ejecutar`, `resolver_rutas`,
  `ErrorEntrada`); `config.json` (última carpeta base recordada, por
  PC/usuario).
- **Produce:** `config.json` actualizado con la carpeta base elegida;
  dispara en `nucleo` la escritura de `Hoja_Medidas.xlsx` dentro de la
  carpeta base del caso.
- **Expone:** `main()` — punto de entrada (`python Balance_BESS.py`).
- **Depende de:** `nucleo.py` (mismo directorio, import directo).

---

## `nucleo.py`

- **Qué hace:** todo el cálculo de la etapa Medidores, sin interfaz. Resuelve
  las rutas de un caso a partir de la carpeta base, valida que existan las
  entradas requeridas, lee `Medidas_SAE.xlsx` y el `SOC_AAMM.xlsx` del caso,
  homologa nombres de central contra `Centrales.xlsx`/`Diccionario`, calcula
  las columnas J (SoC), L (Ventana), N (Clave_Dia_HoraMes) y O
  (Indicador_SoC), deja pendientes K/M/P/Q/R/S/T, y escribe
  `Hoja_Medidas.xlsx` con una hoja `Medidores` y una hoja `Log`.
- **Consume:**
  - `<CARPETA_BASE>/Medidas/Medidas_SAE.xlsx` (hoja `Medidas`)
  - `<CARPETA_BASE>/Medidas/SOC_AAMM.xlsx` (patrón `SOC_AAMM.xlsx`, ej.
    `SOC_2607.xlsx`; debe existir exactamente uno)
  - `<CARPETA_BASE>/Auxiliares/Centrales.xlsx` (hojas `Resumen BESS` y
    `Diccionario`)
- **Produce:** `<CARPETA_BASE>/Hoja_Medidas.xlsx` (hojas `Medidores` y
  `Log`).
- **Expone:**
  - `ErrorEntrada` — excepción para problemas de datos de entrada.
  - `normalizar(texto)` — minúsculas, sin tildes, espacios colapsados.
  - `resolver_rutas(carpeta_base)` — deriva todas las rutas del caso.
  - `revisar_estructura(carpeta_base)` → `(rutas, filas)` — valida entradas
    para pintar el checklist.
  - `buscar_soc(medidas_dir)`, `periodo_desde_aamm(aamm)`.
  - `leer_medidas_sae(ruta)`, `leer_centrales(ruta)`,
    `construir_homologacion(diccionario)`.
  - `detectar_fila_nombres`, `detectar_bloques`, `extraer_soc` — extracción
    de SoC por detección dinámica de bloques (ver `METODOLOGIA.md` §5).
  - `calcular_ventana`, `calcular_indicador_soc`, `calcular_clave_auxiliar`
    — columnas calculadas L, O, N respectivamente.
  - `construir_medidores(df_sae, df_soc, anio, mes, registrar=print)` →
    `(df_medidores, avisos)`.
  - `escribir_salida(df, ruta_salida, avisos, incidencias)`.
  - `ejecutar(carpeta_base, registrar=print, progreso=None)` — orquesta el
    proceso completo de punta a punta.
- **Parámetros fijos:** `INICIO_VENTANA = 10`, `UMBRAL_SOC = 0.06` (ver plan
  de migración §8).
- **Depende de:** `pandas`, `openpyxl` (como engine de
  `pd.ExcelWriter`/`pd.read_excel`).

---

## Diferencias con el documento de dominio

_(vacío — no se detectaron diferencias entre `nucleo.py`/`Balance_BESS.py`
y `docs/Plan_Traspaso_Python_Balance_BESS.md` al organizar el repositorio)._
