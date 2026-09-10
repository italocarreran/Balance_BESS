# MAPA.md — qué hace cada script

Primera lectura obligatoria de cualquier sesión (ver `METODOLOGIA.md` §2).
Un bloque corto por script: qué hace · consume · produce · expone · depende
de.

---

## `Balance_BESS.py`

- **Qué hace:** ventana tkinter única de la etapa Medidores. Deja elegir la
  carpeta base del caso e ingresar el **período (AAMM, 4 dígitos, ej.
  `2607`)** en un campo de texto, valida automáticamente la estructura
  pintando un checklist (OK/FALTA/PENDIENTE), ejecuta el cálculo en un hilo
  aparte con log y barra de progreso, y ofrece abrir la carpeta de salida
  al terminar. El AAMM ya no se infiere del nombre de ningún archivo: lo
  escribe el usuario y ese valor es la fuente de verdad del período.
- **Consume:** `nucleo` (`revisar_estructura(carpeta, aamm)`,
  `ejecutar(carpeta, aamm, ...)`, `resolver_rutas`, `ErrorEntrada`);
  `config.json` (última carpeta base y último AAMM recordados, por
  PC/usuario).
- **Produce:** `config.json` actualizado con la carpeta base y el AAMM
  elegidos; dispara en `nucleo` la escritura de `Hoja_Medidas.xlsx` dentro
  de la carpeta base del caso.
- **Expone:** `main()` — punto de entrada (`python Balance_BESS.py`).
- **Depende de:** `nucleo.py` (mismo directorio, import directo).

---

## `nucleo.py`

- **Qué hace:** todo el cálculo de la etapa Medidores, sin interfaz. Resuelve
  las rutas de un caso a partir de la carpeta base, valida que existan las
  entradas requeridas (incluido el período AAMM que ingresa el usuario), lee
  `Medidas_SAE.xlsx` y el archivo de SoC del período, homologa nombres de
  central contra `Centrales.xlsx`/`Diccionario`, calcula las 31 columnas
  A:AE de `Medidores` según la especificación cerrada del plan (§16.3) y
  escribe `Hoja_Medidas.xlsx` con una hoja `Medidores` y una hoja `Log`.
  Calculadas: J (SoC), K (Copia_Ventana = copia de L), L (Ventana), N
  (Clave_Dia_HoraMes), O (Indicador_SoC). Deliberadamente vacías (no
  pendientes, es diseño confirmado): M, P, Q, U, Z, AA. Pendientes porque
  dependen de macros de Ofertas SSCC cuyo código fuente todavía no se
  entregó: R, S, T, V, W, X, Y, AB, AC, AD, AE (ver plan §19.2).
- **Consume:**
  - `<CARPETA_BASE>/Medidas/Medidas_SAE.xlsx` (hoja `Medidas`)
  - Un archivo `.xlsx` dentro de `<CARPETA_BASE>/Medidas/` cuyo nombre
    contenga "SOC" y el AAMM ingresado por el usuario (no hay un nombre de
    archivo fijo; debe existir exactamente uno)
  - `<CARPETA_BASE>/Auxiliares/Centrales.xlsx` (hojas `Resumen BESS` y
    `Diccionario`)
  - `<CARPETA_BASE>/Ofertas/` — carpeta ya definida en el plan (§16.2);
    todavía no se lee ningún archivo de ahí (falta patrón de nombre y la
    lógica de las macros).
- **Produce:** `<CARPETA_BASE>/Hoja_Medidas.xlsx` (hojas `Medidores` y
  `Log`).
- **Expone:**
  - `ErrorEntrada` — excepción para problemas de datos de entrada.
  - `normalizar(texto)` — minúsculas, sin tildes, espacios colapsados.
  - `validar_aamm(aamm)` — valida que sea texto de 4 dígitos; levanta
    `ErrorEntrada` si no.
  - `resolver_rutas(carpeta_base)` — deriva todas las rutas del caso
    (incluye `ofertas_dir`).
  - `revisar_estructura(carpeta_base, aamm=None)` → `(rutas, filas)` —
    valida entradas para pintar el checklist.
  - `buscar_soc(medidas_dir, aamm)` — busca el archivo cuyo nombre
    contenga "SOC" y el AAMM dado; `periodo_desde_aamm(aamm)`.
  - `leer_medidas_sae(ruta)`, `leer_centrales(ruta)`,
    `construir_homologacion(diccionario)`.
  - `detectar_fila_nombres`, `detectar_bloques`, `extraer_soc` — extracción
    de SoC por detección dinámica de bloques (ver `METODOLOGIA.md` §5).
  - `calcular_ventana`, `calcular_indicador_soc`, `calcular_clave_auxiliar`
    — columnas calculadas L, O, N respectivamente (K se deriva de L
    directamente en `construir_medidores`, sin función propia).
  - `construir_medidores(df_sae, df_soc, anio, mes, registrar=print)` →
    `(df_medidores, avisos)`.
  - `escribir_salida(df, ruta_salida, avisos, incidencias)`.
  - `ejecutar(carpeta_base, aamm, registrar=print, progreso=None)` —
    orquesta el proceso completo de punta a punta.
- **Parámetros fijos:** `INICIO_VENTANA = 10`, `UMBRAL_SOC = 0.06` (ver plan
  de migración §8).
- **Constantes de columnas:** `LETRA_A_CAMPO` (dict A→AE, su orden de
  inserción ES el orden final de columnas — nunca `sorted()` sobre sus
  claves), `COLUMNAS_VACIAS`, `COLUMNAS_PENDIENTES_OFERTAS`.
- **Depende de:** `pandas`, `openpyxl` (como engine de
  `pd.ExcelWriter`/`pd.read_excel`).

---

## Diferencias con el documento de dominio

_(vacío — no se detectaron diferencias entre `nucleo.py`/`Balance_BESS.py`
y `docs/Plan_Traspaso_Python_Balance_BESS.md` al organizar el repositorio)._
