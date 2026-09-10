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
  `Medidas_SAE.xlsx`, el archivo de SoC del período y el archivo
  `*OfertasSSCC*` de `Ofertas/`, homologa nombres de central contra
  `Centrales.xlsx`/`Diccionario`, calcula las columnas A:U de `Medidores`
  (incluye R, S, T) y escribe `Hoja_Medidas.xlsx` con la hoja `Medidores`
  más las tablas auxiliares de Ofertas SSCC (ver plan §16.3, §20) y un
  `Log`. Calculadas: J (SoC), K (Copia_Ventana = copia de L), L (Ventana),
  N (Clave_Dia_HoraMes), O (Indicador_SoC), R (Oferta_Completa_Dia), S
  (Indicador_Ventana_Oferta), T (Ventana_No_Completa). Deliberadamente
  vacías (diseño confirmado, no pendiente): M, P, Q, U. `V, W, X, Y, AB,
  AC, AD, AE` del plan **no son columnas de `Medidores`**: son tablas
  auxiliares de otro largo (central × día, central × ventana) que se
  calculan y se escriben como hojas propias — ver plan §20.1.
- **Consume:**
  - `<CARPETA_BASE>/Medidas/Medidas_SAE.xlsx` (hoja `Medidas`)
  - Un archivo `.xlsx` o `.csv` dentro de `<CARPETA_BASE>/Medidas/` cuyo
    nombre contenga "SOC" y el AAMM ingresado por el usuario (no hay un
    nombre de archivo fijo; debe existir exactamente uno). La estructura
    de bloques horizontales por central es la misma en ambos formatos;
    `leer_soc_crudo()` elige el lector según la extensión.
  - `<CARPETA_BASE>/Auxiliares/Centrales.xlsx` (hojas `Resumen BESS` y
    `Diccionario`; `Diccionario` columnas E/F/G — índices 4/5/6 — se usan
    específicamente para homologar Ofertas SSCC)
  - Un archivo `.xlsx`/`.xlsm`/`.xlsb`/`.xls` dentro de
    `<CARPETA_BASE>/Ofertas/` cuyo nombre contenga "OfertasSSCC" (si hay
    más de uno, a diferencia del SoC, se toma el más reciente por fecha de
    modificación — así lo hace la macro original)
- **Produce:** `<CARPETA_BASE>/Hoja_Medidas.xlsx`, hojas: `Medidores`,
  `Resumen Ofertas SSCC`, `Ofertas SSCC por Dia`, `Resumen Ventana Oferta`,
  `Log`.
- **Expone (además de lo ya listado antes de esta sesión):**
  - `leer_soc_crudo(ruta_soc)` — lee el archivo de SoC sin encabezado,
    con `pd.read_csv` o `pd.read_excel` según la extensión.
  - `buscar_archivo_ofertas(ofertas_dir)` — busca el archivo `*OfertasSSCC*`
    más reciente.
  - `construir_resumen_ofertas_sscc(ruta_ofertas, registrar=print)` —
    replica `Generar_Resumen_Ofertas_SSCC`.
  - `cargar_resumen_en_medidores(df_resumen, claves_medidores, diccionario, registrar=print)`
    → `(df_wxy, (anio, mes), avisos)` — replica `OSSCC_CargarResumenEnMedidores`
    (equivalente a `Medidores!W:Y`).
  - `calcular_r(df_medidores, df_wxy, diccionario, registrar=print)` →
    `(serie_r, avisos)` — replica la fórmula de `Medidores!R` (usa `V`
    internamente, vía `_homologar_fge`/`_mapas_homologacion_fge`).
  - `calcular_s(ventana, r_valor)` — replica la fórmula de `Medidores!S`.
  - `construir_resumen_ventana_oferta(clave, ventana, oferta_r, inicio_ventana=INICIO_VENTANA, registrar=print)`
    — replica `Resumir_Medidores_Central_Ventana_Oferta_Completa`
    (equivalente a `Medidores!AB:AE`).
  - `calcular_t(clave, ventana, resumen_ventana_oferta)` →
    `(serie_t, cantidad_sin_match)` — replica la fórmula de `Medidores!T`.
  - `construir_medidores(df_sae, df_soc, anio, mes, ruta_ofertas, diccionario, registrar=print)`
    → `(df_medidores, avisos, df_resumen_ofertas, df_wxy, df_resumen_ventana)`.
  - `escribir_salida(df, ruta_salida, avisos, incidencias, df_resumen_ofertas=None, df_wxy=None, df_resumen_ventana=None)`.
  - `ejecutar(carpeta_base, aamm, registrar=print, progreso=None)` —
    orquesta el proceso completo de punta a punta.
- **Parámetros fijos:** `INICIO_VENTANA = 10`, `UMBRAL_SOC = 0.06` (ver plan
  de migración §8).
- **Constantes de columnas:** `LETRA_A_CAMPO` (dict A→U, su orden de
  inserción ES el orden final de columnas), `COLUMNAS_VACIAS`.
- **Depende de:** `pandas`, `openpyxl` (como engine de
  `pd.ExcelWriter`/`pd.read_excel`), `calendar` (stdlib, días del mes).

---

## Diferencias con el documento de dominio

_(vacío — no se detectaron diferencias entre `nucleo.py`/`Balance_BESS.py`
y `docs/Plan_Traspaso_Python_Balance_BESS.md` al organizar el repositorio)._
