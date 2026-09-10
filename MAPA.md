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
  elegidos; dispara en `nucleo` la escritura de `Consolidado_entradas.xlsx`
  dentro de la carpeta base del caso.
- **Expone:** `main()` — punto de entrada (`python Balance_BESS.py`).
- **Depende de:** `nucleo.py` (mismo directorio, import directo).

---

## `nucleo.py`

- **Qué hace:** todo el cálculo de la etapa Medidores (Medidores, Ofertas
  SSCC), más la carga de CMg, FD y Subastas, sin interfaz. Resuelve las
  rutas de un caso a partir de la carpeta base, valida que existan las
  entradas requeridas (incluido el período AAMM que ingresa el usuario), y
  escribe `Consolidado_entradas.xlsx` con las hojas `Medidores`, `Ofertas
  SSCC`, `CMg`, `FD`, `Subastas`, `Log`.

  **Medidores** (A:U): calculadas J (SoC), K (Copia_Ventana = copia de L),
  L (Ventana), N (Clave_Dia_HoraMes), O (Indicador_SoC), R
  (Oferta_Completa_Dia), S (Indicador_Ventana_Oferta), T
  (Ventana_No_Completa). Deliberadamente vacías (diseño confirmado, no
  pendiente): M, P, Q, U. `V, W, X, Y, AB, AC, AD, AE` del plan **no son
  columnas de `Medidores`**: son tablas auxiliares de otro largo (central ×
  día, central × ventana) que se calculan y se escriben juntas en una sola
  hoja (`HOJA_OFERTAS_SSCC = "Ofertas SSCC"`) — ver plan §20.1 y §22. El
  resumen intermedio equivalente a la hoja "Resumen Ofertas SSCC" del
  `.xlsm` original es puramente auxiliar para construir la tabla W:Y: no se
  persiste.

  **CMg**, **FD**, **Subastas** (plan §23): replican únicamente las macros
  de *carga* (`Cargar_CMg_Desde_Archivo`, `Cargar_SSCC_Desempeno_En_FD`,
  `Cargar_Remuneracion_Subastas_Rapido`), no las que las consumen después
  (`Asignar_CMg_a_Calculos_Turbo`, `Actualizar_Calculos_Columnas`), que
  pertenecen a una etapa posterior sin implementar. `FD` tiene el mismo
  patrón de "dos tablas de distinto largo compartiendo hoja" que Ofertas
  SSCC, pero por **columnas** en vez de por filas: el bloque CSF (A:M) y el
  CPF (Q:AE) van lado a lado, cada uno con su propio número de filas.
  `Subastas!N` queda vacía y documentada como pendiente (depende de
  `'Calculo E Costos'`, una hoja de la etapa siguiente). Ninguna de las
  tres tiene un documento de dominio tan detallado como Medidores: las
  columnas puramente copiadas se nombran con su letra de Excel tal cual, no
  se les inventa un nombre de negocio no documentado.
- **Consume:**
  - `<CARPETA_BASE>/Medidas/Medidas_SAE.xlsx` (hoja `Medidas`)
  - Un archivo `.xlsx` dentro de `<CARPETA_BASE>/Medidas/` cuyo nombre
    contenga "SOC" y el AAMM ingresado por el usuario (no hay un nombre de
    archivo fijo; debe existir exactamente uno)
  - `<CARPETA_BASE>/Auxiliares/Centrales.xlsx` (hojas `Resumen BESS` y
    `Diccionario`; `Diccionario` columnas E/F/G — índices 4/5/6 — se usan
    específicamente para homologar Ofertas SSCC)
  - Un archivo `.xlsx`/`.xlsm`/`.xlsb`/`.xls` dentro de
    `<CARPETA_BASE>/Ofertas/` cuyo nombre contenga "OfertasSSCC" (más
    reciente si hay varios)
  - `<CARPETA_BASE>/Cmg/cmg.xlsx` (nombre literal fijo, sin AAMM)
  - Un archivo Excel dentro de `<CARPETA_BASE>/SSCC_Desempeño/` cuyo nombre
    empiece con "SSCC_Desempeño_" (más reciente si hay varios), hojas `CPF
    Horario` y `CSF Horario`
  - Un archivo Excel dentro de `<CARPETA_BASE>/Subastas/` cuyo nombre
    empiece con "3_REMUNERACIÓN_SUBASTAS_E_ID_" (más reciente si hay
    varios; carpeta propia — la macro original lo buscaba junto al .xlsm,
    ver plan §23.3), hoja `DB`
- **Produce:** `<CARPETA_BASE>/Consolidado_entradas.xlsx`, hojas:
  `Medidores`, `Ofertas SSCC` (las tablas W:Y y AB:AE equivalentes, una
  debajo de la otra — ver `_escribir_tabla_con_titulo()`), `CMg`, `FD` (los
  bloques CSF y CPF lado a lado, columnas A:M y Q:AE), `Subastas`, `Log`.
- **Expone (funciones clave agregadas hasta ahora, además de las básicas
  de E/S y homologación):**
  - Ofertas SSCC: `buscar_archivo_ofertas`, `construir_resumen_ofertas_sscc`,
    `cargar_resumen_en_medidores`, `calcular_r`, `calcular_s`,
    `construir_resumen_ventana_oferta`, `calcular_t`.
  - CMg/FD/Subastas: `buscar_archivo_sscc_desempeno`,
    `buscar_archivo_subastas`, `leer_cmg(ruta_cmg, registrar=print)` →
    `df_cmg`; `construir_fd(ruta_sscc, registrar=print)` →
    `(df_fd_csf, df_fd_cpf)`; `construir_subastas(ruta_subastas, registrar=print)`
    → `df_subastas`.
  - `construir_medidores(df_sae, df_soc, anio, mes, ruta_ofertas, diccionario, registrar=print)`
    → `(df_medidores, avisos, df_wxy, df_resumen_ventana)`.
  - `escribir_salida(df, ruta_salida, avisos, incidencias, df_wxy=None, df_resumen_ventana=None, df_cmg=None, df_fd_csf=None, df_fd_cpf=None, df_subastas=None)`.
  - `ejecutar(carpeta_base, aamm, registrar=print, progreso=None)` —
    orquesta el proceso completo de punta a punta.
- **Parámetros fijos:** `INICIO_VENTANA = 10`, `UMBRAL_SOC = 0.06` (ver plan
  de migración §8).
- **Constantes de columnas:** `LETRA_A_CAMPO` (dict A→U de `Medidores`, su
  orden de inserción ES el orden final de columnas), `COLUMNAS_VACIAS`.
- **Depende de:** `pandas`, `openpyxl` (como engine de
  `pd.ExcelWriter`/`pd.read_excel`), `calendar` (stdlib, días del mes).

---

## Diferencias con el documento de dominio

_(vacío — no se detectaron diferencias entre `nucleo.py`/`Balance_BESS.py`
y `docs/Plan_Traspaso_Python_Balance_BESS.md` al organizar el repositorio)._
