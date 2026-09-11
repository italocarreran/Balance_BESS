# MAPA.md — qué hace cada script

Primera lectura obligatoria de cualquier sesión (ver `METODOLOGIA.md` §2).
Un bloque corto por script: qué hace · consume · produce · expone · depende
de.

---

## `Balance_BESS.py`

- **Qué hace:** ventana tkinter única. Deja elegir la carpeta base del caso
  e ingresar el **período (AAMM, 4 dígitos, ej. `2607`)** en un campo de
  texto, y debajo dibuja un **diagrama de la estructura del caso** (árbol de
  texto tipo consola, prefijos `├──`/`└──`/`│`, patrón tomado del
  `Revisor_Reliquidacion.py` que el usuario dio como referencia) con el
  estado de cada entrada (OK/FALTA/PENDIENTE). El AAMM no se infiere del
  nombre de ningún archivo: lo escribe el usuario.

  Las dos salidas (`Consolidado_entradas.xlsx`, `Pagos_BESS.xlsx`) son las
  dos últimas filas del mismo diagrama, cada una con su botón **Generar...**
  que abre una ventana aparte:
  - *Generar Consolidado_entradas.xlsx*: una casilla por sección de
    `nucleo.SECCIONES_CONSOLIDADO` ("Medidores + Ofertas SSCC", "CMg", "FD",
    "Subastas"; todas tildadas por defecto). Lo destildado se **conserva**
    tal cual estaba en el archivo existente (no se recalcula ni se borra) —
    ver `nucleo.generar_consolidado`.
  - *Generar Pagos_BESS.xlsx*: sin casillas todavía (una sola hoja de
    salida); explica que usa la hoja `Medidores` ya generada (no la
    recalcula) más `Centrales.xlsx`/`cmg.xlsx` frescos — ver
    `nucleo.generar_pagos_bess`.

  Ambas ventanas corren su función de `nucleo` en un hilo aparte
  (`lanzar_generacion()`, helper compartido) y reportan al log/barra de
  progreso/timer de la ventana **principal**, no a widgets propios: no hay
  un botón "Ejecutar" único, cada salida se genera por separado.
- **Consume:** `nucleo` (`revisar_estructura`, `resolver_rutas`,
  `generar_consolidado`, `generar_pagos_bess`, `SECCIONES_CONSOLIDADO`,
  `ErrorEntrada`); `config.json` (última carpeta base y último AAMM
  recordados, por PC/usuario).
- **Produce:** `config.json` actualizado con la carpeta base y el AAMM
  elegidos; dispara en `nucleo` la escritura de `Consolidado_entradas.xlsx`
  y/o `Pagos_BESS.xlsx` dentro de la carpeta base del caso (por separado,
  según que ventana "Generar" se haya usado).
- **Expone:** `main()` — punto de entrada (`python Balance_BESS.py`);
  helpers de presentación del árbol (`_profundidad_fila`,
  `_es_ultimo_en_su_nivel`, `_prefijos_arbol`) que traducen la lista plana
  de `revisar_estructura()` a prefijos tipo consola — deliberadamente NO
  viven en `nucleo.py`, que no conoce conceptos de interfaz.
- **Depende de:** `nucleo.py` (mismo directorio, import directo).

---

## `nucleo.py`

- **Qué hace:** todo el cálculo de la etapa Medidores (Medidores, Ofertas
  SSCC), la carga de CMg, FD y Subastas, y una primera etapa (base) de
  Calculo E Costos, sin interfaz. Resuelve las rutas de un caso a partir de
  la carpeta base, valida que existan las entradas requeridas (incluido el
  período AAMM que ingresa el usuario), y escribe `Consolidado_entradas.xlsx`
  (hojas `Medidores`, `Ofertas SSCC`, `CMg`, `FD`, `Subastas`, `Log`) y
  `Pagos_BESS.xlsx` (hoja `Calculo E Costos`, nombre y alcance provisorios).

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
  `Subastas!N` ("Energía SSCC") queda vacía y documentada como pendiente
  (depende de `'Calculo E Costos'`, cuyo resto de columnas todavía no se
  implementa). Los nombres de columna de `FD` y `Subastas` (`NOMBRES_FD_CSF`,
  `NOMBRES_FD_CPF`, `NOMBRES_SUBASTAS`) fueron confirmados por el usuario
  contra un caso real (plan §24), no inventados — antes de eso se usaba la
  letra de Excel tal cual por no tener esa información.

  **Calculo E Costos** (plan §25, etapa base — a pedido del usuario, "por
  etapas: primero H + CMg + traspaso de Medidores"): replica parcialmente
  `Traspasar_Medidores_A_Calculos_Rapido` (módulo `B_medidores_a_calculos`)
  y `Asignar_CMg_a_Calculos_Turbo` (módulo `A_Carga_Cmg_a_Destino`). Cubre
  A:G (con D↔E invertidas, igual que la macro), H/`Barra` (antes fórmula
  `=VLOOKUP(G,Resumen!B:G,6,FALSE)`; acá homologada por **nombre** de
  columna contra `Resumen BESS!Nombre activo`/`Barra inyección`, no por
  posición, porque `Centrales.xlsx` no reproduce el layout `Resumen!B:G`
  del libro original), I/J (`Energia_Positiva`/`Energia_Negativa`, la
  energía de `Medidores!Gen_Unidad` separada por signo, solo si
  `Ventana_No_Completa = 1`; si no, la fila es de `Calculo RE545`, fuera de
  alcance), K/`SoC` (copia de `Medidores!SoC`), P/`Copia_Ventana` (copia de
  `Medidores!Copia_Ventana`) y Q/`CMg` (homologado por `Barra` + `Cuarto de
  Hora` normalizado, vía `NormalizaCuarto`). Va a un archivo **separado**
  (`Pagos_BESS.xlsx`, nombre provisorio) a pedido explícito del usuario. Los
  nombres de columna son placeholders derivados de los comentarios de la
  macro — todavía no confirmados contra un archivo real (pendiente: el
  usuario adjuntó dos veces un archivo de encabezados que no traía la hoja
  `Ecostos`).

  **Calculo E Costos, etapa 2** (plan §25.6/25.7): agrega `L, N, O, R, S, T,
  U, W, X, Y, AB, AC, AD`. `L` (¿participó en una subasta?) homologa contra
  `Subastas!Sub_Baj` (confirmado por el usuario) + `Configuración`+`Mes`+
  `Dia`+`Hora_dia` (la central-clave es **inferida**, no confirmada letra
  por letra — ver plan §25.6, puede estar mal si `L` da sospechosamente
  bajo). `N/O/R/Y/AB/AC/AD` se calculan por grupo (central=`clave` +
  ventana=`Copia_Ventana`); `S/T/U` no agrupan; `W/X` son **globales** (no
  por grupo). Bloqueadas: `M`, `AE`, `AF`, `AG:AX`, `AZ` — dependen de una
  hoja `Resumen` del libro original (tabla central→factor + un umbral
  único en `H8`) **distinta** de `Centrales.xlsx!Resumen BESS`, que
  todavía no está mapeada en la migración; tampoco el umbral de subida/
  bajada que usan `AU/AV/AW/AZ`. Toda la hoja `Calculo RE545` también
  queda fuera.
- **Consume:**
  - `<CARPETA_BASE>/Medidas/Medidas_SAE.xlsx` (hoja `Medidas`)
  - Un archivo `.xlsx` dentro de `<CARPETA_BASE>/Medidas/` cuyo nombre
    contenga "SOC" y el AAMM ingresado por el usuario (no hay un nombre de
    archivo fijo; debe existir exactamente uno)
  - `<CARPETA_BASE>/Auxiliares/Centrales.xlsx` (hojas `Resumen BESS` y
    `Diccionario`; `Diccionario` columnas E/F/G — índices 4/5/6 — se usan
    específicamente para homologar Ofertas SSCC; `Resumen BESS` columnas
    `Nombre activo`/`Barra inyección` se usan para `Calculo E Costos!Barra`)
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
- **Produce:**
  - `<CARPETA_BASE>/Consolidado_entradas.xlsx`, hojas: `Medidores`, `Ofertas
    SSCC` (las tablas W:Y y AB:AE equivalentes, una al lado de la otra — ver
    `_escribir_tabla_con_titulo()`), `CMg`, `FD` (los bloques CSF y CPF lado
    a lado, columnas A:M y Q:AE, con sus nombres reales), `Subastas` (con
    sus nombres reales), `Log`.
  - `<CARPETA_BASE>/Pagos_BESS.xlsx` (nombre provisorio), hoja `Calculo E
    Costos` (etapa base, ver más arriba).
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
  - Calculo E Costos (etapa base): `_normaliza_cuarto(valor)` →
    texto (replica `NormalizaCuarto`); `construir_dic_cmg(df_cmg)` →
    `dict` clave `"BARRA|CUARTO"` → valor Q; `construir_mapa_barra(resumen_bess)`
    → `dict` nombre de central normalizado → barra de inyección;
    `construir_calculo_e_costos(df_medidores, mapa_barra, dic_cmg, registrar=print)`
    → `df_ecostos`; `escribir_pagos_bess(ruta_salida, df_ecostos, registrar=print)`.
  - Calculo E Costos (etapa 2): `calcular_l(df_ecostos, df_subastas)`,
    `calcular_n_o(df_ecostos)`, `calcular_r_ecostos(df_ecostos)` (sufijo
    `_ecostos` a propósito: Medidores ya tiene su propia `calcular_r()`,
    lógica no relacionada — no fusionarlas), `calcular_s_t_u(df_ecostos)`,
    `calcular_w_x(df_ecostos)`, `calcular_y_ab_ac_ad(df_ecostos)`, todas
    combinadas por `completar_calculo_e_costos_grupos(df_ecostos, df_subastas, registrar=print)`
    → `df_ecostos` con L/N/O/R/S/T/U/W/X/Y/AB/AC/AD agregadas.
  - `construir_medidores(df_sae, df_soc, anio, mes, ruta_ofertas, diccionario, registrar=print)`
    → `(df_medidores, avisos, df_wxy, df_resumen_ventana)`.
  - `escribir_salida(df, ruta_salida, avisos, incidencias, df_wxy=None, df_resumen_ventana=None, df_cmg=None, df_fd_csf=None, df_fd_cpf=None, df_subastas=None, ruta_existente=None, hojas_regenerar=None, registrar=print)`
    — `hojas_regenerar=None` (por defecto) regenera las 5 hojas de datos;
    si es un `set` con algunos nombres de `_HOJAS_CONSOLIDADO`, las que NO
    estén en el set se copian tal cual desde `ruta_existente`
    (`_copiar_hoja_existente()`, copia cruda vía `openpyxl`, sin fórmulas ni
    formato) en vez de recalcularse.
  - `SECCIONES_CONSOLIDADO` — tupla de `(id, etiqueta, descripción, hojas)`
    por cada casilla de la ventana "Generar" de `Consolidado_entradas.xlsx`
    (`"medidores"` agrupa Medidas_SAE + SoC + Centrales + OfertasSSCC,
    porque `construir_medidores()` los necesita siempre juntos; `"cmg"`,
    `"fd"`, `"subastas"` son independientes).
  - `generar_consolidado(carpeta_base, aamm, secciones_activas, registrar=print, progreso=None)`
    — genera/actualiza `Consolidado_entradas.xlsx` recalculando solo las
    secciones tildadas; valida los archivos de entrada únicamente para las
    secciones tildadas (si `"medidores"` no está tildada, no exige
    Medidas_SAE/SoC/Centrales/Ofertas). Reemplaza a la vieja `ejecutar()`.
  - `generar_pagos_bess(carpeta_base, registrar=print, progreso=None)` —
    genera/actualiza `Pagos_BESS.xlsx`; lee `Medidores` Y `Subastas` desde
    `Consolidado_entradas.xlsx` ya generado (no los recalcula), y
    Centrales.xlsx/cmg.xlsx frescos. Sin `aamm` como parámetro: nada de la
    etapa base ni de la etapa 2 de Calculo E Costos lo necesita (todo sale
    de `Medidores`/`Subastas`, que ya traen Mes/Dia/Hora).
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
