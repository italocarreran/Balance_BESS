# MAPA.md — qué hace cada script

Primera lectura obligatoria de cualquier sesión (ver `METODOLOGIA.md` §2).
Un bloque corto por script: qué hace · consume · produce · expone · depende
de.

---

## Estructura del repositorio

```
Balance_BESS.py            <- la ventana (lo unico que se ejecuta)
Script/
    __init__.py
    nucleo.py              <- todo el calculo del caso
    Cmg/
        __init__.py
        Extrae_CMG_barras.py   <- arma cmg.xlsx desde el CSV 15-minutal
    Medidas/
        __init__.py
        comun.py               <- ErrorMedidas + helpers de texto
        Homologacion.py        <- el Excel de Auxiliares/ (punto+canal -> clave)
        Descarga_PRMTE.py      <- API de medidas, por punto de medida
        Claves_Balance.py      <- calendario de cuartos + agrupacion por clave
        Generacion_Real.py     <- API de operacion real (hoja "Medidas API")
```

`Script/` es un paquete: la ventana hace `from Script import nucleo` y
`nucleo.py` hace `from .Cmg import Extrae_CMG_barras`. La idea (conversada
con el usuario) es ir sacando de `nucleo.py` un módulo por etapa, como ya
se hizo con `Cmg/`; por ahora el resto sigue en un solo archivo grande.

El nombre del módulo de CMg usa guiones bajos, no espacios, para que sea
importable como cualquier módulo.

---

## `Balance_BESS.py`

- **Qué hace:** ventana tkinter única. Deja elegir la carpeta base del caso
  e ingresar el **período (AAMM, 4 dígitos, ej. `2607`)** en un campo de
  texto, y debajo dibuja un **diagrama de la estructura del caso** (árbol de
  texto tipo consola, prefijos `├──`/`└──`/`│`, patrón tomado del
  `Revisor_Reliquidacion.py` que el usuario dio como referencia) con el
  estado de cada entrada (OK/FALTA/PENDIENTE). El AAMM no se infiere del
  nombre de ningún archivo: lo escribe el usuario, y **no** es una fila del
  diagrama (no es parte de la estructura de carpetas) — las dos filas que
  dependen de él (el SoC dentro de `Medidas/`, el CSV dentro de `Cmg/`) lo
  dicen en su propio detalle cuando falta.

  **No hay ventanas intermedias ni un botón "Ejecutar" único**: cada acción
  es un botón en la fila que le corresponde. Las columnas del diagrama son
  `Estructura | Estado | Acción | Detalle` — el botón va a la **izquierda**
  del detalle, en una celda de ancho fijo (`ANCHO_ACCION`, en píxeles) para
  que el detalle arranque siempre en la misma columna tenga o no botón esa
  fila.

  | Fila | Botón | Qué hace |
  |---|---|---|
  | `Medidas/Medidas_SAE.xlsx` | **Actualizar** | `nucleo.generar_medidas_sae` — corre los cuatro pasos de Medidas de un viaje |
  | `Cmg/cmg<AAMM>_def_15minutal.csv` | **Traer cmg_15min** | `nucleo.traer_csv_cmg` — copia el CSV del período desde la unidad de red a `Cmg/` |
  | `Cmg/cmg.xlsx` | **Generar** | `nucleo.generar_cmg` — arma `cmg.xlsx` con el CSV que quedó al lado |
  | `Consolidado_entradas.xlsx` | **Actualizar todo** | `generar_consolidado` con todas las secciones |
  | cada `hoja '...'` de esa salida | **Actualizar** | `generar_consolidado` con esa sola sección |
  | `Pagos_BESS.xlsx` | **Actualizar todo** | `generar_pagos_bess` con todas |
  | cada `hoja '...'` de esa salida | **Actualizar** | `generar_pagos_bess` con esa sola |

  Las dos salidas se desglosan por hoja igual que `Centrales.xlsx`: lo que
  no se actualiza se **conserva** tal cual estaba en el archivo (no se
  recalcula ni se borra — ver `escribir_salida`/`hojas_regenerar`), y si el
  archivo todavía no existe se crea con el resto de las hojas vacías (queda
  registrado en su hoja `Log` y el diagrama las muestra como PENDIENTE).

  Todos los botones corren su función de `nucleo` en un hilo aparte
  (`lanzar()`, helper compartido) reportando al log/barra/timer de la
  ventana, y mientras algo corre **todos** los botones del árbol quedan
  deshabilitados (`corriendo`/`habilitar_botones`). Como el árbol se
  repinta entero en cada `revisar()`, las referencias a los botones se
  renuevan ahí (`botones_arbol`).
- **Consume:** `Script.nucleo` (`revisar_estructura`, `traer_csv_cmg`,
  `generar_cmg`, `generar_consolidado`, `generar_pagos_bess`,
  `SECCIONES_CONSOLIDADO`, `SECCIONES_PAGOS`, `validar_aamm`,
  `ErrorEntrada`, `extrae_cmg`); `config.json` (última carpeta base y
  último AAMM recordados, por PC/usuario).
- **Produce:** `config.json` actualizado con la carpeta base y el AAMM
  elegidos; dispara en `nucleo` la escritura del CSV de CMg, `cmg.xlsx`,
  `Consolidado_entradas.xlsx` y/o `Pagos_BESS.xlsx` dentro de la carpeta
  base del caso (cada uno por su botón).
- **Expone:** `main()` — punto de entrada (`python Balance_BESS.py`);
  helpers de presentación del árbol (`_es_ultimo_en_su_nivel`,
  `_prefijos_arbol`) que traducen la lista plana de `revisar_estructura()`
  a prefijos tipo consola — deliberadamente NO viven en `nucleo.py`, que no
  conoce conceptos de interfaz. El **nivel** de cada fila sí lo pone
  `nucleo` (es estructura, no dibujo), y el **id** de cada fila es lo que
  la ventana usa para decidir qué botón le cuelga (`_boton_de_fila`): así
  `nucleo.py` no sabe nada de botones.
- **Depende de:** el paquete `Script/` (mismo directorio).

---

## `Script/Medidas/` — cómo se arma `Medidas_SAE.xlsx`

- **Qué hace:** los cuatro pasos que antes eran cuatro scripts sueltos que se
  corrían a mano uno detrás de otro (autor original: Freddy.Arriagada).
  Ahora corren de un viaje desde el botón **Actualizar** de la fila
  `Medidas_SAE.xlsx` (`nucleo.generar_medidas_sae`).

  | Módulo | Script original | Qué hace |
  |---|---|---|
  | `Homologacion.py` | `0_diccionario_prmte_a_claves_balance.py` | lee el Excel de homologación (`Punto de Medida` + `Canal` → `clave` + `Flujo`) |
  | `Descarga_PRMTE.py` | `1_generacion_prmte.py` | baja las medidas de cada punto, por lotes, reanudable |
  | `Claves_Balance.py` | `2_generacion_claves_Balance.py` | calendario de cuartos de hora + agrupación por clave |
  | `Generacion_Real.py` | `3_Generacion_Real.py` | agrega las centrales de la hoja `Medidas API` desde la API de operación real |

- **Qué cambió respecto de los scripts sueltos** (todo a pedido del usuario,
  salvo donde se diga):
  - el Excel de homologación vive en `Auxiliares/`, al lado de
    `Centrales.xlsx`, y se busca por patrón (`*homologacion*`) en vez de
    estar al lado del `.py`. El `homol.parquet` intermedio desapareció: se
    lee una vez y queda en memoria;
  - la lista `FILTROS_TOPOLOGY` que vivía dentro del paso 3 salió del código
    y ahora es la hoja `Medidas API` de `Centrales.xlsx`. **Ya no es un
    reemplazo**: esas centrales se sacaron del archivo de homologación, así
    que no llegan por el otro camino — el paso 3 las **agrega**, con la
    `clave` que diga esa hoja;
  - el `user_key` de la API se recibe por parámetro. Es una credencial: no
    puede vivir en el código ni en el repositorio (ver `Balance_BESS.py`);
  - los lotes descargados y la marca de reanudación van a
    `<CARPETA_BASE>/Medidas/_trabajo/`, que la ventana no muestra;
  - **(no pedido, es un bug)** los lotes y la marca de reanudación llevan el
    período en el nombre. Antes, correr dos meses en la misma carpeta
    mezclaba los lotes (`medidas_batch_*.parquet` los levantaba todos) y
    daba por procesados puntos de otro mes;
  - **(no pedido, es un bug)** el umbral de "punto de medida completo" ya no
    es la constante `2976` (= 31 × 96) sino la cantidad de cuartos de hora
    que el mes descargado realmente trae. `2976` estaba mal para cualquier
    mes de 30 días o menos, y para los meses con cambio de hora;
  - se dejó de generar `log_inconsistencias_medidas.xlsx`, que comparaba el
    criterio de desempate viejo contra el de mayor `idMeasure`: esa
    comparación era una investigación ya cerrada (el propio script la titula
    "CRITERIO DEFINITIVO"). Lo que sí se informa en el log es cuántos grupos
    venían duplicados.
- **Detalle que importa:** el `Cuarto de Hora` es un índice global del mes
  que después cruza contra `CMg`, así que las **dos** fuentes comparten un
  solo calendario, el que arma `Claves_Balance` (numerado por `intervaloUtc`,
  porque la hora local se repite en el cambio de hora y desordenaría la
  numeración). Las filas de la API de operación real se le pegan por hora
  local — esa API no entrega UTC —, así que en un día de cambio de hora hacia
  atrás hay ambigüedad: se toma la primera ocurrencia y se avisa en el log.
- **Depende de:** `pandas`, `requests` y `pyarrow` (los lotes son parquet).
  **No importa `nucleo`** (misma regla que `Cmg/`): recibe rutas y datos y
  levanta `ErrorMedidas`, que `nucleo` traduce a `ErrorEntrada`.

---

## `Script/Cmg/Extrae_CMG_barras.py`

- **Qué hace:** todo lo que sabe del CSV 15-minutal de CMg. Viene del
  script suelto que se corría a mano al lado del CSV (autor original:
  Freddy.Arriagada), con tres cambios: el CSV se baja de la unidad de red a
  la carpeta `Cmg/` del caso en vez de buscarse al lado del `.py`, las
  barras a filtrar se reciben por parámetro en vez de estar escritas en el
  código, y no escribe el Excel (devuelve DataFrames).
- **Consume:**
  `T:\CMgReales 15MIN\<AAAA>\<AAMM>\Mensual\CMg\Cmg para balance\cmg<AAMM>_def_15minutal.csv`
  (`RAIZ_CMG_REALES` + `SUBCARPETAS_CMG_REALES` + `PLANTILLA_CSV_CMG_15MIN`
  — la única ruta del programa que apunta fuera de la carpeta base del
  caso; si `T:` cambia de letra se cambia ahí y nada más), y después el
  mismo CSV ya copiado en `<CARPETA_BASE>/Cmg/`.
- **Produce:** la copia local del CSV; el DataFrame de `cmg.xlsx` (lo
  escribe `nucleo.generar_cmg`).
- **Expone:** `ErrorCmg`; `nombre_csv_15min(aamm)`,
  `ruta_csv_en_red(aamm, raiz=None)`, `ruta_csv_local(carpeta_cmg, aamm)`,
  `traer_csv_15min(carpeta_cmg, aamm, raiz=None, registrar=print)`,
  `construir_cmg_desde_csv(ruta_csv, barras, registrar=print)` →
  `(df_salida, resumen_dias)`, `validar_layout(df, registrar=print)`,
  `resumen_dias_anomalos(resumen_dias)`.
- **Depende de:** solo `pandas` (a propósito: **no importa `nucleo`**, así
  no hay ciclos de import cuando se saquen más etapas a módulos propios).
  Los errores previsibles salen como `ErrorCmg` y `nucleo` los traduce a
  `ErrorEntrada`.
- **Detalle que importa:** `nucleo.leer_cmg()` vuelve a leer `cmg.xlsx`
  **por posición** (D = Barra, F = valor de Q, H = Cuarto de Hora,
  I = CMg promedio), así que un cambio de columnas en el CSV rompería la
  etapa siguiente en silencio: `validar_layout()` avisa en el log si
  `BARRA` deja de caer en D o `Cuarto de Hora` en H. El `Cuarto de Hora`
  global se numera con los bloques que el CSV **realmente** trae (no se
  asumen 96 por día: los días de cambio de hora traen 92/100).

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
  (`Pagos_BESS.xlsx`, nombre provisorio) a pedido explícito del usuario.

  **Calculo E Costos, etapas 2 y 3** (plan §25.6-25.10): agrega `L, M, N,
  O, R, S, T, U, W, X, Y, AB, AC, AD, AE, AF, AG, AH, AI, AJ, AK, AL, AM,
  AN, AO, AP, AQ, AR, AS, AT, AU, AV`. `L` (¿participó en una subasta?)
  homologa contra `Subastas!Sub_Baj` (confirmado por el usuario) +
  `Configuración`+`Mes`+`Dia`+`Hora_dia`; `M` (¿SoC sobre el mínimo?) y
  `AE`/`AF` (energía asignada por bloques) usan la hoja `Resumen BESS` de
  `Centrales.xlsx` — resultó ser la MISMA tabla que la hoja `Resumen` del
  libro original (no hacía falta una hoja nueva, ver plan §25.8). `AG:AL`
  (Prorratas) salen de una tabla dinámica **derivada de `Subastas`, no de
  un archivo externo** (`construir_prorrata_sscc()`); `AM:AR` (FD) salen
  de homologar la central contra `Diccionario!A→B` y buscar en `FD` un
  bloque de 4 "Cuarto de Hora" (`construir_dic_mapeo_diccionario()`,
  `calcular_fd_prorrateado()`) — en ambos grupos, `CTF` (`AI/AL/AO/AR`) es
  **0 hardcodeado** (confirmado por el usuario: no existe, y así lo hace
  también el VBA original). `AS/AT` combinan lo anterior con `AE`/`AF`
  (`_calcular_costo_ponderado()`); `AU/AV` promedian `AB`/`AD` por grupo,
  activados solo si la suma GLOBAL de energía por ventana (todas las
  centrales, no por grupo) supera/baja de ±10. `N/O/R/Y/AB/AC/AD/AE/AF`
  se calculan por grupo (central=`clave` + ventana=`Copia_Ventana`);
  `S/T/U` no agrupan; `W/X` son **globales** (no por grupo). **Nombres de
  columna reales, confirmados contra un archivo real**
  (`NOMBRES_CALCULO_E_COSTOS`, plan §25.9) — ya no son placeholders; `AG:AL`
  y `AM:AR` comparten a propósito los mismos 6 nombres cortos (así es en
  el archivo real, se distinguen por un encabezado de grupo que no se
  replica en este esquema de una sola fila; `Total` también se repite
  entre `U` y `AX`, así que a esas columnas hay que llegar por posición,
  no por nombre).

  **Calculo E Costos, etapa 4** (plan §25.11): cierra la hoja con `AW`
  (`Descuento FD`), `AX` (`Total` = `AU+AV-AW`) y `AZ` (`Monto a
  compensar`, por grupo, nunca negativo). El umbral de subida/bajada que
  las bloqueaba **no era un archivo externo**: se deriva de `Subastas` +
  `Subastas!N` contando filas por central+ciclo+tipo
  (`construir_dic_umbrales_subastas()`), igual que la Prorrata SSCC. Y
  `Subastas!N` (llamada `"Ciclo"` en el archivo real, plan §26.8 — se
  creyó "Energía SSCC" hasta corregirse) **no es una energía**: es el
  `Ciclo de Carga del mes` de `Calculo E Costos` homologado por
  `Hora_mes` + `Configuración` (`calcular_subastas_ciclo()`), que viene
  de `Medidores` — por eso la "dependencia circular" que se había
  anotado no existía. Fuera de alcance: la columna `AY` (que la macro
  original tampoco escribe).

  **`Pagos_BESS.xlsx` tiene casillas por hoja** en su ventana "Generar"
  (`SECCIONES_PAGOS`, mismo patron que `SECCIONES_CONSOLIDADO`): una
  para `Calculo E Costos`, otra para `Calculo RE545`. La hoja que se
  destilda se preserva tal cual estaba en el archivo existente (no se
  recalcula ni se borra), mismo criterio de `escribir_salida()` para
  `Consolidado_entradas.xlsx`. Solo `Calculo E Costos` exige el archivo
  `SSCC_Desempeño_*`; `Calculo RE545` no lo necesita, asi que tildar
  solo esa seccion no lo pide.

  **Calculo RE545** (plan §26): la hoja hermana, **completa** (`A:CE`).
  La alimenta la MISMA macro de traspaso: `A:G`, `K` y `P` van iguales a
  las dos hojas y lo que se reparte es la energía, según
  `Medidores!Ventana_No_Completa` (`= 1` → E Costos; cualquier otra cosa,
  incluido vacío → RE545). Solo RE545 recibe `T` (`Ventana de
  valorizacion`) y `R` (`CMg Promedio`, `CMg!I`) — por eso
  `construir_dic_cmg()` guarda el par `(CMg!F, CMg!I)`. Es casi toda
  fórmulas en la hoja, no valores escritos por macro (al revés que
  E Costos). Tiene **dos tablas de distinto largo**: el bloque principal
  (una fila por cuarto de hora) y el resumen `AW:BG` (una fila por
  central+ventana), que se escriben lado a lado. `AY` (`Oferta
  Completa`) no es fórmula ni macro: es la columna `Completa` del
  resumen central+ventana que ya alimenta `Medidores!T`. **Trampa:**
  `R`, `S`, `T`, `U` y `V` existen en las dos hojas y significan cosas
  distintas en cada una (`U` es `Total` en E Costos y `EiniT` en RE545);
  lo mismo con el `VLOOKUP` sobre `Resumen BESS`, donde el índice 2 es
  `Pmax (MW)` y el 4 es `Capacidad (MWh)` (ver plan §26.7).
- **Consume:**
  - `<CARPETA_BASE>/Medidas/Medidas_SAE.xlsx` (hoja `Medidas`)
  - Un archivo `.xlsx` dentro de `<CARPETA_BASE>/Medidas/` cuyo nombre
    contenga "SOC" y el AAMM ingresado por el usuario (no hay un nombre de
    archivo fijo; debe existir exactamente uno)
  - `<CARPETA_BASE>/Auxiliares/<algo>Homologacion<algo>.xlsx` (hoja `homol`:
    `Punto de Medida` + `Canal` → `clave` + `Flujo`), solo para generar
    `Medidas_SAE.xlsx`
  - `<CARPETA_BASE>/Auxiliares/Centrales.xlsx` (hojas `Resumen BESS` y
    `Diccionario`; `Diccionario` columnas E/F/G — índices 4/5/6 — se usan
    específicamente para homologar Ofertas SSCC; `Resumen BESS` columnas
    `Nombre activo`/`Barra inyección` se usan para `Calculo E Costos!Barra`)
  - Un archivo `.xlsx`/`.xlsm`/`.xlsb`/`.xls` dentro de
    `<CARPETA_BASE>/Ofertas/` cuyo nombre contenga "OfertasSSCC" (más
    reciente si hay varios)
  - `<CARPETA_BASE>/Cmg/cmg.xlsx` (nombre literal fijo, sin AAMM). No se
    descarga: lo genera el propio programa con `generar_cmg()` a partir de
    `<CARPETA_BASE>/Cmg/cmg<AAMM>_def_15minutal.csv`, que a su vez se baja
    de la unidad de red con `traer_csv_cmg()` (ver
    `Script/Cmg/Extrae_CMG_barras.py`)
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
    Costos` hasta `AV` (ver más arriba; `generar_pagos_bess()` ahora
    también requiere el archivo `SSCC_Desempeño_*` para `AM:AR`).
- **Expone (funciones clave agregadas hasta ahora, además de las básicas
  de E/S y homologación):**
  - Estructura del caso: `revisar_estructura(carpeta_base, aamm=None)` →
    `(rutas, filas)`, donde cada fila es un dict `{id, etiqueta, nivel,
    estado, detalle}` (`estado`: `ok`/`falta`/`pendiente`). El **nivel**
    (0 = raíz del caso, 1 = dentro de una carpeta/archivo, 2 = un nivel
    más) lo pone `nucleo` porque es estructura, no dibujo; el **id** es
    estable y es lo que la ventana usa para colgarle el botón que
    corresponda. Helpers: `_fila()`, `hojas_de(ruta)` (nombres de hoja de
    un Excel) y `hojas_con_datos(ruta)` (`{hoja: tiene datos}`, usado para
    el estado hoja por hoja de las dos salidas: una hoja preservada que
    nunca se generó queda con una sola celda vacía y tiene que verse
    PENDIENTE, no generada), `_filas_de_hojas()`.
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
    `construir_dic_resumen_factor(resumen_bess)` → `(dict` nombre de central
    normalizado → `Pmax (MW), umbral_soc_minimo)` (misma hoja `Resumen BESS`
    que `construir_mapa_barra`, ver plan §25.8);
    `construir_calculo_e_costos(df_medidores, mapa_barra, dic_cmg, registrar=print)`
    → `df_ecostos`; `escribir_pagos_bess(ruta_salida, df_ecostos, registrar=print)`.
  - Calculo E Costos (etapa 2): `calcular_l(df_ecostos, df_subastas)`,
    `calcular_m(df_ecostos, umbral_soc_minimo)`, `calcular_n_o(df_ecostos)`,
    `calcular_r_ecostos(df_ecostos)` (sufijo `_ecostos` a propósito:
    Medidores ya tiene su propia `calcular_r()`, lógica no relacionada — no
    fusionarlas), `calcular_s_t_u(df_ecostos)`, `calcular_w_x(df_ecostos)`,
    `calcular_y_ab_ac_ad(df_ecostos)`, `calcular_ae_af(df_ecostos, dic_factor)`
    (usa `_calcular_asignacion_energia()`).
  - Calculo E Costos (etapa 3, plan §25.10): `construir_prorrata_sscc(df_subastas)`
    → tabla dinámica (pivot); `construir_dic_prorrata(tabla_prorrata, registrar=print)`
    → `dict` central+hora_mes → `(CPF, CSF)`; `calcular_prorratas(df_ecostos, dic_prorrata)`
    → `(AG, AH)`; `construir_dic_mapeo_diccionario(diccionario)` → `dict`
    (tercera lectura de `Diccionario`, distinta de `construir_homologacion`
    y `_mapas_homologacion_fge` — no fusionar); `_calcular_bloque(valor)`;
    `construir_dic_fd_bloque(df_fd, columna_id, columna_mas, columna_menos)`;
    `calcular_fd_prorrateado(df_ecostos, dic_mapeo, dic_fd_csf, dic_fd_cpf)`
    → `(AM, AN, AP, AQ)`; `_calcular_costo_ponderado(...)` +
    `calcular_as_at(df_ecostos)` → `(AS, AT)`; `calcular_au_av(df_ecostos)`
    → `(AU, AV)`.
  - Todas combinadas por
    `completar_calculo_e_costos_grupos(df_ecostos, df_subastas, dic_factor, umbral_soc_minimo, diccionario, df_fd_csf, df_fd_cpf, registrar=print)`
    → `df_ecostos` con L/M/N/O/R/S/T/U/W/X/Y/AB/AC/AD/AE/AF/AG/AH/AI/AJ/AK/AL/AM/AN/AO/AP/AQ/AR/AS/AT/AU/AV
    agregadas Y renombrada a nombres reales (`NOMBRES_CALCULO_E_COSTOS`,
    plan §25.9) — mismo patrón que `NOMBRES_FD_CSF`/`NOMBRES_SUBASTAS`;
    `AG:AL` y `AM:AR` comparten a propósito los mismos 6 nombres cortos
    (así es en el archivo real).
  - `construir_medidores(df_sae, df_soc, anio, mes, ruta_ofertas, diccionario, registrar=print)`
    → `(df_medidores, avisos, df_wxy, df_resumen_ventana)`.
  - `escribir_salida(df, ruta_salida, avisos, incidencias, df_wxy=None, df_resumen_ventana=None, df_cmg=None, df_fd_csf=None, df_fd_cpf=None, df_subastas=None, ruta_existente=None, hojas_regenerar=None, registrar=print)`
    — `hojas_regenerar=None` (por defecto) regenera las 5 hojas de datos;
    si es un `set` con algunos nombres de `_HOJAS_CONSOLIDADO`, las que NO
    estén en el set se copian tal cual desde `ruta_existente`
    (`_copiar_hoja_existente()`, copia cruda vía `openpyxl`, sin fórmulas ni
    formato) en vez de recalcularse.
  - `SECCIONES_CONSOLIDADO` — tupla de `(id, etiqueta, descripción, hojas)`
    por cada casilla de la ventana "Generar" de `Consolidado_entradas.xlsx`.
    `"medidores"` y `"ofertas_sscc"` son ids SEPARADOS (una casilla cada
    uno, cada una decide si se reescribe su propia hoja) pero comparten
    una unica LECTURA/calculo (`construir_medidores()` arma las dos hojas
    de una sola pasada porque `Medidores!R:S:T` depende de Ofertas SSCC):
    alcanza con que cualquiera de las dos este tildada para que se lean
    Medidas_SAE + SoC + Centrales + OfertasSSCC. `"cmg"`, `"fd"`,
    `"subastas"` si son independientes de punta a punta.
  - `generar_consolidado(carpeta_base, aamm, secciones_activas, registrar=print, progreso=None)`
    — genera/actualiza `Consolidado_entradas.xlsx` recalculando solo las
    secciones pedidas; valida los archivos de entrada únicamente para esas
    secciones (si no se pide `"medidores"`, no exige
    Medidas_SAE/SoC/Centrales/Ofertas). Si el archivo no existe, se crea.
    Reemplaza a la vieja `ejecutar()`.
  - `generar_medidas_sae(carpeta_base, aamm, user_key=None, registrar=print, progreso=None)`
    — genera/actualiza `<CARPETA_BASE>/Medidas/Medidas_SAE.xlsx` corriendo
    los cuatro pasos seguidos (botón **Actualizar** de esa fila). El paso de
    la API de operación real es opcional: sin la hoja `Medidas API` se
    escribe solo lo que viene de la homologación. Helpers:
    `leer_medidas_api(ruta_centrales)` (lee esa hoja: `topologyName`,
    `clave`, `Factor` opcional) y `_resumir_diagnostico_medidas()` (lo que
    antes iba a `reporte_medidas_consolidadas.xlsx`, ahora al log).
  - `_leer_hoja_con_encabezado(ruta, hoja, columnas_buscadas)` — lector
    genérico de hojas cuyo encabezado no está en la primera fila (las hojas
    reales traen un título arriba). `_leer_resumen_bess()` es ahora un caso
    particular de este.
  - `traer_csv_cmg(carpeta_base, aamm, registrar=print, progreso=None)` —
    copia el CSV 15-minutal del período de la unidad de red a
    `<CARPETA_BASE>/Cmg/` (botón **Traer cmg_15min**). Se copia en vez de
    leerlo directo de la red para que el caso quede autocontenido: una vez
    traído, `cmg.xlsx` se puede regenerar sin la unidad conectada y queda
    registrado con qué archivo se trabajó.
  - `generar_cmg(carpeta_base, aamm, ruta_csv=None, registrar=print, progreso=None)`
    — genera/actualiza `<CARPETA_BASE>/Cmg/cmg.xlsx` con el CSV que ya está
    en esa misma carpeta (botón **Generar**). Las barras salen de
    `barras_desde_resumen_bess(resumen_bess)`, que reusa
    `construir_mapa_barra()` — la MISMA fuente que `Calculo E Costos!Barra`,
    así que las dos puntas no se pueden desincronizar. El resto (ruta de
    red, formato del CSV, numeración del `Cuarto de Hora`, validación del
    layout) vive en `Script/Cmg/Extrae_CMG_barras.py`, ver su bloque más
    arriba.
  - `generar_pagos_bess(carpeta_base, registrar=print, progreso=None)` —
    genera/actualiza `Pagos_BESS.xlsx`; lee `Medidores` Y `Subastas` desde
    `Consolidado_entradas.xlsx` ya generado (no los recalcula), y
    Centrales.xlsx/cmg.xlsx/`SSCC_Desempeño_*` frescos (este último, nuevo,
    para `AM:AR`). Sin `aamm` como parámetro: nada de las etapas 2/3 de
    Calculo E Costos lo necesita (todo sale de `Medidores`/`Subastas`, que
    ya traen Mes/Dia/Hora).
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
