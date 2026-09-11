# BITACORA.md — registro de sesiones

Solo se agrega. Nunca se edita ni borra una entrada vieja. La única
excepción es la sección "Pendientes abiertos", que sí se edita porque es un
estado, no un historial.

---

## Pendientes abiertos

- Crear casos de prueba con datos reales y comparar la salida Python
  contra `11_PAGOS_BESS_2607_Definitivo.xlsm` (Medidores, Ofertas SSCC,
  CMg, FD, Subastas, Calculo E Costos) (plan §13 punto 10, §20.3). Todo se
  validó hasta ahora solo con casos sintéticos.
- Confirmar si la columna `V` (clave de homologación día+central, interna
  a `calcular_r`) necesita persistirse en una hoja propia para poder
  auditarla fila a fila contra la planilla 11, o si alcanza con auditar
  "Ofertas SSCC por Dia" + `Diccionario!E:F:G` a mano.
- ~~Bloqueante: confirmar la estructura real de columnas de `Subastas`~~
  — **resuelto**: el usuario entregó un `Libro1.xlsx` con la macro real y
  encabezados+fórmulas de la hoja `subastas` real. Confirmó que faltaba la
  columna `Concepto` (B) — ver entrada de esta sesión ("Corrección grande:
  `NOMBRES_SUBASTAS` estaba mal desde el principio"). `NOMBRES_SUBASTAS`
  ya está corregido y validado con los valores exactos de ese archivo.
- ~~Confirmar con el usuario qué son [Configuración, Ciclo, Clave, SUBIDA,
  BAJADA]~~ — resuelto en la misma entrega: viven en `Subastas!S:W` (antes
  se documentaba como `R:V`, un error menor de una letra). Coincide
  exactamente con lo que `construir_dic_umbrales_subastas()` ya calculaba
  en Python — sigue sirviendo como dato de validación cruzada si el
  usuario comparte valores reales de esa tabla para comparar.
- ~~Revisar con el usuario los valores reales de `Subastas!Sub_Baj`~~ —
  resuelto: `Sub_Baj` SÍ tiene `BAJADA`/`SUBIDA` reales, el dato nunca
  estuvo mal. Lo que estaba mal era `NOMBRES_SUBASTAS`: el nombre
  `"Sub_Baj"` estaba pegado a la posición equivocada (una posición antes
  de donde realmente vive), así que el código, sin saberlo, terminaba
  filtrando sobre los valores de `Control` (`CSF`/`CTF`/`CPF`, sin
  dirección) en vez de sobre `BAJADA`/`SUBIDA` reales. Eso explica
  exactamente el crash de la sesión anterior (0 filas tras el filtro
  `.isin(["BAJADA","SUBIDA"])` — nunca iba a haber match comparando contra
  texto que dice "CSF"). Falta re-confirmar con un caso real que ahora sí
  aparezcan filas `BAJADA`/`SUBIDA` con la corrección aplicada.
- ~~Validar contra un caso real que `Subastas!Control` tenga exactamente
  los valores `CPF`/`CSF`~~ — aclarado: `Control` (real) tiene el tipo SIN
  dirección (`CSF`/`CTF`/`CPF`, confirmado con el archivo real), y
  `construir_dic_prorrata()` ya buscaba por substring "cpf"/"csf" — sigue
  funcionando igual, sin cambios de código. Lo que SÍ cambió es que
  `construir_dic_reservas_subastas()` (reservas por subasta de `Calculo
  RE545`) usaba por error `Control` cuando necesitaba `Concepto` (las
  etiquetas completas `CPF(-)`/`CSF(+)`/etc que usa la fórmula real de
  `Subastas!$B:$B`) — corregido en esta sesión.
- Confirmar con más de un archivo de SoC real que el patrón "fila de
  nombre limpio + fila de ruta SCADA apiladas" (resuelto esta sesión,
  `detectar_fila_nombres()`) es estable. Con el único archivo real visto
  hasta ahora, los nombres de la fila 2 ya vienen idénticos a
  `Medidas_SAE.xlsx` — sospechar que el `Diccionario` quizás ni haga falta
  para el SoC, pero falta confirmarlo con otro período/archivo.
- Correr una vez el botón "Actualizar" de `Medidas_SAE.xlsx` contra la API
  real: confirmar la forma de la respuesta y que el `intervalo` de las dos
  APIs sea el inicio del cuarto de hora en las dos (de eso depende el cruce
  contra el calendario compartido).
- Confirmar si la columna `Canal` de la hoja `Gen real` tiene que significar
  algo: la API de operación real no expone canales, así que hoy se acepta
  (para que la hoja tenga la misma forma que `homol`) pero se ignora.
- Confirmar si la columna `Flujo` de la hoja `Gen real` hace falta o si todas
  las centrales van con 1 (el script original no aplicaba signo).
- Abrir la ventana en Windows y confirmar el ancho de la columna "Acción"
  (`ANCHO_ACCION`, hoy 150 px) contra los botones más largos ("Traer
  cmg_15min", "Actualizar todo") y el alto de fila (`ALTO_ACCION`, 26 px).
- Correr `traer_csv_cmg()`/`generar_cmg()` una vez contra el CSV real de
  `T:\CMgReales 15MIN`
  para confirmar que las barras de `Resumen BESS!Barra inyección` están
  escritas exactamente igual que la columna `BARRA` del CSV (con el relleno
  de guiones bajos, ej. `TOCOPILLA_____110`) y que el CSV real trae las 7
  columnas que dejan `BARRA` en D y `Cuarto de Hora` en H (ver
  `_validar_layout_cmg`).
- Validar la Prorrata SSCC normalizada (corregida esta sesión) contra más
  filas de la planilla 11 real — solo se confirmaron 2 casos puntuales
  (los que el usuario reportó), aunque el mecanismo (normalizar por fila)
  quedó confirmado con certeza matemática, no es una inferencia.
- ~~Investigar el reporte del usuario "En el R545 tengo diferencias igual
  parten en AK:AN"~~ — resuelto esta sesión: no era un problema de
  columnas AK:AN en sí (esas coinciden perfectamente comparando por
  posición/orden, no por letra de Excel — la salida nunca reprodujo la
  letra real, ver más abajo), sino que `AR:AT` (`CPF(+)`/`CSF(+)`/`CTF(+)`
  del bloque "FMA", el tercero de los tres bloques de reservas) se
  calculaban con el mismo `SUMIFS` que el resto, cuando en el archivo real
  son la **constante 1** en las 26.784 filas, no una fórmula. Ver entrada
  de esta sesión.
- Revisar si `construir_dic_mapeo_diccionario()` (columnas A:B, usada por
  el FD homologado de `Calculo E Costos!AM:AR`) y `_mapas_homologacion_
  fge()` (columnas E:F:G, usada por `Medidores!V`) también deberían usar
  `_bloques_columnas_diccionario()` en vez de índices de columna fijos
  (`0,1` y `4,5,6`) — hoy coinciden con los bloques reales por casualidad
  en el único `Diccionario` visto, pero si el orden de los bloques
  cambiara en otro archivo (o se agregara un bloque nuevo a la izquierda)
  esas posiciones fijas se romperían. No se tocó porque HOY funcionan
  bien y cambiarlas sin otro caso real de por medio sería un cambio
  especulativo.
- Correr un caso real completo (`Consolidado_entradas.xlsx` +
  `Pagos_BESS.xlsx`) con la corrección de `Subastas` aplicada, para
  confirmar que ahora sí aparecen filas `L=1` (participa en subasta) y
  que las Prorratas/reservas de RE545 no quedan todas en 0. Es el
  siguiente paso natural después de esta sesión.
- `Calculo E Costos` y `Calculo RE545` están **completas** (ver plan §25
  y §26). Lo que sigue son las hojas de salida que las consumen:
  `PRORRATA_RETIROS`, `Compensacion total`, `Resumen` y el CSV
  (`Verificacion_CSV`) — ninguna analizada todavía. El documento de
  trazabilidad las marca como "capa de cálculo masiva, necesita rastreo
  aguas arriba dedicado" (§7 de ese documento).
- Confirmar contra un caso real cuál de las columnas de `Subastas` suma
  cada bloque de reservas de `Calculo RE545` (`AC:AH`, `AI:AN`, `AO:AT`).
  Se siguió la fórmula (posición `O`/`P`/`Q`), pero los nombres reales de
  esas columnas (`FD`, `FMA`, sin nombre) están corridos una columna
  respecto de los títulos de grupo de RE545 (Subastas/FD/FMA). Ver plan
  §26.3.
- Decidir si la columna `Energía SSCC` de la hoja `Subastas` de
  `Consolidado_entradas.xlsx` tiene que quedar escrita ahí. El cálculo ya
  no es un pendiente (`calcular_subastas_energia_sscc()`), pero se hace
  del lado de `Pagos_BESS.xlsx`, que se genera después y en otro archivo;
  escribirla en la hoja `Subastas` implicaría que el botón "Generar
  Pagos_BESS" modifique el archivo de la otra ventana. Es una decisión de
  presentación, no de cálculo.
- Revisar `Subastas!M` ("Ciclo"). Aplicando el mismo corrimiento de una
  columna que el usuario confirmó (y que se usó para `L`, `AW` y
  `Subastas!N`), la fórmula original `=K&H&I` sería
  `Configuración & Dia & Hora_dia`, no `Propietario & Hora_dia &
  Hora_mes` como está hoy. Hoy no afecta ningún cálculo (ninguna otra
  columna consume `Ciclo`), por eso no se tocó sin preguntar.
- Confirmar si la carpeta `Subastas/` (creada esta sesión, no existe en
  la planilla original) es el nombre/ubicación que se quiere mantener, o
  si se prefiere buscar el archivo directamente en `<CARPETA_BASE>` como
  hacía la macro original (ver plan §23.3).
- Confirmar el nombre definitivo de `Pagos_BESS.xlsx` (provisorio, elegido
  por el usuario como "pagos_bess o algo así por ahora").
- Probar la ventana nueva (diagrama + botones "Generar") con una carpeta
  base real: solo se probó por ahora con `python -m py_compile` (no hay
  entorno grafico en esta sesión para abrir la ventana) y con pruebas
  sintéticas de la lógica de árbol (`_prefijos_arbol`) y de generación
  parcial (`generar_consolidado`, `generar_pagos_bess`) por separado.
- Evaluar si `guardar_config()` necesita escritura atómica (ver
  `METODOLOGIA.md` §7).

---

## 2026-09-10 — Organización inicial del repositorio

Se recibieron los archivos de la etapa Medidores (`Balance_BESS.py`,
`nucleo.py`), el plan de traspaso a Python y una plantilla de metodología.
El repositorio estaba prácticamente vacío (solo un `README.md` de una
línea). Se ordenó siguiendo la plantilla de metodología recibida:

- `Balance_BESS.py` y `nucleo.py` agregados en la raíz del repositorio, sin
  cambios respecto de lo recibido.
- Plan de traspaso agregado como documento de dominio en
  `docs/Plan_Traspaso_Python_Balance_BESS.md`.
- `METODOLOGIA.md` adaptado con la información real del proyecto
  (secciones 1, 2, 4, 5, 7 y 8 completadas; ya no es la plantilla genérica).
- Creados `MAPA.md` (un bloque por script), `BITACORA.md` (este archivo),
  `REGLAS.md` (checklist de inicio/cierre de sesión), `requirements.txt`
  (`pandas`, `openpyxl`) y `.gitignore` (`config.json`, `__pycache__/`,
  salidas `.xlsx` de casos concretos).
- `README.md` actualizado con instalación, uso y tabla de navegación hacia
  el resto de los documentos.

Estado dejado: solo está implementada la etapa Medidores, y dentro de ella
solo las columnas A:J, L, N y O. Las columnas K, M, P, Q, R, S, T quedan
como `pd.NA` (pendientes, ver arriba). No hay implementación de etapas
posteriores del balance ni de `OfertasSSCC`. No se corrió el proceso contra
un caso real en esta sesión (no había datos de un caso disponibles); queda
pendiente crear casos de prueba para validar contra la planilla 11.

---

## 2026-09-10 (2) — Plan actualizado: AAMM manual y especificación cerrada de columnas

Se recibió una versión actualizada del plan de traspaso
(`docs/Plan_Traspaso_Python_Balance_BESS.md`, secciones 16–18 nuevas) más un
pedido explícito: el archivo de SoC no se llama literalmente
`SOC_AAMM.xlsx`; el período AAMM debe ingresarlo el usuario en un recuadro
de la ventana, no inferirse del nombre de un archivo.

Cambios en `nucleo.py`:

- `buscar_soc(medidas_dir, aamm)` ahora recibe el AAMM como parámetro
  (validado por la nueva `validar_aamm()`) y busca cualquier `.xlsx` dentro
  de `Medidas/` cuyo nombre contenga "SOC" y ese AAMM, en vez de exigir el
  patrón literal `SOC_AAMM.xlsx`. Se probó con un nombre deliberadamente no
  literal (`"resumen soc julio 2607 v2.xlsx"`) y detecta correctamente.
- `revisar_estructura(carpeta_base, aamm)` y `ejecutar(carpeta_base, aamm, ...)`
  ahora reciben el AAMM en vez de extraerlo del archivo de SoC. El checklist
  agrega una fila "Periodo (AAMM)" que bloquea Ejecutar si no son 4 dígitos.
- Se agregó `CARPETA_OFERTAS = "Ofertas"` y `ofertas_dir` en
  `resolver_rutas()`, con una fila no bloqueante en el checklist (ubicación
  ya definida por el plan §16.2; falta patrón de archivo y lectura).
- `LETRA_A_CAMPO` se extendió de A a AE siguiendo la tabla cerrada del plan
  (§16.3). Se corrigió el orden final de columnas para que use
  `list(LETRA_A_CAMPO.values())` en vez de `sorted(LETRA_A_CAMPO)`: con
  claves de dos letras ("AA", "AB"...) el orden alfabético de texto ya no
  coincide con el orden real de columnas de Excel.
- Se implementó `K` = copia de `L` fila a fila (`Copia_Ventana`).
- Se separaron `COLUMNAS_VACIAS` (M, P, Q, U, Z, AA — diseño confirmado,
  no pendiente) de `COLUMNAS_PENDIENTES_OFERTAS` (R, S, T, V, W, X, Y, AB,
  AC, AD, AE — dependen de las macros de Ofertas SSCC).

Cambios en `Balance_BESS.py`:

- Nuevo campo "Periodo del caso (AAMM)" en la ventana (`Entry` limitado a
  4 dígitos), con su propio `StringVar` persistido en `config.json`.
  Cambiar el AAMM (Enter o al perder foco) vuelve a correr la validación
  de estructura.
- `revisar()` y la llamada a `nucleo.ejecutar()` pasan el AAMM ingresado.

**No se implementaron** `R, S, T, V, W, X, Y, AB, AC, AD, AE`: el plan
actualizado (§17, §18) exige que las macros de Ofertas SSCC se repliquen
fielmente, pero esta sesión solo tiene el nombre de esas macros y qué
columna produce cada una — no su código VBA. Implementarlas sin eso sería
adivinar la lógica, lo que el propio plan prohíbe (§18: "no reemplazar
lógica conocida por placeholders" — y en este caso la lógica todavía no es
conocida por el asistente). Quedan documentadas en
`docs/Plan_Traspaso_Python_Balance_BESS.md` §19.2 y en "Pendientes
abiertos" arriba.

Verificación: se armó un caso sintético en el scratchpad (carpeta con
`Medidas_SAE.xlsx`, un archivo de SoC con nombre no literal, y
`Centrales.xlsx`) y se corrió `nucleo.ejecutar()` de punta a punta: genera
31 columnas en el orden correcto, `K` coincide con `L`, y las columnas
vacías/pendientes quedan como `NaN`. No se probó contra un caso real ni
contra la planilla 11 (sigue sin datos reales disponibles en el entorno).

---

## 2026-09-10 (3) — Ofertas SSCC implementado a partir del código VBA fuente

Se recibió `Trazabilidad_11_PAGOS_BESS_2607_Definitivo.md` con el código VBA completo de
`Generar_Resumen_Ofertas_SSCC` y `Resumir_Medidores_Central_Ventana_Oferta_Completa`, y las
fórmulas de Excel de `Medidores!K,L,N,O,R,S,T,V`. Esto resuelve el pendiente principal de la
sesión anterior.

**Hallazgo importante antes de programar:** las fórmulas muestran que `V` (y por macro, `W, X,
Y, AB, AC, AD, AE`) NO son columnas por fila de `Medidores` — son tablas auxiliares de otro
largo (central × día, central × ventana) que solo viven en esas letras de columna porque ahí
había espacio libre en la planilla. Mantenerlas como columnas `pd.NA` del mismo largo que A:U
(como se hizo la sesión anterior, cuando eran genuinamente "pendientes") dejó de tener sentido
una vez que se pueden calcular: ahora se escriben como hojas propias de `Hoja_Medidas.xlsx`
("Resumen Ofertas SSCC", "Ofertas SSCC por Dia", "Resumen Ventana Oferta"). Esto está
documentado con más detalle en `docs/Plan_Traspaso_Python_Balance_BESS.md` §20.

Cambios en `nucleo.py`:

- `LETRA_A_CAMPO` ahora va de A a U (se sacaron V, W, X, Y, AB, AC, AD, AE). `COLUMNAS_VACIAS`
  quedó en M, P, Q, U. Se eliminó `COLUMNAS_PENDIENTES_OFERTAS` (ya no queda nada pendiente).
- Nuevas funciones que replican la macro `Generar_Resumen_Ofertas_SSCC`:
  `construir_resumen_ofertas_sscc()` y sus auxiliares privados (`_contiene_bess_o_sae`,
  `_servicio_termina_en_rs`, `_es_respuesta_si`, `_normalizar_periodo`, etc., prefijo `_` como
  las de detección de SoC).
- `cargar_resumen_en_medidores()` replica `OSSCC_CargarResumenEnMedidores` (equivalente a
  `Medidores!W:Y`), incluyendo la homologación vía `Diccionario!E:F:G` y el aviso de nombres de
  `Medidores!clave` no encontrados en el diccionario.
- `calcular_r()` replica la fórmula de `R` (`VLOOKUP` contra la tabla `V:Y`), usando
  `_mapas_homologacion_fge()`/`_homologar_fge()` para la homologación específica vía
  `Diccionario!F/G→E` que usa la fórmula de `V` — un mapeo DISTINTO del que usa
  `construir_homologacion()` para el SoC (una es posicional por columna, la otra trata toda la
  fila como equivalencias simétricas). Si no hay match, `R` queda `NaN` y se registra un aviso
  en vez de fallar (no debería pasar si `Diccionario` está completo, pero no se asume).
- `calcular_s()` replica la fórmula de `S`, vectorizada por "corridas" de `Ventana` constante
  (no se reinicia por central, igual que la fórmula original).
- `construir_resumen_ventana_oferta()` replica `Resumir_Medidores_Central_Ventana_Oferta_
  Completa` (equivalente a `Medidores!AB:AE`).
- `calcular_t()` replica la fórmula de `T` (`1 - Completa`), vía merge contra el resumen
  anterior.
- `buscar_archivo_ofertas()` busca el archivo `*OfertasSSCC*` más reciente en `Ofertas/` — a
  diferencia del SoC, si hay más de uno SÍ se elige automáticamente por fecha de modificación
  (así lo hace la macro `OSSCC_BuscarArchivoOfertas` original).
- `revisar_estructura()`: la fila de Ofertas SSCC pasó de `pendiente` (no bloqueaba) a `falta`
  (bloquea Ejecutar) si no hay carpeta `Ofertas/` o no hay archivo `*OfertasSSCC*` — ahora es
  obligatoria (plan §17-18).
- `construir_medidores()` y `ejecutar()` quedaron con nuevos parámetros (`ruta_ofertas`,
  `diccionario`) y devuelven además las tres tablas auxiliares. `escribir_salida()` las escribe
  como hojas nuevas.

**Trampa encontrada y corregida en la misma sesión:** el primer intento de
`buscar_archivo_ofertas()` comparaba contra el literal `"ofertasscc"` (dos "s" seguidas), pero
`"OfertasSSCC".lower()` da tres "s" seguidas ("Ofertas" + "SSCC"). El patrón ahora se deriva en
tiempo de ejecución con `"OfertasSSCC".lower()` (constante `PATRON_NOMBRE_OFERTAS`) en vez de
transcribirlo a mano, para no repetir el error. Ver `METODOLOGIA.md` §7 si se agrega ahí.

**Verificación:** caso sintético con 1 central, 2 días, `Hora` en convención 1-24 (como usa la
planilla real, confirmado porque `(INICIO_VENTANA-1)*4=36` y `(25-INICIO_VENTANA)*4=60` dieron
exactamente los conteos de filas de las ventanas de inicio y de cierre del caso de prueba) y un
archivo de Ofertas SSCC con oferta completa las 24 horas ambos días: las 3 ventanas (inicio,
normal, última) quedaron `Completa=1` con 36/96/60 filas respectivamente, y R=1/T=0 en las 192
filas de `Medidores`. Un segundo caso con una hora "No" en vez de "Sí" marcó correctamente
`Oferta completa=0`. No se probó contra un caso real ni contra la hoja `Medidores` de la
planilla 11 (sigue pendiente, ver arriba).

---

## 2026-09-10 (4) — El archivo de SoC puede ser .csv, no solo .xlsx

El usuario aclaró que `SOC_AAMM` (el archivo de SoC) puede llegar como `.csv`, manteniendo la
misma estructura de bloques horizontales por central (`Status | Questionable | Time Stamp |
Value`) que ya soportaba `extraer_soc()`. Antes de tocar código se preguntó explícitamente por
la estructura del CSV (¿tabla larga o mantiene los bloques?) para no adivinar mal un formato
que ya funcionaba: el usuario confirmó que mantiene los bloques, solo cambia el contenedor.

Cambios en `nucleo.py`:

- `EXTENSIONES_SOC = {".xlsx", ".csv"}` (antes `buscar_soc()` exigía `.xlsx` a secas).
- Nueva función `leer_soc_crudo(ruta_soc)`: elige `pd.read_csv(ruta_soc, header=None)` o
  `pd.read_excel(ruta_soc, sheet_name=0, header=None)` según la extensión.
  `extraer_soc()` la usa en vez de llamar a `pd.read_excel` directo.
  `detectar_fila_nombres()`/`detectar_bloques()` no cambiaron: ya trabajan sobre el DataFrame
  resultante sin que les importe de dónde vino.

Probado con un CSV sintético (mismo layout de bloque que ya se probaba en Excel, separador
coma, UTF-8): `extraer_soc()` lo lee igual que un `.xlsx` equivalente, sin incidencias.

No se tocó el formato del archivo de OfertasSSCC (sigue siendo Excel, como confirma el propio
VBA que filtra por extensiones `.xlsx/.xlsm/.xlsb/.xls`) ni el de `Medidas_SAE.xlsx`.

---

## 2026-09-10 (5) — Corrección: el archivo de SoC no era .csv, era otro archivo

El usuario subió un `SOC_2607.csv` real para probar el soporte agregado en la entrada anterior,
y falló: `extraer_soc()` no encontró los encabezados `Time Stamp`/`Value`. Al inspeccionar el
archivo, resultó ser una tabla larga con columnas `Fecha_Hora, CONFIGURACION, Central, Pago,
Tipo_pago, Bloque_15min` — un archivo de pagos/liquidación, sin ninguna columna de SoC, que
solo coincidía por casualidad con el patrón de nombre "SOC"+AAMM.

Se preguntó al usuario antes de tocar más código (no se adaptó el parser a esta estructura
nueva a ciegas). Confirmó: era el archivo equivocado, "dejalo como antes es un xlsx" — el
archivo de SoC real siempre es `.xlsx`, con la estructura de bloques horizontales que
`extraer_soc()` ya soportaba desde el principio.

Se revirtió el soporte de `.csv` para el SoC agregado en la entrada anterior:

- `nucleo.buscar_soc()` vuelve a exigir `.xlsx` exclusivamente.
- Se eliminó `nucleo.leer_soc_crudo()`; `extraer_soc()` vuelve a llamar `pd.read_excel`
  directamente, como antes de esa sesión.
- Se revirtieron las menciones a "SoC puede ser .csv" en `README.md`, `MAPA.md` y
  `METODOLOGIA.md` §5 (documentos de estado actual, no de historial).
- Se agregó una trampa en `METODOLOGIA.md` §7: un archivo que matchea el patrón de nombre
  "SOC"+AAMM pero no tiene la estructura de bloques esperada no es el archivo de SoC, aunque
  comparta el nombre — no adaptar el parser a ciegas, preguntar primero.
- `docs/Plan_Traspaso_Python_Balance_BESS.md` sección 21 se corrigió en el mismo lugar (no se
  duplicó una sección nueva) para reflejar que el archivo sigue siendo siempre `.xlsx`.

Esta entrada de bitácora no borra ni edita la entrada anterior (2026-09-10 (4)) según la regla
de solo-agregar; queda como registro de que esa sesión partió de una premisa equivocada y esta
la corrigió.

---

## 2026-09-10 (6) — Hojas auxiliares de Ofertas SSCC: se unen y se recorta una

Pedido del usuario: las tablas "Ofertas SSCC por Dia" y "Resumen Ventana Oferta" deben quedar
en la misma hoja, y "Resumen Ofertas SSCC" es auxiliar — mejor no crearla como hoja.

Cambios en `nucleo.py`:

- `construir_resumen_ofertas_sscc()` sigue calculándose igual dentro de `construir_medidores()`
  (sin cambios en su lógica), pero deja de devolverse/persistirse: es un paso intermedio que
  solo hace falta en memoria para construir la tabla equivalente a `Medidores!W:Y`.
  `construir_medidores()` ahora devuelve `(df_medidores, avisos, df_wxy, df_resumen_ventana)`
  (antes devolvía también `df_resumen_ofertas`).
- Nueva función `_escribir_tabla_con_titulo(writer, hoja, df, titulo, fila_inicio)`: escribe un
  título en negrita y la tabla debajo, dentro de una hoja dada, y devuelve la fila donde debería
  empezar el siguiente bloque (para poder apilar varias tablas en la misma hoja).
- `escribir_salida()` ya no recibe `df_resumen_ofertas`; ahora escribe `df_wxy` y
  `df_resumen_ventana` una debajo de la otra en una sola hoja nueva, `HOJA_OFERTAS_SSCC =
  "Ofertas SSCC"`, cada una con su título.
- `Hoja_Medidas.xlsx` queda con tres hojas: `Medidores`, `Ofertas SSCC`, `Log` (antes tenía
  cinco).

Probado con el mismo caso sintético de sesiones anteriores: la hoja combinada queda con el
título+tabla de "Ofertas SSCC por dia" (31 filas) seguido de una fila en blanco y el
título+tabla de "Resumen ventana oferta" (3 filas), en el orden y con los valores esperados.

---

## 2026-09-10 (7) — Hojas CMg, FD, Subastas + renombre del archivo de salida

Pedido del usuario: agregar las hojas CMg, FD y Subastas, replicando las macros
`Cargar_CMg_Desde_Archivo`, `Cargar_SSCC_Desempeno_En_FD` y `Cargar_Remuneracion_Subastas_Rapido`
del documento de trazabilidad VBA, y renombrar el archivo de salida de `Hoja_Medidas.xlsx` a
`Consolidado_entradas.xlsx`.

Se leyó el código VBA completo de las tres macros y de sus funciones auxiliares
(`UltimaFilaEntreColumnasCMg`, `FiltrarFilasBESSoSAE`/`EsBESSoSAE`, `AjustarBloqueFormulas`,
`BuscarArchivoSSCCMasReciente`, `AbrirConexionExcelSubastas`/`AjustarFormulasMNSubastasRapido`/
`BuscarArchivoSubastasMasRecienteRapido`), más la sección de fórmulas del libro (`5.2 FD`,
`5.3 Subastas`) del documento de trazabilidad, ya usado en sesiones anteriores para Ofertas SSCC.

**Hallazgo estructural (igual patrón que V:Y/AB:AE de Medidores, pero por columnas):** en `FD`,
el bloque CSF (A:M, viene de `CSF Horario`) y el bloque CPF (Q:AE, viene de `CPF Horario`) son
dos tablas independientes de distinto largo que comparten la hoja en rangos de columnas
distintos, no de filas. Se escriben lado a lado (`escribir_salida()` usa `startcol` en
`df.to_excel()`), cada una con su propio número de filas.

Cambios en `nucleo.py`:

- Nuevas carpetas/archivos: `CARPETA_CMG="Cmg"` (+ `ARCHIVO_CMG="cmg.xlsx"`, nombre literal fijo,
  a diferencia de todos los demás archivos externos del proyecto), `CARPETA_SSCC_DESEMPENO=
  "SSCC_Desempeño"`, `CARPETA_SUBASTAS="Subastas"`.
- `_buscar_archivo_excel_mas_reciente()`: generaliza `buscar_archivo_ofertas()` (antes solo
  servía para Ofertas) para reusarse también en `buscar_archivo_sscc_desempeno()` y
  `buscar_archivo_subastas()`, con un flag `desde_inicio` porque Ofertas busca el patrón en
  cualquier posición del nombre, mientras que SSCC_Desempeño y Subastas exigen que el nombre
  *empiece* con el patrón (así lo hacen sus macros originales, con `Dir("patron*.*")`).
  `EXTENSIONES_OFERTAS` se renombra a `EXTENSIONES_EXCEL` (ya no es solo de Ofertas).
- `leer_cmg()`: replica `Cargar_CMg_Desde_Archivo` — hoja `CMg` o la primera si no existe,
  columnas A:I con su encabezado real (no se inventan nombres), ordenadas por columna D
  ascendente y luego H ascendente.
- `construir_fd()` + `_construir_bloque_fd_csf()`/`_construir_bloque_fd_cpf()`: replican
  `Cargar_SSCC_Desempeno_En_FD` y las fórmulas de FD (`A=str(B)&F`, `B=(DAY(D)-1)*24+E+1+
  IF(DAY(D)>100,1,0)`, `C=DAY(D)`, `K=J`, `L=K`, `M=B` para el bloque CSF; análogo con
  Q,R,S,T,U,V,AA,AC,AD,AE para el bloque CPF). El término `IF(DAY(fecha)>100,...)` se conserva
  tal cual aunque nunca sea cierto para un día real (fiel a la fórmula original, no se "limpia").
  Filtro BESS/SAE **sin** BAT (`_contiene_bess_o_sae_sin_bat()`, distinta de la de Ofertas SSCC
  que sí incluye BAT — son dos filtros reales distintos, no se fusionaron).
- `construir_subastas()`: replica `Cargar_Remuneracion_Subastas_Rapido` (que en VBA usa ADO/SQL
  contra la hoja `DB`; en Python se lee directo con pandas aplicando el mismo filtro/selección,
  sin necesitar ADO). Arma B:L (copia), M (fórmula `=K&H&I`), O/P/Q (copias de P/Y/V). `N` queda
  `pd.NA` documentada como pendiente: su fórmula real depende de `'Calculo E Costos'`, una hoja
  de una etapa posterior sin implementar — no se adivina.
- `revisar_estructura()`, `ejecutar()`, `escribir_salida()`: extendidos para validar, leer y
  escribir las tres hojas nuevas. Las tres entradas son obligatorias (bloquean Ejecutar si
  faltan), igual criterio que Ofertas SSCC.
- `ARCHIVO_SALIDA` cambia de `"Hoja_Medidas.xlsx"` a `"Consolidado_entradas.xlsx"`.

**Decisión de arquitectura (a confirmar con el usuario, ver "Pendientes abiertos"):** la macro
original de Subastas buscaba su archivo directamente en la carpeta del `.xlsm`, sin subcarpeta.
Se le creó una carpeta propia `Subastas/` para ser consistente con el resto de las entradas
externas de este proyecto (cada una con su carpeta bajo la carpeta base del caso), no porque la
planilla original lo hiciera así.

**Verificación:** caso sintético con datos para las 6 entradas (Medidas_SAE, SoC, Centrales,
OfertasSSCC, cmg.xlsx, SSCC_Desempeño_\*, 3_REMUNERACIÓN_SUBASTAS_E_ID_\*). Se verificaron a
mano los valores esperados de las fórmulas de FD (A, B, C, K, L, M del bloque CSF; Q, R, S, AC,
AD, AE del bloque CPF) y de Subastas (M, O, P, Q) contra los datos de entrada armados a
propósito, y coincidieron exactamente. `Consolidado_entradas.xlsx` quedó con las 6 hojas
esperadas: `Medidores`, `Ofertas SSCC`, `CMg`, `FD`, `Subastas`, `Log`. No se probó contra un
caso real ni contra la planilla 11.

---

## 2026-09-10 (8) — Encabezados reales de FD/Subastas + Ofertas SSCC lado a lado

El usuario entregó un Excel (`Libro1.xlsx`) con los encabezados reales de `FD` y `Subastas`, y
pidió que las dos tablas de la hoja "Ofertas SSCC" (sesión anterior, apiladas verticalmente)
queden una al lado de la otra en vez de una debajo de la otra.

**Encabezados FD/Subastas (plan §24):** confirman exactamente las fórmulas ya implementadas la
sesión anterior — nada cambió en los VALORES calculados, solo se reemplazaron los nombres de
columna por letra (`"A"`, `"B"`, ...) por los nombres reales (`NOMBRES_FD_CSF`, `NOMBRES_FD_CPF`,
`NOMBRES_SUBASTAS` en `nucleo.py`). Detalle interesante: los encabezados reales confirman que
`FD!M` y `FD!AE` repiten literalmente el nombre "Hora Mes" de `FD!B`/`FD!R` — coincide con que
también repiten su valor (`M=B`, `AE=R`), así que no hay contradicción. Como Python no permite
indexar sin ambigüedad un DataFrame con nombres de columna duplicados, el renombre se aplica con
`set_axis()` al final de cada función (`_construir_bloque_fd_csf`/`_construir_bloque_fd_cpf`),
después de terminar todos los cálculos con nombres de letra únicos — nunca antes.

También se detectó que `Subastas!K` (la columna que se filtra por BESS/SAE) se llama
"Propietario" en la vida real, y `Subastas!M` ("Ciclo") es literalmente
`Propietario & Hora_dia & Hora_mes` — nombres mucho más claros que "K", "H", "I". `Subastas!Q`
no tiene encabezado en el archivo real: se dejó sin nombre (columna `""`), no se le inventó uno.

**Ofertas SSCC lado a lado:** `_escribir_tabla_con_titulo()` ahora acepta `columna_inicio`
además de `fila_inicio`, y devuelve `(fila_siguiente, columna_siguiente)` en vez de solo la
fila — cada llamada usa el que corresponda según cómo se estén acomodando los bloques
(`escribir_salida()` ahora encadena por columna para esta hoja en particular).

Probado con el mismo caso sintético de sesiones anteriores: los encabezados de `FD` y
`Subastas` en el archivo generado coinciden letra por letra con los del `Libro1.xlsx` entregado,
y la hoja "Ofertas SSCC" quedó con "Ofertas SSCC por dia" en A1:C... y "Resumen ventana oferta"
arrancando 2 columnas después (F1:I...), ambas en la fila 1.

---

## 2026-09-11 — `Calculo E Costos`: etapa base (H + CMg + traspaso de Medidores)

El usuario pidió crear la hoja `Calculo E Costos` ("Ecostos"), cuya lógica sale de la macro
`Actualizar_Calculos_Columnas` (columna H de barras por fórmula, CMg asignado por
`Asignar_CMg_a_Calculos_Turbo`), en un archivo **separado** de `Consolidado_entradas.xlsx`
("pagos_bess o algo así por ahora"). El archivo de encabezados reales que el usuario intentó
enviar para esta hoja llegó dos veces sin la hoja `Ecostos` (solo traía `FD`/`Subastas`, ya
confirmadas la sesión anterior) — queda pendiente, ver "Pendientes abiertos".

`Actualizar_Calculos_Columnas` resultó ser una macro de ~1500 líneas con dependencias profundas
(Subastas, Resumen, Diccionario, FD, y una "Prorrata SSCC" que en el documento de trazabilidad
figuraba como fuente externa pendiente). Antes de traducir todo de una vez, se preguntó al
usuario: (a) si la Prorrata SSCC era un archivo externo disponible, y (b) cómo priorizar el
trabajo dado el tamaño de la macro. Respuestas: la Prorrata SSCC **no es un archivo externo**,
es una tabla dinámica derivable de `Subastas` (`Filas: Configuración, Hora_mes` / `Columnas:
Control` / `Valores: Cuenta de Sub_Baj`) — no es un bloqueador real. Y la prioridad elegida fue
explícitamente **"por etapas: primero H + CMg + traspaso de Medidores"**.

Esta sesión implementa exactamente esa primera etapa (plan §25):

- `_normaliza_cuarto()`: replica `NormalizaCuarto` (numérico → texto del entero redondeado,
  otro → texto recortado, vacío/error → `""`).
- `construir_dic_cmg(df_cmg)`: replica el armado del diccionario de
  `Asignar_CMg_a_Calculos_Turbo` a partir de `df_cmg` (columnas por posición, ya sin renombrar
  desde `leer_cmg()`: D=Barra, F=valor a asignar en Q, H=Cuarto de Hora). Clave
  `UCase(Barra)+"|"+NormalizaCuarto(CuartoHora)`; ante clave repetida gana la primera fila.
- `construir_mapa_barra(resumen_bess)`: resuelve `H` (antes fórmula
  `=VLOOKUP(G,Resumen!B:G,6,FALSE)`) homologando **por nombre de columna**
  (`Nombre activo`/`Barra inyección` de `Resumen BESS`) en vez de por posición, porque
  `Centrales.xlsx` no reproduce el layout `Resumen!B:G` del libro original. `leer_centrales()`
  ya devolvía `resumen` pero `ejecutar()` lo descartaba (`_, diccionario = ...`); ahora se
  captura y se usa.
- `construir_calculo_e_costos(df_medidores, mapa_barra, dic_cmg, registrar=print)`: arma A:G
  (con D↔E invertidas, igual que `Traspasar_Medidores_A_Calculos_Rapido`), `Barra` (H),
  `Energia_Positiva`/`Energia_Negativa` (I/J, energía de `Medidores!Gen_Unidad` separada por
  signo, solo si `Ventana_No_Completa=1`), `SoC` (K, copia de `Medidores!SoC`),
  `Copia_Ventana` (P, copia de `Medidores!Copia_Ventana`) y `CMg` (Q, homologado por Barra +
  Cuarto de Hora). Nombres de columna: placeholders, ver "Pendientes abiertos".
- `escribir_pagos_bess(ruta_salida, df_ecostos, registrar=print)`: escribe `Pagos_BESS.xlsx`
  (nueva constante `ARCHIVO_SALIDA_PAGOS`) con la única hoja `Calculo E Costos`
  (`HOJA_CALCULO_ECOSTOS`).
- `ejecutar()`: encadena las funciones anteriores después de escribir
  `Consolidado_entradas.xlsx`, y ahora también escribe `Pagos_BESS.xlsx`.

Explícitamente fuera de esta etapa (decisión del usuario): el resto de columnas de
`Actualizar_Calculos_Columnas` (L, M, N, O, R, S, T, U, W, X, Y, AB:AF, AG:AX, AZ) y toda la
hoja `Calculo RE545`.

**Verificación:** dos scripts de prueba sintéticos (sin persistir en el repo, borrados al
cerrar la sesión, según la convención). El primero prueba cada función nueva contra
`DataFrame`s armados a mano (incluye el caso "clave CMg repetida gana la primera fila" y el
caso "Ventana_No_Completa≠1 → energía en 0 en esta hoja"). El segundo repite las pruebas clave
pasando por un `.xlsx` real (vía `leer_centrales()`/`leer_cmg()`) para confirmar que los tipos
que devuelve `pandas`/`openpyxl` al leer un archivo real (no un `DataFrame` construido a mano)
no rompen `_normaliza_cuarto()` ni la homologación por nombre. Todos los casos coincidieron con
lo esperado a mano. No se probó contra un caso real ni contra la planilla 11.

---

## 2026-09-11 (2) — Fix: `Resumen BESS` con título arriba de los encabezados

Al correr contra un `Centrales.xlsx` real (primera vez que el usuario probó la etapa `Calculo
E Costos` fuera de un caso sintético) salió:

```
La hoja 'Resumen BESS' de Centrales.xlsx debe tener una columna de nombre de central
('Nombre activo') y una de barra de inyeccion ('Barra inyección'). Columnas encontradas:
['Cuadro N° 1: Resumen BESS', 'Unnamed: 1', ..., 'Unnamed: 8']
```

Causa: `leer_centrales()` leía `Resumen BESS` con `pd.read_excel(header=0)`, asumiendo que la
fila 1 ya traía los encabezados. El archivo real trae un título fusionado
("`Cuadro N° 1: Resumen BESS`") en esa fila, y los encabezados reales van una fila más abajo —
mismo problema que ya se había resuelto para el SoC (`detectar_fila_nombres()`), pero acá
todavía no se había aplicado el mismo criterio.

**Fix:** nueva función `_leer_resumen_bess(ruta, nombre_hoja)` que lee la hoja cruda
(`header=None`) y busca, en las primeras 15 filas, la que contiene textos que matchean
`'nombre'+'activ'` y `'barra'` (normalizados) — nunca una posición fija. `leer_centrales()` la
usa para `Resumen BESS`; `Diccionario` no se toca porque ya se leía con `header=None` sin
asumir fila fija. `construir_mapa_barra()` no necesitó cambios: ya buscaba por nombre de
columna, no por posición.

**Verificación:** test sintético con una hoja de 4 filas (título fusionado, encabezados,
2 centrales) reproduciendo exactamente la estructura del error real; `leer_centrales()` +
`construir_mapa_barra()` devolvieron el mapa esperado. Sin persistir en el repo (convención de
pruebas).

---

## 2026-09-11 (3) — Ventana nueva: diagrama de carpetas + "Generar" por salida

El usuario adjuntó `Revisor_Reliquidacion.py` (otro proyecto suyo) como referencia de cómo
quiere que se vea la ventana: un diagrama de texto de la estructura de carpetas/archivos
(prefijos `├──`/`└──`/`│`, monoespaciada) con un botón por fila cuando corresponde, en vez del
checklist plano que había hasta ahora. Pedido concreto: al lado de `Consolidado_entradas.xlsx`
un botón "Generar" que abra una ventana con una casilla por entrada (Medidas, CMg, Ofertas,
etc.) para elegir qué recalcular, y lo mismo para `Pagos_BESS.xlsx` ("ajustamos detalles
después").

Antes de tocar código se preguntó al usuario (`AskUserQuestion`) el punto más consecuente: qué
pasa con una entrada destildada al apretar Actualizar. Eligió explícitamente **"se preserva lo
que ya había"** (frente a "se regenera todo igual" o "queda vacío"), y confirmó el layout
general (carpeta+AAMM igual que hoy, árbol debajo — reemplazando al checklist plano, ya que el
árbol muestra el mismo estado OK/FALTA/PENDIENTE).

**`nucleo.py` — regeneración parcial (cambio de arquitectura, no solo de UI):**

- `_copiar_hoja_existente(wb_origen, nombre_hoja, wb_destino)`: copia una hoja completa (solo
  valores, sin fórmulas ni formato) de un workbook `openpyxl` a otro. Es el mecanismo real de
  "preservar": nunca se intenta reconstruir `Ofertas SSCC` o `FD` (bloques de distinto largo,
  con títulos y `startcol`) a partir de un DataFrame leído de vuelta — se copia la hoja física
  tal cual, evitando reinventar su estructura.
- `escribir_salida()` gana `ruta_existente` y `hojas_regenerar` (`None` = comportamiento
  clásico, regenera las 5 hojas). Cuando `hojas_regenerar` es un set, las hojas fuera de ese set
  se preservan vía `_copiar_hoja_existente()`; si no existían antes, quedan vacías y se registra
  un aviso (log de la corrida + fila del `Log`) en vez de fallar en silencio.
- `SECCIONES_CONSOLIDADO`: agrupa las 4 casillas de la ventana con las hojas que produce cada
  una. Decisión de diseño: **no hay una casilla por archivo de entrada**, sino una casilla
  `"medidores"` que junta Medidas_SAE + SoC + Centrales(Diccionario) + OfertasSSCC, porque
  `construir_medidores()` los necesita siempre los 4 juntos — tildar solo "Ofertas" y dejar
  "Medidas" destildada no permitiría recalcular nada coherente. `"cmg"`, `"fd"`, `"subastas"`
  quedan independientes porque cada uno sale de una sola función/archivo.
- `generar_consolidado(carpeta_base, aamm, secciones_activas, registrar, progreso)`: reemplaza
  a la vieja `ejecutar()`. Valida (y exige) los archivos de entrada **solo para las secciones
  tildadas** — si `"medidores"` no está tildada, no hace falta tener Medidas_SAE/SoC/Centrales/
  Ofertas presentes ni siquiera un AAMM válido.
- `generar_pagos_bess(carpeta_base, registrar, progreso)`: separado de `generar_consolidado()`
  (ya no hace todo un `ejecutar()` monolítico). Lee `Medidores` de `Consolidado_entradas.xlsx`
  ya generado (`pd.read_excel`, no se recalcula) en vez de recibir `df_medidores` en memoria
  como antes — refleja que ahora son dos flujos independientes disparados por botones distintos.
  Perdió el parámetro `aamm` (no lo usaba: todo sale de `Medidores`, que ya trae Mes/Dia/Hora).
- La vieja `ejecutar(carpeta_base, aamm, ...)` se **eliminó** (no solo se dejó como wrapper):
  nada la llama ya que la ventana dispara `generar_consolidado`/`generar_pagos_bess` por
  separado, y mantenerla como código muerto no aportaba nada.

**`Balance_BESS.py` — reescritura de la ventana:**

- `_profundidad_fila()`, `_es_ultimo_en_su_nivel()`, `_prefijos_arbol()`: arman los prefijos
  tipo árbol a partir de la lista plana `(etiqueta, estado, detalle)` que ya devolvía
  `revisar_estructura()`, sin pedirle a `nucleo.py` que sepa de árboles/interfaz. La profundidad
  se deduce del TEXTO de la etiqueta (`"algo/"` = carpeta de primer nivel, `"  hoja ..."` =
  nieto, el resto = archivo hijo de la carpeta anterior) en vez de que `nucleo.py` devuelva un
  campo de profundidad — mantiene a `nucleo.py` sin conceptos de UI.
- El panel "Entradas detectadas" (`pintar_checklist`, filas planas) se reemplaza por "Estructura
  del caso" (`pintar_arbol`, con los prefijos de arriba). Al final del árbol se agregan a mano
  las dos filas de salida (`Consolidado_entradas.xlsx`, `Pagos_BESS.xlsx`), cada una con su
  botón "Generar...".
- Se eliminó el botón único "Ejecutar": cada salida se dispara desde su propia ventana
  (`abrir_ventana_generar_consolidado`, `abrir_ventana_generar_pagos`), con casillas (la primera)
  o solo una explicación (la segunda, sin casillas todavía). `lanzar_generacion()` es el helper
  compartido: corre la función de `nucleo` en un hilo, y hace que `registrar`/`progreso` escriban
  en el log/barra de la ventana PRINCIPAL (no hay log propio por ventana secundaria), para no
  duplicar esos widgets.
- Al terminar una generación se vuelve a llamar `revisar()` para refrescar el árbol (por si el
  archivo de salida cambió de OK/FALTA), y se cierra la ventana "Generar" correspondiente.

**Verificación:** sin entorno gráfico disponible en esta sesión (no hay `tkinter` instalado acá,
solo se pudo correr `python -m py_compile` sobre `Balance_BESS.py`/`nucleo.py`). Se probó por
separado, con scripts sintéticos borrados al cerrar la sesión:
- La lógica de árbol (`_profundidad_fila`/`_prefijos_arbol`, copiadas fuera de `Balance_BESS.py`
  para poder importarlas sin `tkinter`) contra la lista real que devuelve
  `revisar_estructura()` en una carpeta de prueba: el árbol impreso en consola tiene la forma
  esperada.
- `escribir_salida()` con `hojas_regenerar`: preserva hojas no tildadas desde `ruta_existente`,
  regenera las tildadas, y avisa (sin fallar) cuando una hoja a preservar no existía todavía.
- `generar_consolidado()`: no exige Medidas/Ofertas/Centrales si `"medidores"` no está tildada;
  corre de punta a punta con solo `"cmg"` tildada; rechaza secciones desconocidas y el caso de no
  tildar nada.
- `generar_pagos_bess()`: exige `Consolidado_entradas.xlsx` con `Medidores` antes de correr, y
  de punta a punta con datos reales (vía `Centrales.xlsx`/`cmg.xlsx` reales) da los mismos
  resultados que la implementación anterior.

**No probado:** la ventana real (no hay `tkinter` en este entorno) — falta que el usuario la
abra y confirme que el diagrama se ve como esperaba y que los botones "Generar" funcionan en la
práctica.

---

## 2026-09-11 (4) — `Calculo E Costos`, etapa 2: L, N, O, R, S, T, U, W, X, Y, AB, AC, AD

El usuario pidió terminar el cálculo de Ecostos ("necesito que termines con el calculo de
Ecostos"). Antes de implementar se leyó completa la macro `Actualizar_Calculos_Columnas`
(módulo `J_Calculo_Ecostos`, ~1500 líneas) y sus funciones auxiliares (`CrearDiccionarioSubastas`,
`CrearDiccionarioProrrata`, las funciones de ordenamiento por bloques, `CalcularAsignacionEnergia`,
etc.), y se encontró un problema real: buena parte de las columnas restantes (`M`, `AE`, `AF` y
todo `AG:AZ`) depende de una hoja `Resumen` del libro original — **distinta** de
`Centrales.xlsx!Resumen BESS` — con una tabla central→factor y un umbral único (`H8`), y de un
umbral de subida/bajada por central+ventana. Ninguna de las dos cosas está mapeada en la
migración. Se le preguntó al usuario 3 veces con distinto nivel de detalle técnico (la primera
con letras de columna VBA, que no se entendió — "no t entendí" / "no entiendo dime el valoer");
la pregunta se simplificó a lenguaje llano y tampoco se resolvió del todo, pero el usuario dio
una pista concreta que sí resolvió una pieza clave (ver abajo). **Las otras dos (hoja `Resumen`
y umbral de subida/bajada) siguen abiertas** — ver "Pendientes abiertos".

**El problema de `L` y cómo se resolvió:** la fórmula real de `L` compara
central+mes+día+hora de `Calculo E Costos` contra `Subastas`, filtrando por un "tipo"
BAJADA/SUBIDA. El código VBA documentaba esa clave por posición (`Subastas!D` = tipo,
`G,H,I,K` = el resto de la clave), pero esas letras no coinciden con los encabezados reales de
`Subastas` ya confirmados (`D` es `Fecha`, no un texto BAJADA/SUBIDA) — se le pidió al usuario
que revisara la macro `Traspasar_Medidores_A_Calculos_Rapido` por si la respuesta estaba ahí (no
estaba: se confirmó con `grep` que esa macro no menciona nada de Subastas/BAJADA/SUBIDA/Resumen).
El usuario confirmó directamente: **"Es la columna C de la hoja subastas que ya generamos"** —
`Subastas!Sub_Baj`, tal como sugería el propio nombre de la columna. Con eso se reconstruyó la
clave equivalente por NOMBRE de columna real (no por posición): tipo=`Sub_Baj`,
central=`Configuración` (esta última **inferida**, no confirmada letra por letra: es el mismo
campo que usa la tabla dinámica Prorrata SSCC como identificador de central, y da una
alineación semántica limpia con Mes/Dia/Hora_dia; se descartó `Propietario` porque es el campo
que se usa para filtrar por BESS/SAE, no para identificar una central puntual). Queda como
pendiente validar esto contra un caso real (ver "Pendientes abiertos").

**Implementado en `nucleo.py`** (todo lo que NO depende de la hoja `Resumen`):

- `_normaliza_valor_vba(valor)`: replica `NormalizarValor` (mayúsculas+recorte; números sin
  decimales de más, igual que `CStr` en VBA) para armar claves compuestas comparables entre
  `Calculo E Costos` y `Subastas`.
- `calcular_l(df_ecostos, df_subastas)`: ver arriba.
- `calcular_n_o(df_ecostos)`: por grupo (central+ventana), suma acumulada de energía positiva/
  negativa (solo filas `L=1`) ordenando por `Cuarto de Hora` descendente, repartida a todas las
  filas que comparten ese `Cuarto de Hora`.
- `calcular_r_ecostos(df_ecostos)`: ranking por grupo, `CMg` descendente + `Cuarto de Hora`
  descendente, con empates compartiendo el ranking de inicio del bloque ("competition
  ranking"). **Nombre con sufijo `_ecostos` a propósito** — ver el error de abajo.
- `calcular_s_t_u`, `calcular_w_x`, `calcular_y_ab_ac_ad`: ver plan §25.7 para el detalle de
  cada una (`W` tiene una excepción fiel al original: la primera fila de todo el archivo no
  reinicia a 1, toma el valor de `Hora`).
- `completar_calculo_e_costos_grupos(df_ecostos, df_subastas, registrar=print)`: combina todas
  las anteriores. `generar_pagos_bess()` ahora también lee la hoja `Subastas` de
  `Consolidado_entradas.xlsx` (además de `Medidores`) y la llama después de
  `construir_calculo_e_costos()`.

**Error encontrado y corregido durante esta sesión (antes de comitear):** la primera versión de
`calcular_r_ecostos` se llamaba simplemente `calcular_r()`, pisando en silencio a la función que
YA existía con ese nombre para Medidores (`Oferta_Completa_Dia`, lógica no relacionada). Python
no avisa de la redefinición; recién se manifestó como un `TypeError` de argumentos al correr el
test de punta a punta (`construir_medidores()` llama a `calcular_r()` en tiempo de ejecución, y
para ese momento el nombre ya apuntaba a la versión nueva de 1 argumento). Se corrigió
renombrando a `calcular_r_ecostos` y se agregó una trampa en `METODOLOGIA.md` §7: antes de
agregar una función nueva, `grep` para confirmar que el nombre no existe ya.

También se corrigió, en el mismo pase de pruebas, un bug real en `calcular_w_x`: la primera
versión calculaba `W` como "posición dentro del bloque" en vez de "contador que suma 1 desde el
valor de la fila anterior", así que la excepción de la primera fila (`W` = `Hora` en vez de 1)
no se arrastraba al resto de su bloque. Se corrigió sumando el offset (`Hora - 1`) a todas las
filas del primer bloque.

**Verificación:** un test sintético con 2 grupos (central+ventana) y valores elegidos a mano
para poder calcular cada columna manualmente de antemano (incluye empates en `CMg`+`Cuarto de
Hora` para `R`, una fila con `Energia_Positiva=0` para probar el "no calificaI" de `Y`/`AB`, y
un cambio de `Copia_Ventana` a mitad de archivo para probar que `W` es realmente global y no por
grupo) — todas las columnas coincidieron con el cálculo a mano. Un segundo test corrió
`generar_pagos_bess()` de punta a punta con `Centrales.xlsx`/`cmg.xlsx`/`Consolidado_entradas.xlsx`
(con hojas `Medidores` y `Subastas`) reales, confirmando que las 13 columnas nuevas aparecen en
`Pagos_BESS.xlsx` con los tipos y valores esperados. Sin persistir en el repo (convención de
pruebas). No se probó contra un caso real ni contra la planilla 11 — sigue pendiente, y ahora es
más urgente por la inferencia de `L` sin confirmar.

---

## 2026-09-11 (5) — `Calculo E Costos`: M, AE, AF + nombres reales de toda la hoja

El usuario adjuntó `Centrales.xlsx` (real) y un `Libro1.xlsx` con 4 hojas: `FD`, `subastas`,
`E COSTOS` y `Resumen` — respondiendo al bloqueo de la sesión anterior ("Que necesitas de la
hoja original de resumen porque eso es todo lo que hay... puede que los indices esten
diferente").

**Hallazgo clave**: la hoja `Resumen` del libro original y `Centrales.xlsx!Resumen BESS` son
**la misma tabla** — mismos 9 encabezados en el mismo orden (`Nombre activo`, `Pmax (MW)`,
`Horas para descarga forzada`, `Capacidad (MWh)`, `Energía mínima`, `Barra inyección`, `%
Energía sobre mínima (indicador nuevo ciclo)`, `Ciclos max diarios`, `Eficiencia`). No hacía
falta una hoja nueva: el "factor" de `AE`/`AF` es la columna `Pmax (MW)`, y el "umbral" de `M`
es el valor de `% Energía sobre mínima...` en la primera fila de datos (que en el archivo real
cae justo en `H8`, de ahí la referencia fija del VBA original).

**Implementado en `nucleo.py`:**

- `construir_dic_resumen_factor(resumen_bess)`: arma central→`Pmax (MW)` y el umbral de SoC
  mínimo, a partir de la MISMA hoja que ya usa `construir_mapa_barra()` (misma búsqueda por
  nombre de columna normalizado, no por posición).
- `calcular_m(df_ecostos, umbral_soc_minimo)`: `1` si `SoC > umbral`, si no `0`.
- `_calcular_asignacion_energia(bloque, energia_maxima, factor)` + `calcular_ae_af(df_ecostos, dic_factor)`:
  replica `CalcularAsignacionEnergia` — reparte el máximo de `N`/`O` del grupo en bloques de 15
  minutos según el orden `W` y el factor de la central; sin factor o factor=0, queda en blanco
  (`pd.NA`) en vez de fabricar un error de Excel. `Int()` de VBA se replica con `math.floor`
  (redondea hacia abajo incluso en negativos, distinto de truncar hacia cero).
- `completar_calculo_e_costos_grupos()` ahora también agrega `M`, `AE`, `AF`, y al final
  renombra TODAS las columnas con `NOMBRES_CALCULO_E_COSTOS` (nombres reales, de la hoja
  `E COSTOS` de `Libro1.xlsx` — fila 3 tiene el encabezado real de cada columna). Mismo patrón
  que `NOMBRES_FD_CSF`/`NOMBRES_SUBASTAS`: todo el cálculo interno sigue usando los nombres/
  letras de siempre, el rename es el último paso antes de escribir.
- `generar_pagos_bess()` ahora también arma `dic_factor`/`umbral_soc_minimo` (de la misma
  lectura de `Centrales.xlsx` que ya hacía para `mapa_barra`) y se los pasa a
  `completar_calculo_e_costos_grupos()`.

**Confirmación indirecta de la columna `L`**: el archivo de encabezados reales muestra que
`Calculo E Costos!G` se llama literalmente `Configuracion` — el mismo nombre de campo que
`Subastas!Configuración`, que es lo que se venía usando (inferido) para homologar centrales en
`calcular_l()`. Sube la confianza en esa elección, aunque sigue sin confirmarse fila por fila.

**Lo que sigue pendiente (`AG:AZ`)**, según la misma hoja `E COSTOS`:
- `AG:AL` ("Prorratas") necesitan la tabla dinámica Prorrata SSCC, todavía no construida.
- `AM:AR` ("FD") necesitan una categoría `CTF` que no existe en nuestra hoja `FD` (solo tiene
  CSF/CPF) — origen sin identificar.
- `AW` (Descuento FD) necesita el umbral de subida/bajada de Subastas: el archivo de
  encabezados reales muestra `Subastas!R:V` (`Configuración`, `Ciclo`, `Clave`, `SUBIDA`,
  `BAJADA`) como una tabla de resumen aparte, pero sin filas de datos de ejemplo, y no coincide
  con `Subastas!U:W` que usaba el código VBA — la posición real sigue sin confirmarse.

**Verificación:** tests sintéticos (sin persistir en el repo) para `construir_dic_resumen_factor`
(incluye el caso "primera fila con nombre válido" cuando hay filas sin central), `calcular_m`,
`_calcular_asignacion_energia` (bloque completo y bloque parcial con fracción), `calcular_ae_af`
(incluye el caso "central sin factor -> NA"), y que `completar_calculo_e_costos_grupos()`
devuelve exactamente las columnas de `NOMBRES_CALCULO_E_COSTOS` en ese orden. Se repitió también
el test de regresión completo de la etapa 2 (2 grupos, empates en ranking, etc. de la sesión
anterior) contra los nombres reales, sin cambios en los valores esperados. Un test de punta a
punta corrió `generar_pagos_bess()` completo con archivos reales de `Centrales.xlsx`/`cmg.xlsx`/
`Consolidado_entradas.xlsx`, confirmando que las 28 columnas de `Pagos_BESS.xlsx` salen con los
nombres reales y los tipos esperados. No se probó contra un caso real ni contra la planilla 11.

---

## 2026-09-11 (6) — `Calculo E Costos`, etapa 3: `AG:AV` (Prorratas, FD homologado, costo ponderado)

El usuario aportó dos confirmaciones cortas que destrabaron esta etapa: **"creo que no tiene
ctf no está en los FD y en la hoja de los ecostos sale con 0"** y **"Las subastas que te mande
vs las del original están corridas una columna, la primera en el original está vacía"**.

La segunda explica retroactivamente la discrepancia de letras encontrada la sesión anterior
para la columna `L`: aplicando ese corrimiento de una columna a lo que documentaba el VBA
(`Subastas!D`=tipo, `G,H,I,K`=clave), se obtiene exactamente `Sub_Baj` + `Configuración+Mes+
Dia+Hora_dia` — lo mismo que ya se había implementado por inferencia, ahora con una explicación
clara. La primera confirma que `CTF` (columnas `AI`, `AL`, `AO`, `AR`) no necesita ningún
origen: en el VBA original están hardcodeadas en 0 (`salidaAGAX(i,3)=0`, etc.), nunca dependen
de un diccionario — coincide exactamente con lo que el usuario reportó ver en el archivo real.

**Implementado en `nucleo.py`** (todo lo que no depende del umbral de subida/bajada, que sigue
sin resolverse — ver "Pendientes abiertos"):

- `construir_prorrata_sscc(df_subastas)`: arma la tabla dinámica Prorrata SSCC con
  `pandas.pivot_table` directamente desde `Subastas` (`index=[Configuración, Hora_mes],
  columns=Control, values=Sub_Baj, aggfunc=count`) — confirmado hace varias sesiones que NO es
  un archivo externo, pero recién ahora se construye en Python.
- `construir_dic_prorrata()`: busca entre las columnas que deja el pivot la que contenga "cpf"
  y la que contenga "csf" en el nombre (inferido, no confirmado que `Control` tenga esos dos
  valores exactos — avisa si no las encuentra, no falla).
- `calcular_prorratas()`: homologa por central+`Hora Mes` → `(AG, AH)`. `AJ=AG`, `AK=AH`
  (duplicados a propósito, así lo hace el VBA original: `salidaAGAX(i,4)=valorAG`). `AI=AL=0`.
- `construir_dic_mapeo_diccionario()`: TERCERA lectura de la hoja `Diccionario` (columna A→B,
  primera coincidencia gana) — distinta de `construir_homologacion()` y de
  `_mapas_homologacion_fge()`, documentado como trampa nueva en `METODOLOGIA.md` §7.
- `_calcular_bloque()` + `construir_dic_fd_bloque()` + `calcular_fd_prorrateado()`: homologan la
  central contra `Diccionario`, arman una clave "bloque de 4 + central homologada" (uno para
  descarga usando `Y`, otro para carga usando `AC`) y buscan esa clave en `FD!CSF(±)`/`CPF(±)`
  → `(AM, AN, AP, AQ)`. Central no encontrada en `Diccionario` → blanco (`pd.NA`); central
  encontrada pero sin match en `FD` → 0 (fiel al original, que solo registra un aviso). `AO=AR=0`.
- `_calcular_costo_ponderado()` + `calcular_as_at()`: replica `CalcularCostoPonderado` → `(AS,
  AT)`, combinando las Prorratas, el FD homologado y `AE`/`AF`.
- `calcular_au_av()`: promedio de `AB`/`AD` por grupo (central+ventana), activado solo si la
  suma GLOBAL de energía por ventana (TODAS las centrales que comparten esa `Copia_Ventana`, sin
  agrupar por central — una agrupación distinta de la de `N/O/R/Y/AB/AC/AD`) supera ±10 → `(AU,
  AV)`.
- `completar_calculo_e_costos_grupos()` ahora recibe también `diccionario`, `df_fd_csf` y
  `df_fd_cpf`, y agrega las 16 columnas nuevas antes del rename final.
- `generar_pagos_bess()` ahora también busca y lee el archivo `SSCC_Desempeño_*` (con
  `buscar_archivo_sscc_desempeno()` + `construir_fd()`, igual que ya hacía `generar_consolidado()`
  para la sección `"fd"`) y mantiene `diccionario` de `leer_centrales()` en vez de descartarlo.

**`NOMBRES_CALCULO_E_COSTOS` gana valores DUPLICADOS a propósito**: `AG:AL` ("Prorratas") y
`AM:AR` ("FD") comparten los mismos 6 nombres cortos (`CPF(-)`, `CSF(-)`, `CTF(-)`, `CPF(+)`,
`CSF(+)`, `CTF(+)`) porque así están en el archivo real (se distinguen por un encabezado de
grupo en las filas 1-2 que no se replica en nuestro esquema de una sola fila de encabezado) —
mismo criterio que el `"Hora Mes"` duplicado de `FD`. Documentado como trampa en
`METODOLOGIA.md` §7 (indexar por ese nombre después del rename da una `Series` ambigua).

**Verificación:** tests sintéticos (sin persistir en el repo) para cada función nueva por
separado con valores calculados a mano (incluye el caso "central sin match en Diccionario ->
blanco", "central con match pero sin FD -> 0", `CalcularCostoPonderado` con energía/precio en
blanco, y `AU`/`AV` con una tercera central en el mismo `Copia_Ventana` para probar que la suma
global cruza centrales). Un test llamó a `completar_calculo_e_costos_grupos()` completo y
confirmó que el orden final de columnas coincide exactamente con `NOMBRES_CALCULO_E_COSTOS`. Un
test de punta a punta corrió `generar_pagos_bess()` con un archivo `SSCC_Desempeño_*` sintético
pero con la estructura real que espera `construir_fd()` (datos desde la fila 12, filtro BESS/SAE
en columna D) — corrió sin errores; la única discrepancia fue que `pd.read_excel` renombra
columnas duplicadas al releer (`"CPF(-)"` → `"CPF(-).1"`), un comportamiento conocido de pandas
al leer, no un problema de lo que se escribió (confirmado escribiendo y releyendo un `DataFrame`
con columnas duplicadas de prueba). No se probó contra un caso real ni contra la planilla 11.

---

## 2026-09-11 (7) — `Calculo E Costos`, etapa 4: `AW`, `AX`, `AZ` (+ `Subastas!N`) — hoja completa

El usuario pidió terminar lo que quedaba de Ecostos ("Termina lo pendiente de E Costos porfa. Lo
que sigue pendiente (AW, AX, AZ)"). Al empezar la sesión el repositorio **no tenía** el documento
de trazabilidad del `.xlsm`: las sesiones anteriores lo habían leído como adjunto, pero lo único
que quedó escrito fue *qué* faltaba, no el código VBA ni las fórmulas de las tres columnas. Sin
eso había que adivinar la lógica, que es justo lo que prohíben `REGLAS.md` y el plan §18, así que
se pidió el documento antes de tocar código. El usuario lo entregó y pidió dejarlo en el repo:
ahora vive en **`docs/Trazabilidad_11_PAGOS_BESS_2607_Definitivo.md`** (331 KB, incluye el VBA
completo y todas las fórmulas del libro), y `README.md` lo agrega a la tabla de navegación. Es la
causa raíz de este bloqueo: no vuelve a pasar.

**El "bloqueante" resultó no existir.** Lo que estaba anotado como una dependencia circular
(`AW` → umbral → `COUNTIFS` sobre `Subastas!N` → `Calculo E Costos`) se deshizo al leer la
fórmula real de `Subastas!N` (sección 5.3 del documento):

```
=IFERROR(XLOOKUP(1,('Calculo E Costos'!$D$2:$D$50000=J3)*('Calculo E Costos'!$G$2:$G$50000=K3),
        'Calculo E Costos'!$P$2:$P$50000,""),"")
```

`Calculo E Costos!P` es `Ciclo de Carga del mes` (= `Copia_Ventana`), que viene de `Medidores` y
está disponible desde la etapa base: **no depende de ninguna columna calculada**. Y de paso: la
columna `Energía SSCC` no es una energía, es un número de ciclo. El nombre engañaba.

**Implementado en `nucleo.py`** (sección nueva "etapa 4", después de `calcular_au_av`):

- `calcular_subastas_energia_sscc(df_subastas, df_ecostos)`: replica el XLOOKUP de arriba
  homologando por NOMBRE de columna (`Hora_mes` + `Configuración` de nuestra hoja `Subastas`
  contra `Hora Mes` + `clave` de Ecostos), con el mismo corrimiento de una columna ya confirmado
  para `L`. Primera coincidencia gana (XLOOKUP sin modo de búsqueda); sin coincidencia, `""`.
- `construir_dic_umbrales_subastas(df_subastas, energia_sscc)`: replica la tabla `Subastas!U:W`
  del libro original (`U = S&"&"&T`, `V`/`W` = `COUNTIFS` por SUBIDA/BAJADA). **No era un archivo
  externo ni una hoja aparte**: se deriva de `Subastas` + `Subastas!N` contando filas por
  central+ciclo+tipo, exactamente como la Prorrata SSCC. Diferencia deliberada con el original:
  las combinaciones salen de los datos en vez de una lista fija escrita a mano — es equivalente,
  porque una combinación sin datos daría umbral 0 y con umbral 0 ninguna fila pasa el filtro
  `W <= umbral*4` (`W` arranca en 1), o sea `AW = 0` de las dos formas.
- `_clave_central_ciclo()`: la clave `central&ciclo`, que en el original se arma distinto de cada
  lado (concatenación de Excel en `Subastas!U`, `TextoSeguro(G)&"&"&TextoSeguro(P)` en el VBA) y
  tiene que dar lo mismo. Usa `_valor_clave()` para que un `3` y un `3.0` den los dos `"3"`.
- `calcular_aw_ax(df_ecostos, dic_umbrales)`: replica `AW` y `AX` del bloque "AU, AV, AW, AX Y AZ".
  **Ojo con un detalle que parece un error de tipeo y no lo es:** los promedios van cruzados —
  `AB` se promedia con el umbral de **BAJADA** y `AD` con el de **SUBIDA**. Así está en el VBA
  original, se replicó tal cual.
- `calcular_az(df_ecostos)`: `MAX(0, (SUMA(AX) - SUMA(U)) / cantidad de filas)` por grupo
  central+ventana, mismo valor en todas las filas del grupo.
- `NOMBRES_CALCULO_E_COSTOS` suma `AW` → `Descuento FD`, `AX` → `Total`, `AZ` → `Monto a
  compensar`. **`Total` queda duplicado a propósito** (es también el nombre real de `U`), igual
  que los `CPF(-)`/`CSF(-)`/... repetidos entre `AG:AL` y `AM:AR`: en el archivo real se
  distinguen por el encabezado de grupo de las filas 1-2, que este esquema de una sola fila de
  encabezado no replica. Para llegar sin ambigüedad a una de esas columnas hay que ir por
  posición, no por nombre (los tests de esta sesión lo hacen así).

`AY` **no** se calcula: la macro salta de `AX` a `AZ` y el documento no muestra ninguna región de
fórmulas para `AY4:AY26787` (sección 5.4). Se documentó explícitamente para que no parezca un
olvido.

**Verificación** (tests sintéticos, sin persistir en el repo, según la convención):

- `Subastas!N`: cruce por Hora_mes+central, incluido el caso "no hay fila que cruce → `''`" y el
  caso "gana la primera coincidencia".
- Umbrales: los conteos SUBIDA/BAJADA salieron exactamente los esperados a mano, incluido un tipo
  distinto de SUBIDA/BAJADA (no cuenta), una fila sin ciclo (no entra en ninguna clave) y la
  normalización `5.0` → `"5"` en la clave.
- `AW`: un grupo armado a propósito con `W = [1, 4, 5, 9]` y umbrales (subida 2, bajada 1), donde
  los filtros `W<=4` y `W<=8` dejan subconjuntos distintos — promedios 15 y 4 calculados a mano,
  y las 4 filas coincidieron. Casos borde: grupo sin umbrales → 0; umbral 0/0 → 0; `AE` en blanco
  → 0 **solo en esa fila**; `cantidadADW = 0` → 0 en **todo** el grupo.
- `AX = AU + AV - AW` fila a fila, y `AZ` constante por grupo, nunca negativo (probado con una
  suma de `U` enorme) y dividido por la cantidad de filas del grupo.
- Corrida completa de `completar_calculo_e_costos_grupos()` con datos sintéticos: devuelve las 48
  columnas de `NOMBRES_CALCULO_E_COSTOS` en orden, las tres nuevas numéricas en todas las filas.

No se probó contra un caso real ni contra la planilla 11 — sigue siendo el pendiente principal.

---

## 2026-09-11 (8) — `Calculo RE545`, etapa base (`A:V`)

El usuario pidió seguir con `Calculo RE545` y entregó `Calculo_RE545_reducido_para_IA.xlsx`
(hoja real recortada: fila 3 = nombres de columna, filas 1-2 = títulos de grupo, hoja
`Mapa_Formulas` con todas las familias de fórmulas, columna `CF` con la fila original). Se copió
al repo como `docs/Calculo_RE545_reducido_para_IA.xlsx` — mismo criterio que el documento de
trazabilidad: si es la fuente de una decisión, vive en el repo, no en un adjunto de sesión.

Avisó además que quedaba poco contexto y que se iba a dormir, así que esta entrada se escribe
con el detalle suficiente para que la próxima sesión continúe sin preguntarle nada.

**Hallazgo estructural:** `Calculo RE545` es casi toda **fórmulas en la hoja**, no valores
escritos por macro (al revés que `Calculo E Costos`, donde la macro J escribe casi todo). Las
únicas columnas que escribe el VBA son las del traspaso y `Q`/`R`.

**Cómo se reparten las dos hojas** (`Traspasar_Medidores_A_Calculos_Rapido` recorre `Medidores`
una sola vez y escribe en las dos): `A:G`, `K` y `P` van **iguales a las dos**; lo que se reparte
es la energía, según `Medidores!T` (`Ventana_No_Completa`): `= 1` → E Costos; cualquier otro
número, vacío, no numérico o error → RE545. Solo RE545 recibe además `T` (= `Medidores!L`,
"Ventana de valorizacion") y `R` (`CMg!I`, "CMg Promedio", porque `CompletarDestinoTurbo` se
llama con `escribirR:=True` para esta hoja y `False` para la otra).

**Cambios en `nucleo.py`:**

- `construir_dic_cmg()` ahora guarda **el par** `(CMg!F, CMg!I)` en vez de solo `CMg!F` — son los
  dos elementos del `Array()` del diccionario VBA. Se agregó `_buscar_cmg(dic_cmg, barra,
  cuarto_hora)` como helper compartido (replica `CompletarDestinoTurbo`), y
  `construir_calculo_e_costos()` ahora toma `[0]` de ese par. **Si algo se rompe en CMg, mirar
  acá primero**: es un cambio en una estructura que ya usaban las dos hojas.
- `construir_calculo_re545()`: etapa base (traspaso + Barra + Q/R + T).
- `construir_dic_resumen_eficiencia()`: central → `Eficiencia` de `Resumen BESS` (la 9na columna
  de las 9 de esa tabla = el `VLOOKUP(...,9,0)` de `V`). Es una **tercera** lectura de esa hoja,
  distinta de `construir_mapa_barra()` y `construir_dic_resumen_factor()` — no fusionar.
- `calcular_s_re545()`: el "ranking cmg" de esta hoja. **No es el mismo** que el de E Costos:
  agrupa por central + `T` (ventana de valorización) y ordena por `CMg Promedio` + `Hora`, no por
  `CMg` + ciclo.
- `calcular_u_v_re545()`: `U` (`EiniT`) = `SoC % × Pmax (MW) × 1000`; `V` (`EalmT`) = la carga
  total del grupo central+ventana cambiada de signo, por la `Eficiencia`. Central que no está en
  `Resumen BESS` → las dos en blanco (equivale al `#N/A` del VLOOKUP).
- `completar_calculo_re545()`: agrega `L`, `M`, `N`, `O`, `S`, `U`, `V` y renombra a
  `NOMBRES_CALCULO_RE545`. `L`, `M`, `N` y `O` son **literalmente las mismas fórmulas** que en E
  Costos, así que se reusan `calcular_l()`, `calcular_m()` y `calcular_n_o()` (verificado contra
  la fórmula real de RE545, que es la versión explícita de lo que esa función ya hacía).
- `escribir_pagos_bess()` acepta un `df_re545` opcional y escribe la hoja `Calculo RE545`
  (`HOJA_CALCULO_RE545`) en `Pagos_BESS.xlsx`; `generar_pagos_bess()` la arma y la pasa.

**Trampa nueva, importante:** `R`, `S`, `T`, `U` y `V` existen en las dos hojas y **significan
cosas distintas** en cada una (`U` es "Total" en E Costos y "EiniT" en RE545). Nunca reusar una
función de una hoja en la otra sin leer antes la fórmula real de la columna.

**Verificación** (tests sintéticos, sin persistir en el repo): el reparto de energía entre las dos
hojas (incluido el caso `Ventana_No_Completa` vacío → RE545) y que la hoja hermana sigue dando lo
complementario; `T` = `Medidores!L`; `Q`/`R` del par de CMg; `S` con un grupo armado a propósito
con empates de `CMg Promedio` resueltos por `Hora` (valores calculados a mano: 1.5, 1.25, 1.0,
1.75, 1.0); `U`/`V` con una central sin ficha en `Resumen BESS` → blanco; y la corrida completa
de `completar_calculo_re545()` devolviendo las 22 columnas con los nombres reales. Se corrieron
también los tests de la etapa 4 de E Costos como regresión (el cambio de `construir_dic_cmg`
podía romperlos): pasan.

**Lo que sigue** (detalle por bloque en el plan §26.3): `AC:AU`, `AW:BG` y `BI:CE`.

---

## 2026-09-11 (9) — `Calculo RE545`, etapa 2: `AC:AU` (reservas por subasta)

Segunda etapa de RE545, en la misma sesión (el usuario pidió avanzar sin consultarlo).

`AC:AU` son tres bloques de 6 columnas con los **mismos 6 encabezados** (`CPF(-)`, `CSF(-)`,
`CTF(-)`, `CPF(+)`, `CSF(+)`, `CTF(+)`), que se distinguen por el título de grupo de la fila 2:
"Subastas" (`AC:AH`), "FD" (`AI:AN`) y "FMA" (`AO:AT`). Los tres son el mismo `SUMIFS` contra
`Subastas` cambiando solo la columna sumada, y `AU` = `SUMPRODUCT` de los tres bloques `/4*1000`.

- `construir_dic_reservas_subastas(df_subastas)`: arma los tres diccionarios
  `(central, hora del mes, tipo) -> suma`. Criterios por NOMBRE (`Configuración`, `Hora_mes`,
  `Control`), columnas sumadas por POSICIÓN (`O`, `P`, `Q`).
- `calcular_reservas_re545(df_re545, dics)`: devuelve las 18 columnas + `AU`.
- `NOMBRES_CALCULO_RE545` suma los 18 nombres repetidos + `SUMA Reservas*FMA*FD`.

**Decisión documentada (pendiente de validar, no bloquea):** los nombres reales de `Subastas`
llaman `FD` a `O` y `FMA` a `P` — corridos una columna respecto de los títulos de grupo de RE545,
que dicen Subastas/FD/FMA para `O`/`P`/`Q`. Es el mismo corrimiento de una columna que el usuario
ya había descrito para el archivo de Subastas. Se siguió **la fórmula** (posición), no el nombre,
porque la fórmula es la fuente primaria. Si al validar con datos reales los tres bloques salen
corridos entre sí, esto es lo primero que hay que revisar (plan §26.3).

**Verificación:** tests sintéticos con dos filas de `Subastas` que comparten central+hora+tipo
(para probar que el `SUMIFS` suma y no pisa), un tipo que no existe en los datos (→ 0, no blanco),
una fila de RE545 cuya hora del mes no cruza con nada (→ los 18 en 0 y `AU` en 0) y `AU`
calculado a mano (`(8x4x1 + 7x4x0.25)/4*1000 = 9.750`). Regresión completa de E Costos (etapas
2-4) y de la etapa base de RE545: pasan.

---

## 2026-09-11 (10) — `Calculo RE545`, etapa 3: `AW:BG` (resumen por central + ventana)

Tercera etapa de RE545 en la misma sesión.

**Hallazgo estructural:** `AW:BG` **no son más columnas del bloque principal**: son una tabla
aparte de 288 filas (9 centrales × 32 ventanas) contra las 26.787 del bloque principal,
compartiendo la hoja. Tercer caso del mismo patrón en este proyecto (ya había pasado con los
bloques CSF/CPF de `FD` y con las dos tablas de `Ofertas SSCC`). Se escribe al lado del bloque
principal con una columna en blanco de separación (`escribir_pagos_bess()` ahora acepta
`df_resumen_re545` y usa `startcol`).

**De dónde sale `AY` ("Oferta Completa"), que no es fórmula ni la escribe ninguna macro:** es la
columna `Completa` de `construir_resumen_ventana_oferta()` — la misma tabla central+ventana que ya
alimenta `Medidores!T` (`T = 1 - Completa`). Coinciden el nombre, la clave, el dominio (0/1) y el
sentido, y el archivo real lo confirma: la central cuya última ventana queda incompleta tiene
`AY = 0` justo ahí. No hubo que inventar nada ni pedir un archivo nuevo: `generar_pagos_bess()`
reconstruye ese resumen desde la hoja `Medidores` ya generada.

**Otro dato que se resolvió de paso:** `Medidores!$S$1` (que usa `BV`) es la **hora de inicio de
ventana**, el mismo dato que la constante `INICIO_VENTANA` — se deduce de la fórmula de
`Medidores!L`, que incrementa la ventana justo cuando la hora es igual a `S1`.

**Cambios en `nucleo.py`:** `NOMBRES_RESUMEN_RE545`, `calcular_bv_re545()`,
`construir_resumen_ventanas_re545()`. Además `completar_calculo_re545()` ya **no** renombra (el
renombre pasó a `renombrar_calculo_re545()`, que se llama al final): la tabla resumen necesita el
DataFrame con los nombres internos.

**Verificación:** tests sintéticos con grupos armados a propósito — `AZ`/`BA` tomando la PRIMERA
fila del grupo y no la última (el `AGGREGATE(15,6,...,1)` del original), `BB` como suma de `AU`,
`BF` sumando solo las filas de la hora anterior al inicio de ventana, `BC` calculado a mano
(`MIN(MAX(MIN(120,50),20),40) = 40`), el caso "oferta incompleta → `BC = 0`" y el caso "ventana 31
→ `BC = 0`". Más un test de escritura real del `.xlsx` confirmando que las dos tablas quedan lado
a lado con una columna en blanco entre medio. Regresión de E Costos y de las etapas 1-2 de RE545:
pasan.

**Pendiente:** `BI:CE` (Componentes 1 y 2) y, con eso, `BD`/`BE` del resumen.

---

## 2026-09-11 (11) — `Calculo RE545`, etapa 4: `BI:CE` — hoja completa

Última etapa de RE545 en la misma sesión. Con esto quedan **completas las dos hojas de cálculo**
del libro (`Calculo E Costos` y `Calculo RE545`).

**Corrección de un error propio de las etapas anteriores de esta misma sesión** (queda anotado
porque cambia valores ya commiteados): `VLOOKUP(G, Resumen!$B$8:$J$26, 4, 0)` **no es `Pmax
(MW)`**, es **`Capacidad (MWh)`**. El orden real de las 9 columnas de `Resumen BESS` es `Nombre
activo`, `Pmax (MW)`, `Horas para descarga forzada`, `Capacidad (MWh)`, `Energía mínima`, `Barra
inyección`, `% Energía sobre mínima`, `Ciclos max diarios`, `Eficiencia` — y es consistente con
que `H` use el índice 6 para `Barra inyección`. Afectaba a `U` (`EiniT`) y a `BC` (`Edisp_T`), que
se habían implementado con `Pmax`. Se agregó `construir_dic_resumen_capacidad()` y se corrigieron
las dos. Tabla de referencia de qué índice usa cada columna, en el plan §26.7. **`AE`/`AF` de
E Costos siguen bien con `Pmax`** (ahí el VBA usa `Resumen!B:C`, o sea el índice 2), igual que `BN`
de RE545.

**Implementado (`calcular_componentes_re545()` + `completar_checks_resumen_re545()`):** `BI`
(`Orden`, con el salto de 4 filas del original), `BJ` (`Periodo`), `BK`/`BL` (sumas por
central+orden+ventana+periodo), `BM` (`LARGE(IF(...))` matricial: el k-ésimo `CMg` más grande
entre todas las filas con el mismo `BK`, con `k = Periodo/15 + 1`), `BN` (`Edisp_Asig`), `BO`,
`BQ`, `BR`, `BS` (`INDEX/MATCH` matricial), `BT`, `BU`, `BV`, `BW`, `BX`, `BY`, `BZ`, `CA`, `CC` y
`CE` (`Monto a compensar`). Y con eso, `BD`/`BE` del resumen `AW:BG`.

**Las dos recursiones, que es lo único que no se puede vectorizar:** `BN` necesita los `BN`
anteriores de su grupo (se recorre de arriba hacia abajo) y `BU` necesita los `BU` **posteriores**
(se recorre de abajo hacia arriba, porque `BT` mira las filas siguientes). `BY` necesita los `BZ`
anteriores, pero `BZ` no depende de `BY`, así que ahí alcanza con calcular `BZ` primero.

**Verificación:** un caso sintético de 8 filas (2 horas × 4 bloques) con todos los valores
elegidos para poder calcular a mano: `BI` = 1,1,1,1,2,2,2,2; `BM` = 80,70,60,50 repetido; `BN`
cortándose al llegar al `Edisp_T` del grupo (150, 150, 100, 0, ...); `BT`/`BU` con la recursión
hacia arriba (400,400,400,400,300,200,100,0 y 0,0,0,0,100,100,100,100); `BY` como acumulado de
`BZ`; y `CE` en sus dos ramas — la de `MAX(...,0)` (da 0) y una positiva calculada a mano
(`(28.500 - 26.000)/8 = 312,5` por fila). Más el caso de inyección negativa (`BZ = 0`, `BY = BX`,
`BU = 0` por `BS = 0`) y los dos checks del resumen. Regresión de E Costos (etapas 2-4) y de las
etapas 1-3 de RE545: pasan.

**Lo que sigue** (nuevo frente, ninguna analizada todavía): las hojas de salida que consumen estas
dos — `PRORRATA_RETIROS`, `Compensacion total`, `Resumen` y el CSV.

---

## 2026-09-11 (12) — Casillas por hoja en la ventana "Generar" de `Pagos_BESS.xlsx`

El usuario pidió cerrar el pendiente que había quedado anotado ("ajustamos detalles después"):
la ventana "Generar" de `Pagos_BESS.xlsx` era todo o nada (una sola hoja, sin casillas), a
diferencia de la de `Consolidado_entradas.xlsx` (`SECCIONES_CONSOLIDADO`). Ahora que la hoja
tiene dos salidas (`Calculo E Costos` y `Calculo RE545`, completas desde la sesión anterior),
pedido explícito: **dos casillas, una por hoja**.

**Cambios en `nucleo.py`** (mismo patrón que `SECCIONES_CONSOLIDADO`/`generar_consolidado`):

- `SECCIONES_PAGOS`: tupla con `("ecostos", "Calculo E Costos", descripción, ("Calculo E
  Costos",))` y `("re545", "Calculo RE545", descripción, ("Calculo RE545",))`.
- `escribir_pagos_bess()` gana `ruta_existente` y `hojas_regenerar` (antes solo tenía
  `df_ecostos`/`df_re545`/`df_resumen_re545`/`registrar`). Con `hojas_regenerar=None` se
  comporta exactamente igual que antes (retrocompatible: los tests de sesiones anteriores que
  la llaman posicionalmente sin estos parámetros nuevos siguen funcionando tal cual). Con un
  `set`, la hoja que NO está en el set se copia tal cual desde `ruta_existente` en vez de
  escribirse desde el DataFrame — reusa `_copiar_hoja_existente()`, la misma función que ya
  usaba `escribir_salida()` para `Consolidado_entradas.xlsx`. Si no hay versión anterior para
  preservar, la hoja queda vacía y se registra un aviso (mismo criterio, sin duplicar código).
- `generar_pagos_bess()` gana el parámetro obligatorio `secciones_activas` (antes no lo tenía;
  es un cambio incompatible a propósito, como ya había pasado con `generar_consolidado()`).
  Valida secciones desconocidas y "ninguna tildada" antes de tocar ningún archivo. Las lecturas
  compartidas (`Medidores`, `Subastas`, `Centrales.xlsx`, `cmg.xlsx`) se hacen siempre que haga
  falta alguna sección; lo que se condiciona es el CÁLCULO de cada hoja:
  - `dic_eficiencia`/`dic_capacidad` (que solo usa RE545: `V`, `U`/`BC`/`BN`) solo se arman si
    `"re545"` está tildada.
  - El archivo `SSCC_Desempeño_*` (que solo usa `Calculo E Costos`, para `AM:AR`) solo se exige
    y se lee si `"ecostos"` está tildada. **Confirmado con un caso real armado a propósito**:
    tildar solo `"re545"` corre sin pedir ese archivo aunque no exista en la carpeta del caso;
    tildar solo `"ecostos"` sí lo exige y falla con un mensaje claro si falta.
  - `df_ecostos`/`df_re545`/`df_resumen_re545` quedan en `None` si su sección no está tildada, y
    así se le pasan a `escribir_pagos_bess()` junto con el `hojas_regenerar` correspondiente.

**Cambios en `Balance_BESS.py`:** `abrir_ventana_generar_pagos()` reescrita para recorrer
`nucleo.SECCIONES_PAGOS` igual que `abrir_ventana_generar_consolidado()` recorre
`SECCIONES_CONSOLIDADO` — un `LabelFrame` con casilla + descripción por sección, valida que haya
al menos una tildada, y pasa `secciones_activas` a `generar_pagos_bess()`.

**Verificación:** tests sintéticos de `escribir_pagos_bess()` (primera corrida escribe las dos
hojas; segunda corrida con solo `"ecostos"` recalcula esa hoja y preserva RE545 tal cual estaba,
sin vaciarlo; tercera corrida con solo `"re545"` al revés; caso sin versión anterior para
preservar → hoja vacía + aviso) y de `generar_pagos_bess()` (secciones vacías/desconocidas →
`ErrorEntrada`). Además un caso **con archivos reales** armado a propósito (`Centrales.xlsx`,
`cmg.xlsx`, `Consolidado_entradas.xlsx` con `Medidores`/`Subastas`) para confirmar en la práctica
que "solo RE545" no pide `SSCC_Desempeño_*` y que "solo E Costos" sí, y que una corrida que falla
a mitad de camino (por archivo faltante) no toca el `Pagos_BESS.xlsx` ya existente. Regresión
completa de las sesiones anteriores (Calculo E Costos etapas 2-4, Calculo RE545 etapas 1-4):
pasa. No se probó la ventana tkinter en sí (sin entorno gráfico en esta sesión, como siempre).

---

## 2026-09-11 (13) — Confirmado: `Configuración` (no `Propietario`) para homologar centrales

El usuario confirmó explícitamente: **"si se usa configuración"**. Cierra la última inferencia
sin confirmar de las que quedaban documentadas — la homologación de central en `calcular_l()`
(y, por extensión, en todo lo que reusa el mismo criterio: `AG:AL`/`AM:AR` de Prorratas/FD
homologado, `AW` con la tabla de umbrales, `calcular_reservas_re545()` de `Calculo RE545`, y
`calcular_subastas_energia_sscc()`) usa `Subastas!Configuración`, **no** `Subastas!Propietario`.
No se tocó código: la implementación ya usaba `Configuración` desde que se resolvió por
inferencia razonada (mismo nombre de campo que `Calculo E Costos!G`, alineación semántica con
Mes/Dia/Hora_dia); esta sesión solo pasa esa elección de "inferida, pendiente de confirmar" a
"confirmada por el usuario" en toda la documentación (`docs/Plan_Traspaso_Python_Balance_BESS.md`
§25.6, `METODOLOGIA.md` §7, `BITACORA.md`).

**Sigue pendiente** (no es lo mismo que la homologación en sí): validar fila por fila contra un
caso real que la cantidad de filas con `L=1` sea razonable — eso confirma que el CRUCE funciona
bien con datos reales, más allá de que ya esté confirmado qué columna usar.

---

## 2026-09-11 (14) — Fix: crash con datos reales cuando `Subastas` no tiene filas BAJADA/SUBIDA

**Primera corrida contra datos reales** (el usuario corrió `generar_pagos_bess()` desde la
ventana, con `SSCC_Desempeño_Julio_2026_V2.xlsx` real: 6.696 filas CSF/CPF leídas bien). Se
cayó en `completar_calculo_e_costos_grupos()` → `calcular_l()` → `_construir_set_subastas_tipo()`
con:

```
ufunc 'add' did not contain a loop with signature matching types
(dtype('int64'), dtype('<U1')) -> None
```

**Causa raíz:** el filtro por `Subastas!Sub_Baj` en `{BAJADA, SUBIDA}` no encontró **ninguna
fila** en el archivo real del usuario — la tabla filtrada (`sub`) quedó con 0 filas. Con 0 filas,
`pandas.Series.map()` es un no-op que no llega a ejecutar la función: devuelve una Series vacía
con el **mismo dtype que tenía antes** de mapear, en vez del dtype que devolvería la función.
Como `Subastas!Mes` se lee de Excel como `int64` (es una columna numérica), el resultado de
`sub["Mes"].map(_normaliza_valor_vba)` quedó en `int64` en vez de texto, y al concatenarlo con
`"¦"` (separador de clave) para armar la clave compuesta, `numpy` no encuentra una operación
`int64 + texto` y lanza el error. **No se reprodujo con los tests sintéticos de sesiones
anteriores** porque ninguno armó a propósito el caso "cero filas después de filtrar" — siempre
había al menos una fila BAJADA o SUBIDA en los datos de prueba.

**Fix en `nucleo.py`:** nueva función `_columna_clave_vba(serie)` = `serie.map(_normaliza_valor_
vba).astype(str)` — el `.astype(str)` fuerza el dtype a texto **siempre**, esté vacía la Series o
no, corrigiendo el caso que `.map()` no cubre. Reemplaza el patrón `columna.map(_normaliza_valor_
vba)` en los 4 lugares donde participa de una concatenación con "+": `_construir_set_subastas_
tipo()` (el que crasheaba), `calcular_l()`, `calcular_prorratas()` y `calcular_subastas_energia_
sscc()`. Los otros dos usos de `.map(_normaliza_valor_vba)` en el archivo (`calcular_reservas_
re545()`) no se tocaron: ahí el resultado se usa como clave de tupla en un diccionario, no se
concatena con "+", así que no está expuesto a este bug.

**Verificación:** se reprodujo el error exacto con un `DataFrame` armado a propósito (columnas
`object`/`int64` "clásicas", no el dtype `str` nuevo de pandas 3.x que usa este sandbox por
defecto — hubo que forzar los dtypes explícitamente para reproducirlo, porque con el dtype nuevo
la concatenación no fallaba) y una fila cuyo `Sub_Baj` no es ni `BAJADA` ni `SUBIDA` (0 filas tras
filtrar). Confirmado que antes del fix reproduce el `ufunc 'add'` exacto y que después corre sin
error, devolviendo `L=0` en todas las filas (comportamiento correcto: si de verdad no hay ninguna
subasta que cruce, nadie participa). Regresión completa de las 14 sesiones anteriores: pasa.

**Hallazgo de fondo, no un bug de código, queda como pendiente:** que el filtro haya dado 0 filas
significa que en el caso real usado, **ninguna fila de `Subastas!Sub_Baj` normaliza exactamente a
`BAJADA` o `SUBIDA`**. Puede ser el período correcto (sin subastas ese mes) o puede ser que el
texto real use otra palabra/formato — hay que confirmarlo con el usuario (ver "Pendientes
abiertos"). Mientras tanto, con ese caso, `L`, `M`, `N`, `O` y las reservas por subasta de
`Calculo RE545` dan 0/vacío en todas las filas — no es un error, es el resultado correcto de la
fórmula real con ese filtro.

---

## 2026-09-11 (15) — Fix: SoC con nombres de central como ruta SCADA + hallazgo en `Subastas`

Segunda corrida contra datos reales del usuario, dos problemas nuevos.

### Fix: SoC — nombres de bloque como ruta SCADA completa

El log mostró: `Centrales en Medidas_SAE.xlsx sin bloque de SoC: [9 centrales limpias]` y
`Centrales en el SOC que no estan en Medidas_SAE.xlsx: [rutas tipo
'\\SRV-SCADA-AF1\SEN\Generación\SEN\03 Región II\SAE-Del Desierto|Nombre PCP/PID']` — cero
cruces. Causa: en el archivo real de SoC (exportación tipo PI), el nombre de cada bloque de
central **no es el nombre limpio**: es la ruta SCADA completa del punto, con un sufijo `|Nombre
PCP/PID` (o `|Nombre` en al menos un caso) pegado al final. `detectar_bloques()` toma ese texto
tal cual (`nombre_bess_origen = str(valor).strip()`, es lo correcto — no reinterpreta nada), pero
`extraer_soc()` buscaba ese texto LITERAL en el `Diccionario`, y ningún humano escribe esa ruta
completa a mano en una hoja de equivalencias — por eso el cruce daba siempre 0.

**Fix:** `_extraer_nombre_desde_ruta_scada(texto)` — si el texto tiene `\`, devuelve solo el
último tramo (después de la última `\`) sin el sufijo después de `|` (ej. de la ruta de arriba
saca `"SAE-Del Desierto"`). No es una reinterpretación de datos: es separar una estructura ya
presente en el archivo (ruta + sufijo), no adivinar a qué central corresponde. `extraer_soc()`
ahora prueba primero el texto literal (compatibilidad con cualquier SoC "limpio" sin ruta) y, si
no hay match, prueba de nuevo con el nombre extraído — el `Diccionario` puede tener cualquiera de
las dos formas. Si ninguna tiene match, usa el nombre **limpio** (no la ruta completa) como
`central`, solo para que avisos/incidencias sean legibles — sigue sin cruzar contra `Medidores`,
mismo comportamiento que antes.

**Importante, se lo dejo dicho al usuario en el chat:** este fix por sí solo **no alcanza** para
la mayoría de las 9 centrales. Comparando el nombre que queda tras extraerlo de la ruta contra el
nombre limpio real de `Medidas_SAE.xlsx`:

| Extraído de la ruta SoC | Real en Medidas_SAE.xlsx | ¿Cruza solo con el fix? |
|---|---|---|
| `SAE-Tocopilla` | `SAE-TOCOPILLA` | Sí (`normalizar()` ya ignora mayúsculas) — pero el merge final es por texto LITERAL, así que de todas formas necesita un `Diccionario` que devuelva el texto exacto `SAE-TOCOPILLA` |
| `SAE-Del Desierto` | `SAE-DEL-DESIERTO` | No (espacio vs guion) |
| `SAE-PE La Cabaña` | `SAE-CRCA-PE-LA-CABANA` | No (falta prefijo `CRCA-`) |
| `SAE-PFV Victor Jara` | `SAE-CRCA-PFV-VICTOR-JARA` | No (falta prefijo) |
| `SAE-PFV Andes Solar 4` | `SAE-CRCA-PFV-ANDES4` | No |
| `SAE-PFV Andes Solar III` | `SAE-CRCA-PFV-ANDES3` | No (`III` vs `3`) |
| `SAE-PFV Nuevo Quillagua II` | `SAE-CRCA-PFV-NUEVO-QUILLAGUA-2` | No (`II` vs `2`) |
| `SAE-PFV Don Humberto` | `SAE-CRCA-PFV-DON-HUMBERTO` | No (falta prefijo) |

Las 9 (todas, en la práctica) necesitan una fila en `Centrales.xlsx!Diccionario` con el nombre
limpio de `Medidas_SAE` y el nombre extraído de la ruta (o la ruta completa, cualquiera de las
dos funciona ahora) — eso es contenido del archivo del usuario, no algo que el código deba
adivinar (`Diccionario` es mantenimiento manual, ver `METODOLOGIA.md`).

**Caso sin resolver, no es una ruta SCADA:** `'07 Region RM'` apareció como un bloque completo
aparte, sin `\` ni `|`. No encaja con el patrón de los otros 8 — no se adivinó qué es (podría ser
un bloque real mal cortado, o una columna/central que no corresponde). Se le preguntó al usuario.

### Hallazgo (sin tocar código): posible columna faltante en `Subastas`

El usuario compartió una foto de la hoja `Subastas` real de la planilla 11. Una fila de datos
muestra, en celdas consecutivas: `CSF(-)` (que es exactamente el resultado de la fórmula real de
`Control`, sección 5.3 del documento de trazabilidad: `=IF(AND(C1="CSF",D1="SUBIDA"),"CSF(+)",
IF(AND(C1="CSF",D1="BAJADA"),"CSF(-)",...))`) seguido de `CSF` y de `BAJADA`. Esa fórmula necesita
DOS insumos (`C`=servicio SIN dirección, `D`=dirección) para armar el label de `Control` — pero
`NOMBRES_SUBASTAS` hoy solo tiene UNA columna entre `Control` y `Fecha` (`Sub_Baj`, que se asumía
que guardaba `SUBIDA`/`BAJADA` directamente, confirmado hace sesiones por el usuario: "Es la
columna C de la hoja subastas que ya generamos"). Si la foto es correcta, falta una columna
"Servicio" que nunca se mapeó, y **todo lo que sigue después de `Control` queda corrido una
posición**. Esto también explicaría el otro hallazgo de la sesión anterior (`Sub_Baj` dando 0
filas BAJADA/SUBIDA con datos reales — si el filtro estaba comparando contra la columna
equivocada, por supuesto no cruza nada).

**No se tocó código todavía.** Este archivo ya tuvo varias rondas de "corrimiento de columna" mal
resueltas por inferencia (ver sesiones anteriores); antes de tocar `NOMBRES_SUBASTAS` de nuevo
hace falta confirmación letra por letra del usuario, no otra inferencia visual sobre una captura
de pantalla. Se le pidió que confirme el contenido exacto de las primeras columnas de `Subastas`
en la planilla 11 real.

**Verificación:** tests sintéticos de `_extraer_nombre_desde_ruta_scada()` (ruta con los dos
sufijos vistos, texto sin ruta queda intacto) y de `extraer_soc()` de punta a punta (con y sin
`Diccionario`, y confirmando que `nombre_scada_original` sigue guardando la ruta completa sin
tocar, solo `central` cambia). Regresión completa: pasa.

---

## 2026-09-11 (16) — Corrección grande: `NOMBRES_SUBASTAS` estaba mal desde el principio

El usuario adjuntó `Libro1.xlsx` con la hoja `subastas` real (encabezados **y fórmulas**, no solo
nombres) y pidió: "revisa la macro `Cargar_Remuneracion_Subastas_Rapido` y arma según el adjunto
en el mismo formato de salida que tengo en la planilla 11". Se copió el archivo al repo como
`docs/Libro1_Subastas_real.xlsx` (mismo criterio que los otros adjuntos que son fuente de una
decisión). Cruzando ese archivo contra el código VBA de la macro (ya en el documento de
trazabilidad) se encontró la causa raíz de la confusión que venía arrastrándose desde hace varias
sesiones (el "corrimiento de columna" que nunca terminaba de cuadrar del todo).

**La prueba definitiva** es la fórmula real de `Subastas!B1`:

```
=IF(AND(C1="CSF",D1="SUBIDA"),"CSF(+)",IF(AND(C1="CSF",D1="BAJADA"),"CSF(-)",
  IF(AND(C1="CTF",D1="SUBIDA"),"CTF(+)",IF(AND(C1="CTF",D1="BAJADA"),"CTF(-)","REVISAR"))))
```

Esta fórmula arma la columna `B` a partir de DOS insumos: `C` (el tipo de servicio SIN dirección:
`CSF`/`CTF`/`CPF`) y `D` (la dirección: `SUBIDA`/`BAJADA`). `NOMBRES_SUBASTAS` **solo tenía una
columna ahí** (`B`="Control", asumiendo que guardaba el label completo como `"CSF(-)"`, y
`C`="Sub_Baj" asumiendo que ahí vivía directamente `SUBIDA`/`BAJADA`) — le faltaba contar una
columna entera: la real `B` = `Concepto` (el label completo, lo que veníamos llamando "Control"),
`C` = `Control` (el tipo SIN dirección, una columna que nunca se había mapeado) y recién `D` =
`Sub_Baj`. Confirmado letra por letra contra los datos reales del archivo (fila 3: `B3='CSF(-)'`,
`C3='CSF'`, `D3='BAJADA'`, `K3='SAE-CRCA-PFV-DON-HUMBERTO'`, `L3='EGP_CHILE'`).

**`NOMBRES_SUBASTAS` corregido, de punta a punta** (16 posiciones `B:Q`, todas confirmadas contra
el archivo real):

| Letra | Antes (mal) | Ahora (confirmado) |
|---|---|---|
| B | `Control` | **`Concepto`** |
| C | `Sub_Baj` | **`Control`** |
| D | `Fecha` | **`Sub_Baj`** |
| E | `Año` | **`Fecha`** |
| F | `Mes` | **`Año`** |
| G | `Dia` | **`Mes`** |
| H | `Hora_dia` | **`Dia`** |
| I | `Hora_mes` | **`Hora_dia`** |
| J | `Configuración` | **`Hora_mes`** |
| K | `Propietario` | **`Configuración`** |
| L | `Clave horaria` | **`Propietario`** |
| M | `Ciclo` | **`Clave horaria`** |
| N | `Energía SSCC` | **`Ciclo`** |
| O | `FD` | **`Energía SSCC`** |
| P | `FMA` | **`FD`** |
| Q | *(sin nombre)* | **`FMA`** |

Cada nombre "de antes" cayó exactamente una posición más adelante de donde debía — **no** era un
corrimiento uniforme de todo el archivo (la explicación que se venía dando, "la primera columna
del original está vacía"), era que faltaba UNA columna real (`Concepto`) al principio del bloque
de 11 que copia la macro (`DB!B:L → Subastas!B:L`, confirmado con el `MsgBox` de la macro:
`"DB B:L → Subastas B:L"`, `"DB P → Subastas O"`, `"DB Y → Subastas P"`, `"DB V → Subastas Q"`).
`A` (antes de `B`=Concepto) sigue genuinamente sin usar — eso sí estaba bien.

**Consecuencia feliz: casi nada de la lógica ya escrita estaba mal, solo los nombres.**
`construir_subastas()` arma los valores por **posición** (columna 0 del bloque B:L, columna 1,
etc.), y esas posiciones eran correctas — el bug estaba únicamente en qué nombre se le pegaba a
cada posición al final (`NOMBRES_SUBASTAS`). Por ejemplo, la fórmula de "Clave horaria"
(`=Configuración&Dia&Hora_dia`, confirmada con la fórmula real `M3=K3&H3&I3`) ya se calculaba
así en el código (usando las letras internas `K`,`H`,`I` como *posiciones*, que por construcción
coinciden con las letras reales de Excel en el bloque B:L) — solo estaba mal etiquetada como
`"Ciclo"`. Verificado con los valores exactos de dos filas del archivo real (test sintético, ver
abajo): el resultado coincide.

**Cambios de código:**

- `NOMBRES_SUBASTAS`: corregido (tabla de arriba).
- `calcular_subastas_energia_sscc()` → renombrada **`calcular_subastas_ciclo()`**: la columna
  `N` que esta función resuelve **no es "Energía SSCC", es "Ciclo"** — el nombre real de la
  columna que trae el XLOOKUP (`Calculo E Costos!P` = "Ciclo de Carga del mes") coincide
  exactamente con que sea un número de ciclo, no una energía. La lógica/fórmula que ya estaba
  implementada es la correcta — solo el nombre estaba equivocado. `construir_dic_umbrales_
  subastas()` renombra su parámetro `energia_sscc` → `ciclo_subastas` (mismo motivo).
- `construir_dic_reservas_subastas()` (reservas por subasta de `Calculo RE545`, `AC:AU`): usaba
  `df_subastas["Control"]` como criterio de tipo (`CPF(-)`/`CSF(+)`/etc) — **esto sí era un bug
  real, no solo un nombre**: la fórmula real usa `Subastas!$B:$B` como rango de coincidencia
  contra el encabezado de columna (`AC$3="CPF(-)"`), y `Subastas!B` es `Concepto` (el label
  completo), no `Control` (que ahora sabemos que es solo el tipo sin dirección — nunca iba a
  poder distinguir `(-)` de `(+)`). Corregido a `df_subastas["Concepto"]`.
- `construir_prorrata_sscc()`/`construir_dic_prorrata()` (Prorrata SSCC, `AG:AL`): **sin
  cambios** — el pivot ya usaba `columns="Control"`, y ahora que se sabe que `Control` real es
  el tipo sin dirección (`CSF`/`CTF`/`CPF`), sigue siendo exactamente lo que el pivot necesita
  (agrupa por tipo, no por tipo+dirección — coincide con que el VBA original duplique
  literalmente `AJ=AG`/`AK=AH`, sin distinguir dirección en absoluto). Era una decisión correcta
  desde el principio, ahora con más fundamento.

**Pieza extra que este archivo confirmó, sin que hiciera falta pedirla:** la tabla de umbrales
SUBIDA/BAJADA (`AW`, bloqueada durante varias sesiones hasta la "etapa 4") vive en
`Subastas!S:W` (`S`=Configuración, `T`=Ciclo, `U`=Clave, `V`=SUBIDA, `W`=BAJADA) — coincide
exactamente con lo que `construir_dic_umbrales_subastas()` ya calculaba en Python de forma
independiente (no la lee del archivo, la deriva). Buena señal cruzada de que esa parte del
cálculo está bien encaminada.

**Verificación:** nuevo test (`test_subastas_real.py`, no persistido) que arma dos filas de
`DB` con los valores EXACTOS de las filas 3 y 7 del `Libro1.xlsx` real (mismo texto, mismos
números) y confirma que `construir_subastas()` separa correctamente `Concepto`/`Control`/
`Sub_Baj`, ubica `Configuración`/`Propietario` en su lugar, calcula `Clave horaria` igual que la
fórmula real, deja `Ciclo` vacío (se calcula después, en `Pagos_BESS.xlsx`) y copia `Energía
SSCC`/`FD`/`FMA` en las posiciones correctas (`40.6`/`0.8533`/`0` y `74.5`/`1`/`0`, tal cual el
archivo). Se corrigieron los tests de sesiones anteriores que usaban los nombres viejos
(`test_etapa4.py`, `test_re545.py`) para que seteen `Concepto` además de `Control` donde
corresponde. Regresión completa (E Costos etapas 2-4, RE545 etapas 1-4, secciones de
`Pagos_BESS.xlsx`, SoC): pasa.

**No se tocó** (fuera de alcance de esta corrección, quedan igual): `FD`/`E COSTOS`/`Resumen` del
mismo `Libro1.xlsx` — se revisaron de pasada y coinciden con lo ya implementado (`NOMBRES_FD_CSF`,
`NOMBRES_FD_CPF`, `NOMBRES_CALCULO_E_COSTOS`), sin cambios.

**Todavía pendiente** (ver "Pendientes abiertos"): correr un caso real completo con esta
corrección aplicada, para confirmar que ahora sí aparecen filas `L=1` y que Prorratas/reservas de
RE545 no quedan en 0 (la falta de la columna `Concepto` explicaría, retroactivamente, por qué el
caso real de la sesión anterior daba 0 filas `BAJADA`/`SUBIDA`).

---

## 2026-09-11 (17) — Fix: fila de nombres del SoC (fila 2, no la 3) + selector Medidores/Ofertas separado

Dos pedidos cortos del usuario, con un archivo real (`SOC_2607.xlsx`) que resolvió el primero de
punta a punta.

### Fix: `detectar_fila_nombres()` se quedaba con la fila equivocada

El usuario avisó: "el soc sigue sin nada en la fila 2 está el nombre de la central porsi no la
3". Con el archivo real se confirmó la estructura exacta: fila 2 = nombre limpio de la central
(`"SAE-CRCA-PFV-DON-HUMBERTO"`, idéntico al de `Medidas_SAE.xlsx`), fila 3 = la ruta SCADA
completa (`"\\SRV-SCADA-AF2\SEN\Generación\...\SAE-PFV Don Humberto|Nombre"`, la misma pieza que
ya se había resuelto la sesión anterior con `_extraer_nombre_desde_ruta_scada`), fila 4 en blanco,
fila 5 los encabezados `Status/Questionable/Time Stamp/Value`. Dos filas útiles APILADAS (sin
blanco entre medio) antes del hueco en blanco que precede a los encabezados — `detectar_fila_
nombres()` subía desde los encabezados y se quedaba con la PRIMERA fila útil que encontraba
(fila 3, la ruta), sin darse cuenta de que había otra más arriba (fila 2, el nombre limpio).

**Fix:** en vez de devolver la primera fila útil encontrada subiendo, ahora se sigue subiendo
mientras las filas sigan siendo útiles (sin blanco de por medio) y se devuelve la MÁS ARRIBA de
ese bloque contiguo. Con una sola fila útil (el caso más común hasta ahora) el comportamiento es
idéntico a antes — retrocompatible.

**Verificación, con el archivo real completo:** `extraer_soc()` sobre `SOC_2607.xlsx` (sin ningún
`Diccionario`/homologación, `mapa_homologacion={}`) detecta las **9 centrales exactas** de
`Medidas_SAE.xlsx` (`SAE-CRCA-PE-LA-CABANA`, `SAE-CRCA-PFV-ANDES3`, ..., `SAE-TOCOPILLA`),
**26.793 filas**, **cero incidencias**. La misteriosa central suelta `'07 Region RM'` de la sesión
anterior queda explicada: es lo que la fila 3 (ruta SCADA) tiene para el bloque de `SAE-CRCA-
PFV-MANZANO` en particular — un dato incompleto que ya no se usa, porque ahora se lee la fila 2.
Con esto, y salvo que otro archivo real muestre lo contrario, probablemente **ya no hace falta
ningún `Diccionario` para homologar el SoC** — los nombres de fila 2 ya vienen idénticos a
`Medidas_SAE.xlsx`. Se agregó `test_soc_fila_nombres.py` (no persistido) con el caso real (2 filas
apiladas), el caso simple (1 fila, retrocompatibilidad) y un caso con 3 filas apiladas.

### Selector separado: `Medidores` y `Ofertas SSCC` como casillas independientes

Pedido: "separa medidas de ofertas el selector". Hasta ahora `SECCIONES_CONSOLIDADO` tenía una
sola sección `"medidores"` que escribía las dos hojas juntas (`("Medidores", "Ofertas SSCC")`),
porque `construir_medidores()` las arma en una sola pasada (`Medidores!R:S:T` depende de Ofertas
SSCC). Ahora son dos ids separados (`"medidores"` → hoja `Medidores`; `"ofertas_sscc"` → hoja
`Ofertas SSCC`), cada uno decidiendo solo si se REESCRIBE su propia hoja — pero **la lectura
combinada sigue siendo una sola**: alcanza con que cualquiera de las dos esté tildada para que se
lean los 4 archivos de entrada (`Medidas_SAE.xlsx`, SoC, `Centrales.xlsx`, OfertasSSCC) y se corra
`construir_medidores()`; lo que cambia es solo qué hoja(s) se escriben al final (`hojas_
regenerar`, mismo mecanismo de preservación que ya usaban `CMg`/`FD`/`Subastas`). No hacía falta
tocar `Balance_BESS.py`: la ventana ya recorre `SECCIONES_CONSOLIDADO` genéricamente, así que
ahora dibuja 5 casillas en vez de 4 sin ningún cambio de código — solo se agrandó la ventana
(`620x420` → `620x560`) para que entren.

**Verificación:** `test_selector_medidas_ofertas.py` (no persistido), con las funciones de
lectura/cálculo monkeypatcheadas para no depender de archivos Excel completos: confirma que
tildar solo `"ofertas_sscc"` dispara igual la lectura combinada (se llama a `construir_
medidores()`) pero `Medidores` queda preservado tal cual estaba, y viceversa con solo
`"medidores"` tildada. Regresión completa de las 16 sesiones anteriores: pasa. No se probó la
ventana tkinter en sí (sin entorno gráfico en esta sesión, como siempre).

---

## 2026-09-11 (18) — Diagnóstico mejorado: "sin bloque de SoC" ahora dice si es un problema de Diccionario

El usuario corrió de nuevo con el fix de la sesión anterior. Mejoró mucho (de 9 centrales sin
cruzar a 1): `Centrales en Medidas_SAE.xlsx sin bloque de SoC: ['SAE-CRCA-PFV-NUEVO-QUILLAGUA-2']`
— pero avisó que esa central sí está en su archivo de SoC.

**Diagnóstico:** con el `SOC_2607.xlsx` real que había compartido, `extraer_soc()` (aislado, sin
`Diccionario`) SÍ detecta esa central perfectamente — fila 2 trae literalmente
`'SAE-CRCA-PFV-NUEVO-QUILLAGUA-2'`, idéntico a `Medidas_SAE.xlsx`, sin ningún carácter raro (se
revisó con `repr()`, sin espacios/unicode ocultos). Sin `Diccionario`, esta central cruza sola.

**Hipótesis más probable, confirmada como técnicamente posible con una prueba:** el
`Centrales.xlsx!Diccionario` del usuario probablemente tiene una fila para esta central con el
nombre "feo" (estilo ruta SCADA, ej. `"SAE-PFV Nuevo Quillagua II"`) escrito ANTES que el nombre
limpio en esa misma fila. `construir_homologacion()` toma el PRIMER valor de cada fila como
"canónico" — si ese orden quedó así (probablemente porque el `Diccionario` se armó en una época
en que solo se conocía el nombre feo, antes de esta sesión), la homologación **rompe** un cruce
que la fila 2 del SoC ya resolvía sola: convierte el nombre limpio en el feo, y el feo no cruza
contra nada en `Medidas_SAE.xlsx`. Se armó una prueba sintética que reproduce exactamente este
mecanismo (`construir_homologacion()` con una fila `["SAE-PFV Nuevo Quillagua II",
"SAE-CRCA-PFV-NUEVO-QUILLAGUA-2"]` → el nombre limpio homologa hacia el feo).

**No se tocó la lógica de homologación** (cambiar cuál valor de la fila gana como "canónico"
afectaría potencialmente otras centrales que sí dependen del orden actual — cambio de más riesgo
del que amerita una hipótesis todavía sin confirmar con el archivo real del usuario). En cambio,
se mejoró el **diagnóstico**: el aviso "Centrales en Medidas_SAE.xlsx sin bloque de SoC" ahora
busca, para cada central faltante, si existe algún `nombre_scada_original` (el nombre crudo antes
de homologar) que normalice igual a esa central — si lo encuentra, lo dice explícitamente en el
aviso ("el SoC SÍ trae un bloque con nombre crudo [...] — revisar si Diccionario lo está
homologando a otro nombre"). Si el usuario corre de nuevo, el mensaje mismo va a confirmar o
descartar la hipótesis sin necesitar que comparta su `Diccionario`.

**Verificación:** test sintético (`test_soc_diagnostico.py`, no persistido) con la lógica exacta
del bloque nuevo: caso "hay candidato crudo" (muestra la pista) y caso "de verdad no hay bloque"
(mensaje simple, sin inventar pistas falsas). Regresión completa: pasa.

**Pendiente:** confirmar con el próximo aviso (o con el contenido real de `Diccionario`) si la
hipótesis es correcta. Si lo es, la corrección más simple sería reordenar esa fila del
`Diccionario` (poner el nombre limpio primero) — eso ya lo puede hacer el usuario directamente en
su archivo, sin esperar un cambio de código.

---

## 2026-09-11 (19) — Dos bugs reales confirmados y arreglados con la primera comparación fila a fila contra la planilla 11

El usuario compartió, por primera vez, una comparación directa: `Pagos_BESS.xlsx` generado por
Python con una hoja extra pegada a mano (`Ecostos planilla 11`) con los valores reales del
`.xlsm` original para las mismas filas, más su `Centrales.xlsx` real (dos veces) y confirmó que
el aviso mejorado de la sesión anterior efectivamente detectó la pista ("el SoC SÍ trae un
bloque..."). Con estos tres archivos se resolvieron dos bugs reales de una — la primera vez que
el proyecto se valida contra datos reales, no solo sintéticos, y encontró exactamente el tipo de
error que esa validación existe para atrapar.

### Fix 1: `construir_homologacion()` mezclaba tablas de equivalencia distintas

Se leyó el `Diccionario` real completo. Confirmó algo que el plan ya sabía en teoría (sección
4.2/B: "la hoja presenta bloques asociados a FD/Subastas/ofertas") pero que `construir_
homologacion()` nunca implementó correctamente: son **tablas independientes por columnas**
(`A:B`=FD, columnas `C:D` vacías, `E:F:G`=Subastas/ofertas), no una fila = todos los sinónimos de
una central. Para 6 de las 9 centrales las dos tablas coinciden fila a fila por casualidad del
orden en que se cargaron — pero para las últimas 3 (`VICTOR-JARA`, `ANDES4`, `NUEVO-QUILLAGUA-2`)
el orden de la tabla de la derecha está corrido una fila respecto de la izquierda. La función
vieja trataba la fila entera como un solo grupo de sinónimos, así que en esas 3 filas mezclaba
central de una tabla con la central de la fila vecina de la otra tabla — literalmente homologaba
`NUEVO-QUILLAGUA-2` hacia `ANDES4` (la central de la fila anterior en el bloque de la derecha).
Esto explica el aviso persistente que el usuario venía reportando (y que el diagnóstico de la
sesión anterior confirmó correctamente: "el SoC SÍ trae el bloque").

**Pista clave para encontrarlo:** las otras dos funciones que leen esta misma hoja
(`construir_dic_mapeo_diccionario()`, columnas `A:B`; `_mapas_homologacion_fge()`, columnas
`E:F:G`) **ya** usaban posiciones de columna fijas — nunca tuvieron este bug, porque ya estaban
diseñadas sabiendo que son tablas separadas. Solo `construir_homologacion()` (la más vieja de
las tres, escrita antes de que se entendiera bien la estructura de bloques) se había quedado con
el enfoque ingenuo de "toda la fila es un grupo".

**Fix:** `_bloques_columnas_diccionario()` — detecta los bloques de columnas automáticamente
(separador = una columna vacía en TODAS las filas del archivo, no alcanza con mirar una sola
fila porque la fila de encabezados de grupo típicamente solo tiene texto en la primera columna
de cada bloque). `construir_homologacion()` ahora arma el mapa bloque por bloque, sin mezclar
equivalencias entre bloques distintos.

**Verificación:** con el `Diccionario` real completo, las 9 centrales homologan correctamente
hacia sí mismas (antes, `NUEVO-QUILLAGUA-2` homologaba mal). Corrida de punta a punta con el
`SOC_2607.xlsx` real + el `Diccionario` real: 9 centrales, 26.793 filas, **cero** incidencias.
Test sintético (`test_diccionario_bloques.py`, no persistido) que reproduce exactamente la
estructura real (incluida la fila corrida) y confirma retrocompatibilidad con un `Diccionario`
de una sola tabla (sin bloques separados por columnas vacías).

### Fix 2: la Prorrata SSCC (`AG:AL`) devolvía la cuenta cruda, no la proporción

El usuario reportó: "hay algunos pocos casos en los que los controles de frecuencia salen con 2
y otras diferencias donde en la planilla 11 es 0.5 y en la que genera el python es 1 [...] el
problema está en la prorrata de CF". Con la hoja de comparación se encontró la causa exacta:
para un grupo con 1 fila `CPF` + 1 fila `CSF`, la planilla real trae `AG=0.5, AH=0.5` — pero
nuestro código, que hacía `pivot_table(..., aggfunc="count")` y devolvía la cuenta tal cual, daba
`AG=1, AH=1` (la cuenta cruda de cada tipo). Para un grupo con 2 `CPF` + 1 `CSF`, la planilla
real trae `AG=0.6666..., AH=0.3333...` — exactamente `2/3` y `1/3`. El nombre "Prorrata" lo decía
literalmente: es una **proporción** (reparte el 100% del grupo entre los tipos que aparecen), no
una cuenta. El "sale con 2" que reportó el usuario es el mismo bug: cuando había 2 filas del
mismo tipo sin ningún otro tipo compitiendo, la cuenta cruda daba 2 en vez de la proporción
correcta (1.0, ya que 2/2=1).

**Fix:** `construir_prorrata_sscc()` ahora divide cada fila del pivot por la suma de esa misma
fila (entre todos los valores de `Control` presentes), antes de devolverlo — convirtiendo la
cuenta cruda en una proporción que suma exactamente 1 por fila.

**Verificación:** los DOS casos exactos que trajo la comparación real (`1+1 → 0.5/0.5` y
`2+1 → 0.6666../0.3333..`) coinciden ahora al dígito. Test sintético (`test_prorrata_
normalizada.py`, no persistido) con esos dos casos más un caso de una sola fila (retrocompatible,
sigue dando 1.0) y una prueba general de que cada fila del pivot siempre suma 1. Regresión
completa de las 18 sesiones anteriores: pasa (ningún test viejo dependía de un grupo con más de
una fila del mismo tipo, así que ninguno se rompió con la normalización).

**Downstream, sin cambios de código:** `AS`/`AT` (`_calcular_costo_ponderado`, "Energía descarga/
carga con FD") y todo lo que depende de `AG:AL` ya eran fieles a la fórmula real
(`factor = precio1*cantidad1 + precio2*cantidad2`, una suma ponderada, no un promedio) — el bug
estaba enteramente aguas arriba, en qué valores de `cantidad1`/`cantidad2` recibían. Con `AG:AL`
corregidos, `AS`/`AT` deberían salir bien ahora sin tocar esa función.

**Archivos de referencia guardados** (mismo criterio que sesiones anteriores — si es la fuente de
una decisión, vive en el repo): `docs/Centrales_real.xlsx`, `docs/SOC_real_2607.xlsx`,
`docs/Pagos_BESS_comparacion_real.xlsx` (esta última es la primera comparación real vs Python
lado a lado que existe en el proyecto — vale la pena mirarla de nuevo si aparece otra
discrepancia).

**Pendiente:** el usuario todavía no confirmó si con estos dos fixes el resto de la hoja coincide
completamente contra la planilla 11 — falta una corrida nueva de punta a punta.

## 2026-09-11 (20) — Fix: `Calculo RE545!AR:AT` no eran `SUMIFS`, eran la constante 1

El usuario reportó, con una nueva comparación real (`Pagos_BESS.xlsx` con una hoja pegada
`RE545 P11`, análoga a la de la sesión anterior pero para `Calculo RE545`): "En el R545 tengo
diferencias igual parten en AK:AN".

**Primer paso, para no perseguir un fantasma:** como ya se documentó (plan §26.2, BITACORA de
sesiones anteriores), la salida de `Calculo RE545` **no reproduce la letra de Excel real** — las
columnas vacías del original (`W:AB`) no se escriben, así que todo lo que sigue queda corrido de
letra. Comparando por **orden/contenido** (no por letra) los tres bloques de 6 columnas
(`Subastas`, `FD`, `FMA`) contra la hoja `RE545 P11` pegada por el usuario, fila a fila por las
26.784 filas, usando como clave `Mes+Dia+Hora+Minuto+Configuracion`:

- Bloque `Subastas`: 0 diferencias.
- Bloque `FD` (donde el usuario ubicó "AK:AN" en SU comparación, que sí tiene las columnas reales
  sin correr): 0 diferencias.
- Bloque `FMA`: **80.352 diferencias** (positions `CPF(+)`, `CSF(+)`, `CTF(+)` de ese bloque, en
  TODAS las filas).

O sea: la molestia real no estaba en `AK:AN` de nuestra salida (que es el bloque `FD`, sin
diferencias), sino en el bloque `FMA` (que en nuestra salida compacta cae en columnas distintas,
pero el usuario lo estaba mirando alineado con la posición real del `.xlsm`, donde si cae en
`AK:AN`... da igual: la comparación por contenido mostró exactamente dónde estaba el problema real,
sin necesidad de discutir letras).

**Causa exacta**, confirmada contra las fórmulas guardadas del archivo real
(`docs/Calculo_RE545_reducido_para_IA.xlsx`, hoja `Mapa_Formulas`, que lista la fórmula de cada
rango de columnas del `.xlsm` real): `AO:AQ` (`CPF(-)`/`CSF(-)`/`CTF(-)` del bloque `FMA`) SÍ son
`SUMIFS(Subastas!$Q:$Q, ...)`, igual que los otros 15 valores de los tres bloques — pero `AR:AT`
(`CPF(+)`/`CSF(+)`/`CTF(+)`, el mismo bloque) **no aparecen como fórmula en el mapa**: son el
valor literal `1`, constante, en las 26.784 filas del archivo real (confirmado también
directamente en la hoja `RE545 P11` pegada por el usuario: `CPF(+)=CSF(+)=CTF(+)=1` sin ninguna
excepción). El plan (§26.3) asumía "los tres [bloques] son el mismo SUMIFS" — cierto para 15 de
las 18 columnas, falso para estas 3. Nuestro código, al no distinguir el caso, les aplicaba el
mismo `SUMIFS`, que casi siempre da 0 (rara vez hay match exacto central+hora+tipo en `Subastas`
para esas combinaciones) — de ahí el "sale con 2 [en realidad 0] donde debería ser otra cosa" que
reportaba el usuario, y el arrastre a `AU` (`SUMA Reservas*FMA*FD = SUMPRODUCT(...)/4*1000`, que
multiplica por estas columnas).

**Fix:** `calcular_reservas_re545()` — al armar el tercer bloque (`FMA`), las posiciones 3, 4 y 5
(`CPF(+)`, `CSF(+)`, `CTF(+)`) ya no consultan el diccionario de `SUMIFS`: se fijan directamente
en `1.0`, para todas las filas, sin excepción — es dato constante, no una fórmula que dependa de
`Subastas`.

**Verificación:**
- La función nueva es, por construcción, independiente de los diccionarios de `SUMIFS` para esas
  3 columnas (siempre devuelve 1.0 sin mirar `Subastas`) — coincide automáticamente con las
  26.784 filas reales, que también son constantes.
- Recalculando `AU` fila a fila con los valores reales de los otros 15 componentes (que ya
  coincidían) más `AR=AS=AT=1` en vez del `SUMIFS` viejo: **0 diferencias** contra el `AU` real en
  las 26.784 filas (antes del fix: 4.628 filas con diferencia, hasta 36.609 de magnitud).
- Test sintético actualizado/agregado (`test_re545.py`, `test_re545_fma_constante.py`, no
  persistidos) que confirman que `AR`/`AS`/`AT` dan 1.0 sin importar el contenido de `Subastas`
  (incluso con diccionarios vacíos), y que las otras 15 columnas del bloque de reservas siguen
  siendo el `SUMIFS` de siempre, sin tocarse.
- Regresión completa de las 19 sesiones anteriores: pasa (se actualizó a mano el valor esperado
  de un test viejo, `test_re545.py`, que tenía hardcodeado el resultado incorrecto de `AU` para
  un caso con `CSF(+)`; el valor nuevo es el correcto según la fórmula real).

**Archivo de referencia:** `015231ed-Pagos_BESS.xlsx` (hoja `RE545 P11`, pegada por el usuario) —
no se copió a `docs/` porque no aporta nada que `docs/Pagos_BESS_comparacion_real.xlsx` (sesión
anterior) o `docs/Calculo_RE545_reducido_para_IA.xlsx` (ya en el repo, fuente directa de la
fórmula real que confirmó el fix) no tuvieran ya.

## 2026-09-11 (21) — Fix grande: `BK/BL/BS` de `Calculo RE545!BI:CE` cruzaban mal S contra BI

El usuario, todavía mirando la misma comparación real (`RE545 P11`), reportó más diferencias
usando los nombres de columna de **nuestra** salida (no las letras reales): "AR:AT, BI, AZ" y
después, más específico: "Curva Cmg Decendente promedio horario, (blanco), Curva Cmg Decendente".

**Primer chequeo (para no repetir la confusión de letras de la entrada anterior):** comparando por
contenido, `AR:AT` (el fix de la sesión 20) y `BI`/`AZ` con letra REAL (`Orden`, `EiniT` del
resumen) daban 0 diferencias — esos ya estaban bien. El problema real estaba en dos columnas
específicas que el usuario nombró explícitamente: `Curva Cmg Decendente promedio horario` (`BK`
real) y `Curva Cmg Decendente` (`BM` real).

**Causa exacta**, encontrada comparando fila a fila y confirmada contra las fórmulas guardadas del
archivo real (`docs/Calculo_RE545_reducido_para_IA.xlsx`, hoja `Mapa_Formulas`):

```
BK4 = SUMIFS(R:R, G:G,G4, S:S,BI4, T:T,T4, E:E,BJ4)
BL4 = SUMIFS(Q:Q, S:S,BI4, E:E,BJ4, G:G,G4, T:T,T4)
BS4 = IFERROR(INDEX(I, MATCH(1, (BR=BR4)*(S=BI4)*(G=G4)*(E=BJ4), 0)), "")
```

El criterio `S:S,BI4` compara la columna `S` (`ranking cmg`) de las OTRAS filas contra el `BI`
(`Orden`) de LA FILA ACTUAL — no `BI` contra `BI`. `calcular_bk_bl_bm_bs_re545()` armaba una sola
clave usando `BI` de los dos lados (la de acumulación Y la de búsqueda), lo que da el resultado
correcto únicamente cuando `S` y `BI` coinciden fila a fila por casualidad — que es exactamente lo
que pasaba en el único caso sintético que existía hasta ahora (por eso nunca se detectó). Con
datos reales, donde `S` y `BI` difieren, el agrupamiento salía mal.

Esto no se quedaba en `BK`/`BM`: al ser el insumo de `BN` (`Edisp_Asig`), `BO` (`Total C1_545`),
`CC` (`Total C2_545`) y, al final, `CE` (`Monto a compensar` — la ÚLTIMA columna de toda la hoja),
el error se propagaba a lo largo de toda la sección `BI:CE`.

**Fix:** `calcular_bk_bl_bm_bs_re545()` ahora arma DOS listas de claves — `claves_acumulacion`
(con `S` de cada fila, para poblar los diccionarios de suma/primer-valor) y `claves_busqueda` (con
`BI` de cada fila, para leer el resultado) — en vez de una sola clave usada de los dos lados.
También se corrigió el valor por defecto de `BS` sin match: antes daba `KeyError` (nunca pasaba
porque la clave vieja siempre existía en el propio diccionario); ahora, correctamente, da `NA`
(blanco), igual que el `IFERROR` real — a diferencia de `BK`/`BL`, que sin match dan `0` (`SUMIFS`
real).

**Verificación exhaustiva** (con `Pagos_BESS.xlsx` real, hoja `RE545 P11`, 26.784 filas):
- Reconstruyendo `BK`/`BM` a partir de los datos crudos de la propia salida (`Configuracion`,
  `Ventana de valorizacion`, `ranking cmg`, `CMg Promedio`, `CMg`, `Orden`, `Periodo`) y pasándolos
  por la función corregida: **0** diferencias contra los valores reales, en las 26.784 filas.
- Corriendo la etapa completa `BI:CE` (`calcular_componentes_re545()`) con `AU` y el resumen
  `AW:BG` reconstruidos con sus valores REALES (para no arrastrar la contaminación del bug de `AU`
  de la sesión anterior, que también ensuciaba el resumen `AW:BG` generado con el código viejo):
  `Edisp_Asig`, `Total C1_545`, `Total C2_545` y **`Monto a compensar`** (la columna final de toda
  la hoja) dan **0** diferencias en las 26.784 filas.
- Test sintético nuevo (`test_re545_bk_bl_bs_cruzado.py`, no persistido) que reproduce un caso con
  `S != BI` fila a fila y confirma que el resultado cambia respecto del comportamiento viejo (el
  test viejo, `test_re545d.py`, nunca hubiese detectado esto porque su `S` sintético coincidía con
  `BI` por construcción — se le agregó una columna `S` explícita, documentando por qué, para que
  siga siendo válido sin ocultar el hueco de cobertura).
- Regresión completa de las 20 sesiones anteriores: pasa.

**Confirmado también:** la contaminación de `Edisp_Asig`/`Total C1_545`/`Total C2_545` que
aparecía al principio de esta validación (antes de sustituir el resumen `AW:BG` por valores
reales) no era un bug nuevo — era el mismo bug de `AU` de la sesión 20 propagándose a través de
`BV` → `Margen ultima hora`/`Total Reservas* FD *FMA`/`Edisp_T` del resumen. Con los dos fixes (20
y 21) aplicados juntos en el pipeline real (donde el resumen se reconstruye siempre a partir del
`AU` ya corregido), no hace falta ningún fix adicional para esa cadena.

---

## 2026-09-11 (22) — `cmg.xlsx` ahora lo genera el programa (botón "Generar" en su fila del árbol)

**Pedido del usuario:** `cmg.xlsx` (entrada de `Cmg/`) se armaba a mano corriendo un script suelto
(`Extrae_CMG_barras.py`) al lado del CSV. Quería (1) un botón "Generar" al lado del nombre en la
ventana de Balance_BESS, (2) que el CSV de origen se busque en la ruta de red
`T:\CMgReales 15MIN\AAAA\AAMM\Mensual\CMg\Cmg para balance` en vez de al lado del `.py`, y (3) que
las barras a filtrar salgan de `Centrales.xlsx` (hoja `Resumen BESS`) en vez de estar escritas en
el código.

**Lo que se hizo:**

- `nucleo.py`: sección nueva "GENERACION DE cmg.xlsx DESDE EL CSV 15-MINUTAL" con
  `ruta_csv_cmg_15min()`, `barras_desde_resumen_bess()`, `construir_cmg_desde_csv()`,
  `_validar_layout_cmg()` y `generar_cmg(carpeta_base, aamm, ruta_csv=None, ...)` — misma firma
  `registrar`/`progreso` que `generar_consolidado`/`generar_pagos_bess`, así entra sin cambios en
  el `lanzar_generacion()` que ya existe en la ventana.
- Constantes nuevas: `RAIZ_CMG_REALES` (`T:\CMgReales 15MIN`), `SUBCARPETAS_CMG_REALES`,
  `PLANTILLA_CSV_CMG_15MIN`, `SEPARADOR_CSV_CMG`, `CODIFICACION_CSV_CMG`,
  `COLUMNA_CSV_CMG_VALOR`. La letra de unidad queda en UN solo lugar por si cambia.
- Las barras salen de `construir_mapa_barra()` (la misma función que ya alimenta
  `Calculo E Costos!Barra`), no de una lectura nueva: así el filtro del CSV y la homologación
  posterior **no se pueden desincronizar**. Se comparan en mayúsculas (mismo criterio que
  `_buscar_cmg`) pero se conserva el texto tal cual viene del CSV.
- `Balance_BESS.py`: la fila `cmg.xlsx` del árbol lleva su propio botón "Generar". No abre ventana
  con casillas como las dos salidas — no hay nada que elegir. Pide confirmación si el archivo ya
  existe. `pintar_arbol()` guarda ahora las referencias de los botones dibujados dentro del árbol
  (`botones_arbol`), y `terminar()` tolera que ese widget ya no exista (el árbol se repinta entero
  en cada `revisar()`).
- La fila `cmg.xlsx` del diagrama dice además si el CSV de origen del período está disponible o
  no, que es lo que decide si el botón va a poder hacer algo.

**Detalle que importa y es fácil de romper:** `leer_cmg()` lee `cmg.xlsx` **por posición**
(D = Barra, F = valor de Q, H = Cuarto de Hora, I = CMg promedio). El layout que sale de
`construir_cmg_desde_csv()` (columnas del CSV + `Cuarto de Hora` + promedio horario) es justo ese,
pero depende de que el CSV siga trayendo 7 columnas. Por eso `_validar_layout_cmg()` avisa en el
log si `BARRA` deja de caer en D o `Cuarto de Hora` en H, y corta con `ErrorEntrada` si quedan
menos de 9 columnas.

**Lo que NO cambió respecto del script original:** la numeración del `Cuarto de Hora` global sigue
saliendo de los bloques que el CSV realmente trae (no se asumen 96 por día), así que los días de
cambio de hora con 92/100 cuartos siguen funcionando; ahora además se listan en el log los días
que no tienen 24 h.

**Verificación:** caso sintético end-to-end (CSV de 2 días —uno con 23 h—, 3 barras en el CSV y 3
en `Centrales.xlsx`, una de ellas sin datos): genera `cmg.xlsx`, avisa de la barra sin datos,
detecta el día de 23 h, y el archivo resultante se vuelve a leer con `leer_cmg()` +
`construir_dic_cmg()` dando las claves `BARRA|cuarto` esperadas. `python -m py_compile
Balance_BESS.py nucleo.py` pasa. La ventana en sí no se pudo abrir (no hay `tkinter` en el
contenedor de la sesión): el cableado del botón se revisó a mano.

**Pendiente:** correrlo una vez contra el CSV real de la unidad `T:` para confirmar que las barras
de `Centrales.xlsx` están escritas exactamente igual que en el CSV (con el relleno de guiones
bajos, ej. `TOCOPILLA_____110`). Si alguna no coincide, el log lo dice barra por barra.

---

## 2026-09-11 (23) — Reorganización: paquete `Script/`, un botón por fila, salidas desglosadas por hoja

**Pedido del usuario**, seis puntos:

1. los botones a la izquierda del detalle;
2. el cálculo adentro de una carpeta `Script/`, con `nucleo.py` ahí y una subcarpeta `Cmg/`
   con `Extrae_CMG_barras.py` (la idea declarada es ir modularizando `nucleo.py` por etapas);
3. sacar la fila "Periodo (AAMM)" de entre las carpetas del diagrama — el SoC va dentro de
   `Medidas/`;
4. desglosar `Consolidado_entradas.xlsx` como carpeta, una fila por hoja, cada una con botón
   "Actualizar", y eliminar la ventana intermedia de "Generar". Si el archivo no existe, que
   se genere;
5. lo mismo para `Pagos_BESS.xlsx`;
6. el `cmg<AAMM>_def_15minutal.csv` pasa a vivir en la carpeta `Cmg/` del caso, al lado de
   `cmg.xlsx`, con un botón "Traer cmg_15min" que lo baja de la ruta de red.

**Estructura nueva del repo:**

```
Balance_BESS.py
Script/
    __init__.py
    nucleo.py
    Cmg/
        __init__.py
        Extrae_CMG_barras.py
```

`Balance_BESS.py` hace `from Script import nucleo`; `nucleo.py` hace `from .Cmg import
Extrae_CMG_barras` (con fallback a `from Cmg import ...` por si se importa suelto con
`Script/` en el `sys.path`). **Regla nueva, importante para la modularización que viene:**
un módulo de etapa NO importa `nucleo` — recibe rutas y datos, y levanta su propia excepción
(`ErrorCmg`), que `nucleo` traduce a `ErrorEntrada`. Así no hay ciclos de import cuando se
saquen más etapas. El nombre del archivo usa guiones bajos y no espacios para que sea
importable (el usuario lo escribió como "Extrae CMG barras.py").

**Contrato nuevo de `revisar_estructura()`:** antes devolvía tuplas
`(etiqueta, estado, detalle)` y la ventana deducía la profundidad de cada fila mirando el
TEXTO de la etiqueta (`_profundidad_fila()`: ¿termina en "/"?, ¿empieza con dos espacios?).
Eso ya venía frágil y con el desglose por hojas no daba más. Ahora devuelve dicts
`{id, etiqueta, nivel, estado, detalle}`:

- el **nivel** lo pone `nucleo` porque es estructura, no dibujo (que un archivo esté adentro
  de una carpeta, o una hoja adentro de un archivo, no es una decisión de interfaz);
- el **id** es estable y es lo único que la ventana necesita para saber qué botón colgarle a
  cada fila (`_boton_de_fila()`), así `nucleo` sigue sin saber nada de botones.

**Ventana:**

- columnas `Estructura | Estado | Acción | Detalle`. La celda de acción tiene ancho FIJO en
  píxeles (`ANCHO_ACCION`), si no cada fila correría el detalle según el largo de su botón.
- desaparecieron las dos ventanas "Generar" con casillas. Botones: `Traer cmg_15min` y
  `Generar` en `Cmg/`, `Actualizar` en cada fila-hoja de las dos salidas, `Actualizar todo`
  en la fila del archivo.
- mientras corre algo, TODOS los botones del árbol quedan deshabilitados (`corriendo` +
  `habilitar_botones`), y al terminar el árbol se repinta entero (que es lo que los
  rehabilita).
- el campo del período (AAMM) sigue arriba; lo que se sacó es la FILA del diagrama. El SoC
  quedó donde corresponde, colgando de `Medidas/`.

**CMg en dos pasos:** `nucleo.traer_csv_cmg()` copia el CSV de la unidad de red a
`<CARPETA_BASE>/Cmg/` y `nucleo.generar_cmg()` arma `cmg.xlsx` con el CSV que quedó ahí. Se
copia en vez de leer directo de la red a propósito: el caso queda autocontenido (se puede
regenerar `cmg.xlsx` sin la unidad conectada) y queda registrado con qué archivo se trabajó.

**Dos bugs encontrados de paso, los dos arreglados:**

1. **El `Log` de `Consolidado_entradas.xlsx` perdía todos los avisos de preservación.**
   `escribir_salida()` armaba el `df_log` ANTES del bloque `with pd.ExcelWriter(...)`, pero
   los avisos de "no se regeneró la hoja X y no había versión anterior" los agrega
   `_preservar_o_avisar()` DENTRO de ese bloque. Resultado: el `Log` decía "Sin
   observaciones" mientras cuatro hojas quedaban vacías. Nunca se había notado porque hasta
   ahora lo normal era generar todo junto; con un botón por hoja, generar una sola es el caso
   normal y esos avisos son justamente los que hay que ver. El log ahora se arma al final
   (`_armar_log()`).
2. **El estado de una hoja no se podía deducir de que la hoja existiera.** Al generar una sola
   sección, el archivo se crea con TODAS las hojas (las no pedidas, vacías), así que el
   diagrama las mostraba todas como "generada". Se agregó `hojas_con_datos()` (openpyxl en
   `read_only`, `max_row > 1`): una hoja vacía se ve PENDIENTE. El estado de la fila del
   archivo es ahora el rollup de sus hojas ("le faltan hojas por generar").

**Verificación:** caso sintético con `Consolidado_entradas.xlsx` a medias — el árbol
renderizado (niveles, prefijos, estados, botones) sale correcto y el SoC cuelga de `Medidas/`;
`traer_csv_15min` + `generar_cmg` contra una "unidad de red" falsa, con el `cmg.xlsx`
resultante releído por `leer_cmg()`/`construir_dic_cmg()`; `generar_consolidado()` de una sola
sección sobre una carpeta sin el archivo (lo crea, deja el resto de las hojas vacías y ahora
sí las lista en el `Log`); y el error esperado cuando se pide "Generar" sin haber traído el
CSV. `py_compile` de los tres módulos pasa. La ventana en sí no se pudo abrir (no hay
`tkinter` en el contenedor de la sesión): el cableado se revisó a mano y los helpers del árbol
(`_prefijos_arbol`) se probaron aparte, sin tkinter.

**Pendiente que deja esta sesión:** abrir la ventana una vez en Windows para confirmar el
ancho de la columna "Acción" (`ANCHO_ACCION = 150 px`) contra los botones más largos
("Traer cmg_15min", "Actualizar todo") y el alto de fila (`ALTO_ACCION = 26 px`).

---

## 2026-09-11 (24) — `Medidas_SAE.xlsx` ahora lo genera el programa (paquete `Script/Medidas/`)

Misma idea que con `cmg.xlsx` (sesión 22), aplicada a la primera entrada del caso. El usuario
entregó los cuatro scripts sueltos que se corrían a mano uno detrás de otro (autor original:
Freddy.Arriagada) y pidió seis cosas:

1. el Excel de homologación pasa a `Auxiliares/`, al lado de `Centrales.xlsx`;
2. `Medidas_SAE.xlsx` tiene un botón "Actualizar" que corre todo de un viaje;
3. la lista `FILTROS_TOPOLOGY` del script 3 sale del código y pasa a `Centrales.xlsx`; **deja de
   ser un reemplazo**: esas centrales se sacan del archivo de homologación, así que el paso 3 las
   **agrega**, con la clave que indique la lista;
4. la hoja nueva de `Centrales.xlsx` la diseño yo y él la crea;
5. los códigos van en `Script/Medidas/`, hermana de `Script/Cmg/`;
6. las salidas intermedias no se ven en la ventana.

**El hallazgo que ordenó todo:** el paso 2 terminaba escribiendo exactamente las 9 columnas de
`nucleo.COLUMNAS_AI`, en el mismo orden (`Mes, Dia, Hora, Minutos, Hora Mes, Cuarto de Hora,
clave, intervalo, Gen_Unidad`). O sea que esta cadena ya era, sin saberlo, el generador de la
entrada que hasta ahora había que dejar a mano en `Medidas/`.

**Estructura:** `Script/Medidas/` con `comun.py` (`ErrorMedidas` + helpers de texto),
`Homologacion.py`, `Descarga_PRMTE.py`, `Claves_Balance.py` y `Generacion_Real.py` — uno por
script original. Se mantiene la regla de `Cmg/`: un módulo de etapa **no importa `nucleo`**;
recibe rutas y datos y levanta su propia excepción, que `nucleo` traduce a `ErrorEntrada`.
`nucleo.generar_medidas_sae()` orquesta los cuatro pasos.

**La hoja `Medidas API` de `Centrales.xlsx`** (diseñada acá, la crea el usuario): `topologyName`
(el nombre exacto de la API), `clave` (con qué nombre aparece en `Medidas_SAE.xlsx`) y `Factor`
(opcional, 1 por defecto; es el equivalente del `Flujo` del archivo de homologación, `-1` para los
retiros). Dos filas pueden apuntar a la misma clave: se suman. La hoja entera es opcional — sin
ella no se agrega ninguna central por ese camino y el resto corre igual.

**Decisión de diseño que importa:** el `Cuarto de Hora` es un índice global del mes que después
cruza contra `CMg`, así que las dos fuentes **comparten un solo calendario** (el que arma
`Claves_Balance`, numerado por `intervaloUtc` porque la hora local se repite en el cambio de hora).
Si cada fuente numerara por su cuenta, un día de cambio de hora las desalinearía en silencio. La
API de operación real no entrega UTC, así que sus filas se pegan por hora local: en un día de
cambio de hora hacia atrás hay ambigüedad real, se toma la primera ocurrencia y se avisa en el
log. Si el cruce da 0 coincidencias, se corta con un error explícito en vez de escribir un
`Medidas_SAE.xlsx` al que le faltan esas centrales.

**Credencial:** los scripts traían el `user_key` escrito adentro (vacío en el 1, `"-"` en el 3).
Primero lo saqué a un campo de la ventana guardado en `config.json`, señalando que al ser una
credencial no debería quedar versionada; **el usuario decidió dejarlo en el código** y se hizo así.
Queda en una sola constante, `USER_KEY` en `Script/Medidas/comun.py`, para las dos APIs — el
problema real que tenía era estar repetido en dos archivos y con valores distintos. Consecuencia
asumida, anotada en `METODOLOGIA.md` §5: queda versionada, así que el repositorio no puede volverse
público sin rotarla antes.

**Dos bugs de los scripts originales, arreglados** (no estaban en el pedido):

1. **Los lotes descargados no llevaban el período en el nombre.** El paso 2 hacía
   `glob("medidas_batch_*.parquet")` y el paso 1 anotaba los puntos ya procesados en un archivo
   único, así que correr dos meses en la misma carpeta mezclaba los lotes de los dos y daba por
   procesados puntos de otro mes. Ahora los dos llevan el período en el nombre. (De paso se fue el
   `punt∟os_procesados.txt`, con un carácter raro en medio del nombre.)
2. **El umbral de "punto de medida completo" era la constante `2976`** (= 31 × 96). Está mal para
   cualquier mes de 30 días o menos —descartaría todos los puntos— y para los meses con cambio de
   hora. Ahora es la cantidad de cuartos de hora que el mes descargado realmente trae
   (`intervaloUtc` distintos), que es lo mismo que `2976` pretendía ser.

**Lo que se sacó a propósito:** el `log_inconsistencias_medidas.xlsx` del script 3, que comparaba
el criterio de desempate viejo contra el de mayor `idMeasure`. Era una investigación ya cerrada (el
propio script titula esa sección "CRITERIO DEFINITIVO"). Lo que sí queda en el log es cuántos
grupos venían duplicados. Los diagnósticos del paso 2 (`reporte_medidas_consolidadas.xlsx`) también
se mudaron al log, por el punto 6 del pedido. Los lotes y la marca de reanudación viven en
`<CARPETA_BASE>/Medidas/_trabajo/`, que la ventana no muestra.

**Refactor menor:** `_leer_resumen_bess()` se generalizó en `_leer_hoja_con_encabezado(ruta, hoja,
columnas_buscadas)` para poder leer la hoja nueva con el mismo criterio (buscar la fila de
encabezados en vez de asumir la primera, porque las hojas reales traen un título arriba).
`_leer_resumen_bess()` quedó como un caso particular.

**Verificación:** test sintético de punta a punta con DataFrames con la forma de las dos respuestas
de API (monkeypatch de `Descarga_PRMTE.descargar` y `Generacion_Real.descargar_mes`; el resto del
proceso es puro pandas). Dos días, tres puntos de medida —uno incompleto a propósito—, tres
centrales en la hoja `Medidas API` —dos apuntando a la misma clave con factores opuestos— y un
grupo duplicado para probar el desempate. Resultados verificados a mano: el punto incompleto se
descarta (187 vs 192 cuartos), `SAE-UNO` da −1 por cuarto (canal 1 con `Flujo` +1 menos canal 3 con
`Flujo` −1), `SAE-ANDES-III` da 0 (inyección + retiro), `PFV-ANDES-IV` da 192 + 4 del desempate, la
central que no está en la lista no aparece, las tres claves comparten los cuartos 1..192, y el
archivo resultante lo relee `leer_medidas_sae()` con las columnas exactamente iguales a
`COLUMNAS_AI`. Más el caso sin la hoja `Medidas API` (devuelve lista vacía y el proceso sigue), la
regresión de las sesiones 22 y 23, y `py_compile` de todos los módulos.

**Lo que NO se pudo probar:** las dos funciones que hablan con la API (no hay red ni credencial en
el contenedor) y la ventana (no hay `tkinter`). Ver pendientes.

**Pendientes que deja esta sesión:**

- Correr el botón una vez contra la API real: confirmar la forma de la respuesta (que
  `mediciones` traiga `intervalo`/`intervaloUtc`/`canalVal1`/`canalVal3`/`principal`, y que el
  filtro de puntos incompletos siga descartando lo mismo que antes), y que el `intervalo` de las
  dos APIs sea el INICIO del cuarto de hora en las dos — de eso depende el cruce contra el
  calendario, y si no coincidiera el log lo va a decir fuerte ("ninguna fila cruza").
- Confirmar con el usuario si `Factor` hace falta o si todas las centrales de la hoja van con 1.
  El script original no aplicaba ningún signo; lo agregué porque, sin él, un retiro no tiene cómo
  expresarse — pero con el archivo real puede resultar que la API ya entregue el signo.

---

## 2026-09-11 (25) — La homologación de Gen real se muda al archivo de homologación

El usuario entregó el `Homologacion ClavesTF y PRMTE.xlsx` real y cambió de opinión sobre dónde va
la lista de centrales de operación real: **no** en `Centrales.xlsx` (como se había hecho en la
sesión 24) sino en el **mismo archivo de homologación**, en una hoja `Gen real`, "con la misma info
que ahí: clave, Punto de Medida, Canal, Flujo". Tiene sentido: es homologación igual que `homol`, y
así se mantiene con el mismo archivo en vez de repartida en dos.

**El archivo real** (útil para futuras sesiones): una sola hoja `homol`, 74 filas, columnas
`clave | Punto de Medida | Canal | Flujo`. `Punto de Medida` es el `idPuntoMedida`
(`DNHUMBER_033_FB1_EGP`), `Canal` es el `slugCanal` (`kWhD` / `kWhR`) y `Flujo` es ±1. 7 claves,
37 puntos de medida. Los lectores se probaron contra él antes de tocar nada.

**Lo que se movió:** `nucleo.leer_medidas_api()` y la constante `HOJA_MEDIDAS_API` desaparecen;
ahora es `Homologacion.leer_gen_real()` + `HOJA_GEN_REAL`. `generar_medidas_sae()` ya **no necesita
`Centrales.xlsx` para nada** (era su única dependencia con ese archivo). En el diagrama, el archivo
de homologación se desglosa por hojas igual que `Centrales.xlsx`: `homol` (obligatoria) y
`Gen real` (opcional, se ve PENDIENTE si no está).

**La decisión que tuve que tomar solo, porque el usuario está fuera:** de las cuatro columnas, tres
se leen solas (`clave` es la clave del balance; `Flujo` es el ±1 que en la sesión 24 se llamaba
`Factor`; `Punto de Medida` tiene que ser el `topologyName` de la API de operación real, que es lo
único que identifica a una central en esa API). La que no tiene lectura obvia es **`Canal`**: la
API de operación real no expone canales. Se acepta la columna (para que la hoja tenga la misma
forma que `homol`, que es lo que se pidió) pero **no se usa** para nada. Queda anotado como
pendiente por si tenía que significar algo.

**Verificación:** el test sintético de punta a punta de la sesión 24, con el `Gen real` ahora en el
archivo de homologación y con las columnas en el orden real (`clave` primero): mismos resultados
exactos que antes (`SAE-UNO` −192, `SAE-ANDES-III` 0 por inyección+retiro, `PFV-ANDES-IV` 196 con
el desempate), o sea que mover la hoja no cambió ningún comportamiento. Más: `leer_homologacion()`
y `leer_gen_real()` contra el archivo real subido (74 filas / hoja ausente → lista vacía), el árbol
renderizado con ese archivo en `Auxiliares/`, y la regresión de CMg. `py_compile` de todo.

**Pendiente nuevo:** confirmar si la columna `Canal` de la hoja `Gen real` tiene que significar
algo. Hoy se ignora.
