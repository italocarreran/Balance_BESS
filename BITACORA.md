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
- Obtener del usuario los encabezados reales de la hoja `Calculo E Costos`
  ("Ecostos"): las dos veces que adjuntó un archivo pensado para esto solo
  traía las hojas `FD`/`Subastas` (ya confirmadas). Mientras tanto,
  `construir_calculo_e_costos()` usa nombres placeholder derivados de los
  comentarios de la macro (ver plan §25.4, `METODOLOGIA.md` §7).
- Completar el resto de `Actualizar_Calculos_Columnas` (L, M, N, O, R, S,
  T, U, W, X, Y, AB:AF, AG:AX, AZ) y toda la hoja `Calculo RE545` — la
  etapa base (H, CMg, traspaso de Medidores) ya está implementada (ver
  entrada de esta sesión). Incluye construir la "Prorrata SSCC" como tabla
  dinámica derivada de `Subastas` (confirmado por el usuario que no es un
  archivo externo: `Filas: Configuración, Hora_mes` / `Columnas: Control` /
  `Valores: Cuenta de Sub_Baj`).
- Una vez completo `Calculo E Costos`, resolver `Subastas!N` ("Energía
  SSCC"), que depende de columnas de esa hoja.
- Confirmar si la carpeta `Subastas/` (creada esta sesión, no existe en
  la planilla original) es el nombre/ubicación que se quiere mantener, o
  si se prefiere buscar el archivo directamente en `<CARPETA_BASE>` como
  hacía la macro original (ver plan §23.3).
- Confirmar el nombre definitivo de `Pagos_BESS.xlsx` (provisorio, elegido
  por el usuario como "pagos_bess o algo así por ahora").
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
