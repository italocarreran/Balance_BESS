# BITACORA.md — registro de sesiones

Solo se agrega. Nunca se edita ni borra una entrada vieja. La única
excepción es la sección "Pendientes abiertos", que sí se edita porque es un
estado, no un historial.

---

## Pendientes abiertos

- Crear casos de prueba con datos reales y comparar la salida Python
  contra la hoja `Medidores` de `11_PAGOS_BESS_2607_Definitivo.xlsm`,
  incluyendo ahora `R, S, T` y las tres hojas auxiliares de Ofertas SSCC
  (plan §13 punto 10, §20.3). Todavía solo se validó con un caso sintético.
- Confirmar si la columna `V` (clave de homologación día+central, interna
  a `calcular_r`) necesita persistirse en una hoja propia para poder
  auditarla fila a fila contra la planilla 11, o si alcanza con auditar
  "Ofertas SSCC por Dia" + `Diccionario!E:F:G` a mano.
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
