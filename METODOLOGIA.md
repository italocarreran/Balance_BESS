# METODOLOGIA.md — cómo se trabaja en este repositorio

> Documento vivo. Si en una sesión se descubre algo que la siguiente necesita
> saber, se agrega acá antes de cerrar.

---

## 0. Por qué existe este documento

Este repositorio no existe solo para "guardar el código". Existe para que un
asistente de IA pueda entender el sistema **sin leerlo entero**. Cada regla de
acá abajo sale de esa única necesidad: minimizar cuánto hay que leer para
hacer un cambio correcto.

Si el repositorio lo trabaja más de una persona o más de un asistente con
acceso de escritura, sin verse en tiempo real, este documento es además la
única red de seguridad que existe entre ellos. No son sugerencias.

---

## 1. Qué es este repositorio

Herramienta en Python que reemplaza, hoja por hoja, el cálculo hecho hoy en
`11_PAGOS_BESS_2607_Definitivo.xlsm` (Balance BESS / SSCC). Implementa hasta
ahora **Medidores** (incluida Ofertas SSCC), la carga de **CMg**, **FD** y
**Subastas**, y una primera etapa (base) de **Calculo E Costos**: dos
scripts interdependientes que comparten un `config.json` guardado junto al
código (no versionado, es de la herramienta, no del caso).

- `Balance_BESS.py` — ventana tkinter. Único punto de entrada para el
  usuario.
- `nucleo.py` — todo el cálculo, sin dependencias de interfaz. `Balance_BESS.py`
  importa `nucleo` y nunca al revés.

Cada caso a procesar vive en su propia carpeta ("carpeta base"), fuera del
repositorio, con la estructura fija documentada en
`docs/Plan_Traspaso_Python_Balance_BESS.md` §3.3. El repositorio no contiene
datos de ningún caso.

---

## 2. Regla de expansión de contexto

La regla que no se negocia, para cualquier sesión que toque este repositorio:

```
1. Leer MAPA.md                       (mapa corto: qué hace cada script/módulo)
2. Leer de este documento SOLO las secciones que hacen falta
3. Abrir completo ÚNICAMENTE el archivo que se va a modificar
```

**No abrir archivos vecinos "para tener contexto".** Si de verdad hace falta
uno más, pedirlo explícitamente y decir por qué. Leer de más es exactamente
el problema que esta metodología resuelve.

Para las reglas de negocio del cálculo (fórmulas de Excel replicadas, de
dónde sale cada dato, qué columnas siguen pendientes) el documento de
referencia es `docs/Plan_Traspaso_Python_Balance_BESS.md`. Es largo a
propósito y por eso mismo **no se lee entero**: se busca la sección puntual
con grep o por su encabezado (p. ej. "9.4" para la columna `Ventana`).

**Dónde vive el código**, con nombres exactos:

- `Balance_BESS.py` — módulo/entrada principal, se abre siempre primero.
- `nucleo.py` — módulo de cálculo compartido, vive junto al principal en la
  raíz del repositorio, no adentro de una subcarpeta.
- `config.json` — configuración/estado por PC/usuario (última carpeta base
  elegida). No se versiona (ver `.gitignore`). Vive junto a `Balance_BESS.py`
  únicamente porque así lo resuelve `Path(__file__).parent` en el propio
  script; el resto del código nunca debe depender de esa ruta.
- `docs/Plan_Traspaso_Python_Balance_BESS.md` — documento de dominio.

No hay hoy scripts satélite ni un módulo común adicional: solo estos dos
archivos. Si `Balance_BESS.py` se mueve de carpeta, `nucleo.py` debe moverse
con él (import relativo simple, sin paquete).

---

## 3. El set de documentos y para qué sirve cada uno

| Documento | Para qué | Quién lo edita y cuándo |
|---|---|---|
| `REGLAS.md` | Checklist obligatorio de inicio/cierre de sesión. Se lee **entero, primero**, antes que nada. | Se edita poco; es la red de seguridad, no debe crecer sin necesidad. |
| `METODOLOGIA.md` | Cómo se trabaja: regla de expansión, convenciones de código, trampas conocidas, decisiones ya tomadas. | Documento vivo — se actualiza cuando cambia algo estructural o se toma una decisión de diseño. |
| `BITACORA.md` | Registro de qué se hizo en cada sesión y qué quedó pendiente — lo que un `git log` no cuenta. | **Solo se agrega.** Nunca se edita ni borra una entrada vieja. La única excepción es la sección "Pendientes abiertos", que sí se edita porque es un estado, no un historial. |
| `MAPA.md` | Un bloque corto por script: qué hace · consume · produce · expone · depende de. Es la primera lectura de cualquier sesión. | Se actualiza cuando cambia la estructura o el rol de un script. |
| `README.md` | Puerta de entrada: instalación/uso y tabla de "querés X → leé Y". | Se actualiza si cambia la estructura de carpetas o el flujo de instalación. |
| `docs/Plan_Traspaso_Python_Balance_BESS.md` | Reglas de negocio del cálculo: fórmulas, de dónde sale cada dato, qué queda pendiente. Es largo a propósito y no se reescribe por comodidad. | Se corrige ahí mismo cuando el dominio cambia (p. ej. se confirma la lógica de una columna pendiente); nunca se duplica en otro archivo. |

**Regla de prioridad de lectura:** `REGLAS.md` → `METODOLOGIA.md` →
`MAPA.md` → el archivo a modificar → (si hace falta) la sección puntual de
`docs/Plan_Traspaso_Python_Balance_BESS.md`.

**Donde el código y el documento de dominio difieren, manda el código.**
Las diferencias detectadas se anotan en una tabla al final de `MAPA.md`
("Diferencias con el documento de dominio"); al corregir el documento, se
borra la fila correspondiente.

No existe todavía un `INTERFACES.md` generado ni un generador de interfaces:
con dos archivos, `MAPA.md` alcanza. Si el proyecto crece a varios módulos,
reevaluar (ver §10 de la plantilla original de esta metodología).

---

## 4. Cómo se entregan los cambios

Por ahora este repositorio lo trabaja un asistente de IA con revisión
humana (no hay más de un asistente con push directo todavía). Si eso
cambia, esta sección debe actualizarse para documentar cómo se coordinan
entre sí (rama base compartida, `BITACORA.md` como único canal entre
sesiones que no se ven en tiempo real, prefijo de autor en los commits si
alguno autentica con el nombre de otra persona).

- Rama base única (la rama por defecto del repositorio remoto). Todo sale
  de ahí y todo vuelve ahí. Una rama de trabajo se fusiona y se borra; no
  se acumulan ramas vivas.
- Nunca reescribir el historial compartido (`push --force`, `commit --amend`
  sobre algo ya subido, `reset --hard` contra la rama remota).
- Si en algún momento se vuelve a un flujo sin push directo (el usuario sube
  y baja archivos a mano): entregar siempre el archivo completo, nunca un
  diff ni "reemplazá la línea X por esto" — quien lo recibe lo corre tal
  cual, y editar a mano es donde se cuelan los errores.

No hay todavía scripts de apoyo (`scripts/sincronizar.sh`,
`scripts/verificar.sh`). Mientras no existan, la verificación antes de
cerrar una sesión es manual: correr `python -m py_compile Balance_BESS.py
nucleo.py` y, si hay un caso de prueba disponible, `nucleo.ejecutar(...)`
contra él.

---

## 5. Convenciones de código establecidas

- **Interfaz de usuario:** una única ventana tkinter (patrón "carpeta base +
  Examinar + checklist de entradas + Ejecutar + log + barra de progreso +
  contador de tiempo"). Los estados del checklist son exactamente tres:
  `ok` (verde), `falta` (rojo, bloquea Ejecutar) y `pendiente` (ámbar, no
  bloquea). No agregar un cuarto estado sin actualizar `SIMBOLO` y
  `COLOR_ESTADO` en `Balance_BESS.py` a la vez.
- **Persistencia de configuración:** `config.json` junto al `.py`, con una
  clave por PC/usuario (`get_usuario()` = `hostname_usuario`), para que
  varias personas puedan compartir la misma copia del script sin pisarse la
  carpeta recordada. Si el archivo existe pero no se puede leer/parsear, se
  ignora en silencio y se sigue — es preferible perder el ajuste recordado
  a que el programa no abra. La escritura reescribe el `config.json`
  completo (no es atómica todavía; ver §7).
- **Rutas de un caso:** nunca rutas absolutas ni dependientes de
  `Path(__file__).parent` para los archivos de un caso. Todo se deriva de la
  carpeta base vía `resolver_rutas()` en `nucleo.py`. `Path(__file__).parent`
  se usa únicamente para `CONFIG_PATH` en `Balance_BESS.py`.
- **Lectura de Excel:** con `pandas` (`read_excel`/`ExcelFile`) y escritura
  con `openpyxl` como engine. El SoC se extrae por **detección dinámica de
  bloques por encabezados** (`detectar_fila_nombres` → `detectar_bloques` →
  `extraer_soc`): nunca por letra de columna fija ni por offset constante
  entre el nombre de la central y sus columnas `Time Stamp`/`Value` (regla
  del plan, §6.1 de `docs/Plan_Traspaso_Python_Balance_BESS.md`).
- **Normalización de texto:** toda comparación de nombres de central/hoja
  pasa por `normalizar()` en `nucleo.py` (minúsculas, sin tildes, espacios
  colapsados). No reimplementar una variante local de esta función en otro
  archivo.
- **Homologación de nombres:** se resuelve siempre contra la hoja
  `Diccionario` de `Centrales.xlsx` vía `construir_homologacion()`. La
  primera réplica no corrige ni reinterpreta homologaciones aunque parezcan
  desplazadas; cualquier inconsistencia se reporta como aviso/incidencia,
  no se "arregla" en silencio.
- **Período del caso (AAMM):** lo ingresa el usuario en un campo de texto
  de la ventana (4 dígitos, ej. `2607`), no se infiere del nombre de ningún
  archivo. `nucleo.validar_aamm()` es la única función que valida el
  formato; todo lo demás (`buscar_soc`, `revisar_estructura`, `ejecutar`)
  recibe el AAMM ya como parámetro. El archivo de SoC dentro de `Medidas/`
  tampoco tiene un nombre fijo: solo debe contener "SOC" y el AAMM en
  cualquier posición del nombre (`_es_archivo_de_soc()`); el archivo en sí
  siempre es `.xlsx` (confirmado con el usuario — un CSV con "SOC"+AAMM en
  el nombre puede ser un archivo completamente distinto sin relación con
  el SoC, ver `METODOLOGIA.md` §7).
- **Columnas de `Medidores` (A:U):** el orden final de columnas sale de
  `LETRA_A_CAMPO`, cuyo **orden de inserción** es el orden de Excel.
  `COLUMNAS_VACIAS` (M, P, Q, U) son diseño confirmado, no trabajo
  pendiente. `V, W, X, Y, AB, AC, AD, AE` **no están en `LETRA_A_CAMPO`**:
  en la planilla original no son una columna por fila de `Medidores`, son
  tablas auxiliares de otro largo (central × día, central × ventana) que
  comparten esas letras de columna solo porque ahí había espacio libre. Se
  calculan y se escriben juntas en una sola hoja auxiliar
  (`HOJA_OFERTAS_SSCC = "Ofertas SSCC"`, una tabla al lado de la otra con
  su propio título vía `_escribir_tabla_con_titulo()`, que acepta tanto
  `fila_inicio` como `columna_inicio`) en vez de forzarlas a columnas
  `pd.NA` del mismo largo que A:U (ver plan de migración §20.1, §22, §24.3).
  El resumen intermedio equivalente a la hoja
  "Resumen Ofertas SSCC" del `.xlsm` original (con una columna por
  servicio `_RS`) es puramente auxiliar para construir la tabla W:Y — no
  se persiste en `Consolidado_entradas.xlsx`, solo vive en memoria dentro de
  `construir_medidores()`. `R`, `S`, `T` sí son columnas por fila y están
  implementadas: dependen de las macros de Ofertas SSCC
  (`Generar_Resumen_Ofertas_SSCC`, `Resumir_Medidores_Central_Ventana_
  Oferta_Completa`), replicadas fielmente a partir del código VBA y las
  fórmulas de Excel entregados (plan §20).
- **Ofertas SSCC:** archivo obligatorio (plan §17-18), se busca en
  `<CARPETA_BASE>/Ofertas/` con `buscar_archivo_ofertas()` — nombre debe
  contener "OfertasSSCC" (sin importar mayúsculas); a diferencia del SoC,
  si hay más de uno se toma el más reciente por fecha de modificación
  (replica exacta de `OSSCC_BuscarArchivoOfertas`, no una decisión nueva).
  Homologación de nombres para Ofertas SSCC usa específicamente
  `Diccionario!E/F/G` (índices 4/5/6 del DataFrame `header=None`) vía
  `_mapas_homologacion_fge()`/`_homologar_fge()` — es un mapeo DISTINTO del
  que usa `construir_homologacion()` para el SoC (que trata toda la fila
  como equivalencias simétricas); no confundir ni fusionar ambos.
- **Errores de entrada vs. errores inesperados:** un problema de datos de
  entrada (archivo faltante, ambigüedad de SOC, columnas faltantes, bloque
  sin `Time Stamp`/`Value`) se señaliza con `nucleo.ErrorEntrada`, con un
  mensaje explicativo para el usuario. No usar excepciones genéricas para
  esto: `Balance_BESS.py` distingue ambos casos para mostrar un mensaje
  distinto.
- **Columnas deliberadamente vacías:** se agregan igual al DataFrame de
  salida (como `pd.NA`, listadas en `COLUMNAS_VACIAS` para `Medidores`) en
  vez de omitirse, para que la forma de `Consolidado_entradas.xlsx` sea
  comparable con la planilla 11 aunque la columna no tenga valor.
- **CMg, FD, Subastas (plan §23):** replican únicamente las macros de
  *carga* de esas hojas (`Cargar_CMg_Desde_Archivo`,
  `Cargar_SSCC_Desempeno_En_FD`, `Cargar_Remuneracion_Subastas_Rapido`).
  Ninguna de las tres tiene un documento de dominio tan detallado como
  Medidores; sus nombres
  de columna (`NOMBRES_FD_CSF`, `NOMBRES_FD_CPF`, `NOMBRES_SUBASTAS`, plan
  §24) los confirmó el usuario contra un caso real, no se inventaron. Si
  aparece una columna sin ese respaldo, usar su letra de Excel tal cual
  (p. ej. `"N"`) en vez de inventarle un nombre de negocio no documentado —
  mismo criterio de "no adivinar" que el resto del proyecto. Ojo con
  nombres duplicados dentro de un mismo bloque (p. ej. `FD` repite "Hora
  Mes" en B y M, y en R y AE): se renombra con `set_axis()` recién al
  final, después de calcular todo con nombres de letra únicos — Python no
  prohíbe columnas duplicadas, pero indexar por ese nombre durante el
  cálculo sería ambiguo.
- **`FD` tiene el mismo patrón de "tablas de distinto largo compartiendo
  hoja" que Ofertas SSCC, pero por columnas en vez de por filas:** el
  bloque CSF (A:M, viene de `CSF Horario`) y el bloque CPF (Q:AE, viene de
  `CPF Horario`) se filtran y calculan por separado (pueden tener distinta
  cantidad de filas) y se escriben lado a lado (`escribir_salida()`, vía
  `startcol` en `df.to_excel()`), no una debajo de la otra.
- **Filtro BESS/SAE de FD y Subastas vs. el de Ofertas SSCC:**
  `_contiene_bess_o_sae_sin_bat()` (FD, Subastas) NO incluye "BAT";
  `_contiene_bess_o_sae()` (Ofertas SSCC) sí. Son dos filtros distintos que
  se parecen — no fusionarlos en una sola función aunque parezca tentador.
- **Subastas!N queda vacía a propósito:** su fórmula real depende de
  `'Calculo E Costos'!D/G/P`; la hoja ya existe (en `Pagos_BESS.xlsx`,
  etapa base, plan §25) pero todavía no las columnas específicas que esa
  fórmula necesita. No se adivina su valor.
- **`Calculo E Costos` (plan §25), etapa base, en archivo separado
  (`Pagos_BESS.xlsx`, nombre provisorio):** replica solo una parte de
  `Traspasar_Medidores_A_Calculos_Rapido` (traspaso A:G con D↔E
  invertidas, I/J según signo y `Ventana_No_Completa`, J→K, K→P) y de
  `Asignar_CMg_a_Calculos_Turbo` (columna Q, con `escribirR=False` — la
  hoja no incluye `Calculo RE545`). `H` (Barra) es fórmula
  (`VLOOKUP(G,Resumen!B:G,6,FALSE)` en el original) y se homologa por
  **nombre de columna** contra `Resumen BESS!Nombre activo`/`Barra
  inyección` (`construir_mapa_barra()`), no por posición: `Centrales.xlsx`
  no reproduce el layout `Resumen!B:G` del libro original. El resto de
  `Actualizar_Calculos_Columnas` (L, M, N, O, R, S, T, U, W, X, Y, AB:AF,
  AG:AX, AZ) y `Calculo RE545` completo quedan pendientes — decisión
  explícita del usuario de avanzar por etapas. Nombres de columna:
  placeholders derivados de los comentarios de la macro, todavía sin
  confirmar contra un archivo real (ver trampa en §7).

---

## 6. Generación de `INTERFACES.md`

No aplica todavía: no existe un generador de interfaces en este repositorio.
Con dos archivos (`Balance_BESS.py`, `nucleo.py`) alcanza con `MAPA.md`. Si
se agregan más módulos y esto deja de ser suficiente, documentar acá la
decisión de introducir un generador (o no) antes de empezar a usarlo.

---

## 7. Trampas conocidas

Tabla de solo agregar: cada vez que un bug cueste tiempo real de
investigación, se documenta acá con la regla que lo evita, **antes** de
cerrar la sesión que lo encontró. No se borran filas viejas salvo que la
causa raíz deje de existir en el código.

| Trampa | Regla |
|---|---|
| `guardar_config()` reescribe `config.json` entero sin escritura atómica. Un corte a mitad de escritura puede dejar el archivo corrupto. | Al tocar esa función, evaluar escritura atómica (escribir a un temporal y `rename`). Mientras tanto, `leer_config()` ya tolera un JSON corrupto devolviendo `{}`, así que el peor caso es perder la carpeta recordada, no romper el programa. |
| Si `Medidas/` tiene más de un archivo `SOC_AAMM.xlsx`, `buscar_soc()` lanza `ErrorEntrada` a propósito — no elige el más reciente. | No "arreglar" esto para que elija automáticamente por fecha de modificación: es una decisión deliberada del plan (§3.4) para no tomar en silencio el mes equivocado. |
| Las columnas `K, M, P, Q, R, S, T` de `Medidores` están en `COLUMNAS_PENDIENTES` como `pd.NA` porque su lógica exacta o su fuente (Ofertas SSCC) todavía no está definida. | No inventar una fórmula para completarlas "para que quede bonito". Cerrar primero la regla exacta en `docs/Plan_Traspaso_Python_Balance_BESS.md` §9, con el humano que conoce la planilla 11, y recién ahí implementar. |
| `calcular_ventana()` reinicia el contador por **bloque de filas consecutivas con la misma clave**, no por `groupby` sobre toda la central. | Si los datos de entrada no vienen ordenados por `clave` e `intervalo` antes de llamar a esta función, el resultado no coincide con la fórmula de Excel. `construir_medidores()` ya ordena con `sort_values(["clave", "intervalo"])` antes de calcularla; no quitar ese paso ni reordenar después. |
| El orden final de columnas usa `list(LETRA_A_CAMPO.values())` (el dict ya está declarado en el orden correcto de Excel). | Si en el futuro se necesitara reintroducir alguna letra de dos caracteres (AA, AB...) en `LETRA_A_CAMPO`, nunca ordenar sus claves con `sorted()`: "AA" < "B" como texto, lo que rompería el orden real de columnas de Excel. Ya pasó una vez en esta migración (ver `docs/Plan_Traspaso_Python_Balance_BESS.md` §19/§20). |
| La fórmula de `Medidores!V` usa `Diccionario!F` y `Diccionario!G` como alias hacia `Diccionario!E` (columnas 5,6,7 del sheet, índices 4,5,6 en el DataFrame `header=None`) — un mapeo posicional específico, distinto de `construir_homologacion()` (que usa toda la fila, sin posición fija). | No usar `construir_homologacion()` para resolver Ofertas SSCC ni `_mapas_homologacion_fge()` para el SoC: son dos bloques distintos de la misma hoja `Diccionario`, con reglas de lectura distintas. |
| `calcular_s()` no se reinicia por central: sigue siendo "igual a la fila anterior mientras `Ventana` no cambie" incluso cruzando de una central a otra. | Es fiel a la fórmula de Excel (`IF(L3=L2,S2,...)`, sin comparar `G`). Si dos centrales consecutivas terminan/empiezan con la misma `Ventana`, `S` no se reinicia — así es también en la planilla original, no es un bug a corregir. |
| Que un archivo se llame `SOC_2607.csv` (o cualquier nombre que contenga "SOC"+AAMM) no garantiza que sea el archivo de SoC de la etapa Medidores. Ya apareció un CSV con ese patrón de nombre que en realidad era un archivo de pagos/liquidación (columnas `Fecha_Hora, CONFIGURACION, Central, Pago, Tipo_pago, Bloque_15min`, sin ninguna columna de SoC), sin relación con `Medidores!J`. | El archivo de SoC real siempre es `.xlsx`, con la estructura de bloques horizontales `Status/Questionable/Time Stamp/Value` (ver `extraer_soc()`). Si un archivo que matchea el patrón de nombre no tiene esa estructura, **no asumir que el formato cambió**: es señal de que no es el archivo correcto. Preguntar antes de adaptar el parser a una estructura nueva. |
| El `Centrales.xlsx` real trae, en la hoja `Resumen BESS`, un título fusionado en la primera fila (`"Cuadro N° 1: Resumen BESS"`) **antes** de la fila de encabezados reales. Un primer intento leyó la hoja con `pd.read_excel(header=0)` (posición fija) y `construir_mapa_barra()` fallaba: no encontraba `'Nombre activo'`/`'Barra inyección'` porque esas columnas venían como `Unnamed: N`. | `leer_centrales()` ahora usa `_leer_resumen_bess()`, que detecta la fila de encabezados buscando los textos esperados (mismo criterio que `detectar_fila_nombres()` para el SoC), nunca por posición fija. Si en el futuro aparece otra hoja de `Centrales.xlsx` con un título similar, aplicar el mismo patrón, no asumir `header=0`. |
| Los nombres de columna de `Calculo E Costos` (`nucleo.construir_calculo_e_costos`) son placeholders (`Mes`, `Dia`, `Hora`, `Hora Mes`, `Minutos`, `Cuarto de Hora`, `clave`, `Barra`, `Energia_Positiva`, `Energia_Negativa`, `SoC`, `Copia_Ventana`, `CMg`) derivados de los comentarios de la macro, no confirmados. El usuario adjuntó dos veces un archivo pensado para traer los encabezados reales de "Ecostos" y ambas veces solo traía las hojas `FD`/`Subastas` (ya confirmadas). | No dar estos nombres por definitivos ni usarlos como referencia para otra hoja. Corregirlos apenas llegue el archivo correcto, sin tocar la lógica de cálculo ya implementada (plan §25.4). |
| `"OfertasSSCC"` tiene **tres** "s" seguidas al pasarlo a minúsculas (`"Ofertas"` termina en "s" + `"SSCC"` empieza con dos "s" más = `"...tas" + "sscc"` = `"...tasssc c"`). Un primer intento transcribió el literal a mano con solo dos "s" (`"ofertasscc"`) y `buscar_archivo_ofertas()` nunca encontraba ningún archivo real. | No transcribir a mano un literal derivado de un nombre con letras dobles/triples repetidas: calcularlo en tiempo de ejecución (`"OfertasSSCC".lower()`, constante `PATRON_NOMBRE_OFERTAS` en `nucleo.py`) y comparar contra eso. Se detectó con un test sintético antes de llegar a producción; si vuelve a fallar la detección del archivo de Ofertas, este es el primer sospechoso a descartar. |

---

## 8. Decisiones ya tomadas

Lista de solo agregar, para no volver a discutir lo mismo en cada sesión.

- **Primero replicar fielmente la lógica de la planilla 11; después
  simplificar.** No se reinterpretan reglas de negocio todavía no
  documentadas solo porque parezcan mejorables (plan §2, principio 1).
- **La carpeta base es la única selección manual del usuario.** Todo lo
  demás (`Medidas_SAE.xlsx`, `SOC_AAMM.xlsx`, `Centrales.xlsx`, la salida)
  se resuelve por ruta relativa desde ahí, para que mover
  `Balance_BESS.py` a otra ubicación no cambie qué entradas encuentra ni
  dónde escribe (criterio de portabilidad, plan §3.13).
- **Las columnas cuya regla no está confirmada se dejan pendientes
  (`pd.NA`) en vez de adivinarse.** Aplica hoy a `K, M, P, Q, R, S, T` de
  `Medidores`. `OfertasSSCC` en particular ni siquiera tiene ubicación de
  carpeta definida todavía (plan §3.6, §7).
- **El SoC se extrae por semántica de encabezados, nunca por posición
  fija.** Confirmado en el plan §16.1: ni letras de columna ni offsets
  constantes entre el nombre de la central y `Time Stamp`/`Value`, porque
  esa distancia varía entre archivos mensuales.
- **`Consolidado_entradas.xlsx` (antes `Hoja_Medidas.xlsx`) es un artefacto
  de validación, no un archivo versionado.** Se genera por caso en la
  carpeta base del usuario y se ignora en git (ver `.gitignore`); el
  repositorio no guarda salidas de casos concretos. Se renombró porque ya
  no es solo la etapa Medidores: consolida varias entradas materializadas
  (Medidores, CMg, FD, Subastas) que alimentan las siguientes etapas del
  balance.
- **`V, W, X, Y, AB, AC, AD, AE` no son columnas de `Medidores` en Python.**
  Las fórmulas de Excel (`V3:V312`, no `V3:V26786`) muestran que son tablas
  auxiliares de otro largo que solo comparten letra de columna con
  `Medidores` porque ahí había espacio libre en la planilla. Forzarlas a
  columnas `pd.NA` del mismo largo que A:U (como se hizo antes de tener el
  código VBA) ya no es una aproximación razonable una vez que se pueden
  calcular de verdad: se escriben juntas en la hoja auxiliar
  `HOJA_OFERTAS_SSCC` (ver plan §20.1, §22). El resumen "Resumen Ofertas
  SSCC" del `.xlsm` original tampoco se persiste — es un paso intermedio
  que solo hace falta en memoria para construir la tabla W:Y.
- **El período AAMM lo escribe el usuario, no se adivina del nombre de un
  archivo.** `SOC_AAMM.xlsx` era solo un patrón conceptual en el plan
  original; en la práctica el archivo de SoC llega con nombres variables.
  Confiar en un regex sobre el nombre para extraer el período era frágil;
  pedirlo explícitamente en la ventana es la fuente de verdad y además
  sirve para validar el archivo de SoC encontrado (debe contener ese AAMM).
- **`Calculo E Costos` vive en un archivo separado (`Pagos_BESS.xlsx`,
  nombre provisorio), no en `Consolidado_entradas.xlsx`.** Pedido explícito
  del usuario. Y se implementa **por etapas**: primero H (Barra) + CMg +
  traspaso base desde Medidores (elegido explícitamente por el usuario
  frente a la alternativa de traducir de una sola vez toda
  `Actualizar_Calculos_Columnas`, ~1500 líneas con dependencias profundas);
  el resto de columnas y `Calculo RE545` quedan para una etapa posterior
  (plan §25).
