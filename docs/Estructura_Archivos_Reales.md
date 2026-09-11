# Estructura de los archivos reales — referencia rápida

Este documento existe para que una sesión nueva (sin el historial de esta conversación, sin los
archivos reales adjuntos de nuevo) entienda **a la primera** qué archivo Excel es cada cosa, qué
estructura real tiene y qué trampas ya conocemos — sin tener que volver a subirlos ni volver a
explicar lo mismo. Si el usuario dice "te mando el Centrales" o pega una hoja de comparación real,
la sesión debería poder actuar directo con esto, sin preguntar "¿cómo viene estructurado?".

Convención de esta página:

- ✅ **Confirmado contra archivo real**: se leyó un archivo real (guardado en `docs/`, o subido y
  ya usado para corregir código) y esto es lo que trae, letra por letra.
- ⚠️ **Del código, no de un archivo real visto**: así está implementado y viene funcionando, pero
  nadie confirmó esta parte específica contra un archivo real todavía.
- Cada archivo real que tenemos guardado en `docs/` está listado con su ruta exacta — se puede
  abrir directo con `openpyxl`/`pandas` en vez de pedírselo de nuevo al usuario.

---

## A. Archivos de entrada del proceso (arma la carpeta del caso)

Son los que el usuario deja en `<CARPETA_BASE>/...` (ver estructura completa en `README.md`).
Ninguno se selecciona a mano — el programa los encuentra por carpeta + patrón de nombre.

### 1. `Medidas/Medidas_SAE.xlsx`

- **Nombre**: literal, `Medidas_SAE.xlsx`.
- **Hoja que lee el programa**: `"Medidas"` (constante `HOJA_MEDIDAS_SAE`).
- **Encabezados**: fila 1, nombres reales (se lee con `pandas.read_excel` normal, sin offset).
- **Columnas que usa** (`COLUMNAS_AI`, en este orden — de ahí sale `Medidores!A:I`):

  | Columna real | Contenido |
  |---|---|
  | `Mes` | mes (número) |
  | `Dia` | día del mes |
  | `Hora` | hora del día |
  | `Minutos` | minuto del bloque de 15' (0/15/30/45) |
  | `Hora Mes` | hora acumulada del mes |
  | `Cuarto de Hora` | bloque horario (1-4) |
  | `clave` | nombre de la central (formato `SAE-...`) |
  | `intervalo` | fecha/hora del registro (se parsea con `pd.to_datetime`) |
  | `Gen_Unidad` | energía del intervalo (kWh, con signo) |

- **Trampa**: si falta alguna de esas 9 columnas, `leer_medidas_sae()` corta con `ErrorEntrada`
  listando cuáles faltan — no intenta adivinar nombres parecidos.
- **Función que lo lee**: `leer_medidas_sae(ruta)`.
- **Archivo real de referencia**: no tenemos una copia guardada en `docs/` todavía. ⚠️ estructura
  confirmada por funcionamiento del código, no por inspección directa de un archivo real reciente.

---

### 2. `Medidas/SOC_AAMM.xlsx` (o cualquier nombre que contenga "SOC" + el AAMM)

- **Patrón**: cualquier `.xlsx` en `Medidas/` cuyo nombre contenga `SOC` y el período (`AAMM`,
  ej. `2607`) que el usuario tipeó en la ventana. Si hay más de uno que matchea, el programa se
  detiene (no elige por fecha de modificación).
- **Hoja**: la primera hoja del archivo (nombre libre, en el archivo real visto es `"Hoja1"`).
- **Estructura real, confirmada** (✅ `docs/SOC_real_2607.xlsx`): son **bloques apilados
  verticalmente**, uno por central, cada uno así:

  ```text
  fila N-2:  <nombre limpio de la central>              <- coincide con Medidas_SAE!clave
  fila N-1:  \\SERVIDOR\SEN\Generación\SEN\<región>\<central>|<sufijo>   <- ruta SCADA completa
  fila N:    (blanco)
  fila N+1:  Status | Questionable | Time Stamp | Value   <- encabezados del bloque
  fila N+2 en adelante: datos (una fila por cuarto de hora)
  ```

  Ejemplo real (columnas E:H del archivo, 1-indexadas):

  ```text
  fila 2: [_, _, _, _, "SAE-CRCA-PFV-DON-HUMBERTO"]
  fila 3: [_, _, _, _, "\\SRV-SCADA-AF2\SEN\Generación\SEN\07 Región RM\SAE-PFV Don Humberto|Nombre"]
  fila 4: (blanco)
  fila 5: [_, _, "Status", "Questionable", "Time Stamp"]   (Value queda una columna más a la derecha)
  fila 6+: datos
  ```

  El siguiente bloque (central siguiente) empieza pegado después del último dato del bloque
  anterior, con el mismo patrón fila-nombre / fila-ruta / blanco / encabezados.

- **Trampas confirmadas (con datos reales, ver `BITACORA.md` sesión "SOC" y "17"):**
  - El nombre de central que hay que usar es el de la fila de **más arriba** del bloque de dos
    filas de metadata (la limpia), NO la fila pegada justo arriba de "Status/Time Stamp" (esa es
    la ruta SCADA completa). `detectar_fila_nombres()` sube desde los encabezados mientras haya
    contenido y se queda con la fila más arriba del bloque contiguo — no con la primera que
    encuentra.
  - El nombre de la fila de más arriba coincide **exacto** con `Medidas_SAE.xlsx!clave` en el
    único caso visto — sospecha (no confirmada aún con un segundo archivo) de que el
    `Diccionario` de `Centrales.xlsx` podría ni hacer falta para homologar el SoC.
  - Si en cambio hace falta homologar (nombre de SoC distinto al de `Medidas_SAE`), se usa
    `construir_homologacion()` sobre `Centrales.xlsx!Diccionario` — ver más abajo la trampa de
    esa hoja.
- **Función que lo lee**: `extraer_soc(...)` / `detectar_fila_nombres()` (`nucleo.py`).
- **Archivo real de referencia**: `docs/SOC_real_2607.xlsx` ✅ (copia completa del real, 1.4 MB).

---

### 3. `Auxiliares/Centrales.xlsx`

Maestro externo, nombre literal. Dos hojas.

#### 3.a. Hoja `Resumen BESS`

- **Estructura real** (✅ `docs/Centrales_real.xlsx`):

  ```text
  fila 1: "Cuadro N° 1: Resumen BESS"   <- título, no es dato
  fila 2: (blanco)
  fila 3: encabezados reales
  fila 4 en adelante: una fila por central BESS (9 centrales en el caso visto)
  ```

- **Columnas reales, en este orden exacto** (`fila 3`, índice VLOOKUP real entre paréntesis —
  importante: el índice del VLOOKUP real **no** es intuitivo, ver la trampa de abajo):

  | # | Columna real | Índice VLOOKUP real | La usa |
  |---|---|---|---|
  | 1 | `Nombre activo` | (la clave) | homologación por nombre |
  | 2 | `Pmax (MW)` | 2 | `AE`/`AF` de E Costos, `BN` de RE545 |
  | 3 | `Horas para descarga forzada` | — | sin uso confirmado todavía |
  | 4 | `Capacidad (MWh)` | 4 | `U`/`BC` de RE545, `BC` de E Costos |
  | 5 | `Energía mínima` | — | sin uso confirmado todavía |
  | 6 | `Barra inyección` | 6 | `H` de las dos hojas de cálculo |
  | 7 | `% Energía sobre mínima (indicador nuevo ciclo)` | — | umbral de ciclo (`M` de E Costos) |
  | 8 | `Ciclos max diarios` | — | sin uso confirmado todavía |
  | 9 | `Eficiencia` | 9 | `V` de RE545 |

- **Trampa confirmada (sesión de corrección, ver `BITACORA.md`):** el índice 4 del VLOOKUP real
  (`Resumen!$B$8:$J$26,4,0)`) es **`Capacidad (MWh)`**, NO `Pmax (MW)` como se asumió al
  principio — `Pmax` es el índice **2**. Si en algún momento hay que agregar un VLOOKUP nuevo
  contra esta hoja, mirar la tabla de arriba, no adivinar por el nombre de la variable.
- **Función que lo lee**: `_leer_resumen_bess()`, `construir_dic_resumen_capacidad()`,
  `construir_dic_resumen_factor()`, `construir_mapa_barra()`.

#### 3.b. Hoja `Diccionario`

- **Estructura real, la trampa más importante de todo el archivo** (✅ `docs/Centrales_real.xlsx`):
  **NO** es "una fila = todos los sinónimos de una central". Son **varias tablas de equivalencia
  independientes, una al lado de la otra**, separadas por columnas que están **enteramente en
  blanco en todas las filas** del archivo. En el caso real:

  ```text
  fila 1 (títulos de grupo): A="FD"     C:D=(blanco, separador)   E="Subastas"   G="ofertas"
  fila 2: (blanco)
  fila 3 en adelante: datos, una fila por central... pero el ORDEN de las filas
                       de cada bloque NO tiene por qué coincidir con el de al lado.
  ```

  Ejemplo real (9 centrales, columnas A,B | E,F,G):

  ```text
  A                              B                        E                               F                               G
  SAE-CRCA-PFV-MANZANO           BESS PFV MANZANO         SAE-CRCA-PFV-MANZANO           SAE-CRCA-PFV-MANZANO
  SAE-CRCA-PFV-DON-HUMBERTO      BESS PFV DON HUMBERTO    SAE-CRCA-PFV-DON-HUMBERTO      SAE-CRCA-PFV-DON-HUMBERTO
  SAE-CRCA-PE-LA-CABANA          BESS PE LA CABAÑA        SAE-CRCA-PE-LA-CABANA          SAE-CRCA-PE-LA-CABANA
  SAE-TOCOPILLA                  SAE TOCOPILLA            SAE-TOCOPILLA                  BAT_TOCOPILLA
  SAE-DEL-DESIERTO                SAE DEL DESIERTO         SAE-DEL-DESIERTO               SAE-DEL-DESIERTO
  SAE-CRCA-PFV-ANDES3            BESS PFV ANDES SOLAR III SAE-CRCA-PFV-ANDES3            SAE-CRCA-PFV-ANDES3
  SAE-CRCA-PFV-VICTOR-JARA       BESS PFV VICTOR JARA     SAE-CRCA-PFV-ANDES4            SAE-CRCA-PFV-ANDES4           BAT_ANDES_4_FV
  SAE-CRCA-PFV-ANDES4            BESS PFV ANDES SOLAR IV  SAE-CRCA-PFV-NUEVO-QUILLAGUA-2  SAE-CRCA-PFV-NUEVO-QUILLAGUA-2
  SAE-CRCA-PFV-NUEVO-QUILLAGUA-2 SAE-CRCA-PFV-NUEVO...-2  SAE-CRCA-PFV-VICTOR-JARA        SAE-CRCA-PFV-VICTOR-JARA
  ```

  Notar la fila 7 (`VICTOR-JARA` en A:B, pero `ANDES4` en E:F): las dos tablas están **corridas
  entre sí** para 3 de las 9 centrales — tratar la fila entera como "un solo grupo de sinónimos"
  homologa mal (mezcla `VICTOR-JARA` con `ANDES4`). Este fue un bug real, ya corregido.
- **Cómo se lee correctamente**: `_bloques_columnas_diccionario()` detecta los bloques de
  columnas automáticamente (separador = columna vacía en TODAS las filas), y cada bloque se
  homologa **por separado**. `construir_homologacion()` (usada por el SoC) ya usa este mecanismo.
  `construir_dic_mapeo_diccionario()` (columnas A:B, usado por el FD homologado de E Costos) y
  `_mapas_homologacion_fge()` (columnas E:F:G, usado por `Medidores!V`) usan posiciones fijas —
  funcionan bien hoy porque coinciden con los bloques reales, pero no se generalizaron a
  `_bloques_columnas_diccionario()` (ver "Pendientes abiertos" en `BITACORA.md`).
- **Trampa adicional**: cada fila puede tener 2 o 3 columnas de sinónimo (no siempre todo el
  bloque), y una celda vacía dentro del bloque significa "no hay sinónimo ahí" — el primer valor
  no vacío de la fila es el nombre "canónico" al que homologan los demás.

---

### 4. `Ofertas/*OfertasSSCC*` (cualquier Excel cuyo nombre contenga "OfertasSSCC")

- **Si hay más de uno**: se toma el más reciente por fecha de modificación (a diferencia del SoC).
- **Estructura**: se leen **TODAS las hojas del archivo** (no una hoja fija), cada una con
  encabezados en la fila 1 pero el programa lee posicional (`header=None`) y arranca en la fila 2.
- **Columnas que usa, por posición** (⚠️ nombres reales de columna NO confirmados contra un
  archivo real — se usa por posición porque así lo hace la macro original):

  | Posición (0-indexada) | Letra | Contenido |
  |---|---|---|
  | 0 | A | Nombre (se filtra por que contenga "BESS"/"SAE"/"BAT") |
  | 1 | B | Año |
  | 2 | C | Mes |
  | 3 | D | Día |
  | 4-6 | E:G | sin uso |
  | 7 | H | Servicio ofertado (se filtra por que termine en `"_RS"`) |
  | 8 | I | Indicador (`"Sí"` = ofertado esa hora) |

- **Función que lo lee**: `construir_resumen_ofertas_sscc(ruta_ofertas)`.
- **Archivo real de referencia**: no tenemos copia guardada en `docs/`.

---

### 5. `Cmg/cmg.xlsx`

- **Nombre**: única excepción con nombre literal fijo (no cambia con el período).
- **Hoja**: `"CMg"` si existe, si no la primera hoja del archivo.
- **Estructura**: columnas A:I (9), encabezados en la fila 1 (se **preservan tal cual**, no se
  renombran — la macro original tampoco los toca). Datos desde la fila 2.
- **Uso**: se ordena por columna D ascendente y luego H ascendente (mismo orden que aplica la
  macro real antes de pegarlo) y se usa un lookup por Barra (columna homologada aparte) + Cuarto
  de Hora.
- **Función que lo lee**: `leer_cmg(ruta_cmg)`.
- **Archivo real de referencia**: no tenemos copia guardada en `docs/`.

---

### 6. `SSCC_Desempeño/SSCC_Desempeño_*.xlsx`

- **Patrón**: cualquier Excel cuyo nombre empiece con `SSCC_Desempeño_`. Si hay más de uno, el
  más reciente por fecha de modificación.
- **Hojas**: `"CPF Horario"` y `"CSF Horario"` (las dos, obligatorias — si falta alguna, error).
- **Encabezados reales**: arrancan más arriba, pero el programa lee **desde la fila 12** de Excel
  (`header=None`, se descartan las primeras 11 filas de título/metadata).
- **Columnas reales que usa, en este orden** (✅ el CONTENIDO/nombre de cada columna fue
  confirmado por el usuario contra un caso real — ver comentario de `NOMBRES_FD_CSF`/
  `NOMBRES_FD_CPF` en `nucleo.py`; ⚠️ el mapeo letra-por-letra de abajo es una reconstrucción a
  partir de esos nombres reales + la lectura posicional del código, no la lectura directa de la
  fila de encabezados del archivo de origen — nadie guardó un `SSCC_Desempeño_*.xlsx` real en
  `docs/` todavía, solo la hoja `FD` de salida ya confirmada, ver abajo):

  **`CSF Horario`, columnas B:H (7)**:

  | Columna real (origen) | Contenido |
  |---|---|
  | B | Fecha |
  | C | Hora |
  | D | Unidad (nombre de central — **se filtra por BESS/SAE acá**) |
  | E | `Respuesta CSF (Fact_CSF)` |
  | F | `Disponibilidad (Fdis_CSF)` |
  | G | `Desempeño (DCSF)` |
  | H | `Factor de Desempeño (Fd_CSF)` |

  **`CPF Horario`, columnas B:J (9)**:

  | Columna real (origen) | Contenido |
  |---|---|
  | B | Fecha |
  | C | Hora |
  | D | Unidad (nombre de central — **se filtra por BESS/SAE acá**) |
  | E | `Respuesta CPF+ (Fact_CPF+)` |
  | F | `Respuesta CPF- (Fact_CPF-)` |
  | G | `Disponibilidad (Fdis_CPF)` |
  | H | `Desempeño (DCPF)` |
  | I | `Factor de Desempeño (Fd_CPF)` |
  | J | `Cuenta con equipo registrador validado` |

- **Salida**: estas columnas alimentan dos bloques de nuestra propia hoja `FD` (`Consolidado_
  entradas.xlsx`) — CSF en `A:M`, CPF en `Q:AE`, con columnas calculadas agregadas (`id`,
  `Hora Mes` duplicada, `CSF(+)`/`CSF(-)`/`CPF(+)`/`CPF(-)` copiadas de la respuesta) — ver
  `NOMBRES_FD_CSF`/`NOMBRES_FD_CPF` en `nucleo.py` para el detalle exacto de esa hoja de salida
  (no confundir con la estructura del archivo de ENTRADA de arriba, que es distinta).
- **Función que lo lee**: `construir_fd(ruta_sscc)`.
- **Archivo real de referencia**: `docs/Libro1_Subastas_real.xlsx`, hoja `"FD"` ✅ (encabezados
  reales confirmados, sin datos de fila).

---

### 7. `Subastas/3_REMUNERACIÓN_SUBASTAS_E_ID_*.xlsx`

- **Patrón**: cualquier Excel cuyo nombre empiece con `3_REMUNERACIÓN_SUBASTAS_E_ID_`. Si hay más
  de uno, el más reciente por fecha de modificación.
- **Hoja**: `"DB"` (constante `HOJA_SUBASTAS_ORIGEN`) — la macro original la consulta por ADO/SQL.
- **Encabezados reales**: fila 2 de Excel (`header=None`, se lee desde la fila 3 = índice 2).
- **Columnas reales que usa, B:Y (24 columnas, filtradas por posición 9 = columna K)** (✅
  confirmadas contra `docs/Libro1_Subastas_real.xlsx`, hoja `"subastas"`, que es una copia/
  reducción real de esta misma hoja `DB`):

  | Columna real | Contenido |
  |---|---|
  | B | `Concepto` (label completo: `CSF(-)`, `CPF(+)`, etc. — fórmula real a partir de C y D) |
  | C | `Control` (tipo SIN dirección: `CSF`/`CTF`/`CPF`) |
  | D | `Sub_Baj` (dirección: `SUBIDA`/`BAJADA`) |
  | E | `Fecha` |
  | F | `Año` |
  | G | `Mes` |
  | H | `Dia` |
  | I | `Hora_dia` |
  | J | `Hora_mes` |
  | K | `Configuración` (nombre de central — **se filtra por BESS/SAE acá**, aunque la macro real
    dice filtrar por "Propietario"; el resultado es el mismo porque toda central BESS empieza
    con `SAE-`) |
  | L | `Propietario` |
  | — | (M:Y se leen pero solo P, Y, V se copian a la salida, ver abajo) |

  De esas 24 columnas (B:Y), la salida (`Subastas!B:Q` de `Consolidado_entradas.xlsx`) usa:
  `B:L` copiadas directo (11 columnas), `M` calculada (`Clave horaria` = `Configuración & Dia &
  Hora_dia`), `N` vacía (`Ciclo`, se calcula después en `Calculo E Costos`), y `O`/`P`/`Q`
  (`Energía SSCC`/`FD`/`FMA`) que son copias de `DB!P`/`DB!Y`/`DB!V` respectivamente — **por
  posición**, no por nombre (los nombres reales de esas 3 columnas en el archivo de origen vienen
  corridos una posición respecto de los títulos de grupo que dicen "Subastas"/"FD"/"FMA"; se
  siguió la fórmula real, no el nombre, y quedó confirmado con un caso real).
- **Trampa histórica ya corregida**: el diccionario `NOMBRES_SUBASTAS` estuvo mal desde el
  principio — le faltaba contar la columna `Concepto` (B) como columna propia, así que todo
  quedaba corrido una posición. La columna `A` real nunca se usa (confirmado: siempre vacía).
- **Función que lo lee**: `construir_subastas(ruta_subastas)`.
- **Archivo real de referencia**: `docs/Libro1_Subastas_real.xlsx`, hoja `"subastas"` ✅
  (encabezados y fórmulas reales, con filas de datos de ejemplo).

---

## B. Archivos de referencia real para VALIDAR la salida (no son entradas del proceso)

Estos no los usa el programa — son los que el usuario manda para comparar nuestra salida contra
la planilla `11_PAGOS_BESS_2607_Definitivo.xlsm` real, fila a fila. Cuando el usuario dice "tengo
diferencias" y pega o adjunta uno de estos, el flujo es: leer con `openpyxl` (`data_only=True`),
armar una clave estable (`Mes+Dia+Hora+Minuto+Configuracion`, o `central+ventana` para tablas
resumen) y comparar celda a celda contra la hoja equivalente de nuestra salida.

| Archivo | Qué trae | Uso |
|---|---|---|
| `docs/Pagos_BESS_comparacion_real.xlsx` | Hojas `Calculo E Costos`, `Calculo RE545` (nuestra salida en ese momento) + `Ecostos planilla 11` (pegada a mano por el usuario, real) | Primera comparación real vs Python que existió en el proyecto — origen del fix de Prorrata SSCC |
| `docs/Calculo_RE545_reducido_para_IA.xlsx` | Hoja `Calculo RE545 reducido` (real, recortada) con **fórmulas** (no solo valores) + hoja `Mapa_Formulas` (rango de celdas → fórmula real, muy útil para confirmar un cálculo sin tener que pedir el `.xlsm` completo) | Fuente directa de varias correcciones de `Calculo RE545` (reservas AR:AT constante 1, cruce S/BI de BK:BL:BS) |
| `docs/Libro1_Subastas_real.xlsx` | Hojas `FD`, `subastas`, `E COSTOS` (con los **merges reales** de Excel, `ws.merged_cells`), `Resumen` | Fuente de `NOMBRES_SUBASTAS`, `NOMBRES_FD_CSF/CPF`, y de `GRUPOS_CALCULO_E_COSTOS` (encabezados de grupo combinados) |
| `docs/Centrales_real.xlsx` | Copia completa del `Centrales.xlsx` real (`Resumen BESS` + `Diccionario`) | Fuente de la corrección de `construir_homologacion()` y del índice VLOOKUP real de `Resumen BESS` |
| `docs/SOC_real_2607.xlsx` | Copia completa de un `SOC_AAMM.xlsx` real | Fuente de la corrección de `detectar_fila_nombres()` (fila 2, no la 3) |

**Cuando el usuario manda un Excel de comparación nuevo** (ej. una hoja pegada tipo
`"<algo> planilla 11"` o `"<HOJA> P11"`): asumir que trae, en alguna hoja, **nuestra salida actual
más una hoja pegada con valores reales** para comparar — leer ambas con `openpyxl`, armar la
clave y diferenciar columna por columna, en vez de asumir que las letras de Excel coinciden entre
nuestra hoja y la real (**nunca coinciden**: nuestra salida no reproduce columnas vacías del
original, así que todo queda corrido de letra — comparar siempre por nombre de columna/clave de
fila, jamás por posición de letra).

---

## C. Qué hacer con un Excel nuevo que te manden

1. Si el nombre/estructura coincide con alguno de los 7 de la sección A → es una entrada real del
   proceso. Compararlo contra la tabla de columnas de esa sección; si difiere, es una corrección
   a documentar acá (y probablemente un bug a corregir en `nucleo.py`).
2. Si trae una hoja "Calculo E Costos"/"Calculo RE545"/similar **más** una hoja pegada con valores
   reales al lado → es una comparación de validación (sección B). Usar el método de comparación
   por clave, no por letra.
3. Si es un archivo nuevo, nunca antes visto en el proyecto (ej. un reporte distinto) → no
   inventar su estructura ni asumir que "debería" tener tal o cual columna: leerlo con
   `openpyxl`/`pandas` primero y preguntar al usuario qué se espera hacer con él, salvo que sea
   evidente por el nombre/contenido.
4. Cualquier archivo real que resulte "fuente de una decisión" (confirma o corrige una estructura,
   un índice, un orden de columnas) se guarda en `docs/` (ver convención ya usada en toda la
   sección A) y se agrega una fila a este documento — así la próxima sesión no tiene que volver a
   pedirlo.
