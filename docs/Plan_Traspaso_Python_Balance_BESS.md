# Plan de traspaso a Python — Balance BESS / SSCC

## 1. Objetivo

Replicar en Python el cálculo actualmente contenido en `11_PAGOS_BESS_2607_Definitivo.xlsm`.

La planilla 11 se utilizará **solo como referencia de validación** durante la migración. El proceso Python no deberá depender de información almacenada exclusivamente dentro de ese libro.

La migración se realizará **hoja por hoja**, manteniendo inicialmente la lógica actual lo más fielmente posible. Aunque existan columnas cuya utilidad todavía no esté completamente entendida, se replicarán mientras formen parte del flujo vigente. La simplificación vendrá después de obtener una reproducción validada.

---

## 2. Principios de diseño

1. **Primero replicar; después simplificar.**
2. La planilla 11 será una referencia, no una entrada del proceso Python.
3. Toda tabla maestra actualmente embebida en el `.xlsm` que sea necesaria para reproducir el cálculo deberá externalizarse.
4. Los valores constantes relevantes deberán convertirse en:
   - parámetros explícitos dentro del código, o
   - entradas externas, según corresponda.
5. Las columnas auxiliares se mantendrán durante la primera versión aunque posteriormente puedan eliminarse.
6. Cada etapa deberá poder compararse contra su equivalente en el Excel original.
7. Las fuentes entregadas externamente se tratarán como **orígenes iniciales** mientras no sea necesario rastrearlas más atrás.

---


# 3. Arquitectura de carpetas y forma de ejecución

## 3.1. Principio general

El proceso se diseñará con una lógica **centralizada y portable**.

El archivo Python podrá estar guardado en cualquier ubicación:

- disco local;
- carpeta de red;
- OneDrive;
- escritorio;
- otra carpeta de scripts.

Su ubicación **no determinará las rutas de trabajo**.

Al ejecutar el Python se abrirá una ventana. Por ahora, la única selección manual será una **carpeta base de trabajo**.

A partir de esa carpeta base, el programa resolverá automáticamente todas las entradas y salidas mediante rutas relativas.

No se deben escribir rutas absolutas dentro del código para los archivos del balance.

---

## 3.2. Flujo esperado para el usuario

La experiencia inicial será:

```text
Ejecutar Python
      ↓
Se abre ventana
      ↓
Seleccionar carpeta base del caso
      ↓
Python revisa estructura esperada
      ↓
Detecta automáticamente entradas
      ↓
Muestra si los archivos requeridos existen
      ↓
Botón Ejecutar
      ↓
Genera Hoja_Medidas en la carpeta base
```

En esta primera etapa no será necesario seleccionar manualmente cada Excel.

La carpeta base será la única ruta elegida por el usuario.

---

## 3.3. Estructura de carpetas definida

La carpeta seleccionada por el usuario tendrá esta estructura:

```text
<CARPETA_BASE_SELECCIONADA>/
│
├── Medidas/
│   ├── Medidas_SAE.xlsx
│   └── SOC_AAMM.xlsx
│
├── Auxiliares/
│   └── Centrales.xlsx
│
├── [archivo OfertasSSCC]              ← ubicación por definir/confirmar
│
└── Hoja_Medidas.xlsx                  ← salida Python
```

Ejemplo para julio de 2026:

```text
Caso_2607/
│
├── Medidas/
│   ├── Medidas_SAE.xlsx
│   └── SOC_2607.xlsx
│
├── Auxiliares/
│   └── Centrales.xlsx
│
└── Hoja_Medidas.xlsx
```

### Regla importante

`<CARPETA_BASE_SELECCIONADA>` representa la carpeta que el usuario selecciona desde la ventana.

El programa no debe asumir dónde está esa carpeta en Windows.

Por ejemplo, cualquiera de estos casos debe funcionar igual:

```text
T:\Balances\2607\
C:\Users\usuario\Desktop\Prueba_2607\
Z:\Procesos\BESS\2607\
C:\Temp\Balance\
```

siempre que dentro exista la estructura requerida.

---

## 3.4. Carpeta `Medidas`

Ruta relativa:

```text
<CARPETA_BASE>/Medidas/
```

Contendrá las entradas relacionadas con medidas y estado de carga.

### Archivos definidos inicialmente

#### `Medidas_SAE.xlsx`

Nombre esperado:

```text
Medidas_SAE.xlsx
```

Ubicación:

```text
<CARPETA_BASE>/Medidas/Medidas_SAE.xlsx
```

Será la entrada equivalente a `Medidores!A:I`.

#### `SOC_AAMM.xlsx`

Patrón esperado:

```text
SOC_AAMM.xlsx
```

Ejemplos:

```text
SOC_2607.xlsx
SOC_2608.xlsx
SOC_2611.xlsx
```

Ubicación:

```text
<CARPETA_BASE>/Medidas/SOC_AAMM.xlsx
```

Será la fuente para construir la columna de SoC equivalente a `Medidores!J`.

### Detección del archivo SOC

El programa no debe depender de un nombre como:

```text
soc julio 26.xlsx
```

La convención futura será:

```text
SOC_AAMM.xlsx
```

El código deberá buscar dentro de `Medidas/` un archivo que cumpla ese patrón.

Idealmente debe existir **un único SOC válido** para el caso seleccionado.

Si existen:

- ninguno → error claro;
- uno → utilizarlo;
- más de uno → detener o solicitar resolución, para evitar escoger silenciosamente un mes equivocado.

No se debe elegir automáticamente por fecha de modificación cuando exista ambigüedad de período.

---

## 3.5. Carpeta `Auxiliares`

Ruta relativa:

```text
<CARPETA_BASE>/Auxiliares/
```

Contendrá maestros y tablas auxiliares necesarias para reproducir el cálculo.

### `Centrales.xlsx`

Ubicación:

```text
<CARPETA_BASE>/Auxiliares/Centrales.xlsx
```

Actualmente contiene:

- hoja `Resumen BESS`;
- hoja `Diccionario`.

El proceso Python deberá buscarlo automáticamente en esa ruta después de seleccionar la carpeta base.

No deberá pedirse al usuario seleccionarlo manualmente mientras se mantenga esta estructura.

---

## 3.6. Ofertas SSCC

`OfertasSSCC` es una entrada externa recibida.

Su ubicación definitiva dentro de la nueva estructura todavía debe confirmarse.

Por ahora queda como **pendiente de diseño de carpeta**, pero la arquitectura deberá respetar el mismo principio:

> el usuario selecciona solo la carpeta base y Python encuentra automáticamente el archivo de OfertasSSCC.

Cuando se defina la ubicación, deberá quedar expresada como ruta relativa a `<CARPETA_BASE>`.

La lógica de lectura deberá replicar:

- `Generar_Resumen_Ofertas_SSCC`;
- `Resumir_Medidores_Central_Ventana_Oferta_Completa`.

---

## 3.7. Salida de la primera etapa

La salida de la réplica de `Medidores` se llamará inicialmente:

```text
Hoja_Medidas.xlsx
```

Ubicación:

```text
<CARPETA_BASE>/Hoja_Medidas.xlsx
```

Es decir, debe quedar directamente en la carpeta seleccionada por el usuario y **no** dentro de `Medidas/` ni `Auxiliares/`.

Ejemplo:

```text
Caso_2607/
├── Medidas/
├── Auxiliares/
└── Hoja_Medidas.xlsx
```

### Rol de la salida

`Hoja_Medidas.xlsx` será inicialmente:

- resultado de la etapa Python;
- archivo de validación contra la hoja `Medidores` de la planilla 11;
- posible entrada de las etapas posteriores mientras se desarrolla el resto del proceso.

Más adelante podrá reemplazarse por un dataframe/parquet u otra estructura interna si conviene, pero durante la migración Excel facilita la comparación.

---

## 3.8. Ventana inicial

La interfaz seguirá el patrón de trabajo centralizado utilizado en otros proyectos.

### Elementos mínimos

La primera versión tendrá:

1. campo que muestre la **Carpeta base**;
2. botón `Examinar`;
3. estado de los archivos detectados;
4. botón `Ejecutar`;
5. zona de mensajes/log;
6. barra de progreso;
7. contador de tiempo.

Por ahora no se necesitan selectores individuales para:

- `Medidas_SAE.xlsx`;
- `SOC_AAMM.xlsx`;
- `Centrales.xlsx`.

Estos archivos deberán detectarse automáticamente desde la carpeta base.

---

## 3.9. Persistencia de la última ruta

La ventana podrá recordar la última carpeta seleccionada mediante un:

```text
config.json
```

ubicado junto al Python.

Este `config.json` pertenece a la herramienta, no al caso de balance.

Conceptualmente:

```text
<carpeta donde vive el Python>/
├── Balance_BESS.py
└── config.json
```

Mientras que los datos pueden estar en cualquier otra ubicación:

```text
T:\...\2607\
├── Medidas/
├── Auxiliares/
└── Hoja_Medidas.xlsx
```

Esto mantiene separadas:

- la ubicación del código;
- la ubicación de los datos.

Si el script se mueve a otra carpeta, el proceso seguirá funcionando seleccionando nuevamente la carpeta base.

---

## 3.10. Resolución de rutas

Después de seleccionar:

```text
base = <carpeta elegida>
```

conceptualmente se resolverán:

```text
medidas_dir      = base / "Medidas"
auxiliares_dir   = base / "Auxiliares"

medidas_sae      = medidas_dir / "Medidas_SAE.xlsx"
soc              = buscar SOC_AAMM dentro de medidas_dir
centrales        = auxiliares_dir / "Centrales.xlsx"

salida_medidores = base / "Hoja_Medidas.xlsx"
```

Estas rutas se derivan de la carpeta base.

No deberán depender de:

```text
Path(__file__).parent
```

salvo para archivos propios de configuración del programa como `config.json`.

---

## 3.11. Validación automática al seleccionar la carpeta

Al seleccionar una carpeta base, la ventana deberá revisar automáticamente:

### Estructura

- existe `Medidas/`;
- existe `Auxiliares/`.

### Entradas

- existe `Medidas/Medidas_SAE.xlsx`;
- existe exactamente un `Medidas/SOC_AAMM.xlsx`;
- existe `Auxiliares/Centrales.xlsx`;
- existe OfertasSSCC una vez definida su ubicación.

### Maestros

En `Centrales.xlsx`:

- existe hoja `Resumen BESS`;
- existe hoja `Diccionario`.

### Resultado visual

La ventana deberá mostrar claramente algo equivalente a:

```text
Carpeta base          OK
Medidas/              OK
Medidas_SAE.xlsx      OK
SOC_2607.xlsx         OK
Auxiliares/           OK
Centrales.xlsx        OK
OfertasSSCC           PENDIENTE / OK
```

La interfaz debe permitir detectar el problema **antes** de comenzar el cálculo.

---

## 3.12. Identificación del período

No se dependerá inicialmente de una ruta específica para saber el período.

Una fuente natural para obtener `AAMM` será el nombre:

```text
SOC_AAMM.xlsx
```

Por ejemplo:

```text
SOC_2607.xlsx
```

permite obtener:

```text
AAMM = 2607
```

Antes de fijar esta regla como definitiva deberá comprobarse que:

- `Medidas_SAE.xlsx` corresponda al mismo período;
- OfertasSSCC corresponda al mismo período;
- los datos internos coincidan con `AAMM`.

La validación cruzada del período será preferible a confiar solo en el nombre del archivo.

---

## 3.13. Criterio de portabilidad

Debe cumplirse esta prueba:

```text
Mover Balance_BESS.py a otra carpeta
        ↓
ejecutarlo
        ↓
seleccionar exactamente el mismo caso
        ↓
obtener exactamente el mismo resultado
```

Si mover el `.py` cambia qué entradas encuentra o dónde genera la salida, la arquitectura estará mal implementada.

---

## 3.14. Evolución futura

Aunque hoy solo estamos replicando `Medidores`, la carpeta base debe servir para ir incorporando las demás etapas sin cambiar la forma de trabajo del usuario.

La idea futura será:

```text
Ejecutar herramienta central
        ↓
Seleccionar carpeta del caso
        ↓
La herramienta descubre todas las entradas
        ↓
Ejecutar Medidores
        ↓
Ejecutar siguientes cálculos
        ↓
Generar resultado final
```

Por eso desde el principio la carpeta seleccionada debe actuar como la **raíz del caso**.

---


# 4. Archivo maestro externo: `centrales.xlsx`

Se creará un archivo externo llamado:

`centrales.xlsx`

Este archivo concentrará, inicialmente, información maestra que hoy vive dentro de la planilla 11.

## 4.1. Estructura definida de `Centrales.xlsx`

El archivo auxiliar ya fue definido y actualmente contiene dos hojas:

- `Resumen BESS`
- `Diccionario`

### A. Hoja `Resumen BESS`

Contiene la tabla maestra de parámetros y barras de cada BESS. La estructura observada es:

| Campo | Uso esperado |
|---|---|
| Nombre activo | Identificador principal del BESS |
| Pmax (MW) | Parámetro de modelado |
| Horas para descarga forzada | Parámetro de modelado |
| Capacidad (MWh) | Parámetro de modelado |
| Energía mínima | Parámetro de modelado |
| Barra inyección | Homologación con CMg / barra eléctrica |
| % Energía sobre mínima (indicador nuevo ciclo) | Parámetro de lógica de ciclos |
| Ciclos max diarios | Parámetro de modelado |
| Eficiencia | Parámetro de modelado |

Esta hoja reemplazará la dependencia respecto de la tabla equivalente que actualmente vive en `Resumen` de la planilla 11.

### B. Hoja `Diccionario`

Contiene las homologaciones entre nombres usados por distintas fuentes. Se conservará inicialmente tal como está construido hoy.

La hoja presenta bloques asociados a:

- `FD`
- `Subastas`
- `ofertas`

En Python no se debe asumir que el mismo nombre aparece idéntico en todas las fuentes; la homologación debe resolverse mediante esta tabla.

**Regla de migración:** durante la primera réplica no se corregirán ni reinterpretarán homologaciones aunque parezcan desplazadas o poco intuitivas. Se usarán exactamente como estén en `Centrales.xlsx` y cualquier inconsistencia se reportará para validación manual.

## 4.2. Rol de `Centrales.xlsx`

`Centrales.xlsx` será un **maestro externo obligatorio** para el proceso Python.

Su función inicial será concentrar:

1. parámetros de modelado de cada BESS;
2. barra de inyección;
3. homologaciones entre nombres de distintas fuentes.

Más adelante se podrá normalizar su estructura, pero no antes de validar la réplica contra la planilla 11.

---

# 5. Primera hoja a migrar: `Medidores`

## 5.1. Objetivo

Construir mediante Python una tabla equivalente a la hoja `Medidores` del libro original.

La primera versión debe reproducir las columnas actuales y su lógica, incluso en aquellos casos donde todavía no esté completamente definida la utilidad posterior de una columna.

---

# 6. Entradas de `Medidores`

## 6.1. Medidas SAE — columnas A:I

### Estado
Origen conocido.

### Fuente actual
`Medidas_SAE.xlsx`, hoja `Medidas`.

### Cadena conocida

`Homologacion ClavesTF y PRMTE.xlsx`
→ `homol.parquet`
→ API de medidas del Coordinador
→ `medidas_batch_*.parquet`
→ script de generación por clave de balance
→ `Medidas_SAE.xlsx`
→ actualmente macro de carga
→ `Medidores!A:I`

### Diseño Python

No es obligatorio mantener `Medidas_SAE.xlsx` como archivo intermedio si la integración futura permite consumir directamente el dataframe generado en la etapa anterior.

Para la primera versión se puede mantener conceptualmente la misma interfaz:

**Entrada:** tabla equivalente a `Medidas_SAE.xlsx / Medidas`.

---

## 6.2. SoC / SCADA — columna J

### Estado
Origen inicial recibido externamente.

### Archivo de referencia revisado
`soc julio 26.xlsx`

### Función dentro de `Medidores`
Alimenta `Medidores!J`.

Posteriormente este valor se traspasa a:

`Calculo RE545!K`

donde corresponde al `SoC %`.

### Estructura observada del archivo

El archivo revisado contiene una única hoja (`Hoja1`) y utiliza bloques horizontales por BESS.

Se identificaron, entre otros, los siguientes bloques:

- `SAE-CRCA-PFV-DON-HUMBERTO`
- `SAE-CRCA-PFV-MANZANO`
- `SAE-TOCOPILLA`
- `SAE-DEL-DESIERTO`
- `SAE-CRCA-PE-LA-CABANA`
- `SAE-CRCA-PFV-ANDES3`
- `SAE-CRCA-PFV-ANDES4`
- `SAE-CRCA-PFV-NUEVO-QUILLAGUA-2`
- `SAE-CRCA-PFV-VICTOR-JARA`

Cada bloque presenta una estructura semejante a:

`Status | Questionable | Time Stamp | Value`

Sin embargo:

- los bloques no comienzan siempre en columnas contiguas;
- existen columnas vacías entre grupos;
- algunos metadatos superiores tienen formatos distintos;
- por lo tanto, **no se debe programar la lectura en función de letras de columnas fijas**.

---

# 6. Método estandarizado propuesto para extraer el SoC

La extracción deberá basarse en la **semántica de los encabezados**, no en posiciones rígidas.

## 6.1. Detección de bloques por coordenadas de la fila de nombres

La posición horizontal del nombre de una central **no es fija** respecto de sus columnas de datos. Por lo tanto, la extracción no se hará buscando una distancia constante entre el nombre del BESS y `Time Stamp` / `Value`.

El método será:

1. identificar la fila que contiene los nombres lógicos de las centrales/BESS;
2. recorrer esa fila y registrar la coordenada de cada celda no vacía que corresponda a una central;
3. ordenar esas coordenadas de izquierda a derecha;
4. definir para cada central un bloque horizontal:
   - inicio = columna donde aparece su nombre;
   - fin = columna inmediatamente anterior al nombre de la siguiente central;
   - para la última central, fin = última columna relevante de la hoja;
5. dentro de cada bloque, buscar los encabezados por texto, no por posición:
   - `Time Stamp`
   - `Value`;
6. guardar las coordenadas reales de esas dos columnas y extraer los datos desde allí.

Ejemplo conceptual:

```text
Fila nombres:
        CENTRAL_A                 CENTRAL_B                    CENTRAL_C
        ↓                         ↓                            ↓
cols 2..6                    cols 7..12                   cols 13..18

Dentro de cada rango se buscan:
Time Stamp | Value
```

De esta forma, el método tolera:

- columnas vacías entre bloques;
- distinto ancho por central;
- que el nombre esté sobre `Status`, `Questionable` u otra columna del bloque;
- desplazamientos entre archivos mensuales;
- incorporación de nuevas centrales sin modificar posiciones hardcodeadas.

## 6.2. Identificación del BESS

El identificador primario del bloque será el texto detectado en la **fila de nombres de centrales**.

Para cada nombre encontrado se almacenará al menos:

- `nombre_bess_origen`;
- `columna_inicio_bloque`;
- `columna_fin_bloque`;
- `columna_timestamp`;
- `columna_value`.

Posteriormente `nombre_bess_origen` se homologará mediante la hoja `Diccionario` de `Centrales.xlsx`.

Si dentro de un bloque no se encuentra exactamente un `Time Stamp` y un `Value`, el proceso no debe adivinar: debe registrar una inconsistencia y excluir ese bloque hasta revisión.

## 6.3. Extracción

Por cada BESS se generará inicialmente una estructura:

| BESS_SCADA | Time Stamp | SoC |
|---|---|---|

Los datos se leerán desde la primera fila posterior al encabezado hasta el último registro válido.

## 6.4. Normalización temporal

El campo `Time Stamp` puede venir como fecha/hora de Excel.

Python deberá convertirlo a un `datetime` explícito.

Se deberá validar:

- fecha;
- hora;
- minuto;
- orden cronológico;
- duplicados;
- huecos temporales.

En el archivo revisado los registros observados presentan una cadencia de **15 minutos**.

La lógica no deberá asumir silenciosamente que siempre existen exactamente 96 registros diarios; deberá detectar y reportar desviaciones.

## 6.5. Normalización del SoC

`Value` deberá convertirse a número.

Los valores observados están expresados como proporción decimal, por ejemplo:

`0.040609`

equivale a:

`4.0609 %`

La primera versión debe mantener el valor en la misma escala que utiliza actualmente el Excel para evitar diferencias durante la validación.

## 6.6. Salida normalizada del SoC

Formato recomendado:

| central | timestamp | soc |
|---|---|---:|
| ... | ... | ... |

Podrán mantenerse además columnas de trazabilidad como:

- `nombre_scada_original`
- `ruta_scada`
- `archivo_origen`

Estas columnas no necesariamente formarán parte del cálculo final, pero son útiles para auditoría.

## 6.7. Validaciones mínimas de entrada

Antes de utilizar los datos:

- debe existir un nombre de BESS para cada par `Time Stamp` / `Value`;
- no deben existir timestamps inválidos;
- `SoC` debe ser numérico;
- detectar timestamps duplicados por BESS;
- detectar intervalos faltantes;
- reportar BESS presentes en el archivo pero no homologados en `centrales.xlsx`;
- reportar BESS esperados por el balance que no aparezcan en el archivo de SoC.

### Conclusión

La ambigüedad visual del archivo es manejable.

El método recomendado es **detección dinámica de bloques por encabezados**, no lectura por columnas fijas.

---

# 7. Ofertas SSCC

## 7.1. Estado

El archivo origen ya está disponible y se recibe externamente.

Por lo tanto se considera una **entrada inicial** del proceso.

## 7.2. Lógica a replicar

Python deberá replicar fielmente la lógica de las macros:

- `Generar_Resumen_Ofertas_SSCC`
- `Resumir_Medidores_Central_Ventana_Oferta_Completa`

La primera versión no intentará reinterpretar ni simplificar esas reglas.

## 7.3. Resultados internos equivalentes

La lógica asociada a Ofertas SSCC actualmente genera información equivalente a las zonas:

- `Medidores!W:Y`
- `Medidores!AB:AE`

En Python estas zonas no tienen por qué existir físicamente como columnas separadas por letras de Excel, pero sus resultados y significado deberán preservarse inicialmente para permitir una comparación directa.

---

# 8. Parámetros fijos de `Medidores`

Los siguientes parámetros se definirán explícitamente en Python.

## 8.1. Inicio de ventana

Valor actual:

`10`

Equivale al valor utilizado actualmente mediante `S1`.

Primera implementación:

`INICIO_VENTANA = 10`

Por ahora permanecerá como parámetro fijo dentro del código.

## 8.2. Umbral SoC

Valor actual:

`6 %`

La lógica actual de la columna O compara el SoC contra este umbral.

Primera implementación conceptual:

`UMBRAL_SOC = 0.06`

Debe conservarse como parámetro explícito dentro del código y no como número escondido dentro de una expresión.

---

# 9. Columnas de `Medidores` a replicar

## 9.1. A:I

Datos provenientes de Medidas SAE.

**Tipo:** entrada.

---

## 9.2. J

Dato de SoC proveniente del archivo externo SCADA.

**Tipo:** entrada.

Debe realizarse el cruce temporal/por central necesario para asociar cada registro de SoC con la fila correcta de `Medidores`.

El detalle exacto de esa unión todavía debe documentarse al revisar las columnas A:I y la correspondencia utilizada actualmente.

---

## 9.3. K

Mantener la lógica actual.

**Estado funcional:** pendiente de documentar en detalle.

No se elimina en la primera versión.

---

## 9.4. L — contador/ventana

Mantener la lógica actual.

Condición inicial:

`L inicial = 0`

Este cero actualmente está escrito explícitamente y en Python deberá convertirse en una inicialización deliberada.

Lógica Excel conocida:

```text
SI(G_actual <> G_anterior;
   0;
   SI(C_actual = C_anterior;
      0;
      SI(C_actual = INICIO_VENTANA;
         1;
         0
      )
   ) + L_anterior
)
```

Donde:

`INICIO_VENTANA = 10`

La traducción a Python deberá respetar exactamente esta secuencia.

---

## 9.5. M

Mantener la lógica actual.

**Estado funcional:** pendiente de documentar en detalle.

---

## 9.6. N — clave auxiliar

Se entiende actualmente como una clave auxiliar.

La lógica deberá replicarse exactamente.

Expresión conocida conceptualmente:

`N = B + "&" + E`

La utilidad final puede reevaluarse una vez validada la réplica.

---

## 9.7. O — indicador según SoC

Mantener la lógica actual.

Umbral:

`6 %`

Conceptualmente:

`O = 1 * (J > 0.06)`

El valor `0.06` deberá quedar definido mediante el parámetro:

`UMBRAL_SOC = 0.06`

y no repetirse como número mágico.

---

## 9.8. P:Q

Mantener las columnas y lógica existentes durante la primera versión.

**Estado funcional:** pendiente de completar.

---

## 9.9. R

Mantener la lógica de búsqueda/cruce asociada a Ofertas SSCC.

La fuente equivalente será generada en Python desde el archivo externo de Ofertas.

---

## 9.10. S

Mantener la lógica actual.

La columna depende, entre otros elementos, de la continuidad de `L` y del resultado de `R`.

No se simplificará inicialmente.

---

## 9.11. T

Mantener la lógica actual.

Está relacionada con la combinación de central, ventana/ciclo y la información resumida de ofertas.

Debe reproducirse contra el resultado equivalente a `AB:AE`.

---

## 9.12. U en adelante

Se mantendrán todas las columnas necesarias para reproducir la hoja actual.

Las columnas cuya función todavía no esté entendida se clasificarán temporalmente como:

`AUXILIAR - FUNCIÓN PENDIENTE DE DOCUMENTAR`

pero no serán eliminadas.

---

# 10. Salida esperada de la etapa `Medidores`

La etapa Python deberá generar un dataframe equivalente, fila a fila, a la hoja `Medidores` después de haber ejecutado las macros y cálculos necesarios.

Durante la fase de validación se recomienda exportar temporalmente:

`debug_medidores.xlsx`

o un archivo equivalente de comparación.

Este archivo sería exclusivamente de control y no necesariamente formará parte del proceso productivo definitivo.

---

# 11. Validación de `Medidores`

La réplica se considerará validada cuando, para un mismo período y las mismas entradas:

1. coincida la cantidad de filas;
2. coincidan las centrales/configuraciones;
3. coincida el orden temporal esperado;
4. coincidan las columnas de entrada;
5. coincidan las columnas calculadas;
6. cualquier diferencia numérica esté dentro de una tolerancia previamente definida;
7. las diferencias sean explicables y estén documentadas.

La comparación deberá realizarse inicialmente contra la hoja `Medidores` de:

`11_PAGOS_BESS_2607_Definitivo.xlsm`

---

# 12. Entradas definidas para la etapa `Medidores`

| Entrada | Estado | Tipo |
|---|---|---|
| Medidas SAE / generación por clave | Definida | Externa / proceso previo |
| SoC SCADA | Definida como origen externo; extracción a estandarizar | Externa |
| OfertasSSCC | Origen disponible | Externa |
| `centrales.xlsx` — Diccionario | A crear | Maestro |
| `centrales.xlsx` — Barras | A crear | Maestro |
| Inicio de ventana = 10 | Definido | Parámetro código |
| Umbral SoC = 6% | Definido | Parámetro código |

---

# 13. Pendientes específicos antes de programar `Medidores`

1. Documentar **todas las columnas actuales de Medidores**, una por una.
2. Confirmar la lógica exacta de K.
3. Confirmar la lógica exacta de M.
4. Confirmar la lógica exacta de P y Q.
5. Revisar el resto de columnas U:AE y establecer cuáles son:
   - entradas;
   - auxiliares;
   - resultados de ofertas;
   - claves;
   - indicadores.
6. Definir exactamente cómo se cruza el SoC externo con cada fila de Medidores:
   - central;
   - fecha;
   - hora;
   - cuarto de hora/minuto.
7. Definir la estructura definitiva de `centrales.xlsx`.
8. Documentar completamente la traducción funcional de:
   - `Generar_Resumen_Ofertas_SSCC`
   - `Resumir_Medidores_Central_Ventana_Oferta_Completa`
9. Definir nombre/formato de la salida Python de la etapa `Medidores`.
10. Crear casos de prueba para comparar Python vs Excel.

---

# 14. Estrategia general de migración

La secuencia será:

```text
Entradas iniciales
        ↓
Normalización de maestros y archivos externos
        ↓
MEDIDORES
        ↓
Validación contra Excel
        ↓
Siguiente hoja
        ↓
Validación
        ↓
...
        ↓
Resultado final del balance
```

No se avanzará a simplificar la arquitectura hasta disponer de una réplica funcional completa o suficientemente validada.

---

# 15. Estado actual

## Hoja en análisis
`Medidores`

## Estado
**Diseño de entradas/salidas y lógica. Sin implementación todavía.**

## Próximo paso recomendado
Completar el inventario columna por columna de `Medidores`, utilizando la planilla 11 únicamente como referencia, y cerrar las reglas exactas antes de comenzar a programar.


---

# 16. Actualización de diseño — extracción SoC y maestro de centrales

## 16.1. Decisión confirmada para SoC

La extracción del archivo SoC se basará en las **coordenadas detectadas en la fila de nombres de centrales**. Los encabezados `Time Stamp` y `Value` se buscarán dentro del rango horizontal correspondiente a cada central.

No se utilizarán letras de columna ni offsets fijos.

## 16.2. Archivo auxiliar confirmado

Archivo: `Centrales.xlsx`

Hojas actuales:

- `Resumen BESS`: parámetros de modelado y barra de inyección.
- `Diccionario`: homologaciones entre nombres utilizados por FD, Subastas y Ofertas.

Este archivo pasa a formar parte de las entradas maestras oficiales del nuevo proceso Python.


# 16. Especificación cerrada de columnas de `Medidores`

Esta sección consolida las reglas confirmadas para replicar la hoja `Medidores` en Python.

## 16.1. Regla de cruce para SoC

La columna `J` se alimentará desde `SOC_AAMM.xlsx`.

La asociación entre cada registro de SoC y cada fila de `Medidores` se realizará usando:

- central;
- fecha;
- hora;
- cuarto de hora.

La extracción del archivo `SOC_AAMM.xlsx` seguirá la lógica ya definida de detección dinámica de bloques por la fila de nombres de centrales y búsqueda de encabezados `Time Stamp` / `Value` dentro de cada bloque.

---

## 16.2. Ubicación de Ofertas SSCC

La entrada externa de Ofertas SSCC se ubicará en:

```text
<CARPETA_BASE>/Ofertas/
```

El programa deberá buscar ahí el archivo correspondiente al período.

---

## 16.3. Regla exacta por columna

| Columna | Estado | Regla / fuente |
|---|---|---|
| A:I | Copiadas | Provienen de `Medidas/Medidas_SAE.xlsx`, hoja `Medidas`. |
| J | Copiada / cruzada | SoC desde `Medidas/SOC_AAMM.xlsx`, asociado por central + fecha + hora + cuarto de hora. |
| K | Calculada | Copia de `L` fila a fila. |
| L | Calculada | Contador/ventana. Parte con `L inicial = 0`. Mantener la lógica actual usando `INICIO_VENTANA = 10`. |
| M | Vacía | Debe quedar vacía. |
| N | Calculada | Clave auxiliar. `N = B & "&" & E`. |
| O | Calculada | Indicador por SoC con `UMBRAL_SOC = 0.06`. Conceptualmente `O = 1 * (J > 0.06)`. |
| P | Vacía | Debe quedar vacía. |
| Q | Vacía | Debe quedar vacía. |
| R | Calculada | Resultado de búsqueda contra el resumen de Ofertas SSCC generado en Python. |
| S | Calculada | Mantener lógica secuencial actual dependiente de L y R. |
| T | Calculada | Mantener lógica actual contra el resumen central + ventana + oferta completa. |
| U | Vacía | Debe quedar vacía. |
| V | Calculada | ID auxiliar. Fórmula de referencia: `=X3&BUSCARX(W3;Diccionario!F:F;Diccionario!E:E;BUSCARX(W3;Diccionario!G:G;Diccionario!E:E;0;0);0)` |
| W | Calculada por lógica de Ofertas | Central. |
| X | Calculada por lógica de Ofertas | Día. |
| Y | Calculada por lógica de Ofertas | Oferta completa: bandera de si tiene o no oferta completa. |
| Z | Vacía | Debe quedar vacía. |
| AA | Vacía | Debe quedar vacía. |
| AB | Calculada por lógica de Ofertas | Central. |
| AC | Calculada por lógica de Ofertas | Ventana T. |
| AD | Calculada por lógica de Ofertas | Oferta. |
| AE | Calculada por lógica de Ofertas | Completa. |

---

# 17. Traspaso obligatorio de macros de Ofertas SSCC

Las macros de Ofertas SSCC **sí forman parte obligatoria de la réplica en Python**.

No deben reemplazarse por columnas vacías ni por placeholders `_PENDIENTE`.

## 17.1. `Generar_Resumen_Ofertas_SSCC`

Python deberá reproducir fielmente la lógica de:

```text
Generar_Resumen_Ofertas_SSCC
```

Su salida corresponde a:

| Columna | Significado |
|---|---|
| W | Central |
| X | Día |
| Y | Oferta completa |

---

## 17.2. `Resumir_Medidores_Central_Ventana_Oferta_Completa`

Python deberá reproducir fielmente la lógica de:

```text
Resumir_Medidores_Central_Ventana_Oferta_Completa
```

Su salida corresponde a:

| Columna | Significado |
|---|---|
| AB | Central |
| AC | Ventana T |
| AD | Oferta |
| AE | Completa |

Estas columnas alimentan posteriormente la lógica de `T`.

---

# 18. Regla de réplica fiel

Para esta primera versión:

- no eliminar columnas auxiliares;
- no simplificar macros;
- no reemplazar lógica conocida por placeholders;
- replicar primero el comportamiento del Excel;
- validar contra la hoja `Medidores` original;
- simplificar únicamente después de lograr equivalencia.

---

# 19. Nota de implementación — período AAMM y columnas de Ofertas SSCC (sesión de organización del repo)

## 19.1. AAMM como entrada explícita del usuario

`SOC_AAMM.xlsx` es un patrón conceptual (§3.4), no un nombre de archivo literal: en la práctica el archivo de SoC puede llegar con otro nombre siempre que contenga `SOC` y el período. Por eso el período **AAMM ya no se infiere del nombre del archivo**: la ventana pide al usuario ingresarlo en un campo de texto (4 dígitos, ej. `2607`), y ese valor es la fuente de verdad para:

- calcular año/mes (`periodo_desde_aamm`);
- buscar el archivo de SoC dentro de `Medidas/` (debe contener `SOC` y el AAMM ingresado en el nombre, en cualquier posición/separador — ya no una regex estricta de la forma `SOC_AAMM.xlsx`).

Si el AAMM no se ingresa o no son 4 dígitos, el checklist de la ventana lo marca como `FALTA` y bloquea la ejecución.

## 19.2. Qué se implementó de §16.3 y qué sigue pendiente

Implementado en `nucleo.py` en esta sesión:

- `K` = copia de `L` fila a fila (`Copia_Ventana`).
- `M`, `P`, `Q`, `U`, `Z`, `AA` quedan explícitamente vacías (`pd.NA`) — es diseño confirmado, no una tarea pendiente.

Sigue pendiente, porque el texto de las macros VBA (`Generar_Resumen_Ofertas_SSCC` y
`Resumir_Medidores_Central_Ventana_Oferta_Completa`) todavía no fue entregado a esta sesión —
solo se conoce el nombre de las macros y qué columnas producen, no la lógica interna —, y por
lo tanto no se pueden replicar fielmente sin adivinar (viola la regla de réplica fiel, §18):

- `R`, `S`, `T` (dependen del resumen de Ofertas SSCC);
- `V` (la fórmula de Excel es conocida, pero depende de `W`/`X`, que son salida de una macro pendiente, y de la disposición exacta por columnas de la hoja `Diccionario`, todavía no confirmada);
- `W`, `X`, `Y` (salida de `Generar_Resumen_Ofertas_SSCC`);
- `AB`, `AC`, `AD`, `AE` (salida de `Resumir_Medidores_Central_Ventana_Oferta_Completa`).

Estas columnas quedan en `Hoja_Medidas.xlsx` como `pd.NA`, con nombre de campo terminado en
`_PENDIENTE_OFERTAS`, hasta que se entregue el código fuente de ambas macros. La carpeta
`Ofertas/` bajo `<CARPETA_BASE>` ya se detecta en el checklist de la ventana (no bloqueante),
pero todavía no se define el patrón de nombre del archivo dentro de ella ni se lee su contenido.
