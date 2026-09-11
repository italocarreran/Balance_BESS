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

---

# 20. Implementación de Ofertas SSCC (sesión con el código VBA fuente)

Se recibió `Trazabilidad_11_PAGOS_BESS_2607_Definitivo.md`, que incluye el código VBA completo
de `Generar_Resumen_Ofertas_SSCC` y `Resumir_Medidores_Central_Ventana_Oferta_Completa`
(extraído de `xl/vbaProject.bin`) y, en su sección 5.1, las fórmulas de Excel de `Medidores!K`,
`L`, `N`, `O`, `R`, `S`, `T` y `V`. Esto resuelve el pendiente de la sección 19.2: ya no falta
el código fuente de las macros.

## 20.1. Hallazgo estructural: V, W, X, Y, AB, AC, AD, AE no son columnas por fila

Las fórmulas de la sección 5.1 muestran que `V3:V312` (no `V3:V26786` como el resto de las
columnas de `Medidores`) — es decir, `V` (y `W`, `X`, `Y`, que la macro `H_Leer_Ofertas` escribe
directamente como valores, sin fórmula) solo ocupan tantas filas como necesite la tabla
auxiliar "central × día del mes" que arma `OSSCC_CargarResumenEnMedidores`, no una fila por
registro de `Medidores`. Lo mismo ocurre con `AB:AE`, que ocupan tantas filas como grupos
"central × ventana" existan (los escribe `Resumir_Medidores_Central_Ventana_Oferta_Completa`).
En la planilla original conviven en las mismas letras de columna que el resto de `Medidores`
porque ahí había espacio libre, no porque compartan el mismo "largo" conceptual.

**Decisión de diseño:** en Python, `V`, `W`, `X`, `Y`, `AB`, `AC`, `AD`, `AE` ya no se
representan como columnas de `pd.NA` del mismo largo que `A:U` (eso nunca fue fiel, aunque
serviía como marcador de "pendiente"). Se calculan como tablas auxiliares de su propio largo y
se escriben como hojas separadas de `Hoja_Medidas.xlsx`:

- **"Resumen Ofertas SSCC"** — salida de `Generar_Resumen_Ofertas_SSCC` (Nombre, Año, Mes, Día,
  una columna por servicio `_RS` encontrado, Oferta completa). Mismo nombre que la hoja
  homónima del `.xlsm` original, para comparación directa.
- **"Ofertas SSCC por Dia"** — equivalente a `Medidores!W:Y` (Nombre, Dia, Oferta completa),
  una fila por central × día del mes.
- **"Resumen Ventana Oferta"** — equivalente a `Medidores!AB:AE` (Central, Ventana T, Oferta,
  Completa), una fila por grupo central × ventana.

`V` en sí (la clave de homologación día+central usada solo para resolver `R`) no se persiste
como tabla propia: es un paso intermedio interno de `nucleo.calcular_r()`.

## 20.2. Columnas R, S, T — ahora calculadas

`R`, `S` y `T` SÍ son columnas por fila (sus rangos de fórmula cubren todo `Medidores`) y se
implementaron tal cual en `nucleo.py`:

- **R** (`Oferta_Completa_Dia`) = `VLOOKUP(B&G, V:Y, 4, FALSE)`: para cada fila, busca (Día +
  central homologada vía `Diccionario!F/G→E`) en la tabla "Ofertas SSCC por Dia" y devuelve la
  Oferta completa de ese día.
- **S** (`Indicador_Ventana_Oferta`) = `IF(L=L_anterior, S_anterior, IF(R=1,1,2))`: se mantiene
  mientras la Ventana no cambie; al cambiar, toma 1 si R=1 en esa fila, si no 2. No se reinicia
  aparte por central — la fórmula original tampoco lo hace, se apoya en que L ya cambia al
  cambiar de central.
- **T** (`Ventana_No_Completa`) = `1 - Completa(central=G, ventana=L)`, buscando en "Resumen
  Ventana Oferta" (T=0 → ventana completa, T=1 → incompleta).

## 20.3. Verificación

Se armó un caso sintético (1 central, 2 días, `Hora` en convención 1-24 como usa la planilla
real) con un archivo de OfertasSSCC donde la central ofertó las 24 horas del servicio `_RS`
ambos días. Resultado: las 3 ventanas del caso (inicio, normal, última) dieron exactamente 36,
96 y 60 filas — los valores esperados de la fórmula
`(INICIO_VENTANA-1)*4` / `96` / `(25-INICIO_VENTANA)*4` — y las tres quedaron marcadas
`Completa=1`; R=1 y T=0 en las 192 filas. Un segundo caso con una hora "No" en vez de "Sí"
marcó correctamente `Oferta completa=0` para ese grupo. No se validó todavía contra un caso
real ni contra la hoja `Medidores` de la planilla 11 (sigue sin datos reales disponibles).

## 20.4. Lo que sigue pendiente

- Validar contra un caso real y contra la planilla 11 (plan §13, punto 10).
- Confirmar que la columna `V` no necesita persistirse (por ahora es puramente interna a
  `calcular_r`); si en la validación contra Excel hiciera falta auditarla fila a fila, agregarla
  a la hoja "Ofertas SSCC por Dia".
- `U:AE` fuera de lo ya cubierto por `R,S,T,V,W,X,Y,AB,AC,AD,AE` no aplica: esas eran todas las
  columnas pendientes identificadas hasta ahora (plan §9.12 original).

---

# 21. El archivo de SoC — corrección: sigue siendo siempre `.xlsx`

Una sesión anterior agregó soporte para que el archivo de SoC llegara como `.csv`, a partir de
una aclaración del usuario que resultó estar equivocada (confundió el archivo de SoC con otro
CSV de pagos/liquidación que por coincidencia también tenía "SOC" y el AAMM en el nombre —
columnas `Fecha_Hora, CONFIGURACION, Central, Pago, Tipo_pago, Bloque_15min`, sin ninguna
columna de SoC). El usuario confirmó después que el archivo de SoC real siempre es `.xlsx`.

Se revirtió el soporte de `.csv`: `nucleo.buscar_soc()` vuelve a exigir `.xlsx`, y se eliminó
`leer_soc_crudo()` (la sección 20/19 de este documento ya describía correctamente la estructura
de bloques horizontales por central — eso no cambió, solo el contenedor sí volvió a ser
únicamente Excel). Ver `METODOLOGIA.md` §7 (trampas conocidas) para la regla que evita repetir
esta confusión: un archivo que matchea el patrón de nombre pero no tiene la estructura de
bloques esperada no es el archivo de SoC, aunque comparta "SOC"+AAMM en el nombre.

---

# 22. Ajuste de las hojas auxiliares de Ofertas SSCC

Tras usar `Hoja_Medidas.xlsx` en la práctica, se ajustó cómo se exponen las tablas auxiliares
descritas en la sección 20.1:

- **"Resumen Ofertas SSCC"** (salida de `Generar_Resumen_Ofertas_SSCC`: Nombre, Año, Mes, Día,
  una columna por servicio `_RS`, Oferta completa) deja de escribirse como hoja. Es puramente
  auxiliar — solo existe para construir la tabla equivalente a `Medidores!W:Y` — y no aporta
  valor de validación por sí sola. Sigue calculándose en memoria dentro de
  `construir_medidores()` (función `construir_resumen_ofertas_sscc()`, sin cambios), solo que ya
  no se persiste ni se devuelve fuera de esa función.
- **"Ofertas SSCC por Dia"** (equivalente a `Medidores!W:Y`) y **"Resumen Ventana Oferta"**
  (equivalente a `Medidores!AB:AE`) se unen en una sola hoja, `"Ofertas SSCC"`
  (`nucleo.HOJA_OFERTAS_SSCC`), una tabla debajo de la otra, cada una con un título en negrita
  arriba (`_escribir_tabla_con_titulo()`) para distinguirlas al abrir el archivo.

`Hoja_Medidas.xlsx` queda entonces con tres hojas: `Medidores`, `Ofertas SSCC`, `Log`.

---

# 23. Nuevas hojas: CMg, FD, Subastas — y cambio de nombre del archivo de salida

El archivo de salida deja de llamarse `Hoja_Medidas.xlsx` y pasa a llamarse
**`Consolidado_entradas.xlsx`**, reflejando que ya no es solo la etapa Medidores: reúne varias
entradas materializadas (Medidores, CMg, FD, Subastas) que las siguientes etapas del balance
(`Calculo E Costos`, `Calculo RE545`) van a consumir.

Se agregan tres hojas, replicando las macros de carga (no las que las consumen después, como
`Asignar_CMg_a_Calculos_Turbo` o `Actualizar_Calculos_Columnas`, que pertenecen a una etapa
posterior todavía sin implementar):

## 23.1. `CMg` — replica `Cargar_CMg_Desde_Archivo`

- **Entrada:** `<CARPETA_BASE>/Cmg/cmg.xlsx` — a diferencia de todos los demás archivos
  externos de este proyecto, el nombre es **literal y fijo** (así lo exige la macro original;
  no sigue un patrón con AAMM).
- **Lectura:** hoja `CMg` si existe, si no la primera hoja del archivo; columnas A:I completas,
  con su encabezado real (no se le inventa nombre).
- **Transformación:** se ordena por columna D ascendente y luego por columna H ascendente —
  igual que hace la macro sobre el origen antes de pegarlo en la hoja `CMg` del balance.
- **No implementado (fuera de alcance de esta sesión):** `Asignar_CMg_a_Calculos_Turbo`, que
  cruza esta hoja contra `Calculo E Costos`/`Calculo RE545` (columnas D, F, H, I) — pertenece a
  una etapa posterior.

## 23.2. `FD` — replica `Cargar_SSCC_Desempeno_En_FD`

- **Entrada:** `<CARPETA_BASE>/SSCC_Desempeño/SSCC_Desempeño_*.xlsx` (o `.xlsm`/`.xlsb`/`.xls`) —
  el más reciente por fecha de modificación si hay varios, igual que la macro original
  (`BuscarArchivoSSCCMasReciente`).
- **Lectura:** hojas `CPF Horario` y `CSF Horario` del archivo, datos desde la fila 12,
  filtrando las filas cuya columna D (dentro de cada bloque) contenga "BESS" o "SAE" — **sin**
  "BAT" (a diferencia del filtro de Ofertas SSCC; son dos filtros distintos aunque se parezcan).
- **Estructura resultante — hallazgo clave:** igual que con V/W/X/Y/AB/AC/AD/AE de `Medidores`,
  los bloques CSF y CPF de `FD` son **dos tablas independientes de distinto largo** que
  comparten la misma hoja en rangos de columnas distintos (A:M para CSF, Q:AE para CPF; N:P
  quedan vacías/fuera de alcance porque la macro no las toca). Se escriben lado a lado en la
  hoja `FD`, cada una con su propio largo de filas.
- **Fórmulas replicadas** (columnas calculadas, exactas según la sección de fórmulas del
  libro): para el bloque CSF (destino D:J = copia directa de `CSF Horario!B:H` filtrado):
  - `A = str(B) & F`
  - `B = (DAY(D)-1)*24 + E + 1 + IF(DAY(D)>100,1,0)` (fórmula "Hora Mes"; el término
    `DAY(D)>100` nunca es cierto para un día real, se conserva tal cual sin "corregirlo")
  - `C = DAY(D)`
  - `K = J`, `L = K`, `M = B`

  Para el bloque CPF (destino T:AB = copia directa de `CPF Horario!B:J` filtrado):
  - `Q = str(R) & V`
  - `R = (DAY(T)-1)*24 + U + 1 + IF(DAY(T)>100,1,0)`
  - `S = DAY(T)`
  - `AC = AA`, `AD = AC`, `AE = R`
- **Columnas sin nombre de negocio documentado:** ni el plan ni la trazabilidad dan nombres
  reales para las columnas puramente copiadas (D:J, T:AB) más allá de lo que las fórmulas
  revelan (D/T = fecha, E/U = hora). Se dejaron con su letra de Excel tal cual, en vez de
  inventarles un nombre — igual criterio que "no adivinar" ya aplicado en otras partes de este
  documento.

## 23.3. `Subastas` — replica `Cargar_Remuneracion_Subastas_Rapido`

- **Entrada:** `<CARPETA_BASE>/Subastas/3_REMUNERACIÓN_SUBASTAS_E_ID_*.xlsx` (o
  `.xlsm`/`.xlsb`/`.xls`) — el más reciente si hay varios.

  **Nota de arquitectura:** la macro original buscaba este archivo directamente en la carpeta
  del `.xlsm` (sin subcarpeta). Se le dio una carpeta propia (`Subastas/`) para ser consistente
  con el resto de las entradas externas de este proyecto (`Medidas/`, `Auxiliares/`, `Ofertas/`,
  `Cmg/`, `SSCC_Desempeño/`), cada una con su propia carpeta bajo la carpeta base del caso. Es
  una decisión de estructura de carpetas, no de contenido.
- **Lectura:** la macro original consulta la hoja `DB` por ADO/SQL (equivalente a filtrar y
  seleccionar columnas de una tabla); en Python se lee directamente con pandas aplicando el
  mismo filtro y la misma selección de columnas, sin necesidad de ADO. Datos desde la fila 3,
  columnas B:Y, filtrando donde la columna K (10ª del bloque B:Y) contenga "BESS" o "SAE".
- **Columnas resultantes (B:Q):**
  - `B:L`: copia directa de `DB!B:L` (11 columnas, filtradas).
  - `M` (fórmula): `= K & H & I`.
  - `N`: **pendiente** — la fórmula real (`N3:N6997` en la sección de fórmulas del libro) hace
    referencia a `'Calculo E Costos'!D:D`, `G:G` y `P:P`, una hoja de una etapa posterior que
    todavía no está implementada. Se deja como columna vacía documentada, no se adivina.
  - `O, P, Q`: copias directas de `DB!P`, `DB!Y`, `DB!V` respectivamente (así lo indica la
    macro original).
- **Fuera de alcance de esta sesión:** las columnas `U:W` de `Subastas` (resumen que depende de
  `N` y de otras hojas) y el formato del encabezado `B1` (fórmula puramente decorativa que arma
  una etiqueta de texto, no un dato por fila).

## 23.4. Hojas resultantes de `Consolidado_entradas.xlsx`

`Medidores`, `Ofertas SSCC`, `CMg`, `FD`, `Subastas`, `Log`.

---

# 24. Encabezados reales de FD y Subastas, y ajuste de la hoja Ofertas SSCC

El usuario entregó un Excel con los encabezados reales de `FD` y `Subastas`, confirmando (y
corrigiendo el nombre genérico por letra que se había usado en la sección 23) las columnas que
solo se copian:

## 24.1. `FD` — encabezados confirmados

Bloque CSF (A:M):

| Letra | Nombre |
|---|---|
| A | `id` |
| B | `Hora Mes` |
| C | `Dia` |
| D | `Fecha` |
| E | `Hora` |
| F | `Unidad` |
| G | `Respuesta CSF (Fact_CSF)` |
| H | `Disponibilidad (Fdis_CSF)` |
| I | `Desempeño (DCSF)` |
| J | `Factor de Desempeño (Fd_CSF)` |
| K | `CSF(+)` |
| L | `CSF(-)` |
| M | `Hora Mes` (repite el nombre de B; también repite su valor — `M = B`, confirmado antes) |

Bloque CPF (Q:AE), mismo patrón:

| Letra | Nombre |
|---|---|
| Q | `id` |
| R | `Hora Mes` |
| S | `Dia` |
| T | `Fecha` |
| U | `Hora` |
| V | `Unidad` |
| W | `Respuesta CPF+ (Fact_CPF+)` |
| X | `Respuesta CPF- (Fact_CPF-)` |
| Y | `Disponibilidad (Fdis_CPF)` |
| Z | `Desempeño (DCPF)` |
| AA | `Factor de Desempeño (Fd_CPF)` |
| AB | `Cuenta con equipo registrador validado` |
| AC | `CPF(+)` |
| AD | `CPF(-)` |
| AE | `Hora Mes` (repite el nombre de R; repite su valor — `AE = R`) |

N:P siguen vacías/fuera de alcance (el archivo de ejemplo tampoco tiene encabezado ahí).

Confirma, de paso, todas las fórmulas ya implementadas en la sesión anterior: `K = J` con
`K="CSF(+)"`, `J="Factor de Desempeño (Fd_CSF)"` — es decir, la columna "positiva" simplemente
copia el factor de desempeño sin condición de signo (y `L="CSF(-)"` copia a su vez el valor de
`K`). Aunque el nombre sugiera una lógica de signo, la fórmula extraída del `.xlsm` real es una
copia directa; no se "corrige" para que tenga más sentido semántico — replicar primero.

## 24.2. `Subastas` — encabezados confirmados

| Letra | Nombre | Origen |
|---|---|---|
| B | `Control` | copiado de `DB!B` |
| C | `Sub_Baj` | copiado de `DB!C` |
| D | `Fecha` | copiado de `DB!D` |
| E | `Año` | copiado de `DB!E` |
| F | `Mes` | copiado de `DB!F` |
| G | `Dia` | copiado de `DB!G` |
| H | `Hora_dia` | copiado de `DB!H` |
| I | `Hora_mes` | copiado de `DB!I` |
| J | `Configuración` | copiado de `DB!J` |
| K | `Propietario` | copiado de `DB!K` — **esta es la columna que se filtra por BESS/SAE** |
| L | `Clave horaria` | copiado de `DB!L` |
| M | `Ciclo` | fórmula: `= Propietario & Hora_dia & Hora_mes` (`K&H&I`) |
| N | `Energía SSCC` | **pendiente** — depende de `'Calculo E Costos'`, etapa sin implementar |
| O | `FD` | copiado de `DB!P` |
| P | `FMA` | copiado de `DB!Y` |
| Q | *(sin nombre en el archivo real)* | copiado de `DB!V` |

`A` (`Concepto`) no forma parte de lo que escribe `Cargar_Remuneracion_Subastas_Rapido`, y
`R:V` (`Configuración`, `Ciclo`, `Clave`, `SUBIDA`, `BAJADA`) tampoco: son columnas de otra
lógica (probablemente relacionada con las fórmulas `U:W` ya descartadas como fuera de alcance
en la sección 23.3).

## 24.3. Hoja "Ofertas SSCC": las dos tablas van lado a lado, no una debajo de la otra

Ajuste pedido por el usuario: "Ofertas SSCC por dia" y "Resumen ventana oferta" (sección 22)
pasan de estar apiladas verticalmente a estar **una al lado de la otra**, con una separación de
2 columnas en blanco entre ambas. `_escribir_tabla_con_titulo()` ahora acepta `columna_inicio`
además de `fila_inicio`, y devuelve tanto la fila como la columna donde podría continuar el
siguiente bloque (cada llamada usa el dato que corresponda a cómo se estén acomodando los
bloques en ese momento).

---

# 25. `Calculo E Costos` — etapa base (H + CMg + traspaso de Medidores)

El usuario pidió crear la hoja `Calculo E Costos` ("Ecostos"), cuya lógica de cálculo vive en la
macro `Actualizar_Calculos_Columnas` (módulo `J_Calculo_Ecostos`), con la barra (columna H)
resuelta por fórmula y el CMg asignado por la macro `Asignar_CMg_a_Calculos_Turbo` (módulo
`A_Carga_Cmg_a_Destino`). También pidió explícitamente que esta hoja viva en un **archivo
separado** de `Consolidado_entradas.xlsx` ("esto que esté en otra planilla que se llame
pagos_bess o algo así por ahora").

`Actualizar_Calculos_Columnas` es una macro muy extensa (~1500 líneas) con dependencias
profundas (Subastas, Resumen, una "Prorrata SSCC" que resultó ser una tabla dinámica derivable
de Subastas, Diccionario, FD). Frente a esa escala, se preguntó al usuario cómo priorizar y
eligió explícitamente: **"por etapas: primero H + CMg + traspaso de Medidores"**. Esta sección
documenta esa primera etapa; el resto queda pendiente (ver 25.4).

## 25.1. `Traspasar_Medidores_A_Calculos_Rapido` (módulo `B_medidores_a_calculos`)

Traspasa, para cada fila de `Medidores`, hacia **ambas** hojas `Calculo E Costos` y `Calculo
RE545` (esta última fuera de alcance por ahora):

- **A:G**, con **D y E invertidas** respecto a `Medidores` (`Destino!D = Medidores!E`,
  `Destino!E = Medidores!D`). El resto de A:G se copia tal cual.
- **H no se toca** — es fórmula (`=VLOOKUP(G4,Resumen!B:G,6,FALSE)`), ver 25.2.
- **I/J**: la energía de `Medidores!I` (`Gen_Unidad`) se separa por signo (positivo → I,
  negativo → J, cero → ambas 0) y se enruta según `Medidores!T` (`Ventana_No_Completa`):
  `T = 1` → la energía va a `Calculo E Costos!I/J`; `T <> 1` o vacío → va a `Calculo
  RE545!I/J` (y en `Calculo E Costos` queda 0/0 esa fila).
- **`Medidores!J` (`SoC`) → `Destino!K`.**
- **`Medidores!K` (`Copia_Ventana`) → `Destino!P`.**
- `Medidores!L` (`Ventana`) → `Calculo RE545!T` únicamente (no aplica a `Calculo E Costos`).

Todas las filas de `Medidores` se traspasan a ambas hojas destino (no se filtran): lo que
cambia según `T` es si la energía I/J queda con valor real o en 0 en cada una.

## 25.2. `H` (Barra) — antes fórmula `VLOOKUP`

En el `.xlsm` original, `Calculo E Costos!H4:H26787` es la fórmula:

```
=VLOOKUP(G4,Resumen!B:G,6,FALSE)
```

Busca el valor de `G` (la central, columna `clave` en nuestro `Medidores`) en la columna `B` de
la hoja `Resumen` del libro original, y devuelve la 6ª columna del rango `B:G` (es decir, la
columna `G` de esa hoja) — la barra de inyección de esa central.

**No se replica por posición** (`Resumen!B:G` del `.xlsm` original) sino **por nombre de
columna**, homologando `Medidores!clave` contra `Resumen BESS!Nombre activo` y devolviendo
`Resumen BESS!Barra inyección` (`construir_mapa_barra()`). Motivo: `Centrales.xlsx` (nuestro
maestro, sección 4) no reproduce el layout `Resumen!B:G` del libro original — homologar por
nombre es más fiel a la intención de la fórmula (encontrar la barra de la central) que asumir
una posición de columna no confirmada.

## 25.3. `Asignar_CMg_a_Calculos_Turbo` (módulo `A_Carga_Cmg_a_Destino`) — solo Q

La macro arma un diccionario a partir de la hoja `CMg` (columnas por posición: `D` = Barra, `F`
= valor a asignar en `Q`, `H` = Cuarto de Hora, `I` = valor a asignar en `R`, este último solo
para `Calculo RE545`):

```
clave = UCase(Trim(CMg!D)) & "|" & NormalizaCuarto(CMg!H)
```

Si la clave se repite, gana la **primera** fila (`If Not dictCMg.Exists(clave) Then Add`).

`NormalizaCuarto` (replicada en `_normaliza_cuarto()`):

```
Error     -> ""
Numérico  -> CStr(CLng(valor))   ' texto del entero redondeado
Otro      -> Trim(CStr(valor))
```

Para completar el destino (`CompletarDestinoTurbo`), la clave de búsqueda en cada fila de
destino es `UCase(Trim(Destino!H)) & "|" & NormalizaCuarto(Destino!F)` (`H` = Barra, `F` =
Cuarto de Hora) y solo se asigna `Q = CMg!F` — para `Calculo E Costos` la macro llama con
`escribirR:=False`, así que `R` (que en `Calculo E Costos` ni siquiera existe como columna
usada) nunca se toca. `R` sólo aplica a `Calculo RE545`, fuera de alcance.

`construir_dic_cmg(df_cmg)` arma ese diccionario a partir del `df_cmg` que ya lee `leer_cmg()`
(columnas por posición, sin renombrar — `D` = `df_cmg.columns[3]`, `F` = `columns[5]`, `H` =
`columns[7]`).

## 25.4. Alcance de esta etapa y lo que queda pendiente

`construir_calculo_e_costos()` arma únicamente:

| Columna (nombre placeholder) | Origen |
|---|---|
| `Mes`, `Dia`, `Hora`, `clave` | copia directa de `Medidores` |
| `Hora Mes`, `Minutos` | copia de `Medidores` (D↔E invertidas, ver 25.1) |
| `Cuarto de Hora` | copia directa de `Medidores` |
| `Barra` (H) | homologada por nombre contra `Resumen BESS` (ver 25.2) |
| `Energia_Positiva`, `Energia_Negativa` (I/J) | `Medidores!Gen_Unidad` separada por signo, solo si `Ventana_No_Completa = 1` |
| `SoC` (K) | copia de `Medidores!SoC` |
| `Copia_Ventana` (P) | copia de `Medidores!Copia_Ventana` |
| `CMg` (Q) | homologado por `Barra` + `Cuarto de Hora` contra `cmg.xlsx` (ver 25.3) |

Los nombres de columna son **placeholders** derivados de los comentarios de la macro: el
usuario adjuntó dos veces un archivo pensado para traer los encabezados reales de `Ecostos`,
pero ambas veces solo incluía las hojas `FD`/`Subastas` (ya confirmadas en la sección 24). Se
corrigen los nombres apenas se reciba el archivo correcto, sin reinterpretar el resto de la
lógica ya implementada.

Explícitamente **fuera de alcance** de esta etapa (decisión del usuario, "por etapas"):

- El resto de columnas de `Actualizar_Calculos_Columnas`: `L`, `M`, `N`, `O`, `R`, `S`, `T`,
  `U`, `W`, `X`, `Y`, `AB:AF`, `AG:AX`, `AZ`. Incluye la lógica de grupos por `G+P`
  (`CrearClave2`), rankings por `Q`/`F` dentro de cada grupo, y la "Prorrata SSCC" — que el
  usuario confirmó que **no** es un archivo externo sino una tabla dinámica derivable de
  `Subastas` (`Filas: Configuración, Hora_mes` / `Columnas: Control` / `Valores: Cuenta de
  Sub_Baj`), por lo que no es un bloqueador real para una futura etapa.
- Toda la hoja `Calculo RE545` (la misma macro de traspaso la alimenta, pero con las columnas
  I/J invertidas por la condición de `T`, y `R`/`T` propios que no aplican a `Calculo E
  Costos`).

## 25.5. `Pagos_BESS.xlsx`

Salida nueva y separada de `Consolidado_entradas.xlsx`, a pedido explícito del usuario. Nombre
y alcance provisorios (`ARCHIVO_SALIDA_PAGOS = "Pagos_BESS.xlsx"`). Por ahora tiene una sola
hoja, `Calculo E Costos` (`HOJA_CALCULO_ECOSTOS`), con las columnas descritas en 25.4 y 25.7.

## 25.6. Etapa 2: `L, N, O, R, S, T, U, W, X, Y, AB, AC, AD` — y lo que en su momento quedó bloqueado

Segunda etapa de `Actualizar_Calculos_Columnas`, implementada después de que el usuario pidiera
explícitamente continuar ("necesito que termines con el calculo de Ecostos"). Antes de
implementar se encontró un problema real: buena parte de las columnas restantes dependían de
datos que en ese momento **no existían en ningún archivo ya mapeado** en la migración:

- Una hoja `Resumen` del libro original con una tabla central→factor (columnas B:C desde la
  fila 8) y un umbral único en `H8`, usada por `M` (umbral) y `AE`/`AF` (factor). **Resuelto en
  25.8**: es la misma tabla que `Centrales.xlsx!Resumen BESS`, no una hoja aparte.
- Un umbral de subida/bajada por `Configuración+P`, usado en `AU`/`AV`/`AW`/`AZ`. **Sigue sin
  resolverse** (ver 25.8 y "Pendientes abiertos" en `BITACORA.md`).

En ese momento se implementó únicamente lo que no dependía de la hoja `Resumen`: `L`, `N`, `O`,
`R`, `S`, `T`, `U`, `W`, `X`, `Y`, `AB`, `AC`, `AD`. `M`, `AE` y `AF` se agregaron después (25.8);
`AG:AZ` sigue pendiente.

### El problema de la columna `L` y cómo se resolvió

La fórmula real del `.xlsm` para `L` es:

```
=1*(OR(COUNTIFS(Subastas!$K:$K,G4,Subastas!$G:$G,A4,Subastas!$H:$H,B4,Subastas!$I:$I,C4,
        Subastas!$D:$D,"BAJADA"),
       COUNTIFS(Subastas!$K:$K,G4,Subastas!$G:$G,A4,Subastas!$H:$H,B4,Subastas!$I:$I,C4,
        Subastas!$D:$D,"SUBIDA")))
```

Es decir: marca 1 si existe una fila en `Subastas` de tipo BAJADA o SUBIDA cuya
central+mes+día+hora coincide con la fila de `Calculo E Costos`. El código VBA
(`CrearDiccionarioSubastas`) documenta esa clave por posición de columna (`Subastas!D` = tipo,
`G,H,I,K` = la clave), pero esas posiciones **no coinciden** con los encabezados reales de
`Subastas` que el usuario ya había confirmado (sección 24.2): ahí `D` es `Fecha`, no un texto
BAJADA/SUBIDA.

Se le preguntó al usuario dónde vive ese dato. Respuesta: **"Es la columna C de la hoja
subastas que ya generamos"** — o sea `Subastas!Sub_Baj` (literalmente "Subida/Bajada"), como ya
sugería el propio nombre de esa columna. Con esa confirmación, la clave equivalente se
reconstruyó **por nombre de columna real**, no por posición:

- tipo → `Subastas!Sub_Baj` (confirmado por el usuario)
- central → `Subastas!Configuración` (**inferido**, no confirmado letra por letra: es el mismo
  campo que usa la tabla dinámica Prorrata SSCC como identificador de central, y el que produce
  una alineación semántica limpia con `Mes`/`Dia`/`Hora_dia` de Subastas contra `Mes`/`Dia`/
  `Hora` de `Calculo E Costos` — `Propietario`, la otra columna candidata, se descartó porque es
  el campo que se usa para filtrar por BESS/SAE, no para identificar una central puntual)
- mes/día/hora → `Subastas!Mes`, `Subastas!Dia`, `Subastas!Hora_dia`

**Pendiente de validar contra un caso real**: si al correr esto la cantidad de filas con `L=1`
sale sospechosamente baja o en cero, la primera sospechosa es esta inferencia (`Configuración`
en vez de `Propietario`). Dicho esto, el archivo de encabezados reales que confirmó los nombres
de `Calculo E Costos` (25.8) subió bastante la confianza en esta elección: la columna `G` de
`Calculo E Costos` se llama literalmente **`Configuracion`** — el mismo nombre de campo que
`Subastas!Configuración` — lo que es un indicio fuerte (aunque no una prueba fila por fila) de
que es el campo correcto para homologar centrales entre las dos hojas.

## 25.7. Detalle columna por columna de la etapa 2

Implementadas en `nucleo.py` (`completar_calculo_e_costos_grupos()` y sus funciones internas):

| Columna | Lógica |
|---|---|
| `L` | Ver 25.6. |
| `N`, `O` | Por grupo (central=`clave` + ventana=`Copia_Ventana`): ordenando por `Cuarto de Hora` descendente, suma acumulada de `Energia_Positiva` (N) y de `-Energia_Negativa` (O), contando solo filas con `L=1`, repartida a **todas** las filas que comparten el mismo `Cuarto de Hora` del grupo (no solo a las que tienen `L=1`). |
| `R` | Por grupo: ranking ordenando por `CMg` descendente y `Cuarto de Hora` descendente; las filas empatadas en ambos comparten el mismo ranking (la posición donde empieza el empate — "competition ranking", no denso). |
| `S`, `T`, `U` | `S = Energia_Positiva × CMg`, `T = Energia_Negativa × CMg`, `U = L × (S + T)`. No dependen de agrupar. |
| `W`, `X` | `X = Copia_Ventana` (copia directa). `W` es un contador **global** (no por grupo) que se reinicia a 1 cada vez que `X` cambia respecto de la fila anterior; la primera fila de todo el archivo es una excepción fiel al original: `W` toma el valor de `Hora` en vez de 1 (y esa diferencia se arrastra en el resto de su bloque). |
| `Y`, `AB` | Por grupo: de las filas con `L=1` y `Energia_Positiva≠0`, ordenadas por `CMg` descendente, la fila en la posición `j` (dentro del orden ORIGINAL del grupo) recibe el `Cuarto de Hora` (`Y`) y el `CMg` (`AB`) de la `j`-ésima fila calificada; si el grupo tiene menos filas calificadas que filas totales, las posiciones sobrantes toman el `Cuarto de Hora` de las filas NO calificadas en su orden original (y `AB` queda vacío). |
| `AC`, `AD` | Igual que `Y`/`AB` pero para `Energia_Negativa≠0`, ordenadas por `CMg` **ascendente**. |

En su momento bloqueadas, resueltas en 25.8: `M`, `AE`, `AF`. Todavía pendiente: `AG:AX`, `AZ`
(ver 25.9).

`generar_pagos_bess()` ahora exige también la hoja `Subastas` de `Consolidado_entradas.xlsx`
(además de `Medidores`) para poder calcular `L`.

## 25.8. `M`, `AE`, `AF` — la hoja "Resumen" resultó ser `Resumen BESS`

El usuario adjuntó `Centrales.xlsx` (real) y un `Libro1.xlsx` con 4 hojas: `FD`, `subastas`,
`E COSTOS` y — la pieza que faltaba — **`Resumen`**. Comparando esa hoja `Resumen` del libro
original contra `Centrales.xlsx!Resumen BESS`, resultaron ser **la misma tabla**: mismos 9
encabezados en el mismo orden (`Nombre activo`, `Pmax (MW)`, `Horas para descarga forzada`,
`Capacidad (MWh)`, `Energía mínima`, `Barra inyección`, `% Energía sobre mínima (indicador
nuevo ciclo)`, `Ciclos max diarios`, `Eficiencia`). No hacía falta una hoja nueva ni un archivo
aparte — la sección 25.2 ya venía usando esta misma tabla para `H`/`Barra`.

Con eso, `Resumen!B:C` (factor, `AE`/`AF`) y `Resumen!H8` (umbral, `M`) se resuelven así:

- **Factor** (`AE`/`AF`): columna `Pmax (MW)` de `Resumen BESS`, homologada por `Nombre activo`
  igual que `Barra inyección` (`construir_dic_resumen_factor()`).
- **Umbral** (`M`): en el archivo real, la celda `H8` cae justo en la primera fila de datos
  (fila de encabezados = 7, primera central = fila 8), columna `% Energía sobre mínima
  (indicador nuevo ciclo)`. Se toma igual acá: el valor de esa columna en la **primera fila con
  nombre de central** de `Resumen BESS` (no una fila fija — el encabezado de `Centrales.xlsx`
  no siempre cae en la misma posición, ver `_leer_resumen_bess()`).

Columnas agregadas en `nucleo.py` (`calcular_m()`, `_calcular_asignacion_energia()`,
`calcular_ae_af()`):

| Columna | Lógica |
|---|---|
| `M` | `1` si `SoC > umbral`, si no `0`. |
| `AE`, `AF` | Por grupo (central+ventana): reparte el máximo de `N` (para `AE`) / `O` (para `AF`) del grupo en bloques de 15 minutos según el `Bloque ordenado` (`W`) de cada fila y el factor (`Pmax`) de la central — replica `CalcularAsignacionEnergia`: el bloque recibe el 100% del factor si cae dentro de la cantidad de bloques llenos, una fracción en el bloque siguiente si sobra un resto, y 0 en el resto. `AF` lleva el signo cambiado. Si la central no tiene factor, o el factor es 0, `AE`/`AF` quedan en blanco (`pd.NA`) — equivalente a los `#N/A`/`#DIV/0!` del original, sin fabricar un tipo de error de Excel en Python. |

## 25.9. Encabezados reales de `Calculo E Costos` — y lo que queda pendiente

El mismo `Libro1.xlsx` trae la hoja `E COSTOS` con los encabezados reales de **toda** la hoja
(fila 3 tiene el nombre de cada columna; filas 1-2 son títulos de grupo). Se aplicaron a todas
las columnas ya implementadas (`NOMBRES_CALCULO_E_COSTOS` en `nucleo.py`, mismo patrón que
`NOMBRES_FD_CSF`/`NOMBRES_SUBASTAS`: se calcula todo con los nombres/letras internos usados
hasta acá y se renombra recién al final, en `completar_calculo_e_costos_grupos()`):

| Interno | Real | Interno | Real |
|---|---|---|---|
| `Mes` | `Mes` | `S` | `Valorizacion Descarga` |
| `Dia` | `Dia` | `T` | `Valorizacion Carga` |
| `Hora` | `Hora` | `U` | `Total` |
| `Hora Mes` | `Hora mes` | `W` | `Bloque ordenado` |
| `Minutos` | `Minuto` | `X` | `Ciclo` |
| `Cuarto de Hora` | `Bloque horario` | `Y` | `Bloque Mes Descarga` |
| `clave` | `Configuracion` | `AB` | `Curva monotona CMg Descarga` |
| `Barra` | `Barra` | `AC` | `Bloque Mes  Carga` (dos espacios, tal cual el archivo) |
| `Energia_Positiva` | `Descarga kWh` | `AD` | `Curva monotona CMg Carga` |
| `Energia_Negativa` | `Carga kWh` | `AE` | `Energía descargada` |
| `SoC` | `SoC %` | `AF` | `Energía cargada` |
| `Copia_Ventana` | `Ciclo de Carga del mes` | `L` | `Adj SSCC` |
| `CMg` | `CMg` | `M` | `SoC sobre el minimo` |
| `N` | `Energía SSCC (-) por remunerar` | `O` | `Energía SSCC (+) por remunerar` |
| `R` | `ranking cmg` | | |

Lo que en su momento quedó pendiente (`AG:AZ`) según esta misma hoja — resuelto en su mayoría
en 25.10, salvo `AW:AZ`:

| Real | Rol |
|---|---|
| `CPF(-)`, `CSF(-)`, `CTF(-)`, `CPF(+)`, `CSF(+)`, `CTF(+)` (`AG:AL`, grupo "Prorratas") | **Implementado (25.10)**. |
| `CPF(-)`, `CSF(-)`, `CTF(-)`, `CPF(+)`, `CSF(+)`, `CTF(+)` (`AM:AR`, grupo "FD") | **Implementado (25.10)**. `CTF` confirmado por el usuario: no existe, sale en 0 — coincide con que el VBA original también las deja hardcodeadas en 0. |
| `Energía descarga/carga con FD` (`AS`, `AT`) | **Implementado (25.10)**. |
| `Ingreso descarga`, `Costo carga` (`AU`, `AV`) | **Implementado (25.10)**. |
| `Descuento FD`, `Total` (`AW`, `AX`, grupo "Componente 1") | **Todavía pendiente** — `AW` necesita el umbral de subida/bajada por central+ventana; ver 25.10. |
| `Monto a compensar` (`AZ`) | **Todavía pendiente** — depende de `AX`. |

## 25.10. `AG:AV` — Prorrata SSCC, homologación FD y costo ponderado

El usuario confirmó dos cosas que destrabaron esta etapa:

1. **`CTF` no existe**: "creo que no tiene ctf no está en los FD y en la hoja de los ecostos
   sale con 0". Coincide exactamente con el código: `salidaAGAX(i, 3) = 0` (`AI`),
   `salidaAGAX(i, 6) = 0` (`AL`), `salidaAGAX(i, 9) = 0` (`AO`), `salidaAGAX(i, 12) = 0` (`AR`)
   — están hardcodeadas en 0 en el VBA original, no dependen de ningún diccionario. No hacía
   falta encontrar un origen para `CTF`: nunca existió como dato real en esta parte del cálculo.

2. **El archivo de Subastas está corrido una columna**: "Las subastas que te mande vs las del
   original están corridas una columna, la primera en el original está vacía". Esto explica la
   discrepancia de la sección 25.6: aplicando ese corrimiento de una columna a lo que el VBA
   documentaba (`Subastas!D` = tipo, `G,H,I,K` = clave), se obtiene exactamente
   `Subastas!Sub_Baj` (tipo) + `Configuración+Mes+Dia+Hora_dia` (clave) — lo mismo que ya se
   había implementado por inferencia en 25.6, ahora con una explicación clara de por qué las
   letras del VBA no coincidían con los encabezados reales.

Con eso, se implementó:

### Prorrata SSCC (`AG, AH, AI, AJ, AK, AL`)

`construir_prorrata_sscc()` arma la tabla dinámica en pandas (`pivot_table`, `index=
[Configuración, Hora_mes]`, `columns=Control`, `values=Sub_Baj`, `aggfunc=count`) — **no es un
archivo externo**, se deriva de `Subastas` (confirmado por el usuario, ver 25.4).
`construir_dic_prorrata()` busca, entre las columnas que deja el pivot (una por cada valor de
`Control`), la que contenga "CPF" y la que contenga "CSF" en el nombre — **inferido**, no
confirmado letra por letra que `Control` tenga exactamente esos dos valores; si no las
encuentra, avisa y usa 0. `calcular_prorratas()` homologa por central+`Hora Mes`.

Replicando el VBA (`salidaAGAX(i,1)=valorAG; salidaAGAX(i,4)=valorAG` — **AG y AJ son el mismo
valor**, igual `AH`/`AK`): `AJ = AG`, `AK = AH` (Prorratas "+" duplica literalmente las
Prorratas "-"; no es un error, así está en el original). `AI = AL = 0` (`CTF`, ver arriba).

### FD homologado (`AM, AN, AO, AP, AQ, AR`)

`construir_dic_mapeo_diccionario()` replica `CrearDiccionarioPrimerValor(Diccionario, 1, 2)`:
columna A → columna B de `Diccionario` (una TERCERA lectura de esa hoja, distinta de
`construir_homologacion()` y de `_mapas_homologacion_fge()` — no fusionar). `_calcular_bloque()`
replica `CalcularBloque` (`Int((valor-1)/4)+1`). `calcular_fd_prorrateado()` arma la central
homologada + el bloque de `Y` (para descarga) o de `AC` (para carga), y busca esa clave en
`FD!CSF(±)`/`CPF(±)` (ya construidos por `construir_fd()`, requiere ahora también el archivo
`SSCC_Desempeño_*` para `generar_pagos_bess()`, igual que ya lo exige `generar_consolidado()`
para la sección `"fd"`). Si la central no está en `Diccionario`, las 4 quedan en blanco
(`pd.NA`, equivalente al `#N/A` del original); si está pero no hay match en `FD`, quedan en 0
(fiel al original: ahí solo se registra un aviso, no se propaga un error). `AO = AR = 0` (`CTF`).

### `AS`, `AT` — costo ponderado

`_calcular_costo_ponderado()` replica `CalcularCostoPonderado`: si `AG+AH` (las cantidades de
Prorrata) es mayor a 0, pondera `AM`/`AN` (o `AP`/`AQ`) por esas cantidades; si no, el factor es
1. Si `AE`/`AF` es blanco, el resultado es blanco; si `AG+AH>0` pero el precio (`AM`/`AN`/`AP`/
`AQ`) es blanco, también.

### `AU`, `AV` — ingreso/costo por Componente 1

`calcular_au_av()` replica el promedio de `AB` (para `AU`) / `AD` (para `AV`) dentro del grupo
central+ventana, entre las filas con `AE`/`AF` válido y distinto de 0, multiplicado por `AE`/
`AF` — pero solo si la suma **global** de energía (todas las centrales que comparten la misma
`Copia_Ventana`, sin agrupar por central) supera 10 (`AU`) o es menor a -10 (`AV`). Esta suma
global (`sumaIP`/`sumaJP` en el VBA) es una agrupación **distinta** de la de `N/O/R/Y/AB/AC/AD`
(que sí es por central+ventana) — no confundirlas.

### Lo que sigue pendiente: `AW`, `AX`, `AZ`

`AW` ("Descuento FD") necesita un umbral de subida/bajada por central+ciclo que, en el `.xlsm`
original, vive en una tabla resumen aparte de `Subastas` (no en el bloque principal B:Q).
Aplicando el mismo corrimiento de columna de más arriba a lo que mostraba el archivo de
encabezados (`Subastas!R:V` = `Configuración`, `Ciclo`, `Clave`, `SUBIDA`, `BAJADA`), la tabla
real parecería estar en `S:W`, con `U`("Clave") = `Configuración & "&" & Ciclo`, y `V`/`W`
conteniendo un valor bajo el título "SUBIDA"/"BAJADA" calculado con una fórmula `COUNTIFS` que
a su vez depende de `Subastas!N` ("Energía SSCC"), que a su vez depende de `Calculo E Costos!P`
— una dependencia circular con nuestro propio cálculo que todavía no se terminó de decantar. No
se implementa `AW`/`AX`/`AZ` hasta resolver esto.
