# Balance_BESS

Herramienta en Python que reemplaza, hoja por hoja, el cálculo hecho hoy en
`11_PAGOS_BESS_2607_Definitivo.xlsm` (Balance BESS / SSCC). La planilla se
usa solo como referencia de validación; el proceso Python no depende de
información almacenada exclusivamente en ese libro.

Etapas implementadas hasta ahora: **Medidores** (incluye Ofertas SSCC), la
carga de **CMg**, **FD** y **Subastas**, y las dos hojas de cálculo del libro
completas: **Calculo E Costos** y **Calculo RE545**.

## Instalación

```bash
pip install -r requirements.txt
```

## Uso

```bash
python Balance_BESS.py
```

1. Elegir la **carpeta base** del caso (ver estructura abajo). El programa
   recuerda la última carpeta usada, por PC/usuario, en `config.json`.
2. Ingresar el **período (AAMM)** en el recuadro de la ventana: 4 dígitos,
   año+mes simplificado (ej. `2607` para julio de 2026). No se adivina del
   nombre de ningún archivo — es el dato que el programa usa para ubicar el
   SoC del período dentro de `Medidas/`.
3. La ventana detecta automáticamente las entradas y las dibuja como un
   diagrama de carpetas (`OK` / `FALTA` / `PENDIENTE` por cada una).
4. Al final del diagrama están `Consolidado_entradas.xlsx` y
   `Pagos_BESS.xlsx`, cada una con su botón **Generar...**. Ese botón abre
   una ventana aparte donde se elige qué partes recalcular esta vez (una
   casilla por hoja); lo que no se tilda se conserva tal cual estaba en el
   archivo existente. `Pagos_BESS.xlsx` tiene dos casillas: `Calculo E
   Costos` y `Calculo RE545`.

## Estructura de carpeta de un caso

```text
<CARPETA_BASE>/
├── Medidas/
│   ├── Medidas_SAE.xlsx
│   └── <algún archivo .xlsx cuyo nombre contenga "SOC" y el AAMM,
│        ej. SOC_2607.xlsx, "resumen soc julio 2607.xlsx">
├── Auxiliares/
│   └── Centrales.xlsx       (hojas "Resumen BESS" y "Diccionario")
├── Ofertas/
│   └── <algún archivo Excel cuyo nombre contenga "OfertasSSCC">
├── Cmg/
│   └── cmg.xlsx             (nombre literal, no cambia con el período)
├── SSCC_Desempeño/
│   └── <algún archivo Excel cuyo nombre empiece con "SSCC_Desempeño_">
├── Subastas/
│   └── <algún archivo Excel cuyo nombre empiece con
│        "3_REMUNERACIÓN_SUBASTAS_E_ID_">
├── Consolidado_entradas.xlsx    <- salida generada por el programa
└── Pagos_BESS.xlsx              <- salida generada por el programa
                                     (nombre provisorio)
```

Ningún archivo (salvo `cmg.xlsx`) sigue un nombre fijo:

- **SoC**: cualquier `.xlsx` en `Medidas/` cuyo nombre contenga "SOC" y el
  AAMM ingresado en la ventana. Si hay más de un archivo que cumple la
  condición, el programa se detiene y pide dejar solo el del período
  correspondiente (no elige por fecha de modificación).
- **OfertasSSCC**, **SSCC_Desempeño_\*** y **3_REMUNERACIÓN_SUBASTAS_E_ID_\***:
  cualquier archivo Excel (`.xlsx`/`.xlsm`/`.xlsb`/`.xls`) en su carpeta
  correspondiente cuyo nombre contenga (Ofertas) o empiece con (los otros
  dos) ese texto. Si hay más de uno, a diferencia del SoC, se toma
  automáticamente el más reciente por fecha de modificación — así lo hacen
  las macros originales de la planilla.
- **cmg.xlsx**: única excepción con nombre literal fijo, dentro de `Cmg/`.

La carpeta base puede estar en cualquier ubicación (disco local, red,
OneDrive); moverla o mover `Balance_BESS.py` a otro lugar no cambia el
resultado, siempre que la carpeta base seleccionada sea la misma.

## Documentación

| Querés... | Leé |
|---|---|
| Entender cómo se trabaja en este repo (para un asistente de IA o alguien nuevo) | `METODOLOGIA.md` |
| El checklist obligatorio de inicio/cierre de sesión | `REGLAS.md` |
| Qué hace cada script, en dos líneas | `MAPA.md` |
| Historial de sesiones y pendientes abiertos | `BITACORA.md` |
| Reglas de negocio del cálculo y el plan completo de migración | `docs/Plan_Traspaso_Python_Balance_BESS.md` |
| El código VBA original, las fórmulas del `.xlsm` y de dónde sale cada dato | `docs/Trazabilidad_11_PAGOS_BESS_2607_Definitivo.md` |
| La hoja `Calculo RE545` real (recortada), con sus nombres de columna y fórmulas | `docs/Calculo_RE545_reducido_para_IA.xlsx` |
| La hoja `Subastas` real, con encabezados y fórmulas (fuente de la corrección de `NOMBRES_SUBASTAS`); también trae la hoja `E COSTOS` con los encabezados de grupo reales (celdas combinadas), fuente de `GRUPOS_CALCULO_E_COSTOS` | `docs/Libro1_Subastas_real.xlsx` |
| `Centrales.xlsx` y `SOC_AAMM.xlsx` reales (fuente de las correcciones de `detectar_fila_nombres()` y `construir_homologacion()`) | `docs/Centrales_real.xlsx`, `docs/SOC_real_2607.xlsx` |
| Primera comparación real vs Python de `Calculo E Costos` fila a fila (fuente de la corrección de la Prorrata SSCC) | `docs/Pagos_BESS_comparacion_real.xlsx` |

## Estado actual

`Consolidado_entradas.xlsx` (antes `Hoja_Medidas.xlsx`) tiene seis hojas:

- `Medidores` — columnas A:U (A:J entrada, K copia de L, L, N, O, R, S, T
  calculadas; M, P, Q, U deliberadamente vacías por diseño).
- `Ofertas SSCC` — las tablas auxiliares equivalentes a `Medidores!W:Y`
  ("Ofertas SSCC por dia") y `Medidores!AB:AE` ("Resumen ventana oferta"),
  una al lado de la otra. El resumen intermedio equivalente a la hoja
  "Resumen Ofertas SSCC" del `.xlsm` original es puramente auxiliar y no se
  persiste.
- `CMg` — copia ordenada de `Cmg/cmg.xlsx` (replica
  `Cargar_CMg_Desde_Archivo`).
- `FD` — datos de `CPF Horario`/`CSF Horario` filtrados por BESS/SAE, más
  sus columnas calculadas, con sus nombres reales de columna (replica
  `Cargar_SSCC_Desempeno_En_FD`). Los bloques CSF (A:M) y CPF (Q:AE) son dos
  tablas de distinto largo, lado a lado en la misma hoja.
- `Subastas` — datos de la hoja `DB` filtrados por "Configuración" contiene
  BESS/SAE (funcionalmente equivalente a filtrar por Propietario: los
  nombres de central BESS empiezan con "SAE-"), más la columna "Clave
  horaria" calculada, con sus nombres reales de columna (`Concepto`,
  `Control`, `Sub_Baj`, ..., `Energía SSCC`, `FD`, `FMA` — corregidos en
  una sesión posterior, ver `BITACORA.md`; replica
  `Cargar_Remuneracion_Subastas_Rapido`). La columna "Ciclo" queda vacía:
  depende de `Calculo E Costos`.
- `Log` — avisos e incidencias detectadas durante el cálculo.

`Pagos_BESS.xlsx` (nombre provisorio, a pedido del usuario) tiene dos hojas:

- `Calculo E Costos` — traspaso desde `Medidores`, asignación de CMg, y casi
  toda `Actualizar_Calculos_Columnas` (`L, M, N, O, R, S, T, U, W, X, Y, AB,
  AC, AD, AE, AF, AG:AX, AZ`), con **nombres reales de columna** (confirmados
  contra un archivo real, hoja "E COSTOS"): `Configuracion`, `Barra`,
  `Descarga kWh`/`Carga kWh`, `SoC %`, `CMg`, `Adj SSCC`, `SoC sobre el
  minimo`, `ranking cmg`, `Valorizacion Descarga`/`Carga`, `Bloque ordenado`,
  `Ciclo`, `Curva monotona CMg Descarga`/`Carga`, `Energía descargada`/
  `cargada`, las Prorratas y el FD homologado (`CPF(±)`/`CSF(±)`/`CTF(±)`,
  este último siempre en 0 — confirmado que no existe), `Ingreso descarga`,
  `Costo carga`, `Descuento FD`, `Total` y `Monto a compensar`, entre otros.
  Requiere ahora también el archivo `SSCC_Desempeño_*` (para el FD
  homologado). La hoja queda **completa** (`A:AZ`, sin la columna `AY`, que
  la macro original tampoco escribe).
- `Calculo RE545` — la hoja hermana, también **completa** (`A:CE`): mismo
  traspaso desde `Medidores` (la energía se reparte entre las dos hojas según
  `Ventana_No_Completa`), las reservas por subasta (`AC:AU`), el resumen por
  central+ventana (`AW:BG`, una tabla de otro largo que se escribe al lado) y
  los Componentes 1 y 2 (`BI:CE`), hasta el `Monto a compensar`.

Las dos hojas de `Pagos_BESS.xlsx` llevan además, arriba de los nombres de
columna, los encabezados de grupo con celdas combinadas del archivo real
(`Dia`, `Nombre`, `BESS`, `Prorratas (-)/(+)`, `FD`, `Subastas`, `FMA`,
`Componente 1`/`Componente 2` — `GRUPOS_CALCULO_E_COSTOS`/
`GRUPOS_CALCULO_RE545`); como la salida no reproduce la letra de Excel real
(solo el orden y el contenido), cada grupo cae en la columna que le toca en
**nuestro** orden, no en la del archivo original.

Las macros de Ofertas SSCC, CMg, FD y Subastas replicadas son solo las de
**carga** de esas hojas.

Validado con casos sintéticos (no con datos reales todavía): ver
`BITACORA.md` → "Pendientes abiertos" para lo que falta antes de dar por
cerrada cada etapa (validación contra un caso real y contra la planilla 11).
