# Balance_BESS

Herramienta en Python que reemplaza, hoja por hoja, el cálculo hecho hoy en
`11_PAGOS_BESS_2607_Definitivo.xlsm` (Balance BESS / SSCC). La planilla se
usa solo como referencia de validación; el proceso Python no depende de
información almacenada exclusivamente en ese libro.

Etapas implementadas hasta ahora: **Medidores** (incluye Ofertas SSCC) y la
carga de **CMg**, **FD** y **Subastas**.

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
3. La ventana detecta automáticamente las entradas y muestra un checklist
   (`OK` / `FALTA` / `PENDIENTE`).
4. El botón **Ejecutar** se habilita cuando no falta nada requerido y genera
   `Consolidado_entradas.xlsx` directamente en la carpeta base.

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
└── Consolidado_entradas.xlsx    <- salida generada por el programa
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

## Estado actual

`Consolidado_entradas.xlsx` (antes `Hoja_Medidas.xlsx`) tiene seis hojas:

- `Medidores` — columnas A:U (A:J entrada, K copia de L, L, N, O, R, S, T
  calculadas; M, P, Q, U deliberadamente vacías por diseño).
- `Ofertas SSCC` — las tablas auxiliares equivalentes a `Medidores!W:Y`
  ("Ofertas SSCC por dia") y `Medidores!AB:AE` ("Resumen ventana oferta"),
  una debajo de la otra. El resumen intermedio equivalente a la hoja
  "Resumen Ofertas SSCC" del `.xlsm` original es puramente auxiliar y no se
  persiste.
- `CMg` — copia ordenada de `Cmg/cmg.xlsx` (replica
  `Cargar_CMg_Desde_Archivo`).
- `FD` — datos de `CPF Horario`/`CSF Horario` filtrados por BESS/SAE, más
  sus columnas calculadas (replica `Cargar_SSCC_Desempeno_En_FD`). Los
  bloques CSF (A:M) y CPF (Q:AE) son dos tablas de distinto largo, lado a
  lado en la misma hoja.
- `Subastas` — datos de la hoja `DB` filtrados por BESS/SAE en la columna K,
  más la columna M calculada (replica `Cargar_Remuneracion_Subastas_Rapido`).
  La columna N queda vacía: depende de `Calculo E Costos`, una etapa
  posterior todavía sin implementar.
- `Log` — avisos e incidencias detectadas durante el cálculo.

Las macros de Ofertas SSCC, CMg, FD y Subastas replicadas son solo las de
**carga** de esas hojas; las macros que las consumen después (`Asignar_CMg_
a_Calculos_Turbo`, `Actualizar_Calculos_Columnas`) pertenecen a la etapa
`Calculo E Costos` / `Calculo RE545`, todavía sin implementar.

Validado con casos sintéticos (no con datos reales todavía): ver
`BITACORA.md` → "Pendientes abiertos" para lo que falta antes de dar por
cerrada cada etapa (validación contra un caso real y contra la planilla 11).
