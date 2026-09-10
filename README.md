# Balance_BESS

Herramienta en Python que reemplaza, hoja por hoja, el cálculo hecho hoy en
`11_PAGOS_BESS_2607_Definitivo.xlsm` (Balance BESS / SSCC). La planilla se
usa solo como referencia de validación; el proceso Python no depende de
información almacenada exclusivamente en ese libro.

Etapa implementada hasta ahora: **Medidores**.

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
   `Hoja_Medidas.xlsx` directamente en la carpeta base.

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
└── Hoja_Medidas.xlsx        <- salida generada por el programa
```

Ni el archivo de SoC ni el de OfertasSSCC siguen un nombre fijo:

- **SoC**: cualquier `.xlsx` en `Medidas/` cuyo nombre contenga "SOC" y el
  AAMM ingresado en la ventana. Si hay más de un archivo que cumple la
  condición, el programa se detiene y pide dejar solo el del período
  correspondiente (no elige por fecha de modificación).
- **OfertasSSCC**: cualquier archivo Excel en `Ofertas/` cuyo nombre
  contenga "OfertasSSCC". Si hay más de uno, a diferencia del SoC, se toma
  automáticamente el más reciente por fecha de modificación (así lo hace
  la macro original de la planilla).

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

Etapa Medidores completa: columnas A:U de `Medidores` implementadas
(A:J entrada, K copia de L, L, N, O, R, S, T calculadas; M, P, Q, U
deliberadamente vacías por diseño). Las macros de Ofertas SSCC
(`Generar_Resumen_Ofertas_SSCC`, `Resumir_Medidores_Central_Ventana_
Oferta_Completa`) están replicadas a partir del código VBA original.
`Hoja_Medidas.xlsx` tiene tres hojas:

- `Medidores` — la tabla A:U, una fila por registro.
- `Ofertas SSCC` — las tablas auxiliares equivalentes a `Medidores!W:Y`
  ("Ofertas SSCC por dia") y `Medidores!AB:AE` ("Resumen ventana oferta"),
  una debajo de la otra con su propio título. El resumen intermedio
  equivalente a la hoja "Resumen Ofertas SSCC" del `.xlsm` original es
  puramente auxiliar y no se persiste.
- `Log` — avisos e incidencias detectadas durante el cálculo.

Validado con un caso sintético (no con datos reales todavía): ver
`BITACORA.md` → "Pendientes abiertos" para lo que falta antes de dar por
cerrada la etapa (validación contra un caso real y contra la planilla 11).
