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
├── Ofertas/                 (ubicación definida; lectura aún pendiente)
└── Hoja_Medidas.xlsx        <- salida generada por el programa
```

El nombre del archivo de SoC no sigue un patrón fijo: solo debe contener
"SOC" y el AAMM ingresado en la ventana, en cualquier posición. Si hay más
de un archivo que cumple esa condición, el programa se detiene y pide
dejar solo el del período que corresponde (no elige por fecha de
modificación).

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

Etapa Medidores: las 31 columnas A:AE están definidas según la
especificación cerrada del plan. Implementadas y calculadas: A:J (entrada),
K (copia de L), L, N, O. Deliberadamente vacías (diseño confirmado, no
pendiente): M, P, Q, U, Z, AA. Pendientes porque dependen de las macros de
Ofertas SSCC, cuyo código fuente todavía no se entregó: R, S, T, V, W, X,
Y, AB, AC, AD, AE (ver `BITACORA.md` → "Pendientes abiertos"). La carpeta
`Ofertas/` ya tiene ubicación definida (`<CARPETA_BASE>/Ofertas/`) pero
todavía no se lee ningún archivo de ahí.
