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
2. La ventana detecta automáticamente las entradas y muestra un checklist
   (`OK` / `FALTA` / `PENDIENTE`).
3. El botón **Ejecutar** se habilita cuando no falta nada requerido y genera
   `Hoja_Medidas.xlsx` directamente en la carpeta base.

## Estructura de carpeta de un caso

```text
<CARPETA_BASE>/
├── Medidas/
│   ├── Medidas_SAE.xlsx
│   └── SOC_AAMM.xlsx        (ej. SOC_2607.xlsx — debe haber exactamente uno)
├── Auxiliares/
│   └── Centrales.xlsx       (hojas "Resumen BESS" y "Diccionario")
└── Hoja_Medidas.xlsx        <- salida generada por el programa
```

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

Etapa Medidores: columnas A:J, L, N, O implementadas. Las columnas K, M, P,
Q, R, S, T quedan pendientes (ver `BITACORA.md` → "Pendientes abiertos").
`OfertasSSCC` todavía no tiene ubicación definida dentro de la carpeta del
caso.
