# BITACORA.md — registro de sesiones

Solo se agrega. Nunca se edita ni borra una entrada vieja. La única
excepción es la sección "Pendientes abiertos", que sí se edita porque es un
estado, no un historial.

---

## Pendientes abiertos

- Confirmar la lógica exacta de las columnas K y M de `Medidores` (plan
  §9.3, §9.5).
- Confirmar la lógica exacta de las columnas P y Q (plan §9.8).
- Definir la ubicación de `OfertasSSCC` dentro de `<CARPETA_BASE>` y su
  lectura (plan §3.6, §7).
- Replicar `Generar_Resumen_Ofertas_SSCC` y
  `Resumir_Medidores_Central_Ventana_Oferta_Completa` para las columnas R,
  S, T (plan §9.9–9.11).
- Documentar las columnas U en adelante de `Medidores` (plan §9.12).
- Crear casos de prueba para comparar la salida Python contra la hoja
  `Medidores` de `11_PAGOS_BESS_2607_Definitivo.xlsm` (plan §13, punto 10).
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
