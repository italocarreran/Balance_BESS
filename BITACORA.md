# BITACORA.md — registro de sesiones

Solo se agrega. Nunca se edita ni borra una entrada vieja. La única
excepción es la sección "Pendientes abiertos", que sí se edita porque es un
estado, no un historial.

---

## Pendientes abiertos

- Conseguir el código fuente real (VBA) de `Generar_Resumen_Ofertas_SSCC` y
  `Resumir_Medidores_Central_Ventana_Oferta_Completa`. Sin eso no se pueden
  implementar fielmente las columnas `R, S, T, V, W, X, Y, AB, AC, AD, AE`
  de `Medidores` (plan §16.3, §17, §19.2) — solo se conoce qué columna
  produce cada macro, no su lógica interna.
- Definir el patrón de nombre del archivo de OfertasSSCC dentro de
  `<CARPETA_BASE>/Ofertas/` (la ubicación de la carpeta ya está definida,
  plan §16.2, pero no el nombre del archivo ni su lectura).
- Confirmar la disposición exacta por columnas de la hoja `Diccionario` de
  `Centrales.xlsx` (qué es la columna E, F, G) para poder implementar la
  fórmula de `V` (plan §16.3), que depende de `BUSCARX` contra esas
  columnas.
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

---

## 2026-09-10 (2) — Plan actualizado: AAMM manual y especificación cerrada de columnas

Se recibió una versión actualizada del plan de traspaso
(`docs/Plan_Traspaso_Python_Balance_BESS.md`, secciones 16–18 nuevas) más un
pedido explícito: el archivo de SoC no se llama literalmente
`SOC_AAMM.xlsx`; el período AAMM debe ingresarlo el usuario en un recuadro
de la ventana, no inferirse del nombre de un archivo.

Cambios en `nucleo.py`:

- `buscar_soc(medidas_dir, aamm)` ahora recibe el AAMM como parámetro
  (validado por la nueva `validar_aamm()`) y busca cualquier `.xlsx` dentro
  de `Medidas/` cuyo nombre contenga "SOC" y ese AAMM, en vez de exigir el
  patrón literal `SOC_AAMM.xlsx`. Se probó con un nombre deliberadamente no
  literal (`"resumen soc julio 2607 v2.xlsx"`) y detecta correctamente.
- `revisar_estructura(carpeta_base, aamm)` y `ejecutar(carpeta_base, aamm, ...)`
  ahora reciben el AAMM en vez de extraerlo del archivo de SoC. El checklist
  agrega una fila "Periodo (AAMM)" que bloquea Ejecutar si no son 4 dígitos.
- Se agregó `CARPETA_OFERTAS = "Ofertas"` y `ofertas_dir` en
  `resolver_rutas()`, con una fila no bloqueante en el checklist (ubicación
  ya definida por el plan §16.2; falta patrón de archivo y lectura).
- `LETRA_A_CAMPO` se extendió de A a AE siguiendo la tabla cerrada del plan
  (§16.3). Se corrigió el orden final de columnas para que use
  `list(LETRA_A_CAMPO.values())` en vez de `sorted(LETRA_A_CAMPO)`: con
  claves de dos letras ("AA", "AB"...) el orden alfabético de texto ya no
  coincide con el orden real de columnas de Excel.
- Se implementó `K` = copia de `L` fila a fila (`Copia_Ventana`).
- Se separaron `COLUMNAS_VACIAS` (M, P, Q, U, Z, AA — diseño confirmado,
  no pendiente) de `COLUMNAS_PENDIENTES_OFERTAS` (R, S, T, V, W, X, Y, AB,
  AC, AD, AE — dependen de las macros de Ofertas SSCC).

Cambios en `Balance_BESS.py`:

- Nuevo campo "Periodo del caso (AAMM)" en la ventana (`Entry` limitado a
  4 dígitos), con su propio `StringVar` persistido en `config.json`.
  Cambiar el AAMM (Enter o al perder foco) vuelve a correr la validación
  de estructura.
- `revisar()` y la llamada a `nucleo.ejecutar()` pasan el AAMM ingresado.

**No se implementaron** `R, S, T, V, W, X, Y, AB, AC, AD, AE`: el plan
actualizado (§17, §18) exige que las macros de Ofertas SSCC se repliquen
fielmente, pero esta sesión solo tiene el nombre de esas macros y qué
columna produce cada una — no su código VBA. Implementarlas sin eso sería
adivinar la lógica, lo que el propio plan prohíbe (§18: "no reemplazar
lógica conocida por placeholders" — y en este caso la lógica todavía no es
conocida por el asistente). Quedan documentadas en
`docs/Plan_Traspaso_Python_Balance_BESS.md` §19.2 y en "Pendientes
abiertos" arriba.

Verificación: se armó un caso sintético en el scratchpad (carpeta con
`Medidas_SAE.xlsx`, un archivo de SoC con nombre no literal, y
`Centrales.xlsx`) y se corrió `nucleo.ejecutar()` de punta a punta: genera
31 columnas en el orden correcto, `K` coincide con `L`, y las columnas
vacías/pendientes quedan como `NaN`. No se probó contra un caso real ni
contra la planilla 11 (sigue sin datos reales disponibles en el entorno).
