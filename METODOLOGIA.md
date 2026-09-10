# METODOLOGIA.md — cómo se trabaja en este repositorio

> Documento vivo. Si en una sesión se descubre algo que la siguiente necesita
> saber, se agrega acá antes de cerrar.

---

## 0. Por qué existe este documento

Este repositorio no existe solo para "guardar el código". Existe para que un
asistente de IA pueda entender el sistema **sin leerlo entero**. Cada regla de
acá abajo sale de esa única necesidad: minimizar cuánto hay que leer para
hacer un cambio correcto.

Si el repositorio lo trabaja más de una persona o más de un asistente con
acceso de escritura, sin verse en tiempo real, este documento es además la
única red de seguridad que existe entre ellos. No son sugerencias.

---

## 1. Qué es este repositorio

Herramienta en Python que reemplaza, hoja por hoja, el cálculo hecho hoy en
`11_PAGOS_BESS_2607_Definitivo.xlsm` (Balance BESS / SSCC). Por ahora
implementa únicamente la primera etapa, **Medidores**: dos scripts
interdependientes que comparten un `config.json` guardado junto al código
(no versionado, es de la herramienta, no del caso).

- `Balance_BESS.py` — ventana tkinter. Único punto de entrada para el
  usuario.
- `nucleo.py` — todo el cálculo, sin dependencias de interfaz. `Balance_BESS.py`
  importa `nucleo` y nunca al revés.

Cada caso a procesar vive en su propia carpeta ("carpeta base"), fuera del
repositorio, con la estructura fija documentada en
`docs/Plan_Traspaso_Python_Balance_BESS.md` §3.3. El repositorio no contiene
datos de ningún caso.

---

## 2. Regla de expansión de contexto

La regla que no se negocia, para cualquier sesión que toque este repositorio:

```
1. Leer MAPA.md                       (mapa corto: qué hace cada script/módulo)
2. Leer de este documento SOLO las secciones que hacen falta
3. Abrir completo ÚNICAMENTE el archivo que se va a modificar
```

**No abrir archivos vecinos "para tener contexto".** Si de verdad hace falta
uno más, pedirlo explícitamente y decir por qué. Leer de más es exactamente
el problema que esta metodología resuelve.

Para las reglas de negocio del cálculo (fórmulas de Excel replicadas, de
dónde sale cada dato, qué columnas siguen pendientes) el documento de
referencia es `docs/Plan_Traspaso_Python_Balance_BESS.md`. Es largo a
propósito y por eso mismo **no se lee entero**: se busca la sección puntual
con grep o por su encabezado (p. ej. "9.4" para la columna `Ventana`).

**Dónde vive el código**, con nombres exactos:

- `Balance_BESS.py` — módulo/entrada principal, se abre siempre primero.
- `nucleo.py` — módulo de cálculo compartido, vive junto al principal en la
  raíz del repositorio, no adentro de una subcarpeta.
- `config.json` — configuración/estado por PC/usuario (última carpeta base
  elegida). No se versiona (ver `.gitignore`). Vive junto a `Balance_BESS.py`
  únicamente porque así lo resuelve `Path(__file__).parent` en el propio
  script; el resto del código nunca debe depender de esa ruta.
- `docs/Plan_Traspaso_Python_Balance_BESS.md` — documento de dominio.

No hay hoy scripts satélite ni un módulo común adicional: solo estos dos
archivos. Si `Balance_BESS.py` se mueve de carpeta, `nucleo.py` debe moverse
con él (import relativo simple, sin paquete).

---

## 3. El set de documentos y para qué sirve cada uno

| Documento | Para qué | Quién lo edita y cuándo |
|---|---|---|
| `REGLAS.md` | Checklist obligatorio de inicio/cierre de sesión. Se lee **entero, primero**, antes que nada. | Se edita poco; es la red de seguridad, no debe crecer sin necesidad. |
| `METODOLOGIA.md` | Cómo se trabaja: regla de expansión, convenciones de código, trampas conocidas, decisiones ya tomadas. | Documento vivo — se actualiza cuando cambia algo estructural o se toma una decisión de diseño. |
| `BITACORA.md` | Registro de qué se hizo en cada sesión y qué quedó pendiente — lo que un `git log` no cuenta. | **Solo se agrega.** Nunca se edita ni borra una entrada vieja. La única excepción es la sección "Pendientes abiertos", que sí se edita porque es un estado, no un historial. |
| `MAPA.md` | Un bloque corto por script: qué hace · consume · produce · expone · depende de. Es la primera lectura de cualquier sesión. | Se actualiza cuando cambia la estructura o el rol de un script. |
| `README.md` | Puerta de entrada: instalación/uso y tabla de "querés X → leé Y". | Se actualiza si cambia la estructura de carpetas o el flujo de instalación. |
| `docs/Plan_Traspaso_Python_Balance_BESS.md` | Reglas de negocio del cálculo: fórmulas, de dónde sale cada dato, qué queda pendiente. Es largo a propósito y no se reescribe por comodidad. | Se corrige ahí mismo cuando el dominio cambia (p. ej. se confirma la lógica de una columna pendiente); nunca se duplica en otro archivo. |

**Regla de prioridad de lectura:** `REGLAS.md` → `METODOLOGIA.md` →
`MAPA.md` → el archivo a modificar → (si hace falta) la sección puntual de
`docs/Plan_Traspaso_Python_Balance_BESS.md`.

**Donde el código y el documento de dominio difieren, manda el código.**
Las diferencias detectadas se anotan en una tabla al final de `MAPA.md`
("Diferencias con el documento de dominio"); al corregir el documento, se
borra la fila correspondiente.

No existe todavía un `INTERFACES.md` generado ni un generador de interfaces:
con dos archivos, `MAPA.md` alcanza. Si el proyecto crece a varios módulos,
reevaluar (ver §10 de la plantilla original de esta metodología).

---

## 4. Cómo se entregan los cambios

Por ahora este repositorio lo trabaja un asistente de IA con revisión
humana (no hay más de un asistente con push directo todavía). Si eso
cambia, esta sección debe actualizarse para documentar cómo se coordinan
entre sí (rama base compartida, `BITACORA.md` como único canal entre
sesiones que no se ven en tiempo real, prefijo de autor en los commits si
alguno autentica con el nombre de otra persona).

- Rama base única (la rama por defecto del repositorio remoto). Todo sale
  de ahí y todo vuelve ahí. Una rama de trabajo se fusiona y se borra; no
  se acumulan ramas vivas.
- Nunca reescribir el historial compartido (`push --force`, `commit --amend`
  sobre algo ya subido, `reset --hard` contra la rama remota).
- Si en algún momento se vuelve a un flujo sin push directo (el usuario sube
  y baja archivos a mano): entregar siempre el archivo completo, nunca un
  diff ni "reemplazá la línea X por esto" — quien lo recibe lo corre tal
  cual, y editar a mano es donde se cuelan los errores.

No hay todavía scripts de apoyo (`scripts/sincronizar.sh`,
`scripts/verificar.sh`). Mientras no existan, la verificación antes de
cerrar una sesión es manual: correr `python -m py_compile Balance_BESS.py
nucleo.py` y, si hay un caso de prueba disponible, `nucleo.ejecutar(...)`
contra él.

---

## 5. Convenciones de código establecidas

- **Interfaz de usuario:** una única ventana tkinter (patrón "carpeta base +
  Examinar + checklist de entradas + Ejecutar + log + barra de progreso +
  contador de tiempo"). Los estados del checklist son exactamente tres:
  `ok` (verde), `falta` (rojo, bloquea Ejecutar) y `pendiente` (ámbar, no
  bloquea). No agregar un cuarto estado sin actualizar `SIMBOLO` y
  `COLOR_ESTADO` en `Balance_BESS.py` a la vez.
- **Persistencia de configuración:** `config.json` junto al `.py`, con una
  clave por PC/usuario (`get_usuario()` = `hostname_usuario`), para que
  varias personas puedan compartir la misma copia del script sin pisarse la
  carpeta recordada. Si el archivo existe pero no se puede leer/parsear, se
  ignora en silencio y se sigue — es preferible perder el ajuste recordado
  a que el programa no abra. La escritura reescribe el `config.json`
  completo (no es atómica todavía; ver §7).
- **Rutas de un caso:** nunca rutas absolutas ni dependientes de
  `Path(__file__).parent` para los archivos de un caso. Todo se deriva de la
  carpeta base vía `resolver_rutas()` en `nucleo.py`. `Path(__file__).parent`
  se usa únicamente para `CONFIG_PATH` en `Balance_BESS.py`.
- **Lectura de Excel:** con `pandas` (`read_excel`/`ExcelFile`) y escritura
  con `openpyxl` como engine. El SoC se extrae por **detección dinámica de
  bloques por encabezados** (`detectar_fila_nombres` → `detectar_bloques` →
  `extraer_soc`): nunca por letra de columna fija ni por offset constante
  entre el nombre de la central y sus columnas `Time Stamp`/`Value` (regla
  del plan, §6.1 de `docs/Plan_Traspaso_Python_Balance_BESS.md`).
- **Normalización de texto:** toda comparación de nombres de central/hoja
  pasa por `normalizar()` en `nucleo.py` (minúsculas, sin tildes, espacios
  colapsados). No reimplementar una variante local de esta función en otro
  archivo.
- **Homologación de nombres:** se resuelve siempre contra la hoja
  `Diccionario` de `Centrales.xlsx` vía `construir_homologacion()`. La
  primera réplica no corrige ni reinterpreta homologaciones aunque parezcan
  desplazadas; cualquier inconsistencia se reporta como aviso/incidencia,
  no se "arregla" en silencio.
- **Errores de entrada vs. errores inesperados:** un problema de datos de
  entrada (archivo faltante, ambigüedad de SOC, columnas faltantes, bloque
  sin `Time Stamp`/`Value`) se señaliza con `nucleo.ErrorEntrada`, con un
  mensaje explicativo para el usuario. No usar excepciones genéricas para
  esto: `Balance_BESS.py` distingue ambos casos para mostrar un mensaje
  distinto.
- **Columnas todavía no definidas:** se agregan igual al DataFrame de salida
  (como `pd.NA`, listadas en `COLUMNAS_PENDIENTES`) en vez de omitirse, para
  que la forma de `Hoja_Medidas.xlsx` sea comparable con `Medidores` de la
  planilla 11 aunque el valor todavía no se calcule.

---

## 6. Generación de `INTERFACES.md`

No aplica todavía: no existe un generador de interfaces en este repositorio.
Con dos archivos (`Balance_BESS.py`, `nucleo.py`) alcanza con `MAPA.md`. Si
se agregan más módulos y esto deja de ser suficiente, documentar acá la
decisión de introducir un generador (o no) antes de empezar a usarlo.

---

## 7. Trampas conocidas

Tabla de solo agregar: cada vez que un bug cueste tiempo real de
investigación, se documenta acá con la regla que lo evita, **antes** de
cerrar la sesión que lo encontró. No se borran filas viejas salvo que la
causa raíz deje de existir en el código.

| Trampa | Regla |
|---|---|
| `guardar_config()` reescribe `config.json` entero sin escritura atómica. Un corte a mitad de escritura puede dejar el archivo corrupto. | Al tocar esa función, evaluar escritura atómica (escribir a un temporal y `rename`). Mientras tanto, `leer_config()` ya tolera un JSON corrupto devolviendo `{}`, así que el peor caso es perder la carpeta recordada, no romper el programa. |
| Si `Medidas/` tiene más de un archivo `SOC_AAMM.xlsx`, `buscar_soc()` lanza `ErrorEntrada` a propósito — no elige el más reciente. | No "arreglar" esto para que elija automáticamente por fecha de modificación: es una decisión deliberada del plan (§3.4) para no tomar en silencio el mes equivocado. |
| Las columnas `K, M, P, Q, R, S, T` de `Medidores` están en `COLUMNAS_PENDIENTES` como `pd.NA` porque su lógica exacta o su fuente (Ofertas SSCC) todavía no está definida. | No inventar una fórmula para completarlas "para que quede bonito". Cerrar primero la regla exacta en `docs/Plan_Traspaso_Python_Balance_BESS.md` §9, con el humano que conoce la planilla 11, y recién ahí implementar. |
| `calcular_ventana()` reinicia el contador por **bloque de filas consecutivas con la misma clave**, no por `groupby` sobre toda la central. | Si los datos de entrada no vienen ordenados por `clave` e `intervalo` antes de llamar a esta función, el resultado no coincide con la fórmula de Excel. `construir_medidores()` ya ordena con `sort_values(["clave", "intervalo"])` antes de calcularla; no quitar ese paso ni reordenar después. |

---

## 8. Decisiones ya tomadas

Lista de solo agregar, para no volver a discutir lo mismo en cada sesión.

- **Primero replicar fielmente la lógica de la planilla 11; después
  simplificar.** No se reinterpretan reglas de negocio todavía no
  documentadas solo porque parezcan mejorables (plan §2, principio 1).
- **La carpeta base es la única selección manual del usuario.** Todo lo
  demás (`Medidas_SAE.xlsx`, `SOC_AAMM.xlsx`, `Centrales.xlsx`, la salida)
  se resuelve por ruta relativa desde ahí, para que mover
  `Balance_BESS.py` a otra ubicación no cambie qué entradas encuentra ni
  dónde escribe (criterio de portabilidad, plan §3.13).
- **Las columnas cuya regla no está confirmada se dejan pendientes
  (`pd.NA`) en vez de adivinarse.** Aplica hoy a `K, M, P, Q, R, S, T` de
  `Medidores`. `OfertasSSCC` en particular ni siquiera tiene ubicación de
  carpeta definida todavía (plan §3.6, §7).
- **El SoC se extrae por semántica de encabezados, nunca por posición
  fija.** Confirmado en el plan §16.1: ni letras de columna ni offsets
  constantes entre el nombre de la central y `Time Stamp`/`Value`, porque
  esa distancia varía entre archivos mensuales.
- **`Hoja_Medidas.xlsx` es un artefacto de validación, no un archivo
  versionado.** Se genera por caso en la carpeta base del usuario y se
  ignora en git (ver `.gitignore`); el repositorio no guarda salidas de
  casos concretos.
