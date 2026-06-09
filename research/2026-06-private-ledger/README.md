# Investigacion PDF privado -> DOCX editable

Fecha: 2026-06-09
Caso: diario contable privado de 2 paginas

## Pregunta

Que pipeline conviene mas para este tipo de PDF si el objetivo es:

- mantener el formato lo mas parecido posible
- seguir pudiendo editar en Word
- no depender de software comercial ni de OCR innecesario

## Muestra analizada

- PDF con capa de texto real, no escaneado
- 2 paginas
- tablas detectables con reglas/bordes claros
- PyMuPDF detecta `2` tablas en la pagina 1 y `3` en la pagina 2

## Experimentos locales

### 1. PyMuPDF `find_tables()`

Comando base:

```bash
cd work/pdf-a-word-desktop
. .venv/bin/activate
python - <<'PY'
from pathlib import Path
import fitz
pdf = Path('/home/fernando/.openclaw/media/inbound/Diario_26-01---d3285e39-1ac2-4cdd-b5de-18c300048d27.pdf')
doc = fitz.open(pdf)
print({'pages': doc.page_count, 'tables_per_page': [len(page.find_tables().tables) for page in doc]})
PY
```

Resultado:

```text
{'pages': 2, 'tables_per_page': [2, 3]}
```

Lectura: el documento ya viene muy bien estructurado para una reconstruccion nativa de tablas en Word.

### 2. `pdftohtml`

Comando usado:

```bash
pdftohtml -c -hidden -noframes INPUT.pdf diario.html
```

Resultado observado:

- genera HTML visualmente fiel
- pero usa `position:absolute`, coordenadas `top/left` y una imagen de fondo por pagina
- no crea tablas HTML semanticas reutilizables

Consecuencia:

- sirve para inspeccion visual o depuracion
- mala base para un `.docx` realmente editable

Nota de privacidad:

- los artefactos generados desde el PDF privado se movieron fuera del repo a `/tmp/pdf-a-word-private-research/`
- en este repo solo quedan conclusiones, no contenido privado

### 3. `pdftohtml -> pandoc -> docx`

Comando base:

```bash
pandoc -f html -t docx diario.html -o diario-pandoc.docx
```

Resultado inspeccionado:

- el `.docx` se genera
- queda con `340` parrafos y `0` tablas
- se pierden completamente la estructura tabular y la maquetacion util para editar

Lectura: convertir primero a HTML y luego a DOCX no resuelve este caso si el HTML intermedio es de posicionamiento absoluto.

## Open source evaluado

## 1. `pdf2docx`

Repo:

- https://github.com/ArtifexSoftware/pdf2docx

Por que importa:

- ya es la base del modo visual actual
- sigue activo y con releases recientes
- la release `v0.5.9` menciona arreglos de tablas vacias y compatibilidad con `PyMuPDF>=1.26.7`

Punto fuerte:

- intenta preservar apariencia general

Punto debil:

- para este tipo de diario el modo visual puede ser lento y no garantiza una edicion limpia de tablas

## 2. `PyMuPDF`

Docs:

- https://pymupdf.readthedocs.io/en/latest/faq/index.html

Por que importa:

- `find_tables()` detecta tablas a partir de lineas y rectangulos sin raster intermedio
- la propia doc sugiere combinar tabla detectada + logica espacial personalizada para casos raros

Punto fuerte:

- encaja muy bien con un PDF de tablas con bordes como este

## 3. `pdfplumber`

Repo:

- https://github.com/jsvine/pdfplumber

Por que importa:

- es una alternativa muy buena para inspeccion fina de texto, lineas y debugging
- la documentacion de PyMuPDF indica que su extraccion de tablas esta portada desde pdfplumber

Punto fuerte:

- muy util como herramienta de contraste o fallback

Punto debil:

- no aporta por si sola una salida DOCX mejor que una reconstruccion hecha a medida

## 4. `Camelot`

Repo:

- https://github.com/camelot-dev/camelot

Por que importa:

- proyecto activo para extraccion tabular
- ofrece varios parsers y metricas de calidad

Punto fuerte:

- muy interesante si aparecen PDFs con tablas mas complicadas o sin bordes claros

Punto debil:

- su salida natural es estructurada (`DataFrame`, CSV, HTML), no DOCX maquetado

## 5. `pdf2htmlEX`

Repo:

- https://github.com/pdf2htmlEX/pdf2htmlEX

Por que importa:

- es la referencia clasica para PDF -> HTML muy fiel

Punto fuerte:

- conserva mucho la apariencia

Punto debil:

- esa fidelidad se apoya precisamente en posicionamiento absoluto y CSS visual
- eso no se traduce bien a DOCX editable

## 6. `html2docx` / `pandoc`

Repos/docs:

- https://github.com/pqzx/html2docx
- https://pandoc.org/

Punto fuerte:

- son utiles cuando el HTML de entrada es semantico

Punto debil:

- no son buena solucion si el HTML viene de un PDF “pintado” con coordenadas absolutas

## Veredicto

## PARTIAL

La idea `PDF -> HTML -> DOCX` solo merece la pena si el HTML intermedio es estructural de verdad. En este caso no lo es: `pdftohtml` y herramientas similares conservan bien la vista, pero no la semantica editable.

Para este PDF concreto, la mejor estrategia no es HTML como paso principal. La via mas prometedora es:

1. detectar tablas nativas desde el PDF
2. reconstruir tablas DOCX con anchos, altos, celdas combinadas, alineaciones y estilos
3. dejar `pdf2docx` visual como modo alternativo/experimental

## Recomendacion de arquitectura

### Opcion recomendada

`PDF estructurado -> PyMuPDF/pdfplumber -> python-docx`

Por que:

- mejor control de editabilidad real
- mejor encaje con diarios contables y listados
- permite reglas especificas por tipo de tabla
- no necesita OCR ni HTML intermedio

### Opcion secundaria

`PDF -> pdf2docx`

Usarla cuando:

- el PDF tenga maquetacion mas libre
- la prioridad sea apariencia aproximada y no tanto editar celdas con limpieza

### Opcion descartada como pipeline principal

`PDF -> HTML -> DOCX`

Descartada para este caso por perdida de estructura semantica.

## Siguiente paso propuesto

Retomar la implementacion del modo principal como “tabla editable y fiel”, pero ya con esta premisa:

- optimizado para PDFs tabulares con texto real
- usando geometria real de tabla
- con fallback visual separado
