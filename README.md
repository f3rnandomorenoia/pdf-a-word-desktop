# PDF a Word Desktop

Aplicacion sencilla de escritorio para Windows, Linux y macOS que permite seleccionar un archivo PDF y convertirlo a Word.

## Que hace

- Interfaz grafica con Tkinter.
- Seleccion de archivo PDF desde el ordenador.
- Salida DOCX en dos modos:
  - **Editable y fiel**: recomendado para contabilidades, diarios y listados con tablas; reconstruye tablas nativas de Word intentando respetar anchos, alturas, celdas combinadas y estilos basicos.
  - **Fiel al PDF (experimental)**: intenta conservar mas la apariencia original, aunque puede tardar mas y en tablas complejas ajustar columnas de forma irregular.
- Salida DOC opcional usando LibreOffice en modo headless, si esta instalado.
- Build automatizada para Windows y macOS con GitHub Actions.

## Descargar las builds

Abre la seccion **Releases** del repositorio y descarga:

- `PdfAWord.exe` para Windows
- `PdfAWord-macOS.zip` para macOS

## Uso desde codigo fuente

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
pip install -e .
python -m pdf_to_word_app
```

Tambien se puede convertir desde terminal:

```bash
python -m pdf_to_word_app documento.pdf documento.docx --mode table
python -m pdf_to_word_app documento.pdf documento-fiel.docx --mode visual
```

## PDF de prueba incluido

En `samples/` hay un caso simulado de contabilidad parroquial:

- `ejemplo-contabilidad-parroquia.pdf`: PDF de entrada, parecido al caso real descrito.
- `ejemplo-contabilidad-parroquia.docx`: salida generada por la aplicacion en modo tabla editable.
- `visual/comparativa-original-vs-docx-editable.png`: comparacion visual entre PDF y DOCX reexportado a PDF.

Para regenerarlo:

```bash
python samples/create_accounting_sample_pdf.py
python -m pdf_to_word_app samples/ejemplo-contabilidad-parroquia.pdf samples/ejemplo-contabilidad-parroquia.docx --mode table
```

## Nota sobre PDF a Word

La conversion conserva texto, imagenes y disposicion cuando el PDF lo permite. Si el PDF es una imagen escaneada, antes haria falta OCR; esta primera version no incluye OCR.

Para PDFs claramente tabulares, el modo principal no pasa por HTML intermedio: reconstruye directamente la estructura del PDF en DOCX para mantener mejor la editabilidad.

Para generar `.doc`, instala LibreOffice. Sin LibreOffice, usa `.docx`, que es el formato moderno de Office.
