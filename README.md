# PDF a Word Desktop

Aplicacion sencilla de escritorio para Windows, Linux y macOS que permite seleccionar un archivo PDF y convertirlo a Word.

## Que hace

- Interfaz grafica con Tkinter.
- Seleccion de archivo PDF desde el ordenador.
- Salida DOCX mediante `pdf2docx`.
- Salida DOC opcional usando LibreOffice en modo headless, si esta instalado.
- Ejecutable Windows generado automaticamente con GitHub Actions.

## Descargar el ejecutable de Windows

Abre la seccion **Releases** del repositorio y descarga `PdfAWord.exe`.

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
python -m pdf_to_word_app documento.pdf documento.docx
```

## Nota sobre PDF a Word

La conversion conserva texto, imagenes y disposicion cuando el PDF lo permite. Si el PDF es una imagen escaneada, antes haria falta OCR; esta primera version no incluye OCR.

Para generar `.doc`, instala LibreOffice. Sin LibreOffice, usa `.docx`, que es el formato moderno de Office.
