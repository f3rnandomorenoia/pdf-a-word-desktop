from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

from pdf2docx import Converter


class ConversionError(RuntimeError):
    """Raised when a PDF conversion cannot be completed."""


def default_output_path(pdf_path: Path, extension: str = "docx") -> Path:
    """Return an output path beside the PDF without overwriting existing files."""
    extension = extension.lower().lstrip(".")
    candidate = pdf_path.with_suffix(f".{extension}")
    if not candidate.exists():
        return candidate

    index = 2
    while True:
        candidate = pdf_path.with_name(f"{pdf_path.stem}-{index}.{extension}")
        if not candidate.exists():
            return candidate
        index += 1


def convert_pdf_to_docx(pdf_path: Path, output_path: Path) -> Path:
    pdf_path = Path(pdf_path)
    output_path = Path(output_path)
    validate_pdf(pdf_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    converter = Converter(str(pdf_path))
    try:
        converter.convert(str(output_path))
    finally:
        converter.close()

    if not output_path.exists() or output_path.stat().st_size == 0:
        raise ConversionError("La conversion no genero un archivo DOCX valido.")
    return output_path


def convert_pdf_to_word(pdf_path: Path, output_path: Path) -> Path:
    output_path = Path(output_path)
    extension = output_path.suffix.lower()
    if extension == ".docx":
        return convert_pdf_to_docx(Path(pdf_path), output_path)
    if extension == ".doc":
        return convert_pdf_to_doc(Path(pdf_path), output_path)
    raise ConversionError("El destino debe terminar en .docx o .doc.")


def convert_pdf_to_doc(pdf_path: Path, output_path: Path) -> Path:
    """Create a legacy .doc file by using LibreOffice after DOCX conversion."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise ConversionError(
            "Para generar .doc hace falta LibreOffice instalado. "
            "Prueba con DOCX o instala LibreOffice."
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="pdf-a-word-") as temp_dir:
        temp_docx = Path(temp_dir) / f"{Path(output_path).stem}.docx"
        convert_pdf_to_docx(pdf_path, temp_docx)
        command = [
            soffice,
            "--headless",
            "--convert-to",
            "doc",
            "--outdir",
            str(output_path.parent),
            str(temp_docx),
        ]
        result = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            timeout=120,
        )
        generated = output_path.parent / f"{temp_docx.stem}.doc"
        if result.returncode != 0 or not generated.exists():
            detail = (result.stderr or result.stdout or "").strip()
            raise ConversionError(
                "LibreOffice no pudo generar el archivo .doc."
                + (f" Detalle: {detail}" if detail else "")
            )
        if generated != output_path:
            if output_path.exists():
                output_path.unlink()
            generated.rename(output_path)

    if not output_path.exists() or output_path.stat().st_size == 0:
        raise ConversionError("La conversion no genero un archivo DOC valido.")
    return output_path


def validate_pdf(pdf_path: Path) -> None:
    if not pdf_path.exists():
        raise ConversionError("El archivo PDF no existe.")
    if not pdf_path.is_file():
        raise ConversionError("La ruta seleccionada no es un archivo.")
    if pdf_path.suffix.lower() != ".pdf":
        raise ConversionError("Selecciona un archivo con extension .pdf.")
    if pdf_path.stat().st_size == 0:
        raise ConversionError("El PDF esta vacio.")
