from pathlib import Path
from zipfile import ZipFile
import re

from pdf_to_word_app.converter import convert_pdf_to_editable_docx, default_output_path


def test_default_output_path_uses_pdf_name(tmp_path: Path) -> None:
    pdf = tmp_path / "entrada.pdf"
    pdf.write_bytes(b"%PDF-1.4\n")

    assert default_output_path(pdf) == tmp_path / "entrada.docx"


def test_default_output_path_does_not_overwrite(tmp_path: Path) -> None:
    pdf = tmp_path / "entrada.pdf"
    pdf.write_bytes(b"%PDF-1.4\n")
    (tmp_path / "entrada.docx").write_text("exists", encoding="utf-8")

    assert default_output_path(pdf) == tmp_path / "entrada-2.docx"


def test_editable_table_conversion_creates_word_table(tmp_path: Path) -> None:
    source = Path("samples/ejemplo-contabilidad-parroquia.pdf")
    output = tmp_path / "salida.docx"

    convert_pdf_to_editable_docx(source, output)

    assert output.exists()
    with ZipFile(output) as docx:
        document_xml = docx.read("word/document.xml").decode("utf-8")
    assert "Concepto generado por el programa" in document_xml
    assert "<w:tbl>" in document_xml
    grid_widths = [int(value) for value in re.findall(r'<w:gridCol w:w="(\d+)"/>', document_xml)]
    tc_widths = [int(value) for value in re.findall(r'<w:tcW w:type="dxa" w:w="(\d+)"/>', document_xml)]
    assert 1260 not in tc_widths
    assert any(width in tc_widths for width in grid_widths)
