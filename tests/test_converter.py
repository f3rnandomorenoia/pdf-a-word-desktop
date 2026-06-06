from pathlib import Path

from pdf_to_word_app.converter import default_output_path


def test_default_output_path_uses_pdf_name(tmp_path: Path) -> None:
    pdf = tmp_path / "entrada.pdf"
    pdf.write_bytes(b"%PDF-1.4\n")

    assert default_output_path(pdf) == tmp_path / "entrada.docx"


def test_default_output_path_does_not_overwrite(tmp_path: Path) -> None:
    pdf = tmp_path / "entrada.pdf"
    pdf.write_bytes(b"%PDF-1.4\n")
    (tmp_path / "entrada.docx").write_text("exists", encoding="utf-8")

    assert default_output_path(pdf) == tmp_path / "entrada-2.docx"
