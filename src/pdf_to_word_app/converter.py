from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

import fitz
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
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


def convert_pdf_to_word(pdf_path: Path, output_path: Path, mode: str = "table") -> Path:
    output_path = Path(output_path)
    extension = output_path.suffix.lower()
    if extension == ".docx":
        if mode == "table":
            return convert_pdf_to_editable_docx(Path(pdf_path), output_path)
        return convert_pdf_to_docx(Path(pdf_path), output_path)
    if extension == ".doc":
        return convert_pdf_to_doc(Path(pdf_path), output_path, mode=mode)
    raise ConversionError("El destino debe terminar en .docx o .doc.")


def convert_pdf_to_doc(pdf_path: Path, output_path: Path, mode: str = "table") -> Path:
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
        if mode == "table":
            convert_pdf_to_editable_docx(pdf_path, temp_docx)
        else:
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


def convert_pdf_to_editable_docx(pdf_path: Path, output_path: Path) -> Path:
    """Build a clean Word document, prioritizing editable tables over exact layout."""
    pdf_path = Path(pdf_path)
    output_path = Path(output_path)
    validate_pdf(pdf_path)

    source = fitz.open(pdf_path)
    document = Document()
    page_tables = [page.find_tables().tables for page in source]
    max_col_count = max(
        (max((len(row) for row in table.extract()), default=0) for tables in page_tables for table in tables),
        default=0,
    )
    section = document.sections[0]
    if max_col_count > 5:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
    for margin in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(section, margin, Inches(0.45))

    normal_style = document.styles["Normal"]
    normal_style.font.name = "Arial"
    normal_style.font.size = Pt(9)

    wrote_anything = False
    for page_index, page in enumerate(source):
        if page_index:
            document.add_page_break()
        tables = page_tables[page_index]
        table_rects = [fitz.Rect(table.bbox) for table in tables]
        text_blocks = _non_table_text_blocks(page, table_rects)

        before_first_table_y = min((rect.y0 for rect in table_rects), default=page.rect.height)
        after_last_table_y = max((rect.y1 for rect in table_rects), default=0)
        for block in [b for b in text_blocks if b[0] < before_first_table_y]:
            _add_paragraph(document, block[1])
            wrote_anything = True

        for table in tables:
            rows = table.extract()
            if not rows:
                continue
            _add_word_table(document, rows)
            wrote_anything = True

        for block in [b for b in text_blocks if b[0] > after_last_table_y]:
            _add_paragraph(document, block[1])
            wrote_anything = True

    source.close()
    if not wrote_anything:
        return convert_pdf_to_docx(pdf_path, output_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    if not output_path.exists() or output_path.stat().st_size == 0:
        raise ConversionError("La conversion no genero un archivo DOCX valido.")
    return output_path


def _non_table_text_blocks(page: fitz.Page, table_rects: list[fitz.Rect]) -> list[tuple[float, str]]:
    blocks: list[tuple[float, str]] = []
    for block in page.get_text("blocks"):
        rect = fitz.Rect(block[:4])
        text = " ".join(str(block[4]).split())
        if not text:
            continue
        if any(rect.intersects(table_rect) for table_rect in table_rects):
            continue
        blocks.append((rect.y0, text))
    return sorted(blocks, key=lambda item: item[0])


def _add_paragraph(document: Document, text: str) -> None:
    paragraph = document.add_paragraph(text)
    if len(text) < 90 and text.upper() == text:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in paragraph.runs:
            run.bold = True
            run.font.size = Pt(14)
    else:
        paragraph.paragraph_format.space_after = Pt(8)


def _add_word_table(document: Document, rows: list[list[str | None]]) -> None:
    col_count = max(len(row) for row in rows)
    table = document.add_table(rows=len(rows), cols=col_count)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    widths = _column_widths(col_count)
    _set_fixed_table_layout(table)
    for col_index, width in enumerate(widths):
        table.columns[col_index].width = width
    has_header = any((value or "").strip().lower() == "fecha" for value in rows[0])
    for row_index, row in enumerate(rows):
        if has_header:
            row_height = 0.34 if row_index == 0 else 0.62
        else:
            row_height = 0.28
        table.rows[row_index].height = Inches(row_height)
        table.rows[row_index].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        for col_index in range(col_count):
            cell = table.cell(row_index, col_index)
            value = row[col_index] if col_index < len(row) and row[col_index] is not None else ""
            cell.text = str(value).replace("\n", " ")
            _set_cell_width(cell, widths[col_index])
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)
                for run in paragraph.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(8)
                    if has_header and row_index == 0:
                        run.bold = True
            if has_header and row_index == 0:
                _shade_cell(cell, "E6E6E6")
    document.add_paragraph()


def _column_widths(col_count: int) -> list[Inches]:
    if col_count == 5:
        return [Inches(0.7), Inches(2.55), Inches(0.65), Inches(0.65), Inches(2.0)]
    if col_count == 3:
        return [Inches(3.8), Inches(1.35), Inches(1.0)]
    available_width = 10.0
    return [Inches(available_width / max(col_count, 1)) for _ in range(col_count)]


def _set_cell_width(cell, width) -> None:
    cell.width = width
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_w = tc_pr.first_child_found_in("w:tcW")
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(int(width.inches * 1440)))
    tc_w.set(qn("w:type"), "dxa")


def _set_fixed_table_layout(table) -> None:
    tbl_pr = table._tbl.tblPr
    layout = tbl_pr.first_child_found_in("w:tblLayout")
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")


def _shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    tc_pr.append(shading)


def validate_pdf(pdf_path: Path) -> None:
    if not pdf_path.exists():
        raise ConversionError("El archivo PDF no existe.")
    if not pdf_path.is_file():
        raise ConversionError("La ruta seleccionada no es un archivo.")
    if pdf_path.suffix.lower() != ".pdf":
        raise ConversionError("Selecciona un archivo con extension .pdf.")
    if pdf_path.stat().st_size == 0:
        raise ConversionError("El PDF esta vacio.")
