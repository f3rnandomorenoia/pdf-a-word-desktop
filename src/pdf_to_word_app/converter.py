from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

import fitz
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm, Pt, RGBColor, Twips
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
    """Build an editable Word document that stays visually close for table-based PDFs."""
    pdf_path = Path(pdf_path)
    output_path = Path(output_path)
    validate_pdf(pdf_path)

    source = fitz.open(pdf_path)
    document = Document()
    section = document.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    for margin in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(section, margin, Mm(5))
    usable_width_inches = section.page_width.inches - section.left_margin.inches - section.right_margin.inches

    normal_style = document.styles["Normal"]
    normal_style.font.name = "Arial"
    normal_style.font.size = Pt(7.5)

    wrote_anything = False
    for page_index, page in enumerate(source):
        if page_index:
            document.add_page_break()
        tables = page.find_tables().tables
        table_rects = [fitz.Rect(table.bbox) for table in tables]
        text_blocks = _non_table_text_blocks(page, table_rects)

        before_first_table_y = min((rect.y0 for rect in table_rects), default=page.rect.height)
        after_last_table_y = max((rect.y1 for rect in table_rects), default=0)
        for block in [b for b in text_blocks if b[0] < before_first_table_y]:
            _add_paragraph(document, block[1])
            wrote_anything = True

        for table in tables:
            if table.row_count == 0 or table.col_count == 0 or not table.extract():
                continue
            _add_word_table(document, table, usable_width_inches)
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


def _add_word_table(document: Document, source_table, usable_width_inches: float) -> None:
    values = source_table.extract()
    row_count = source_table.row_count
    col_count = source_table.col_count
    x_edges = _x_edges(source_table)
    y_edges = _y_edges(source_table)
    table = document.add_table(rows=row_count, cols=col_count)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_fixed_table_layout(table)

    column_widths, row_heights = _table_dimensions(source_table, usable_width_inches, x_edges, y_edges)
    for col_index, width in enumerate(column_widths):
        table.columns[col_index].width = width
    for row_index, row_height in enumerate(row_heights):
        table.rows[row_index].height = row_height
        table.rows[row_index].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

    for row_index, row in enumerate(source_table.rows):
        for col_index, bbox in enumerate(row.cells):
            if bbox is None:
                continue
            start_row, end_row, start_col, end_col = _span_from_bbox(bbox, x_edges, y_edges)
            if start_row != row_index or start_col != col_index:
                continue
            if start_row == end_row and start_col == end_col:
                continue
            anchor = table.cell(start_row, start_col)
            for merge_row in range(start_row, end_row + 1):
                for merge_col in range(start_col, end_col + 1):
                    if merge_row == start_row and merge_col == start_col:
                        continue
                    anchor = anchor.merge(table.cell(merge_row, merge_col))

    for row_index, row in enumerate(values):
        for col_index in range(col_count):
            bbox = source_table.rows[row_index].cells[col_index]
            if bbox is None:
                continue
            start_row, _, start_col, end_col = _span_from_bbox(bbox, x_edges, y_edges)
            if start_row != row_index or start_col != col_index:
                continue
            value = row[col_index] if col_index < len(row) and row[col_index] is not None else ""
            cell = table.rows[row_index].cells[col_index]
            span_width = Twips(sum(width.twips for width in column_widths[start_col : end_col + 1]))
            _set_cell_width(cell, span_width)
            _write_cell_text(cell, value, row_index, col_index, row_count, col_count)
            _style_cell(cell, row_index, col_index, row_count, col_count)



def _table_dimensions(source_table, usable_width_inches: float, x_edges: list[float], y_edges: list[float]) -> tuple[list[Inches], list[Inches]]:
    table_width_points = max(source_table.bbox[2] - source_table.bbox[0], 1)
    scale = usable_width_inches / (table_width_points / 72.0)
    widths = [Inches(max(((x_edges[i + 1] - x_edges[i]) / 72.0) * scale, 0.18)) for i in range(len(x_edges) - 1)]
    heights = [Inches(max(((y_edges[i + 1] - y_edges[i]) / 72.0) * scale, 0.16)) for i in range(len(y_edges) - 1)]
    return widths, heights


def _span_from_bbox(bbox: tuple[float, float, float, float], x_edges: list[float], y_edges: list[float]) -> tuple[int, int, int, int]:
    x0, y0, x1, y1 = (round(value, 3) for value in bbox)
    start_col = x_edges.index(x0)
    end_col = x_edges.index(x1) - 1
    start_row = y_edges.index(y0)
    end_row = y_edges.index(y1) - 1
    return start_row, end_row, start_col, end_col


def _x_edges(source_table) -> list[float]:
    return sorted({round(cell[0], 3) for cell in source_table.cells} | {round(cell[2], 3) for cell in source_table.cells})


def _y_edges(source_table) -> list[float]:
    return sorted({round(cell[1], 3) for cell in source_table.cells} | {round(cell[3], 3) for cell in source_table.cells})


def _write_cell_text(cell, value: str, row_index: int, col_index: int, row_count: int, col_count: int) -> None:
    lines = [line.strip() for line in str(value).splitlines() if line.strip()]
    if not lines:
        lines = [""]
    first = cell.paragraphs[0]
    _clear_paragraph(first)

    if col_count == 2 and row_count == 2 and row_index == 1 and col_index == 0:
        for idx, line in enumerate(lines):
            paragraph = first if idx == 0 else cell.add_paragraph()
            if line.lower().startswith("título "):
                _add_label_value(paragraph, "Título", line[7:].strip())
            elif line.lower().startswith("subtítulo "):
                _add_label_value(paragraph, "Subtítulo", line[10:].strip())
            else:
                paragraph.add_run(line)
        return

    for idx, line in enumerate(lines):
        paragraph = first if idx == 0 else cell.add_paragraph()
        paragraph.add_run(line)


def _clear_paragraph(paragraph) -> None:
    element = paragraph._element
    for child in list(element):
        element.remove(child)


def _add_label_value(paragraph, label: str, value: str) -> None:
    label_run = paragraph.add_run(f"{label} ")
    label_run.font.size = Pt(7)
    label_run.font.color.rgb = RGBColor(115, 115, 115)
    value_run = paragraph.add_run(value)
    value_run.font.size = Pt(8)


def _style_cell(cell, row_index: int, col_index: int, row_count: int, col_count: int) -> None:
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    _set_cell_margins(cell, top=10, start=15, bottom=10, end=15)
    for paragraph in cell.paragraphs:
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        for run in paragraph.runs:
            run.font.name = "Arial"
            run.font.size = Pt(7)

    if col_count == 2 and row_count == 2:
        _style_header_table(cell, row_index, col_index)
    elif col_count == 9:
        _style_ledger_table(cell, row_index, col_index)
    elif col_count == 3:
        _style_summary_table(cell, row_index, col_index)


def _style_header_table(cell, row_index: int, col_index: int) -> None:
    if row_index == 0 and col_index == 0:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.size = Pt(12)
    elif row_index == 0 and col_index == 1:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in cell.paragraphs[0].runs:
            run.font.size = Pt(8)
    elif row_index == 1 and col_index == 1:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in cell.paragraphs[0].runs:
            run.font.size = Pt(8)
    else:
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT


def _style_ledger_table(cell, row_index: int, col_index: int) -> None:
    numeric_cols = {6, 7, 8}
    centered_cols = {0, 1, 2, 3, 5}
    if row_index == 0:
        _shade_cell(cell, "E6E6E6")
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.size = Pt(6.8)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        return

    if col_index in numeric_cols:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif col_index in centered_cols:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

    for run in cell.paragraphs[0].runs:
        if col_index == 2:
            run.bold = True
            run.font.color.rgb = RGBColor(128, 0, 128)
        if col_index == 8:
            run.bold = True
        if col_index in numeric_cols:
            run.font.size = Pt(6.8)


def _style_summary_table(cell, row_index: int, col_index: int) -> None:
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_index else WD_ALIGN_PARAGRAPH.LEFT
    for run in cell.paragraphs[0].runs:
        if row_index == 0 or col_index > 0:
            run.bold = True

def _set_fixed_table_layout(table) -> None:
    tbl_pr = table._tbl.tblPr
    layout = tbl_pr.first_child_found_in("w:tblLayout")
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")


def _set_cell_width(cell, width) -> None:
    cell.width = width
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.first_child_found_in("w:tcW")
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:type"), "dxa")
    tc_w.set(qn("w:w"), str(width.twips))


def _shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    tc_pr.append(shading)


def _set_cell_margins(cell, top: int, start: int, bottom: int, end: int) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for tag, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{tag}"))
        if node is None:
            node = OxmlElement(f"w:{tag}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def validate_pdf(pdf_path: Path) -> None:
    if not pdf_path.exists():
        raise ConversionError("El archivo PDF no existe.")
    if not pdf_path.is_file():
        raise ConversionError("La ruta seleccionada no es un archivo.")
    if pdf_path.suffix.lower() != ".pdf":
        raise ConversionError("Selecciona un archivo con extension .pdf.")
    if pdf_path.stat().st_size == 0:
        raise ConversionError("El PDF esta vacio.")
