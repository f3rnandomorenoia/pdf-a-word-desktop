from __future__ import annotations

from pathlib import Path

import fitz
from PIL import Image, ImageChops, ImageDraw


ROOT = Path(__file__).resolve().parent
VISUAL = ROOT / "visual"


def render_first_page(pdf_path: Path, output_path: Path) -> Image.Image:
    doc = fitz.open(pdf_path)
    pix = doc[0].get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
    pix.save(output_path)
    doc.close()
    return Image.open(output_path).convert("RGB")


def build_comparison() -> None:
    VISUAL.mkdir(exist_ok=True)
    original = render_first_page(ROOT / "ejemplo-contabilidad-parroquia.pdf", VISUAL / "original-pdf.png")
    converted_pdf = VISUAL / "ejemplo-contabilidad-parroquia.pdf"
    converted = render_first_page(converted_pdf, VISUAL / "docx-editable-rendered.png")

    width = max(original.width, converted.width)
    height = max(original.height, converted.height)
    pad = 40
    label_height = 46
    canvas = Image.new("RGB", (width * 2 + pad * 3, height + label_height + pad * 2), "white")
    draw = ImageDraw.Draw(canvas)
    draw.text((pad, 16), "PDF original", fill="black")
    right_x = pad * 2 + width
    draw.text((right_x, 16), "DOCX modo tabla editable, reexportado a PDF", fill="black")
    canvas.paste(original, (pad, label_height))
    canvas.paste(converted, (right_x, label_height))
    canvas.save(VISUAL / "comparativa-original-vs-docx-editable.png")

    left = original.resize((width, height))
    right = converted.resize((width, height))
    diff = ImageChops.difference(left, right).convert("L")
    diff = diff.point(lambda pixel: 255 if pixel > 18 else 0)
    diff_rgb = Image.new("RGB", diff.size, "white")
    red = Image.new("RGB", diff.size, (220, 0, 0))
    diff_rgb.paste(red, mask=diff)
    diff_rgb.save(VISUAL / "diferencias-original-vs-docx-editable.png")


if __name__ == "__main__":
    build_comparison()
