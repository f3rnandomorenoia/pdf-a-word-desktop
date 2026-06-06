from __future__ import annotations

from pathlib import Path

import fitz


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "ejemplo-contabilidad-parroquia.pdf"


def add_text(page: fitz.Page, rect: fitz.Rect, text: str, size: int = 10, bold: bool = False) -> None:
    page.insert_textbox(
        rect,
        text,
        fontsize=size,
        fontname="helv",
        color=(0, 0, 0),
        align=fitz.TEXT_ALIGN_LEFT,
    )


def draw_cell(page: fitz.Page, rect: fitz.Rect, text: str, size: int = 8, fill: tuple[float, float, float] | None = None) -> None:
    page.draw_rect(rect, color=(0.35, 0.35, 0.35), fill=fill, width=0.45)
    inner = fitz.Rect(rect.x0 + 4, rect.y0 + 4, rect.x1 - 4, rect.y1 - 4)
    page.insert_textbox(inner, text, fontsize=size, fontname="helv", color=(0, 0, 0))


def build_pdf() -> None:
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)

    add_text(page, fitz.Rect(48, 36, 547, 62), "PARROQUIA DE SAN JOSE - RESUMEN CONTABLE PARA TABLON", 15)
    add_text(page, fitz.Rect(48, 64, 547, 84), "Periodo: enero 2026 | Documento de ejemplo para probar conversion PDF a Word", 9)
    page.draw_line((48, 90), (547, 90), color=(0, 0, 0), width=0.8)

    intro = (
        "Este documento simula el tipo de PDF contable que se recibe desde un programa externo. "
        "El objetivo de la conversion es poder abrirlo en Word, rectificar conceptos largos o poco claros, "
        "revisar importes y despues imprimirlo para exponerlo en el tablon de anuncios."
    )
    add_text(page, fitz.Rect(48, 104, 547, 150), intro, 9)

    headers = ["Fecha", "Concepto generado por el programa", "Debe", "Haber", "Concepto rectificado"]
    widths = [58, 210, 62, 62, 155]
    x = 48
    y = 170
    header_h = 28
    for header, width in zip(headers, widths):
        draw_cell(page, fitz.Rect(x, y, x + width, y + header_h), header, 8, fill=(0.88, 0.88, 0.88))
        x += width

    rows = [
        [
            "03/01/26",
            "APUNTE AUTOMATICO 000431 - SERVICIOS GENERALES DIVERSOS NO CLASIFICADOS PARROQUIA DIECISEIS",
            "0,00",
            "124,50",
            "Donativos colecta primer domingo",
        ],
        [
            "07/01/26",
            "CARGO BANCARIO SUMINISTRO ELECTRICO RECINTO ANEXO Y DEPENDENCIAS AUXILIARES",
            "86,42",
            "0,00",
            "Factura electricidad parroquia",
        ],
        [
            "12/01/26",
            "LIQUIDACION MOVIMIENTOS EFECTIVO CAJA MENOR ACTIVIDAD ORDINARIA SIN DETALLE",
            "32,10",
            "0,00",
            "Material limpieza y oficina",
        ],
        [
            "16/01/26",
            "INGRESO TRANSFERENCIA AGRUPADA CONCEPTO NO EDITABLE SEGUN PAQUETE CONTABILIDAD",
            "0,00",
            "210,00",
            "Aportaciones grupo catequesis",
        ],
        [
            "21/01/26",
            "PAGO PROVEEDOR GENERICO MANTENIMIENTO LOCAL SOCIAL PARROQUIAL CODIGO INTERNO 78-A",
            "145,20",
            "0,00",
            "Reparacion puerta salon parroquial",
        ],
        [
            "28/01/26",
            "REGULARIZACION FINAL DE PERIODO POR DIFERENCIAS DE REDONDEO Y AJUSTE AUTOMATICO",
            "1,34",
            "0,00",
            "Ajuste contable menor",
        ],
    ]

    y += header_h
    row_h = 52
    for row in rows:
        x = 48
        for value, width in zip(row, widths):
            align_value = value
            draw_cell(page, fitz.Rect(x, y, x + width, y + row_h), align_value, 7.3)
            x += width
        y += row_h

    add_text(page, fitz.Rect(48, y + 24, 547, y + 46), "Totales del periodo", 11)
    y += 52
    draw_cell(page, fitz.Rect(48, y, 315, y + 30), "Total gastos", 9, fill=(0.94, 0.94, 0.94))
    draw_cell(page, fitz.Rect(315, y, 425, y + 30), "265,06", 9)
    draw_cell(page, fitz.Rect(425, y, 547, y + 30), "EUR", 9)
    y += 30
    draw_cell(page, fitz.Rect(48, y, 315, y + 30), "Total ingresos", 9, fill=(0.94, 0.94, 0.94))
    draw_cell(page, fitz.Rect(315, y, 425, y + 30), "334,50", 9)
    draw_cell(page, fitz.Rect(425, y, 547, y + 30), "EUR", 9)
    y += 30
    draw_cell(page, fitz.Rect(48, y, 315, y + 30), "Saldo del periodo", 9, fill=(0.86, 0.93, 0.86))
    draw_cell(page, fitz.Rect(315, y, 425, y + 30), "69,44", 9, fill=(0.86, 0.93, 0.86))
    draw_cell(page, fitz.Rect(425, y, 547, y + 30), "EUR", 9, fill=(0.86, 0.93, 0.86))

    note = (
        "Notas para la prueba: tras convertir a Word, deberia poder editarse la columna 'Concepto rectificado', "
        "cambiar cualquier texto de concepto y conservar una tabla suficientemente legible para imprimir."
    )
    add_text(page, fitz.Rect(48, y + 48, 547, y + 92), note, 8)

    doc.save(OUTPUT)
    doc.close()
    print(OUTPUT)


if __name__ == "__main__":
    build_pdf()
