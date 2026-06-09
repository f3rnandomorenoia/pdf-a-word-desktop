from __future__ import annotations

import argparse
from pathlib import Path

from .converter import convert_pdf_to_word, default_output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Convierte PDF a Word desde GUI o terminal.")
    parser.add_argument("pdf", nargs="?", help="Archivo PDF de entrada.")
    parser.add_argument("output", nargs="?", help="Archivo .docx de salida.")
    parser.add_argument(
        "--mode",
        choices=["table", "visual"],
        default="table",
        help="table reconstruye tablas editables manteniendo mejor el formato; visual intenta clonar la apariencia del PDF.",
    )
    args = parser.parse_args()

    if not args.pdf:
        from .app import main as run_gui

        run_gui()
        return

    pdf_path = Path(args.pdf)
    output_path = Path(args.output) if args.output else default_output_path(pdf_path, "docx")
    output_path = output_path.with_suffix(".docx") if output_path.suffix.lower() != ".docx" else output_path
    result = convert_pdf_to_word(pdf_path, output_path, mode=args.mode)
    print(result)


if __name__ == "__main__":
    main()
