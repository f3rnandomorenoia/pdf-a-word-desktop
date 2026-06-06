from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .converter import ConversionError, convert_pdf_to_word, default_output_path


class PdfToWordApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PDF a Word")
        self.geometry("680x360")
        self.minsize(620, 340)

        self.pdf_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.output_format = tk.StringVar(value="docx")
        self.status = tk.StringVar(value="Selecciona un PDF para empezar.")
        self._result_queue: queue.Queue[tuple[str, str]] = queue.Queue()

        self._build_ui()
        self.after(150, self._poll_result_queue)

    def _build_ui(self) -> None:
        container = ttk.Frame(self, padding=18)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)

        title = ttk.Label(container, text="Convertidor PDF a Word", font=("Segoe UI", 18, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))

        subtitle = ttk.Label(
            container,
            text="Convierte un archivo PDF a DOCX. La salida DOC requiere LibreOffice instalado.",
        )
        subtitle.grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 18))

        ttk.Label(container, text="PDF").grid(row=2, column=0, sticky="w", pady=6)
        pdf_entry = ttk.Entry(container, textvariable=self.pdf_path)
        pdf_entry.grid(row=2, column=1, sticky="ew", padx=8, pady=6)
        ttk.Button(container, text="Seleccionar...", command=self._choose_pdf).grid(
            row=2, column=2, sticky="e", pady=6
        )

        ttk.Label(container, text="Salida").grid(row=3, column=0, sticky="w", pady=6)
        output_entry = ttk.Entry(container, textvariable=self.output_path)
        output_entry.grid(row=3, column=1, sticky="ew", padx=8, pady=6)
        ttk.Button(container, text="Guardar como...", command=self._choose_output).grid(
            row=3, column=2, sticky="e", pady=6
        )

        ttk.Label(container, text="Formato").grid(row=4, column=0, sticky="w", pady=6)
        format_frame = ttk.Frame(container)
        format_frame.grid(row=4, column=1, sticky="w", padx=8, pady=6)
        ttk.Radiobutton(
            format_frame,
            text="DOCX",
            value="docx",
            variable=self.output_format,
            command=self._refresh_output_extension,
        ).pack(side="left", padx=(0, 18))
        ttk.Radiobutton(
            format_frame,
            text="DOC",
            value="doc",
            variable=self.output_format,
            command=self._refresh_output_extension,
        ).pack(side="left")

        self.progress = ttk.Progressbar(container, mode="indeterminate")
        self.progress.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(20, 8))

        status_label = ttk.Label(container, textvariable=self.status)
        status_label.grid(row=6, column=0, columnspan=3, sticky="w", pady=(0, 16))

        actions = ttk.Frame(container)
        actions.grid(row=7, column=0, columnspan=3, sticky="e")
        ttk.Label(actions, text=f"v{__version__}").pack(side="left", padx=(0, 18))
        self.convert_button = ttk.Button(actions, text="Convertir", command=self._start_conversion)
        self.convert_button.pack(side="left")

    def _choose_pdf(self) -> None:
        selected = filedialog.askopenfilename(
            title="Seleccionar PDF",
            filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")],
        )
        if not selected:
            return
        pdf_path = Path(selected)
        self.pdf_path.set(str(pdf_path))
        self.output_path.set(str(default_output_path(pdf_path, self.output_format.get())))
        self.status.set("PDF seleccionado. Puedes convertirlo cuando quieras.")

    def _choose_output(self) -> None:
        extension = self.output_format.get()
        selected = filedialog.asksaveasfilename(
            title="Guardar archivo Word",
            defaultextension=f".{extension}",
            filetypes=[
                ("Word DOCX", "*.docx"),
                ("Word 97-2003 DOC", "*.doc"),
                ("Todos los archivos", "*.*"),
            ],
        )
        if selected:
            self.output_path.set(str(Path(selected).with_suffix(f".{extension}")))

    def _refresh_output_extension(self) -> None:
        current = self.output_path.get().strip()
        pdf = self.pdf_path.get().strip()
        extension = self.output_format.get()
        if current:
            self.output_path.set(str(Path(current).with_suffix(f".{extension}")))
        elif pdf:
            self.output_path.set(str(default_output_path(Path(pdf), extension)))

    def _start_conversion(self) -> None:
        pdf = self.pdf_path.get().strip()
        output = self.output_path.get().strip()
        if not pdf:
            messagebox.showwarning("Falta el PDF", "Selecciona primero un archivo PDF.")
            return
        if not output:
            output = str(default_output_path(Path(pdf), self.output_format.get()))
            self.output_path.set(output)

        self._set_busy(True)
        self.status.set("Convirtiendo... Puede tardar si el PDF tiene muchas paginas.")
        thread = threading.Thread(target=self._convert_worker, args=(pdf, output), daemon=True)
        thread.start()

    def _convert_worker(self, pdf: str, output: str) -> None:
        try:
            result = convert_pdf_to_word(Path(pdf), Path(output))
        except ConversionError as exc:
            self._result_queue.put(("error", str(exc)))
        except Exception as exc:  # Defensive UI boundary for unexpected library errors.
            self._result_queue.put(("error", f"Error inesperado: {exc}"))
        else:
            self._result_queue.put(("ok", str(result)))

    def _poll_result_queue(self) -> None:
        try:
            kind, payload = self._result_queue.get_nowait()
        except queue.Empty:
            self.after(150, self._poll_result_queue)
            return

        self._set_busy(False)
        if kind == "ok":
            self.status.set(f"Archivo creado: {payload}")
            messagebox.showinfo("Conversion terminada", f"Archivo creado:\n{payload}")
        else:
            self.status.set(payload)
            messagebox.showerror("No se pudo convertir", payload)
        self.after(150, self._poll_result_queue)

    def _set_busy(self, busy: bool) -> None:
        if busy:
            self.convert_button.state(["disabled"])
            self.progress.start(12)
        else:
            self.convert_button.state(["!disabled"])
            self.progress.stop()


def main() -> None:
    app = PdfToWordApp()
    app.mainloop()


if __name__ == "__main__":
    main()
