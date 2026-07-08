"""Tkinter fallback interface for environments without PySide6."""
from __future__ import annotations

import logging
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from fet_to_xlsx.parser.fet_reader import FetReader, FetReaderError
from fet_to_xlsx.parser.models import FetData
from fet_to_xlsx.parser.parser import FetParser
from fet_to_xlsx.parser.solution_finder import SolutionFinder

LOGGER = logging.getLogger(__name__)


class TkMainWindow(tk.Tk):
    """Simple desktop window that works with Python's standard library."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Conversor FET → Excel")
        self.geometry("720x520")
        self.reader = FetReader()
        self.parser = FetParser()
        self.solution_finder = SolutionFinder()
        self.current_path: Path | None = None
        self.data: FetData | None = None
        self.path_var = tk.StringVar(value="No se ha seleccionado ningún archivo.")
        self.progress_var = tk.IntVar(value=0)
        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(container, textvariable=self.path_var).pack(anchor=tk.W, pady=(0, 8))
        ttk.Label(container, text="Resumen del contenido").pack(anchor=tk.W)
        self.summary = scrolledtext.ScrolledText(container, height=18, state=tk.DISABLED)
        self.summary.pack(fill=tk.BOTH, expand=True, pady=(4, 8))
        ttk.Progressbar(container, variable=self.progress_var, maximum=100).pack(fill=tk.X, pady=(0, 8))

        buttons = ttk.Frame(container)
        buttons.pack(fill=tk.X)
        ttk.Button(buttons, text="Seleccionar archivo", command=self.select_file).pack(side=tk.LEFT, padx=(0, 8))
        self.export_button = ttk.Button(buttons, text="Exportar a Excel", command=self.export_file, state=tk.DISABLED)
        self.export_button.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(buttons, text="Salir", command=self.destroy).pack(side=tk.LEFT)

    def select_file(self) -> None:
        file_name = filedialog.askopenfilename(title="Seleccionar archivo FET", filetypes=[("FET", "*.fet")])
        if not file_name:
            return
        self.progress_var.set(10)
        try:
            root = self.reader.read(file_name)
            self.progress_var.set(45)
            self.data = self.parser.parse(root)
            if not self.data.scheduled_activities:
                self.data.scheduled_activities = self.solution_finder.find_for(file_name, self.data)
            self.current_path = Path(file_name)
            self.path_var.set(str(self.current_path))
            self._set_summary(self._summary_text(self.data))
            self.export_button.configure(state=tk.NORMAL)
            self.progress_var.set(100)
        except FetReaderError as exc:
            self._error(str(exc))
        except Exception as exc:  # noqa: BLE001 - GUI boundary must not leak exceptions.
            LOGGER.exception("Unexpected parse error")
            self._error(f"No fue posible analizar el archivo: {exc}")

    def export_file(self) -> None:
        if self.data is None:
            self._error("Primero seleccione un archivo .fet válido.")
            return
        output = filedialog.asksaveasfilename(
            title="Guardar Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="horario.xlsx",
        )
        if not output:
            return
        try:
            from fet_to_xlsx.exporter.excel_exporter import ExcelExporter
        except ModuleNotFoundError as exc:
            if exc.name == "openpyxl":
                self._error("Falta la dependencia openpyxl. Instale dependencias con: python -m pip install -r requirements.txt")
                return
            raise
        try:
            self.progress_var.set(20)
            path = ExcelExporter().export(self.data, output)
            self.progress_var.set(100)
            messagebox.showinfo("Exportación completa", f"Archivo generado:\n{path}")
        except Exception as exc:  # noqa: BLE001 - GUI boundary must not leak exceptions.
            LOGGER.exception("Unexpected export error")
            self._error(f"No fue posible exportar el Excel: {exc}")

    def _set_summary(self, text: str) -> None:
        self.summary.configure(state=tk.NORMAL)
        self.summary.delete("1.0", tk.END)
        self.summary.insert(tk.END, text)
        self.summary.configure(state=tk.DISABLED)

    def _summary_text(self, data: FetData) -> str:
        return "\n".join([
            f"Institución: {data.institution.name or 'No especificada'}",
            f"Versión del archivo: {data.institution.version or 'No especificada'}",
            f"Número de docentes: {len(data.teachers)}",
            f"Número de materias: {len(data.subjects)}",
            f"Número de grupos: {len(data.groups)}",
            f"Número de actividades: {len(data.activities)}",
            f"Número de salones: {len(data.rooms)}",
            f"Número de restricciones: {len(data.constraints)}",
        ])

    def _error(self, message: str) -> None:
        self.progress_var.set(0)
        messagebox.showerror("Error", message)


def run_app() -> int:
    """Start the Tkinter event loop."""
    window = TkMainWindow()
    window.mainloop()
    return 0
