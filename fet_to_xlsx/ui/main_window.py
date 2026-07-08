"""PySide6 graphical interface for the FET to Excel converter."""
from __future__ import annotations

import logging
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QFileDialog, QLabel, QMainWindow, QMessageBox, QPushButton,
    QProgressBar, QTextEdit, QVBoxLayout, QWidget, QHBoxLayout,
)

from fet_to_xlsx.parser.fet_reader import FetReader, FetReaderError
from fet_to_xlsx.parser.models import FetData
from fet_to_xlsx.parser.parser import FetParser
from fet_to_xlsx.parser.solution_finder import SolutionFinder

LOGGER = logging.getLogger(__name__)


class MainWindow(QMainWindow):
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Conversor FET → Excel")
        self.resize(720, 520)
        self.reader = FetReader()
        self.parser = FetParser()
        self.solution_finder = SolutionFinder()
        self.current_path: Path | None = None
        self.data: FetData | None = None
        self.path_label = QLabel("No se ha seleccionado ningún archivo.")
        self.summary = QTextEdit(readOnly=True)
        self.progress = QProgressBar()
        self.export_button = QPushButton("Exportar a Excel")
        self.export_button.setEnabled(False)
        self._setup_ui()

    def _setup_ui(self) -> None:
        select_button = QPushButton("Seleccionar archivo")
        exit_button = QPushButton("Salir")
        select_button.clicked.connect(self.select_file)
        self.export_button.clicked.connect(self.export_file)
        exit_button.clicked.connect(self.close)
        buttons = QHBoxLayout()
        buttons.addWidget(select_button)
        buttons.addWidget(self.export_button)
        buttons.addWidget(exit_button)
        layout = QVBoxLayout()
        layout.addWidget(self.path_label)
        layout.addWidget(QLabel("Resumen del contenido"))
        layout.addWidget(self.summary)
        layout.addWidget(self.progress)
        layout.addLayout(buttons)
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_file(self) -> None:
        file_name, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo FET", "", "FET (*.fet)")
        if not file_name:
            return
        self.progress.setValue(10)
        try:
            root = self.reader.read(file_name)
            self.progress.setValue(45)
            self.data = self.parser.parse(root)
            if not self.data.scheduled_activities:
                self.data.scheduled_activities = self.solution_finder.find_for(file_name, self.data)
            self.current_path = Path(file_name)
            self.path_label.setText(str(self.current_path))
            self.summary.setPlainText(self._summary_text(self.data))
            self.export_button.setEnabled(True)
            self.progress.setValue(100)
        except FetReaderError as exc:
            self._error(str(exc))
        except Exception as exc:  # noqa: BLE001 - GUI boundary must not leak exceptions.
            LOGGER.exception("Unexpected parse error")
            self._error(f"No fue posible analizar el archivo: {exc}")

    def export_file(self) -> None:
        if self.data is None:
            self._error("Primero seleccione un archivo .fet válido.")
            return
        output, _ = QFileDialog.getSaveFileName(self, "Guardar Excel", "horario.xlsx", "Excel (*.xlsx)")
        if not output:
            return
        try:
            from fet_to_xlsx.exporter.excel_exporter import ExcelExporter

            self.progress.setValue(20)
            path = ExcelExporter().export(self.data, output)
            self.progress.setValue(100)
            QMessageBox.information(self, "Exportación completa", f"Archivo generado:\n{path}")
        except Exception as exc:  # noqa: BLE001 - GUI boundary must not leak exceptions.
            LOGGER.exception("Unexpected export error")
            self._error(f"No fue posible exportar el Excel: {exc}")

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
        self.progress.setValue(0)
        QMessageBox.critical(self, "Error", message)


def run_app() -> int:
    """Start the Qt event loop."""
    app = QApplication([])
    window = MainWindow()
    window.show()
    return app.exec()
