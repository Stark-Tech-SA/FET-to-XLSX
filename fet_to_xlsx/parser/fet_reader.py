"""Low-level XML loading and validation for .fet files."""
from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET

from fet_to_xlsx.parser.xml_utils import descendant, normalized_name


class FetReaderError(Exception):
    """Raised when a FET file cannot be read or validated."""


class FetReader:
    """Read a FET XML document from disk without depending on FET itself."""

    def read(self, path: str | Path) -> ET.Element:
        file_path = Path(path)
        if not file_path.exists():
            raise FetReaderError(f"El archivo no existe: {file_path}")
        if not file_path.is_file():
            raise FetReaderError(f"La ruta no corresponde a un archivo: {file_path}")
        if file_path.suffix.lower() != ".fet":
            raise FetReaderError("Seleccione un archivo con extensión .fet")
        try:
            root = ET.parse(file_path).getroot()
        except ET.ParseError as exc:
            raise FetReaderError(f"XML inválido o archivo corrupto: {exc}") from exc
        if normalized_name(root.tag) != "fet" and descendant(root, "Institution_Name") is None:
            raise FetReaderError("La estructura no parece compatible con FET.")
        return root
