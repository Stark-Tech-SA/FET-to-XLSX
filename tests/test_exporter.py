import pytest

openpyxl = pytest.importorskip("openpyxl")
from openpyxl import load_workbook

from fet_to_xlsx.exporter.excel_exporter import ExcelExporter
from fet_to_xlsx.parser.models import Activity, FetData, Institution, Room, Subject, Teacher


def test_exporter_creates_expected_sheets(tmp_path):
    data = FetData(
        institution=Institution(name="Colegio Demo", version="6.0"),
        teachers=[Teacher("Ana")],
        subjects=[Subject("Matemáticas")],
        rooms=[Room("101")],
        activities=[Activity(id="1", teachers=["Ana"], subject="Matemáticas", students=["1A"], duration=1)],
    )
    output = ExcelExporter().export(data, tmp_path / "demo.xlsx")
    workbook = load_workbook(output)
    assert "Información General" in workbook.sheetnames
    assert "Docentes" in workbook.sheetnames
    assert "Restricciones" in workbook.sheetnames
