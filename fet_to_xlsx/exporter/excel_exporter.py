"""Excel exporter for parsed FET data."""
from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path
import re
from typing import Iterable

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from fet_to_xlsx.exporter.formatting import style_table
from fet_to_xlsx.parser.models import Activity, FetData, ScheduledActivity


class ExcelExporter:
    """Create a formatted .xlsx workbook from typed FET data."""

    WEEKDAYS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]

    def export(self, data: FetData, output_path: str | Path) -> Path:
        path = Path(output_path)
        if path.suffix.lower() != ".xlsx":
            path = path.with_suffix(".xlsx")
        wb = Workbook()
        self._general(wb.active, data)
        self._teachers(wb.create_sheet("Docentes"), data)
        self._subjects(wb.create_sheet("Materias"), data)
        self._groups(wb.create_sheet("Grupos"), data)
        self._rooms(wb.create_sheet("Salones"), data)
        self._activities(wb.create_sheet("Actividades"), data)
        self._constraints(wb.create_sheet("Restricciones"), data)
        self._resource_timetables(wb, data, "teacher")
        self._resource_timetables(wb, data, "group")
        self._resource_timetables(wb, data, "room")
        wb.save(path)
        return path

    def _general(self, ws: Worksheet, data: FetData) -> None:
        ws.title = "Información General"
        rows = [
            ("Campo", "Valor"),
            ("Institución", data.institution.name),
            ("Versión", data.institution.version),
            ("Fecha de exportación", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ("Número de docentes", len(data.teachers)),
            ("Número de estudiantes", sum(g.number_of_students or 0 for g in data.groups)),
            ("Número de grupos", len(data.groups)),
            ("Número de actividades", len(data.activities)),
            ("Número de materias", len(data.subjects)),
            ("Número de salones", len(data.rooms)),
            ("Número de restricciones", len(data.constraints)),
        ]
        for row in rows:
            ws.append(row)
        style_table(ws)

    def _teachers(self, ws: Worksheet, data: FetData) -> None:
        ws.append(["Nombre", "Comentarios", "Disponibilidad"])
        for teacher in data.teachers:
            ws.append([teacher.name, teacher.comments, self._stringify(teacher.availability)])
        style_table(ws)

    def _subjects(self, ws: Worksheet, data: FetData) -> None:
        ws.append(["Nombre", "Comentarios"])
        for subject in data.subjects:
            ws.append([subject.name, subject.comments])
        style_table(ws)

    def _groups(self, ws: Worksheet, data: FetData) -> None:
        ws.append(["Año", "Grupo", "Subgrupo", "Número de estudiantes"])
        for group in data.groups:
            ws.append([group.year, group.group, group.subgroup, group.number_of_students])
        style_table(ws)

    def _rooms(self, ws: Worksheet, data: FetData) -> None:
        ws.append(["Nombre", "Capacidad", "Edificio", "Comentarios"])
        for room in data.rooms:
            ws.append([room.name, room.capacity, room.building, room.comments])
        style_table(ws)

    def _activities(self, ws: Worksheet, data: FetData) -> None:
        ws.append(["ID", "Docente", "Materia", "Grupo", "Duración", "Cantidad de estudiantes", "Actividad activa", "Comentarios"])
        for activity in data.activities:
            ws.append([
                activity.id,
                ", ".join(activity.teachers),
                activity.subject,
                ", ".join(activity.students),
                activity.duration,
                activity.student_count,
                "Sí" if activity.active else "No",
                activity.comments,
            ])
        style_table(ws)

    def _constraints(self, ws: Worksheet, data: FetData) -> None:
        ws.append(["Categoría", "Tipo", "Peso", "Campos"])
        for constraint in data.constraints:
            ws.append([constraint.category, constraint.type, constraint.weight, self._stringify(constraint.fields)])
        style_table(ws)

    def _resource_timetables(self, wb: Workbook, data: FetData, mode: str) -> None:
        resources = self._resources(data, mode)
        if not resources:
            self._empty_timetable(wb.create_sheet(self._base_title(mode)), data)
            return
        for resource in resources:
            ws = wb.create_sheet(self._unique_title(wb, f"{self._prefix(mode)} {resource}"))
            self._timetable(ws, data, mode, resource)

    def _empty_timetable(self, ws: Worksheet, data: FetData) -> None:
        ws.append(["Hora", *self._days(data)])
        ws.append(["Sin recursos definidos", *[""] * len(self._days(data))])
        style_table(ws)

    def _timetable(self, ws: Worksheet, data: FetData, mode: str, resource: str) -> None:
        ws.append(["Hora", *self._days(data)])
        if not data.scheduled_activities:
            ws.append(["Sin solución de horario incluida", *[""] * len(self._days(data))])
            style_table(ws)
            return
        activities = {activity.id: activity for activity in data.activities}
        placements = self._placements(data.scheduled_activities)
        for hour in data.hours or ["Hora"]:
            row = [hour]
            for day in self._days(data):
                row.append(self._cell_text(resource, day, hour, placements, activities, mode))
            ws.append(row)
        style_table(ws)
        self._merge_duration_cells(ws, data, activities, mode, resource)

    def _merge_duration_cells(
        self,
        ws: Worksheet,
        data: FetData,
        activities: dict[str, Activity],
        mode: str,
        resource: str,
    ) -> None:
        if not data.hours:
            return
        days = self._days(data)
        for placement in data.scheduled_activities:
            activity = activities.get(placement.activity_id)
            if not activity or not self._matches(resource, placement, activity, mode):
                continue
            duration = max(activity.duration or 1, 1)
            if duration <= 1 or placement.day not in days or placement.hour not in data.hours:
                continue
            col = 2 + days.index(placement.day)
            row = 2 + data.hours.index(placement.hour)
            end_row = min(row + duration - 1, 1 + len(data.hours))
            if end_row > row:
                ws.merge_cells(start_row=row, start_column=col, end_row=end_row, end_column=col)

    def _cell_text(
        self,
        resource: str,
        day: str,
        hour: str,
        placements: dict[tuple[str, str], list[ScheduledActivity]],
        activities: dict[str, Activity],
        mode: str,
    ) -> str:
        parts: list[str] = []
        for placement in placements.get((day, hour), []):
            activity = activities.get(placement.activity_id)
            if activity and self._matches(resource, placement, activity, mode):
                parts.append(self._format_activity(activity, placement, mode))
        return "\n---\n".join(parts)

    def _format_activity(self, activity: Activity, placement: ScheduledActivity, mode: str) -> str:
        lines = [activity.subject]
        if mode != "teacher" and activity.teachers:
            lines.append(", ".join(activity.teachers))
        if mode != "group" and activity.students:
            lines.append(", ".join(activity.students))
        if mode != "room" and placement.room:
            lines.append(placement.room)
        return "\n".join(line for line in lines if line)

    def _matches(self, resource: str, placement: ScheduledActivity, activity: Activity, mode: str) -> bool:
        return (
            (mode == "teacher" and resource in activity.teachers)
            or (mode == "group" and resource in activity.students)
            or (mode == "room" and resource == placement.room)
        )

    def _resources(self, data: FetData, mode: str) -> list[str]:
        if mode == "teacher":
            return [teacher.name for teacher in data.teachers]
        if mode == "group":
            return sorted({group.subgroup or group.group or group.year for group in data.groups if group.subgroup or group.group or group.year})
        return [room.name for room in data.rooms]

    def _placements(self, scheduled: Iterable[ScheduledActivity]) -> dict[tuple[str, str], list[ScheduledActivity]]:
        placements: dict[tuple[str, str], list[ScheduledActivity]] = defaultdict(list)
        for placement in scheduled:
            placements[(placement.day, placement.hour)].append(placement)
        return placements

    def _days(self, data: FetData) -> list[str]:
        return data.days or self.WEEKDAYS

    def _base_title(self, mode: str) -> str:
        return {"teacher": "Horario Docente", "group": "Horario Grupo", "room": "Horario Salón"}[mode]

    def _prefix(self, mode: str) -> str:
        return {"teacher": "Docente", "group": "Grupo", "room": "Salón"}[mode]

    def _unique_title(self, wb: Workbook, title: str) -> str:
        safe_title = re.sub(r"[\\/*?:\[\]]", "-", title).strip() or "Horario"
        safe_title = safe_title[:31]
        if safe_title not in wb.sheetnames:
            return safe_title
        for index in range(1, 1000):
            suffix = f" {index}"
            candidate = f"{safe_title[:31 - len(suffix)]}{suffix}"
            if candidate not in wb.sheetnames:
                return candidate
        return safe_title[:28] + " 999"

    def _stringify(self, value: object) -> str:
        return str(value).replace("{", "").replace("}", "") if value else ""
