"""Robust parser for FET XML documents."""
from __future__ import annotations

from collections.abc import Iterable
import xml.etree.ElementTree as ET

from fet_to_xlsx.parser.models import (
    Activity, Constraint, FetData, Institution, Room, ScheduledActivity,
    StudentsGroup, Subject, Teacher,
)


def _text(element: ET.Element | None, default: str = "") -> str:
    return (element.text or default).strip() if element is not None else default


def _int(value: str) -> int | None:
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _bool(value: str) -> bool:
    return value.strip().lower() not in {"false", "0", "no", "n"}


class FetParser:
    """Convert FET XML into typed domain models.

    The parser intentionally keeps unknown constraint fields instead of failing,
    because FET versions may introduce new tags or minor naming differences.
    """

    def parse(self, root: ET.Element) -> FetData:
        data = FetData(
            institution=Institution(
                name=_text(root.find("Institution_Name")),
                version=root.attrib.get("version", _text(root.find("Version"))),
                comments=_text(root.find("Comments")),
            ),
            days=[_text(day.find("Name")) for day in root.findall("./Days_List/Day")],
            hours=[_text(hour.find("Name")) for hour in root.findall("./Hours_List/Hour")],
        )
        data.teachers = self._parse_teachers(root)
        data.subjects = self._parse_subjects(root)
        data.groups = self._parse_students(root)
        data.rooms = self._parse_rooms(root)
        data.activities = self._parse_activities(root)
        data.constraints = self._parse_constraints(root)
        data.scheduled_activities = self._parse_scheduled(root)
        self._attach_teacher_availability(data)
        return data

    def _parse_teachers(self, root: ET.Element) -> list[Teacher]:
        return [Teacher(_text(t.find("Name")), _text(t.find("Comments"))) for t in root.findall("./Teachers_List/Teacher")]

    def _parse_subjects(self, root: ET.Element) -> list[Subject]:
        return [Subject(_text(s.find("Name")), _text(s.find("Comments"))) for s in root.findall("./Subjects_List/Subject")]

    def _parse_students(self, root: ET.Element) -> list[StudentsGroup]:
        groups: list[StudentsGroup] = []
        for year in root.findall("./Students_List/Year"):
            year_name = _text(year.find("Name"))
            groups.append(StudentsGroup(year=year_name, number_of_students=_int(_text(year.find("Number_of_Students")))))
            for group in year.findall("Group"):
                group_name = _text(group.find("Name"))
                groups.append(StudentsGroup(year=year_name, group=group_name, number_of_students=_int(_text(group.find("Number_of_Students")))))
                for subgroup in group.findall("Subgroup"):
                    groups.append(StudentsGroup(
                        year=year_name,
                        group=group_name,
                        subgroup=_text(subgroup.find("Name")),
                        number_of_students=_int(_text(subgroup.find("Number_of_Students"))),
                    ))
        return groups

    def _parse_rooms(self, root: ET.Element) -> list[Room]:
        return [Room(
            name=_text(r.find("Name")),
            capacity=_int(_text(r.find("Capacity"))),
            building=_text(r.find("Building")),
            comments=_text(r.find("Comments")),
        ) for r in root.findall("./Rooms_List/Room")]

    def _parse_activities(self, root: ET.Element) -> list[Activity]:
        activities: list[Activity] = []
        for item in root.findall("./Activities_List/Activity"):
            teachers = [_text(e) for e in item.findall("Teacher") if _text(e)]
            students = [_text(e) for e in item.findall("Students") if _text(e)]
            activities.append(Activity(
                id=_text(item.find("Id")),
                teachers=teachers,
                subject=_text(item.find("Subject")),
                students=students,
                duration=_int(_text(item.find("Duration"))),
                total_duration=_int(_text(item.find("Total_Duration"))),
                student_count=_int(_text(item.find("Number_Of_Students"))),
                active=_bool(_text(item.find("Active"), "true")),
                comments=_text(item.find("Comments")),
            ))
        return activities

    def _parse_constraints(self, root: ET.Element) -> list[Constraint]:
        constraints: list[Constraint] = []
        for category, path in (("Tiempo", "./Time_Constraints_List"), ("Espacio", "./Space_Constraints_List")):
            parent = root.find(path)
            if parent is None:
                continue
            for node in list(parent):
                fields = self._flatten_children(node)
                constraints.append(Constraint(category=category, type=node.tag, weight=str(fields.pop("Weight_Percentage", "")), fields=fields))
        return constraints

    def _parse_scheduled(self, root: ET.Element) -> list[ScheduledActivity]:
        scheduled: list[ScheduledActivity] = []
        for parent in root.findall(".//Activities_Timetable") + root.findall(".//Activity_Timetable"):
            nodes: Iterable[ET.Element] = parent.findall("Activity") if parent.tag != "Activity" else [parent]
            for node in nodes:
                activity_id = _text(node.find("Id")) or _text(node.find("Activity_Id"))
                day = _text(node.find("Day"))
                hour = _text(node.find("Hour"))
                if activity_id and day and hour:
                    scheduled.append(ScheduledActivity(activity_id, day, hour, _text(node.find("Room"))))
        return scheduled

    def _attach_teacher_availability(self, data: FetData) -> None:
        by_name = {teacher.name: teacher for teacher in data.teachers}
        for constraint in data.constraints:
            if "Teacher" in constraint.fields and "notavailable" in constraint.type.lower():
                teacher = by_name.get(str(constraint.fields["Teacher"]))
                if teacher:
                    teacher.availability.append(constraint.fields)

    def _flatten_children(self, node: ET.Element) -> dict[str, object]:
        fields: dict[str, object] = {}
        counters: dict[str, int] = {}
        for child in list(node):
            value: object = self._flatten_children(child) if list(child) else _text(child)
            key = child.tag
            if key in fields:
                counters[key] = counters.get(key, 1) + 1
                existing = fields[key]
                if not isinstance(existing, list):
                    fields[key] = [existing]
                fields[key].append(value)  # type: ignore[union-attr]
            else:
                counters[key] = 1
                fields[key] = value
        return fields
