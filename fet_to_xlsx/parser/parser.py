"""Robust parser for FET XML documents."""
from __future__ import annotations

from collections.abc import Iterable
import xml.etree.ElementTree as ET

from fet_to_xlsx.parser.models import (
    Activity, Constraint, FetData, Institution, Room, ScheduledActivity,
    StudentsGroup, Subject, Teacher,
)
from fet_to_xlsx.parser.xml_utils import child, children, descendant, descendants, local_name, text


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
    because FET versions may introduce new tags or minor naming differences. It
    also ignores XML namespaces and underscore/case differences in known tags.
    """

    def parse(self, root: ET.Element) -> FetData:
        data = FetData(
            institution=Institution(
                name=text(descendant(root, "Institution_Name", "InstitutionName")),
                version=root.attrib.get("version", text(descendant(root, "Version"))),
                comments=text(descendant(root, "Comments")),
            ),
            days=[text(child(day, "Name")) for day in self._list_items(root, "Days_List", "Day")],
            hours=[text(child(hour, "Name")) for hour in self._list_items(root, "Hours_List", "Hour")],
        )
        data.teachers = self._parse_teachers(root)
        data.subjects = self._parse_subjects(root)
        data.groups = self._parse_students(root)
        data.rooms = self._parse_rooms(root)
        data.activities = self._parse_activities(root)
        data.constraints = self._parse_constraints(root)
        data.scheduled_activities = self._parse_scheduled(root, data.constraints)
        scheduled_ids = {placement.activity_id for placement in data.scheduled_activities}
        data.unscheduled_activity_ids = [activity.id for activity in data.activities if activity.id not in scheduled_ids]
        self._attach_teacher_availability(data)
        return data

    def _parse_teachers(self, root: ET.Element) -> list[Teacher]:
        return [
            Teacher(text(child(teacher, "Name")), text(child(teacher, "Comments")))
            for teacher in self._list_items(root, "Teachers_List", "Teacher")
            if text(child(teacher, "Name"))
        ]

    def _parse_subjects(self, root: ET.Element) -> list[Subject]:
        return [
            Subject(text(child(subject, "Name")), text(child(subject, "Comments")))
            for subject in self._list_items(root, "Subjects_List", "Subject")
            if text(child(subject, "Name"))
        ]

    def _parse_students(self, root: ET.Element) -> list[StudentsGroup]:
        groups: list[StudentsGroup] = []
        for year in self._list_items(root, "Students_List", "Year"):
            year_name = text(child(year, "Name"))
            groups.append(StudentsGroup(year=year_name, number_of_students=_int(text(child(year, "Number_of_Students")))))
            for group in children(year, "Group"):
                group_name = text(child(group, "Name"))
                groups.append(StudentsGroup(
                    year=year_name,
                    group=group_name,
                    number_of_students=_int(text(child(group, "Number_of_Students"))),
                ))
                for subgroup in children(group, "Subgroup"):
                    groups.append(StudentsGroup(
                        year=year_name,
                        group=group_name,
                        subgroup=text(child(subgroup, "Name")),
                        number_of_students=_int(text(child(subgroup, "Number_of_Students"))),
                    ))
        return groups

    def _parse_rooms(self, root: ET.Element) -> list[Room]:
        return [
            Room(
                name=text(child(room, "Name")),
                capacity=_int(text(child(room, "Capacity"))),
                building=text(child(room, "Building")),
                comments=text(child(room, "Comments")),
            )
            for room in self._list_items(root, "Rooms_List", "Room")
            if text(child(room, "Name"))
        ]

    def _parse_activities(self, root: ET.Element) -> list[Activity]:
        activities: list[Activity] = []
        for item in self._list_items(root, "Activities_List", "Activity"):
            teachers = [text(teacher) for teacher in children(item, "Teacher") if text(teacher)] or ["SIN DOCENTE"]
            students = [text(student) for student in children(item, "Students", "Student") if text(student)] or ["SIN GRUPO"]
            activity_id = text(child(item, "Id", "Activity_Id"))
            if not activity_id:
                continue
            activities.append(Activity(
                id=activity_id,
                teachers=teachers,
                subject=text(child(item, "Subject")),
                students=students,
                duration=_int(text(child(item, "Duration"))),
                total_duration=_int(text(child(item, "Total_Duration"))),
                student_count=_int(text(child(item, "Number_Of_Students"))),
                active=_bool(text(child(item, "Active"), "true")),
                comments=text(child(item, "Comments")),
            ))
        return activities

    def _parse_constraints(self, root: ET.Element) -> list[Constraint]:
        constraints: list[Constraint] = []
        for category, list_name in (("Tiempo", "Time_Constraints_List"), ("Espacio", "Space_Constraints_List")):
            parent = descendant(root, list_name)
            if parent is None:
                continue
            for node in list(parent):
                fields = self._flatten_children(node)
                constraints.append(Constraint(
                    category=category,
                    type=local_name(node.tag),
                    weight=str(fields.pop("Weight_Percentage", fields.pop("Weight", ""))),
                    fields=fields,
                ))
        return constraints

    def _parse_scheduled(self, root: ET.Element, constraints: list[Constraint]) -> list[ScheduledActivity]:
        scheduled = self._parse_solution_nodes(root)
        if not scheduled:
            scheduled = self._parse_locked_time_constraints(constraints)
        self._apply_room_constraints(scheduled, constraints)
        return scheduled

    def _parse_solution_nodes(self, root: ET.Element) -> list[ScheduledActivity]:
        scheduled: list[ScheduledActivity] = []
        parents = descendants(root, "Activities_Timetable", "Activity_Timetable", "Timetable")
        for parent in parents:
            nodes: Iterable[ET.Element] = children(parent, "Activity") if local_name(parent.tag) != "Activity" else [parent]
            for node in nodes:
                activity_id = text(child(node, "Id", "Activity_Id"))
                day = text(child(node, "Day", "Preferred_Day"))
                hour = text(child(node, "Hour", "Preferred_Hour"))
                if activity_id and day and hour:
                    scheduled.append(ScheduledActivity(activity_id, day, hour, text(child(node, "Room"))))
        return scheduled

    def _parse_locked_time_constraints(self, constraints: list[Constraint]) -> list[ScheduledActivity]:
        scheduled: list[ScheduledActivity] = []
        for constraint in constraints:
            normalized_type = constraint.type.lower().replace("_", "")
            if "activitypreferredstartingtime" not in normalized_type:
                continue
            if not self._constraint_is_effective_solution_hint(constraint):
                continue
            activity_id = str(constraint.fields.get("Activity_Id") or constraint.fields.get("ActivityId") or "")
            day = str(constraint.fields.get("Day") or constraint.fields.get("Preferred_Day") or constraint.fields.get("PreferredDay") or "")
            hour = str(constraint.fields.get("Hour") or constraint.fields.get("Preferred_Hour") or constraint.fields.get("PreferredHour") or "")
            if activity_id and day and hour:
                scheduled.append(ScheduledActivity(activity_id=activity_id, day=day, hour=hour))
        return scheduled

    def _apply_room_constraints(self, scheduled: list[ScheduledActivity], constraints: list[Constraint]) -> None:
        rooms_by_activity: dict[str, str] = {}
        for constraint in constraints:
            normalized_type = constraint.type.lower().replace("_", "")
            if "activitypreferredroom" not in normalized_type:
                continue
            if not self._constraint_is_effective_solution_hint(constraint):
                continue
            activity_id = str(constraint.fields.get("Activity_Id") or constraint.fields.get("ActivityId") or "")
            room = str(constraint.fields.get("Room") or constraint.fields.get("Preferred_Room") or "")
            if activity_id and room:
                rooms_by_activity[activity_id] = room
        for placement in scheduled:
            if not placement.room:
                placement.room = rooms_by_activity.get(placement.activity_id, "")

    def _constraint_is_effective_solution_hint(self, constraint: Constraint) -> bool:
        active = _bool(str(constraint.fields.get("Active", "true")))
        weight = str(constraint.weight or "100").replace("%", "")
        locked = _bool(str(constraint.fields.get("Permanently_Locked", "true")))
        return active and (weight in {"", "100", "100.0"} or locked)

    def _attach_teacher_availability(self, data: FetData) -> None:
        by_name = {teacher.name: teacher for teacher in data.teachers}
        for constraint in data.constraints:
            if "Teacher" in constraint.fields and "notavailable" in constraint.type.lower():
                teacher = by_name.get(str(constraint.fields["Teacher"]))
                if teacher:
                    teacher.availability.append(constraint.fields)

    def _list_items(self, root: ET.Element, list_name: str, item_name: str) -> list[ET.Element]:
        parent = descendant(root, list_name)
        return children(parent, item_name) if parent is not None else []

    def _flatten_children(self, node: ET.Element) -> dict[str, object]:
        fields: dict[str, object] = {}
        for item in list(node):
            value: object = self._flatten_children(item) if list(item) else text(item)
            key = local_name(item.tag)
            if key in fields:
                existing = fields[key]
                if not isinstance(existing, list):
                    fields[key] = [existing]
                fields[key].append(value)  # type: ignore[union-attr]
            else:
                fields[key] = value
        return fields
