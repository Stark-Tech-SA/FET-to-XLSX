"""Domain models for FET timetabling data."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(slots=True)
class Institution:
    """General metadata extracted from a FET file."""

    name: str = ""
    version: str = ""
    comments: str = ""


@dataclass(slots=True)
class Teacher:
    """Teacher declared in the FET file."""

    name: str
    comments: str = ""
    availability: list[dict[str, Any]] = field(default_factory=list)


@dataclass(slots=True)
class Subject:
    """Subject declared in the FET file."""

    name: str
    comments: str = ""


@dataclass(slots=True)
class StudentsGroup:
    """Students set. FET can nest years, groups and subgroups."""

    year: str = ""
    group: str = ""
    subgroup: str = ""
    number_of_students: int | None = None


@dataclass(slots=True)
class Room:
    """Room and optional building metadata."""

    name: str
    capacity: int | None = None
    building: str = ""
    comments: str = ""


@dataclass(slots=True)
class Activity:
    """Teaching activity from FET."""

    id: str
    teachers: list[str] = field(default_factory=list)
    subject: str = ""
    students: list[str] = field(default_factory=list)
    duration: int | None = None
    total_duration: int | None = None
    student_count: int | None = None
    active: bool = True
    comments: str = ""


@dataclass(slots=True)
class Constraint:
    """Generic FET time or space constraint."""

    category: str
    type: str
    weight: str = ""
    fields: dict[str, Any] = field(default_factory=dict)


@dataclass(slots=True)
class ScheduledActivity:
    """Optional solved timetable placement, if present in the XML."""

    activity_id: str
    day: str
    hour: str
    room: str = ""


@dataclass(slots=True)
class FetData:
    """Aggregate root for parsed FET content."""

    institution: Institution = field(default_factory=Institution)
    teachers: list[Teacher] = field(default_factory=list)
    subjects: list[Subject] = field(default_factory=list)
    groups: list[StudentsGroup] = field(default_factory=list)
    rooms: list[Room] = field(default_factory=list)
    activities: list[Activity] = field(default_factory=list)
    constraints: list[Constraint] = field(default_factory=list)
    days: list[str] = field(default_factory=list)
    hours: list[str] = field(default_factory=list)
    scheduled_activities: list[ScheduledActivity] = field(default_factory=list)

    @property
    def student_sets_count(self) -> int:
        return len(self.groups)

    @property
    def restrictions_count(self) -> int:
        return len(self.constraints)
