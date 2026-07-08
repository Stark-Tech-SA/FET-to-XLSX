"""Find FET-generated timetable placements stored next to a .fet file."""
from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET

from fet_to_xlsx.parser.models import FetData, ScheduledActivity
from fet_to_xlsx.parser.xml_utils import child, children, descendants, text


class SolutionFinder:
    """Locate companion XML timetable files produced by FET.

    FET input files (`.fet`) do not always contain the generated timetable. In
    typical workflows, FET writes result files beside the input file, often under
    a `timetables` directory and with names containing `activities`. This helper
    searches those companion XML files and extracts activity placements.
    """

    MAX_FILES = 200

    def find_for(self, fet_path: str | Path, data: FetData) -> list[ScheduledActivity]:
        """Return placements from the best companion solution file, if any."""
        path = Path(fet_path)
        known_ids = {activity.id for activity in data.activities}
        best: list[ScheduledActivity] = []
        for candidate in self._candidate_files(path):
            placements = self._read_candidate(candidate, known_ids)
            if len(placements) > len(best):
                best = placements
        return best

    def _candidate_files(self, fet_path: Path) -> list[Path]:
        base = fet_path.stem.lower()
        roots = [fet_path.parent, fet_path.parent / "timetables"]
        candidates: list[Path] = []
        seen: set[Path] = set()
        for root in roots:
            if not root.exists():
                continue
            for file_path in root.rglob("*.xml"):
                if file_path in seen or file_path.name.startswith("~"):
                    continue
                seen.add(file_path)
                name = file_path.name.lower()
                if base in name or "activities" in name or "timetable" in name or "horario" in name:
                    candidates.append(file_path)
                if len(candidates) >= self.MAX_FILES:
                    return candidates
        return candidates

    def _read_candidate(self, file_path: Path, known_ids: set[str]) -> list[ScheduledActivity]:
        try:
            root = ET.parse(file_path).getroot()
        except ET.ParseError:
            return []
        placements: list[ScheduledActivity] = []
        for node in descendants(root, "Activity"):
            placement = self._placement_from_activity_node(node)
            if placement and (not known_ids or placement.activity_id in known_ids):
                placements.append(placement)
        return placements

    def _placement_from_activity_node(self, node: ET.Element) -> ScheduledActivity | None:
        activity_id = text(child(node, "Id", "Activity_Id", "ActivityId"))
        day = text(child(node, "Day", "Preferred_Day", "Real_Day"))
        hour = text(child(node, "Hour", "Preferred_Hour", "Real_Hour"))
        room = text(child(node, "Room", "Real_Room", "Preferred_Room"))
        if not activity_id:
            activity_id = text(child(node, "Activity"))
        # Some FET result XML variants wrap details one level down.
        if not (activity_id and day and hour):
            for nested in children(node, "Starting_Time", "Time", "Placement"):
                activity_id = activity_id or text(child(nested, "Id", "Activity_Id", "ActivityId"))
                day = day or text(child(nested, "Day", "Preferred_Day", "Real_Day"))
                hour = hour or text(child(nested, "Hour", "Preferred_Hour", "Real_Hour"))
                room = room or text(child(nested, "Room", "Real_Room", "Preferred_Room"))
        if activity_id and day and hour:
            return ScheduledActivity(activity_id=activity_id, day=day, hour=hour, room=room)
        return None
