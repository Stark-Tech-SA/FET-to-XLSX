from fet_to_xlsx.parser.models import Activity, FetData
from fet_to_xlsx.parser.solution_finder import SolutionFinder


def test_solution_finder_loads_companion_activities_xml(tmp_path):
    fet_path = tmp_path / "school.fet"
    fet_path.write_text("<fet />", encoding="utf-8")
    result_dir = tmp_path / "timetables" / "school"
    result_dir.mkdir(parents=True)
    (result_dir / "school_activities.xml").write_text(
        """
        <Activities_Timetable>
          <Activity><Id>10</Id><Day>Lunes</Day><Hour>08:00</Hour><Room>101</Room></Activity>
        </Activities_Timetable>
        """,
        encoding="utf-8",
    )
    data = FetData(activities=[Activity(id="10")])

    placements = SolutionFinder().find_for(fet_path, data)

    assert len(placements) == 1
    assert placements[0].activity_id == "10"
    assert placements[0].day == "Lunes"
    assert placements[0].hour == "08:00"
    assert placements[0].room == "101"
