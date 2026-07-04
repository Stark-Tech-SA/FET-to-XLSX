from xml.etree import ElementTree as ET

from fet_to_xlsx.parser.parser import FetParser


def test_parser_extracts_core_fet_entities():
    xml = """
    <fet version="6.0">
      <Institution_Name>Colegio Demo</Institution_Name>
      <Teachers_List><Teacher><Name>Ana</Name><Comments>Tutora</Comments></Teacher></Teachers_List>
      <Subjects_List><Subject><Name>Matemáticas</Name></Subject></Subjects_List>
      <Students_List><Year><Name>1</Name><Number_of_Students>30</Number_of_Students><Group><Name>1A</Name><Subgroup><Name>1A-1</Name></Subgroup></Group></Year></Students_List>
      <Rooms_List><Room><Name>101</Name><Capacity>35</Capacity><Building>A</Building></Room></Rooms_List>
      <Activities_List><Activity><Id>1</Id><Teacher>Ana</Teacher><Subject>Matemáticas</Subject><Students>1A</Students><Duration>2</Duration><Number_Of_Students>30</Number_Of_Students><Active>true</Active></Activity></Activities_List>
      <Time_Constraints_List><ConstraintTeacherNotAvailableTimes><Weight_Percentage>100</Weight_Percentage><Teacher>Ana</Teacher><Day>Lunes</Day></ConstraintTeacherNotAvailableTimes></Time_Constraints_List>
    </fet>
    """
    data = FetParser().parse(ET.fromstring(xml))
    assert data.institution.name == "Colegio Demo"
    assert len(data.teachers) == 1
    assert len(data.subjects) == 1
    assert len(data.groups) == 3
    assert len(data.rooms) == 1
    assert data.activities[0].teachers == ["Ana"]
    assert data.constraints[0].type == "ConstraintTeacherNotAvailableTimes"
    assert data.teachers[0].availability
