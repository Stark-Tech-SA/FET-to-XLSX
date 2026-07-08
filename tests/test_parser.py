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


def test_parser_ignores_namespaces_and_builds_schedule_from_locked_constraints():
    xml = """
    <ns:fet xmlns:ns="urn:fet" version="6.0">
      <ns:Institution_Name>Colegio Demo</ns:Institution_Name>
      <ns:Days_List><ns:Day><ns:Name>Lunes</ns:Name></ns:Day></ns:Days_List>
      <ns:Hours_List><ns:Hour><ns:Name>08:00</ns:Name></ns:Hour></ns:Hours_List>
      <ns:Teachers_List><ns:Teacher><ns:Name>Ana</ns:Name></ns:Teacher></ns:Teachers_List>
      <ns:Subjects_List><ns:Subject><ns:Name>Matemáticas</ns:Name></ns:Subject></ns:Subjects_List>
      <ns:Students_List><ns:Year><ns:Name>1</ns:Name><ns:Group><ns:Name>1A</ns:Name></ns:Group></ns:Year></ns:Students_List>
      <ns:Rooms_List><ns:Room><ns:Name>101</ns:Name></ns:Room></ns:Rooms_List>
      <ns:Activities_List>
        <ns:Activity><ns:Id>1</ns:Id><ns:Teacher>Ana</ns:Teacher><ns:Subject>Matemáticas</ns:Subject><ns:Students>1A</ns:Students><ns:Duration>1</ns:Duration></ns:Activity>
      </ns:Activities_List>
      <ns:Time_Constraints_List>
        <ns:ConstraintActivityPreferredStartingTime>
          <ns:Weight_Percentage>100</ns:Weight_Percentage><ns:Activity_Id>1</ns:Activity_Id><ns:Preferred_Day>Lunes</ns:Preferred_Day><ns:Preferred_Hour>08:00</ns:Preferred_Hour><ns:Permanently_Locked>true</ns:Permanently_Locked><ns:Active>true</ns:Active>
        </ns:ConstraintActivityPreferredStartingTime>
      </ns:Time_Constraints_List>
      <ns:Space_Constraints_List>
        <ns:ConstraintActivityPreferredRoom>
          <ns:Weight_Percentage>100</ns:Weight_Percentage><ns:Activity_Id>1</ns:Activity_Id><ns:Room>101</ns:Room><ns:Permanently_Locked>true</ns:Permanently_Locked>
        </ns:ConstraintActivityPreferredRoom>
      </ns:Space_Constraints_List>
    </ns:fet>
    """
    data = FetParser().parse(ET.fromstring(xml))
    assert data.institution.name == "Colegio Demo"
    assert data.days == ["Lunes"]
    assert data.hours == ["08:00"]
    assert data.scheduled_activities[0].activity_id == "1"
    assert data.scheduled_activities[0].day == "Lunes"
    assert data.scheduled_activities[0].hour == "08:00"
    assert data.scheduled_activities[0].room == "101"
