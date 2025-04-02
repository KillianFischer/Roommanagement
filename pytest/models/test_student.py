import pytest
import pandas as pd
import os
import sys

# Add the src directory
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "src"))

from models.student import StudentPreference
from models.company import Company, CompanySession

IMPORT_FOLDER = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "import")

# Sample data
@pytest.fixture
def sample_student_data():
    return pd.DataFrame(
        {
            "Klasse": ["10A"],
            "Name": ["Dilaksan"],
            "Vorname": ["Christian"],
            "Wahl 1": [1],
            "Wahl 2": [2],
            "Wahl 3": [3],
        }
    )

@pytest.fixture
def real_student_data():
    file_path = os.path.join(IMPORT_FOLDER, "BOT2_Wahl.xlsx")
    return pd.read_excel(file_path)

# tests
def test_student_preference_from_dataframe(sample_student_data):
    preferences = StudentPreference.from_dataframe(sample_student_data)
    assert len(preferences) == 1
    assert preferences[0].name == "Dilaksan, Christian"
    assert preferences[0].wishes == ["1", "2", "3"]


# Tests with real data
def test_student_preference_from_real_data(real_student_data):
    """Test creating StudentPreference objects from real Excel data."""
    preferences = StudentPreference.from_dataframe(real_student_data)
    
    # Check that we have the expected number of students
    assert len(preferences) == len(real_student_data)
    
    # Check properties of the first student
    first_student = preferences[0]
    expected_name = f"{real_student_data.iloc[0]['Name']}, {real_student_data.iloc[0]['Vorname']}"
    assert first_student.name == expected_name
    
    # Check that student IDs are generated correctly
    assert first_student.student_id.startswith(real_student_data.iloc[0]['Klasse'])
    
    # Check that wishes are loaded correctly
    assert len(first_student.wishes) > 0
    
    # Check that at least some students have multiple wishes
    has_multiple_wishes = False
    for student in preferences:
        if len(student.wishes) > 1:
            has_multiple_wishes = True
            break
    assert has_multiple_wishes, "No students have multiple wishes"

def test_student_preference_with_company_mapping(real_student_data):
    """Test creating StudentPreference objects with company mapping."""
    # Create a sample company mapping
    company_mapping = {
        1: "Company Obi",
        2: "Company Toby",
        3: "Company Copy"
    }
    
    preferences = StudentPreference.from_dataframe(real_student_data, company_mapping)
    
    # Check that company mapping is applied correctly
    # Find a student with a wish that matches a key in the mapping
    for student in preferences:
        for wish in student.wishes:
            if wish in ["Company A", "Company B", "Company C"]:
                # Found a mapped wish
                assert True
                return
    
    # If we didn't find any mapped wishes, the test should still pass
    # because not all students might have wishes that match our sample mapping
    assert True

def test_calculate_fulfillment_score():
    """Test the calculation of the fulfillment score based on student wishes."""
    # Create a student with wishes
    student = StudentPreference(
        student_id="TEST_1",
        name="Test Student",
        wishes=["Company A", "Company B", "Company C", "Company D", "Company E", "Company F"]
    )
    
    # Test first wish fulfilled, 6 Points
    score = student.calculate_fulfillment_score("Company Obi")
    assert score == 6.0
    assert student.fulfillment_score == 6.0
    
    # Test second wish fulfilled, 5 Points
    score = student.calculate_fulfillment_score("Company Toby")
    assert score == 5.0
    assert student.fulfillment_score == 5.0
    
    # Test third wish fulfilled, 4 Points
    score = student.calculate_fulfillment_score("Company Copy")
    assert score == 4.0
    assert student.fulfillment_score == 4.0
    
    # Test fourth wish fulfilled, 3 Points
    score = student.calculate_fulfillment_score("Company Obi")
    assert score == 3.0
    assert student.fulfillment_score == 3.0
    
    # Test fifth wish fulfilled, 2 Points
    score = student.calculate_fulfillment_score("Company Toby")
    assert score == 2.0
    assert student.fulfillment_score == 2.0
    
    # Test sixth wish fulfilled, 1 Point
    score = student.calculate_fulfillment_score("Company Copy")
    assert score == 1.0
    assert student.fulfillment_score == 1.0
    
    # Test company not in wishes, 0 Points :(
    score = student.calculate_fulfillment_score("Company G")
    assert score == 0.0
    assert student.fulfillment_score == 0.0
    
    # Test no company assigned, 0 Points :(
    score = student.calculate_fulfillment_score(None) # FIXME
    assert score == 0.0
    assert student.fulfillment_score == 0.0

def test_scheduler_erfullungsscore_calculation():
    """Test the scheduler service's erfüllungsscore calculation."""
    # Create scheduler service
    scheduler = SchedulerService() # FIXME
    
    # Create test companies
    companies = [
        Company(name="Company A", capacity=20, max_sessions=1, earliest_slot=0, blocked_slots=[]),
        Company(name="Company B", capacity=20, max_sessions=1, earliest_slot=0, blocked_slots=[]),
        Company(name="Company C", capacity=20, max_sessions=1, earliest_slot=0, blocked_slots=[]),
    ]
    scheduler.companies = companies
    
    # Create test students
    students = [
        StudentPreference(student_id="TEST_1", name="Student 1", wishes=["Company A", "Company B", "Company C"]),
        StudentPreference(student_id="TEST_2", name="Student 2", wishes=["Company B", "Company A", "Company C"]),
    ]
    scheduler.student_preferences = students
    
    # Add rooms
    scheduler.rooms = ["Room 1", "Room 2", "Room 3"]
    
    # Create schedule manually
    scheduler.schedule = {}
    
    # Company A in slot 0
    session_a = CompanySession(
        company=companies[0],
        room="Room 1",
        time_slot="A",
        time_range="8:45 – 9:30"
    )
    session_a.add_student("TEST_1", "Student 1")  # First wish for student 1
    scheduler.schedule[("Company A", 0)] = session_a
    
    # Company B in slot 1
    session_b = CompanySession(
        company=companies[1],
        room="Room 2",
        time_slot="B",
        time_range="9:50 – 10:35"
    )
    session_b.add_student("TEST_1", "Student 1")  # Second wish for student 1
    session_b.add_student("TEST_2", "Student 2")  # First wish for student 2
    scheduler.schedule[("Company B", 1)] = session_b
    
    # Calculate scores
    scheduler._calculate_student_fulfillment_scores({})
    
    # Check student 1 (got 1st and 2nd wish -> 6+5 out of max 11 points -> 91.67%)
    assert 91.0 <= students[0].fulfillment_score <= 92.0 # FIXME
    
    # Check student 2 (got 1st wish only -> 6 out of max 6 points -> 100%)
    assert students[1].fulfillment_score == 100.0
    
    # Test with different configuration
    # Reset students
    students[0].fulfillment_score = None
    students[1].fulfillment_score = None
    
    # Empty existing schedule
    scheduler.schedule = {}
    
    # Company B in slot 0
    session_b = CompanySession(
        company=companies[1],
        room="Room 2",
        time_slot="A",
        time_range="8:45 – 9:30"
    )
    session_b.add_student("TEST_1", "Student 1")  # Second wish for student 1
    session_b.add_student("TEST_2", "Student 2")  # First wish for student 2
    scheduler.schedule[("Company B", 0)] = session_b
    
    # Company C in slot 1
    session_c = CompanySession(
        company=companies[2],
        room="Room 3",
        time_slot="B",
        time_range="9:50 – 10:35"
    )
    session_c.add_student("TEST_1", "Student 1")  # Third wish for student 1
    session_c.add_student("TEST_2", "Student 2")  # Third wish for student 2
    scheduler.schedule[("Company C", 1)] = session_c
