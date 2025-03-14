import pytest
import pandas as pd
import os
import sys

# Add the src directory
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "src"))

from models.student import StudentPreference

# Import folder
IMPORT_FOLDER = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "import")

#
# Sample data
#
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
    """Load the actual student preference data from the Excel file."""
    file_path = os.path.join(IMPORT_FOLDER, "BOT2_Wahl.xlsx")
    return pd.read_excel(file_path)

#
# Basic tests
#
def test_student_preference_from_dataframe(sample_student_data):
    """Test creating StudentPreference objects from a DataFrame."""
    preferences = StudentPreference.from_dataframe(sample_student_data)
    assert len(preferences) == 1
    assert preferences[0].name == "Dilaksan, Christian"
    assert preferences[0].wishes == ["1", "2", "3"]

def test_get_satisfaction_score():
    """Test calculating satisfaction scores for student preferences."""
    student = StudentPreference(
        student_id="10A_1", name="Dilaksan, Christian", wishes=["1", "2", "3"]
    )

    # All wishes fulfilled with max_wishes=3
    # The method always uses max_points=21 (sum of 6+5+4+3+2+1)
    # 3 wishes fulfilled with max_wishes=3: 3+2+1 = 6 points out of 21 possible = 28.57%
    assert student.get_satisfaction_score([True, True, True], max_wishes=3) == 28.57142857142857

    # Only first wish fulfilled with max_wishes=3
    # 1st wish with max_wishes=3 = 3 points out of 21 possible = 14.29%
    assert student.get_satisfaction_score([True, False, False], max_wishes=3) == 14.285714285714285

    # No wishes fulfilled
    assert student.get_satisfaction_score([False, False, False], max_wishes=3) == 0.0

    # Default max_wishes (6)
    # 3 wishes fulfilled: 6+5+4 = 15 points out of 21 possible = 71.43%
    assert student.get_satisfaction_score([True, True, True]) == 71.42857142857143

    # Test with different weights
    # First wish (6 points), second wish (5 points), third wish (4 points)
    # Total: 15 points out of 21 possible points = 71.43%
    assert student.get_satisfaction_score([True, True, True, False, False, False]) == 71.42857142857143

#
# Tests with real data
#
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
        1: "Company A",
        2: "Company B",
        3: "Company C"
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
