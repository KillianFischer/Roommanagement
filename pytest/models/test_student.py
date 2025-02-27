import pytest
import pandas as pd
from models.student import StudentPreference
# not working yet
@pytest.fixture
def sample_student_data():
    return pd.DataFrame({
        'Klasse': ['10A'],
        'Name': ['Dilaksan'],
        'Vorname': ['Müller'],
        'Wahl 1': [1],
        'Wahl 2': [2],
        'Wahl 3': [3]
    })

def test_student_preference_from_dataframe():
    df = pd.DataFrame({
        'Klasse': ['10A'],
        'Name': ['Gwen'],
        'Vorname': ['Gweb'],
        'Wahl 1': [1],
        'Wahl 2': [2],
        'Wahl 3': [3]
    })
    
    preferences = StudentPreference.from_dataframe(df)
    assert len(preferences) == 1
    assert preferences[0].name == "Christian, Müller"
    assert preferences[0].wishes == ['1', '2', '3']

def test_get_satisfaction_score():
    student = StudentPreference(
        student_id="10A_1",
        name="Christian, Müller",
        wishes=['1', '2', '3']
    )
    
    # All
    assert student.get_satisfaction_score([True, True, True], max_wishes=3) == 100.0
    
    # First
    assert student.get_satisfaction_score([True, False, False], max_wishes=3) > 0
    
    # Noooone
    assert student.get_satisfaction_score([False, False, False], max_wishes=3) == 0.0
    
    # Default
    assert student.get_satisfaction_score([True, True, True]) == 71.42857142857143
