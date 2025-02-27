import pytest
import pandas as pd
from services.scheduler import SchedulerService
# not working yet
@pytest.fixture
def scheduler():
    return SchedulerService()

@pytest.fixture
def sample_student_data():
    return pd.DataFrame({
        'Klasse': ['10A', '10A'],
        'Name': ['Dilaksan', 'Müller'],
        'Vorname': ['Christian', 'Gwen'],
        'Wahl 1': [1, 2],
        'Wahl 2': [2, 1],
        'Wahl 3': [3, 3]
    })

@pytest.fixture
def sample_company_data():
    return pd.DataFrame({
        'Unternehmen': ['Company A', 'Company B', 'Company C'],
        'Fachrichtung': ['IT', 'Engineering', 'Marketing'],
        'Max. Teilnehmer': [5, 4, 3],
        'Max. Veranstaltungen': [2, 2, 1],
        'Frühester Zeitpunkt': ['A', 'B', 'A']
    })

@pytest.fixture
def sample_room_data():
    return pd.DataFrame({
        0: [101, 102, 103, 'Aula', 104]
    })

def test_load_student_preferences(scheduler, sample_student_data):
    result = scheduler.load_student_preferences(sample_student_data)
    assert result == True
    assert len(scheduler.student_preferences) == 2
    assert scheduler.student_preferences[0].name == "Christian, Müller"

def test_load_companies(scheduler, sample_company_data):
    result = scheduler.load_companies(sample_company_data)
    assert result == True
    assert len(scheduler.companies) == 3
    assert scheduler.companies[0].name == "Company A"
    assert scheduler.companies[0].capacity == 5

def test_load_rooms(scheduler, sample_room_data):
    result = scheduler.load_rooms(sample_room_data)
    assert result == True
    assert len(scheduler.rooms) == 4  # not Aula
    assert '101' in scheduler.rooms

def test_is_data_loaded(scheduler, sample_student_data, sample_company_data, sample_room_data):
    assert scheduler.is_data_loaded() == False
    
    scheduler.load_student_preferences(sample_student_data)
    assert scheduler.is_data_loaded() == False
    
    scheduler.load_companies(sample_company_data)
    assert scheduler.is_data_loaded() == False
    
    scheduler.load_rooms(sample_room_data)
    assert scheduler.is_data_loaded() == True

def test_generate_schedule(scheduler, sample_student_data, sample_company_data, sample_room_data):
    scheduler.load_student_preferences(sample_student_data)
    scheduler.load_companies(sample_company_data)
    scheduler.load_rooms(sample_room_data)
    
    result = scheduler.generate_schedule()
    assert result == True
    assert len(scheduler.schedule) > 0