import pytest
import pandas as pd
import os
import sys

# Add the src directory
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "src"))

# Import folder
IMPORT_FOLDER = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "import")

# Sample data
@pytest.fixture
def sample_student_data():
    return pd.DataFrame(
        {
            "Klasse": ["10A", "10A"],
            "Name": ["Dilaksan", "Müller"],
            "Vorname": ["Christian", "Gwen"],
            "Wahl 1": [1, 2],
            "Wahl 2": [2, 1],
            "Wahl 3": [3, 3],
        }
    )

@pytest.fixture
def sample_company_data():
    return pd.DataFrame(
        {
            "Unternehmen": ["Company A", "Company B", "Company C"],
            "Fachrichtung": ["IT", "Engineering", "Marketing"],
            "Max. Teilnehmer": [5, 4, 3],
            "Min. Teilnehmer": [2, 2, 1],
            "Frühester Zeitpunkt": ["A", "B", "A"],
        }
    )

@pytest.fixture
def sample_room_data():
    return pd.DataFrame({0: [101, 102, 103, "Aula", 104]})

# Real tests
@pytest.fixture
def real_room_data():
    """Load the actual room data from the Excel file."""
    file_path = os.path.join(IMPORT_FOLDER, "BOT0_Raumliste.xlsx")
    return pd.read_excel(file_path)

@pytest.fixture
def real_company_data():
    """Load the actual company data from the Excel file."""
    file_path = os.path.join(IMPORT_FOLDER, "BOT1_Veranstaltungsliste.xlsx")
    return pd.read_excel(file_path)

@pytest.fixture
def real_student_data():
    """Load the actual student preference data from the Excel file."""
    file_path = os.path.join(IMPORT_FOLDER, "BOT2_Wahl.xlsx")
    return pd.read_excel(file_path)

# Basic tests with sample data
def test_load_student_preferences(scheduler, sample_student_data):
    result = scheduler.load_student_preferences(sample_student_data)
    assert result == True
    assert len(scheduler.student_preferences) == 2
    assert scheduler.student_preferences[0].name == "Dilaksan, Christian"

def test_load_companies(scheduler, sample_company_data):
    result = scheduler.load_companies(sample_company_data)
    assert result == True
    assert len(scheduler.companies) == 3
    assert scheduler.companies[0].name == "Company A"
    assert scheduler.companies[0].capacity == 5

def test_load_rooms(scheduler, sample_room_data):
    result = scheduler.load_rooms(sample_room_data)
    assert result == True
    assert len(scheduler.rooms) == 5  # includes Aula
    assert "101" in scheduler.rooms

def test_is_data_loaded(
    scheduler, sample_student_data, sample_company_data, sample_room_data
):
    assert scheduler.is_data_loaded() == False

    scheduler.load_student_preferences(sample_student_data)
    assert scheduler.is_data_loaded() == False

    scheduler.load_companies(sample_company_data)
    assert scheduler.is_data_loaded() == False

    scheduler.load_rooms(sample_room_data)
    assert scheduler.is_data_loaded() == True

def test_generate_schedule(
    scheduler, sample_student_data, sample_company_data, sample_room_data
):
    scheduler.load_student_preferences(sample_student_data)
    scheduler.load_companies(sample_company_data)
    scheduler.load_rooms(sample_room_data)

    result = scheduler.generate_schedule()
    assert result == True
    assert len(scheduler.schedule) > 0

#
# Tests with real data from Excel files
#
def test_load_real_room_data(scheduler, real_room_data):
    """Test loading the actual room data."""
    result = scheduler.load_rooms(real_room_data)
    assert result == True
    assert scheduler.rooms is not None
    assert len(scheduler.rooms) > 0
    # Check that at least some rooms are loaded
    assert any(room.isdigit() for room in scheduler.rooms)

def test_load_real_company_data(scheduler, real_company_data):
    """Test loading the actual company data."""
    result = scheduler.load_companies(real_company_data)
    assert result == True
    assert scheduler.companies is not None
    assert len(scheduler.companies) > 0
    assert len(scheduler.companies) == len(real_company_data)
    first_company = scheduler.companies[0]
    assert first_company.name == real_company_data.iloc[0]['Unternehmen'].strip()
    assert first_company.capacity == real_company_data.iloc[0]['Max. Teilnehmer']
    
    assert isinstance(first_company.min_participants, int)
    assert first_company.min_participants >= 0

def test_load_real_student_preferences(scheduler, real_student_data):
    """Test loading the actual student preference data."""
    result = scheduler.load_student_preferences(real_student_data)
    assert result == True
    assert scheduler.student_preferences is not None
    assert len(scheduler.student_preferences) > 0
    assert len(scheduler.student_preferences) == len(real_student_data)
    first_student = scheduler.student_preferences[0]
    expected_name = f"{real_student_data.iloc[0]['Name']}, {real_student_data.iloc[0]['Vorname']}"
    assert first_student.name == expected_name
    assert len(first_student.wishes) > 0

def test_is_real_data_loaded(scheduler, real_student_data, real_company_data, real_room_data):
    """Test that all real data is properly loaded."""
    assert scheduler.is_data_loaded() == False

    scheduler.load_student_preferences(real_student_data)
    assert scheduler.is_data_loaded() == False

    scheduler.load_companies(real_company_data)
    assert scheduler.is_data_loaded() == False

    scheduler.load_rooms(real_room_data)
    assert scheduler.is_data_loaded() == True

def test_generate_schedule_with_real_data(scheduler, real_student_data, real_company_data, real_room_data):
    """Test generating a schedule with the actual data."""
    scheduler.load_student_preferences(real_student_data)
    scheduler.load_companies(real_company_data)
    scheduler.load_rooms(real_room_data)

    result = scheduler.generate_schedule()
    assert result == True
    assert scheduler.schedule is not None
    assert len(scheduler.schedule) > 0

    scheduled_companies = set(company for company, _ in scheduler.schedule.keys())
    assert len(scheduled_companies) > 0
    
    for (_, _), session in scheduler.schedule.items():
        assert session.room in scheduler.rooms

def test_student_assignments_with_real_data(scheduler, real_student_data, real_company_data, real_room_data):
    """Test that students are assigned to companies based on their preferences."""
    scheduler.load_student_preferences(real_student_data)
    scheduler.load_companies(real_company_data)
    scheduler.load_rooms(real_room_data)

    result = scheduler.generate_schedule()
    assert result == True
    
    assigned_students = set()
    for (_, _), session in scheduler.schedule.items():
        for student in session.students:
            assigned_students.add(student['name'])
    
    assert len(assigned_students) > 0
    
    assert len(assigned_students) >= len(scheduler.student_preferences) * 0.1

#
# Advanced tests with real data
#
def test_student_preference_satisfaction(scheduler, real_student_data, real_company_data, real_room_data):
    """Test that student preferences are satisfied to a reasonable degree."""
    scheduler.load_student_preferences(real_student_data)
    scheduler.load_companies(real_company_data)
    scheduler.load_rooms(real_room_data)

    result = scheduler.generate_schedule()
    assert result == True
    
    company_to_id = {}
    for idx, row in real_company_data.iterrows():
        company_to_id[row['Unternehmen'].strip()] = idx + 1  # Use 1-based index as company ID
    
    id_to_company = {}
    for company_name, company_id in company_to_id.items():
        id_to_company[str(company_id)] = company_name
    
    student_satisfaction = []
    for student in scheduler.student_preferences:
        assigned_companies = set()
        for (company, _), session in scheduler.schedule.items():
            for assigned_student in session.students:
                if assigned_student['name'] == student.name:
                    assigned_companies.add(company)
        
        wishes_fulfilled = []
        for wish in student.wishes:
            # The wish might be a company name or a company ID (as string)
            wish_fulfilled = False
            if wish in assigned_companies:
                wish_fulfilled = True
            elif wish in id_to_company and id_to_company[wish] in assigned_companies:
                wish_fulfilled = True
            wishes_fulfilled.append(wish_fulfilled)
        
        if wishes_fulfilled:
            satisfaction = student.get_satisfaction_score(wishes_fulfilled)
            student_satisfaction.append(satisfaction)
    
    assert len(student_satisfaction) > 0
    
    if student_satisfaction:
        avg_satisfaction = sum(student_satisfaction) / len(student_satisfaction)
        assert avg_satisfaction > 0.0

def test_company_capacity_constraints(scheduler, real_student_data, real_company_data, real_room_data):
    """Test that company capacity constraints are respected."""
    scheduler.load_student_preferences(real_student_data)
    scheduler.load_companies(real_company_data)
    scheduler.load_rooms(real_room_data)

    result = scheduler.generate_schedule()
    assert result == True
    
    for (company_name, _), session in scheduler.schedule.items():
        company = None
        for c in scheduler.companies:
            if c.name == company_name:
                company = c
                break
        
        assert company is not None
        assert len(session.students) <= company.capacity

def test_time_slot_constraints(scheduler, real_student_data, real_company_data, real_room_data):
    """Test that time slot constraints are respected."""
    scheduler.load_student_preferences(real_student_data)
    scheduler.load_companies(real_company_data)
    scheduler.load_rooms(real_room_data)

    result = scheduler.generate_schedule()
    assert result == True
    
    for (company_name, slot), _ in scheduler.schedule.items():
        company = None
        for c in scheduler.companies:
            if c.name == company_name:
                company = c
                break
        
        assert company is not None
        assert slot >= company.earliest_slot

def test_student_assignment_uniqueness(scheduler, real_student_data, real_company_data, real_room_data):
    """Test that students are not assigned to the same company multiple times in the same time slot."""
    scheduler.load_student_preferences(real_student_data)
    scheduler.load_companies(real_company_data)
    scheduler.load_rooms(real_room_data)

    result = scheduler.generate_schedule()
    assert result == True
    
    student_assignments = {}
    
    for (company_name, slot), session in scheduler.schedule.items():
        for student in session.students:
            student_name = student['name']
            if student_name not in student_assignments:
                student_assignments[student_name] = set()
            
            company_slot_pair = (company_name, slot)
            assert company_slot_pair not in student_assignments[student_name], \
                f"Student {student_name} is assigned to {company_name} in slot {slot} multiple times"
            
            student_assignments[student_name].add(company_slot_pair)

def test_schedule_with_subset_of_data(scheduler, real_student_data, real_company_data, real_room_data):
    """Test generating a schedule with a subset of the data."""
    student_sample = real_student_data.sample(min(50, len(real_student_data)))
    
    company_sample = real_company_data.sample(min(10, len(real_company_data)))
    
    room_sample = real_room_data.sample(min(5, len(real_room_data)))
    
    scheduler.load_student_preferences(student_sample)
    scheduler.load_companies(company_sample)
    scheduler.load_rooms(room_sample)

    result = scheduler.generate_schedule()
    assert result == True
    assert scheduler.schedule is not None
    
    company_names = set(company.name for company in scheduler.companies)
    for company_name, _ in scheduler.schedule.keys():
        assert company_name in company_names
