import pytest
import pandas as pd
import os
import sys

# Add the src directory 
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "src"))

from models.company import Company, CompanySession

# Import folder
IMPORT_FOLDER = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "import")

#
# Sample data
#
@pytest.fixture
def sample_company_data():
    return pd.DataFrame(
        {
            "Unternehmen": ["Company A"],
            "Fachrichtung": ["IT"],
            "Max. Teilnehmer": [5],
            "Min. Teilnehmer": [2],
            "Frühester Zeitpunkt": ["A"],
        }
    )

@pytest.fixture
def real_company_data():
    """Load the actual company data from the Excel file."""
    file_path = os.path.join(IMPORT_FOLDER, "BOT1_Veranstaltungsliste.xlsx")
    return pd.read_excel(file_path)

#
# Basic tests
#
def test_company_from_dataframe(sample_company_data):
    """Test creating Company objects from a DataFrame."""
    companies = Company.from_dataframe(sample_company_data)
    assert len(companies) == 1
    assert companies[0].name == "Company A"
    assert companies[0].capacity == 5
    assert companies[0].min_participants == 2
    assert companies[0].earliest_slot == 0

def test_company_session():
    """Test creating and using a CompanySession."""
    company = Company(
        name="Test Company",
        capacity=5,
        min_participants=2,
        earliest_slot=0,
        blocked_slots=[],
    )

    session = CompanySession(
        company=company, room="101", time_slot="A", time_range="8:45 – 9:30"
    )

    assert session.is_full() == False
    assert session.add_student("10A_1", "Jane Doe") == True
    assert len(session.students) == 1
    
    # Test adding students until capacity is reached
    for i in range(2, 6):
        assert session.add_student(f"10A_{i}", f"Student {i}") == True
    
    # Test that the session is now full
    assert session.is_full() == True
    
    # Test that we can't add more students
    assert session.add_student("10A_6", "Extra Student") == False
    assert len(session.students) == 5

#
# Tests with real data
#
def test_company_from_real_data(real_company_data):
    """Test creating Company objects from real Excel data."""
    companies = Company.from_dataframe(real_company_data)
    
    # Expected number of companies
    assert len(companies) == len(real_company_data)
    
    # Properties of the first company
    first_company = companies[0]
    assert first_company.name == real_company_data.iloc[0]['Unternehmen'].strip()
    assert first_company.capacity == real_company_data.iloc[0]['Max. Teilnehmer']
    
    # Earliest_slot is correctly calculated from the letter
    if pd.notna(real_company_data.iloc[0]['Frühester Zeitpunkt']):
        expected_slot = ord(real_company_data.iloc[0]['Frühester Zeitpunkt'].strip().upper()[0]) - ord('A')
        assert first_company.earliest_slot == expected_slot
    
    # Blocked_slots contains all slots before earliest_slot
    assert first_company.blocked_slots == list(range(first_company.earliest_slot))
