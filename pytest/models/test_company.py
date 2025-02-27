import pytest
import pandas as pd
from models.company import Company, CompanySession
# not working yet
@pytest.fixture
def sample_company_data():
    return pd.DataFrame({
        'Unternehmen': ['Company A'],
        'Fachrichtung': ['IT'],
        'Max. Teilnehmer': [5],
        'Max. Veranstaltungen': [2],
        'Frühester Zeitpunkt': ['A']
    })

def test_company_from_dataframe(sample_company_data):
    companies = Company.from_dataframe(sample_company_data)
    assert len(companies) == 1
    assert companies[0].name == "Company A"
    assert companies[0].capacity == 5
    assert companies[0].max_sessions == 2
    assert companies[0].earliest_slot == 0

def test_company_session():
    company = Company(
        name="Test Company",
        capacity=5,
        max_sessions=2,
        earliest_slot=0,
        blocked_slots=[]
    )
    
    session = CompanySession(
        company=company,
        room="101",
        time_slot="A",
        time_range="8:45 – 9:30"
    )
    
    assert session.is_full() == False
    assert session.add_student("10A_1", "Jane Doe") == True
    assert len(session.students) == 1