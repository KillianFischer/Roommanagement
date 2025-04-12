from dataclasses import dataclass
from typing import List
import pandas as pd

@dataclass
class Company:
    name: str
    capacity: int
    max_sessions: int
    earliest_slot: int
    blocked_slots: List[int]
    field: str = ""
    always_show_field: bool = False

    @property
    def unique_id(self) -> str:
        return f"{self.name}_{self.field}" if self.field else self.name
        
    def __str__(self) -> str:
        if self.field:
            return f"{self.name} ({self.field})"
        return self.name

    @classmethod
    def from_dataframe(cls, df: pd.DataFrame) -> List["Company"]:
        companies = []
        for _, row in df.iterrows():
            comp_name = str(row["Unternehmen"]).strip()
            
            field = ""
            if "Fachrichtung" in df.columns and pd.notna(row["Fachrichtung"]):
                field = str(row["Fachrichtung"]).strip()
                
            max_teilnehmer = int(row["Max. Teilnehmer"])
            
            if "Max. Veranstaltungen" in df.columns:
                max_sessions = int(row["Max. Veranstaltungen"])
            else:
                max_sessions = 5 # default 5 if no found
                
            if pd.isna(row["Frühester Zeitpunkt"]):
                earliest = 0
            else:
                earliest = ord(
                    str(row["Frühester Zeitpunkt"]).strip().upper()[0]
                ) - ord("A")
            companies.append(
                cls(
                    name=comp_name,
                    capacity=max_teilnehmer,
                    max_sessions=max_sessions,
                    earliest_slot=earliest,
                    blocked_slots=list(range(earliest)),
                    field=field,
                )
            )
        return companies


@dataclass
class CompanySession:
    company: Company
    room: str
    time_slot: str
    time_range: str
    students: List[dict] = None

    def __post_init__(self):
        if self.students is None:
            self.students = []

    def add_student(self, student_id: str, name: str) -> bool:
        if self.is_full():
            return False
        self.students.append({"id": student_id, "name": name})
        return True

    def is_full(self) -> bool:
        return len(self.students) >= self.company.capacity
        
    def get_company_display_name(self) -> str:
        if self.company.field:
            return f"{self.company.name} ({self.company.field})"
        return self.company.name
