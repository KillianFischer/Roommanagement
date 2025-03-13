from dataclasses import dataclass
from typing import List
import pandas as pd


# not working yet
@dataclass
class Company:
    name: str
    capacity: int
    min_participants: int
    earliest_slot: int
    blocked_slots: List[int]

    @classmethod
    def from_dataframe(cls, df: pd.DataFrame) -> List["Company"]:
        companies = []
        for _, row in df.iterrows():
            # strip extra spaces
            comp_name = str(row["Unternehmen"]).strip()
            fachrichtung = row["Fachrichtung"]
            max_teilnehmer = int(row["Max. Teilnehmer"])
            
            # Handle both "Min. Teilnehmer" and "Min." column names
            if "Min. Teilnehmer" in df.columns:
                min_teilnehmer = int(row["Min. Teilnehmer"])
            elif "Min." in df.columns:
                min_teilnehmer = int(row["Min."])
            else:
                # Default to 1 if no minimum column is found
                min_teilnehmer = 1
                
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
                    min_participants=min_teilnehmer,
                    earliest_slot=earliest,
                    blocked_slots=list(range(earliest)),
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
        # Standard check - never exceed the room's physical capacity
        return len(self.students) >= self.company.capacity
