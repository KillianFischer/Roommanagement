from dataclasses import dataclass
from typing import List, Dict, Optional
import pandas as pd

@dataclass
class StudentPreference:
    student_id: str
    name: str
    wishes: List[str]
    assigned_session: Optional[str] = None
    fulfillment_score: Optional[float] = None

    def calculate_fulfillment_score(self, assigned_session: str) -> float:
        self.assigned_session = assigned_session
        
        if not assigned_session:
            self.fulfillment_score = 0.0
            return 0.0
            
        weights = [6, 5, 4, 3, 2, 1]
        if assigned_session in self.wishes:
            position = self.wishes.index(assigned_session)
            if position < len(weights):
                self.fulfillment_score = float(weights[position])
                return self.fulfillment_score
        
        self.fulfillment_score = 0.0
        return 0.0

    @classmethod
    def from_dataframe(
        cls, df: pd.DataFrame, company_mapping: Dict[int, str] = None
    ) -> List["StudentPreference"]:
        preferences = []
        for idx, row in df.iterrows():
            student_id, full_name = cls._extract_student_info(row, idx)
            wishes = cls._extract_wishes(row, df.columns, company_mapping)
            preferences.append(cls(student_id=student_id, name=full_name, wishes=wishes))
        return preferences
    
    @staticmethod
    def _extract_student_info(row, idx):
        klasse = str(row["Klasse"]).strip()
        name = str(row["Name"]).strip()
        vorname = str(row["Vorname"]).strip()
        student_id = f"{klasse}_{idx + 1}"
        full_name = f"{name}, {vorname}"
        return student_id, full_name
    
    @staticmethod
    def _extract_wishes(row, columns, company_mapping):
        wishes = []
        for i in range(1, 7):
            col1 = f"Wahl {i}"
            col2 = f"Wahl{i}"
            wish = None
            
            if col1 in columns:
                wish = row.get(col1)
            elif col2 in columns:
                wish = row.get(col2)
                
            if pd.notna(wish):
                wishes.append(StudentPreference._format_wish(wish, company_mapping))
        return wishes
    
    @staticmethod
    def _format_wish(wish, company_mapping):
        try:
            wish_num = int(float(str(wish).strip()))
            if company_mapping and wish_num in company_mapping:
                return company_mapping[wish_num]
            return str(wish_num)
        except ValueError:
            return str(wish).strip()

