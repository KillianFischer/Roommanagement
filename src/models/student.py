from dataclasses import dataclass
from typing import List, Dict
import pandas as pd
# not working yet
@dataclass
class StudentPreference:
    student_id: str   
    name: str         
    wishes: List[str]

    @classmethod
    def from_dataframe(cls, df: pd.DataFrame, company_mapping: Dict[int, str] = None) -> List['StudentPreference']:
        preferences = []
        for idx, row in df.iterrows():
            klasse = str(row['Klasse']).strip()
            name = str(row['Name']).strip()
            vorname = str(row['Vorname']).strip()
            student_id = f"{klasse}_{idx+1}"
            full_name = f"{name}, {vorname}"
            wishes = []
            for i in range(1, 7):
                col1 = f'Wahl {i}'
                col2 = f'Wahl{i}'
                wish = None
                if col1 in df.columns:
                    wish = row.get(col1)
                elif col2 in df.columns:
                    wish = row.get(col2)
                if pd.notna(wish):
                    try:
                        wish_num = int(float(str(wish).strip()))
                        if company_mapping and wish_num in company_mapping:
                            wishes.append(company_mapping[wish_num])
                        else:
                            wishes.append(str(wish_num))
                    except ValueError:
                        wishes.append(str(wish).strip())
            preferences.append(cls(student_id=student_id, name=full_name, wishes=wishes))
        return preferences

    def get_satisfaction_score(self, realized_wishes: List[bool], max_wishes: int = 6) -> float:
        total_points = 0
        max_points = sum(range(max_wishes, 0, -1))
        for i, wish in enumerate(realized_wishes):
            if wish:
                total_points += (max_wishes - i)
        return (total_points / max_points) * 100
