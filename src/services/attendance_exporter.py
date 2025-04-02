from typing import Dict, Tuple, List
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from tkinter import messagebox


class AttendanceExporter:
    def __init__(self):
        self.styles = getSampleStyleSheet()
        
    def export_attendance_lists(self, filepath: str, schedule: Dict[Tuple[str, int], any], 
                              time_slots: List[Tuple[str, str]], preview_mode=False):
        try:
            os.makedirs(os.path.dirname(filepath), exist_ok=True)
            
            doc = SimpleDocTemplate(
                filepath,
                pagesize=A4,
                rightMargin=10 * mm,
                leftMargin=10 * mm,
                topMargin=10 * mm,
                bottomMargin=10 * mm,
            )

            story = []

            sorted_sessions = sorted(
                schedule.items(),
                key=lambda x: (x[0][0], x[0][1]),
            )

            if preview_mode:
                company_names = list(
                    set(company_name for (company_name, _), _ in sorted_sessions)
                )
                if len(company_names) > 6:
                    company_names = company_names[:6]
                sorted_sessions = [
                    (key, session)
                    for (key, session) in sorted_sessions
                    if key[0] in company_names
                ]

            for (company_name, slot_idx), session in sorted_sessions:
                # Skip excluded companies (slot_idx == -1)
                if slot_idx == -1:
                    continue
                    
                slot_letter, time_range = time_slots[slot_idx]
                
                story.append(
                    Paragraph(
                        f"<b>{company_name}</b>",
                        self.styles["Heading1"],
                    )
                )
                
                story.append(
                    Paragraph(
                        f"Zeitfenster: {slot_letter} ({time_range})<br/>"
                        f"Raum: {session.room}",
                        self.styles["Normal"],
                    )
                )

                data = [["Nr.", "Name", "Klasse", "Anwesend"]]
                
                if len(session.students) == 0:
                    data.append(["", "Kein Schülerinteresse", "", ""])
                else:
                    for i, student in enumerate(
                        sorted(session.students, key=lambda x: x["name"]), 1
                    ):
                        class_name = student["id"].split("_")[0]
                        data.append([str(i), student["name"], class_name, ""])

                    # Check if there are no students
                    current_students = len(session.students)
                    if current_students == 0:
                        data.append(["", "Keine Teilnehmer", "", ""])

                t = Table(
                    data,
                    colWidths=[20 * mm, 80 * mm, 30 * mm, 50 * mm],
                    style=TableStyle(
                        [
                            ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
                            ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                            ("FONTSIZE", (0, 0), (-1, 0), 10),
                            ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
                            ("BACKGROUND", (0, 1), (-1, -1), colors.white),
                            ("TEXTCOLOR", (0, 1), (-1, -1), colors.black),
                            ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                            ("FONTSIZE", (0, 1), (-1, -1), 10),
                            ("TOPPADDING", (0, 1), (-1, -1), 6),
                            ("BOTTOMPADDING", (0, 1), (-1, -1), 6),
                            ("LEADING", (0, 1), (-1, -1), 8),
                        ]
                    ),
                )
                story.append(t)
                story.append(Spacer(1, 12))

            doc.build(story)
            return True

        except Exception as e:
            messagebox.showerror(
                "Export Error",
                f"Fehler beim Exportieren der Anwesenheitslisten: {str(e)}",
            )
            return False 