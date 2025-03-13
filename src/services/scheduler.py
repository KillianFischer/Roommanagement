from typing import List, Dict, Optional, Tuple
import pandas as pd
from tkinter import messagebox

from models.student import StudentPreference
from models.company import Company, CompanySession


class SchedulerService:
    def __init__(self):
        self.student_preferences: Optional[List[StudentPreference]] = None
        self.companies: Optional[List[Company]] = None
        self.rooms: Optional[List[str]] = None
        # Schedule: maps, company name, slot
        self.schedule: Dict[Tuple[str, int], CompanySession] = {}
        # list of tuples: slot letter, time range
        self.time_slots = [
            ("A", "8:45 – 9:30"),
            ("B", "9:50 – 10:35"),
            ("C", "10:35 – 11:20"),
            ("D", "11:40 – 12:25"),
            ("E", "12:25 – 13:10"),
        ]

    def load_student_preferences(self, df: pd.DataFrame) -> bool:
        if df is None or df.empty:
            return False

        company_mapping = {}
        if self.companies:
            for idx, company in enumerate(self.companies, 1):
                normalized_name = company.name.strip()
                company_mapping[idx] = normalized_name
                company_mapping[str(idx)] = normalized_name

        df.columns = df.columns.str.strip()
        self.student_preferences = StudentPreference.from_dataframe(df, company_mapping)
        return True

    def load_companies(self, df: pd.DataFrame) -> bool:
        if df is None or df.empty:
            return False
        df.columns = df.columns.str.strip()
        self.companies = Company.from_dataframe(df)
        return True

    def load_rooms(self, df: pd.DataFrame) -> bool:
        if df is None or df.empty:
            return False
        
        # Check if "Raum" is in the columns, if not use the first column
        if "Raum" in df.columns:
            self.rooms = [str(row["Raum"]).strip() for _, row in df.iterrows()]
        else:
            # Use the first column (index 0)
            self.rooms = [str(row[0]).strip() for _, row in df.iterrows()]
        
        return True

    def is_data_loaded(self) -> bool:
        return (
            self.student_preferences is not None
            and self.companies is not None
            and self.rooms is not None
        )

    def generate_schedule(self) -> bool:
        try:
            company_to_number = {}
            number_to_company = {}
            for idx, company in enumerate(self.companies, 1):
                normalized_name = company.name.strip()
                company_to_number[normalized_name] = str(idx)
                number_to_company[str(idx)] = normalized_name
                number_to_company[idx] = normalized_name

            all_wish_counts = {}
            first_wish_counts = {}
            
            for student in self.student_preferences:
                if not student.wishes:
                    continue
                    
                first_wish = str(student.wishes[0]).strip()
                try:
                    wish_num = int(float(first_wish))
                    company_name = number_to_company.get(wish_num, first_wish)
                except (ValueError, TypeError):
                    company_name = first_wish
                first_wish_counts[company_name] = first_wish_counts.get(company_name, 0) + 1
                
                for wish in student.wishes:
                    if not wish:
                        continue
                    try:
                        wish_num = int(float(str(wish).strip()))
                        company_name = number_to_company.get(wish_num, str(wish).strip())
                    except (ValueError, TypeError):
                        company_name = str(wish).strip()
                    all_wish_counts[company_name] = all_wish_counts.get(company_name, 0) + 1

            excluded_companies = []
            filtered_companies = []
            for company in self.companies:
                normalized_name = company.name.strip()
                wish_count = all_wish_counts.get(normalized_name, 0)
                if wish_count >= company.min_participants:
                    filtered_companies.append(company)
                else:
                    excluded_companies.append(company)

            sorted_companies = sorted(
                filtered_companies,
                key=lambda x: first_wish_counts.get(x.name.strip(), 0),
                reverse=True,
            )

            self.schedule.clear()
            company_rooms = {}
            available_rooms = self.rooms.copy()
            
            # Remove Aula from available rooms to ensure only Polizei gets it
            available_rooms = [room for room in available_rooms if room.strip().lower() != "aula"]

            for company in excluded_companies:
                company_rooms[company.name] = "Hat nicht die Min. Teilnehmer erreicht"
                session = CompanySession(
                    company=company,
                    room="Hat nicht die Min. Teilnehmer erreicht",
                    time_slot="-",
                    time_range="-",
                )
                self.schedule[(company.name, -1)] = session

            # Handle special company: Only Polizei gets Aula
            polizei_company = next(
                (
                    company
                    for company in sorted_companies
                    if "polizei" in company.name.strip().lower()
                ),
                None,
            )
            if polizei_company:
                company_rooms["Polizei"] = "Aula"
                sorted_companies.remove(polizei_company)
                for slot_idx, (slot_letter, time_range) in enumerate(self.time_slots):
                    if slot_idx >= polizei_company.earliest_slot:
                        session = CompanySession(
                            company=polizei_company,
                            room="Aula",
                            time_slot=slot_letter,
                            time_range=time_range,
                        )
                        self.schedule[(polizei_company.name, slot_idx)] = session
            
            sorted_companies = sorted(
                [c for c in self.companies if c.name.strip() != "Polizei"],
                key=lambda x: x.capacity,
                reverse=True,
            )

            for company in sorted_companies:
                if not available_rooms:
                    available_rooms = self.rooms.copy()
                    if "Aula" in available_rooms:
                        available_rooms.remove("Aula")
                
                company_room = available_rooms.pop(0)
                company_rooms[company.name] = company_room
                
                company_name = company.name.strip()
                total_interest = all_wish_counts.get(company_name, 0)
                
                # If interest ≤ 20: 1 session
                # If 20 < interest ≤ 40: 2 sessions
                # If 40 < interest ≤ 60: 3 sessions, etc.
                if total_interest <= 20:
                    needed_slots = 1
                else:
                    # Calculate needed slots based on 20 students per session rule
                    needed_slots = (total_interest + 19) // 20  # Ceiling division by 20
                    needed_slots = min(needed_slots, len(self.time_slots) - company.earliest_slot)
                
                # Create the required number of sessions
                for slot_offset in range(needed_slots):
                    slot_idx = company.earliest_slot + slot_offset
                    slot_letter, time_range = self.time_slots[slot_idx]



                    session = CompanySession(
                        company=company,
                        room=company_room,
                        time_slot=slot_letter,
                        time_range=time_range,
                    )
                    self.schedule[(company.name, slot_idx)] = session

            # Company_sessions for student assignment
            company_sessions = {}
            for (company_name, slot_idx), session in self.schedule.items():
                if company_name not in company_sessions:
                    company_sessions[company_name] = []
                company_sessions[company_name].append((slot_idx, session))
            
            for company_name in company_sessions:
                company_sessions[company_name].sort(key=lambda x: x[0])
                
                if len(company_sessions[company_name]) > 1:
                    total_interest = all_wish_counts.get(company_name, 0)
                    sessions_count = len(company_sessions[company_name])
                    ideal_per_session = total_interest / sessions_count

            for student in self.student_preferences:
                assigned_slots = set()
                
                for wish_idx, wish in enumerate(student.wishes):
                    if not wish:
                        continue
                        
                    try:
                        wish_num = int(float(str(wish).strip()))
                        company_name = number_to_company.get(wish_num, str(wish).strip())
                    except (ValueError, TypeError):
                        company_name = str(wish).strip()
                    
                    if (company_name, -1) in self.schedule:
                        continue
                    
                    # All available sessions for this company
                    if company_name in company_sessions:
                        sessions = company_sessions[company_name]
                        
                        # Sessions that are not in assigned slots
                        available_sessions = []
                        for slot_idx, session in sessions:
                            if slot_idx not in assigned_slots:
                                available_sessions.append((slot_idx, session, len(session.students)))
                        
                        if available_sessions:
                            if len(sessions) > 1:
                                total_interest = all_wish_counts.get(company_name, 0)
                                ideal_per_session = total_interest / len(sessions)
                                
                                available_sessions.sort(key=lambda x: abs(x[2] - ideal_per_session))
                            
                            slot_idx, session, current_count = available_sessions[0]
                            
                            if current_count < session.company.capacity:
                                session.add_student(student.student_id, student.name)
                                assigned_slots.add(slot_idx)

            return True

        except Exception as e:
            messagebox.showerror(
                "Error", f"Fehler bei der Zeitplangenerierung: {str(e)}"
            )
            self.schedule.clear()
            return False

    def get_schedule(self) -> Dict[Tuple[str, int], CompanySession]:
        return self.schedule

    def export_student_schedules(self):
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import mm
            from reportlab.platypus import (
                SimpleDocTemplate,
                Table,
                TableStyle,
                Paragraph,
            )

            number_to_company = {}
            for idx, company in enumerate(self.companies, 1):
                normalized_name = company.name.strip()
                number_to_company[str(idx)] = normalized_name
                number_to_company[idx] = normalized_name

            class_schedules = {}
            for student in self.student_preferences:
                class_name = student.student_id.split("_")[0]
                if class_name not in class_schedules:
                    class_schedules[class_name] = []

                student_schedule = []
                realized_wishes = [False] * len(student.wishes)
                
                for wish_idx, wish in enumerate(student.wishes):
                    if not wish:
                        continue
                        
                    try:
                        wish_num = int(float(str(wish).strip()))
                        company_name = number_to_company.get(wish_num, str(wish).strip())
                    except (ValueError, TypeError):
                        company_name = str(wish).strip()
                    
                    for slot_idx, (slot_letter, time_range) in enumerate(self.time_slots):
                        key = (company_name, slot_idx)
                        if key in self.schedule:
                            session = self.schedule[key]
                            if any(s["id"] == student.student_id for s in session.students):
                                realized_wishes[wish_idx] = True
                                student_schedule.append({
                                    "time": f"{slot_letter} ({time_range})",
                                    "company": company_name,
                                    "room": session.room,
                                    "wish_number": wish_idx + 1,
                                })
                                break

                satisfaction_score = student.get_satisfaction_score(realized_wishes)
                class_schedules[class_name].append(
                    {
                        "name": student.name,
                        "schedule": sorted(student_schedule, key=lambda x: x["time"]),
                        "score": satisfaction_score,
                    }
                )

            doc = SimpleDocTemplate(
                "student_schedules.pdf",
                pagesize=A4,
                rightMargin=10 * mm,
                leftMargin=10 * mm,
                topMargin=10 * mm,
                bottomMargin=10 * mm,
            )

            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                "CustomTitle", parent=styles["Heading1"], fontSize=12, spaceAfter=10
            )
            class_style = ParagraphStyle(
                "ClassTitle", parent=styles["Heading1"], fontSize=14, spaceAfter=10
            )

            story = []

            for class_name, students in sorted(class_schedules.items()):
                story.append(Paragraph(f"Klasse {class_name}", class_style))
                
                students_processed = 0
                while students_processed < len(students):
                    page_students = students[students_processed:students_processed + 4]

                    for student in page_students:
                        story.append(
                            Paragraph(
                                f"{student['name']} - Bewertung: {student['score']:.1f}%",
                                title_style,
                            )
                        )

                        schedule_data = [["Zeit", "Unternehmen", "Raum", "Wunsch"]]
                        for appointment in student["schedule"]:
                            schedule_data.append(
                                [
                                    appointment["time"],
                                    appointment["company"],
                                    appointment["room"],
                                    str(appointment["wish_number"]),
                                ]
                            )

                        t = Table(
                            schedule_data,
                            colWidths=[60 * mm, 60 * mm, 30 * mm, 20 * mm],
                            style=TableStyle(
                                [
                                    ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
                                    ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                                    ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                                    ("ALIGN", (0, 1), (-1, -1), "LEFT"),
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
                        story.append(Paragraph("<br/><br/>", styles["Normal"]))

                    students_processed += 4

            doc.build(story)
            messagebox.showinfo(
                "Export erfolgreich",
                "Schülerzeitpläne wurden unter student_schedules.pdf gespeichert.",
            )

        except Exception as e:
            messagebox.showerror(
                "Export Error",
                f"Fehler beim Exportieren der Schülerzeitpläne: {str(e)}",
            )

    def export_attendance_lists(self, preview_mode=False):
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.styles import getSampleStyleSheet
            from reportlab.lib.units import mm
            from reportlab.platypus import (
                SimpleDocTemplate,
                Table,
                TableStyle,
                Paragraph,
            )

            doc = SimpleDocTemplate(
                "attendance_lists.pdf",
                pagesize=A4,
                rightMargin=10 * mm,
                leftMargin=10 * mm,
                topMargin=10 * mm,
                bottomMargin=10 * mm,
            )

            styles = getSampleStyleSheet()
            story = []

            sorted_sessions = sorted(
                self.schedule.items(),
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
                story.append(
                    Paragraph(
                        f"<b>{company_name}</b>",
                        styles["Heading1"],
                    )
                )
                
                story.append(
                    Paragraph(
                        f"Zeitfenster: {session.time_slot} ({session.time_range})<br/>"
                        f"Raum: {session.room}",
                        styles["Normal"],
                    )
                )

                data = [["Nr.", "Name", "Klasse", "Unterschrift"]]
                for i, student in enumerate(
                    sorted(session.students, key=lambda x: x["name"]), 1
                ):
                    class_name = student["id"].split("_")[0]
                    data.append([str(i), student["name"], class_name, ""])

                empty_rows = [["", "", "", ""] for _ in range(5)]
                for i, empty_row in enumerate(empty_rows, len(data)):
                    empty_row[0] = str(i)
                data.extend(empty_rows)

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
                story.append(Paragraph("<br/><br/>", styles["Normal"]))

            doc.build(story)
            messagebox.showinfo(
                "Export erfolgreich",
                "Anwesenheitslisten wurden unter attendance_lists.pdf gespeichert.",
            )

        except Exception as e:
            messagebox.showerror(
                "Export Error",
                f"Fehler beim Exportieren der Anwesenheitslisten: {str(e)}",
            )
