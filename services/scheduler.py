from typing import List, Dict, Optional, Tuple
import pandas as pd
from tkinter import messagebox

from models.student import StudentPreference
from models.company import Company, CompanySession
# does not work yet
class SchedulerService:
    def __init__(self):
        self.student_preferences: Optional[List[StudentPreference]] = None
        self.companies: Optional[List[Company]] = None
        self.rooms: Optional[List[str]] = None
        # Schedule: maps, company name, slot index
        self.schedule: Dict[Tuple[str, int], CompanySession] = {}
        # list of tuples: slot letter, time range
        self.time_slots = [
            ('A', '8:45 – 9:30'),
            ('B', '9:50 – 10:35'),
            ('C', '10:35 – 11:20'),
            ('D', '11:40 – 12:25'),
            ('E', '12:25 – 13:10')
        ]

    def load_student_preferences(self, df: pd.DataFrame) -> bool:
        if df is None or df.empty:
            return False
        
        # mapping from company number to name
        company_mapping = {}
        if self.companies:
            for idx, company in enumerate(self.companies, 1):
                company_mapping[idx] = company.name.strip()
                company_mapping[str(idx)] = company.name.strip()
        
        # strip extra spaces
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
        # ToDo: not ignoring "Aula"
        self.rooms = [str(int(room)) for room in df.iloc[:, 0].tolist() 
                      if pd.notna(room) and str(room).strip() != 'Aula' and str(room).strip().isdigit()]
        return True

    def is_data_loaded(self) -> bool:
        return all([
            self.student_preferences is not None and len(self.student_preferences) > 0,
            self.companies is not None and len(self.companies) > 0,
            self.rooms is not None and len(self.rooms) > 0
        ])

    def generate_schedule(self) -> bool:
        try:
            # mapping between company names and numbers
            company_to_number = {}
            number_to_company = {}
            for idx, company in enumerate(self.companies, 1):
                normalized_name = company.name.strip()
                company_to_number[normalized_name] = str(idx)
                number_to_company[str(idx)] = normalized_name

            # first wishes to ensure capacity
            first_wish_counts = {}
            for student in self.student_preferences:
                if student.wishes:
                    first_wish = str(student.wishes[0]).strip()
                    # Convert company name to number if it's a name
                    wish_number = company_to_number.get(first_wish, first_wish)
                    first_wish_counts[wish_number] = first_wish_counts.get(wish_number, 0) + 1

            # Calculate sessions needed for first wishes first
            sessions_per_company = {}
            for company in self.companies:
                normalized_name = company.name.strip()
                company_number = company_to_number[normalized_name]
                first_wish_count = first_wish_counts.get(company_number, 0)
                if first_wish_count > 0:
                    min_sessions = -(-first_wish_count // company.capacity)
                    sessions_per_company[normalized_name] = min(min_sessions, company.max_sessions)

            # Sort companies by number of first wishes
            sorted_companies = sorted(
                self.companies,
                key=lambda x: first_wish_counts.get(company_to_number.get(x.name.strip(), ''), 0),
                reverse=True
            )

            # Count all wishes for total session calculation
            company_counts = {}
            for student in self.student_preferences:
                for comp in student.wishes:
                    comp_normalized = str(comp).strip()
                    company_counts[comp_normalized] = company_counts.get(comp_normalized, 0) + 1

            self.schedule.clear()
            company_rooms = {}
            available_rooms = self.rooms.copy()

            # Distribute companies across time slots
            slot_companies = {i: [] for i in range(len(self.time_slots))}

            for company in sorted_companies:
                if not available_rooms:
                    available_rooms = self.rooms.copy()
                company_room = available_rooms.pop(0)
                company_rooms[company.name] = company_room

                sessions_needed = sessions_per_company[company.name]

                # Create sessions for all allocated slots
                for slot_idx in range(len(self.time_slots)):
                    slot_letter, time_range = self.time_slots[slot_idx]
                    session = CompanySession(
                        company=company,
                        room=company_room,
                        time_slot=slot_letter,
                        time_range=time_range
                    )
                    self.schedule[(company.name, slot_idx)] = session
                    slot_companies[slot_idx].append(company.name)

            # Assign students to sessions
            student_assignments = {s.student_id: set() for s in self.student_preferences}
            
            # First, fulfill first wishes
            for student in self.student_preferences:
                if not student.wishes:
                    continue
                first_wish = str(student.wishes[0]).strip()
                assigned = False
                for slot_idx in range(len(self.time_slots)):
                    key = (first_wish, slot_idx)
                    if key in self.schedule and not self.schedule[key].is_full():
                        self.schedule[key].add_student(student.student_id, student.name)
                        student_assignments[student.student_id].add(slot_idx)
                        assigned = True
                        break
                if not assigned:
                    messagebox.showerror(
                        "Fehler",
                        f"Erster Wunsch konnte nicht erfüllt werden für: {student.name}"
                    )
                    return False

            # Then handle remaining wishes
            for student in self.student_preferences:
                for wish in student.wishes[1:]:
                    wish = str(wish).strip()
                    for slot_idx in range(len(self.time_slots)):
                        if slot_idx in student_assignments[student.student_id]:
                            continue
                        key = (wish, slot_idx)
                        if key in self.schedule and not self.schedule[key].is_full():
                            self.schedule[key].add_student(student.student_id, student.name)
                            student_assignments[student.student_id].add(slot_idx)
                            break

            return True

        except Exception as e:
            messagebox.showerror("Error", f"Fehler bei der Zeitplangenerierung: {str(e)}")
            self.schedule.clear()
            return False

    def get_schedule(self) -> Dict[Tuple[str,int], CompanySession]:
        return self.schedule

    def export_student_schedules(self):
        """
        Exportiert Schülerzeitpläne als PDF mit 4 Schülern pro Seite,
        sortiert nach Klassen.
        """
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import mm
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
            
            # Gruppiere Schüler nach Klassen und erstelle ihre Zeitpläne
            class_schedules = {}
            for student in self.student_preferences:
                class_name = student.student_id.split('_')[0]
                if class_name not in class_schedules:
                    class_schedules[class_name] = []
                
                # Sammle alle Termine für diesen Schüler
                student_schedule = []
                realized_wishes = []
                for slot_idx, (slot_letter, time_range) in enumerate(self.time_slots):
                    session_found = False
                    for wish_idx, wish in enumerate(student.wishes):
                        key = (str(wish).strip(), slot_idx)
                        if key in self.schedule:
                            session = self.schedule[key]
                            if any(s['id'] == student.student_id for s in session.students):
                                student_schedule.append({
                                    'time': f"{slot_letter} ({time_range})",
                                    'company': str(wish).strip(),
                                    'room': session.room,
                                    'wish_number': wish_idx + 1
                                })
                                realized_wishes.append(True)
                                session_found = True
                                break
                    if not session_found:
                        realized_wishes.append(False)
                
                # Berechne Erfüllungsscore
                satisfaction_score = student.get_satisfaction_score(realized_wishes)
                
                class_schedules[class_name].append({
                    'name': student.name,
                    'schedule': sorted(student_schedule, key=lambda x: x['time']),
                    'score': satisfaction_score
                })

            # Create PDF
            doc = SimpleDocTemplate(
                "student_schedules.pdf",
                pagesize=A4,
                rightMargin=10*mm,
                leftMargin=10*mm,
                topMargin=10*mm,
                bottomMargin=10*mm
            )
            
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=12,
                spaceAfter=10
            )
            
            story = []
            
            # For each class
            for class_name, students in sorted(class_schedules.items()):
                students_processed = 0
                while students_processed < len(students):
                    # Take up to 4 students for a page
                    page_students = students[students_processed:students_processed+4]
                    
                    # Create 4 equal areas on the page
                    for student in page_students:
                        # Header for each student
                        story.append(Paragraph(
                            f"{student['name']} - Klasse {class_name} - Score: {student['score']:.1f}%",
                            title_style
                        ))
                        
                        # Schedule table
                        schedule_data = [['Time', 'Company', 'Room', 'Wish']]
                        for appointment in student['schedule']:
                            schedule_data.append([
                                appointment['time'],
                                appointment['company'],
                                appointment['room'],
                                str(appointment['wish_number'])
                            ])
                        
                        t = Table(
                            schedule_data,
                            colWidths=[60*mm, 60*mm, 30*mm, 20*mm],
                            style=TableStyle([
                                ('GRID', (0,0), (-1,-1), 0.25, colors.black),
                                ('BACKGROUND', (0,0), (-1,0), colors.grey),
                                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                                ('FONTSIZE', (0,0), (-1,0), 10),
                                ('BOTTOMPADDING', (0,0), (-1,0), 12),
                                ('BACKGROUND', (0,1), (-1,-1), colors.white),
                                ('TEXTCOLOR', (0,1), (-1,-1), colors.black),
                                ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
                                ('FONTSIZE', (0,1), (-1,-1), 10),
                                ('TOPPADDING', (0,1), (-1,-1), 6),
                                ('BOTTOMPADDING', (0,1), (-1,-1), 6),
                                ('LEADING', (0,1), (-1,-1), 8)
                            ])
                        )
                        story.append(t)
                        story.append(Paragraph("<br/><br/>", styles['Normal']))
                    
                    students_processed += 4
            
            doc.build(story)
            messagebox.showinfo(
                "Export erfolgreich",
                "Schülerzeitpläne wurden unter student_schedules.pdf gespeichert."
            )
            
        except Exception as e:
            messagebox.showerror(
                "Export Error",
                f"Fehler beim Exportieren der Schülerzeitpläne: {str(e)}"
            )

    def export_attendance_lists(self):
        """
        Exportiert Anwesenheitslisten für jede Veranstaltung als PDF.
        """
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.styles import getSampleStyleSheet
            from reportlab.lib.units import mm
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
            
            doc = SimpleDocTemplate(
                "attendance_lists.pdf",
                pagesize=A4,
                rightMargin=10*mm,
                leftMargin=10*mm,
                topMargin=10*mm,
                bottomMargin=10*mm
            )
            
            styles = getSampleStyleSheet()
            story = []
            
            # Sort by company name and time slot
            sorted_sessions = sorted(
                self.schedule.items(),
                key=lambda x: (x[0][0], x[0][1])  # Sort by company name, then slot
            )
            
            for (company_name, slot_idx), session in sorted_sessions:
                # Header
                story.append(Paragraph(
                    f"<b>{company_name}</b><br/>"
                    f"Time slot: {session.time_slot} ({session.time_range})<br/>"
                    f"Room: {session.room}",
                    styles['Heading1']
                ))
                
                # Attendee list
                data = [['Nr.', 'Name', 'Class', 'Signature']]
                for i, student in enumerate(sorted(session.students, key=lambda x: x['name']), 1):
                    class_name = student['id'].split('_')[0]
                    data.append([str(i), student['name'], class_name, ''])
                
                # does not work yet
                for i in range(5):
                    data.append([str(len(data)), '', '', ''])
                
                t = Table(
                    data,
                    colWidths=[20*mm, 80*mm, 30*mm, 50*mm],
                    style=TableStyle([
                        ('GRID', (0,0), (-1,-1), 0.25, colors.black),
                        ('BACKGROUND', (0,0), (-1,0), colors.grey),
                        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,0), (-1,0), 10),
                        ('BOTTOMPADDING', (0,0), (-1,0), 12),
                        ('BACKGROUND', (0,1), (-1,-1), colors.white),
                        ('TEXTCOLOR', (0,1), (-1,-1), colors.black),
                        ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
                        ('FONTSIZE', (0,1), (-1,-1), 10),
                        ('TOPPADDING', (0,1), (-1,-1), 6),
                        ('BOTTOMPADDING', (0,1), (-1,-1), 6)
                    ])
                )
                story.append(t)
                story.append(Paragraph("<br/><br/>", styles['Normal']))
            
            doc.build(story)
            messagebox.showinfo(
                "Export erfolgreich",
                "Anwesenheitslisten wurden unter attendance_lists.pdf gespeichert."
            )
            
        except Exception as e:
            messagebox.showerror(
                "Export Error",
                f"Fehler beim Exportieren der Anwesenheitslisten: {str(e)}"
            )
