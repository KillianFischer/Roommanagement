from typing import List, Dict, Optional, Tuple, Callable
import pandas as pd
import os
from tkinter import messagebox

from services.scheduler_core import SchedulerCore
from services.pdf_exporter import PDFExporter
from services.attendance_exporter import AttendanceExporter
from models.student import StudentPreference


class Scheduler:
    def __init__(self, error_handler: Optional[Callable[[str], None]] = None):
        self.core = SchedulerCore(error_handler)
        self.pdf_exporter = PDFExporter()
        self.attendance_exporter = AttendanceExporter()
        self.time_slots = self.core.time_slots
        self.error_handler = error_handler

    def clear_error(self):
        """Clear any errors in the scheduler"""
        if self.error_handler:
            self.error_handler("")
            
    def on_error(self, message):
        """Report an error via the error handler"""
        if self.error_handler:
            self.error_handler(message)
        else:
            messagebox.showerror("Error", message)

    def load_student_preferences(self, df: pd.DataFrame) -> bool:
        return self.core.load_student_preferences(df)

    def load_companies(self, df: pd.DataFrame) -> bool:
        self.clear_error()
        required_columns = [
            "Unternehmen",
            "Max. Teilnehmer",
            "Min. Teilnehmer",
            "Frühester Zeitpunkt",
        ]

        # Need to handle both field name options
        field_column = None
        if "Fachrichtung" in df.columns:
            field_column = "Fachrichtung"
        
        # Check required columns
        for col in required_columns:
            if col not in df.columns:
                self.on_error(f"Erforderliche Spalte fehlt: {col}")
                return False

        try:
            # Load companies with specialization field if available
            companies = []
            for _, row in df.iterrows():
                name = str(row["Unternehmen"]).strip()
                field = str(row[field_column]).strip() if field_column and pd.notna(row[field_column]) else ""
                
                max_participants = int(row["Max. Teilnehmer"])
                min_participants = int(row["Min. Teilnehmer"])
                
                # Handle earliest slot letter representation (A, B, C, ...)
                earliest_slot = 0  # Default to A (first slot)
                if pd.notna(row["Frühester Zeitpunkt"]):
                    slot_letter = str(row["Frühester Zeitpunkt"]).strip().upper()
                    if slot_letter in ["A", "B", "C", "D", "E"]:
                        earliest_slot = ord(slot_letter) - ord("A")
                
                companies.append(
                    Company(
                        name=name,
                        field=field,
                        capacity=max_participants,
                        min_participants=min_participants,
                        earliest_slot=earliest_slot,
                    )
                )
            
            self.core.companies = companies
            return True
        except Exception as e:
            self.on_error(f"Fehler beim Verarbeiten der Unternehmensliste: {str(e)}")
            return False

    def load_rooms(self, df: pd.DataFrame) -> bool:
        return self.core.load_rooms(df)

    def is_data_loaded(self) -> bool:
        return self.core.is_data_loaded()

    def generate_schedule(self) -> bool:
        """Generate a schedule using the core scheduler"""
        self.clear_error()
        
        if not self.is_data_loaded():
            self.on_error("Bitte laden Sie zuerst alle Daten.")
            return False
        
        try:
            # Use the core scheduler to generate the schedule
            success = self.core.generate_schedule()
            
            if not success:
                self.on_error("Fehler bei der Generierung des Zeitplans.")
                return False
                
            return True
        
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.on_error(f"Fehler bei der Generierung des Zeitplans: {str(e)}")
            return False

    def get_schedule(self):
        """Get the current schedule"""
        return self.core.schedule

    def calculate_overall_fulfillment_score(self) -> float:
        """Calculate the overall score for how well student wishes were fulfilled"""
        if not self.core.schedule:
            return 0
        
        total_students = len(self.core.student_preferences)
        if total_students == 0:
            return 0
        
        total_slots = len(self.time_slots)
        
        # Count how many students got their wishes
        wish_counts = {i: 0 for i in range(1, 7)}  # 1-6 wish numbers
        missing_wish_count = 0
        
        for (company_id, slot_idx), session in self.core.schedule.items():
            if slot_idx == -1:  # Skip excluded companies
                continue
            
            for student in session.students:
                wish_number = student.get("wish_number", None)
                # Check if wish_number is an integer or can be converted to one
                try:
                    if wish_number is not None:
                        wish_number = int(wish_number)
                        if 1 <= wish_number <= 6:
                            wish_counts[wish_number] += 1
                        else:
                            missing_wish_count += 1
                    else:
                        missing_wish_count += 1
                except (ValueError, TypeError):
                    # If wish_number is "-" or some other non-numeric value
                    missing_wish_count += 1
                
        # Calculate a weighted score: 
        # 1st wish = 100%, 2nd = 80%, 3rd = 60%, 4th = 40%, 5th = 20%, 6th = 10%, none = 0%
        weights = {1: 1.0, 2: 0.8, 3: 0.6, 4: 0.4, 5: 0.2, 6: 0.1}
        
        max_possible_score = total_students * total_slots * 1.0  # If everyone gets 1st wish for all slots
        achieved_score = sum(wish_counts[i] * weights[i] for i in range(1, 7))
        
        if max_possible_score == 0:
            return 0
        
        return (achieved_score / max_possible_score) * 100

    def export_student_schedules_pdf(self, filepath: str) -> bool:
        """Export student schedules as PDF"""
        self.clear_error()
        
        if not self.get_schedule():
            self.on_error("No schedule available. Please generate a schedule first.")
            return False
            
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import mm
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
            
            # Create the PDF document
            doc = SimpleDocTemplate(
                filepath,
                pagesize=A4,
                rightMargin=10*mm,
                leftMargin=10*mm,
                topMargin=10*mm,
                bottomMargin=10*mm
            )
            
            story = []
            styles = getSampleStyleSheet()
            
            # Add title
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=16,
                spaceAfter=10
            )
            story.append(Paragraph("Schülerzeitpläne", title_style))
            
            # Add overall score
            overall_score = self.calculate_overall_fulfillment_score()
            score_style = ParagraphStyle(
                'Score',
                parent=styles['Normal'],
                fontSize=12,
                spaceAfter=5
            )
            story.append(Paragraph(f"Gesamter Erfüllungsscore: {overall_score:.1f}%", score_style))
            story.append(Spacer(1, 5*mm))
            
            # Get student schedules organized by class for better readability
            student_schedules = self.get_student_schedules()
            
            # Organize by class
            class_schedules = {}
            for student_name, appointments in student_schedules.items():
                # Find the student to get their class
                student = next((s for s in self.core.student_preferences if s.name == student_name), None)
                if student:
                    class_name = student.student_id.split("_")[0]
                    if class_name not in class_schedules:
                        class_schedules[class_name] = []
                        
                    class_schedules[class_name].append({
                        "name": student_name,
                        "schedule": appointments
                    })
                    
            # Add schedules by class
            for class_name, students in sorted(class_schedules.items()):
                # Add class header
                class_style = ParagraphStyle(
                    'ClassTitle',
                    parent=styles['Heading2'],
                    fontSize=14,
                    spaceAfter=5
                )
                story.append(Paragraph(f"Klasse {class_name}", class_style))
                
                # For each student in this class
                for student in sorted(students, key=lambda x: x["name"]):
                    # Add student name
                    student_style = ParagraphStyle(
                        'StudentName',
                        parent=styles['Heading3'],
                        fontSize=12,
                        spaceAfter=5
                    )
                    story.append(Paragraph(f"{student['name']}", student_style))
                    
                    # Create a table for this student's schedule
                    schedule_data = [["Zeit", "Unternehmen", "Raum", "Wunsch Nr."]]
                    
                    # Sort appointments by time slot
                    for time_slot, time_range, company, room, wish_num in student["schedule"]:
                        schedule_data.append([f"{time_slot} ({time_range})", company, room, str(wish_num) if wish_num != "-" else "-"])
                        
                    # Create table with styling
                    t = Table(schedule_data, colWidths=[40*mm, 70*mm, 30*mm, 30*mm])
                    
                    # Define styles for the table
                    table_style = [
                        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ]
                    
                    # Color the wish numbers by importance
                    for i in range(1, len(schedule_data)):
                        wish = schedule_data[i][3]
                        try:
                            wish_num = int(wish)
                            if wish_num == 1:
                                table_style.append(('TEXTCOLOR', (3, i), (3, i), colors.green))
                            elif wish_num == 2:
                                table_style.append(('TEXTCOLOR', (3, i), (3, i), colors.darkgreen))
                            elif wish_num == 3:
                                table_style.append(('TEXTCOLOR', (3, i), (3, i), colors.blue))
                            elif wish_num in [4, 5, 6]:
                                table_style.append(('TEXTCOLOR', (3, i), (3, i), colors.orange))
                        except:
                            pass
                            
                    t.setStyle(TableStyle(table_style))
                    story.append(t)
                    story.append(Spacer(1, 5*mm))
                    
                # Add space after each class
                story.append(Spacer(1, 5*mm))
                
            # Build the PDF
            doc.build(story)
            return True
            
        except Exception as e:
            self.on_error(f"Error exporting student schedules: {str(e)}")
            return False

    def export_company_overview_pdf(self, filepath: str) -> bool:
        try:
            if not self.is_data_loaded() or not self.get_schedule():
                if self.error_handler:
                    self.error_handler("Bitte laden Sie alle Dateien und generieren Sie einen Zeitplan.")
                else:
                    messagebox.showerror(
                        "Error", "Bitte laden Sie alle Dateien und generieren Sie einen Zeitplan."
                    )
                return False

            self.pdf_exporter.export_company_overview(
                filepath,
                self.get_schedule(),
                self.time_slots
            )
            return True
        except Exception as e:
            if self.error_handler:
                self.error_handler(f"Fehler beim Exportieren: {str(e)}")
            else:
                messagebox.showerror("Error", f"Fehler beim Exportieren: {str(e)}")
            return False

    def export_attendance_lists_pdf(self, filepath: str, preview_mode=False) -> bool:
        """Export attendance lists as PDF"""
        self.clear_error()
        
        if not self.get_schedule():
            self.on_error("No schedule available. Please generate a schedule first.")
            return False
            
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import mm
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
            
            # Create the PDF document
            doc = SimpleDocTemplate(
                filepath,
                pagesize=A4,
                rightMargin=10*mm,
                leftMargin=10*mm,
                topMargin=10*mm,
                bottomMargin=10*mm
            )
            
            story = []
            styles = getSampleStyleSheet()
            
            # Add title
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=16,
                spaceAfter=10
            )
            story.append(Paragraph("Anwesenheitslisten", title_style))
            story.append(Spacer(1, 10*mm))
            
            # Sort sessions by company and time slot for predictable order
            sorted_sessions = sorted(
                self.get_schedule().items(),
                key=lambda x: (x[0][0], x[0][1])  # Sort by company ID then slot
            )
            
            # Add a table for each session
            for (company_id, slot_idx), session in sorted_sessions:
                if slot_idx == -1:  # Skip excluded companies
                    continue
                    
                # Get company name (with field if available)
                company_name = session.get_company_display_name()
                
                # Get time slot information
                slot_letter, time_range = self.time_slots[slot_idx]
                
                # Add company header
                section_style = ParagraphStyle(
                    'SectionTitle',
                    parent=styles['Heading2'],
                    fontSize=14,
                    spaceAfter=5
                )
                story.append(Paragraph(f"{company_name}", section_style))
                
                # Add time slot and room information
                info_style = ParagraphStyle(
                    'Info',
                    parent=styles['Normal'],
                    fontSize=11,
                    spaceAfter=5
                )
                story.append(Paragraph(f"Zeitfenster: {slot_letter} ({time_range}) - Raum: {session.room}", info_style))
                
                # Check if minimum participants are reached
                attendance_data = []
                attendance_header = ["Nr.", "Name", "Klasse", "Anwesend"]
                
                if session.company.min_participants > 0 and len(session.students) < session.company.min_participants:
                    # If not enough students, show a message
                    attendance_data.append(attendance_header)
                    attendance_data.append(["", "Mindest Anzahl nicht erreicht", "", ""])
                    
                    # Create a simple table
                    t = Table(attendance_data, colWidths=[10*mm, 100*mm, 30*mm, 40*mm])
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ]))
                    
                elif len(session.students) == 0:
                    # If no students assigned, show a message
                    attendance_data.append(attendance_header)
                    attendance_data.append(["", "Keine Teilnehmer", "", ""])
                    
                    # Create a simple table
                    t = Table(attendance_data, colWidths=[10*mm, 100*mm, 30*mm, 40*mm])
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ]))
                    
                else:
                    # Add student rows (sorted by name for easier lookup)
                    attendance_data.append(attendance_header)
                    for i, student in enumerate(sorted(session.students, key=lambda x: x["name"]), 1):
                        class_name = student["id"].split("_")[0]
                        attendance_data.append([str(i), student["name"], class_name, ""])
                        
                    # Create the table with styling
                    t = Table(attendance_data, colWidths=[10*mm, 100*mm, 30*mm, 40*mm])
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ]))
                
                story.append(t)
                story.append(Spacer(1, 10*mm))
            
            # Build the PDF
            doc.build(story)
            return True
            
        except Exception as e:
            self.on_error(f"Error exporting attendance lists: {str(e)}")
            return False

    def get_student_schedules(self):
        """Get schedules organized by student"""
        self.clear_error()
        
        if not self.core.schedule:
            return {}
            
        student_schedules = {}
        
        # Process each session
        for (company_id, slot_idx), session in self.core.schedule.items():
            if slot_idx == -1:  # Skip excluded companies
                continue
                
            slot_letter, time_range = self.time_slots[slot_idx]
            company_name = session.get_company_display_name()  # Use display name with field
            room = session.room
            
            # Add this session to each assigned student's schedule
            for student in session.students:
                student_name = student["name"]
                if student_name not in student_schedules:
                    student_schedules[student_name] = []
                
                # Include the wish number if available
                wish_number = student.get("wish_number", "-")
                    
                student_schedules[student_name].append((slot_letter, time_range, company_name, room, wish_number))
                
        # Sort each student's schedule by time slot
        for student_name in student_schedules:
            student_schedules[student_name].sort()
            
        return student_schedules
        
    def get_company_overview(self):
        """
        Get company schedule overview in a format suitable for UI display
        
        Returns:
            dict: Slot letter -> list of (company, room, num_students)
        """
        if not self.is_data_loaded() or not self.get_schedule():
            return {}
            
        slot_to_companies = {}
        
        # Fill in the slots based on company sessions
        for (company_id, slot_idx), session in self.get_schedule().items():
            if slot_idx == -1:  # Skip excluded companies
                continue
                
            slot_letter, _ = self.time_slots[slot_idx]
            
            if slot_letter not in slot_to_companies:
                slot_to_companies[slot_letter] = []
            
            # Use the display name which includes the field if needed
            display_name = session.get_company_display_name() 
                
            slot_to_companies[slot_letter].append((
                display_name,
                session.room,
                len(session.students)
            ))
                
        # Sort each slot's companies by name
        for slot_letter in slot_to_companies:
            slot_to_companies[slot_letter].sort(key=lambda x: x[0].lower())
            
        return slot_to_companies
        
    def get_excluded_companies(self):
        """
        Get list of excluded companies
        
        Returns:
            list: List of excluded company display names
        """
        if not self.is_data_loaded() or not self.get_schedule():
            return []
            
        excluded_companies = []
        
        for (company_id, slot_idx), session in self.get_schedule().items():
            if slot_idx == -1:
                excluded_companies.append(session.get_company_display_name())
                
        return sorted(excluded_companies)
        
    @property
    def student_preferences(self) -> Optional[List[StudentPreference]]:
        return self.core.student_preferences


class Company:
    def __init__(
        self, name, field="", capacity=0, min_participants=0, earliest_slot=0, blocked_slots=None
    ):
        self.name = name
        self.field = field
        self.capacity = capacity
        self.min_participants = min_participants
        self.earliest_slot = earliest_slot
        self.blocked_slots = blocked_slots or []
        # Create a unique identifier that combines name and field
        self.unique_id = f"{name}_{field}" if field else name
        self.always_show_field = False  # Flag to always show field in display name
    
    def __str__(self):
        if self.field:
            return f"{self.name} ({self.field})"
        return self.name


class Session:
    def __init__(self, company, room):
        self.company = company
        self.room = room
        self.students = []
        
    def __str__(self):
        return f"Session({self.company}, {self.room}, {len(self.students)} students)"
        
    def get_company_display_name(self):
        """Return a display name for the company that includes the field if available"""
        if self.company.field:
            return f"{self.company.name} ({self.company.field})"
        return self.company.name
