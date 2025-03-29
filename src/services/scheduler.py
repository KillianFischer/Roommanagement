from typing import List, Optional, Callable
import pandas as pd
from tkinter import messagebox

from services.scheduler_core import SchedulerCore
from services.pdf_exporter import PDFExporter
from services.excel_exporter import ExcelExporter
from services.attendance_exporter import AttendanceExporter
from models.student import StudentPreference


class Scheduler:
    def __init__(self, error_handler: Optional[Callable[[str], None]] = None):
        self.core = SchedulerCore(error_handler)
        self.pdf_exporter = PDFExporter()
        self.excel_exporter = ExcelExporter()
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
            "Max. Veranstaltungen",
            "Frühester Zeitpunkt",
        ]

        field_column = None
        if "Fachrichtung" in df.columns:
            field_column = "Fachrichtung"
        
        # Check required columns
        for col in required_columns:
            if col not in df.columns:
                # Special case for backward compatibility
                if col == "Max. Veranstaltungen" and "Min. Teilnehmer" in df.columns:
                    continue  # Allow using Min. Teilnehmer for backward compatibility
                self.on_error(f"Erforderliche Spalte fehlt: {col}")
                return False

        try:
            # Load companies with specialization field if available
            companies = []
            for _, row in df.iterrows():
                name = str(row["Unternehmen"]).strip()
                field = str(row[field_column]).strip() if field_column and pd.notna(row[field_column]) else "" # FIXME
                
                max_participants = int(row["Max. Teilnehmer"])

                if "Max. Veranstaltungen" in df.columns:
                    max_sessions = int(row["Max. Veranstaltungen"])
                else:
                    # FIXME: Remove this
                    max_sessions = int(row["Min. Teilnehmer"])
                
                earliest_slot = 0  # Default to A (first slot)
                if pd.notna(row["Frühester Zeitpunkt"]): # FIXME
                    slot_letter = str(row["Frühester Zeitpunkt"]).strip().upper()
                    if slot_letter in ["A", "B", "C", "D", "E"]:
                        earliest_slot = ord(slot_letter) - ord("A")
                
                companies.append(
                    Company(
                        name=name,
                        field=field,
                        capacity=max_participants,
                        max_sessions=max_sessions,
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
            # Invoke the core scheduler.py
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
        
        fulfillment_data = self.core.calculate_fulfillment()
        if not fulfillment_data or "overall_stats" not in fulfillment_data:
            return 0
            
        return fulfillment_data["overall_stats"].get("fulfillment_percentage", 0)
        
    def get_student_fulfillment_scores(self) -> pd.DataFrame:
        if not self.core.schedule:
            return pd.DataFrame()
            
        fulfillment_data = self.core.calculate_fulfillment()
        if not fulfillment_data or "by_student" not in fulfillment_data:
            return pd.DataFrame()
            
        student_data = []
        
        for student_id, data in fulfillment_data["by_student"].items():
            weighted_pct = data.get("weighted_fulfillment", 0)
            
            # Count wishes by rank
            wish_counts = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, None: 0}
            for _, wish_number in data.get("wishes_fulfilled", []):
                if wish_number in wish_counts:
                    wish_counts[wish_number] += 1
                
            # Prepare record for DataFrame
            record = {
                "Student ID": student_id,
                "Name": data.get("name", ""),
                "Total Score": data.get("weighted_score", 0),
                "Max Score": 21,  # Maximum possible weight sum (6+5+4+3+2+1)
                "Fulfillment %": round(weighted_pct, 2),
                "1st Wishes": wish_counts.get(1, 0),
                "2nd Wishes": wish_counts.get(2, 0),
                "3rd Wishes": wish_counts.get(3, 0),
                "4th Wishes": wish_counts.get(4, 0),
                "5th Wishes": wish_counts.get(5, 0),
                "6th Wishes": wish_counts.get(6, 0),
                "No Match": data.get("total_sessions", 0) - len(data.get("wishes_fulfilled", [])),
            }
            
            # Add schedule for each time slot
            student_schedule = self.get_student_schedule(student_id)
            for slot_letter, _ in self.time_slots:
                slot_entry = next((e for e in student_schedule if e[0] == slot_letter), None)
                if slot_entry:
                    record[f"Slot {slot_letter}"] = slot_entry[2]  # Company name
                    record[f"Wish {slot_letter}"] = slot_entry[4]  # Wish number
                else:
                    record[f"Slot {slot_letter}"] = ""
                    record[f"Wish {slot_letter}"] = ""
                
            student_data.append(record)
            
        # Create DataFrame
        df = pd.DataFrame(student_data)
        
        # Add overall statistics
        if len(df) > 0:
            overall_stats = fulfillment_data["overall_stats"]
            print(f"Overall weighted fulfillment: {overall_stats.get('average_weighted_fulfillment', 0):.2f}%")
            print(f"Students with at least one wish: {overall_stats.get('students_with_at_least_one_wish_pct', 0):.2f}%")
            print(f"Students with top three wishes: {overall_stats.get('students_with_top_three_wishes_pct', 0):.2f}%")
            
        return df
        
    def get_student_schedule(self, student_id):
        if not self.core.schedule:
            return []
            
        schedule = []
        for (company_id, slot_idx), session in self.core.schedule.items():
            if slot_idx == -1:  # Skip excluded companies
                continue
                
            # Check if this student is in this session
            for student_info in session.students:
                if student_info["id"] == student_id:
                    # Get wish number if available
                    wish_number = student_info.get("wish_number", "-")
                    
                    # Get time slot info
                    slot_letter, time_range = self.time_slots[slot_idx]
                    
                    schedule.append((
                        slot_letter,
                        time_range,
                        session.company.name,
                        session.room,
                        wish_number
                    ))
                    break
                    
        # Sort by time slot
        return sorted(schedule, key=lambda x: x[0])
        
    def get_fulfillment_statistics(self):
        """Get statistics on wish fulfillment"""
        if not self.core.schedule:
            return {}
            
        return self.core.calculate_fulfillment()["overall_stats"]

    def export_student_schedules_pdf(self, filepath: str) -> bool:
        """Export student schedules as PDF"""
        self.clear_error()
        
        if not self.core.schedule:
            self.on_error("Bitte zuerst den Zeitplan generieren.")
            return False
        try:
            self.pdf_exporter.export_student_schedules(
                filepath=filepath,
                schedule=self.core.schedule,
                student_preferences=self.core.student_preferences, # FIXME
                time_slots=self.time_slots
            )
            return True
        except Exception as e:
            self.on_error(f"Fehler beim Exportieren: {str(e)}")
            return False
            
    def export_student_schedules_excel(self, filepath: str) -> bool:
        """Export student schedules as Excel"""
        self.clear_error()
        
        if not self.core.schedule:
            self.on_error("Bitte zuerst den Zeitplan generieren.")
            return False   
        try:
            self.excel_exporter.export_student_schedules(
                filepath=filepath,
                schedule=self.core.schedule,
                student_preferences=self.core.student_preferences, # FIXME
                time_slots=self.time_slots
            )
            return True
        except Exception as e:
            self.on_error(f"Fehler beim Exportieren: {str(e)}")
            return False

    def export_company_overview_pdf(self, filepath: str) -> bool:
        """Export company overview as PDF"""
        self.clear_error()
        
        if not self.core.schedule:
            self.on_error("Bitte zuerst den Zeitplan generieren.")
            return False
            
        try:
            # Use the PDF exporter to export company overview
            self.pdf_exporter.export_company_overview(
                filepath=filepath,
                schedule=self.core.schedule,
                time_slots=self.time_slots
            )
            return True
        except Exception as e:
            self.on_error(f"Fehler beim Exportieren: {str(e)}")
            return False
            
    def export_company_overview_excel(self, filepath: str) -> bool:
        """Export company overview as Excel"""
        self.clear_error()
        
        if not self.core.schedule:
            self.on_error("Bitte zuerst den Zeitplan generieren.")
            return False
        try:
            self.excel_exporter.export_company_overview(
                filepath=filepath,
                schedule=self.core.schedule,
                time_slots=self.time_slots
            )
            return True
        except Exception as e:
            self.on_error(f"Fehler beim Exportieren: {str(e)}")
            return False

    def export_attendance_lists_pdf(self, filepath: str, preview_mode=False) -> bool:
        """Export attendance lists to PDF"""
        self.clear_error()
        if not self.core.schedule:
            self.on_error("Bitte erst den Zeitplan generieren!")
            return False
            
        return self.attendance_exporter.export_attendance_lists(filepath, self.core.schedule, self.core.time_slots, preview_mode=preview_mode)
            
    def export_attendance_lists_excel(self, filepath: str, preview_mode=False) -> bool:
        """Export attendance lists to Excel"""
        self.clear_error()
        if not self.core.schedule:
            self.on_error("Bitte erst den Zeitplan generieren!")
            return False
            
        return self.excel_exporter.export_attendance_lists(filepath, self.core.schedule, self.core.time_slots, preview_mode=preview_mode)
        
    def export_room_list_excel(self, filepath: str) -> bool:
        """Export room list with capacities to Excel"""
        self.clear_error()
        if not self.core.rooms or not self.core.room_capacities:
            self.on_error("Bitte erst die Räume importieren!")
            return False
            
        return self.excel_exporter.export_room_list(filepath, self.core.rooms, self.core.room_capacities)

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
            company_name = session.get_company_display_name()
            room = session.room
            
            # Add this session to each assigned students schedule
            for student in session.students:
                student_name = student["name"]
                if student_name not in student_schedules:
                    student_schedules[student_name] = []
                
                wish_number = student.get("wish_number", "-")
                    
                student_schedules[student_name].append((slot_letter, time_range, company_name, room, wish_number))
                
        for student_name in student_schedules:
            student_schedules[student_name].sort()
            
        return student_schedules
        
    def get_company_overview(self):
        if not self.is_data_loaded() or not self.get_schedule():
            return {}
            
        slot_to_companies = {}
        
        # Fill in the slots
        for (company_id, slot_idx), session in self.get_schedule().items():
            if slot_idx == -1:  # Skip excluded companies
                continue
                
            slot_letter, _ = self.time_slots[slot_idx]
            
            if slot_letter not in slot_to_companies:
                slot_to_companies[slot_letter] = []
            
            display_name = session.get_company_display_name() 
                
            slot_to_companies[slot_letter].append((
                display_name,
                session.room,
                len(session.students)
            ))
                
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
        self, name, field="", capacity=0, max_sessions=0, earliest_slot=0, blocked_slots=None, fixed_room=None
    ):
        self.name = name
        self.field = field
        self.capacity = capacity
        self.max_sessions = max_sessions
        self.earliest_slot = earliest_slot
        self.blocked_slots = blocked_slots or []
        self.fixed_room = fixed_room
        # Unique id that combines name and field
        self.unique_id = f"{name}_{field}" if field else name
        self.always_show_field = False
    
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
        if self.company.field:
            return f"{self.company.name} ({self.company.field})"
        return self.company.name
