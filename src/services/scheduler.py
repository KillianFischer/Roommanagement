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

    def load_student_preferences(self, df: pd.DataFrame) -> bool:
        return self.core.load_student_preferences(df)

    def load_companies(self, df: pd.DataFrame) -> bool:
        return self.core.load_companies(df)

    def load_rooms(self, df: pd.DataFrame) -> bool:
        return self.core.load_rooms(df)

    def is_data_loaded(self) -> bool:
        return self.core.is_data_loaded()

    def generate_schedule(self) -> bool:
        return self.core.generate_schedule()

    def get_schedule(self):
        return self.core.get_schedule()

    def calculate_overall_fulfillment_score(self) -> float:
        return self.core.calculate_overall_fulfillment_score()

    def export_student_schedules_pdf(self, filepath: str) -> bool:
        try:
            if not self.is_data_loaded() or not self.get_schedule():
                if self.error_handler:
                    self.error_handler("Bitte laden Sie alle Dateien und generieren Sie einen Zeitplan.")
                else:
                    messagebox.showerror(
                        "Error", "Bitte laden Sie alle Dateien und generieren Sie einen Zeitplan."
                    )
                return False

            self.pdf_exporter.export_student_schedules(
                filepath,
                self.get_schedule(),
                self.core.student_preferences,
                self.time_slots
            )
            return True
        except Exception as e:
            if self.error_handler:
                self.error_handler(f"Fehler beim Exportieren: {str(e)}")
            else:
                messagebox.showerror("Error", f"Fehler beim Exportieren: {str(e)}")
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
        try:
            if not self.is_data_loaded() or not self.get_schedule():
                if self.error_handler:
                    self.error_handler("Bitte laden Sie alle Dateien und generieren Sie einen Zeitplan.")
                else:
                    messagebox.showerror(
                        "Error", "Bitte laden Sie alle Dateien und generieren Sie einen Zeitplan."
                    )
                return False

            return self.attendance_exporter.export_attendance_lists(
                filepath,
                self.get_schedule(),
                self.time_slots,
                preview_mode
            )
        except Exception as e:
            if self.error_handler:
                self.error_handler(f"Fehler beim Exportieren: {str(e)}")
            else:
                messagebox.showerror("Error", f"Fehler beim Exportieren: {str(e)}")
            return False

    def get_student_schedules(self):
        """
        Get student schedules in a format suitable for UI display
        
        Returns:
            dict: Student name -> list of (slot_letter, time_range, company, room)
        """
        if not self.is_data_loaded() or not self.get_schedule():
            return {}
            
        student_schedules = {}
        
        # Initialize schedules for all students
        for student in self.core.student_preferences:
            student_schedules[student.name] = []
            
        # Fill in the schedules based on company sessions
        for (company_name, slot_idx), session in self.get_schedule().items():
            if slot_idx == -1:  # Skip excluded companies
                continue
                
            slot_letter, time_range = self.time_slots[slot_idx]
            
            # Add each student in this session to their schedule
            for student in session.students:
                student_name = student["name"]
                
                if student_name not in student_schedules:
                    student_schedules[student_name] = []
                    
                student_schedules[student_name].append((
                    slot_letter,
                    time_range,
                    company_name,
                    session.room
                ))
                
        # Sort each student's schedule by slot
        for student_name in student_schedules:
            student_schedules[student_name].sort(key=lambda x: x[0])
            
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
        for (company_name, slot_idx), session in self.get_schedule().items():
            if slot_idx == -1:  # Skip excluded companies
                continue
                
            slot_letter, _ = self.time_slots[slot_idx]
            
            if slot_letter not in slot_to_companies:
                slot_to_companies[slot_letter] = []
                
            slot_to_companies[slot_letter].append((
                company_name,
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
            list: List of excluded company names
        """
        if not self.is_data_loaded() or not self.get_schedule():
            return []
            
        excluded_companies = []
        
        for (company_name, slot_idx), _ in self.get_schedule().items():
            if slot_idx == -1:
                excluded_companies.append(company_name)
                
        return sorted(excluded_companies)
        
    @property
    def student_preferences(self) -> Optional[List[StudentPreference]]:
        return self.core.student_preferences
