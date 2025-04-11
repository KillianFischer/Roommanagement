from typing import List, Optional, Callable
import pandas as pd
from tkinter import messagebox

from services.scheduler_core import SchedulerCore
from services.excel_exporter import ExcelExporter
from services.attendance_exporter import AttendanceExporter
from models.student import StudentPreference
from models.company import Company


class Scheduler:
    def __init__(self):
        self.core = SchedulerCore()
        self.excel_exporter = ExcelExporter()
        self.attendance_exporter = AttendanceExporter()
        self.time_slots = self.core.time_slots

    # student preferences loading
    def load_student_preferences(self, df: pd.DataFrame) -> bool:
        return self.core.load_student_preferences(df)

    # company data loading
    def load_companies(self, df: pd.DataFrame) -> bool:
        required_columns = [
            "Unternehmen",
            "Max. Teilnehmer",
            "Max. Veranstaltungen",
            "Frühester Zeitpunkt",
        ]
        # Optional columns
        field_column = "Fachrichtung"
        blocked_slots_col = "Gesperrte Slots" # Define column name for blocked slots

        for col in required_columns:
            if col not in df.columns:
                messagebox.showerror("Fehler bei Import", f"Erforderliche Spalte fehlt in Unternehmensliste: {col}")
                return False

        try:
            companies = []
            for _, row in df.iterrows():
                name = str(row["Unternehmen"]).strip()
                field = str(row[field_column]).strip() if field_column in df.columns and pd.notna(row[field_column]) else ""
                max_participants = int(row["Max. Teilnehmer"])
                max_sessions = int(row["Max. Veranstaltungen"])

                earliest_slot = 0
                if pd.notna(row["Frühester Zeitpunkt"]):
                    slot_letter = str(row["Frühester Zeitpunkt"]).strip().upper()
                    if slot_letter in ["A", "B", "C", "D", "E"]:
                        earliest_slot = ord(slot_letter) - ord("A")

                # Process blocked slots
                blocked_slots_list = []
                if blocked_slots_col in df.columns and pd.notna(row[blocked_slots_col]):
                    slots_str = str(row[blocked_slots_col]).strip()
                    if slots_str: # Check if the string is not empty
                        slot_letters = [s.strip().upper() for s in slots_str.split(',') if s.strip()] # Split and clean
                        blocked_slots_list = [ord(letter) - ord('A') for letter in slot_letters if letter in ['A', 'B', 'C', 'D', 'E']] # Convert valid letters to indices

                companies.append(
                    Company(
                        name=name,
                        field=field,
                        capacity=max_participants,
                        max_sessions=max_sessions,
                        earliest_slot=earliest_slot,
                        blocked_slots=blocked_slots_list # Pass the list
                    )
                )

            self.core.companies = companies
            return True
        except Exception as e:
            # Add more specific error handling if needed
            import traceback
            traceback.print_exc() # Print stack trace for debugging
            messagebox.showerror("Fehler bei Import", f"Fehler beim Verarbeiten der Unternehmensliste: {str(e)}")
            return False

    # room data loading
    def load_rooms(self, df: pd.DataFrame) -> bool:
        return self.core.load_rooms(df)

    # data loaded check
    def is_data_loaded(self) -> bool:
        return self.core.is_data_loaded()

    # schedule generation
    def generate_schedule(self) -> bool:
        if not self.is_data_loaded():
            messagebox.showerror("Fehler bei Zeitplanerstellung", "Bitte laden Sie zuerst alle Daten (Schüler, Unternehmen, Räume).")
            return False
        
        try:
            success = self.core.generate_schedule()
            
            if not success:
                return False 
                
            return True
        
        except Exception as e:
            messagebox.showerror("Fehler bei Zeitplanerstellung", f"Unerwarteter Fehler bei der Generierung des Zeitplans: {str(e)}")
            return False

    # schedule getter
    def get_schedule(self):
        return self.core.schedule

    # overall fulfillment score calculation
    def calculate_overall_fulfillment_score(self) -> float:
        if not self.core.schedule:
            return 0
        return self.core.calculate_overall_fulfillment_score()

    # student fulfillment score dataframe
    def get_student_fulfillment_scores(self) -> pd.DataFrame:
        if not self.core.schedule:
            return pd.DataFrame()
            
        student_data = []
        
        company_id_to_num = {}
        for company in self.core.companies:
            try:
                num_id = int(float(company.name.strip()))
                company_id_to_num[company.unique_id] = num_id
            except (ValueError, TypeError):
                pass
        
        for student in self.core.student_preferences:
            if not student.wishes:
                continue
                
            student_schedule = self.get_student_schedule(student.student_id)
            
            wish_counts = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, None: 0}
            total_score = 0
            
            for slot_letter, time_range, company_name, room, wish_number in student_schedule:
                if wish_number and wish_number != "-":
                    wish_counts[wish_number] += 1
                    total_score += (7 - wish_number)
            
            max_possible_score = 21
            fulfillment_pct = (total_score / max_possible_score) * 100 if max_possible_score > 0 else 0
            
            record = {
                "Student ID": student.student_id,
                "Name": student.name,
                "Total Score": total_score,
                "Max Score": max_possible_score,
                "Fulfillment %": round(fulfillment_pct, 2),
                "1st Wishes": wish_counts.get(1, 0),
                "2nd Wishes": wish_counts.get(2, 0),
                "3rd Wishes": wish_counts.get(3, 0),
                "4th Wishes": wish_counts.get(4, 0),
                "5th Wishes": wish_counts.get(5, 0),
                "6th Wishes": wish_counts.get(6, 0),
                "No Match": len(student_schedule) - sum(w[1] for w in wish_counts.items() if w[0] is not None),
            }
            
            for slot_letter, time_range, company_name, room, wish_number in student_schedule:
                record[f"Slot {slot_letter}"] = company_name
                record[f"Wish {slot_letter}"] = wish_number
                
            student_data.append(record)
            
        df = pd.DataFrame(student_data)
        
        if len(df) > 0:
            # Optionally calculate overall stats here if needed
            pass
            
        return df
        
    # individual student schedule getter
    def get_student_schedule(self, student_id):
        if not self.core.schedule:
            return []
            
        schedule = []
        for (company_id, slot_idx), session in self.core.schedule.items():
            if slot_idx == -1:
                continue
                
            for student_info in session.students:
                if student_info["id"] == student_id:
                    wish_number = student_info.get("wish_number", "-")
                    
                    slot_letter, time_range = self.time_slots[slot_idx]
                    
                    schedule.append((
                        slot_letter,
                        time_range,
                        session.company.name,
                        session.room,
                        wish_number
                    ))
                    break
                    
        return sorted(schedule, key=lambda x: x[0])
        
    # fulfillment statistics getter
    def get_fulfillment_statistics(self):
        if not self.core.schedule:
            return {}
            
        score = self.core.calculate_overall_fulfillment_score()
        
        students_with_wishes = len([s for s in self.core.student_preferences if s.wishes and any(wish for wish in s.wishes)])
        students_with_top_three = len([s for s in self.core.student_preferences if s.wishes and any(wish for wish in s.wishes[:3])])
        
        students_with_at_least_one_wish_pct = (students_with_wishes / len(self.core.student_preferences) * 100) if self.core.student_preferences else 0
        students_with_top_three_wishes_pct = (students_with_top_three / len(self.core.student_preferences) * 100) if self.core.student_preferences else 0

        return {
            "fulfillment_percentage": score,
            "students_with_wishes": students_with_wishes, 
            "students_with_at_least_one_wish": students_with_wishes,
            "students_with_top_three_wishes": students_with_top_three,
            "students_with_at_least_one_wish_pct": round(students_with_at_least_one_wish_pct, 1),
            "students_with_top_three_wishes_pct": round(students_with_top_three_wishes_pct, 1),
        }

    # student schedules pdf export
    def export_student_schedules_pdf(self, filepath: str) -> bool:
        if not self.core.schedule:
            messagebox.showerror("Exportfehler", "Bitte zuerst den Zeitplan generieren.")
            return False
        try:
            messagebox.showwarning("Nicht Implementiert", "PDF Export für Schülerpläne nicht vollständig implementiert.")
            return False # Placeholder
        except AttributeError:
             messagebox.showerror("Exportfehler", "PDF Exporter nicht initialisiert.")
             return False
        except Exception as e:
            messagebox.showerror("Exportfehler", f"Fehler beim PDF-Export der Schülerzeitpläne: {str(e)}")
            return False
            
    # student schedules excel export
    def export_student_schedules_excel(self, filepath: str) -> bool:
        if not self.core.schedule:
            messagebox.showerror("Exportfehler", "Bitte zuerst den Zeitplan generieren.")
            return False   
        try:
            self.excel_exporter.export_student_schedules(
                filepath=filepath,
                schedule=self.core.schedule,
                student_preferences=self.core.student_preferences,
                time_slots=self.time_slots
            )
            return True
        except Exception as e:
            messagebox.showerror("Exportfehler", f"Fehler beim Excel-Export der Schülerzeitpläne: {str(e)}")
            return False

    # company overview pdf export
    def export_company_overview_pdf(self, filepath: str) -> bool:
        if not self.core.schedule:
            messagebox.showerror("Exportfehler", "Bitte zuerst den Zeitplan generieren.")
            return False
            
        try:
            # Assuming pdf_exporter exists and has the method
            # self.pdf_exporter.export_company_overview(...)
            messagebox.showwarning("Nicht Implementiert", "PDF Export für Unternehmensübersicht nicht vollständig implementiert.")
            return False # Placeholder
        except AttributeError:
             messagebox.showerror("Exportfehler", "PDF Exporter nicht initialisiert.")
             return False
        except Exception as e:
            messagebox.showerror("Exportfehler", f"Fehler beim PDF-Export der Unternehmensübersicht: {str(e)}")
            return False
            
    # company overview excel export
    def export_company_overview_excel(self, filepath: str) -> bool:
        if not self.core.schedule:
            messagebox.showerror("Exportfehler", "Bitte zuerst den Zeitplan generieren.")
            return False
        try:
            self.excel_exporter.export_company_overview(
                filepath=filepath,
                schedule=self.core.schedule,
                time_slots=self.time_slots
            )
            return True
        except Exception as e:
            messagebox.showerror("Exportfehler", f"Fehler beim Excel-Export der Unternehmensübersicht: {str(e)}")
            return False

    # attendance lists pdf export
    def export_attendance_lists_pdf(self, filepath: str, preview_mode=False) -> bool:
        if not self.core.schedule:
            messagebox.showerror("Exportfehler", "Bitte erst den Zeitplan generieren!")
            return False
            
        return self.attendance_exporter.export_attendance_lists(filepath, self.core.schedule, self.core.time_slots, preview_mode=preview_mode)
            
    # attendance lists excel export
    def export_attendance_lists_excel(self, filepath: str, preview_mode=False) -> bool:
        if not self.core.schedule:
            messagebox.showerror("Exportfehler", "Bitte erst den Zeitplan generieren!")
            return False
            
        return self.excel_exporter.export_attendance_lists(filepath, self.core.schedule, self.core.time_slots, preview_mode=preview_mode)
        
    # room list excel export
    def export_room_list_excel(self, filepath: str) -> bool:
        if not self.core.rooms or not self.core.room_capacities:
            messagebox.showerror("Exportfehler", "Bitte erst die Räume importieren!")
            return False
            
        return self.excel_exporter.export_room_list(filepath, self.core.rooms, self.core.room_capacities)

    # schedule excel export (main grid)
    def export_schedule_excel(self, filepath: str) -> bool:
        if not self.core.schedule:
            messagebox.showerror("Exportfehler", "Bitte zuerst den Zeitplan generieren.")
            return False
            
        try:
            success = self.excel_exporter.export_schedule(
                filepath=filepath,
                schedule=self.core.schedule,
                time_slots=self.time_slots,
                companies=self.core.companies
            )
            return success
        except Exception as e:
            messagebox.showerror("Exportfehler", f"Fehler beim Exportieren des Zeitplans nach Excel: {str(e)}")
            return False

    # student schedules dictionary getter
    def get_student_schedules(self):
        if not self.core.schedule:
            return {}
            
        student_schedules = {}
        
        for (company_id, slot_idx), session in self.core.schedule.items():
            if slot_idx == -1:
                continue
                
            slot_letter, time_range = self.time_slots[slot_idx]
            company_name = session.get_company_display_name()
            room = session.room
            
            for student in session.students:
                student_name = student["name"]
                if student_name not in student_schedules:
                    student_schedules[student_name] = []
                
                wish_number = student.get("wish_number", "-")
                    
                student_schedules[student_name].append((slot_letter, time_range, company_name, room, wish_number))
                
        for student_name in student_schedules:
            student_schedules[student_name].sort()
            
        return student_schedules
        
    # company overview dictionary getter
    def get_company_overview(self):
        if not self.is_data_loaded() or not self.get_schedule():
            return {}
            
        slot_to_companies = {}
        
        for (company_id, slot_idx), session in self.get_schedule().items():
            if slot_idx == -1:
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
        
    # excluded companies list getter
    def get_excluded_companies(self):
        if not self.is_data_loaded() or not self.get_schedule():
            return []
            
        excluded_companies = []
        
        for (company_id, slot_idx), session in self.get_schedule().items():
            if slot_idx == -1:
                excluded_companies.append(session.get_company_display_name())
                
        return sorted(excluded_companies)
        
    # student preferences property
    @property
    def student_preferences(self) -> Optional[List[StudentPreference]]:
        return self.core.student_preferences
