from typing import List, Dict, Optional, Tuple, Callable
import pandas as pd
from tkinter import messagebox

from models.student import StudentPreference
from models.company import Company, CompanySession


class SchedulerCore:
    def __init__(self, error_handler: Optional[Callable[[str], None]] = None):
        self.student_preferences: Optional[List[StudentPreference]] = None
        self.companies: Optional[List[Company]] = None
        self.rooms: Optional[List[str]] = None
        # Schedule: maps, company name, slot
        self.schedule: Dict[Tuple[str, int], CompanySession] = {}
        # Time slots
        self.time_slots = [
            ("A", "8:45 – 9:30"),
            ("B", "9:50 – 10:35"),
            ("C", "10:35 – 11:20"),
            ("D", "11:40 – 12:25"),
            ("E", "12:25 – 13:10"),
        ]
        self.error_handler = error_handler

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
            # Use the first column
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
            company_to_number, number_to_company = self._create_company_mappings()
            all_wish_counts, first_wish_counts = self._count_student_wishes(number_to_company)
            filtered_companies, excluded_companies = self._filter_companies_by_min_participants(all_wish_counts)
            sorted_companies = sorted(
                filtered_companies,
                key=lambda x: first_wish_counts.get(x.name.strip(), 0),
                reverse=True,
            )

            self.schedule.clear()
            company_rooms = {}
            available_rooms = self._prepare_available_rooms()

            self._handle_excluded_companies(excluded_companies)
            self._handle_special_companies(sorted_companies, company_rooms)
            
            sorted_companies = sorted(
                [c for c in self.companies if c.name.strip() != "Polizei"],
                key=lambda x: x.capacity,
                reverse=True,
            )

            self._assign_rooms_to_companies(sorted_companies, available_rooms, company_rooms, all_wish_counts)
            company_sessions = self._prepare_company_sessions()
            self._assign_students_to_sessions(company_sessions, number_to_company)
            self._calculate_student_fulfillment_scores(number_to_company)

            return True

        except Exception as e:
            error_message = f"Fehler bei der Zeitplangenerierung: {str(e)}"
            if self.error_handler:
                self.error_handler(error_message)
            else:
                messagebox.showerror("Error", error_message)
            self.schedule.clear()
            return False
            
    def _create_company_mappings(self):
        company_to_number = {}
        number_to_company = {}
        for idx, company in enumerate(self.companies, 1):
            normalized_name = company.name.strip()
            company_to_number[normalized_name] = str(idx)
            number_to_company[str(idx)] = normalized_name
            number_to_company[idx] = normalized_name
        return company_to_number, number_to_company
        
    def _count_student_wishes(self, number_to_company):
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
                
        return all_wish_counts, first_wish_counts
        
    def _filter_companies_by_min_participants(self, all_wish_counts):
        excluded_companies = []
        filtered_companies = []
        for company in self.companies:
            normalized_name = company.name.strip()
            wish_count = all_wish_counts.get(normalized_name, 0)
            if wish_count >= company.min_participants:
                filtered_companies.append(company)
            else:
                excluded_companies.append(company)
        return filtered_companies, excluded_companies
        
    def _prepare_available_rooms(self):
        available_rooms = self.rooms.copy()
        return [room for room in available_rooms if room.strip().lower() != "aula"]
        
    def _handle_excluded_companies(self, excluded_companies):
        for company in excluded_companies:
            session = CompanySession(
                company=company,
                room="Hat nicht die Min. Teilnehmer erreicht",
                time_slot="-",
                time_range="-",
            )
            self.schedule[(company.name, -1)] = session
            
    def _handle_special_companies(self, sorted_companies, company_rooms):
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
                    
    def _assign_rooms_to_companies(self, sorted_companies, available_rooms, company_rooms, all_wish_counts):
        for company in sorted_companies:
            if not available_rooms:
                available_rooms = self.rooms.copy()
                if "Aula" in available_rooms:
                    available_rooms.remove("Aula")
            
            company_room = available_rooms.pop(0)
            company_rooms[company.name] = company_room
            
            company_name = company.name.strip()
            total_interest = all_wish_counts.get(company_name, 0)
            
            if total_interest <= 20:
                needed_slots = 1
            else:
                needed_slots = (total_interest + 19) // 20
                needed_slots = min(needed_slots, len(self.time_slots) - company.earliest_slot)
            
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
                
    def _prepare_company_sessions(self):
        company_sessions = {}
        for (company_name, slot_idx), session in self.schedule.items():
            if company_name not in company_sessions:
                company_sessions[company_name] = []
            company_sessions[company_name].append((slot_idx, session))
        
        for company_name in company_sessions:
            company_sessions[company_name].sort(key=lambda x: x[0])
            
        return company_sessions
        
    def _assign_students_to_sessions(self, company_sessions, number_to_company):
        for student in self.student_preferences:
            assigned_slots = set()
            assigned_companies = set()
            
            # First pass - try to assign students based on their wishes
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
                
                if company_name in company_sessions:
                    sessions = company_sessions[company_name]
                    
                    available_sessions = []
                    for slot_idx, session in sessions:
                        if slot_idx not in assigned_slots:
                            available_sessions.append((slot_idx, session, len(session.students)))
                    
                    if available_sessions:
                        if len(sessions) > 1:
                            available_sessions.sort(key=lambda x: x[2])
                        
                        slot_idx, session, current_count = available_sessions[0]
                        
                        if current_count < session.company.capacity:
                            session.add_student(student.student_id, student.name)
                            assigned_slots.add(slot_idx)
                            assigned_companies.add(company_name)
            
            # Second pass - prioritize companies from wishes that couldn't be assigned in the first pass
            remaining_wishes = [wish for wish in student.wishes if wish and wish not in assigned_companies]
            for wish in remaining_wishes:
                if len(assigned_slots) >= len(self.time_slots):
                    break
                    
                try:
                    wish_num = int(float(str(wish).strip()))
                    company_name = number_to_company.get(wish_num, str(wish).strip())
                except (ValueError, TypeError):
                    company_name = str(wish).strip()
                
                if (company_name, -1) in self.schedule or company_name in assigned_companies:
                    continue
                
                if company_name in company_sessions:
                    # Get all slots for this company that are not yet assigned
                    available_slots = []
                    for slot_idx, session in company_sessions[company_name]:
                        if slot_idx not in assigned_slots and len(session.students) < session.company.capacity:
                            available_slots.append((slot_idx, session))
                    
                    if available_slots:
                        slot_idx, session = available_slots[0]
                        session.add_student(student.student_id, student.name)
                        assigned_slots.add(slot_idx)
                        assigned_companies.add(company_name)
            
            # Third pass - ensure every student has all 5 slots filled
            # Try to find available slots for companies with space
            if len(assigned_slots) < len(self.time_slots):
                # Create list of unassigned slots
                unassigned_slots = [i for i in range(len(self.time_slots)) if i not in assigned_slots]
                
                # For each unassigned slot, try to find a company with space
                for slot_idx in unassigned_slots:
                    found_session = False
                    
                    # First try companies that have sessions on other slots but not this one
                    # This ensures variety in companies assigned to each student
                    available_sessions = []
                    for company_name, sessions in company_sessions.items():
                        if company_name not in assigned_companies:
                            for s_idx, session in sessions:
                                if s_idx == slot_idx and len(session.students) < session.company.capacity:
                                    available_sessions.append((session, len(session.students), company_name))
                    
                    if available_sessions:
                        available_sessions.sort(key=lambda x: x[1])  # Sort by student count
                        session, _, company_name = available_sessions[0]
                        session.add_student(student.student_id, student.name)
                        assigned_slots.add(slot_idx)
                        assigned_companies.add(company_name)
                        found_session = True
                    
                    # If no new company found with space, try any company with space
                    if not found_session:
                        available_companies = []
                        for (c_name, c_slot), session in self.schedule.items():
                            if c_slot == slot_idx and c_slot != -1:
                                if len(session.students) < session.company.capacity:
                                    available_companies.append((session, len(session.students), c_name))
                        
                        # Sort companies by student count (least full first)
                        if available_companies:
                            available_companies.sort(key=lambda x: x[1])
                            session, _, company_name = available_companies[0]
                            session.add_student(student.student_id, student.name)
                            assigned_slots.add(slot_idx)
                            assigned_companies.add(company_name)
                            found_session = True
                    
                    # If no session was found with space, create a new one for an unallocated company if possible
                    if not found_session:
                        # Find companies that don't have a session in this slot
                        for company in self.companies:
                            # Skip if company is excluded or already has a session in this slot
                            if (company.name, -1) in self.schedule or (company.name, slot_idx) in self.schedule:
                                continue
                            
                            # Skip if this slot is before the company's earliest slot
                            if slot_idx < company.earliest_slot or slot_idx in company.blocked_slots:
                                continue
                            
                            # Find a room for this company
                            if self.rooms:
                                # Try to use company's existing room if it has one
                                existing_room = None
                                for (c_name, c_slot), c_session in self.schedule.items():
                                    if c_name == company.name:
                                        existing_room = c_session.room
                                        break
                                
                                room = existing_room or self.rooms[0]
                                
                                # Create a new session
                                slot_letter, time_range = self.time_slots[slot_idx]
                                session = CompanySession(
                                    company=company,
                                    room=room,
                                    time_slot=slot_letter,
                                    time_range=time_range,
                                )
                                session.add_student(student.student_id, student.name)
                                self.schedule[(company.name, slot_idx)] = session
                                
                                # Update company_sessions
                                if company.name not in company_sessions:
                                    company_sessions[company.name] = []
                                company_sessions[company.name].append((slot_idx, session))
                                company_sessions[company.name].sort(key=lambda x: x[0])
                                
                                assigned_slots.add(slot_idx)
                                assigned_companies.add(company.name)
                                break
                
    def _calculate_student_fulfillment_scores(self, number_to_company):
        for student in self.student_preferences:
            assigned_companies = []
            for (company_name, slot_idx), session in self.schedule.items():
                if any(s["id"] == student.student_id for s in session.students):
                    assigned_companies.append(company_name)
            
            total_score = 0.0
            for company_name in assigned_companies:
                try:
                    company_idx = next((i for i, wish in enumerate(student.wishes) if 
                                        wish and (
                                            wish == company_name or 
                                            (number_to_company.get(wish) == company_name) or
                                            (number_to_company.get(int(float(str(wish)))) == company_name if 
                                             wish.isdigit() or str(wish).replace('.', '', 1).isdigit() else False)
                                        )), None)
                    if company_idx is not None:
                        weights = [6, 5, 4, 3, 2, 1]
                        if company_idx < len(weights):
                            total_score += weights[company_idx]
                except (ValueError, TypeError, AttributeError):
                    continue
            
            max_possible_score = min(len(assigned_companies) * 6, sum([6, 5, 4, 3, 2, 1][:len(student.wishes)]))
            if max_possible_score > 0:
                student.fulfillment_score = (total_score / max_possible_score) * 100
            else:
                student.fulfillment_score = 0.0
                
    def calculate_overall_fulfillment_score(self):
        """
        Calculate the overall erfüllungsscore for all students combined.
        
        Returns:
            float: The overall fulfillment score as a percentage
        """
        if not self.student_preferences:
            return 0.0
            
        # Sum the individual scores
        total_score = 0.0
        student_count = 0
        
        for student in self.student_preferences:
            if student.fulfillment_score is not None:
                total_score += student.fulfillment_score
                student_count += 1
                
        # Calculate average score (overall erfüllungsscore)
        if student_count > 0:
            return total_score / student_count
        else:
            return 0.0

    def get_schedule(self) -> Dict[Tuple[str, int], CompanySession]:
        return self.schedule
        
    def _create_number_to_company_map(self):
        number_to_company = {}
        for idx, company in enumerate(self.companies, 1):
            normalized_name = company.name.strip()
            number_to_company[str(idx)] = normalized_name
            number_to_company[idx] = normalized_name
        return number_to_company 