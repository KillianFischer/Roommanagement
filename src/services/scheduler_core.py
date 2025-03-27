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
            self.schedule.clear()
            
            # Count student wishes to determine company popularity
            _, number_to_company = self._create_company_mappings()
            all_wish_counts, first_wish_counts = self._count_student_wishes(number_to_company)
            
            # Filter companies by minimum participants
            filtered_companies, excluded_companies = self._filter_companies_by_min_participants(all_wish_counts)
            
            # Sort companies by popularity (first wish count)
            sorted_companies = sorted(
                filtered_companies,
                key=lambda x: first_wish_counts.get(x.name.strip(), 0),
                reverse=True,
            )
            
            # First identify Polizei and assign to Aula for all time slots
            polizei_company = self._identify_polizei_company(sorted_companies)
            if polizei_company and polizei_company in sorted_companies:
                sorted_companies.remove(polizei_company)
            
            # Handle companies with duplicate names by setting flags
            self._mark_duplicate_companies(sorted_companies)
            
            # Initialize room availability tracking
            self._room_usage = {room: {t: None for t in range(len(self.time_slots))} for room in self.rooms}
            
            # Always reserve Aula for Polizei in all time slots - ONLY if Aula exists
            if "Aula" in self.rooms and polizei_company:
                # Mark Aula as reserved for all slots
                for slot_idx in range(len(self.time_slots)):
                    self._room_usage["Aula"][slot_idx] = polizei_company.unique_id  # Use unique ID
                
                # Create dedicated Aula sessions for Polizei
                for slot_idx, (slot_letter, time_range) in enumerate(self.time_slots):
                    session = CompanySession(
                        company=polizei_company,
                        room="Aula",  # Always Aula
                        time_slot=slot_letter,
                        time_range=time_range,
                    )
                    self.schedule[(polizei_company.unique_id, slot_idx)] = session
            
            # Assign rooms to the remaining companies
            self._assign_companies_to_rooms(sorted_companies, all_wish_counts)
            
            # Handle excluded companies
            self._handle_excluded_companies(excluded_companies)
            
            # Now assign students to the sessions
            company_sessions = self._prepare_company_sessions()
            number_to_company = self._create_number_to_company_map()
            self._assign_students_to_sessions(company_sessions, number_to_company)
            
            # Calculate fulfillment scores
            self._calculate_student_fulfillment_scores(number_to_company)
            
            # Filter out empty sessions
            self.schedule = {k: v for k, v in self.schedule.items() 
                             if len(v.students) > 0 or k[1] == -1}  # Keep excluded companies
            
            # Validate the schedule
            is_valid = self.debug_room_assignments()
            if not is_valid:
                print("WARNING: Schedule validation found issues! See debug output above.")
                             
            return True
        except Exception as e:
            import traceback
            traceback.print_exc()
            if self.error_handler:
                self.error_handler(f"Fehler bei der Zeitplangenerierung: {str(e)}")
            else:
                messagebox.showerror("Error", f"Fehler bei der Zeitplangenerierung: {str(e)}")
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
        
    def _identify_polizei_company(self, companies):
        """Find the Polizei company in the list of companies"""
        for company in companies:
            if "polizei" in company.name.strip().lower():
                return company
        return None
        
    def _mark_duplicate_companies(self, companies):
        """Mark companies with duplicate names to always show their fields"""
        company_name_count = {}
        for company in companies:
            company_name = company.name.strip()
            company_name_count[company_name] = company_name_count.get(company_name, 0) + 1
            
        for company in companies:
            company_name = company.name.strip()
            if company_name_count.get(company_name, 0) > 1 and hasattr(company, 'field') and company.field:
                company.always_show_field = True
                
    def _assign_companies_to_rooms(self, companies, wish_counts):
        """Assign companies to rooms and time slots without conflicts"""
        # For each company, determine how many time slots they need
        for company in companies:
            company_name = company.name.strip()
            total_interest = wish_counts.get(company_name, 0)
            
            # Calculate needed slots based on total interest and capacity
            if total_interest <= company.capacity:
                needed_slots = 1
            else:
                # Instead of a simple 20 students per slot, use the company's capacity
                needed_slots = (total_interest + (company.capacity - 1)) // company.capacity
                needed_slots = min(needed_slots, len(self.time_slots) - company.earliest_slot)
            
            print(f"Company {company_name}: total interest {total_interest}, capacity {company.capacity}, needed slots {needed_slots}")
            
            # Find a room available for consecutive time slots starting from earliest
            found_room = False
            
            # Try to find a room for this company for all its needed slots
            for room in self.rooms:
                # Skip Aula - it's reserved for Polizei
                if room == "Aula":
                    continue
                    
                can_use_room = True
                # Check if room is available for all needed consecutive slots
                for slot_offset in range(needed_slots):
                    slot_idx = company.earliest_slot + slot_offset
                    if slot_idx >= len(self.time_slots):
                        can_use_room = False
                        break
                        
                    if self._room_usage[room][slot_idx] is not None:
                        can_use_room = False
                        break
                
                if can_use_room:
                    # Assign company to this room for all needed slots
                    for slot_offset in range(needed_slots):
                        slot_idx = company.earliest_slot + slot_offset
                        self._room_usage[room][slot_idx] = company.unique_id  # Use unique ID instead of name
                        
                        slot_letter, time_range = self.time_slots[slot_idx]
                        session = CompanySession(
                            company=company,
                            room=room,
                            time_slot=slot_letter,
                            time_range=time_range,
                        )
                        self.schedule[(company.unique_id, slot_idx)] = session
                    
                    found_room = True
                    break
            
            # If no room was found for all consecutive slots, try to assign to different rooms
            if not found_room:
                for slot_offset in range(needed_slots):
                    slot_idx = company.earliest_slot + slot_offset
                    if slot_idx >= len(self.time_slots):
                        continue
                        
                    # Find any available room for this slot
                    assigned = False
                    for room in self.rooms:
                        if room == "Aula":  # Skip Aula
                            continue
                            
                        if self._room_usage[room][slot_idx] is None:
                            self._room_usage[room][slot_idx] = company.unique_id  # Use unique ID instead of name
                            
                            slot_letter, time_range = self.time_slots[slot_idx]
                            session = CompanySession(
                                company=company,
                                room=room,
                                time_slot=slot_letter,
                                time_range=time_range,
                            )
                            self.schedule[(company.unique_id, slot_idx)] = session
                            
                            assigned = True
                            break
                    
                    if not assigned:
                        # If we can't find a room for this slot, skip it
                        # This means the company doesn't get all its needed slots
                        continue
            
    def _prepare_company_sessions(self):
        company_sessions = {}
        for (company_id, slot_idx), session in self.schedule.items():
            if slot_idx == -1:  # Skip excluded companies
                continue
            if company_id not in company_sessions:
                company_sessions[company_id] = []
            company_sessions[company_id].append((slot_idx, session))
        
        for company_id in company_sessions:
            company_sessions[company_id].sort(key=lambda x: x[0])
            
        return company_sessions
        
    def _assign_students_to_sessions(self, company_sessions, number_to_company):
        """Assign students to company sessions based on their wishes with balanced distribution"""
        # Create a mapping from company name and various representations to unique_id for easier lookup
        company_id_map = {}
        for company in self.companies:
            name = company.name.strip()
            company_id_map[name] = company.unique_id
            company_id_map[str(company)] = company.unique_id
            company_id_map[company.unique_id] = company.unique_id
        
        # First, get all student wishes and organize them by company
        company_wish_lists = {}
        for student in self.student_preferences:
            for wish_idx, wish in enumerate(student.wishes):
                if not wish:
                    continue
                
                # Find the company_id this wish refers to
                company_id = None
                try:
                    wish_num = int(float(str(wish).strip()))
                    for company in self.companies:
                        if str(wish_num) == company.name.strip():
                            company_id = company.unique_id
                            break
                    if not company_id and wish_num in number_to_company:
                        name = number_to_company[wish_num]
                        company_id = company_id_map.get(name)
                except (ValueError, TypeError):
                    wish_str = str(wish).strip()
                    company_id = company_id_map.get(wish_str)
                
                if not company_id:
                    continue
                
                # Add this student to the company's wish list with priority
                if company_id not in company_wish_lists:
                    company_wish_lists[company_id] = []
                
                company_wish_lists[company_id].append({
                    'student': student,
                    'wish_number': wish_idx + 1,  # 1-indexed
                    'priority': wish_idx,  # Lower is higher priority
                })
        
        # Initialize student assignments tracking
        student_assignments = {student.student_id: set() for student in self.student_preferences}
        total_time_slots = len(self.time_slots)
        
        # First, sort companies by popularity (number of wishes)
        sorted_companies = sorted(
            company_wish_lists.keys(),
            key=lambda c_id: len(company_wish_lists.get(c_id, [])),
            reverse=True
        )
        
        # First pass: Assign students to their wishes where possible
        for company_id in sorted_companies:
            if company_id not in company_sessions:
                continue
            
            # Get all sessions for this company
            company_slots = sorted(company_sessions[company_id])
            if not company_slots:
                continue
            
            # Get the student wish list for this company
            wish_list = company_wish_lists[company_id]
            
            # Sort the wish list by priority (lower wish number = higher priority)
            wish_list.sort(key=lambda x: x['priority'])
            
            # Calculate how many students we can assign per session
            if len(company_slots) > 1:
                students_per_session = max(1, len(wish_list) // len(company_slots))
                print(f"Company {company_id}: {len(wish_list)} students, {len(company_slots)} slots, {students_per_session} per slot")
            else:
                students_per_session = len(wish_list)
            
            # First assign high-priority wishes
            for wish_data in wish_list:
                student = wish_data['student']
                wish_number = wish_data['wish_number']
                
                # If student already has all slots filled, skip
                if len(student_assignments[student.student_id]) >= total_time_slots:
                    continue
                
                # Find a session with available space where the student doesn't have a conflict
                assigned = False
                for slot_idx, session in company_slots:
                    # Skip if student already has an assignment in this slot
                    if slot_idx in student_assignments[student.student_id]:
                        continue
                    
                    # Skip if session is already at target capacity or full
                    if len(session.students) >= students_per_session and len(session.students) >= session.company.capacity:
                        continue
                    
                    # Assign student to this session
                    session.add_student(student.student_id, student.name)
                    session.students[-1]["wish_number"] = wish_number
                    
                    # Mark this slot as assigned for this student
                    student_assignments[student.student_id].add(slot_idx)
                    assigned = True
                    break
        
        # Second pass: Make sure each student has all 5 time slots filled
        for student in self.student_preferences:
            assigned_slots = student_assignments[student.student_id]
            
            # If student already has all slots, continue
            if len(assigned_slots) >= total_time_slots:
                continue
            
            # Find which slots need to be filled
            empty_slots = [slot_idx for slot_idx in range(total_time_slots) if slot_idx not in assigned_slots]
            
            # For each empty slot, find an available session
            for slot_idx in empty_slots:
                # Find companies with available space in this slot
                available_sessions = []
                for (c_id, s_idx), session in self.schedule.items():
                    if s_idx == slot_idx and s_idx != -1 and not session.is_full():
                        available_sessions.append((session, len(session.students)))
                
                if available_sessions:
                    # Pick the session with the fewest students
                    available_sessions.sort(key=lambda x: x[1])
                    session = available_sessions[0][0]
                    
                    # Add student to this session
                    session.add_student(student.student_id, student.name)
                    session.students[-1]["wish_number"] = "-"  # Not a wish
                    
                    # Mark slot as assigned
                    assigned_slots.add(slot_idx)
                else:
                    # If no space is available in existing sessions, create a new one if possible
                    # Find a company not already assigned to this student
                    assigned_companies = set()
                    for s_idx in assigned_slots:
                        for (c_id, company_slot_idx), s in self.schedule.items():
                            if company_slot_idx == s_idx and any(sd["id"] == student.student_id for sd in s.students):
                                assigned_companies.add(c_id)
                                break
                    
                    # Find available company not assigned to this student
                    for company in self.companies:
                        if company.unique_id in assigned_companies:
                            continue
                            
                        # Skip if this slot is before the company's earliest slot
                        if slot_idx < company.earliest_slot or slot_idx in company.blocked_slots:
                            continue
                            
                        # Find a room for this slot
                        for room in self.rooms:
                            if room == "Aula":  # Skip Aula
                                continue
                                
                            # Check if room is free for this slot
                            if self._room_usage[room][slot_idx] is None:
                                # Create a new session
                                self._room_usage[room][slot_idx] = company.unique_id
                                
                                slot_letter, time_range = self.time_slots[slot_idx]
                                session = CompanySession(
                                    company=company,
                                    room=room,
                                    time_slot=slot_letter,
                                    time_range=time_range,
                                )
                                
                                # Add student
                                session.add_student(student.student_id, student.name)
                                session.students[-1]["wish_number"] = "-"  # Not a wish
                                
                                # Add to schedule
                                self.schedule[(company.unique_id, slot_idx)] = session
                                
                                # Update company_sessions
                                if company.unique_id not in company_sessions:
                                    company_sessions[company.unique_id] = []
                                company_sessions[company.unique_id].append((slot_idx, session))
                                
                                # Mark slot as assigned
                                assigned_slots.add(slot_idx)
                                break
                        
                        # If we assigned a session, move to the next slot
                        if slot_idx in assigned_slots:
                            break
        
        # Third pass: Handle student fairness - try to spread out unassigned wishes
        # This ensures that one student doesn't get all their wishes while another gets none
        student_wish_counts = {}
        for student in self.student_preferences:
            # Count how many wishes were fulfilled
            wish_count = 0
            for (c_id, slot_idx), session in self.schedule.items():
                if slot_idx == -1:
                    continue
                for sd in session.students:
                    if sd["id"] == student.student_id and sd.get("wish_number", "-") != "-":
                        wish_count += 1
            student_wish_counts[student.student_id] = wish_count
        
        # Sort students by wish fulfillment (fewer fulfilled wishes first)
        sorted_students = sorted(
            self.student_preferences,
            key=lambda s: student_wish_counts.get(s.student_id, 0)
        )
        
        # Try to improve fulfillment for students with few wishes granted
        for student in sorted_students:
            # Only consider students with few wishes fulfilled
            if student_wish_counts.get(student.student_id, 0) >= 2:
                continue
                
            # Check which slots have non-wish assignments
            non_wish_slots = []
            for (c_id, slot_idx), session in self.schedule.items():
                if slot_idx == -1:
                    continue
                for sd in session.students:
                    if sd["id"] == student.student_id and sd.get("wish_number", "-") == "-":
                        non_wish_slots.append((slot_idx, c_id, session))
            
            # For each non-wish slot, try to swap with a wish if possible
            for slot_idx, current_c_id, current_session in non_wish_slots:
                # See if any of the student's wishes can be fulfilled
                for wish_idx, wish in enumerate(student.wishes):
                    if not wish:
                        continue
                        
                    company_id = None
                    try:
                        wish_num = int(float(str(wish).strip()))
                        for company in self.companies:
                            if str(wish_num) == company.name.strip():
                                company_id = company.unique_id
                                break
                        if not company_id and wish_num in number_to_company:
                            name = number_to_company[wish_num]
                            company_id = company_id_map.get(name)
                    except (ValueError, TypeError):
                        wish_str = str(wish).strip()
                        company_id = company_id_map.get(wish_str)
                    
                    if not company_id or company_id == current_c_id:
                        continue
                    
                    # Check if this company has any sessions in this time slot
                    for (c_id, s_idx), session in self.schedule.items():
                        if c_id == company_id and s_idx == slot_idx:
                            # Found a session for the wish company in this slot
                            if not session.is_full():
                                # Remove from current session
                                current_session.students = [sd for sd in current_session.students if sd["id"] != student.student_id]
                                
                                # Add to new session with wish number
                                session.add_student(student.student_id, student.name)
                                session.students[-1]["wish_number"] = wish_idx + 1
                                
                                # Update wish count
                                student_wish_counts[student.student_id] += 1
                                break
                    
                    # If wish was fulfilled, move to next slot
                    if student_wish_counts[student.student_id] > 0:
                        break
        
    def _calculate_student_fulfillment_scores(self, number_to_company):
        # Create a mapping from company name to unique_id for easier lookup
        company_id_map = {}
        for company in self.companies:
            company_id_map[company.name.strip()] = company.unique_id
            company_id_map[str(company)] = company.unique_id
            company_id_map[company.unique_id] = company.unique_id
        
        # First update wish numbers in sessions
        for (company_id, slot_idx), session in self.schedule.items():
            if slot_idx == -1:  # Skip excluded companies
                continue
            
            # For each assigned student, find out which wish was fulfilled
            for student_data in session.students:
                student_id = student_data["id"]
                student = next((s for s in self.student_preferences if s.student_id == student_id), None)
                
                if not student:
                    continue
                    
                # Find which wish number this was for the student
                for i, wish in enumerate(student.wishes):
                    if not wish:
                        continue
                    
                    # Find unique ID for this wish
                    wish_company_id = None
                    
                    try:
                        # Try numeric wish
                        wish_num = int(float(str(wish).strip()))
                        wish_company_id = number_to_company.get(wish_num)
                    except (ValueError, TypeError):
                        # Non-numeric wish, treat as company name or ID
                        wish_str = str(wish).strip()
                        wish_company_id = company_id_map.get(wish_str)
                    
                    # Check if this wish matches the session company
                    if wish_company_id == company_id:
                        # Found the matching wish
                        wish_number = i + 1  # Store wish as 1-indexed
                        student_data["wish_number"] = wish_number
                        break
        
        # Calculate fulfillment scores for students
        for student in self.student_preferences:
            # Find all companies assigned to this student
            assigned_company_ids = []
            for (company_id, slot_idx), session in self.schedule.items():
                if slot_idx == -1:  # Skip excluded companies
                    continue
                if any(s["id"] == student.student_id for s in session.students):
                    assigned_company_ids.append(company_id)
            
            total_score = 0.0
            for company_id in assigned_company_ids:
                # Find which wish this company was for the student
                for i, wish in enumerate(student.wishes):
                    if not wish:
                        continue
                    
                    # Get company ID from wish
                    wish_company_id = None
                    try:
                        wish_num = int(float(str(wish).strip()))
                        wish_company_id = number_to_company.get(wish_num)
                    except (ValueError, TypeError):
                        wish_str = str(wish).strip()
                        wish_company_id = company_id_map.get(wish_str)
                    
                    if wish_company_id == company_id:
                        # Found matching wish
                        weights = [6, 5, 4, 3, 2, 1]
                        if i < len(weights):
                            total_score += weights[i]
                        break
            
            max_possible_score = min(len(assigned_company_ids) * 6, sum([6, 5, 4, 3, 2, 1][:len(student.wishes)]))
            if max_possible_score > 0:
                student.fulfillment_score = (total_score / max_possible_score) * 100
            else:
                student.fulfillment_score = 0.0
                
    def calculate_overall_fulfillment_score(self) -> float:
        """Calculate the overall score for how well student wishes were fulfilled"""
        if not self.schedule:
            return 0
        
        total_students = len(self.student_preferences)
        if total_students == 0:
            return 0
        
        total_slots = len(self.time_slots)
        
        # Count how many students got their wishes
        wish_counts = {i: 0 for i in range(1, 7)}  # 1-6 wish numbers
        missing_wish_count = 0
        
        for (company_id, slot_idx), session in self.schedule.items():
            if slot_idx == -1:  # Skip excluded companies
                continue
            
            for student in session.students:
                wish_number = student.get("wish_number", None)
                if wish_number is not None and 1 <= wish_number <= 6:
                    wish_counts[wish_number] += 1
                else:
                    missing_wish_count += 1
                
        # Calculate a weighted score: 
        # 1st wish = 100%, 2nd = 80%, 3rd = 60%, 4th = 40%, 5th = 20%, 6th = 10%, none = 0%
        weights = {1: 1.0, 2: 0.8, 3: 0.6, 4: 0.4, 5: 0.2, 6: 0.1}
        
        max_possible_score = total_students * total_slots * 1.0  # If everyone gets 1st wish for all slots
        achieved_score = sum(wish_counts[i] * weights[i] for i in range(1, 7))
        
        if max_possible_score == 0:
            return 0
        
        return (achieved_score / max_possible_score) * 100

    def get_schedule(self) -> Dict[Tuple[str, int], CompanySession]:
        return self.schedule
        
    def _create_number_to_company_map(self):
        """Create a mapping from numeric identifiers to company unique IDs"""
        number_to_company = {}
        for idx, company in enumerate(self.companies, 1):
            number_to_company[str(idx)] = company.unique_id
            number_to_company[idx] = company.unique_id
        return number_to_company 

    def _handle_excluded_companies(self, excluded_companies):
        """Handle companies that don't meet minimum participant requirements"""
        for company in excluded_companies:
            session = CompanySession(
                company=company,
                room="Hat nicht die Min. Teilnehmer erreicht",
                time_slot="-",
                time_range="-",
            )
            self.schedule[(company.unique_id, -1)] = session 

    def debug_room_assignments(self):
        """Print out the room assignments to debug scheduling issues"""
        # Create a grid of time slots x rooms
        room_grid = {room: {i: "---" for i in range(len(self.time_slots))} for room in self.rooms}
        
        # Fill in the grid with company names
        for (company_id, slot_idx), session in self.schedule.items():
            if slot_idx == -1:  # Skip excluded companies
                continue
                
            room = session.room
            if room in room_grid and slot_idx in room_grid[room]:
                company_name = session.company.name
                # Truncate long names
                if len(company_name) > 20:
                    company_name = company_name[:17] + "..."
                room_grid[room][slot_idx] = company_name
                
        # Print the grid
        print("\nRoom Assignment Debug Grid:")
        print("-" * 80)
        header = "Room".ljust(15)
        for slot_idx, (slot_letter, time_range) in enumerate(self.time_slots):
            header += f"| {slot_letter} ({time_range}) ".ljust(20)
        print(header)
        print("-" * 80)
        
        for room in sorted(room_grid.keys()):
            row = str(room).ljust(15)
            for slot_idx in range(len(self.time_slots)):
                row += f"| {room_grid[room][slot_idx]} ".ljust(20)
            print(row)
        print("-" * 80)
            
        # Check for Polizei in Aula
        polizei_in_aula = True
        polizei_elsewhere = False
        aula_used_by_others = False
        
        for (company_id, slot_idx), session in self.schedule.items():
            if slot_idx == -1:
                continue
                
            is_polizei = "polizei" in session.company.name.lower()
            in_aula = session.room == "Aula"
            
            if is_polizei and not in_aula:
                polizei_in_aula = False
                polizei_elsewhere = True
                print(f"ERROR: Polizei found in room {session.room} at slot {slot_idx}")
                
            if in_aula and not is_polizei:
                aula_used_by_others = True
                print(f"ERROR: Aula used by {session.company.name} at slot {slot_idx}")
                
        # Check for room conflicts
        room_conflicts = False
        room_usage = {}
        for (company_id, slot_idx), session in self.schedule.items():
            if slot_idx == -1:
                continue
                
            room_slot = (session.room, slot_idx)
            if room_slot in room_usage:
                room_conflicts = True
                print(f"ERROR: Room conflict in {session.room} at slot {slot_idx} between {room_usage[room_slot]} and {session.company.name}")
            else:
                room_usage[room_slot] = session.company.name
                
        # Print summary
        print("\nSchedule Validation Summary:")
        print(f"Polizei only in Aula: {'✓' if polizei_in_aula and not polizei_elsewhere else '✗'}")
        print(f"Aula only used by Polizei: {'✓' if not aula_used_by_others else '✗'}")
        print(f"No room conflicts: {'✓' if not room_conflicts else '✗'}")
        print(f"Total scheduled sessions: {len([k for k in self.schedule.keys() if k[1] != -1])}")
        
        return polizei_in_aula and not polizei_elsewhere and not aula_used_by_others and not room_conflicts 