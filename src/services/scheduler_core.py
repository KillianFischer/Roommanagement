from typing import List, Dict, Optional, Tuple, Callable
import pandas as pd
from tkinter import messagebox
import logging

from models.student import StudentPreference
from models.company import Company, CompanySession

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class SchedulerCore:
    def __init__(self, error_handler: Optional[Callable[[str], None]] = None):
        self.student_preferences: Optional[List[StudentPreference]] = None
        self.companies: Optional[List[Company]] = None
        self.rooms: Optional[List[str]] = None
        self.room_capacities: Dict[str, int] = {}  # Store room capacities
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
        logger.info(f"Loaded {len(self.student_preferences)} student preferences")
        return True

    def load_companies(self, df: pd.DataFrame) -> bool:
        if df is None or df.empty:
            return False
        df.columns = df.columns.str.strip()
        self.companies = Company.from_dataframe(df)
        logger.info(f"Loaded {len(self.companies)} companies")
        return True

    def load_rooms(self, df: pd.DataFrame) -> bool:
        if df is None or df.empty:
            return False
        
        self.rooms = []
        self.room_capacities = {}
        
        # Check if "Raum" and "Kapazität" are in the columns
        if "Raum" in df.columns:
            room_col = "Raum"
            capacity_col = "Kapazität" if "Kapazität" in df.columns else None
        else:
            # Use the first column for room and second for capacity if available
            room_col = df.columns[0]
            capacity_col = df.columns[1] if len(df.columns) > 1 else None
        
        # List of values to exclude (headers, empty values, etc.)
        excluded_values = ["raum", "room", "räume", "rooms", ""]
        
        for _, row in df.iterrows():
            # Get the room name and clean it
            room_name = str(row[room_col]).strip()
            
            # Skip if the room name is empty or matches an excluded value
            if not room_name or room_name.lower() in excluded_values:
                logger.info(f"Skipping room entry: '{room_name}' (likely a header or empty value)")
                continue
            
            # Add the room to our list
            self.rooms.append(room_name)
            
            # Store capacity if available
            if capacity_col and pd.notna(row[capacity_col]):
                try:
                    capacity = int(row[capacity_col])
                    self.room_capacities[room_name] = capacity
                except (ValueError, TypeError):
                    logger.warning(f"Invalid capacity value for room {room_name}: {row[capacity_col]}")
                    self.room_capacities[room_name] = 30  # Default capacity
            else:
                self.room_capacities[room_name] = 30  # Default capacity if not specified
        
        logger.info(f"Loaded {len(self.rooms)} rooms with capacities: {self.room_capacities}")
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
            
            # Initialize room usage tracking
            self._room_usage = {room: {t: None for t in range(len(self.time_slots))} for room in self.rooms}
            
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
                logger.warning("WARNING: Schedule validation found issues! See debug output above.")
                             
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
        """Assign companies to rooms, considering room capacities"""
        logger.info("Assigning companies to rooms")
        
        # Sort rooms by capacity (larger rooms first)
        sorted_rooms = sorted(self.rooms, key=lambda r: self.room_capacities.get(r, 0), reverse=True)
        logger.debug(f"Rooms sorted by capacity: {sorted_rooms}")
        
        # Debug: List all room names to identify problematic entries
        logger.info(f"Available rooms: {sorted_rooms}")
        # Check for 'Raum' entries in the room list and warn about them
        raum_entries = [r for r in sorted_rooms if r.lower() == "raum"]
        if raum_entries:
            logger.warning(f"Found {len(raum_entries)} 'Raum' entries in the room list. These may cause problems: {raum_entries}")
        
        # Filter out any "Raum" entries that might have slipped through
        filtered_rooms = [r for r in sorted_rooms if r.lower() != "raum"]
        if len(filtered_rooms) != len(sorted_rooms):
            logger.info(f"Filtered out {len(sorted_rooms) - len(filtered_rooms)} 'Raum' entries from room list")
            sorted_rooms = filtered_rooms
        
        # Log which companies we're assigning
        logger.info(f"Companies to assign: {[company.name for company in companies]}")
        
        # Check for Finanzamt company
        finanzamt_companies = [c for c in companies if "finanzamt" in c.name.lower()]
        if finanzamt_companies:
            logger.info(f"Found Finanzamt companies: {[c.name for c in finanzamt_companies]}")
            
            # Set a fixed room for Finanzamt if we don't have one already
            for finanzamt in finanzamt_companies:
                if not hasattr(finanzamt, 'fixed_room') or not finanzamt.fixed_room:
                    # Look for a room with finanzamt in the name
                    finanzamt_rooms = [r for r in sorted_rooms if any(term in r.lower() for term in ["finanz", "steuer", "amt"])]
                    if finanzamt_rooms:
                        finanzamt.fixed_room = finanzamt_rooms[0]
                        logger.info(f"Set fixed room for {finanzamt.name} to {finanzamt.fixed_room}")
        
        # Assign popular companies to larger rooms
        for company in companies:
            wish_count = wish_counts.get(company.name.strip(), 0)
            logger.info(f"Assigning {company.name} (popularity: {wish_count}) to rooms")
            
            # Skip if no students are interested
            if wish_count == 0:
                logger.debug(f"Skipping {company.name} - no student interest")
                continue
                
            # Determine how many students we expect (up to company capacity)
            expected_students = min(wish_count, company.capacity)
            
            # Find suitable rooms - rooms with enough capacity
            suitable_rooms = [r for r in sorted_rooms if self.room_capacities.get(r, 0) >= expected_students]
            if not suitable_rooms:
                logger.warning(f"No room with sufficient capacity for {company.name} ({expected_students} students)")
                # Fall back to the largest available room
                suitable_rooms = sorted_rooms
            
            logger.info(f"Suitable rooms for {company.name}: {suitable_rooms}")
            
            # Check for a fixed room assignment
            fixed_room = None
            
            # First check if company has a fixed_room property
            if hasattr(company, 'fixed_room') and company.fixed_room:
                if company.fixed_room in sorted_rooms:
                    fixed_room = company.fixed_room
                    logger.info(f"Using company's fixed room: {fixed_room} for {company.name}")
                else:
                    logger.warning(f"Company {company.name} has fixed room {company.fixed_room} but it's not available")
            
            # Finanzamt special handling
            if not fixed_room and "finanzamt" in company.name.lower():
                logger.info(f"Special handling for Finanzamt company: {company.name}")
                
                # First try to find a room specifically for Finanzamt
                finanzamt_rooms = [r for r in suitable_rooms if any(term in r.lower() for term in ["finanz", "steuer", "amt"])]
                logger.info(f"Potential Finanzamt rooms: {finanzamt_rooms}")
                
                # If found, use it
                if finanzamt_rooms:
                    fixed_room = finanzamt_rooms[0]
                    logger.info(f"Found dedicated Finanzamt room: {fixed_room}")
                    # Store it for future use
                    company.fixed_room = fixed_room
                else:
                    # Otherwise, assign to one of the larger rooms that's not "Raum"
                    non_raum_rooms = [r for r in suitable_rooms if r.lower() != "raum"]
                    if non_raum_rooms:
                        fixed_room = non_raum_rooms[0]
                        logger.info(f"No dedicated Finanzamt room found, using: {fixed_room}")
                        # Store it for future use
                        company.fixed_room = fixed_room
                    else:
                        logger.warning(f"No suitable room found for Finanzamt, using any available room")
                        # Use any suitable room if we can't find a better one
                        if suitable_rooms:
                            fixed_room = suitable_rooms[0]
                            company.fixed_room = fixed_room
            
            # Assign to rooms and time slots
            for slot_idx in range(company.earliest_slot, len(self.time_slots)):
                # Skip this slot if it's in the company's blocked slots
                if hasattr(company, 'blocked_slots') and slot_idx in company.blocked_slots:
                    logger.info(f"Skipping slot {slot_idx} for {company.name} as it's in blocked slots")
                    continue
                
                slot_letter, time_range = self.time_slots[slot_idx]
                
                # Try to use fixed room first if available
                if fixed_room and self._room_usage.get(fixed_room, {}).get(slot_idx) is None:
                    assigned_room = fixed_room
                    logger.info(f"Using fixed room {fixed_room} for {company.name} in slot {slot_letter}")
                else:
                    # Try to find an available room for this slot
                    assigned_room = None
                    for room in suitable_rooms:
                        if room.lower() != "raum" and self._room_usage.get(room, {}).get(slot_idx) is None:
                            # Room is available for this slot
                            assigned_room = room
                            break
                
                if assigned_room:
                    # Mark room as used for this slot
                    if assigned_room not in self._room_usage:
                        self._room_usage[assigned_room] = {t: None for t in range(len(self.time_slots))}
                    self._room_usage[assigned_room][slot_idx] = company.unique_id
                    
                    # Create the session
                    session = CompanySession(
                        company=company,
                        room=assigned_room,
                        time_slot=slot_letter,
                        time_range=time_range,
                    )
                    self.schedule[(company.unique_id, slot_idx)] = session
                    
                    # Extra logging for Finanzamt
                    if "finanzamt" in company.name.lower():
                        logger.info(f"*** FINANZAMT ASSIGNMENT: {company.name} assigned to room {assigned_room} for slot {slot_letter} ***")
                    else:
                        logger.info(f"Assigned {company.name} to room {assigned_room} for slot {slot_letter}")
                else:
                    logger.warning(f"Could not find available room for {company.name} in slot {slot_letter}")
                    
        logger.info(f"Assigned {len(self.schedule)} sessions in total")
        
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
        student_wish_map = {}  # Map from student_id to their wishes as company_ids
        
        for student in self.student_preferences:
            student_wish_map[student.student_id] = []
            
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
                
                # Store in student wish map
                student_wish_map[student.student_id].append((company_id, wish_idx + 1))
        
        # Initialize student assignments tracking
        student_assignments = {student.student_id: set() for student in self.student_preferences}
        student_fulfilled_wishes = {student.student_id: [] for student in self.student_preferences}
        total_time_slots = len(self.time_slots)
        
        # First, sort companies by popularity (number of wishes)
        sorted_companies = sorted(
            company_wish_lists.keys(),
            key=lambda c_id: len(company_wish_lists.get(c_id, [])),
            reverse=True
        )
        
        # First pass: Ensure every student gets an assignment for the first time slot (A)
        # This ensures that student schedules always start with slot A
        first_slot_idx = 0
        logger.info(f"First making sure every student has an assignment for slot A (index {first_slot_idx})")
        
        # Get all sessions available for the first slot
        first_slot_sessions = []
        for company_id, slots in company_sessions.items():
            for slot_idx, session in slots:
                if slot_idx == first_slot_idx:
                    first_slot_sessions.append((company_id, session))
        
        if first_slot_sessions:
            # Assign each student to one session in the first slot
            unassigned_students = [s for s in self.student_preferences]
            
            # First try to assign students to their wish companies in slot A
            for student in list(unassigned_students):
                if student.student_id not in student_wish_map:
                    continue
                    
                # Check if any of student's wishes are available in slot A
                assigned = False
                for company_id, wish_number in student_wish_map[student.student_id]:
                    # Find a session with this company in slot A
                    for i, (session_company_id, session) in enumerate(first_slot_sessions):
                        if session_company_id == company_id and not session.is_full():
                            # Assign student
                            session.add_student(student.student_id, student.name)
                            session.students[-1]["wish_number"] = wish_number
                            
                            # Mark this slot as assigned
                            student_assignments[student.student_id].add(first_slot_idx)
                            student_fulfilled_wishes[student.student_id].append((company_id, wish_number))
                            assigned = True
                            
                            # If session is now full, remove it from available sessions
                            if session.is_full():
                                first_slot_sessions.pop(i)
                                
                            # Remove student from unassigned list
                            unassigned_students.remove(student)
                            break
                    
                    if assigned:
                        break
            
            # For remaining unassigned students, assign to any available session
            while unassigned_students and first_slot_sessions:
                # Get a student
                student = unassigned_students.pop(0)
                
                # Find a session with space
                assigned = False
                for i, (company_id, session) in enumerate(first_slot_sessions):
                    if not session.is_full():
                        # Check if this student has this company in their wishes
                        wish_number = None
                        for wish_idx, wish in enumerate(student.wishes):
                            if not wish:
                                continue
                                
                            try:
                                wish_num = int(float(str(wish).strip()))
                                wish_company = number_to_company.get(wish_num, str(wish).strip())
                                if company_id == company_id_map.get(wish_company):
                                    wish_number = wish_idx + 1
                                    break
                            except (ValueError, TypeError):
                                wish_str = str(wish).strip()
                                if company_id == company_id_map.get(wish_str):
                                    wish_number = wish_idx + 1
                                    break
                        
                        # Assign student
                        session.add_student(student.student_id, student.name)
                        session.students[-1]["wish_number"] = wish_number if wish_number else "-"
                        
                        # Mark this slot as assigned
                        student_assignments[student.student_id].add(first_slot_idx)
                        if wish_number:
                            student_fulfilled_wishes[student.student_id].append((company_id, wish_number))
                        assigned = True
                        
                        # If session is now full, remove it from available sessions
                        if session.is_full():
                            first_slot_sessions.pop(i)
                            
                        break
                
                # If student couldn't be assigned to any session, create a new one if possible
                if not assigned:
                    logger.info(f"Need to create a new session for student {student.student_id} in first slot")
                    
                    # First try to create session with one of student's wishes
                    if student.student_id in student_wish_map:
                        for company_id, wish_number in student_wish_map[student.student_id]:
                            # Find the company object
                            company = None
                            for c in self.companies:
                                if c.unique_id == company_id:
                                    company = c
                                    break
                                    
                            if not company or company.earliest_slot > first_slot_idx:
                                continue  # Skip if company can't be scheduled in first slot
                                
                            # Find an available room
                            for room in self.rooms:
                                if room == "Aula":  # Skip Aula
                                    continue
                                    
                                # Check if room is free for this slot
                                if self._room_usage[room][first_slot_idx] is None:
                                    # Create a new session
                                    self._room_usage[room][first_slot_idx] = company_id
                                    
                                    slot_letter, time_range = self.time_slots[first_slot_idx]
                                    new_session = CompanySession(
                                        company=company,
                                        room=room,
                                        time_slot=slot_letter,
                                        time_range=time_range,
                                    )
                                    
                                    # Add student
                                    new_session.add_student(student.student_id, student.name)
                                    new_session.students[-1]["wish_number"] = wish_number
                                    
                                    # Add to schedule
                                    self.schedule[(company_id, first_slot_idx)] = new_session
                                    
                                    # Update company_sessions
                                    if company_id not in company_sessions:
                                        company_sessions[company_id] = []
                                    company_sessions[company_id].append((first_slot_idx, new_session))
                                    
                                    # Add to first slot sessions
                                    first_slot_sessions.append((company_id, new_session))
                                    
                                    # Mark slot as assigned
                                    student_assignments[student.student_id].add(first_slot_idx)
                                    student_fulfilled_wishes[student.student_id].append((company_id, wish_number))
                                    assigned = True
                                    break
                            
                            # If we assigned a session, stop
                            if assigned:
                                break
                    
                    # If still not assigned, try with any available company
                    if not assigned:
                        # Find an available company and room
                        for company in self.companies:
                            if company.earliest_slot > first_slot_idx:
                                continue  # Company can't be scheduled in first slot
                                
                            # Find an available room
                            for room in self.rooms:
                                if room == "Aula":  # Skip Aula
                                    continue
                                    
                                # Check if room is free for this slot
                                if self._room_usage[room][first_slot_idx] is None:
                                    # Create a new session
                                    self._room_usage[room][first_slot_idx] = company.unique_id
                                    
                                    slot_letter, time_range = self.time_slots[first_slot_idx]
                                    new_session = CompanySession(
                                        company=company,
                                        room=room,
                                        time_slot=slot_letter,
                                        time_range=time_range,
                                    )
                                    
                                    # Add student
                                    new_session.add_student(student.student_id, student.name)
                                    new_session.students[-1]["wish_number"] = "-"  # Not a wish
                                    
                                    # Add to schedule
                                    self.schedule[(company.unique_id, first_slot_idx)] = new_session
                                    
                                    # Update company_sessions
                                    if company.unique_id not in company_sessions:
                                        company_sessions[company.unique_id] = []
                                    company_sessions[company.unique_id].append((first_slot_idx, new_session))
                                    
                                    # Add to first slot sessions
                                    first_slot_sessions.append((company.unique_id, new_session))
                                    
                                    # Mark slot as assigned
                                    student_assignments[student.student_id].add(first_slot_idx)
                                    assigned = True
                                    break
                            
                            # If we assigned a session, stop
                            if assigned:
                                break
        
        # Second pass: Focus on top wishes (1-3) for each student
        logger.info("Prioritizing top wishes (1-3) for each student")
        for student in self.student_preferences:
            # Skip if student already has 3 or more wishes fulfilled
            if len(student_fulfilled_wishes[student.student_id]) >= 3:
                continue
                
            # Get this student's top wishes that haven't been fulfilled yet
            top_wishes = []
            if student.student_id in student_wish_map:
                for company_id, wish_number in student_wish_map[student.student_id]:
                    # Only consider wishes 1-3
                    if wish_number <= 3:
                        # Check if this wish has been fulfilled
                        if not any(w[0] == company_id for w in student_fulfilled_wishes[student.student_id]):
                            top_wishes.append((company_id, wish_number))
            
            # Sort by wish number (lower is better)
            top_wishes.sort(key=lambda x: x[1])
            
            # Try to fulfill each top wish
            for company_id, wish_number in top_wishes:
                # If student already has all slots filled, skip
                if len(student_assignments[student.student_id]) >= total_time_slots:
                    break
                
                # Get available sessions for this company
                company_slots = []
                if company_id in company_sessions:
                    for slot_idx, session in company_sessions[company_id]:
                        # Skip if student already has an assignment in this slot
                        if slot_idx in student_assignments[student.student_id]:
                            continue
                            
                        # Skip if session is full
                        if session.is_full():
                            continue
                            
                        company_slots.append((slot_idx, session))
                
                # If there are available slots, assign student
                if company_slots:
                    # Choose the slot with the fewest students
                    company_slots.sort(key=lambda x: len(x[1].students))
                    slot_idx, session = company_slots[0]
                    
                    # Assign student
                    session.add_student(student.student_id, student.name)
                    session.students[-1]["wish_number"] = wish_number
                    
                    # Mark slot as assigned
                    student_assignments[student.student_id].add(slot_idx)
                    student_fulfilled_wishes[student.student_id].append((company_id, wish_number))
        
        # Third pass: Assign students to their remaining wishes where possible
        logger.info("Continuing with regular assignment for remaining slots and wishes")
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
                logger.info(f"Company {company_id}: {len(wish_list)} students, {len(company_slots)} slots, {students_per_session} per slot")
            else:
                students_per_session = len(wish_list)
            
            # Assign remaining wishes
            for wish_data in wish_list:
                student = wish_data['student']
                wish_number = wish_data['wish_number']
                
                # Skip if this wish has already been fulfilled
                if any(w[0] == company_id for w in student_fulfilled_wishes[student.student_id]):
                    continue
                
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
                    student_fulfilled_wishes[student.student_id].append((company_id, wish_number))
                    assigned = True
                    break
        
        # Fourth pass: Make sure each student has all time slots filled
        logger.info("Filling in remaining slots for all students")
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

    def _calculate_student_fulfillment_scores(self, number_to_company):
        """
        Calculate fulfillment scores for each student based on how well their wishes were met
        Uses the weighting from the example Excel sheet: 
        - Wish 1 = 6 points
        - Wish 2 = 5 points 
        - Wish 3 = 4 points
        - Wish 4 = 3 points
        - Wish 5 = 2 points
        - Wish 6 = 1 point
        """
        logger.info("Calculating student fulfillment scores")
        
        # Create wish weighting
        wish_weights = {1: 6, 2: 5, 3: 4, 4: 3, 5: 2, 6: 1}
        max_score_per_student = 20  # Maximum possible score per student per session
        
        # Create a mapping from company unique ID to company name
        company_id_to_name = {}
        for company in self.companies:
            company_id_to_name[company.unique_id] = company.name
        
        # Create a mapping to track what sessions each student has been assigned to
        student_assignments = {}
        
        # Create a mapping to track which wish corresponds to which session for each student
        student_wish_fulfillment = {}
        
        # First, build a dictionary of which company each student is assigned to for each slot
        for (company_id, slot_idx), session in self.schedule.items():
            if slot_idx == -1:  # Skip excluded sessions
                continue
                
            for student in session.students:
                student_id = student["id"]
                
                if student_id not in student_assignments:
                    student_assignments[student_id] = {}
                    student_wish_fulfillment[student_id] = {}
                
                # Map slot to company for this student
                student_assignments[student_id][slot_idx] = company_id
        
        # Now evaluate if students got their wishes
        total_score = 0
        total_possible_score = 0
        student_scores = {}
        
        for student in self.student_preferences:
            student_id = student.student_id
            if student_id not in student_assignments:
                student_scores[student_id] = 0
                continue
            
            # Calculate score for this student
            student_score = 0
            
            for slot_idx, company_id in student_assignments[student_id].items():
                company_name = company_id_to_name.get(company_id, "")
                
                # Check if this company was in student's wishes
                wish_fulfilled = False
                for wish_idx, wish in enumerate(student.wishes, 1):
                    if not wish:
                        continue
                        
                    try:
                        # Try to match by company number
                        wish_num = int(float(str(wish).strip()))
                        wish_company = number_to_company.get(wish_num, str(wish).strip())
                    except (ValueError, TypeError):
                        wish_company = str(wish).strip()
                    
                    if wish_company.lower() == company_name.lower():
                        wish_fulfilled = True
                        wish_number = wish_idx
                        student_wish_fulfillment[student_id][slot_idx] = wish_number
                        student_score += wish_weights.get(wish_number, 0)
                        logger.debug(f"Student {student_id} got wish {wish_number} for {company_name} in slot {slot_idx}")
                        break
                
                if not wish_fulfilled:
                    logger.debug(f"Student {student_id} didn't have {company_name} in wishes for slot {slot_idx}")
                    student_wish_fulfillment[student_id][slot_idx] = None
            
            # Store the score for this student
            student_scores[student_id] = student_score
            total_score += student_score
            
            # Calculate max possible score for this student (20 points per session)
            max_student_score = len(student_assignments[student_id]) * max_score_per_student
            total_possible_score += max_student_score
            
        # Calculate overall score percentage
        fulfillment_percentage = (total_score / total_possible_score * 100) if total_possible_score > 0 else 0
        logger.info(f"Overall fulfillment score: {fulfillment_percentage:.2f}% ({total_score}/{total_possible_score})")
        
        # Store wish fulfillment in session data
        for (company_id, slot_idx), session in self.schedule.items():
            if slot_idx == -1:  # Skip excluded sessions
                continue
                
            for i, student in enumerate(session.students):
                student_id = student["id"]
                wish_number = student_wish_fulfillment.get(student_id, {}).get(slot_idx)
                session.students[i]["wish_number"] = wish_number
        
        return fulfillment_percentage

    def calculate_overall_fulfillment_score(self) -> float:
        """
        Calculate the overall fulfillment score as a percentage of maximum possible score.
        
        The score is calculated based on:
        - Wish 1 = 6 points
        - Wish 2 = 5 points
        - Wish 3 = 4 points
        - Wish 4 = 3 points
        - Wish 5 = 2 points
        - Wish 6 = 1 point
        - No match = 0 points
        
        Each student can earn a maximum of 20 points per session.
        """
        if not self.schedule:
            return 0.0
            
        logger.info("Calculating overall fulfillment score")
        
        # Create wish weighting
        wish_weights = {1: 6, 2: 5, 3: 4, 4: 3, 5: 2, 6: 1}
        max_score_per_session = 20  # Maximum possible score per student per session
        
        # Initialize counters
        total_score = 0
        total_students = 0
        total_sessions = 0
        wish_counts = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, None: 0}
        
        # Count students in sessions
        for (company_id, slot_idx), session in self.schedule.items():
            if slot_idx == -1:  # Skip excluded companies
                continue
                
            for student in session.students:
                total_students += 1
                total_sessions += 1
                wish_number = student.get("wish_number")
                
                # Add to wish count
                wish_counts[wish_number] = wish_counts.get(wish_number, 0) + 1
                
                # Add to total score
                if wish_number is not None:
                    score = wish_weights.get(wish_number, 0)
                    total_score += score
        
        # Calculate maximum possible score
        max_possible_score = total_sessions * max_score_per_session
        
        # Calculate percentage
        fulfillment_percentage = 0.0
        if max_possible_score > 0:
            fulfillment_percentage = (total_score / max_possible_score) * 100
            
        # Log statistics
        logger.info(f"Fulfillment score statistics:")
        logger.info(f"- Total students in sessions: {total_students}")
        logger.info(f"- Total sessions: {total_sessions}")
        logger.info(f"- Total actual score: {total_score}")
        logger.info(f"- Maximum possible score: {max_possible_score}")
        logger.info(f"- Fulfillment percentage: {fulfillment_percentage:.2f}%")
        logger.info(f"- Wish distribution: {wish_counts}")
        
        return fulfillment_percentage

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

    def debug_room_assignments(self) -> bool:
        """Validate the generated schedule and check for conflicts"""
        logger.info("Validating the generated schedule")
        
        valid = True
        room_schedule = {}
        student_schedule = {}
        company_schedule = {}
        
        # Collect all assignments
        for (company_id, slot_idx), session in self.schedule.items():
            if slot_idx == -1:  # Skip excluded companies
                continue
                
            room = session.room
            time_slot = slot_idx
            
            # Check room conflicts
            if room not in room_schedule:
                room_schedule[room] = {}
            
            if time_slot in room_schedule[room]:
                existing_company = room_schedule[room][time_slot]
                logger.error(f"CONFLICT: Room {room} double-booked for slot {time_slot}: {existing_company} and {company_id}")
                valid = False
            else:
                room_schedule[room][time_slot] = company_id
            
            # Check company conflicts
            if company_id not in company_schedule:
                company_schedule[company_id] = {}
                
            if time_slot in company_schedule[company_id]:
                existing_room = company_schedule[company_id][time_slot]
                logger.error(f"CONFLICT: Company {company_id} scheduled in multiple rooms for slot {time_slot}: {existing_room} and {room}")
                valid = False
            else:
                company_schedule[company_id][time_slot] = room
            
            # Check student conflicts
            for student in session.students:
                student_id = student["id"]
                
                if student_id not in student_schedule:
                    student_schedule[student_id] = {}
                
                if time_slot in student_schedule[student_id]:
                    existing_company = student_schedule[student_id][time_slot]
                    logger.error(f"CONFLICT: Student {student_id} double-booked for slot {time_slot}: {existing_company} and {company_id}")
                    valid = False
                else:
                    student_schedule[student_id][time_slot] = company_id
        
        # Check room capacity constraints
        for (company_id, slot_idx), session in self.schedule.items():
            if slot_idx == -1:  # Skip excluded companies
                continue
                
            room = session.room
            company = session.company
            student_count = len(session.students)
            
            # Check room capacity limits
            room_capacity = self.room_capacities.get(room, 30)
            if student_count > room_capacity:
                logger.warning(f"Capacity exceeded: Room {room} (capacity {room_capacity}) has {student_count} students for {company.name}")
                valid = False
            
            # Check company capacity limits
            if student_count > company.capacity:
                logger.warning(f"Company capacity exceeded: {company.name} (capacity {company.capacity}) has {student_count} students")
                valid = False
        
        # Print room assignments summary
        logger.info("Room assignments summary:")
        for room, slots in room_schedule.items():
            room_capacity = self.room_capacities.get(room, "unknown")
            slot_info = ", ".join([f"{self.time_slots[slot][0]}: {company}" for slot, company in sorted(slots.items())])
            logger.info(f"Room {room} (capacity: {room_capacity}): {slot_info}")
            
        # Print summary statistics
        total_students = sum(len(session.students) for _, session in self.schedule.items() if _ and _[1] != -1)
        total_student_slots = sum(len(slots) for slots in student_schedule.values())
        total_company_slots = sum(len(slots) for slots in company_schedule.values())
        
        logger.info(f"Schedule statistics:")
        logger.info(f"- Total rooms used: {len(room_schedule)}")
        logger.info(f"- Total companies scheduled: {len(company_schedule)}")
        logger.info(f"- Total students scheduled: {len(student_schedule)}")
        logger.info(f"- Total student slots filled: {total_student_slots}")
        logger.info(f"- Total company slots: {total_company_slots}")
        logger.info(f"- Average students per session: {total_students / total_company_slots if total_company_slots else 0:.2f}")
        
        if valid:
            logger.info("Schedule validation PASSED - No conflicts found")
        else:
            logger.error("Schedule validation FAILED - See errors above")
        
        return valid 

    def calculate_fulfillment(self):
        """Calculate the fulfillment of student wishes"""
        total_students = len(self.student_preferences)
        if total_students == 0:
            return {}
        
        # Track wish fulfillment statistics
        wish_stats = {
            "wish1_fulfilled": 0,
            "wish2_fulfilled": 0,
            "wish3_fulfilled": 0,
            "wish4_fulfilled": 0,
            "wish5_fulfilled": 0,
            "no_wish_fulfilled": 0,
            "students_with_at_least_one_wish": 0,
            "students_with_top_three_wishes": 0,
            "students_with_all_five_sessions": 0,
            "total_wish_fulfillment": 0,
            "weighted_fulfillment": 0
        }
        
        # Track wish fulfillment per student
        fulfillment_by_student = {}
        
        # Weight for wishes (higher weight for top wishes)
        wish_weights = {1: 5, 2: 4, 3: 3, 4: 2, 5: 1}
        
        # For each student, check if their wishes were fulfilled
        for student in self.student_preferences:
            student_id = student.student_id
            fulfillment_by_student[student_id] = {
                "name": student.name,
                "wishes_fulfilled": [],
                "total_sessions": 0,
                "fulfillment_score": 0,
                "weighted_score": 0
            }
            
            wishes_fulfilled = []
            wishes_by_company = {}
            
            # Map each wish to a company ID
            for wish_idx, wish in enumerate(student.wishes):
                if not wish:
                    continue
                    
                company_id = self._get_company_id_from_wish(wish)
                if company_id:
                    wishes_by_company[company_id] = wish_idx + 1  # 1-indexed
            
            # Check each assigned session for this student
            for (company_id, slot_idx), session in self.schedule.items():
                for student_data in session.students:
                    if student_data["id"] == student_id:
                        fulfillment_by_student[student_id]["total_sessions"] += 1
                        
                        # Check if this assignment fulfills a wish
                        if company_id in wishes_by_company:
                            wish_number = wishes_by_company[company_id]
                            wishes_fulfilled.append((company_id, wish_number))
                            
                            # For debugging
                            student_data["wish_number"] = wish_number
                            
                            # Update statistics for this wish
                            wish_key = f"wish{wish_number}_fulfilled"
                            if wish_key in wish_stats:
                                wish_stats[wish_key] += 1
                                
                            # Calculate weighted score for this wish
                            if wish_number in wish_weights:
                                fulfillment_by_student[student_id]["weighted_score"] += wish_weights[wish_number]
            
            # Store wishes fulfilled for this student
            fulfillment_by_student[student_id]["wishes_fulfilled"] = sorted(wishes_fulfilled, key=lambda x: x[1])
            
            # Calculate fulfillment score (percentage of wishes fulfilled)
            fulfilled_count = len(wishes_fulfilled)
            if fulfilled_count > 0:
                wish_stats["students_with_at_least_one_wish"] += 1
                
                # Check if student has any of their top 3 wishes
                if any(wish[1] <= 3 for wish in wishes_fulfilled):
                    wish_stats["students_with_top_three_wishes"] += 1
                
                # Calculate as percentage of 5 total possible wishes
                fulfillment_by_student[student_id]["fulfillment_score"] = (fulfilled_count / 5) * 100
                wish_stats["total_wish_fulfillment"] += fulfilled_count
            else:
                wish_stats["no_wish_fulfilled"] += 1
                
            # Calculate weighted score as percentage of maximum possible weighted score
            max_weighted_score = sum(wish_weights.values())  # 15 for our weights
            weighted_pct = (fulfillment_by_student[student_id]["weighted_score"] / max_weighted_score) * 100
            fulfillment_by_student[student_id]["weighted_fulfillment"] = weighted_pct
            wish_stats["weighted_fulfillment"] += weighted_pct
            
            # Check if student has all 5 sessions
            if fulfillment_by_student[student_id]["total_sessions"] == 5:
                wish_stats["students_with_all_five_sessions"] += 1
        
        # Calculate overall statistics
        if total_students > 0:
            total_possible_wishes = total_students * 5
            wish_stats["fulfillment_percentage"] = (wish_stats["total_wish_fulfillment"] / total_possible_wishes) * 100
            wish_stats["students_with_at_least_one_wish_pct"] = (wish_stats["students_with_at_least_one_wish"] / total_students) * 100
            wish_stats["students_with_top_three_wishes_pct"] = (wish_stats["students_with_top_three_wishes"] / total_students) * 100
            wish_stats["students_with_all_five_sessions_pct"] = (wish_stats["students_with_all_five_sessions"] / total_students) * 100
            wish_stats["average_weighted_fulfillment"] = wish_stats["weighted_fulfillment"] / total_students
        
        # Return all statistics
        return {
            "overall_stats": wish_stats,
            "by_student": fulfillment_by_student
        }
        
    def _get_company_id_from_wish(self, wish):
        """Helper method to get company ID from a student wish"""
        if not wish:
            return None
            
        # Try to map wish to company ID
        try:
            # If wish is a number, it might be a company number
            wish_num = int(float(str(wish).strip()))
            # Look for company with this number as name
            for company in self.companies:
                if str(wish_num) == company.name.strip():
                    return company.unique_id
                    
            # Otherwise check if it's a company number in the mapping
            if hasattr(self, 'number_to_company') and wish_num in self.number_to_company:
                name = self.number_to_company[wish_num]
                for company in self.companies:
                    if name.strip() == company.name.strip():
                        return company.unique_id
        except (ValueError, TypeError):
            # If wish is not a number, it might be a company name
            wish_str = str(wish).strip()
            for company in self.companies:
                if wish_str == company.name.strip():
                    return company.unique_id
                    
        return None 