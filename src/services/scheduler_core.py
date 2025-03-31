from typing import List, Dict, Optional, Tuple, Callable
from tkinter import messagebox
import pandas as pd
import logging

from models.student import StudentPreference
from models.company import Company, CompanySession

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class SchedulerCore:
    def __init__(self):
        self.student_preferences: Optional[List[StudentPreference]] = None
        self.companies: Optional[List[Company]] = None
        self.rooms: Optional[List[str]] = None
        self.room_capacities: Dict[str, int] = {}
        self.schedule: Dict[Tuple[str, int], CompanySession] = {}
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
        
        if "Raum" in df.columns:
            room_col = "Raum"
            capacity_col = "Kapazität" if "Kapazität" in df.columns else None
        else:
            room_col = df.columns[0]
            capacity_col = df.columns[1] if len(df.columns) > 1 else None
        
        excluded_values = ["raum", "room", "räume", "rooms", ""]
        
        for _, row in df.iterrows():
            room_name = str(row[room_col]).strip()
            
            if not room_name or room_name.lower() in excluded_values:
                logger.info(f"Skipping room entry: '{room_name}' (likely a header or empty value)")
                continue
            
            self.rooms.append(room_name)
            
            if capacity_col and pd.notna(row[capacity_col]):
                try:
                    capacity = int(row[capacity_col])
                    self.room_capacities[room_name] = capacity
                except (ValueError, TypeError):
                    logger.warning(f"Invalid capacity value for room {room_name}: {row[capacity_col]}")
                    self.room_capacities[room_name] = 20  # Default capacity
            else:
                self.room_capacities[room_name] = 20  # Default capacity if not specified
        
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
            
            self._room_usage = {room: {t: None for t in range(len(self.time_slots))} for room in self.rooms}
            
            _, number_to_company = self._create_company_mappings()
            all_wish_counts, first_wish_counts = self._count_student_wishes(number_to_company)
            
            filtered_companies, excluded_companies = self._filter_companies_by_interest(all_wish_counts)
            
            sorted_companies = sorted(
                filtered_companies,
                key=lambda x: first_wish_counts.get(x.name.strip(), 0),
                reverse=True,
            )
            
            polizei_company = self._identify_polizei_company(sorted_companies)
            if polizei_company and polizei_company in sorted_companies:
                sorted_companies.remove(polizei_company)
            
            self._mark_duplicate_companies(sorted_companies)
            
            if "Aula" in self.rooms and polizei_company:
                for slot_idx in range(len(self.time_slots)):
                    self._room_usage["Aula"][slot_idx] = polizei_company.unique_id  # Use unique ID
                
                for slot_idx, (slot_letter, time_range) in enumerate(self.time_slots):
                    session = CompanySession(
                        company=polizei_company,
                        room="Aula",
                        time_slot=slot_letter,
                        time_range=time_range,
                    )
                    self.schedule[(polizei_company.unique_id, slot_idx)] = session
            
            self._assign_companies_to_rooms(sorted_companies, all_wish_counts)
            
            self._handle_excluded_companies(excluded_companies)
            
            company_sessions = self._prepare_company_sessions()
            number_to_company = self._create_number_to_company_map()
            self._assign_students_to_sessions(company_sessions, number_to_company)
            
            self._calculate_student_fulfillment_scores(number_to_company)
            
            self.schedule = {k: v for k, v in self.schedule.items() 
                             if len(v.students) > 0 or k[1] == -1}  # Keep excluded companies
            
            is_valid = self.debug_room_assignments()
            if not is_valid:
                logger.warning("Schedule validation found issues.")
                             
            return True
        except Exception as e:
            messagebox.showerror("Fehler bei Zeitplanerstellung", f"Interner Fehler bei der Zeitplangenerierung: {str(e)}")
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
        
    def _filter_companies_by_interest(self, all_wish_counts):
        excluded_companies = []
        filtered_companies = []
        for company in self.companies:
            normalized_name = company.name.strip()
            wish_count = all_wish_counts.get(normalized_name, 0)
            if wish_count > 0:
                filtered_companies.append(company)
            else:
                excluded_companies.append(company)
        return filtered_companies, excluded_companies
        
    def _identify_polizei_company(self, companies):
        for company in companies:
            if "polizei" in company.name.strip().lower():
                return company
        return None
        
    def _mark_duplicate_companies(self, companies):
        company_name_count = {}
        for company in companies:
            company_name = company.name.strip()
            company_name_count[company_name] = company_name_count.get(company_name, 0) + 1
            
        for company in companies:
            company_name = company.name.strip()
            if company_name_count.get(company_name, 0) > 1 and hasattr(company, 'field') and company.field:
                company.always_show_field = True
                
    def _assign_companies_to_rooms(self, companies, wish_counts):
        logger.info("Assigning companies to rooms")
        
        sorted_rooms = sorted(self.rooms, key=lambda r: self.room_capacities.get(r, 0), reverse=True)
        
        raum_entries = [r for r in sorted_rooms if r.lower() == "raum"]
        if raum_entries:
            logger.warning(f"Found {len(raum_entries)} 'Raum' entries in the room list. These may cause problems: {raum_entries}")
        
        filtered_rooms = [r for r in sorted_rooms if r.lower() != "raum"]
        if len(filtered_rooms) != len(sorted_rooms):
            logger.info(f"Filtered out {len(sorted_rooms) - len(filtered_rooms)} 'Raum' entries from room list")
            sorted_rooms = filtered_rooms
        
        company_session_count = {company.unique_id: 0 for company in companies}
        
        finanzamt_companies = [c for c in companies if "finanzamt" in c.name.lower()]
        if finanzamt_companies:
            logger.info(f"Found Finanzamt companies: {[c.name for c in finanzamt_companies]}")
            
            for finanzamt in finanzamt_companies:
                if not hasattr(finanzamt, 'fixed_room') or not finanzamt.fixed_room:
                    finanzamt_rooms = [r for r in sorted_rooms if any(term in r.lower() for term in ["finanz", "steuer", "amt"])]
                    if finanzamt_rooms:
                        finanzamt.fixed_room = finanzamt_rooms[0]
                        logger.info(f"Set fixed room for {finanzamt.name} to {finanzamt.fixed_room}")
        
        for company in companies:
            wish_count = wish_counts.get(company.name.strip(), 0)
            logger.info(f"Assigning {company.name} (popularity: {wish_count}, max_sessions: {company.max_sessions}) to rooms")
            
            if wish_count == 0:
                logger.debug(f"Skipping {company.name} - no student interest")
                continue
                
            expected_students = min(wish_count, company.capacity)
            
            suitable_rooms = [r for r in sorted_rooms if self.room_capacities.get(r, 0) >= expected_students]
            if not suitable_rooms:
                logger.warning(f"No room with sufficient capacity for {company.name} ({expected_students} students)")
                suitable_rooms = sorted_rooms
            
            logger.info(f"Suitable rooms for {company.name}: {suitable_rooms}")
            
            fixed_room = None
            
            if hasattr(company, 'fixed_room') and company.fixed_room:
                if company.fixed_room in sorted_rooms:
                    fixed_room = company.fixed_room
                    logger.info(f"Using company's fixed room: {fixed_room} for {company.name}")
                else:
                    logger.warning(f"Company {company.name} has fixed room {company.fixed_room} but it's not available")
            
            if not fixed_room and "finanzamt" in company.name.lower():
                logger.info(f"Special handling for Finanzamt company: {company.name}")
                
                finanzamt_rooms = [r for r in suitable_rooms if any(term in r.lower() for term in ["finanz", "steuer", "amt"])]
                logger.info(f"Potential Finanzamt rooms: {finanzamt_rooms}")
                
                if finanzamt_rooms:
                    fixed_room = finanzamt_rooms[0]
                    logger.info(f"Found dedicated Finanzamt room: {fixed_room}")
                    company.fixed_room = fixed_room
                else:
                    non_raum_rooms = [r for r in suitable_rooms if r.lower() != "raum"]
                    if non_raum_rooms:
                        fixed_room = non_raum_rooms[0]
                        logger.info(f"No dedicated Finanzamt room found, using: {fixed_room}")
                        company.fixed_room = fixed_room
                    else:
                        logger.warning(f"No suitable room found for Finanzamt, using any available room")
                        if suitable_rooms:
                            fixed_room = suitable_rooms[0]
                            company.fixed_room = fixed_room
            
            sessions_assigned = 0
            
            for slot_idx in range(company.earliest_slot, len(self.time_slots)):
                if sessions_assigned >= company.max_sessions:
                    logger.info(f"Reached max sessions ({company.max_sessions}) for {company.name}, not assigning more slots")
                    break
                
                if hasattr(company, 'blocked_slots') and slot_idx in company.blocked_slots:
                    logger.info(f"Skipping slot {slot_idx} for {company.name} as it's in blocked slots")
                    continue
                
                slot_letter, time_range = self.time_slots[slot_idx]
                
                if fixed_room and self._room_usage.get(fixed_room, {}).get(slot_idx) is None:
                    assigned_room = fixed_room
                    logger.info(f"Using fixed room {fixed_room} for {company.name} in slot {slot_letter}")
                else:
                    assigned_room = None
                    for room in suitable_rooms:
                        if room.lower() != "raum" and self._room_usage.get(room, {}).get(slot_idx) is None:
                            assigned_room = room
                            break
                
                if assigned_room:
                    if assigned_room not in self._room_usage:
                        self._room_usage[assigned_room] = {t: None for t in range(len(self.time_slots))}
                    self._room_usage[assigned_room][slot_idx] = company.unique_id
                    
                    session = CompanySession(
                        company=company,
                        room=assigned_room,
                        time_slot=slot_letter,
                        time_range=time_range,
                    )
                    self.schedule[(company.unique_id, slot_idx)] = session
                    
                    sessions_assigned += 1
                    company_session_count[company.unique_id] += 1
                    
                    if "finanzamt" in company.name.lower():
                        logger.info(f"*** FINANZAMT ASSIGNMENT: {company.name} assigned to room {assigned_room} for slot {slot_letter} ***")
                    else:
                        logger.info(f"Assigned {company.name} to room {assigned_room} for slot {slot_letter} ({sessions_assigned}/{company.max_sessions})")
                else:
                    logger.warning(f"Could not find available room for {company.name} in slot {slot_letter}")
                    
        for company in companies:
            logger.info(f"Assigned {company_session_count[company.unique_id]}/{company.max_sessions} sessions to {company.name}")
            
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
        company_id_map = {}
        for company in self.companies:
            name = company.name.strip()
            company_id_map[name] = company.unique_id
            company_id_map[str(company)] = company.unique_id
            company_id_map[company.unique_id] = company.unique_id
        
        company_wish_lists = {}
        student_wish_map = {}  # Map from student_id to their wishes as company_ids
        
        for student in self.student_preferences:
            student_wish_map[student.student_id] = []
            
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
                
                if not company_id:
                    continue
                
                if company_id not in company_wish_lists:
                    company_wish_lists[company_id] = []
                
                company_wish_lists[company_id].append({
                    'student': student,
                    'wish_number': wish_idx + 1,  # 1-indexed
                    'priority': wish_idx,  # Lower is higher priority
                })
                
                student_wish_map[student.student_id].append((company_id, wish_idx + 1))
        
        student_assignments = {student.student_id: set() for student in self.student_preferences}
        student_fulfilled_wishes = {student.student_id: [] for student in self.student_preferences}
        total_time_slots = len(self.time_slots)
        
        sorted_companies = sorted(
            company_wish_lists.keys(),
            key=lambda c_id: len(company_wish_lists.get(c_id, [])),
            reverse=True
        )
        
        first_slot_idx = 0
        logger.info(f"First making sure every student has an assignment for slot A (index {first_slot_idx})")
        
        first_slot_sessions = []
        for company_id, slots in company_sessions.items():
            for slot_idx, session in slots:
                if slot_idx == first_slot_idx:
                    first_slot_sessions.append((company_id, session))
        
        if first_slot_sessions:
            unassigned_students = [s for s in self.student_preferences]
            
            for student in list(unassigned_students):
                if student.student_id not in student_wish_map:
                    continue
                    
                assigned = False
                for company_id, wish_number in student_wish_map[student.student_id]:
                    for i, (session_company_id, session) in enumerate(first_slot_sessions):
                        if session_company_id == company_id and not session.is_full():
                            session.add_student(student.student_id, student.name)
                            session.students[-1]["wish_number"] = wish_number
                            
                            student_assignments[student.student_id].add(first_slot_idx)
                            student_fulfilled_wishes[student.student_id].append((company_id, wish_number))
                            assigned = True
                            
                            if session.is_full():
                                first_slot_sessions.pop(i)
                                
                            unassigned_students.remove(student)
                            break
                    
                    if assigned:
                        break
            
            while unassigned_students and first_slot_sessions:
                student = unassigned_students.pop(0)
                
                assigned = False
                for i, (company_id, session) in enumerate(first_slot_sessions):
                    if not session.is_full():
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
                        
                        session.add_student(student.student_id, student.name)
                        session.students[-1]["wish_number"] = wish_number if wish_number else "-"
                        
                        student_assignments[student.student_id].add(first_slot_idx)
                        if wish_number:
                            student_fulfilled_wishes[student.student_id].append((company_id, wish_number))
                        assigned = True
                        
                        if session.is_full():
                            first_slot_sessions.pop(i)
                            
                        break
                
                if not assigned:
                    logger.info(f"Need to create a new session for student {student.student_id} in first slot")
                    
                    if student.student_id in student_wish_map:
                        for company_id, wish_number in student_wish_map[student.student_id]:
                            company = None
                            for c in self.companies:
                                if c.unique_id == company_id:
                                    company = c
                                    break
                                    
                            if not company or company.earliest_slot > first_slot_idx:
                                continue  # Skip if company can't be scheduled in first slot
                                
                            for room in self.rooms:
                                if room == "Aula":  # Skip Aula
                                    continue
                                    
                                if self._room_usage[room][first_slot_idx] is None:
                                    self._room_usage[room][first_slot_idx] = company_id
                                    
                                    slot_letter, time_range = self.time_slots[first_slot_idx]
                                    new_session = CompanySession(
                                        company=company,
                                        room=room,
                                        time_slot=slot_letter,
                                        time_range=time_range,
                                    )
                                    
                                    new_session.add_student(student.student_id, student.name)
                                    new_session.students[-1]["wish_number"] = wish_number
                                    
                                    self.schedule[(company_id, first_slot_idx)] = new_session
                                    
                                    if company_id not in company_sessions:
                                        company_sessions[company_id] = []
                                    company_sessions[company_id].append((first_slot_idx, new_session))
                                    
                                    first_slot_sessions.append((company_id, new_session))
                                    
                                    student_assignments[student.student_id].add(first_slot_idx)
                                    student_fulfilled_wishes[student.student_id].append((company_id, wish_number))
                                    assigned = True
                                    break
                            
                            if assigned:
                                break
                    
                    if not assigned:
                        for company in self.companies:
                            if company.earliest_slot > first_slot_idx:
                                continue  # Company can't be scheduled in first slot
                                
                            for room in self.rooms:
                                if room == "Aula":  # Skip Aula
                                    continue
                                    
                                if self._room_usage[room][first_slot_idx] is None:
                                    self._room_usage[room][first_slot_idx] = company.unique_id
                                    
                                    slot_letter, time_range = self.time_slots[first_slot_idx]
                                    new_session = CompanySession(
                                        company=company,
                                        room=room,
                                        time_slot=slot_letter,
                                        time_range=time_range,
                                    )
                                    
                                    new_session.add_student(student.student_id, student.name)
                                    new_session.students[-1]["wish_number"] = "-"  # Not a wish
                                    
                                    self.schedule[(company.unique_id, first_slot_idx)] = new_session
                                    
                                    if company.unique_id not in company_sessions:
                                        company_sessions[company.unique_id] = []
                                    company_sessions[company.unique_id].append((first_slot_idx, new_session))
                                    
                                    first_slot_sessions.append((company.unique_id, new_session))
                                    
                                    student_assignments[student.student_id].add(first_slot_idx)
                                    assigned = True
                                    break
                            
                            if assigned:
                                break
        
        for student in self.student_preferences:
            if len(student_fulfilled_wishes[student.student_id]) >= 3:
                continue
                
            top_wishes = []
            if student.student_id in student_wish_map:
                for company_id, wish_number in student_wish_map[student.student_id]:
                    if wish_number <= 3:
                        if not any(w[0] == company_id for w in student_fulfilled_wishes[student.student_id]):
                            top_wishes.append((company_id, wish_number))
            
            top_wishes.sort(key=lambda x: x[1])
            
            for company_id, wish_number in top_wishes:
                if len(student_assignments[student.student_id]) >= total_time_slots:
                    break
                
                company_slots = []
                if company_id in company_sessions:
                    for slot_idx, session in company_sessions[company_id]:
                        if slot_idx in student_assignments[student.student_id]:
                            continue
                            
                        if session.is_full():
                            continue
                            
                        company_slots.append((slot_idx, session))
                
                if company_slots:
                    company_slots.sort(key=lambda x: len(x[1].students))
                    slot_idx, session = company_slots[0]
                    
                    session.add_student(student.student_id, student.name)
                    session.students[-1]["wish_number"] = wish_number
                    
                    student_assignments[student.student_id].add(slot_idx)
                    student_fulfilled_wishes[student.student_id].append((company_id, wish_number))
        
        for student in self.student_preferences:
            assigned_slots = student_assignments[student.student_id]
            
            if len(assigned_slots) >= total_time_slots:
                continue
            
            empty_slots = [slot_idx for slot_idx in range(total_time_slots) if slot_idx not in assigned_slots]
            
            for slot_idx in empty_slots:
                available_sessions = []
                for (c_id, s_idx), session in self.schedule.items():
                    if s_idx == slot_idx and s_idx != -1 and not session.is_full():
                        available_sessions.append((session, len(session.students)))
                
                if available_sessions:
                    available_sessions.sort(key=lambda x: x[1])
                    session = available_sessions[0][0]
                    
                    session.add_student(student.student_id, student.name)
                    session.students[-1]["wish_number"] = "-"  # Not a wish
                    
                    assigned_slots.add(slot_idx)
                else:
                    assigned_companies = set()
                    for s_idx in assigned_slots:
                        for (c_id, company_slot_idx), s in self.schedule.items():
                            if company_slot_idx == s_idx and any(sd["id"] == student.student_id for sd in s.students):
                                assigned_companies.add(c_id)
                                break
                    
                    for company in self.companies:
                        if company.unique_id in assigned_companies:
                            continue
                            
                        if slot_idx < company.earliest_slot or slot_idx in company.blocked_slots:
                            continue
                            
                        for room in self.rooms:
                            if room == "Aula":  # Skip Aula
                                continue
                                
                            if self._room_usage[room][slot_idx] is None:
                                self._room_usage[room][slot_idx] = company.unique_id
                                
                                slot_letter, time_range = self.time_slots[slot_idx]
                                session = CompanySession(
                                    company=company,
                                    room=room,
                                    time_slot=slot_letter,
                                    time_range=time_range,
                                )
                                
                                session.add_student(student.student_id, student.name)
                                session.students[-1]["wish_number"] = "-"  # Not a wish
                                
                                self.schedule[(company.unique_id, slot_idx)] = session
                                
                                if company.unique_id not in company_sessions:
                                    company_sessions[company.unique_id] = []
                                company_sessions[company.unique_id].append((slot_idx, session))
                                
                                assigned_slots.add(slot_idx)
                                break
                        
                        if slot_idx in assigned_slots:
                            break

    def _calculate_student_fulfillment_scores(self, number_to_company):
        """
        Calculate fulfillment scores for each student based on how well their wishes were met
        
        Weighting:
        - Wish 1 = 6 points
        - Wish 2 = 5 points 
        - Wish 3 = 4 points
        - Wish 4 = 3 points
        - Wish 5 = 2 points
        - Wish 6 = 1 point
        
        Maximum possible score per student is 21 points (if all wishes are fulfilled).
        Only students with at least one wish are counted in the denominator.
        """
        logger.info("Calculating student fulfillment scores")
        
        wish_weights = {1: 6, 2: 5, 3: 4, 4: 3, 5: 2, 6: 1}
        max_score_per_student = 21  # Maximum possible score per student (sum of all weights)
        
        company_id_to_name = {}
        for company in self.companies:
            company_id_to_name[company.unique_id] = company.name
        
        student_assignments = {}
        
        student_wish_fulfillment = {}
        
        for (company_id, slot_idx), session in self.schedule.items():
            if slot_idx == -1:  # Skip excluded sessions
                continue
                
            for student in session.students:
                student_id = student["id"]
                
                if student_id not in student_assignments:
                    student_assignments[student_id] = {}
                    student_wish_fulfillment[student_id] = {}
                
                student_assignments[student_id][slot_idx] = company_id
        
        students_with_wishes = 0
        
        total_score = 0
        student_scores = {}
        
        for student in self.student_preferences:
            student_id = student.student_id
            
            if not student.wishes or all(not wish for wish in student.wishes):
                student_scores[student_id] = 0
                continue
                
            students_with_wishes += 1
            
            if student_id not in student_assignments:
                student_scores[student_id] = 0
                continue
            
            fulfilled_wishes = set()
            
            student_score = 0
            
            for slot_idx, company_id in student_assignments[student_id].items():
                company_name = company_id_to_name.get(company_id, "")
                
                wish_fulfilled = False
                for wish_idx, wish in enumerate(student.wishes, 1):
                    if not wish:
                        continue
                        
                    try:
                        wish_num = int(float(str(wish).strip()))
                        wish_company = number_to_company.get(wish_num, str(wish).strip())
                    except (ValueError, TypeError):
                        wish_company = str(wish).strip()
                    
                    if wish_company.lower() == company_name.lower():
                        wish_fulfilled = True
                        wish_number = wish_idx
                        
                        if wish_number not in fulfilled_wishes:
                            fulfilled_wishes.add(wish_number)
                            student_score += wish_weights.get(wish_number, 0)
                            logger.debug(f"Student {student_id} got wish {wish_number} for {company_name}")
                        
                        student_wish_fulfillment[student_id][slot_idx] = wish_number
                        break
                
                if not wish_fulfilled:
                    logger.debug(f"Student {student_id} didn't have {company_name} in wishes for slot {slot_idx}")
                    student_wish_fulfillment[student_id][slot_idx] = None
            
            student_scores[student_id] = student_score
            total_score += student_score
        
        total_possible_score = students_with_wishes * max_score_per_student
        fulfillment_percentage = (total_score / total_possible_score * 100) if total_possible_score > 0 else 0
        logger.info(f"Overall fulfillment score: {fulfillment_percentage:.2f}% ({total_score}/{total_possible_score})")
        logger.info(f"Students with wishes: {students_with_wishes}")
        
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
        
        Each student can earn a maximum of 21 points (sum of all wish weights).
        Only students with at least one wish are included in the calculation.
        """
        if not self.schedule:
            return 0.0
            
        logger.info("Calculating overall fulfillment score")
        
        wish_weights = {1: 6, 2: 5, 3: 4, 4: 3, 5: 2, 6: 1}
        max_score_per_student = 21  # Maximum possible score per student
        
        company_name_to_id = {}
        for company in self.companies:
            name = company.name.strip()
            company_name_to_id[name] = company.unique_id
            # Also add numeric ID if the name is a number
            try:
                num_id = int(float(name))
                company_name_to_id[str(num_id)] = company.unique_id
            except (ValueError, TypeError):
                pass
        
        students_with_wishes = sum(1 for student in self.student_preferences 
                                 if student.wishes and any(wish for wish in student.wishes))
        
        total_score = 0
        wish_counts = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, None: 0}
        
        student_fulfilled_wishes = {}
        
        for student in self.student_preferences:
            student_id = student.student_id
            
            if not student.wishes or all(not wish for wish in student.wishes):
                continue
                
            student_fulfilled_wishes[student_id] = set()
            wishes_by_company = {}
            
            for wish_idx, wish in enumerate(student.wishes):
                if not wish:
                    continue
                    
                try:
                    wish_num = int(float(str(wish).strip()))
                    wish_str = str(wish_num)
                except (ValueError, TypeError):
                    wish_str = str(wish).strip()
                
                if wish_str in company_name_to_id:
                    wishes_by_company[company_name_to_id[wish_str]] = wish_idx + 1  # 1-indexed
            
            for (company_id, slot_idx), session in self.schedule.items():
                if slot_idx == -1:  # Skip excluded companies
                    continue
                    
                if not any(s["id"] == student_id for s in session.students):
                    continue
                    
                if company_id in wishes_by_company:
                    wish_number = wishes_by_company[company_id]
                    
                    if wish_number not in student_fulfilled_wishes[student_id]:
                        student_fulfilled_wishes[student_id].add(wish_number)
                        wish_counts[wish_number] = wish_counts.get(wish_number, 0) + 1
                        logger.debug(f"Student {student_id} got wish {wish_number} for company {session.company.name}")
            
            student_score = sum(wish_weights.get(wish, 0) for wish in student_fulfilled_wishes[student_id])
            total_score += student_score
        
        max_possible_score = students_with_wishes * max_score_per_student
        
        fulfillment_percentage = 0.0
        if max_possible_score > 0:
            fulfillment_percentage = (total_score / max_possible_score) * 100
            
        logger.info(f"Fulfillment score statistics:")
        logger.info(f"- Students with wishes: {students_with_wishes}")
        logger.info(f"- Students with at least one fulfilled wish: {len(student_fulfilled_wishes)}")
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
        """Handle companies that don't have any student interest"""
        for company in excluded_companies:
            session = CompanySession(
                company=company,
                room="Kein Schülerinteresse",  # Updated message
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
        
        for (company_id, slot_idx), session in self.schedule.items():
            if slot_idx == -1:  # Skip excluded companies
                continue
                
            room = session.room
            time_slot = slot_idx
            
            if room not in room_schedule:
                room_schedule[room] = {}
            
            if time_slot in room_schedule[room]:
                existing_company = room_schedule[room][time_slot]
                logger.error(f"CONFLICT: Room {room} double-booked for slot {time_slot}: {existing_company} and {company_id}")
                valid = False
            else:
                room_schedule[room][time_slot] = company_id
            
            if company_id not in company_schedule:
                company_schedule[company_id] = {}
                
            if time_slot in company_schedule[company_id]:
                existing_room = company_schedule[company_id][time_slot]
                logger.error(f"CONFLICT: Company {company_id} scheduled in multiple rooms for slot {time_slot}: {existing_room} and {room}")
                valid = False
            else:
                company_schedule[company_id][time_slot] = room
            
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
        
        for (company_id, slot_idx), session in self.schedule.items():
            if slot_idx == -1:  # Skip excluded companies
                continue
                
            room = session.room
            company = session.company
            student_count = len(session.students)
            
            room_capacity = self.room_capacities.get(room, 30)
            if student_count > room_capacity:
                logger.warning(f"Capacity exceeded: Room {room} (capacity {room_capacity}) has {student_count} students for {company.name}")
                valid = False
            
            if student_count > company.capacity:
                logger.warning(f"Company capacity exceeded: {company.name} (capacity {company.capacity}) has {student_count} students")
                valid = False
        
        logger.info("Room assignments summary:")
        for room, slots in room_schedule.items():
            room_capacity = self.room_capacities.get(room, "unknown")
            slot_info = ", ".join([f"{self.time_slots[slot][0]}: {company}" for slot, company in sorted(slots.items())])
            logger.info(f"Room {room} (capacity: {room_capacity}): {slot_info}")
            
        total_student_slots = sum(len(slots) for slots in student_schedule.values())
        total_company_slots = sum(len(slots) for slots in company_schedule.values())
        
        logger.info(f"Schedule statistics:")
        logger.info(f"- Total rooms used: {len(room_schedule)}")
        logger.info(f"- Total companies scheduled: {len(company_schedule)}")
        logger.info(f"- Total students scheduled: {len(student_schedule)}")
        logger.info(f"- Total student slots filled: {total_student_slots}")
        logger.info(f"- Total company slots: {total_company_slots}")
        logger.info(f"- Average students per session: {total_student_slots / total_company_slots if total_company_slots else 0:.2f}")
        
        if valid:
            logger.info("Schedule validation PASSED - No conflicts found")
        else:
            logger.error("Schedule validation FAILED - See errors above")
        
        return valid 

    def _get_company_id_from_wish(self, wish):
        """Helper method to get company ID from a student wish"""
        if not wish:
            return None
            
        try:
            wish_num = int(float(str(wish).strip()))
            for company in self.companies:
                if str(wish_num) == company.name.strip():
                    return company.unique_id
                    
            if hasattr(self, 'number_to_company') and wish_num in self.number_to_company:
                name = self.number_to_company[wish_num]
                for company in self.companies:
                    if name.strip() == company.name.strip():
                        return company.unique_id
        except (ValueError, TypeError):
            wish_str = str(wish).strip()
            for company in self.companies:
                if wish_str == company.name.strip():
                    return company.unique_id
                    
        return None 