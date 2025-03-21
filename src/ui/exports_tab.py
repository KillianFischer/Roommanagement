import tkinter as tk
from tkinter import ttk, filedialog
import os


class ExportsTab:
    def __init__(self, parent, scheduler, on_mousewheel, app):
        self.parent = parent
        self.scheduler = scheduler
        self.app = app
        
        self.export_frame = ttk.Frame(parent)
        parent.add(self.export_frame, text="Exportieren")

        # nested notebook for export previews
        self.export_notebook = ttk.Notebook(self.export_frame)
        self.export_notebook.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        # Student Schedules tab
        self._setup_student_schedules_tab()
        
        # Attendance Lists tab
        self._setup_attendance_lists_tab()

        # export frame grid
        self.export_frame.columnconfigure(0, weight=1)
        self.export_frame.rowconfigure(0, weight=1)

    def _setup_student_schedules_tab(self):
        self.student_schedules_frame = ttk.Frame(self.export_notebook)
        self.export_notebook.add(self.student_schedules_frame, text="Schülerzeitpläne")

        ttk.Button(
            self.student_schedules_frame,
            text="Vorschau",
            command=self.preview_student_schedules,
        ).grid(row=0, column=0, pady=5, padx=5)
        ttk.Button(
            self.student_schedules_frame,
            text="Als PDF exportieren",
            command=self.export_student_schedules,
        ).grid(row=0, column=1, pady=5, padx=5)

        # canvas and scrollbar for the preview
        self.student_preview_canvas = tk.Canvas(self.student_schedules_frame)
        self.student_preview_scrollbar = ttk.Scrollbar(
            self.student_schedules_frame,
            orient="vertical",
            command=self.student_preview_canvas.yview,
        )
        self.student_preview_frame = ttk.Frame(self.student_preview_canvas)

        self.student_preview_canvas.configure(
            yscrollcommand=self.student_preview_scrollbar.set
        )

        # Bind mouse wheel for student preview
        self.student_preview_canvas.bind_all(
            "<MouseWheel>",
            lambda e: self.app._on_mousewheel(e, self.student_preview_canvas),
        )

        self.student_preview_canvas.grid(
            row=1, column=0, columnspan=2, sticky="nsew", padx=5, pady=5
        )
        self.student_preview_scrollbar.grid(row=1, column=2, sticky="ns")
        self.student_preview_canvas.create_window(
            (0, 0), window=self.student_preview_frame, anchor="nw"
        )

        self.student_schedules_frame.rowconfigure(1, weight=1)
        self.student_schedules_frame.columnconfigure(0, weight=1)
        self.student_schedules_frame.columnconfigure(1, weight=1)
        
    def _setup_attendance_lists_tab(self):
        self.attendance_lists_frame = ttk.Frame(self.export_notebook)
        self.export_notebook.add(self.attendance_lists_frame, text="Anwesenheitslisten")

        ttk.Button(
            self.attendance_lists_frame,
            text="Vorschau",
            command=self.preview_attendance_lists,
        ).grid(row=0, column=0, pady=5, padx=5)
        ttk.Button(
            self.attendance_lists_frame,
            text="Als PDF exportieren",
            command=self.export_attendance_lists,
        ).grid(row=0, column=1, pady=5, padx=5)

        # canvas and scrollbar for the preview
        self.attendance_preview_canvas = tk.Canvas(self.attendance_lists_frame)
        self.attendance_preview_scrollbar = ttk.Scrollbar(
            self.attendance_lists_frame,
            orient="vertical",
            command=self.attendance_preview_canvas.yview,
        )
        self.attendance_preview_frame = ttk.Frame(self.attendance_preview_canvas)

        self.attendance_preview_canvas.configure(
            yscrollcommand=self.attendance_preview_scrollbar.set
        )

        # Bind mouse wheel for attendance preview
        self.attendance_preview_canvas.bind_all(
            "<MouseWheel>",
            lambda e: self.app._on_mousewheel(e, self.attendance_preview_canvas),
        )

        self.attendance_preview_canvas.grid(
            row=1, column=0, columnspan=2, sticky="nsew", padx=5, pady=5
        )
        self.attendance_preview_scrollbar.grid(row=1, column=2, sticky="ns")
        self.attendance_preview_canvas.create_window(
            (0, 0), window=self.attendance_preview_frame, anchor="nw"
        )

        self.attendance_lists_frame.rowconfigure(1, weight=1)
        self.attendance_lists_frame.columnconfigure(0, weight=1)
        self.attendance_lists_frame.columnconfigure(1, weight=1)
        
    def export_student_schedules(self):
        self.app.clear_error()
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
            return
        
        filepath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if filepath:
            if self.scheduler.export_student_schedules_pdf(filepath):
                self.app.clear_error()

    def export_attendance_lists(self):
        self.app.clear_error()
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
            return
        
        filepath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if filepath:
            if self.scheduler.export_attendance_lists_pdf(filepath):
                self.app.clear_error()
                
    def preview_student_schedules(self):
        self.app.clear_error()
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
            return

        for widget in self.student_preview_frame.winfo_children():
            widget.destroy()

        # Display the overall erfüllungsscore at the top
        overall_score = self.scheduler.calculate_overall_fulfillment_score()
        ttk.Label(
            self.student_preview_frame,
            text=f"Gesamter Erfüllungsscore: {overall_score:.1f}%",
            font=("Helvetica", 12, "bold")
        ).grid(row=0, column=0, columnspan=4, pady=(5, 20), sticky="w")

        # Get student schedules by class
        class_schedules = {}
        
        # Get a lookup from company name to numeric ID to better handle wishes
        company_to_number = {}
        for idx, company in enumerate(self.scheduler.core.companies, 1):
            company_to_number[company.name.strip()] = str(idx)
        
        for student_name, appointments in self.scheduler.get_student_schedules().items():
            # Extract class name from student ID
            student = next((s for s in self.scheduler.student_preferences if s.name == student_name), None)
            if student:
                class_name = student.student_id.split("_")[0]
                if class_name not in class_schedules:
                    class_schedules[class_name] = []
                
                # Prepare schedule data
                schedule_data = []
                for slot_letter, time_range, company, room in appointments:
                    # Find which wish number this is
                    wish_number = None
                    
                    # Handle both direct company name matches and numeric matches
                    for i, wish in enumerate(student.wishes):
                        if not wish:
                            continue
                            
                        normalized_wish = str(wish).strip()
                        
                        # Check for direct company name match
                        if normalized_wish == company:
                            wish_number = i + 1
                            break
                            
                        # Check for company number match
                        try:
                            wish_num = int(float(normalized_wish))
                            company_for_number = None
                            for c in self.scheduler.core.companies:
                                if str(wish_num) == company_to_number.get(c.name.strip()):
                                    company_for_number = c.name.strip()
                                    break
                            
                            if company_for_number == company:
                                wish_number = i + 1
                                break
                        except (ValueError, TypeError):
                            pass
                    
                    if wish_number is None:
                        wish_number = "-"
                    
                    schedule_data.append({
                        "time": f"{slot_letter} ({time_range})",
                        "company": company,
                        "room": room,
                        "wish_number": wish_number
                    })
                
                class_schedules[class_name].append({
                    "name": student_name,
                    "schedule": schedule_data
                })

        row = 1  # Start from row 1 since row 0 is used for overall score

        for class_name, students in sorted(class_schedules.items()):
            ttk.Label(
                self.student_preview_frame,
                text=f"Klasse {class_name}",
                font=("Helvetica", 11, "bold")
            ).grid(row=row, column=0, columnspan=4, pady=(20, 10), sticky="w")
            row += 1

            for student in students:
                # Display student name
                ttk.Label(
                    self.student_preview_frame,
                    text=f"{student['name']}",
                    font=("Helvetica", 10, "bold")
                ).grid(row=row, column=0, columnspan=4, pady=(10, 5), sticky="w")
                row += 1

                for col, header in enumerate(["Zeit", "Unternehmen", "Raum", "Wunsch Nr."]):
                    ttk.Label(
                        self.student_preview_frame,
                        text=header,
                        font=("Helvetica", 9, "bold")
                    ).grid(row=row, column=col, padx=5, pady=2, sticky="w")
                row += 1

                for appointment in student["schedule"]:
                    ttk.Label(
                        self.student_preview_frame,
                        text=appointment["time"],
                    ).grid(row=row, column=0, padx=5, pady=2, sticky="w")
                    ttk.Label(
                        self.student_preview_frame,
                        text=appointment["company"],
                    ).grid(row=row, column=1, padx=5, pady=2, sticky="w")
                    ttk.Label(
                        self.student_preview_frame,
                        text=appointment["room"],
                    ).grid(row=row, column=2, padx=5, pady=2, sticky="w")
                    
                    # Show wish number with color coding
                    wish_label = ttk.Label(
                        self.student_preview_frame,
                        text=str(appointment["wish_number"]),
                    )
                    
                    # Color the wish numbers
                    if appointment["wish_number"] == 1:
                        wish_label.configure(foreground="green")
                    elif appointment["wish_number"] == 2:
                        wish_label.configure(foreground="darkgreen")
                    elif appointment["wish_number"] == 3:
                        wish_label.configure(foreground="forestgreen") 
                    elif appointment["wish_number"] in [4, 5, 6]:
                        wish_label.configure(foreground="orange")
                    
                    wish_label.grid(row=row, column=3, padx=5, pady=2, sticky="w")
                    row += 1

            self.student_preview_frame.update_idletasks()
            self.student_preview_canvas.configure(
                scrollregion=self.student_preview_canvas.bbox("all")
            )
            
    def preview_attendance_lists(self):
        self.app.clear_error()
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
            return

        # Clear previous preview
        for widget in self.attendance_preview_frame.winfo_children():
            widget.destroy()

        # Create a temporary PDF for preview
        temp_filepath = "temp_attendance_preview.pdf"
        if self.scheduler.export_attendance_lists_pdf(temp_filepath, preview_mode=True):
            # Show the preview directly in the UI
            sorted_sessions = sorted(
                self.scheduler.get_schedule().items(),
                key=lambda x: (x[0][0], x[0][1]),
            )

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

            row = 0

            for (company_name, slot_idx), session in sorted_sessions:
                # Skip excluded companies
                if slot_idx == -1:
                    continue
                    
                # Get time slot information
                slot_letter, time_range = self.scheduler.time_slots[slot_idx]
                
                # Company header
                ttk.Label(
                    self.attendance_preview_frame,
                    text=f"{company_name}",
                ).grid(row=row, column=0, columnspan=4, pady=(20, 5), sticky="w")
                row += 1

                # Time slot and room information
                ttk.Label(
                    self.attendance_preview_frame,
                    text=f"Zeitfenster: {slot_letter} ({time_range}) - Raum: {session.room}",
                ).grid(row=row, column=0, columnspan=4, pady=(0, 5), sticky="w")
                row += 1

                # Attendance list headers
                ttk.Label(
                    self.attendance_preview_frame,
                    text="Nr.",
                ).grid(row=row, column=0, padx=5, pady=2, sticky="w")
                ttk.Label(
                    self.attendance_preview_frame,
                    text="Name",
                ).grid(row=row, column=1, padx=5, pady=2, sticky="w")
                ttk.Label(
                    self.attendance_preview_frame,
                    text="Klasse",
                ).grid(row=row, column=2, padx=5, pady=2, sticky="w")
                ttk.Label(
                    self.attendance_preview_frame,
                    text="Anwesend",
                ).grid(row=row, column=3, padx=5, pady=2, sticky="w")
                row += 1

                # Check if this company has reached its minimum participants
                if session.company.min_participants > 0 and len(session.students) < session.company.min_participants:
                    # If minimum participants not reached, just show a message
                    ttk.Label(
                        self.attendance_preview_frame,
                        text="",
                    ).grid(row=row, column=0, padx=5, pady=2, sticky="w")
                    ttk.Label(
                        self.attendance_preview_frame,
                        text="Mindest Anzahl nicht erreicht",
                        font=("Helvetica", 10, "bold"),
                    ).grid(row=row, column=1, padx=5, pady=2, sticky="w")
                    ttk.Label(
                        self.attendance_preview_frame,
                        text="",
                    ).grid(row=row, column=2, padx=5, pady=2, sticky="w")
                    ttk.Label(
                        self.attendance_preview_frame,
                        text="",
                    ).grid(row=row, column=3, padx=5, pady=2, sticky="w")
                    row += 1
                else:
                    # Student rows - sort by name
                    for i, student in enumerate(sorted(session.students, key=lambda x: x["name"]), 1):
                        class_name = student["id"].split("_")[0]
                        ttk.Label(
                            self.attendance_preview_frame,
                            text=str(i),
                        ).grid(row=row, column=0, padx=5, pady=2, sticky="w")
                        ttk.Label(
                            self.attendance_preview_frame,
                            text=student["name"],
                        ).grid(row=row, column=1, padx=5, pady=2, sticky="w")
                        ttk.Label(
                            self.attendance_preview_frame,
                            text=class_name,
                        ).grid(row=row, column=2, padx=5, pady=2, sticky="w")
                        ttk.Label(
                            self.attendance_preview_frame,
                            text="________________",
                        ).grid(row=row, column=3, padx=5, pady=2, sticky="w")
                        row += 1

                    # Check if there are no students
                    current_students = len(session.students)
                    if current_students == 0:
                        # If no students, add a message row
                        ttk.Label(
                            self.attendance_preview_frame,
                            text="",
                        ).grid(row=row, column=0, padx=5, pady=2, sticky="w")
                        ttk.Label(
                            self.attendance_preview_frame,
                            text="Keine Teilnehmer",
                            font=("Helvetica", 10, "bold"),
                        ).grid(row=row, column=1, padx=5, pady=2, sticky="w")
                        ttk.Label(
                            self.attendance_preview_frame,
                            text="",
                        ).grid(row=row, column=2, padx=5, pady=2, sticky="w")
                        ttk.Label(
                            self.attendance_preview_frame,
                            text="",
                        ).grid(row=row, column=3, padx=5, pady=2, sticky="w")
                        row += 1

            # Update canvas scroll region
            self.attendance_preview_frame.update_idletasks()
            self.attendance_preview_canvas.configure(
                scrollregion=self.attendance_preview_canvas.bbox("all")
            )
            
            # Try to clean up the temporary file
            try:
                if os.path.exists(temp_filepath):
                    os.remove(temp_filepath)
            except:
                pass 