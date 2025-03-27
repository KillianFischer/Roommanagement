import tkinter as tk
from tkinter import ttk, filedialog
import os
import tempfile


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

        # Button frame at the top
        button_frame = ttk.Frame(self.student_schedules_frame)
        button_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        
        ttk.Button(
            button_frame,
            text="Vorschau",
            command=self.preview_student_schedules,
        ).grid(row=0, column=0, pady=5, padx=5, sticky="w")
        
        ttk.Button(
            button_frame,
            text="Als PDF exportieren",
            command=self.export_student_schedules,
        ).grid(row=0, column=1, pady=5, padx=5, sticky="e")
        
        # Configure button frame
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)

        # Canvas and scrollbar for the preview - take full width
        self.student_preview_canvas = tk.Canvas(self.student_schedules_frame, width=800)
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

        # Position canvas and scrollbar - canvas takes full width
        self.student_preview_canvas.grid(
            row=1, column=0, sticky="nsew", padx=5, pady=5
        )
        self.student_preview_scrollbar.grid(row=1, column=1, sticky="ns", pady=5)
        
        # Create window inside canvas that fills the width
        self.student_preview_canvas.create_window(
            (0, 0), window=self.student_preview_frame, anchor="nw", width=self.student_preview_canvas.winfo_width()
        )

        # Configure weight for full expansion
        self.student_schedules_frame.rowconfigure(1, weight=1)
        self.student_schedules_frame.columnconfigure(0, weight=1)
        
        # Bind to configure event to adjust the window width when canvas changes size
        self.student_preview_canvas.bind('<Configure>', self._on_canvas_configure)
        
    def _on_canvas_configure(self, event):
        # Update the width of the window to match the canvas width
        width = event.width - 10  # A little less than full width to prevent horizontal scrollbar
        self.student_preview_canvas.itemconfigure(self.student_preview_canvas.find_all()[0], width=width)
        
    def _setup_attendance_lists_tab(self):
        self.attendance_lists_frame = ttk.Frame(self.export_notebook)
        self.export_notebook.add(self.attendance_lists_frame, text="Anwesenheitslisten")

        # Button frame at the top
        button_frame = ttk.Frame(self.attendance_lists_frame)
        button_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        
        ttk.Button(
            button_frame,
            text="Vorschau",
            command=self.preview_attendance_lists,
        ).grid(row=0, column=0, pady=5, padx=5, sticky="w")
        
        ttk.Button(
            button_frame,
            text="Als PDF exportieren",
            command=self.export_attendance_lists,
        ).grid(row=0, column=1, pady=5, padx=5, sticky="e")
        
        # Configure button frame
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)

        # Canvas and scrollbar for the preview - take full width
        self.attendance_preview_canvas = tk.Canvas(self.attendance_lists_frame, width=800)
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

        # Position canvas and scrollbar - canvas takes full width
        self.attendance_preview_canvas.grid(
            row=1, column=0, sticky="nsew", padx=5, pady=5
        )
        self.attendance_preview_scrollbar.grid(row=1, column=1, sticky="ns", pady=5)
        
        # Create window inside canvas that fills the width
        self.attendance_preview_canvas.create_window(
            (0, 0), window=self.attendance_preview_frame, anchor="nw", width=self.attendance_preview_canvas.winfo_width()
        )

        # Configure weight for full expansion
        self.attendance_lists_frame.rowconfigure(1, weight=1)
        self.attendance_lists_frame.columnconfigure(0, weight=1)
        
        # Bind to configure event to adjust the window width when canvas changes size
        self.attendance_preview_canvas.bind('<Configure>', self._on_attendance_canvas_configure)
        
    def _on_attendance_canvas_configure(self, event):
        # Update the width of the window to match the canvas width
        width = event.width - 10  # A little less than full width to prevent horizontal scrollbar
        if self.attendance_preview_canvas.find_all():  # Check if canvas has items
            self.attendance_preview_canvas.itemconfigure(self.attendance_preview_canvas.find_all()[0], width=width)
        
    def export_student_schedules(self):
        self.app.clear_error()
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
            return
        
        filepath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if filepath:  # Only proceed if the user didn't cancel the dialog
            if self.scheduler.export_student_schedules_pdf(filepath):
                self.app.clear_error()
        # If filepath is empty (user cancelled), do nothing

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

        # Configure the frame columns for appropriate widths
        self.student_preview_frame.columnconfigure(0, weight=1)   # Zeit column
        self.student_preview_frame.columnconfigure(1, weight=4)   # Unternehmen column (much wider)
        self.student_preview_frame.columnconfigure(2, weight=1)   # Raum column
        self.student_preview_frame.columnconfigure(3, weight=1)   # Wunsch column

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
                for slot_letter, time_range, company, room, wish_number in appointments:
                    # Use the wish_number provided directly from the tuple
                    # This is more accurate than trying to recalculate it
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

                # Create headers with appropriate widths
                headers = ["Zeit", "Unternehmen", "Raum", "Wunsch Nr."]
                sticky_values = ["w", "w", "w", "w"]
                
                for col, (header, sticky) in enumerate(zip(headers, sticky_values)):
                    ttk.Label(
                        self.student_preview_frame,
                        text=header,
                        font=("Helvetica", 9, "bold")
                    ).grid(row=row, column=col, padx=5, pady=2, sticky=sticky)
                row += 1

                for appointment in student["schedule"]:
                    ttk.Label(
                        self.student_preview_frame,
                        text=appointment["time"],
                    ).grid(row=row, column=0, padx=5, pady=2, sticky="w")
                    
                    # Company name with more space
                    ttk.Label(
                        self.student_preview_frame,
                        text=appointment["company"],
                        wraplength=400  # Allow wrapping for very long company names
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

        # Configure column weights for attendance list display
        self.attendance_preview_frame.columnconfigure(0, weight=1)    # Nr column
        self.attendance_preview_frame.columnconfigure(1, weight=6)    # Name column (wider)
        self.attendance_preview_frame.columnconfigure(2, weight=2)    # Klasse column
        self.attendance_preview_frame.columnconfigure(3, weight=2)    # Anwesend column

        # Create a temporary PDF for preview - use an actual temp filepath
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_file:
            temp_filepath = temp_file.name
            
        # Ensure we have a valid filepath for the preview
        if temp_filepath and self.scheduler.export_attendance_lists_pdf(temp_filepath, preview_mode=True):
            # Show the preview directly in the UI
            sorted_sessions = sorted(
                self.scheduler.get_schedule().items(),
                key=lambda x: (x[0][0], x[0][1]),
            )

            # Get unique company IDs (now includes specialization)
            company_ids = list(
                set(company_id for (company_id, _), _ in sorted_sessions)
            )
            if len(company_ids) > 6:
                company_ids = company_ids[:6]

            sorted_sessions = [
                (key, session)
                for (key, session) in sorted_sessions
                if key[0] in company_ids
            ]

            row = 0

            for (company_id, slot_idx), session in sorted_sessions:
                # Skip excluded companies
                if slot_idx == -1:
                    continue
                    
                # Get time slot information
                slot_letter, time_range = self.scheduler.time_slots[slot_idx]
                
                # Company header with field info if available
                company_name = session.get_company_display_name()
                
                # Company header - full width, larger font
                header_label = ttk.Label(
                    self.attendance_preview_frame,
                    text=f"{company_name}",
                    font=("Helvetica", 11, "bold")
                )
                header_label.grid(row=row, column=0, columnspan=4, pady=(20, 5), sticky="w")
                row += 1

                # Time slot and room information
                time_label = ttk.Label(
                    self.attendance_preview_frame,
                    text=f"Zeitfenster: {slot_letter} ({time_range}) - Raum: {session.room}",
                    font=("Helvetica", 10, "italic")
                )
                time_label.grid(row=row, column=0, columnspan=4, pady=(0, 10), sticky="w")
                row += 1

                # Attendance list headers - bold
                headers = [("Nr.", 0), ("Name", 1), ("Klasse", 2), ("Anwesend", 3)]
                for header_text, col in headers:
                    header = ttk.Label(
                        self.attendance_preview_frame,
                        text=header_text,
                        font=("Helvetica", 10, "bold")
                    )
                    header.grid(row=row, column=col, padx=5, pady=5, sticky="w")
                row += 1

                # Check if this company has reached its minimum participants
                if session.company.min_participants > 0 and len(session.students) < session.company.min_participants:
                    # If minimum participants not reached, just show a message
                    ttk.Label(
                        self.attendance_preview_frame,
                        text="",
                    ).grid(row=row, column=0, padx=5, pady=2, sticky="w")
                    
                    warning_label = ttk.Label(
                        self.attendance_preview_frame,
                        text="Mindest Anzahl nicht erreicht",
                        font=("Helvetica", 10, "bold")
                    )
                    # Use foreground color if possible (ttk needs style)
                    try:
                        warning_label.configure(foreground="red")
                    except:
                        pass
                        
                    warning_label.grid(row=row, column=1, columnspan=3, padx=5, pady=10, sticky="w")
                    row += 1
                elif len(session.students) == 0:
                    # No students assigned
                    ttk.Label(
                        self.attendance_preview_frame,
                        text="",
                    ).grid(row=row, column=0, padx=5, pady=2, sticky="w")
                    
                    empty_label = ttk.Label(
                        self.attendance_preview_frame,
                        text="Keine Teilnehmer",
                        font=("Helvetica", 10)
                    )
                    empty_label.grid(row=row, column=1, columnspan=3, padx=5, pady=10, sticky="w")
                    row += 1
                else:
                    # Student rows - sort by name
                    for i, student in enumerate(sorted(session.students, key=lambda x: x["name"]), 1):
                        class_name = student["id"].split("_")[0]
                        
                        # Number
                        ttk.Label(
                            self.attendance_preview_frame,
                            text=str(i),
                        ).grid(row=row, column=0, padx=5, pady=2, sticky="w")
                        
                        # Name - wider column
                        ttk.Label(
                            self.attendance_preview_frame,
                            text=student["name"],
                            wraplength=300  # Allow wrapping for very long names
                        ).grid(row=row, column=1, padx=5, pady=2, sticky="w")
                        
                        # Class
                        ttk.Label(
                            self.attendance_preview_frame,
                            text=class_name,
                        ).grid(row=row, column=2, padx=5, pady=2, sticky="w")
                        
                        # Attendance checkbox placeholder
                        ttk.Label(
                            self.attendance_preview_frame,
                            text="□",
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