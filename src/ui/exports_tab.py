import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import tempfile
import pandas as pd


class ExportsTab:
    def __init__(self, parent, scheduler, on_mousewheel, app):
        self.parent = parent
        self.scheduler = scheduler
        self.app = app
        self._on_mousewheel = on_mousewheel
        
        self.export_frame = ttk.Frame(parent)
        parent.add(self.export_frame, text="Exportieren")

        # nested notebook for export previews
        self.export_notebook = ttk.Notebook(self.export_frame)
        self.export_notebook.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        # Student Schedules tab
        self._setup_student_schedules_tab()
        
        # Attendance Lists tab
        self._setup_attendance_lists_tab()
        
        # Fulfillment Report tab
        self._setup_fulfillment_tab()

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
            command=self.export_student_schedules_pdf,
        ).grid(row=0, column=1, pady=5, padx=5, sticky="e")
        
        ttk.Button(
            button_frame,
            text="Als Excel exportieren",
            command=self.export_student_schedules_excel,
        ).grid(row=0, column=2, pady=5, padx=5, sticky="e")
        
        # Configure button frame
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=0)
        button_frame.columnconfigure(2, weight=0)

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
            command=self.export_attendance_lists_pdf,
        ).grid(row=0, column=1, pady=5, padx=5, sticky="e")
        
        ttk.Button(
            button_frame,
            text="Als Excel exportieren",
            command=self.export_attendance_lists_excel,
        ).grid(row=0, column=2, pady=5, padx=5, sticky="e")
        
        # Configure button frame
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=0)
        button_frame.columnconfigure(2, weight=0)

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
        
        filetypes = [("PDF files", "*.pdf"), ("Excel files", "*.xlsx")]
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=filetypes
        )
        
        if filepath:  # Only proceed if the user didn't cancel the dialog
            # Determine export type based on file extension
            if filepath.lower().endswith('.pdf'):
                if self.scheduler.export_student_schedules_pdf(filepath):
                    self.app.clear_error()
            elif filepath.lower().endswith('.xlsx'):
                if self.scheduler.export_student_schedules_excel(filepath):
                    self.app.clear_error()
            else:
                # Default to Excel if extension is unclear
                if self.scheduler.export_student_schedules_excel(filepath + '.xlsx'):
                    self.app.clear_error()
        # If filepath is empty (user cancelled), do nothing

    def export_attendance_lists(self):
        self.app.clear_error()
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
            return
        
        filetypes = [("PDF files", "*.pdf"), ("Excel files", "*.xlsx")]
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=filetypes
        )
        
        if filepath:
            # Determine export type based on file extension
            if filepath.lower().endswith('.pdf'):
                if self.scheduler.export_attendance_lists_pdf(filepath, preview_mode=False):
                    self.app.clear_error()
            elif filepath.lower().endswith('.xlsx'):
                if self.scheduler.export_attendance_lists_excel(filepath, preview_mode=False):
                    self.app.clear_error()
            else:
                # Default to Excel if extension is unclear
                if self.scheduler.export_attendance_lists_excel(filepath + '.xlsx', preview_mode=False):
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

                # Check if the session has any students
                if len(session.students) == 0:
                    # If no students, show a message
                    ttk.Label(
                        self.attendance_preview_frame,
                        text="",
                    ).grid(row=row, column=0, padx=5, pady=2, sticky="w")
                    
                    warning_label = ttk.Label(
                        self.attendance_preview_frame,
                        text="Kein Schülerinteresse",
                        font=("Helvetica", 10, "bold")
                    )
                    # Use foreground color if possible (ttk needs style)
                    try:
                        warning_label.configure(foreground="red")
                    except:
                        pass
                        
                    warning_label.grid(row=row, column=1, columnspan=3, padx=5, pady=10, sticky="w")
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

    def _setup_fulfillment_tab(self):
        """Set up the tab for displaying fulfillment score reports"""
        self.fulfillment_frame = ttk.Frame(self.export_notebook)
        self.export_notebook.add(self.fulfillment_frame, text="Erfüllungsgrad")
        
        # Button frame at the top
        button_frame = ttk.Frame(self.fulfillment_frame)
        button_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        
        ttk.Button(
            button_frame,
            text="Aktualisieren",
            command=self._refresh_fulfillment_display,
        ).grid(row=0, column=0, pady=5, padx=5, sticky="w")
        
        ttk.Button(
            button_frame,
            text="Als Excel exportieren",
            command=self._export_fulfillment_excel,
        ).grid(row=0, column=1, pady=5, padx=5, sticky="e")
        
        ttk.Button(
            button_frame,
            text="Raumliste exportieren",
            command=self.export_room_list,
        ).grid(row=0, column=2, pady=5, padx=5, sticky="e")
        
        # Configure button frame
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=0)
        button_frame.columnconfigure(2, weight=0)
        
        # Informational text
        info_frame = ttk.Frame(self.fulfillment_frame, padding=10)
        info_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        
        ttk.Label(
            info_frame,
            text="Der Erfüllungsbericht zeigt, wie gut die Wünsche der Schüler erfüllt wurden.",
            font=("Helvetica", 11),
            wraplength=600,
        ).grid(row=0, column=0, sticky="w")
        
        ttk.Label(
            info_frame,
            text="Punkteverteilung für Wünsche:",
            font=("Helvetica", 11, "bold"),
        ).grid(row=1, column=0, sticky="w", pady=(10, 5))
        
        points_text = (
            "1. Wunsch: 5 Punkte\n"
            "2. Wunsch: 4 Punkte\n"
            "3. Wunsch: 3 Punkte\n"
            "4. Wunsch: 2 Punkte\n"
            "5. Wunsch: 1 Punkt\n"
            "Kein Wunsch: 0 Punkte"
        )
        
        ttk.Label(
            info_frame,
            text=points_text,
            font=("Helvetica", 11),
            justify="left",
        ).grid(row=2, column=0, sticky="w")
        
        ttk.Label(
            info_frame,
            text="Der Gesamterfüllungsgrad berechnet sich aus der Summe aller erzielten Punkte\n"
            "geteilt durch die maximal mögliche Punktzahl (15 pro Schüler).",
            font=("Helvetica", 11),
            justify="left",
        ).grid(row=3, column=0, sticky="w", pady=(10, 0))
        
        # Create a frame to contain all statistics
        self.stats_container = ttk.Frame(self.fulfillment_frame)
        self.stats_container.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        
        # Score display
        self.score_frame = ttk.LabelFrame(self.stats_container, text="Aktueller Erfüllungsgrad", padding=10)
        self.score_frame.pack(fill="x", pady=5)
        
        # Calculate and display the current score
        fulfillment_score = self.scheduler.calculate_overall_fulfillment_score()
        
        self.score_label = ttk.Label(
            self.score_frame,
            text=f"{fulfillment_score:.2f}%",
            font=("Helvetica", 16, "bold"),
        )
        self.score_label.pack(anchor="w")
        
        # Statistics frames - will be populated when refreshed
        self.wish_frame = ttk.LabelFrame(self.stats_container, text="Erfüllte Wünsche", padding=10)
        self.wish_frame.pack(fill="x", pady=5)
        
        self.student_frame = ttk.LabelFrame(self.stats_container, text="Schülerstatistik", padding=10)
        self.student_frame.pack(fill="x", pady=5)
        
        # Frame for student table
        self.table_frame = ttk.Frame(self.fulfillment_frame)
        self.table_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        
        # Create a scrollable frame for the student table
        self.student_table_canvas = tk.Canvas(self.table_frame)
        self.student_table_scrollbar = ttk.Scrollbar(
            self.table_frame, orient="vertical", command=self.student_table_canvas.yview
        )
        self.student_table = ttk.Frame(self.student_table_canvas)
        
        # Configure canvas
        self.student_table_canvas.configure(yscrollcommand=self.student_table_scrollbar.set)
        self.student_table_canvas.bind(
            "<Configure>",
            lambda e: self.student_table_canvas.configure(scrollregion=self.student_table_canvas.bbox("all"))
        )
        self.student_table_canvas.create_window((0, 0), window=self.student_table, anchor="nw")
        
        # Enable mousewheel scrolling
        self.student_table_canvas.bind_all(
            "<MouseWheel>", lambda event: self._on_mousewheel(event, self.student_table_canvas)
        )
        
        # Grid layout for canvas and scrollbar
        self.student_table_canvas.grid(row=0, column=0, sticky="nsew")
        self.student_table_scrollbar.grid(row=0, column=1, sticky="ns")
        self.table_frame.columnconfigure(0, weight=1)
        self.table_frame.rowconfigure(0, weight=1)
        
        # Configure weights for fulfillment frame
        self.fulfillment_frame.columnconfigure(0, weight=1)
        self.fulfillment_frame.rowconfigure(3, weight=1)  # Make the table expandable
        
        # Initialize the display
        self._refresh_fulfillment_display()

    def _refresh_fulfillment_display(self):
        """Refresh the fulfillment statistics display"""
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
            return
            
        # Update the score
        fulfillment_score = self.scheduler.calculate_overall_fulfillment_score()
        self.score_label.config(text=f"{fulfillment_score:.2f}%")
        
        # Get statistics
        stats = self.scheduler.get_fulfillment_statistics()
        student_df = self.scheduler.get_student_fulfillment_scores()
        
        # Clear existing widgets in statistics frames
        for widget in self.wish_frame.winfo_children():
            widget.destroy()
        
        for widget in self.student_frame.winfo_children():
            widget.destroy()
            
        for widget in self.student_table.winfo_children():
            widget.destroy()
        
        # Populate wish statistics
        wish_counts = [
            ("1. Wunsch", stats.get("wish1_fulfilled", 0)),
            ("2. Wunsch", stats.get("wish2_fulfilled", 0)),
            ("3. Wunsch", stats.get("wish3_fulfilled", 0)),
            ("4. Wunsch", stats.get("wish4_fulfilled", 0)),
            ("5. Wunsch", stats.get("wish5_fulfilled", 0)),
            ("Kein Wunsch", stats.get("no_wish_fulfilled", 0))
        ]
        
        for i, (label, count) in enumerate(wish_counts):
            ttk.Label(
                self.wish_frame,
                text=f"{label}: {count}",
                font=("Helvetica", 10)
            ).grid(row=i, column=0, sticky="w", padx=5, pady=2)
        
        # Populate student statistics
        student_stats = [
            ("Gesamtanzahl Schüler", stats.get("total_students", 0)),
            ("Schüler mit mind. einem Wunsch", 
             f"{stats.get('students_with_at_least_one_wish', 0)} ({stats.get('students_with_at_least_one_wish_pct', 0):.1f}%)"),
            ("Schüler mit Top-3 Wunsch", 
             f"{stats.get('students_with_top_three_wishes', 0)} ({stats.get('students_with_top_three_wishes_pct', 0):.1f}%)"),
            ("Schüler mit allen Slots", 
             f"{stats.get('students_with_all_five_sessions', 0)} ({stats.get('students_with_all_five_sessions_pct', 0):.1f}%)")
        ]
        
        for i, (label, value) in enumerate(student_stats):
            ttk.Label(
                self.student_frame,
                text=f"{label}: {value}",
                font=("Helvetica", 10)
            ).grid(row=i, column=0, sticky="w", padx=5, pady=2)
        
        # Populate student table
        if not student_df.empty:
            # Table headers
            headers = ["ID", "Name", "Erfüllung %", "1.", "2.", "3.", "4.", "5.", "Keine"]
            
            # Create header row
            for col, header in enumerate(headers):
                ttk.Label(
                    self.student_table,
                    text=header,
                    font=("Helvetica", 10, "bold")
                ).grid(row=0, column=col, padx=5, pady=5, sticky="w")
            
            # Add student rows
            for i, (_, row) in enumerate(student_df.iterrows(), 1):
                ttk.Label(
                    self.student_table,
                    text=row["Student ID"],
                    font=("Helvetica", 9)
                ).grid(row=i, column=0, padx=5, pady=2, sticky="w")
                
                ttk.Label(
                    self.student_table,
                    text=row["Name"],
                    font=("Helvetica", 9)
                ).grid(row=i, column=1, padx=5, pady=2, sticky="w")
                
                ttk.Label(
                    self.student_table,
                    text=f"{row['Fulfillment %']:.1f}%",
                    font=("Helvetica", 9)
                ).grid(row=i, column=2, padx=5, pady=2, sticky="w")
                
                # Wish counts
                ttk.Label(
                    self.student_table,
                    text=str(row["1st Wishes"]),
                    font=("Helvetica", 9)
                ).grid(row=i, column=3, padx=5, pady=2, sticky="w")
                
                ttk.Label(
                    self.student_table,
                    text=str(row["2nd Wishes"]),
                    font=("Helvetica", 9)
                ).grid(row=i, column=4, padx=5, pady=2, sticky="w")
                
                ttk.Label(
                    self.student_table,
                    text=str(row["3rd Wishes"]),
                    font=("Helvetica", 9)
                ).grid(row=i, column=5, padx=5, pady=2, sticky="w")
                
                ttk.Label(
                    self.student_table,
                    text=str(row["4th Wishes"]),
                    font=("Helvetica", 9)
                ).grid(row=i, column=6, padx=5, pady=2, sticky="w")
                
                ttk.Label(
                    self.student_table,
                    text=str(row["5th Wishes"]),
                    font=("Helvetica", 9)
                ).grid(row=i, column=7, padx=5, pady=2, sticky="w")
                
                ttk.Label(
                    self.student_table,
                    text=str(row["No Match"]),
                    font=("Helvetica", 9)
                ).grid(row=i, column=8, padx=5, pady=2, sticky="w")
                
                # Add alternating row colors
                if i % 2 == 0:
                    for col in range(len(headers)):
                        self.student_table.grid_columnconfigure(col, weight=1)
        else:
            ttk.Label(
                self.student_table,
                text="Keine Daten verfügbar. Bitte generieren Sie zuerst einen Zeitplan.",
                font=("Helvetica", 10)
            ).grid(row=0, column=0, padx=5, pady=5, sticky="w")

    def _export_fulfillment_excel(self):
        """Export the fulfillment report to Excel"""
        try:
            if not self.scheduler.get_schedule():
                messagebox.showinfo("Information", "Bitte generieren Sie zuerst einen Zeitplan.")
                return
            
            # Get fulfillment data
            stats = self.scheduler.get_fulfillment_statistics()
            student_df = self.scheduler.get_student_fulfillment_scores()
            
            if student_df.empty:
                messagebox.showinfo("Information", "Keine Daten für einen Erfüllungsbericht verfügbar.")
                return
            
            # Get file path
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Export Fulfillment Report"
            )
            
            if not file_path:
                return
            
            # Create Excel file
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Export the detailed student DataFrame
                student_df.to_excel(writer, sheet_name="Schülerdetails", index=False)
                
                # Create a summary sheet
                summary_data = {
                    "Wunsch": ["1. Wunsch", "2. Wunsch", "3. Wunsch", "4. Wunsch", "5. Wunsch", "Kein Wunsch"],
                    "Anzahl": [
                        stats.get("wish1_fulfilled", 0),
                        stats.get("wish2_fulfilled", 0),
                        stats.get("wish3_fulfilled", 0),
                        stats.get("wish4_fulfilled", 0),
                        stats.get("wish5_fulfilled", 0),
                        stats.get("no_wish_fulfilled", 0)
                    ],
                    "Gewichtung": [5, 4, 3, 2, 1, 0],
                    "Punkte": [
                        5 * stats.get("wish1_fulfilled", 0),
                        4 * stats.get("wish2_fulfilled", 0),
                        3 * stats.get("wish3_fulfilled", 0), 
                        2 * stats.get("wish4_fulfilled", 0),
                        1 * stats.get("wish5_fulfilled", 0),
                        0
                    ]
                }
                
                summary_df = pd.DataFrame(summary_data)
                
                # Add totals
                total_points = sum(summary_data["Punkte"])
                total_wishes = sum(summary_data["Anzahl"])
                summary_df.loc["Total"] = ["Gesamt", total_wishes, "", total_points]
                
                # Add the overall statistics
                summary_df2 = pd.DataFrame({
                    "Metrik": [
                        "Gesamterfüllungsgrad (gewichtet)", 
                        "Schüler mit mindestens einem Wunsch",
                        "Schüler mit einem Top-3 Wunsch",
                        "Schüler mit allen 5 Zeitslots",
                        "Anzahl Schüler gesamt"
                    ],
                    "Wert": [
                        f"{stats.get('average_weighted_fulfillment', 0):.2f}%",
                        f"{stats.get('students_with_at_least_one_wish_pct', 0):.2f}% ({stats.get('students_with_at_least_one_wish', 0)} von {stats.get('total_students', 0)})",
                        f"{stats.get('students_with_top_three_wishes_pct', 0):.2f}% ({stats.get('students_with_top_three_wishes', 0)} von {stats.get('total_students', 0)})",
                        f"{stats.get('students_with_all_five_sessions_pct', 0):.2f}% ({stats.get('students_with_all_five_sessions', 0)} von {stats.get('total_students', 0)})",
                        str(stats.get('total_students', 0))
                    ]
                })
                
                # Export summaries
                summary_df.to_excel(writer, sheet_name="Zusammenfassung", index=False, startrow=0)
                summary_df2.to_excel(writer, sheet_name="Zusammenfassung", index=False, startrow=len(summary_df) + 3)
            
            messagebox.showinfo("Information", f"Erfüllungsbericht wurde nach {file_path} exportiert.")
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Fehler beim Exportieren: {str(e)}")
            
    def _refresh_fulfillment_score(self):
        """Refresh just the fulfillment score (called from the Refresh button)"""
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
            return
            
        fulfillment_score = self.scheduler.calculate_overall_fulfillment_score()
        self.score_label.config(text=f"{fulfillment_score:.2f}%")

    def setup_ui(self):
        self.exports_frame = ttk.Frame(self, padding="10")
        self.exports_frame.grid(row=0, column=0, sticky="nsew")

        # Instructions label
        ttk.Label(
            self.exports_frame,
            text="Nachdem der Zeitplan generiert wurde, können Sie verschiedene Exporte erstellen:",
            font=("Helvetica", 12),
        ).grid(row=0, column=0, columnspan=4, pady=10, sticky="w")

        # Buttons frame
        buttons_frame = ttk.Frame(self.exports_frame, padding="5")
        buttons_frame.grid(row=1, column=0, columnspan=4, pady=10, sticky="w")

        # Export buttons
        ttk.Button(
            buttons_frame,
            text="Schüler-Zeitpläne exportieren",
            command=self.export_student_schedules,
            width=25,
        ).grid(row=0, column=0, padx=5, pady=5)

        ttk.Button(
            buttons_frame,
            text="Unternehmen-Zeitpläne exportieren",
            command=self.export_company_schedules,
            width=25,
        ).grid(row=0, column=1, padx=5, pady=5)

        ttk.Button(
            buttons_frame,
            text="Anwesenheitslisten exportieren",
            command=self.export_attendance_lists,
            width=25,
        ).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Button(
            buttons_frame,
            text="Raumbelegungsplan (PDF)",
            command=self.export_room_schedule,
            width=25,
        ).grid(row=1, column=0, padx=5, pady=5)
        
        ttk.Button(
            buttons_frame,
            text="Erfüllungsbericht (Ansicht)",
            command=self._refresh_fulfillment_display,
            width=25,
        ).grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Button(
            buttons_frame,
            text="Komplett-Export (ZIP)",
            command=self.export_all,
            width=25,
        ).grid(row=1, column=2, padx=5, pady=5)

        # Results display - add a treeview to show export preview
        self.results_frame = ttk.LabelFrame(self.exports_frame, text="Vorschau", padding="10")
        self.results_frame.grid(row=2, column=0, columnspan=4, sticky="nsew", pady=10)
        self.exports_frame.rowconfigure(2, weight=1)
        self.exports_frame.columnconfigure(0, weight=1)

        # Create a canvas with scrollbar for the preview
        self.canvas = tk.Canvas(self.results_frame)
        self.scrollbar = ttk.Scrollbar(
            self.results_frame, orient="vertical", command=self.canvas.yview
        )
        self.preview_frame = ttk.Frame(self.canvas)

        # Configure canvas
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.create_window((0, 0), window=self.preview_frame, anchor="nw")

        # Enable mousewheel scrolling
        self.canvas.bind_all(
            "<MouseWheel>", lambda event: self._on_mousewheel(event, self.canvas)
        )

        # Grid layout for canvas and scrollbar
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.results_frame.columnconfigure(0, weight=1)
        self.results_frame.rowconfigure(0, weight=1) 

    def export_company_schedules(self):
        self.app.clear_error()
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
            return
        
        filetypes = [("PDF files", "*.pdf"), ("Excel files", "*.xlsx")]
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=filetypes
        )
        
        if filepath:
            # Determine export type based on file extension
            if filepath.lower().endswith('.pdf'):
                if self.scheduler.export_company_overview_pdf(filepath):
                    self.app.clear_error()
            elif filepath.lower().endswith('.xlsx'):
                if self.scheduler.export_company_overview_excel(filepath):
                    self.app.clear_error()
            else:
                # Default to Excel if extension is unclear
                if self.scheduler.export_company_overview_excel(filepath + '.xlsx'):
                    self.app.clear_error()

    def export_room_schedule(self):
        self.app.clear_error()
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
            return
            
        # Open file save dialog
        filepath = filedialog.asksaveasfilename(
            title="Raumbelegungsplan speichern",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if filepath:
            messagebox.showinfo("Information", "Diese Funktion wird in einer zukünftigen Version verfügbar sein.")
            # TODO: Implement room schedule export functionality in a future version

    def export_all(self):
        self.app.clear_error()
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
            return
            
        # Ask for a directory to save all exports
        export_dir = filedialog.askdirectory(title="Wählen Sie einen Speicherort für alle Exporte")
        if not export_dir:
            return  # User cancelled
            
        try:
            import zipfile
            from datetime import datetime
            
            # Create a timestamp for the export
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Create a ZIP file to contain all exports
            zip_filepath = os.path.join(export_dir, f"Zeitplan_Export_{timestamp}.zip")
            
            # Create temporary directory for all files
            import tempfile
            temp_dir = tempfile.mkdtemp()
            
            # Export all files
            student_file_pdf = os.path.join(temp_dir, "Schüler_Zeitpläne.pdf")
            student_file_xlsx = os.path.join(temp_dir, "Schüler_Zeitpläne.xlsx")
            company_file_pdf = os.path.join(temp_dir, "Unternehmen_Übersicht.pdf")
            company_file_xlsx = os.path.join(temp_dir, "Unternehmen_Übersicht.xlsx")
            attendance_file_pdf = os.path.join(temp_dir, "Anwesenheitslisten.pdf")
            attendance_file_xlsx = os.path.join(temp_dir, "Anwesenheitslisten.xlsx")
            room_list_xlsx = os.path.join(temp_dir, "Raumliste.xlsx")
            
            # Export all files
            self.scheduler.export_student_schedules_pdf(student_file_pdf)
            self.scheduler.export_student_schedules_excel(student_file_xlsx)
            self.scheduler.export_company_overview_pdf(company_file_pdf)
            self.scheduler.export_company_overview_excel(company_file_xlsx)
            self.scheduler.export_attendance_lists_pdf(attendance_file_pdf, preview_mode=False)
            self.scheduler.export_attendance_lists_excel(attendance_file_xlsx, preview_mode=False)
            self.scheduler.export_room_list_excel(room_list_xlsx)
            
            # Create fulfillment report
            fulfillment_file = os.path.join(temp_dir, "Erfüllungsbericht.xlsx")
            stats = self.scheduler.get_fulfillment_statistics()
            student_df = self.scheduler.get_student_fulfillment_scores()
            
            # Create Excel file for fulfillment report
            with pd.ExcelWriter(fulfillment_file, engine='openpyxl') as writer:
                # Export the detailed student DataFrame
                student_df.to_excel(writer, sheet_name="Schülerdetails", index=False)
                
                # Create a summary sheet
                summary_data = {
                    "Wunsch": ["1. Wunsch", "2. Wunsch", "3. Wunsch", "4. Wunsch", "5. Wunsch", "Kein Wunsch"],
                    "Anzahl": [
                        stats.get("wish1_fulfilled", 0),
                        stats.get("wish2_fulfilled", 0),
                        stats.get("wish3_fulfilled", 0),
                        stats.get("wish4_fulfilled", 0),
                        stats.get("wish5_fulfilled", 0),
                        stats.get("no_wish_fulfilled", 0)
                    ],
                    "Gewichtung": [5, 4, 3, 2, 1, 0],
                    "Punkte": [
                        5 * stats.get("wish1_fulfilled", 0),
                        4 * stats.get("wish2_fulfilled", 0),
                        3 * stats.get("wish3_fulfilled", 0), 
                        2 * stats.get("wish4_fulfilled", 0),
                        1 * stats.get("wish5_fulfilled", 0),
                        0
                    ]
                }
                
                summary_df = pd.DataFrame(summary_data)
                
                # Add totals
                total_points = sum(summary_data["Punkte"])
                total_wishes = sum(summary_data["Anzahl"])
                summary_df.loc["Total"] = ["Gesamt", total_wishes, "", total_points]
                
                # Add the overall statistics
                summary_df2 = pd.DataFrame({
                    "Metrik": [
                        "Gesamterfüllungsgrad (gewichtet)", 
                        "Schüler mit mindestens einem Wunsch",
                        "Schüler mit einem Top-3 Wunsch",
                        "Schüler mit allen 5 Zeitslots",
                        "Anzahl Schüler gesamt"
                    ],
                    "Wert": [
                        f"{stats.get('average_weighted_fulfillment', 0):.2f}%",
                        f"{stats.get('students_with_at_least_one_wish_pct', 0):.2f}% ({stats.get('students_with_at_least_one_wish', 0)} von {stats.get('total_students', 0)})",
                        f"{stats.get('students_with_top_three_wishes_pct', 0):.2f}% ({stats.get('students_with_top_three_wishes', 0)} von {stats.get('total_students', 0)})",
                        f"{stats.get('students_with_all_five_sessions_pct', 0):.2f}% ({stats.get('students_with_all_five_sessions', 0)} von {stats.get('total_students', 0)})",
                        str(stats.get('total_students', 0))
                    ]
                })
                
                # Export summaries
                summary_df.to_excel(writer, sheet_name="Zusammenfassung", index=False)
            
            # Create the ZIP file containing all exports
            with zipfile.ZipFile(zip_filepath, 'w') as zipf:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, temp_dir)
                        zipf.write(file_path, arcname)
            
            # Clean up temporary directory
            import shutil
            shutil.rmtree(temp_dir)
            
            messagebox.showinfo("Information", f"Alle Exporte wurden nach {zip_filepath} exportiert.")
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Fehler beim Exportieren: {str(e)}") 

    def export_room_list(self):
        self.app.clear_error()
        # Open file save dialog
        filepath = filedialog.asksaveasfilename(
            title="Raumliste speichern",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if filepath:
            if self.scheduler.export_room_list_excel(filepath):
                self.app.clear_error()

    def _setup_exports_section(self):
        """Set up the exports section tab with buttons for all export types"""
        self.exports_overview_frame = ttk.Frame(self.export_notebook)
        self.export_notebook.add(self.exports_overview_frame, text="Exporte")
        
        # Main exports frame with label frame
        exports_frame = ttk.LabelFrame(self.exports_overview_frame, text="Exportieren")
        exports_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=5)
        
        # Instructions label
        ttk.Label(
            exports_frame,
            text="Nachdem der Zeitplan generiert wurde, können Sie verschiedene Exporte erstellen:",
            font=("Helvetica", 11),
            wraplength=600
        ).grid(row=0, column=0, columnspan=3, padx=5, pady=10, sticky="w")
        
        # Column 0: Student schedules
        ttk.Label(exports_frame, text="Schülerzeitpläne:").grid(
            row=1, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Button(
            exports_frame,
            text="PDF",
            command=self.export_student_schedules_pdf,
        ).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(
            exports_frame,
            text="Excel",
            command=self.export_student_schedules_excel,
        ).grid(row=1, column=2, padx=5, pady=5)
        
        # Column 1: Company overview
        ttk.Label(exports_frame, text="Unternehmensübersicht:").grid(
            row=2, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Button(
            exports_frame,
            text="PDF",
            command=self.export_company_overview_pdf,
        ).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(
            exports_frame,
            text="Excel",
            command=self.export_company_overview_excel,
        ).grid(row=2, column=2, padx=5, pady=5)
        
        # Row 2: Attendance lists
        ttk.Label(exports_frame, text="Anwesenheitslisten:").grid(
            row=3, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Button(
            exports_frame,
            text="PDF",
            command=self.export_attendance_lists_pdf,
        ).grid(row=3, column=1, padx=5, pady=5)
        ttk.Button(
            exports_frame,
            text="Excel",
            command=self.export_attendance_lists_excel,
        ).grid(row=3, column=2, padx=5, pady=5)
        
        # Row 3: Room list
        ttk.Label(exports_frame, text="Raumliste:").grid(
            row=4, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Button(
            exports_frame,
            text="Excel",
            command=self.export_room_list,
        ).grid(row=4, column=1, padx=5, pady=5)
        
        # Row 4: Room schedule (currently disabled)
        ttk.Label(exports_frame, text="Raumbelegungsplan:").grid(
            row=5, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Button(
            exports_frame,
            text="Excel",
            command=self.export_room_schedule,
        ).grid(row=5, column=1, padx=5, pady=5)
        
        # Row 5: All-in-one export
        ttk.Label(exports_frame, text="Alles exportieren:").grid(
            row=6, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Button(
            exports_frame,
            text="ZIP",
            command=self.export_all,
        ).grid(row=6, column=1, padx=5, pady=5)
        
        # Configure column weights
        exports_frame.columnconfigure(0, weight=1)
        exports_frame.columnconfigure(1, weight=0)
        exports_frame.columnconfigure(2, weight=0)
        
        # Configure frame weights
        self.exports_overview_frame.columnconfigure(0, weight=1)
        self.exports_overview_frame.rowconfigure(0, weight=1) 

    def export_student_schedules_pdf(self):
        self.app.clear_error()
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
            return
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".pdf", 
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if filepath:
            if self.scheduler.export_student_schedules_pdf(filepath):
                self.app.clear_error()

    def export_student_schedules_excel(self):
        self.app.clear_error()
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
            return
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if filepath:
            if self.scheduler.export_student_schedules_excel(filepath):
                self.app.clear_error()

    def export_company_overview_pdf(self):
        self.app.clear_error()
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
            return
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".pdf", 
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if filepath:
            if self.scheduler.export_company_overview_pdf(filepath):
                self.app.clear_error()

    def export_company_overview_excel(self):
        self.app.clear_error()
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
            return
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if filepath:
            if self.scheduler.export_company_overview_excel(filepath):
                self.app.clear_error()

    def export_attendance_lists_pdf(self):
        self.app.clear_error()
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
            return
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".pdf", 
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if filepath:
            if self.scheduler.export_attendance_lists_pdf(filepath, preview_mode=False):
                self.app.clear_error()

    def export_attendance_lists_excel(self):
        self.app.clear_error()
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
            return
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if filepath:
            if self.scheduler.export_attendance_lists_excel(filepath, preview_mode=False):
                self.app.clear_error() 