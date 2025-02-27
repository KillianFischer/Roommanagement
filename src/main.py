import os
from dotenv import load_dotenv
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

from services.scheduler import SchedulerService

load_dotenv()

class RoomManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Room Management")
        self.root.geometry("1200x800")

        # color scheme ToDo: dark mode, modern
        self.colors = {
            'bg': '#1e1e1e',
            'fg': '#ffffff',
            'accent': '#007acc',
            'accent_light': '#0098ff',
            'secondary_bg': '#252526',
            'border': '#333333',
            'error': '#ff3333',
            'success': '#33cc33',
            'header_bg': '#2d2d2d',
            'hover': '#2a2d2e'
        }

        # styles
        style = ttk.Style()
        
        # main theme
        style.configure("TFrame", background=self.colors['bg'])
        style.configure("Secondary.TFrame", background=self.colors['secondary_bg'])
        
        # notebook style ToDo: maybe change?
        style.configure("TNotebook", background=self.colors['bg'])
        style.configure("TNotebook.Tab",
            padding=[15, 8],
            background=self.colors['secondary_bg'],
            foreground=self.colors['fg'],
            font=("Helvetica", 10)
        )
        style.map("TNotebook.Tab",
            background=[("selected", self.colors['accent'])],
            foreground=[("selected", self.colors['fg'])]
        )
        
        # button styles
        style.configure("TButton",
            padding=8,
            font=("Helvetica", 10),
            background=self.colors['secondary_bg'],
            foreground=self.colors['fg']
        )
        style.map("TButton",
            background=[("active", self.colors['hover'])],
            foreground=[("active", self.colors['fg'])]
        )
        
        style.configure("Action.TButton",
            padding=[12, 8],
            font=("Helvetica", 10, "bold"),
            background=self.colors['accent'],
            foreground=self.colors['fg']
        )
        style.map("Action.TButton",
            background=[("active", self.colors['accent_light'])],
            foreground=[("active", self.colors['fg'])]
        )
        
        # Conftreeview styles
        style.configure("Treeview",
            background=self.colors['secondary_bg'],
            foreground=self.colors['fg'],
            fieldbackground=self.colors['secondary_bg'],
            font=("Helvetica", 10),
            rowheight=30
        )
        style.configure("Treeview.Heading",
            background=self.colors['header_bg'],
            foreground=self.colors['fg'],
            font=("Helvetica", 10, "bold")
        )
        style.map("Treeview",
            background=[("selected", self.colors['accent'])],
            foreground=[("selected", self.colors['fg'])]
        )
        
        style.configure("Schedule.Treeview",
            background=self.colors['secondary_bg'],
            foreground=self.colors['fg'],
            fieldbackground=self.colors['secondary_bg'],
            font=("Helvetica", 10),
            rowheight=45
        )
        
        # label styles
        style.configure("TLabel",
            background=self.colors['bg'],
            foreground=self.colors['fg'],
            font=("Helvetica", 10)
        )
        style.configure("Header.TLabel",
            background=self.colors['bg'],
            foreground=self.colors['fg'],
            font=("Helvetica", 12, "bold")
        )
        style.configure("Title.TLabel",
            background=self.colors['bg'],
            foreground=self.colors['fg'],
            font=("Helvetica", 14, "bold")
        )
        
        # scrollbar style
        style.configure("Vertical.TScrollbar",
            background=self.colors['secondary_bg'],
            troughcolor=self.colors['bg'],
            arrowcolor=self.colors['fg']
        )
        
        # canvas style for previews
        style.configure("Preview.TCanvas",
            background=self.colors['secondary_bg']
        )

        # root window background
        self.root.configure(bg=self.colors['bg'])
        
        # Scheduler instance
        self.scheduler = SchedulerService()

        # Load environment variables and setup import folder
        self.dev_mode = os.getenv('DEV_MODE', 'false').lower() == 'true'
        self.import_folder = os.getenv('IMPORT_FOLDER', 'import/')
        
        # Create import folder if it doesn't exist
        if self.dev_mode and not os.path.exists(self.import_folder):
            os.makedirs(self.import_folder)
        
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky="nsew")

        # three tabs: Import, Schedule, Export
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        # Import: Student wishes, Company list, Room list
        self.import_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.import_frame, text="Daten importieren")

        # canvas and scrollbar for the entire import tab
        self.import_canvas = tk.Canvas(
            self.import_frame,
            background=self.colors['bg'],
            highlightthickness=0
        )
        self.import_scrollbar = ttk.Scrollbar(
            self.import_frame,
            orient="vertical",
            command=self.import_canvas.yview,
            style="Vertical.TScrollbar"
        )
        
        # canvas
        self.import_canvas.configure(yscrollcommand=self.import_scrollbar.set)
        
        # main container for sections
        self.import_sections = ttk.Frame(self.import_canvas, style="Secondary.TFrame")
        
        # window in canvas
        self.import_canvas_window = self.import_canvas.create_window(
            (0, 0),
            window=self.import_sections,
            anchor="nw",
            width=self.import_canvas.winfo_width()
        )

        # Grid layout for canvas and scrollbar
        self.import_canvas.grid(row=0, column=0, sticky="nsew", padx=(20,0), pady=20)
        self.import_scrollbar.grid(row=0, column=1, sticky="ns", pady=20)

        # weights for import frame
        self.import_frame.columnconfigure(0, weight=1)
        self.import_frame.rowconfigure(0, weight=1)

        # Bind events for scrolling and resizing
        self.import_canvas.bind('<Configure>', self._on_canvas_configure)
        self.import_sections.bind('<Configure>', self._on_frame_configure)
        self.import_canvas.bind_all("<MouseWheel>", lambda e: self._on_mousewheel(e, self.import_canvas))

        # Student wishes section
        section_frame = ttk.Frame(self.import_sections, style="Secondary.TFrame")
        section_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 20))
        
        ttk.Label(
            section_frame,
            text="Schülerwünsche",
            style="Title.TLabel"
        ).grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky="w")
        
        import_btn = ttk.Button(
            section_frame,
            text="Import",
            command=self.import_preferences,
            style="Action.TButton"
        )
        import_btn.grid(row=1, column=0, pady=2, padx=(0, 10), sticky="w")
        
        self.preferences_status = ttk.Label(
            section_frame,
            text="No file imported",
            foreground=self.colors['fg'],
            style="TLabel"
        )
        self.preferences_status.grid(row=1, column=1, pady=2, sticky="w")
        
        # frame for preview with scrollbar
        preview_frame = ttk.Frame(section_frame, style="Secondary.TFrame")
        preview_frame.grid(row=2, column=0, columnspan=2, pady=(5, 0), sticky="nsew")
        
        self.preferences_preview = ttk.Treeview(
            preview_frame,
            height=6,
            style="Treeview"
        )
        preferences_scrollbar = ttk.Scrollbar(
            preview_frame,
            orient="vertical",
            command=self.preferences_preview.yview,
            style="Vertical.TScrollbar"
        )
        self.preferences_preview.configure(yscrollcommand=preferences_scrollbar.set)
        
        self.preferences_preview.grid(row=0, column=0, sticky="nsew")
        preferences_scrollbar.grid(row=0, column=1, sticky="ns")
        
        preview_frame.columnconfigure(0, weight=1)
        section_frame.columnconfigure(1, weight=1)

        # Company list section
        section_frame = ttk.Frame(self.import_sections, style="Secondary.TFrame")
        section_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 20))
        
        ttk.Label(
            section_frame,
            text="Unternehmensliste",
            style="Title.TLabel"
        ).grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky="w")
        
        import_btn = ttk.Button(
            section_frame,
            text="Import",
            command=self.import_companies,
            style="Action.TButton"
        )
        import_btn.grid(row=1, column=0, pady=2, padx=(0, 10), sticky="w")
        
        self.companies_status = ttk.Label(
            section_frame,
            text="No file imported",
            foreground=self.colors['fg'],
            style="TLabel"
        )
        self.companies_status.grid(row=1, column=1, pady=2, sticky="w")
        
        # frame for preview with scrollbar
        preview_frame = ttk.Frame(section_frame, style="Secondary.TFrame")
        preview_frame.grid(row=2, column=0, columnspan=2, pady=(5, 0), sticky="nsew")
        
        self.companies_preview = ttk.Treeview(
            preview_frame,
            height=6,
            style="Treeview"
        )
        companies_scrollbar = ttk.Scrollbar(
            preview_frame,
            orient="vertical",
            command=self.companies_preview.yview,
            style="Vertical.TScrollbar"
        )
        self.companies_preview.configure(yscrollcommand=companies_scrollbar.set)
        
        self.companies_preview.grid(row=0, column=0, sticky="nsew")
        companies_scrollbar.grid(row=0, column=1, sticky="ns")
        
        preview_frame.columnconfigure(0, weight=1)
        section_frame.columnconfigure(1, weight=1)

        # Room list section
        section_frame = ttk.Frame(self.import_sections, style="Secondary.TFrame")
        section_frame.grid(row=2, column=0, sticky="nsew")
        
        ttk.Label(
            section_frame,
            text="Raumliste",
            style="Title.TLabel"
        ).grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky="w")
        
        import_btn = ttk.Button(
            section_frame,
            text="Import",
            command=self.import_rooms,
            style="Action.TButton"
        )
        import_btn.grid(row=1, column=0, pady=2, padx=(0, 10), sticky="w")
        
        self.rooms_status = ttk.Label(
            section_frame,
            text="No file imported",
            foreground=self.colors['fg'],
            style="TLabel"
        )
        self.rooms_status.grid(row=1, column=1, pady=2, sticky="w")
        
        # frame for preview with scrollbar
        preview_frame = ttk.Frame(section_frame, style="Secondary.TFrame")
        preview_frame.grid(row=2, column=0, columnspan=2, pady=(5, 0), sticky="nsew")
        
        self.rooms_preview = ttk.Treeview(
            preview_frame,
            height=6,
            style="Treeview"
        )
        rooms_scrollbar = ttk.Scrollbar(
            preview_frame,
            orient="vertical",
            command=self.rooms_preview.yview,
            style="Vertical.TScrollbar"
        )
        self.rooms_preview.configure(yscrollcommand=rooms_scrollbar.set)
        
        self.rooms_preview.grid(row=0, column=0, sticky="nsew")
        rooms_scrollbar.grid(row=0, column=1, sticky="ns")
        
        preview_frame.columnconfigure(0, weight=1)
        section_frame.columnconfigure(1, weight=1)

        # import sections layout
        self.import_sections.columnconfigure(0, weight=1)
        
        # import frame layout
        self.import_frame.columnconfigure(0, weight=1)
        self.import_frame.rowconfigure(0, weight=1)

        # Schedule Tab
        self.schedule_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.schedule_frame, text="Zeitplan")
        
        # Control buttons frame
        self.schedule_controls = ttk.Frame(self.schedule_frame)
        self.schedule_controls.grid(row=0, column=0, sticky="ew", padx=15, pady=(15,5))
        
        ttk.Button(
            self.schedule_controls,
            text="Zeitplan generieren",
            command=self.generate_schedule,
            style="Action.TButton"
        ).grid(row=0, column=0, padx=5)
        
        ttk.Button(
            self.schedule_controls,
            text="Zeitplan exportieren",
            command=self.export_schedule,
            style="Action.TButton"
        ).grid(row=0, column=1, padx=5)
        
        # Schedule display frame with scrollbar
        self.schedule_frame_inner = ttk.Frame(self.schedule_frame)
        self.schedule_frame_inner.grid(row=1, column=0, sticky="nsew", padx=15, pady=15)
        
        # scrollbar for schedule tree
        self.schedule_scrollbar = ttk.Scrollbar(self.schedule_frame_inner)
        self.schedule_scrollbar.grid(row=0, column=1, sticky="ns")
        
        self.schedule_tree = ttk.Treeview(
            self.schedule_frame_inner,
            yscrollcommand=self.schedule_scrollbar.set,
            style="Schedule.Treeview"
        )
        self.schedule_tree.grid(row=0, column=0, sticky="nsew")
        
        self.schedule_scrollbar.config(command=self.schedule_tree.yview)
        
        # weights for schedule frames
        self.schedule_frame.columnconfigure(0, weight=1)
        self.schedule_frame.rowconfigure(1, weight=1)
        self.schedule_frame_inner.columnconfigure(0, weight=1)
        self.schedule_frame_inner.rowconfigure(0, weight=1)

        # Export Tab
        self.export_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.export_frame, text="Exportieren")
        
        # nested notebook for export previews
        self.export_notebook = ttk.Notebook(self.export_frame)
        self.export_notebook.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        # Student Schedules tab
        self.student_schedules_frame = ttk.Frame(self.export_notebook)
        self.export_notebook.add(self.student_schedules_frame, text="Schülerzeitpläne")
        
        ttk.Button(self.student_schedules_frame, text="Vorschau", command=self.preview_student_schedules).grid(row=0, column=0, pady=5, padx=5)
        ttk.Button(self.student_schedules_frame, text="Als PDF exportieren", command=self.export_student_schedules).grid(row=0, column=1, pady=5, padx=5)
        
        # canvas and scrollbar for the preview
        self.student_preview_canvas = tk.Canvas(self.student_schedules_frame)
        self.student_preview_scrollbar = ttk.Scrollbar(self.student_schedules_frame, orient="vertical", command=self.student_preview_canvas.yview)
        self.student_preview_frame = ttk.Frame(self.student_preview_canvas)
        
        self.student_preview_canvas.configure(yscrollcommand=self.student_preview_scrollbar.set)
        
        # Bind mouse wheel for student preview
        self.student_preview_canvas.bind_all("<MouseWheel>", lambda e: self._on_mousewheel(e, self.student_preview_canvas))
        
        self.student_preview_canvas.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        self.student_preview_scrollbar.grid(row=1, column=2, sticky="ns")
        self.student_preview_canvas.create_window((0, 0), window=self.student_preview_frame, anchor="nw")
        
        self.student_schedules_frame.rowconfigure(1, weight=1)
        self.student_schedules_frame.columnconfigure(0, weight=1)
        self.student_schedules_frame.columnconfigure(1, weight=1)
        
        # Attendance Lists tab
        self.attendance_lists_frame = ttk.Frame(self.export_notebook)
        self.export_notebook.add(self.attendance_lists_frame, text="Anwesenheitslisten")
        
        ttk.Button(self.attendance_lists_frame, text="Vorschau", command=self.preview_attendance_lists).grid(row=0, column=0, pady=5, padx=5)
        ttk.Button(self.attendance_lists_frame, text="Als PDF exportieren", command=self.export_attendance_lists).grid(row=0, column=1, pady=5, padx=5)
        
        # canvas and scrollbar for the preview
        self.attendance_preview_canvas = tk.Canvas(self.attendance_lists_frame)
        self.attendance_preview_scrollbar = ttk.Scrollbar(self.attendance_lists_frame, orient="vertical", command=self.attendance_preview_canvas.yview)
        self.attendance_preview_frame = ttk.Frame(self.attendance_preview_canvas)
        
        self.attendance_preview_canvas.configure(yscrollcommand=self.attendance_preview_scrollbar.set)
        
        # Bind mouse wheel for attendance preview
        self.attendance_preview_canvas.bind_all("<MouseWheel>", lambda e: self._on_mousewheel(e, self.attendance_preview_canvas))
        
        self.attendance_preview_canvas.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        self.attendance_preview_scrollbar.grid(row=1, column=2, sticky="ns")
        self.attendance_preview_canvas.create_window((0, 0), window=self.attendance_preview_frame, anchor="nw")
        
        self.attendance_lists_frame.rowconfigure(1, weight=1)
        self.attendance_lists_frame.columnconfigure(0, weight=1)
        self.attendance_lists_frame.columnconfigure(1, weight=1)
        
        # export frame grid
        self.export_frame.columnconfigure(0, weight=1)
        self.export_frame.rowconfigure(0, weight=1)

        # grid settings
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(0, weight=1)

    def setup_preview_tree(self, tree, columns):
        tree['columns'] = columns
        tree.column('#0', width=0, stretch=tk.NO)
        for col in columns:
            tree.column(col, anchor=tk.W, width=150)  # Increased width
            tree.heading(col, text=col, anchor=tk.W)
        
        # alternating row colors
        tree.tag_configure('oddrow', background=self.colors['bg'])
        tree.tag_configure('evenrow', background=self.colors['secondary_bg'])

    def update_preview(self, tree, df, columns):
        for item in tree.get_children():
            tree.delete(item)
        for idx, row in df.head(6).iterrows():  # Show more rows
            values = [str(row[col]) if col in row and pd.notna(row[col]) else '' for col in columns]
            tree.insert('', tk.END, values=values, tags=('evenrow' if idx % 2 == 0 else 'oddrow'))

    def get_import_file(self, env_key, dialog_title="Select file"):
        if self.dev_mode:
            filename = os.getenv(env_key)
            if filename:
                filepath = os.path.normpath(os.path.join(self.import_folder, filename))
                if os.path.exists(filepath):
                    return filepath
        
        return filedialog.askopenfilename(
            title=dialog_title,
            filetypes=[("Excel files", "*.xlsx")]
        )

    def import_preferences(self):
        file_path = self.get_import_file('STUDENT_PREFERENCES', "Import Student Preferences")
        if file_path:
            try:
                df = pd.read_excel(file_path)
                df.columns = df.columns.str.strip()
                if self.scheduler.load_student_preferences(df):
                    self.preferences_status.config(text=f"Imported: {os.path.basename(file_path)}", foreground="green")
                    cols = ['Klasse', 'Name', 'Vorname'] + [f'Wahl {i}' for i in range(1, 7)]
                    self.setup_preview_tree(self.preferences_preview, cols)
                    self.update_preview(self.preferences_preview, df, cols)
                else:
                    self.preferences_status.config(text="Ungültiges Format", foreground="red")
            except Exception as e:
                self.preferences_status.config(text=f"Error: {str(e)}", foreground="red")

    def import_companies(self):
        file_path = self.get_import_file('COMPANY_LIST', "Import Company List")
        if file_path:
            try:
                df = pd.read_excel(file_path)
                df.columns = df.columns.str.strip()
                if self.scheduler.load_companies(df):
                    self.companies_status.config(text=f"Imported: {os.path.basename(file_path)}", foreground="green")
                    cols = ['Unternehmen', 'Fachrichtung', 'Max. Teilnehmer', 'Max. Veranstaltungen', 'Frühester Zeitpunkt']
                    self.setup_preview_tree(self.companies_preview, cols)
                    self.update_preview(self.companies_preview, df, cols)
                else:
                    self.companies_status.config(text="Ungültiges Format", foreground="red")
            except Exception as e:
                self.companies_status.config(text=f"Error: {str(e)}", foreground="red")

    def import_rooms(self):
        file_path = self.get_import_file('ROOM_LIST', "Import Room List")
        if file_path:
            try:
                df = pd.read_excel(file_path, header=None)
                if self.scheduler.load_rooms(df):
                    self.rooms_status.config(text=f"Imported: {os.path.basename(file_path)}", foreground="green")
                    self.setup_preview_tree(self.rooms_preview, ['Raum'])
                    self.update_preview(self.rooms_preview, df.rename(columns={df.columns[0]: 'Raum'}), ['Raum'])
                else:
                    self.rooms_status.config(text="Ungültiges Format", foreground="red")
            except Exception as e:
                self.rooms_status.config(text=f"Error: {str(e)}", foreground="red")

    def generate_schedule(self):
        if not self.scheduler.is_data_loaded():
            messagebox.showerror("Fehler", "Bitte alle erforderlichen Daten importieren!")
            return
        if self.scheduler.generate_schedule():
            self.update_schedule_display()
            messagebox.showinfo("Erfolg", "Zeitplan erfolgreich generiert!")
        else:
            messagebox.showerror("Fehler", "Zeitplan konnte nicht generiert werden. Bitte prüfen Sie Ihre Daten und Zeitslots.")

    def update_schedule_display(self):
        for item in self.schedule_tree.get_children():
            self.schedule_tree.delete(item)
        
        time_slots = [
            ('A', '8:45 – 9:30'),
            ('B', '9:50 – 10:35'),
            ('C', '10:35 – 11:20'),
            ('D', '11:40 – 12:25'),
            ('E', '12:25 – 13:10')
        ]
        self.scheduler.time_slots = time_slots

        columns = ['Company'] + [slot for slot, _ in time_slots]
        self.schedule_tree['columns'] = columns
        self.schedule_tree.column('#0', width=0, stretch=tk.NO)
        self.schedule_tree.column('Company', anchor=tk.W, width=250)
        self.schedule_tree.heading('Company', text='Unternehmen', anchor=tk.W)
        
        for col in columns:
            self.schedule_tree.column(col, anchor=tk.W, width=150)  
            if col == 'Company':
                self.schedule_tree.heading(col, text='Unternehmen', anchor=tk.W)
            else:
                time_range = dict(time_slots).get(col, '')
                self.schedule_tree.heading(col, text=f"{col} ({time_range})", anchor=tk.W)

        for company in self.scheduler.companies:
            row = [company.name]
            for slot_idx, (slot_letter, time_range) in enumerate(time_slots):
                if slot_idx < company.earliest_slot:
                    text = ""
                else:
                    session = self.scheduler.schedule.get((company.name, slot_idx))
                    if session:
                        count = len(session.students)
                        text = f"Raum: {session.room}"
                        if count > 0:
                            text += f"\n({count} Schü{'' if count == 1 else 'ler:innen'})"
                    else:
                        text = ""
                row.append(text)
            self.schedule_tree.insert('', tk.END, values=row)

    def export_student_schedules(self):
        if not self.scheduler.get_schedule():
            messagebox.showerror("Fehler", "Bitte erst den Zeitplan generieren!")
            return
        self.scheduler.export_student_schedules()

    def export_attendance_lists(self):
        if not self.scheduler.get_schedule():
            messagebox.showerror("Fehler", "Bitte erst den Zeitplan generieren!")
            return
        self.scheduler.export_attendance_lists(preview_mode=False)

    def preview_student_schedules(self):
        if not self.scheduler.get_schedule():
            messagebox.showerror("Fehler", "Bitte erst den Zeitplan generieren!")
            return

        # Clear previous preview
        for widget in self.student_preview_frame.winfo_children():
            widget.destroy()

        # Group students by class
        class_schedules = {}
        for student in self.scheduler.student_preferences:
            class_name = student.student_id.split('_')[0]
            if class_name not in class_schedules:
                class_schedules[class_name] = []
            
            # Collect all appointments for this student
            student_schedule = []
            realized_wishes = []
            for slot_idx, (slot_letter, time_range) in enumerate(self.scheduler.time_slots):
                session_found = False
                for wish_idx, wish in enumerate(student.wishes):
                    key = (str(wish).strip(), slot_idx)
                    if key in self.scheduler.schedule:
                        session = self.scheduler.schedule[key]
                        if any(s['id'] == student.student_id for s in session.students):
                            student_schedule.append({
                                'time': f"{slot_letter} ({time_range})",
                                'company': str(wish).strip(),
                                'room': session.room,
                                'wish_number': wish_idx + 1
                            })
                            realized_wishes.append(True)
                            session_found = True
                            break
                if not session_found:
                    realized_wishes.append(False)
            
            satisfaction_score = student.get_satisfaction_score(realized_wishes)
            class_schedules[class_name].append({
                'name': student.name,
                'schedule': sorted(student_schedule, key=lambda x: x['time']),
                'score': satisfaction_score
            })

        # Create preview for each class
        row = 0
        style = ttk.Style()
        style.configure("Preview.TLabel", font=("Helvetica", 10))
        style.configure("PreviewHeader.TLabel", font=("Helvetica", 10, "bold"))

        for class_name, students in sorted(class_schedules.items()):
            ttk.Label(
                self.student_preview_frame,
                text=f"Klasse {class_name}",
                style="PreviewHeader.TLabel"
            ).grid(row=row, column=0, columnspan=4, pady=(20, 10), sticky="w")
            row += 1

            for student in students:
                # Student header
                ttk.Label(
                    self.student_preview_frame,
                    text=f"{student['name']} - Bewertung: {student['score']:.1f}%",
                    style="PreviewHeader.TLabel"
                ).grid(row=row, column=0, columnspan=4, pady=(10, 5), sticky="w")
                row += 1

                # Schedule headers
                for col, header in enumerate(['Zeit', 'Unternehmen', 'Raum', 'Wunsch']):
                    ttk.Label(
                        self.student_preview_frame,
                        text=header,
                        style="PreviewHeader.TLabel"
                    ).grid(row=row, column=col, padx=5, pady=2, sticky="w")
                row += 1

                # Schedule rows
                for appointment in student['schedule']:
                    ttk.Label(
                        self.student_preview_frame,
                        text=appointment['time'],
                        style="Preview.TLabel"
                    ).grid(row=row, column=0, padx=5, pady=2, sticky="w")
                    ttk.Label(
                        self.student_preview_frame,
                        text=appointment['company'],
                        style="Preview.TLabel"
                    ).grid(row=row, column=1, padx=5, pady=2, sticky="w")
                    ttk.Label(
                        self.student_preview_frame,
                        text=appointment['room'],
                        style="Preview.TLabel"
                    ).grid(row=row, column=2, padx=5, pady=2, sticky="w")
                    ttk.Label(
                        self.student_preview_frame,
                        text=str(appointment['wish_number']),
                        style="Preview.TLabel"
                    ).grid(row=row, column=3, padx=5, pady=2, sticky="w")
                    row += 1

        # Update canvas scroll region
        self.student_preview_frame.update_idletasks()
        self.student_preview_canvas.configure(
            scrollregion=self.student_preview_canvas.bbox("all")
        )

    def preview_attendance_lists(self):
        if not self.scheduler.get_schedule():
            messagebox.showerror("Fehler", "Bitte erst den Zeitplan generieren!")
            return

        # Clear previous preview
        for widget in self.attendance_preview_frame.winfo_children():
            widget.destroy()
        
        # Sort sessions by company name and time slot
        sorted_sessions = sorted(
            self.scheduler.schedule.items(),
            key=lambda x: (x[0][0], x[0][1])  # Sort by company name, then slot
        )
        
        # Get unique company names and limit to first 6
        company_names = list(set(company_name for (company_name, _), _ in sorted_sessions))
        if len(company_names) > 6:
            company_names = company_names[:6]
            
        # Filter sessions to only include these companies
        sorted_sessions = [(key, session) for (key, session) in sorted_sessions 
                          if key[0] in company_names]

        # Create preview for each session
        row = 0
        style = ttk.Style()
        style.configure("Preview.TLabel", font=("Helvetica", 10))
        style.configure("PreviewHeader.TLabel", font=("Helvetica", 10, "bold"))

        for (company_name, slot_idx), session in sorted_sessions:
            # Session header
            ttk.Label(
                self.attendance_preview_frame,
                text=f"{company_name}",
                style="PreviewHeader.TLabel"
            ).grid(row=row, column=0, columnspan=4, pady=(20, 5), sticky="w")
            row += 1

            ttk.Label(
                self.attendance_preview_frame,
                text=f"Zeitfenster: {session.time_slot} ({session.time_range})",
                style="Preview.TLabel"
            ).grid(row=row, column=0, columnspan=4, pady=2, sticky="w")
            row += 1

            ttk.Label(
                self.attendance_preview_frame,
                text=f"Raum: {session.room}",
                style="Preview.TLabel"
            ).grid(row=row, column=0, columnspan=4, pady=(2, 10), sticky="w")
            row += 1

            # Attendance list headers
            for col, header in enumerate(['Nr.', 'Name', 'Klasse', 'Unterschrift']):
                ttk.Label(
                    self.attendance_preview_frame,
                    text=header,
                    style="PreviewHeader.TLabel"
                ).grid(row=row, column=col, padx=5, pady=2, sticky="w")
            row += 1

            # Attendee rows
            for i, student in enumerate(sorted(session.students, key=lambda x: x['name']), 1):
                class_name = student['id'].split('_')[0]
                ttk.Label(
                    self.attendance_preview_frame,
                    text=str(i),
                    style="Preview.TLabel"
                ).grid(row=row, column=0, padx=5, pady=2, sticky="w")
                ttk.Label(
                    self.attendance_preview_frame,
                    text=student['name'],
                    style="Preview.TLabel"
                ).grid(row=row, column=1, padx=5, pady=2, sticky="w")
                ttk.Label(
                    self.attendance_preview_frame,
                    text=class_name,
                    style="Preview.TLabel"
                ).grid(row=row, column=2, padx=5, pady=2, sticky="w")
                ttk.Label(
                    self.attendance_preview_frame,
                    text="________________",
                    style="Preview.TLabel"
                ).grid(row=row, column=3, padx=5, pady=2, sticky="w")
                row += 1

            # Add empty signature lines
            for i in range(5):
                ttk.Label(
                    self.attendance_preview_frame,
                    text=str(len(session.students) + i + 1),
                    style="Preview.TLabel"
                ).grid(row=row, column=0, padx=5, pady=2, sticky="w")
                ttk.Label(
                    self.attendance_preview_frame,
                    text="",
                    style="Preview.TLabel"
                ).grid(row=row, column=1, padx=5, pady=2, sticky="w")
                ttk.Label(
                    self.attendance_preview_frame,
                    text="",
                    style="Preview.TLabel"
                ).grid(row=row, column=2, padx=5, pady=2, sticky="w")
                ttk.Label(
                    self.attendance_preview_frame,
                    text="________________",
                    style="Preview.TLabel"
                ).grid(row=row, column=3, padx=5, pady=2, sticky="w")
                row += 1

        # Update canvas scroll region
        self.attendance_preview_frame.update_idletasks()
        self.attendance_preview_canvas.configure(
            scrollregion=self.attendance_preview_canvas.bbox("all")
        )

    def export_schedule(self):
        if not self.scheduler.get_schedule():
            messagebox.showerror("Fehler", "Bitte erst den Zeitplan generieren!")
            return
            
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import mm
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
            
            # Create PDF
            doc = SimpleDocTemplate(
                "schedule.pdf",
                pagesize=landscape(A4),
                rightMargin=10*mm,
                leftMargin=10*mm,
                topMargin=10*mm,
                bottomMargin=10*mm
            )
            
            story = []
            styles = getSampleStyleSheet()
            
            # Add title
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=16,
                spaceAfter=20
            )
            story.append(Paragraph("Zeitplan Übersicht", title_style))
            
            # Prepare table data
            time_slots = self.scheduler.time_slots
            headers = ['Unternehmen'] + [f"{slot} ({time})" for slot, time in time_slots]
            table_data = [headers]
            
            for company in self.scheduler.companies:
                row = [company.name]
                for slot_idx, _ in enumerate(time_slots):
                    if slot_idx < company.earliest_slot:
                        text = "---"
                    else:
                        session = self.scheduler.schedule.get((company.name, slot_idx))
                        if session:
                            count = len(session.students)
                            text = f"Raum {session.room}\n({count} TN)"
                        else:
                            text = "---"
                    row.append(text)
                table_data.append(row)
            
            # Create and style the table
            col_widths = [40*mm] + [30*mm] * len(time_slots)
            t = Table(table_data, colWidths=col_widths, repeatRows=1)
            t.setStyle(TableStyle([
                ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
                ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                ('TEXTCOLOR', (0,0), (-1,0), colors.black),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,0), 10),
                ('BOTTOMPADDING', (0,0), (-1,0), 12),
                ('BACKGROUND', (0,1), (-1,-1), colors.white),
                ('TEXTCOLOR', (0,1), (-1,-1), colors.black),
                ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
                ('FONTSIZE', (0,1), (-1,-1), 9),
                ('ALIGN', (0,1), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('GRID', (0,0), (-1,-1), 1, colors.black),
                ('BOX', (0,0), (-1,-1), 2, colors.black),
                ('LINEBELOW', (0,0), (-1,0), 2, colors.black),
            ]))
            
            story.append(t)
            doc.build(story)
            
            messagebox.showinfo(
                "Export erfolgreich",
                "Zeitplan wurde als schedule.pdf gespeichert."
            )
            
        except Exception as e:
            messagebox.showerror(
                "Export Fehler",
                f"Fehler beim Exportieren des Zeitplans: {str(e)}"
            )

    def _on_mousewheel(self, event, canvas):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_canvas_configure(self, event):
        # Update the scroll region when the canvas is resized
        self.import_canvas.itemconfig(
            self.import_canvas_window,
            width=event.width
        )

    def _on_frame_configure(self, event):
        # scrolling
        self.import_canvas.configure(
            scrollregion=self.import_canvas.bbox("all")
        )

if __name__ == "__main__":
    root = tk.Tk()
    app = RoomManagementApp(root)
    root.mainloop()
