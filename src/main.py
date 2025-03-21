import tkinter as tk
import os
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import configparser
import sys
from dotenv import load_dotenv

# Load environment variables early
load_dotenv()

from services.scheduler import Scheduler
from ui import ErrorDisplay

# TODO: Maybe separate class into several files to improve maintainability
class RoomManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Raumverwaltung für Berufsorientierungstag")
        self.root.geometry("1200x800")

        self.scheduler = Scheduler(error_handler=self.show_error)

        self.dev_mode = os.getenv("DEV_MODE", "false").lower() == "true"
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.import_folder = os.path.join(base_dir, os.getenv("IMPORT_FOLDER", "import/"))
        
        print(f"DEV_MODE: {self.dev_mode}")
        print(f"IMPORT_FOLDER: {self.import_folder}")
        print(f"Files in import folder: {os.listdir(self.import_folder) if os.path.exists(self.import_folder) else 'folder does not exist'}")

        if self.dev_mode and not os.path.exists(self.import_folder):
            os.makedirs(self.import_folder)
            
        # Set up global ttk styles
        self.setup_styles()

        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky="nsew")

        # Error display
        self.error_display = ErrorDisplay(self.main_frame, row=1, column=0)
        
        # Setup UI
        self.setup_ui()
        
        # Auto import in dev mode - Execute directly instead of waiting for timer
        if self.dev_mode:
            print("Dev mode detected, performing auto-import immediately")
            self.auto_import_files()

    def auto_import_files(self):
        """Automatically import files in dev mode if they exist in the import folder"""
        try:
            print("Starting auto import...")
            
            # Import rooms
            rooms_file = os.path.join(self.import_folder, os.getenv("ROOM_LIST"))
            if os.path.exists(rooms_file):
                print(f"Loading rooms from {rooms_file}")
                df = pd.read_excel(rooms_file, header=None)
                if self.scheduler.load_rooms(df):
                    self.rooms_status.config(
                        text=f"Imported: {os.path.basename(rooms_file)}",
                    )
                    self.setup_preview_tree(self.rooms_preview, ["Raum"])
                    self.update_preview(
                        self.rooms_preview,
                        df.rename(columns={df.columns[0]: "Raum"}),
                        ["Raum"],
                    )
                    print("Rooms imported successfully")
                else:
                    print("Failed to import rooms: Invalid format")
            else:
                print(f"Rooms file not found: {rooms_file}")
                
            # Import companies
            companies_file = os.path.join(self.import_folder, os.getenv("COMPANY_LIST"))
            if os.path.exists(companies_file):
                print(f"Loading companies from {companies_file}")
                df = pd.read_excel(companies_file)
                df.columns = df.columns.str.strip()
                
                # Handle different column names for minimum participants
                if "Min." in df.columns and "Min. Teilnehmer" not in df.columns:
                    # Rename "Min." to "Min. Teilnehmer"
                    df = df.rename(columns={"Min.": "Min. Teilnehmer"})
                
                if self.scheduler.load_companies(df):
                    self.companies_status.config(
                        text=f"Imported: {os.path.basename(companies_file)}",
                    )
                    cols = [
                        "Unternehmen",
                        "Fachrichtung",
                        "Max. Teilnehmer",
                        "Min. Teilnehmer",
                        "Frühester Zeitpunkt",
                    ]
                    self.setup_preview_tree(self.companies_preview, cols)
                    self.update_preview(self.companies_preview, df, cols)
                    print("Companies imported successfully")
                else:
                    print("Failed to import companies: Invalid format")
            else:
                print(f"Companies file not found: {companies_file}")
                
            # Import preferences
            preferences_file = os.path.join(self.import_folder, os.getenv("STUDENT_PREFERENCES"))
            if os.path.exists(preferences_file):
                print(f"Loading preferences from {preferences_file}")
                df = pd.read_excel(preferences_file)
                df.columns = df.columns.str.strip()
                if self.scheduler.load_student_preferences(df):
                    self.preferences_status.config(
                        text=f"Imported: {os.path.basename(preferences_file)}",
                    )
                    cols = ["Klasse", "Name", "Vorname"] + [
                        f"Wahl {i}" for i in range(1, 7)
                    ]
                    self.setup_preview_tree(self.preferences_preview, cols)
                    self.update_preview(self.preferences_preview, df, cols)
                    print("Preferences imported successfully")
                else:
                    print("Failed to import preferences: Invalid format")
            else:
                print(f"Preferences file not found: {preferences_file}")
                
        except Exception as e:
            print(f"Auto-import error: {str(e)}")
            self.show_error(f"Auto-import error: {str(e)}")

    def setup_ui(self):
        """Set up the user interface"""
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        self.import_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.import_frame, text="Daten importieren")

        self.import_canvas = tk.Canvas(self.import_frame, highlightthickness=0)
        self.import_scrollbar = ttk.Scrollbar(
            self.import_frame,
            orient="vertical",
            command=self.import_canvas.yview,
        )

        self.import_canvas.configure(yscrollcommand=self.import_scrollbar.set)

        self.import_sections = ttk.Frame(self.import_canvas)

        self.import_canvas_window = self.import_canvas.create_window(
            (0, 0),
            window=self.import_sections,
            anchor="nw",
            width=self.import_canvas.winfo_width(),
        )

        self.import_canvas.grid(row=0, column=0, sticky="nsew", padx=(20, 0), pady=20)
        self.import_scrollbar.grid(row=0, column=1, sticky="ns", pady=20)

        self.import_frame.columnconfigure(0, weight=1)
        self.import_frame.rowconfigure(0, weight=1)

        self.import_canvas.bind("<Configure>", self._on_canvas_configure)
        self.import_sections.bind("<Configure>", self._on_frame_configure)
        self.import_canvas.bind_all(
            "<MouseWheel>", lambda e: self._on_mousewheel(e, self.import_canvas)
        )

        # Student wishes section
        section_frame = ttk.Frame(self.import_sections)
        section_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 20))

        ttk.Label(section_frame, text="Schülerwünsche").grid(
            row=0, column=0, columnspan=2, pady=(0, 10), sticky="w"
        )

        import_btn = ttk.Button(
            section_frame,
            text="Import",
            command=self.import_preferences,
        )
        import_btn.grid(row=1, column=0, pady=2, padx=(0, 10), sticky="w")

        self.preferences_status = ttk.Label(
            section_frame,
            text="Noch keine Excel Datei importiert",
        )
        self.preferences_status.grid(row=1, column=1, pady=2, sticky="w")

        # frame for preview
        preview_frame = ttk.Frame(section_frame)
        preview_frame.grid(row=2, column=0, columnspan=2, pady=(5, 0), sticky="nsew")

        self.preferences_preview = ttk.Treeview(
            preview_frame, height=6
        )
        preferences_scrollbar = ttk.Scrollbar(
            preview_frame,
            orient="vertical",
            command=self.preferences_preview.yview,
        )
        self.preferences_preview.configure(yscrollcommand=preferences_scrollbar.set)
        
        # Setup initial empty tree with styling
        self.setup_preview_tree(self.preferences_preview, ["Klasse", "Name", "Vorname", "Wahl 1", "Wahl 2", "Wahl 3", "Wahl 4", "Wahl 5", "Wahl 6"])

        self.preferences_preview.grid(row=0, column=0, sticky="nsew")
        preferences_scrollbar.grid(row=0, column=1, sticky="ns")

        preview_frame.columnconfigure(0, weight=1)
        section_frame.columnconfigure(1, weight=1)

        # Company list
        section_frame = ttk.Frame(self.import_sections)
        section_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 20))

        ttk.Label(section_frame, text="Unternehmensliste").grid(
            row=0, column=0, columnspan=2, pady=(0, 10), sticky="w"
        )

        import_btn = ttk.Button(
            section_frame,
            text="Import",
            command=self.import_companies,
        )
        import_btn.grid(row=1, column=0, pady=2, padx=(0, 10), sticky="w")

        self.companies_status = ttk.Label(
            section_frame,
            text="No file imported",
        )
        self.companies_status.grid(row=1, column=1, pady=2, sticky="w")

        # frame for preview
        preview_frame = ttk.Frame(section_frame)
        preview_frame.grid(row=2, column=0, columnspan=2, pady=(5, 0), sticky="nsew")

        self.companies_preview = ttk.Treeview(preview_frame, height=6)
        companies_scrollbar = ttk.Scrollbar(
            preview_frame,
            orient="vertical",
            command=self.companies_preview.yview,
        )
        self.companies_preview.configure(yscrollcommand=companies_scrollbar.set)
        
        # Setup initial empty tree with styling
        self.setup_preview_tree(self.companies_preview, ["Unternehmen", "Fachrichtung", "Max. Teilnehmer", "Min. Teilnehmer", "Frühester Zeitpunkt"])

        self.companies_preview.grid(row=0, column=0, sticky="nsew")
        companies_scrollbar.grid(row=0, column=1, sticky="ns")

        preview_frame.columnconfigure(0, weight=1)
        section_frame.columnconfigure(1, weight=1)

        # Room list
        section_frame = ttk.Frame(self.import_sections)
        section_frame.grid(row=2, column=0, sticky="nsew")

        ttk.Label(section_frame, text="Raumliste").grid(
            row=0, column=0, columnspan=2, pady=(0, 10), sticky="w"
        )

        import_btn = ttk.Button(
            section_frame,
            text="Import",
            command=self.import_rooms,
        )
        import_btn.grid(row=1, column=0, pady=2, padx=(0, 10), sticky="w")

        self.rooms_status = ttk.Label(
            section_frame,
            text="No file imported",
        )
        self.rooms_status.grid(row=1, column=1, pady=2, sticky="w")

        # frame for preview
        preview_frame = ttk.Frame(section_frame)
        preview_frame.grid(row=2, column=0, columnspan=2, pady=(5, 0), sticky="nsew")

        self.rooms_preview = ttk.Treeview(preview_frame, height=6)
        rooms_scrollbar = ttk.Scrollbar(
            preview_frame,
            orient="vertical",
            command=self.rooms_preview.yview,
        )
        self.rooms_preview.configure(yscrollcommand=rooms_scrollbar.set)
        
        # Setup initial empty tree with styling
        self.setup_preview_tree(self.rooms_preview, ["Raum"])

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
        self.schedule_controls.grid(row=0, column=0, sticky="ew", padx=15, pady=(15, 5))

        ttk.Button(
            self.schedule_controls,
            text="Zeitplan generieren",
            command=self.generate_schedule,
        ).grid(row=0, column=0, padx=5)

        ttk.Button(
            self.schedule_controls,
            text="Zeitplan exportieren",
            command=self.export_schedule,
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
            lambda e: self._on_mousewheel(e, self.student_preview_canvas),
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

        # Attendance Lists tab
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
            lambda e: self._on_mousewheel(e, self.attendance_preview_canvas),
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

        # export frame grid
        self.export_frame.columnconfigure(0, weight=1)
        self.export_frame.rowconfigure(0, weight=1)

        # grid settings
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(0, weight=1)

    def setup_styles(self):
        """Set up ttk styles for the entire application"""
        style = ttk.Style()
        
        # Set up Treeview colors for both dark and light mode
        if self.is_dark_mode():
            # Dark mode
            style.configure("Treeview", background="#2d2d2d", fieldbackground="#2d2d2d", foreground="white")
            style.configure("Treeview.Heading", background="#3f3f3f", foreground="white")
            style.map('Treeview', background=[('selected', '#4a6984')], foreground=[('selected', 'white')])
            
            # Tag configs for alternating rows
            style.map('Treeview', foreground=[])  # Reset the map
            style.configure("Treeview", rowheight=25)
            
            # Define tag styles directly
            self.oddrow_bg = "#3f3f3f"
            self.evenrow_bg = "#2d2d2d"
        else:
            # Light mode
            style.configure("Treeview", background="white", fieldbackground="white", foreground="black")
            style.configure("Treeview.Heading", background="#f2f2f2", foreground="black")
            style.map('Treeview', background=[('selected', '#3584e4')], foreground=[('selected', 'white')])
            
            # Tag configs for alternating rows
            style.map('Treeview', foreground=[])  # Reset the map
            style.configure("Treeview", rowheight=25)
            
            # Define tag styles directly
            self.oddrow_bg = "#f2f2f2"
            self.evenrow_bg = "white"
            
    def is_dark_mode(self):
        """Detect if system is using dark mode"""
        # Check if background of Frame is dark
        style = ttk.Style()
        bg_color = style.lookup('TFrame', 'background')
        
        # If no background color found, assume light mode
        if not bg_color:
            return False
            
        # Try to detect based on Mac system appearance
        try:
            # macOS specific check
            if sys.platform == "darwin":
                import subprocess
                cmd = "defaults read -g AppleInterfaceStyle"
                result = subprocess.run(cmd, shell=True, text=True, capture_output=True)
                return result.stdout.strip() == "Dark"
        except:
            pass
            
        # Try to detect based on color brightness
        try:
            if bg_color.startswith('#'):
                # Hex color
                rgb = tuple(int(bg_color.lstrip('#')[i:i+2], 16) for i in (0, 2, 4))
                brightness = (0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]) / 255
                return brightness < 0.5
        except:
            pass
            
        return False

    def setup_preview_tree(self, tree, columns):
        tree["columns"] = columns
        tree["show"] = "headings"

        for col in columns:
            tree.heading(col, text=col, anchor="w")  # Alle Überschriften linksbündig
            tree.column(col, anchor="w", width=100)  # Alle Werte linksbündig

        # Apply the tag configurations
        tree.tag_configure("oddrow", background=self.oddrow_bg)
        tree.tag_configure("evenrow", background=self.evenrow_bg)

    def update_preview(self, tree, df, columns):
        for item in tree.get_children():
            tree.delete(item)
        
        for idx, row in df.head(6).iterrows():
            values = []
            for col in columns:
                if col in row and pd.notna(row[col]):
                    value = row[col]
                    if isinstance(value, (int, float)) and value.is_integer():
                        values.append(str(int(value)))  # Ganze Zahl ohne Nachkommastellen
                    else:
                        values.append(str(value))  # Sonst als String belassen
                else:
                    values.append("")
            
            tree.insert(
                "", tk.END, values=values, tags=("evenrow" if idx % 2 == 0 else "oddrow")
            )


    def get_import_file(self, env_key, dialog_title="Select file", auto_mode=False):
        if self.dev_mode:
            filename = os.getenv(env_key)
            if filename:
                filepath = os.path.join(self.import_folder, filename)
                if os.path.exists(filepath):
                    print(f"Using file from import folder: {filepath}")
                    return filepath
                else:
                    print(f"Warning: File not found in import folder")
                    print(f"filename from env: {filename}")
                    print(f"filepath: {filepath}")
                    print(f"import_folder: {self.import_folder}")
                    print(f"exists: {os.path.exists(filepath)}")
                    print(f"files in import folder: {os.listdir(self.import_folder) if os.path.exists(self.import_folder) else 'folder does not exist'}")
                    if auto_mode:
                        return None

        if auto_mode:
            return None
            
        return filedialog.askopenfilename(
            title=dialog_title, filetypes=[("Excel files", "*.xlsx")]
        )

    def import_preferences(self, auto_mode=False):
        file_path = self.get_import_file(
            "STUDENT_PREFERENCES", "Import Student Preferences", auto_mode
        )
        if file_path:
            try:
                df = pd.read_excel(file_path)
                df.columns = df.columns.str.strip()
                if self.scheduler.load_student_preferences(df):
                    self.preferences_status.config(
                        text=f"Imported: {os.path.basename(file_path)}",
                    )
                    cols = ["Klasse", "Name", "Vorname"] + [
                        f"Wahl {i}" for i in range(1, 7)
                    ]
                    self.setup_preview_tree(self.preferences_preview, cols)
                    self.update_preview(self.preferences_preview, df, cols)
                else:
                    self.preferences_status.config(
                        text="Ungültiges Format",
                    )
            except Exception as e:
                self.preferences_status.config(
                    text=f"Error: {str(e)}",
                )

    def import_companies(self, auto_mode=False):
        file_path = self.get_import_file("COMPANY_LIST", "Import Company List", auto_mode)
        if file_path:
            try:
                df = pd.read_excel(file_path)
                df.columns = df.columns.str.strip()
                
                # Handle different column names for minimum participants
                if "Min." in df.columns and "Min. Teilnehmer" not in df.columns:
                    # Rename "Min." to "Min. Teilnehmer"
                    df = df.rename(columns={"Min.": "Min. Teilnehmer"})
                
                if self.scheduler.load_companies(df):
                    self.companies_status.config(
                        text=f"Imported: {os.path.basename(file_path)}",
                    )
                    cols = [
                        "Unternehmen",
                        "Fachrichtung",
                        "Max. Teilnehmer",
                        "Min. Teilnehmer",
                        "Frühester Zeitpunkt",
                    ]
                    self.setup_preview_tree(self.companies_preview, cols)
                    self.update_preview(self.companies_preview, df, cols)
                else:
                    self.companies_status.config(
                        text="Ungültiges Format",
                    )
            except Exception as e:
                self.companies_status.config(
                    text=f"Error: {str(e)}",
                )

    def import_rooms(self, auto_mode=False):
        file_path = self.get_import_file("ROOM_LIST", "Import Room List", auto_mode)
        if file_path:
            try:
                df = pd.read_excel(file_path, header=None)
                if self.scheduler.load_rooms(df):
                    self.rooms_status.config(
                        text=f"Imported: {os.path.basename(file_path)}",
                    )
                    self.setup_preview_tree(self.rooms_preview, ["Raum"])
                    self.update_preview(
                        self.rooms_preview,
                        df.rename(columns={df.columns[0]: "Raum"}),
                        ["Raum"],
                    )
                else:
                    self.rooms_status.config(text="Ungültiges Format", foreground="red")
            except Exception as e:
                self.rooms_status.config(text=f"Error: {str(e)}", foreground="red")

    def generate_schedule(self):
        self.clear_error()
        if not self.scheduler.is_data_loaded():
            self.show_error("Bitte laden Sie zuerst alle Daten (Schülerwünsche, Unternehmen und Räume).")
            return

        if self.scheduler.generate_schedule():
            self.update_schedule_display()
        else:
            self.show_error("Es gab ein Problem bei der Generierung des Zeitplans. Bitte überprüfen Sie die Daten.")

    def update_schedule_display(self):
        for item in self.schedule_tree.get_children():
            self.schedule_tree.delete(item)

        time_slots = [
            ("A", "8:45 – 9:30"),
            ("B", "9:50 – 10:35"),
            ("C", "10:35 – 11:20"),
            ("D", "11:40 – 12:25"),
            ("E", "12:25 – 13:10"),
        ]
        
        # Remove overall erfüllungsscore display
        if hasattr(self, 'overall_score_label'):
            self.overall_score_label.destroy()

        columns = ["Company"] + [slot for slot, _ in time_slots]
        self.schedule_tree["columns"] = columns
        self.schedule_tree.column("#0", width=0, stretch=tk.NO)
        self.schedule_tree.column("Company", anchor=tk.W, width=250)
        self.schedule_tree.heading("Company", text="Unternehmen", anchor=tk.W)

        for i, (slot, time_range) in enumerate(time_slots):
            self.schedule_tree.column(slot, anchor=tk.CENTER, width=150)
            self.schedule_tree.heading(
                slot, text=f"{slot} ({time_range})", anchor=tk.CENTER
            )
            
        # Apply tag configurations
        self.schedule_tree.tag_configure("oddrow", background=self.oddrow_bg)
        self.schedule_tree.tag_configure("evenrow", background=self.evenrow_bg)

        # Get company data from scheduler
        schedule = self.scheduler.get_schedule()
        companies = [c for c in self.scheduler.core.companies if (c.name, -1) not in schedule]

        for idx, company in enumerate(companies):
            row = [company.name]
            for slot_idx, (slot_letter, time_range) in enumerate(time_slots):
                if slot_idx < company.earliest_slot or slot_idx in company.blocked_slots:
                    text = "---"
                else:
                    session = schedule.get((company.name, slot_idx))
                    if session:
                        count = len(session.students)
                        capacity = session.company.capacity
                        
                        sessions_for_company = [s for (c, _), s in schedule.items() if c == company.name]
                        
                        if len(sessions_for_company) > 1:
                            text = f"Raum {session.room}"
                        else:
                            text = f"Raum {session.room}"
                    else:
                        text = "---"
                row.append(text)
            self.schedule_tree.insert("", tk.END, values=row, tags=("evenrow" if idx % 2 == 0 else "oddrow"))

    def export_student_schedules(self):
        self.clear_error()
        if not self.scheduler.get_schedule():
            self.show_error("Bitte erst den Zeitplan generieren!")
            return
        
        filepath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if filepath:
            if self.scheduler.export_student_schedules_pdf(filepath):
                self.clear_error()

    def export_attendance_lists(self):
        self.clear_error()
        if not self.scheduler.get_schedule():
            self.show_error("Bitte erst den Zeitplan generieren!")
            return
        
        filepath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if filepath:
            if self.scheduler.export_attendance_lists_pdf(filepath):
                self.clear_error()

    def preview_student_schedules(self):
        self.clear_error()
        if not self.scheduler.get_schedule():
            self.show_error("Bitte erst den Zeitplan generieren!")
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
        self.clear_error()
        if not self.scheduler.get_schedule():
            self.show_error("Bitte erst den Zeitplan generieren!")
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

    def export_schedule(self):
        self.clear_error()
        if not self.scheduler.get_schedule():
            self.show_error("Bitte erst den Zeitplan generieren!")
            return

        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import mm
            from reportlab.platypus import (
                SimpleDocTemplate,
                Table,
                TableStyle,
                Paragraph,
            )

            # Get file path from user
            filepath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if not filepath:
                return

            # Create PDF
            doc = SimpleDocTemplate(
                filepath,
                pagesize=landscape(A4),
                rightMargin=10 * mm,
                leftMargin=10 * mm,
                topMargin=10 * mm,
                bottomMargin=10 * mm,
            )

            story = []
            styles = getSampleStyleSheet()

            # Add title
            title_style = ParagraphStyle(
                "CustomTitle", parent=styles["Heading1"], fontSize=16, spaceAfter=20
            )
            story.append(Paragraph("Zeitplan Übersicht", title_style))

            # Prepare table data
            time_slots = self.scheduler.time_slots
            headers = ["Unternehmen"] + [
                f"{slot} ({time})" for slot, time in time_slots
            ]
            table_data = [headers]

            # Get all wish counts to determine total interest
            all_wish_counts = {}
            for student in self.scheduler.student_preferences:
                for wish in student.wishes:
                    if not wish:
                        continue
                    try:
                        wish_num = int(float(str(wish).strip()))
                        company_name = str(wish_num)
                        for company in self.scheduler.companies:
                            if str(wish_num) == str(company.name.strip()):
                                company_name = company.name.strip()
                                break
                    except (ValueError, TypeError):
                        company_name = str(wish).strip()
                    all_wish_counts[company_name] = all_wish_counts.get(company_name, 0) + 1
            
            # Group sessions by company
            company_sessions = {}
            for (company_name, slot_idx), session in self.scheduler.schedule.items():
                if company_name not in company_sessions:
                    company_sessions[company_name] = []
                company_sessions[company_name].append((slot_idx, session))
            
            for company in self.scheduler.companies:
                row = [company.name]
                company_name = company.name.strip()
                total_interest = all_wish_counts.get(company_name, 0)
                
                for slot_idx, _ in enumerate(time_slots):
                    if slot_idx < company.earliest_slot:
                        text = "---"
                    else:
                        session = self.scheduler.schedule.get((company.name, slot_idx))
                        if session:
                            count = len(session.students)
                            capacity = session.company.capacity
                            
                            # Check if this company has multiple sessions
                            sessions_for_company = [s for (c, _), s in self.scheduler.schedule.items() if c == company.name]
                            
                            if len(sessions_for_company) > 1:
                                # For companies with multiple sessions, show the actual count
                                # We'll rely on the scheduler to distribute students evenly
                                text = f"Raum {session.room}"
                            else:
                                # For companies with a single session, show the actual count
                                text = f"Raum {session.room}"
                        else:
                            text = "---"
                    row.append(text)
                table_data.append(row)

            # Style the table
            col_widths = [40 * mm] + [30 * mm] * len(time_slots)
            t = Table(table_data, colWidths=col_widths, repeatRows=1)
            t.setStyle(
                TableStyle(
                    [
                        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                        ("FONTSIZE", (0, 0), (-1, 0), 10),
                        ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
                        ("BACKGROUND", (0, 1), (-1, -1), colors.white),
                        ("TEXTCOLOR", (0, 1), (-1, -1), colors.black),
                        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                        ("FONTSIZE", (0, 1), (-1, -1), 9),
                        ("ALIGN", (0, 1), (-1, -1), "CENTER"),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("GRID", (0, 0), (-1, -1), 1, colors.black),
                        ("BOX", (0, 0), (-1, -1), 2, colors.black),
                        ("LINEBELOW", (0, 0), (-1, 0), 2, colors.black),
                    ]
                )
            )

            story.append(t)
            doc.build(story)
            
            self.clear_error()

        except Exception as e:
            self.show_error(f"Fehler beim Exportieren des Zeitplans: {str(e)}")

    def _on_mousewheel(self, event, canvas):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_canvas_configure(self, event):
        # Update the scroll region when the canvas is resized
        self.import_canvas.itemconfig(self.import_canvas_window, width=event.width)

    def _on_frame_configure(self, event):
        # scrolling
        self.import_canvas.configure(scrollregion=self.import_canvas.bbox("all"))

    def show_error(self, message):
        """Display an error message in the UI"""
        self.error_display.show(message)
        
    def clear_error(self):
        """Clear the error message"""
        self.error_display.clear()


if __name__ == "__main__":
    root = tk.Tk()
    app = RoomManagementApp(root)
    root.mainloop()
