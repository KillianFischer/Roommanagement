import tkinter as tk
from tkinter import ttk, filedialog
import os
import pandas as pd


class ImportsTab:
    def __init__(self, parent, scheduler, on_mousewheel, app):
        self.parent = parent
        self.scheduler = scheduler
        self.app = app  # Reference to main app to access methods like show_error
        
        self.import_frame = ttk.Frame(parent)
        parent.add(self.import_frame, text="Daten importieren")

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
            "<MouseWheel>", lambda e: on_mousewheel(e, self.import_canvas)
        )

        # Student wishes section
        self._setup_student_section()
        
        # Company list
        self._setup_company_section()
        
        # Room list
        self._setup_room_section()

        # import sections layout
        self.import_sections.columnconfigure(0, weight=1)

        # import frame layout
        self.import_frame.columnconfigure(0, weight=1)
        self.import_frame.rowconfigure(0, weight=1)
        
    def _setup_student_section(self):
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
        self.app.setup_preview_tree(self.preferences_preview, ["Klasse", "Name", "Vorname", "Wahl 1", "Wahl 2", "Wahl 3", "Wahl 4", "Wahl 5", "Wahl 6"])

        self.preferences_preview.grid(row=0, column=0, sticky="nsew")
        preferences_scrollbar.grid(row=0, column=1, sticky="ns")

        preview_frame.columnconfigure(0, weight=1)
        section_frame.columnconfigure(1, weight=1)
        
    def _setup_company_section(self):
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
        self.app.setup_preview_tree(self.companies_preview, ["Unternehmen", "Fachrichtung", "Max. Teilnehmer", "Min. Teilnehmer", "Frühester Zeitpunkt"])

        self.companies_preview.grid(row=0, column=0, sticky="nsew")
        companies_scrollbar.grid(row=0, column=1, sticky="ns")

        preview_frame.columnconfigure(0, weight=1)
        section_frame.columnconfigure(1, weight=1)
        
    def _setup_room_section(self):
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
        self.app.setup_preview_tree(self.rooms_preview, ["Raum"])

        self.rooms_preview.grid(row=0, column=0, sticky="nsew")
        rooms_scrollbar.grid(row=0, column=1, sticky="ns")

        preview_frame.columnconfigure(0, weight=1)
        section_frame.columnconfigure(1, weight=1)
        
    def _on_canvas_configure(self, event):
        # Update the scroll region when the canvas is resized
        self.import_canvas.itemconfig(self.import_canvas_window, width=event.width)

    def _on_frame_configure(self, event):
        # scrolling
        self.import_canvas.configure(scrollregion=self.import_canvas.bbox("all"))
        
    def get_import_file(self, env_key, dialog_title="Select file", auto_mode=False):
        if self.app.dev_mode:
            filename = os.getenv(env_key)
            if filename:
                filepath = os.path.join(self.app.import_folder, filename)
                if os.path.exists(filepath):
                    print(f"Using file from import folder: {filepath}")
                    return filepath
                else:
                    print(f"Warning: File not found in import folder")
                    print(f"filename from env: {filename}")
                    print(f"filepath: {filepath}")
                    print(f"import_folder: {self.app.import_folder}")
                    print(f"exists: {os.path.exists(filepath)}")
                    print(f"files in import folder: {os.listdir(self.app.import_folder) if os.path.exists(self.app.import_folder) else 'folder does not exist'}")
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
                    self.app.setup_preview_tree(self.preferences_preview, cols)
                    self.app.update_preview(self.preferences_preview, df, cols)
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
                    self.app.setup_preview_tree(self.companies_preview, cols)
                    self.app.update_preview(self.companies_preview, df, cols)
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
                    self.app.setup_preview_tree(self.rooms_preview, ["Raum"])
                    self.app.update_preview(
                        self.rooms_preview,
                        df.rename(columns={df.columns[0]: "Raum"}),
                        ["Raum"],
                    )
                else:
                    self.rooms_status.config(text="Ungültiges Format", foreground="red")
            except Exception as e:
                self.rooms_status.config(text=f"Error: {str(e)}", foreground="red")
                
    def auto_import_files(self):
        """Automatically import files in dev mode if they exist in the import folder"""
        try:
            print("Starting auto import...")
            
            # Import rooms
            rooms_file = os.path.join(self.app.import_folder, os.getenv("ROOM_LIST"))
            if os.path.exists(rooms_file):
                print(f"Loading rooms from {rooms_file}")
                df = pd.read_excel(rooms_file, header=None)
                if self.scheduler.load_rooms(df):
                    self.rooms_status.config(
                        text=f"Imported: {os.path.basename(rooms_file)}",
                    )
                    self.app.setup_preview_tree(self.rooms_preview, ["Raum"])
                    self.app.update_preview(
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
            companies_file = os.path.join(self.app.import_folder, os.getenv("COMPANY_LIST"))
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
                    self.app.setup_preview_tree(self.companies_preview, cols)
                    self.app.update_preview(self.companies_preview, df, cols)
                    print("Companies imported successfully")
                else:
                    print("Failed to import companies: Invalid format")
            else:
                print(f"Companies file not found: {companies_file}")
                
            # Import preferences
            preferences_file = os.path.join(self.app.import_folder, os.getenv("STUDENT_PREFERENCES"))
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
                    self.app.setup_preview_tree(self.preferences_preview, cols)
                    self.app.update_preview(self.preferences_preview, df, cols)
                    print("Preferences imported successfully")
                else:
                    print("Failed to import preferences: Invalid format")
            else:
                print(f"Preferences file not found: {preferences_file}")
                
        except Exception as e:
            print(f"Auto-import error: {str(e)}")
            self.app.show_error(f"Auto-import error: {str(e)}") 