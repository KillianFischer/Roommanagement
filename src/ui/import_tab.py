import tkinter as tk
from tkinter import ttk, filedialog
import os
import pandas as pd


class ImportsTab:
    def __init__(self, parent, scheduler, on_mousewheel, app):
        self.parent = parent
        self.scheduler = scheduler
        self.app = app  # Reference to main app
        
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

        self._setup_student_section()

        self._setup_company_section()
        
        self._setup_room_section()

        self.import_sections.columnconfigure(0, weight=1)

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

        preview_frame = ttk.Frame(section_frame)
        preview_frame.grid(row=2, column=0, columnspan=2, pady=(5, 0), sticky="nsew")

        self.companies_preview = ttk.Treeview(preview_frame, height=6)
        companies_scrollbar = ttk.Scrollbar(
            preview_frame,
            orient="vertical",
            command=self.companies_preview.yview,
        )
        self.companies_preview.configure(yscrollcommand=companies_scrollbar.set)
        
        self.app.setup_preview_tree(self.companies_preview, ["Unternehmen", "Fachrichtung", "Max. Teilnehmer", "Max. Veranstaltungen", "Frühester Zeitpunkt"])

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
        self.app.setup_preview_tree(self.rooms_preview, ["Raum", "Kapazität"])

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
                
                # Check for Min. column being used instead of full name
                if "Min." in df.columns and "Max. Veranstaltungen" not in df.columns:
                    # Rename "Min." to "Max. Veranstaltungen"
                    df = df.rename(columns={"Min.": "Max. Veranstaltungen"})
                
                # For backward compatibility
                if "Min. Teilnehmer" in df.columns and "Max. Veranstaltungen" not in df.columns:
                    # Rename "Min. Teilnehmer" to "Max. Veranstaltungen"
                    df = df.rename(columns={"Min. Teilnehmer": "Max. Veranstaltungen"})

                # Get required columns
                required = [
                    "Unternehmen",
                    "Max. Teilnehmer",
                    "Max. Veranstaltungen",
                    "Frühester Zeitpunkt",
                ]

                if self.scheduler.load_companies(df):
                    self.companies_status.config(
                        text=f"Imported: {os.path.basename(file_path)}",
                    )
                    self.app.setup_preview_tree(self.companies_preview, required)
                    self.app.update_preview(self.companies_preview, df, required)
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
                # Read Excel file, assuming it might not have header row
                df = pd.read_excel(file_path, header=None)
                
                # If the file has headers (first row contains "Raum"), use them
                if len(df.columns) >= 1 and isinstance(df.iloc[0, 0], str) and df.iloc[0, 0].lower() in ["raum", "room"]:
                    # The file has headers - reread with headers
                    df = pd.read_excel(file_path)
                    
                    # Make sure we have the right column names
                    if "Raum" not in df.columns and "Room" in df.columns:
                        df = df.rename(columns={"Room": "Raum"})
                    if "Kapazität" not in df.columns and "Kapazitaet" in df.columns:
                        df = df.rename(columns={"Kapazitaet": "Kapazität"})
                    if "Kapazität" not in df.columns and "Capacity" in df.columns:
                        df = df.rename(columns={"Capacity": "Kapazität"})
                else:
                    # No headers - assign our own column names
                    column_names = ["Raum"]
                    if len(df.columns) >= 2:
                        column_names.append("Kapazität")
                    
                    df.columns = column_names
                
                if self.scheduler.load_rooms(df):
                    self.rooms_status.config(
                        text=f"Imported: {os.path.basename(file_path)}",
                    )
                    
                    # Determine the columns to display in preview
                    preview_columns = ["Raum"]
                    if "Kapazität" in df.columns:
                        preview_columns.append("Kapazität")
                    
                    self.app.setup_preview_tree(self.rooms_preview, preview_columns)
                    self.app.update_preview(self.rooms_preview, df, preview_columns)
                else:
                    self.rooms_status.config(text="Ungültiges Format", foreground="red")
            except Exception as e:
                import traceback
                traceback.print_exc()
                self.rooms_status.config(text=f"Error: {str(e)}", foreground="red")
                
    def auto_import_files(self):
        """Automatically import files in dev mode if they exist in the import folder"""
        try:
            print("Starting auto import...")
            
            # Import rooms
            rooms_file = os.path.join(self.app.import_folder, os.getenv("ROOM_LIST"))
            if os.path.exists(rooms_file):
                print(f"Loading rooms from {rooms_file}")
                # Check if excel has headers
                temp_df = pd.read_excel(rooms_file, header=None)
                
                # If the file has headers (first row contains "Raum"), use them
                if len(temp_df.columns) >= 1 and isinstance(temp_df.iloc[0, 0], str) and temp_df.iloc[0, 0].lower() in ["raum", "room"]:
                    # The file has headers - reread with headers
                    df = pd.read_excel(rooms_file)
                    
                    # Make sure we have the right column names
                    if "Raum" not in df.columns and "Room" in df.columns:
                        df = df.rename(columns={"Room": "Raum"})
                    if "Kapazität" not in df.columns and "Kapazitaet" in df.columns:
                        df = df.rename(columns={"Kapazitaet": "Kapazität"})
                    if "Kapazität" not in df.columns and "Capacity" in df.columns:
                        df = df.rename(columns={"Capacity": "Kapazität"})
                else:
                    # No headers - assign our own column names
                    df = temp_df
                    column_names = ["Raum"]
                    if len(df.columns) >= 2:
                        column_names.append("Kapazität")
                    
                    df.columns = column_names
                
                if self.scheduler.load_rooms(df):
                    self.rooms_status.config(
                        text=f"Imported: {os.path.basename(rooms_file)}",
                    )
                    
                    # Determine the columns to display in preview
                    preview_columns = ["Raum"]
                    if "Kapazität" in df.columns:
                        preview_columns.append("Kapazität")
                    
                    self.app.setup_preview_tree(self.rooms_preview, preview_columns)
                    self.app.update_preview(self.rooms_preview, df, preview_columns)
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
                
                # Check for Min. column being used instead of full name
                if "Min." in df.columns and "Max. Veranstaltungen" not in df.columns:
                    # Rename "Min." to "Max. Veranstaltungen"
                    df = df.rename(columns={"Min.": "Max. Veranstaltungen"})
                
                # For backward compatibility
                if "Min. Teilnehmer" in df.columns and "Max. Veranstaltungen" not in df.columns:
                    # Rename "Min. Teilnehmer" to "Max. Veranstaltungen"
                    df = df.rename(columns={"Min. Teilnehmer": "Max. Veranstaltungen"})

                # Get required columns
                required = [
                    "Unternehmen",
                    "Max. Teilnehmer",
                    "Max. Veranstaltungen",
                    "Frühester Zeitpunkt",
                ]

                if self.scheduler.load_companies(df):
                    self.companies_status.config(
                        text=f"Imported: {os.path.basename(companies_file)}",
                    )
                    self.app.setup_preview_tree(self.companies_preview, required)
                    self.app.update_preview(self.companies_preview, df, required)
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