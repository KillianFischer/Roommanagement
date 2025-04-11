import tkinter as tk
from tkinter import ttk, filedialog
import os
import pandas as pd


class ImportsTab:
    def __init__(self, parent, scheduler, on_mousewheel, app):
        self.parent = parent
        self.scheduler = scheduler
        self.app = app
        
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
        
    # student section setup
    def _setup_student_section(self):
        section_frame = ttk.Frame(self.import_sections)
        section_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 20))

        ttk.Label(section_frame, text="Schülerwünsche").grid(
            row=0, column=0, columnspan=2, pady=(0, 10), sticky="w"
        )
        ttk.Button(
            section_frame,
            text="Import",
            command=self.import_preferences,
        ).grid(row=1, column=0, pady=2, padx=(0, 10), sticky="w")
        self.preferences_status = ttk.Label(
            section_frame,
            text="Noch keine Excel Datei importiert",
        )
        self.preferences_status.grid(row=1, column=1, pady=2, sticky="w")
        self.preferences_error_label = ttk.Label(section_frame, text="", foreground="red", wraplength=800)
        self.preferences_error_label.grid(row=2, column=0, columnspan=2, pady=(2, 5), sticky="w")

        preview_frame = ttk.Frame(section_frame)
        preview_frame.grid(row=3, column=0, columnspan=2, pady=(5, 0), sticky="nsew")
        self.preferences_preview = ttk.Treeview(preview_frame, height=6)
        preferences_scrollbar = ttk.Scrollbar(
            preview_frame,
            orient="vertical",
            command=self.preferences_preview.yview,
        )
        self.preferences_preview.configure(yscrollcommand=preferences_scrollbar.set)
        self.app.setup_preview_tree(self.preferences_preview, ["Klasse", "Nachname", "Vorname", "Wahl 1", "Wahl 2", "Wahl 3", "Wahl 4", "Wahl 5", "Wahl 6"])
        self.preferences_preview.grid(row=0, column=0, sticky="nsew")
        preferences_scrollbar.grid(row=0, column=1, sticky="ns")
        preview_frame.columnconfigure(0, weight=1)
        section_frame.columnconfigure(1, weight=1)
        
    # company section setup
    def _setup_company_section(self):
        section_frame = ttk.Frame(self.import_sections)
        section_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 20))

        ttk.Label(section_frame, text="Unternehmensliste").grid(
            row=0, column=0, columnspan=2, pady=(0, 10), sticky="w"
        )
        ttk.Button(
            section_frame,
            text="Import",
            command=self.import_companies,
        ).grid(row=1, column=0, pady=2, padx=(0, 10), sticky="w")
        self.companies_status = ttk.Label(
            section_frame,
            text="Keine Datei importiert",
        )
        self.companies_status.grid(row=1, column=1, pady=2, sticky="w")
        self.companies_error_label = ttk.Label(section_frame, text="", foreground="red", wraplength=800)
        self.companies_error_label.grid(row=2, column=0, columnspan=2, pady=(2, 5), sticky="w")

        preview_frame = ttk.Frame(section_frame)
        preview_frame.grid(row=3, column=0, columnspan=2, pady=(5, 0), sticky="nsew")
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
        
    # room section setup
    def _setup_room_section(self):
        section_frame = ttk.Frame(self.import_sections)
        section_frame.grid(row=2, column=0, sticky="nsew")

        ttk.Label(section_frame, text="Raumliste").grid(
            row=0, column=0, columnspan=2, pady=(0, 10), sticky="w"
        )
        ttk.Button(
            section_frame,
            text="Import",
            command=self.import_rooms,
        ).grid(row=1, column=0, pady=2, padx=(0, 10), sticky="w")
        self.rooms_status = ttk.Label(
            section_frame,
            text="Keine Datei importiert",
        )
        self.rooms_status.grid(row=1, column=1, pady=2, sticky="w")
        self.rooms_error_label = ttk.Label(section_frame, text="", foreground="red", wraplength=800)
        self.rooms_error_label.grid(row=2, column=0, columnspan=2, pady=(2, 5), sticky="w")

        preview_frame = ttk.Frame(section_frame)
        preview_frame.grid(row=3, column=0, columnspan=2, pady=(5, 0), sticky="nsew")
        self.rooms_preview = ttk.Treeview(preview_frame, height=6)
        rooms_scrollbar = ttk.Scrollbar(
            preview_frame,
            orient="vertical",
            command=self.rooms_preview.yview,
        )
        self.rooms_preview.configure(yscrollcommand=rooms_scrollbar.set)
        self.app.setup_preview_tree(self.rooms_preview, ["Raum", "Kapazität"])
        self.rooms_preview.grid(row=0, column=0, sticky="nsew")
        rooms_scrollbar.grid(row=0, column=1, sticky="ns")
        preview_frame.columnconfigure(0, weight=1)
        section_frame.columnconfigure(1, weight=1)
        
    # canvas configure handler
    def _on_canvas_configure(self, event):
        self.import_canvas.itemconfig(self.import_canvas_window, width=event.width)

    # frame configure handler
    def _on_frame_configure(self, event):
        self.import_canvas.configure(scrollregion=self.import_canvas.bbox("all"))
        
    # import file getter
    def get_import_file(self, env_key, dialog_title="Select file", auto_mode=False):
        if self.app.dev_mode:
            filename = os.getenv(env_key)
            if filename:
                filepath = os.path.join(self.app.import_folder, filename)
                if os.path.exists(filepath):
                    return filepath
                else:
                    if auto_mode:
                        return None

        if auto_mode:
            return None
            
        return filedialog.askopenfilename(
            title=dialog_title, filetypes=[("Excel files", "*.xlsx")]
        )

    # error display
    def show_error(self, section, message):
        if section == "preferences":
            self.preferences_error_label.config(text=message)
        elif section == "companies":
            self.companies_error_label.config(text=message)
        elif section == "rooms":
            self.rooms_error_label.config(text=message)

    # error clearing
    def clear_error(self, section=None):
        if section == "preferences" or section is None:
            self.preferences_error_label.config(text="")
        if section == "companies" or section is None:
            self.companies_error_label.config(text="")
        if section == "rooms" or section is None:
            self.rooms_error_label.config(text="")

    # preferences import
    def import_preferences(self, auto_mode=False):
        self.clear_error("preferences")
        file_path = self.get_import_file(
            "STUDENT_PREFERENCES", "Import Student Preferences", auto_mode
        )
        if file_path:
            try:
                df = pd.read_excel(file_path)
                df.columns = df.columns.str.strip()
                if self.scheduler.load_student_preferences(df):
                    self.preferences_status.config(
                        text=f"Importiert: {os.path.basename(file_path)}",
                    )
                    cols = ["Klasse", "Nachname", "Vorname"] + [
                        f"Wahl {i}" for i in range(1, 7)
                    ]
                    self.app.setup_preview_tree(self.preferences_preview, cols)
                    self.app.update_preview(self.preferences_preview, df, cols)
                    self.clear_error("preferences")
                else:
                    self.preferences_status.config(text="Ungültiges Format")
                    self.show_error("preferences", "Fehler beim Import: Ungültiges Dateiformat. Details siehe Popup.")
            except Exception as e:
                self.show_error("preferences", f"Fehler beim Lesen der Datei: {str(e)}")
                self.preferences_status.config(text=f"Error: {str(e)}")

    # companies import
    def import_companies(self, auto_mode=False):
        self.clear_error("companies")
        file_path = self.get_import_file("COMPANY_LIST", "Import Company List", auto_mode)
        if file_path:
            try:
                df = pd.read_excel(file_path)
                df.columns = df.columns.str.strip()
                
                if "Min." in df.columns and "Max. Veranstaltungen" not in df.columns:
                    df = df.rename(columns={"Min.": "Max. Veranstaltungen"})
                if "Min. Teilnehmer" in df.columns and "Max. Veranstaltungen" not in df.columns:
                    df = df.rename(columns={"Min. Teilnehmer": "Max. Veranstaltungen"})

                required = [
                    "Unternehmen",
                    "Max. Teilnehmer",
                    "Max. Veranstaltungen",
                    "Frühester Zeitpunkt",
                ]

                if self.scheduler.load_companies(df):
                    self.companies_status.config(
                        text=f"Importiert: {os.path.basename(file_path)}",
                    )
                    self.app.setup_preview_tree(self.companies_preview, required)
                    self.app.update_preview(self.companies_preview, df, required)
                    self.clear_error("companies")
                else:
                    self.companies_status.config(text="Ungültiges Format")
                    self.show_error("companies", "Fehler beim Import: Ungültiges Dateiformat. Details siehe Popup.")
            except Exception as e:
                self.show_error("companies", f"Fehler beim Lesen der Datei: {str(e)}")
                self.companies_status.config(text=f"Error: {str(e)}")

    # rooms import
    def import_rooms(self, auto_mode=False):
        self.clear_error("rooms")
        file_path = self.get_import_file("ROOM_LIST", "Import Room List", auto_mode)
        if file_path:
            try:
                df = pd.read_excel(file_path, header=None)
                if len(df.columns) >= 1 and isinstance(df.iloc[0, 0], str) and df.iloc[0, 0].lower() in ["raum", "room"]:
                    df = pd.read_excel(file_path)
                    if "Raum" not in df.columns and "Room" in df.columns: df = df.rename(columns={"Room": "Raum"})
                    if "Kapazität" not in df.columns and "Kapazitaet" in df.columns: df = df.rename(columns={"Kapazitaet": "Kapazität"})
                    if "Kapazität" not in df.columns and "Capacity" in df.columns: df = df.rename(columns={"Capacity": "Kapazität"})
                else:
                    column_names = ["Raum"]
                    if len(df.columns) >= 2:
                        column_names.append("Kapazität")
                    df.columns = column_names
                
                if self.scheduler.load_rooms(df):
                    self.rooms_status.config(text=f"Importiert: {os.path.basename(file_path)}")
                    preview_columns = ["Raum"]
                    if "Kapazität" in df.columns: preview_columns.append("Kapazität")
                    self.app.setup_preview_tree(self.rooms_preview, preview_columns)
                    self.app.update_preview(self.rooms_preview, df, preview_columns)
                    self.clear_error("rooms")
                else:
                    self.rooms_status.config(text="Ungültiges Format", foreground="red")
                    self.show_error("rooms", "Fehler beim Import: Ungültiges Dateiformat. Details siehe Popup.")
            except Exception as e:
                import traceback
                traceback.print_exc()
                self.show_error("rooms", f"Fehler beim Lesen der Datei: {str(e)}")
                self.rooms_status.config(text=f"Error: {str(e)}", foreground="red")
                
    # automatic file import
    def auto_import_files(self):
        self.clear_error()
        try:
            rooms_file = os.path.join(self.app.import_folder, os.getenv("ROOM_LIST"))
            if os.path.exists(rooms_file):
                temp_df = pd.read_excel(rooms_file, header=None)
                if len(temp_df.columns) >= 1 and isinstance(temp_df.iloc[0, 0], str) and temp_df.iloc[0, 0].lower() in ["raum", "room"]:
                    df = pd.read_excel(rooms_file)
                    if "Raum" not in df.columns and "Room" in df.columns: df = df.rename(columns={"Room": "Raum"})
                    if "Kapazität" not in df.columns and "Kapazitaet" in df.columns: df = df.rename(columns={"Kapazitaet": "Kapazität"})
                    if "Kapazität" not in df.columns and "Capacity" in df.columns: df = df.rename(columns={"Capacity": "Kapazität"})
                else:
                    df = temp_df
                    column_names = ["Raum"]
                    if len(df.columns) >= 2: column_names.append("Kapazität")
                    df.columns = column_names

                if self.scheduler.load_rooms(df):
                    self.rooms_status.config(text=f"Importiert: {os.path.basename(rooms_file)}")
                    preview_columns = ["Raum"]
                    if "Kapazität" in df.columns: preview_columns.append("Kapazität")
                    self.app.setup_preview_tree(self.rooms_preview, preview_columns)
                    self.app.update_preview(self.rooms_preview, df, preview_columns)
                else:
                    self.show_error("rooms", "Fehler beim Auto Import der Raumliste.")
            
            companies_file = os.path.join(self.app.import_folder, os.getenv("COMPANY_LIST"))
            if os.path.exists(companies_file):
                df = pd.read_excel(companies_file)
                df.columns = df.columns.str.strip()
                if "Min." in df.columns and "Max. Veranstaltungen" not in df.columns:
                    df = df.rename(columns={"Min.": "Max. Veranstaltungen"})
                if "Min. Teilnehmer" in df.columns and "Max. Veranstaltungen" not in df.columns:
                    df = df.rename(columns={"Min. Teilnehmer": "Max. Veranstaltungen"})
                required = ["Unternehmen", "Max. Teilnehmer", "Max. Veranstaltungen", "Frühester Zeitpunkt"]
                if self.scheduler.load_companies(df):
                    self.companies_status.config(text=f"Importiert: {os.path.basename(companies_file)}")
                    self.app.setup_preview_tree(self.companies_preview, required)
                    self.app.update_preview(self.companies_preview, df, required)
                else:
                    self.show_error("companies", "Fehler beim Auto Import der Unternehmensliste.")
                
            preferences_file = os.path.join(self.app.import_folder, os.getenv("STUDENT_PREFERENCES"))
            if os.path.exists(preferences_file):
                df = pd.read_excel(preferences_file)
                df.columns = df.columns.str.strip()
                if self.scheduler.load_student_preferences(df):
                    self.preferences_status.config(text=f"Importiert: {os.path.basename(preferences_file)}")
                    cols = ["Klasse", "Nachname", "Vorname"] + [f"Wahl {i}" for i in range(1, 7)]
                    self.app.setup_preview_tree(self.preferences_preview, cols)
                    self.app.update_preview(self.preferences_preview, df, cols)
                else:
                    self.show_error("preferences", "Fehler beim Auto Import der Schülerwünsche.")
                
        except Exception as e:
            self.show_error("preferences", f"Fehler beim automatischen Import: {str(e)}") 