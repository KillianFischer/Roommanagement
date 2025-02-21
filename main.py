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

        # Initialize the Scheduler instance
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
        self.notebook.add(self.import_frame, text="Import Data")

        # Student wishes
        ttk.Label(self.import_frame, text="Schülerwünsche", font=("Helvetica", 10, "bold")).grid(row=0, column=0, columnspan=2, pady=5, sticky=tk.W)
        ttk.Button(self.import_frame, text="Import", command=self.import_preferences).grid(row=1, column=0, pady=2, padx=5)
        self.preferences_status = ttk.Label(self.import_frame, text="No file imported", foreground="gray")
        self.preferences_status.grid(row=1, column=1, pady=2, padx=5, sticky=tk.W)
        self.preferences_preview = ttk.Treeview(self.import_frame, height=3)
        self.preferences_preview.grid(row=2, column=0, columnspan=2, pady=5, sticky="nsew")

        # Company list
        ttk.Label(self.import_frame, text="Unternehmensliste", font=("Helvetica", 10, "bold")).grid(row=3, column=0, columnspan=2, pady=5, sticky=tk.W)
        ttk.Button(self.import_frame, text="Import", command=self.import_companies).grid(row=4, column=0, pady=2, padx=5)
        self.companies_status = ttk.Label(self.import_frame, text="No file imported", foreground="gray")
        self.companies_status.grid(row=4, column=1, pady=2, padx=5, sticky=tk.W)
        self.companies_preview = ttk.Treeview(self.import_frame, height=3)
        self.companies_preview.grid(row=5, column=0, columnspan=2, pady=5, sticky="nsew")

        # Room list
        ttk.Label(self.import_frame, text="Raumliste", font=("Helvetica", 10, "bold")).grid(row=6, column=0, columnspan=2, pady=5, sticky=tk.W)
        ttk.Button(self.import_frame, text="Import", command=self.import_rooms).grid(row=7, column=0, pady=2, padx=5)
        self.rooms_status = ttk.Label(self.import_frame, text="No file imported", foreground="gray")
        self.rooms_status.grid(row=7, column=1, pady=2, padx=5, sticky=tk.W)
        self.rooms_preview = ttk.Treeview(self.import_frame, height=3)
        self.rooms_preview.grid(row=8, column=0, columnspan=2, pady=5, sticky="nsew")

        self.import_frame.columnconfigure(1, weight=1)

        # Schedule Tab
        self.schedule_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.schedule_frame, text="Schedule")
        ttk.Button(self.schedule_frame, text="Generate Schedule", command=self.generate_schedule).grid(row=0, column=0, pady=15, padx=15)
        self.schedule_tree = ttk.Treeview(self.schedule_frame)
        self.schedule_tree.grid(row=1, column=0, sticky="nsew", pady=15, padx=15)
        self.schedule_frame.rowconfigure(1, weight=1)
        self.schedule_frame.columnconfigure(0, weight=1)

        # Export Tab
        self.export_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.export_frame, text="Export")
        ttk.Button(self.export_frame, text="Export Student Schedules", command=self.export_student_schedules).grid(row=0, column=0, pady=5, padx=5)
        ttk.Button(self.export_frame, text="Export Attendance Lists", command=self.export_attendance_lists).grid(row=1, column=0, pady=5, padx=5)

        # Grid settings
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(0, weight=1)

    def setup_preview_tree(self, tree, columns):
        tree['columns'] = columns
        tree.column('#0', width=0, stretch=tk.NO)
        for col in columns:
            tree.column(col, anchor=tk.W, width=100)
            tree.heading(col, text=col, anchor=tk.W)

    def update_preview(self, tree, df, columns):
        for item in tree.get_children():
            tree.delete(item)
        for _, row in df.head(3).iterrows():
            values = [str(row[col]) if col in row and pd.notna(row[col]) else '' for col in columns]
            tree.insert('', tk.END, values=values)

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
        self.scheduler.export_attendance_lists()

if __name__ == "__main__":
    root = tk.Tk()
    app = RoomManagementApp(root)
    root.mainloop()
