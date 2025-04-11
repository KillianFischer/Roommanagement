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

        self.export_notebook = ttk.Notebook(self.export_frame)
        self.export_notebook.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        self._setup_student_schedules_tab()
        self._setup_attendance_lists_tab()
        
        self.export_frame.columnconfigure(0, weight=1)
        self.export_frame.rowconfigure(0, weight=1)

    # error display
    def show_error(self, section, message):
        if section == "student":
            if hasattr(self, 'student_error_label'):
                self.student_error_label.config(text=message)
        elif section == "attendance":
            if hasattr(self, 'attendance_error_label'):
                self.attendance_error_label.config(text=message)

    # error clearing
    def clear_error(self, section=None):
        if section == "student" or section is None:
            if hasattr(self, 'student_error_label'):
                self.student_error_label.config(text="")
        if section == "attendance" or section is None:
            if hasattr(self, 'attendance_error_label'):
                self.attendance_error_label.config(text="")

    # student schedules tab setup
    def _setup_student_schedules_tab(self):
        self.student_schedules_frame = ttk.Frame(self.export_notebook)
        self.export_notebook.add(self.student_schedules_frame, text="Schülerzeitpläne")

        button_frame = ttk.Frame(self.student_schedules_frame)
        button_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        
        ttk.Button(
            button_frame,
            text="Vorschau",
            command=self._preview_student_command,
        ).grid(row=0, column=0, pady=5, padx=5, sticky="w")
        ttk.Button(
            button_frame,
            text="Als Excel exportieren",
            command=self._export_student_excel_command,
        ).grid(row=0, column=1, pady=5, padx=5, sticky="e")
        
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=0)
        
        self.student_error_label = ttk.Label(self.student_schedules_frame, text="", foreground="red", wraplength=800)
        self.student_error_label.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 5))

        self.student_preview_canvas = tk.Canvas(self.student_schedules_frame, width=800)
        self.student_preview_scrollbar = ttk.Scrollbar(
            self.student_schedules_frame,
            orient="vertical",
            command=self.student_preview_canvas.yview,
        )
        self.student_preview_frame = ttk.Frame(self.student_preview_canvas)
        self.student_preview_canvas.configure(yscrollcommand=self.student_preview_scrollbar.set)

        self.student_preview_canvas.bind_all(
            "<MouseWheel>",
            lambda e: self.app._on_mousewheel(e, self.student_preview_canvas),
        )

        self.student_preview_canvas.grid(row=2, column=0, sticky="nsew", padx=5, pady=5)
        self.student_preview_scrollbar.grid(row=2, column=1, sticky="ns", pady=5)
        
        self.student_preview_canvas_window = self.student_preview_canvas.create_window(
            (0, 0), window=self.student_preview_frame, anchor="nw", width=self.student_preview_canvas.winfo_width()
        )

        self.student_schedules_frame.rowconfigure(1, weight=0)
        self.student_schedules_frame.rowconfigure(2, weight=1)
        self.student_schedules_frame.columnconfigure(0, weight=1)
        
        self.student_preview_canvas.bind('<Configure>', self._on_student_canvas_configure)
        
    # student canvas configure handler
    def _on_student_canvas_configure(self, event):
        width = event.width - 10
        if hasattr(self, 'student_preview_canvas_window') and self.student_preview_canvas.winfo_exists() and self.student_preview_canvas.find_all():
             self.student_preview_canvas.itemconfigure(self.student_preview_canvas_window, width=width)
        
    # attendance lists tab setup
    def _setup_attendance_lists_tab(self):
        self.attendance_lists_frame = ttk.Frame(self.export_notebook)
        self.export_notebook.add(self.attendance_lists_frame, text="Anwesenheitslisten")

        button_frame = ttk.Frame(self.attendance_lists_frame)
        button_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        
        ttk.Button(
            button_frame,
            text="Vorschau",
            command=self._preview_attendance_command,
        ).grid(row=0, column=0, pady=5, padx=5, sticky="w")
        ttk.Button(
            button_frame,
            text="Als Excel exportieren",
            command=self._export_attendance_excel_command,
        ).grid(row=0, column=1, pady=5, padx=5, sticky="e")
        
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=0)
        
        self.attendance_error_label = ttk.Label(self.attendance_lists_frame, text="", foreground="red", wraplength=800)
        self.attendance_error_label.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 5))

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

        self.attendance_preview_canvas.bind_all(
            "<MouseWheel>",
            lambda e: self.app._on_mousewheel(e, self.attendance_preview_canvas),
        )

        self.attendance_preview_canvas.grid(
            row=2, column=0, sticky="nsew", padx=5, pady=5
        )
        self.attendance_preview_scrollbar.grid(row=2, column=1, sticky="ns", pady=5)
        
        self.attendance_preview_canvas_window = self.attendance_preview_canvas.create_window(
            (0, 0), window=self.attendance_preview_frame, anchor="nw", width=self.attendance_preview_canvas.winfo_width()
        )

        self.attendance_lists_frame.rowconfigure(1, weight=0)
        self.attendance_lists_frame.rowconfigure(2, weight=1)
        self.attendance_lists_frame.columnconfigure(0, weight=1)
        
        self.attendance_preview_canvas.bind('<Configure>', self._on_attendance_canvas_configure)
        
    # attendance canvas configure handler
    def _on_attendance_canvas_configure(self, event):
        width = event.width - 10
        if hasattr(self, 'attendance_preview_canvas_window') and self.attendance_preview_canvas.winfo_exists() and self.attendance_preview_canvas.find_all():
             self.attendance_preview_canvas.itemconfigure(self.attendance_preview_canvas_window, width=width)

    # student preview command
    def _preview_student_command(self):
        self.clear_error("student")
        if not self.scheduler.get_schedule():
            self.show_error("student", "Bitte erst den Zeitplan generieren.")
            return
        self.preview_student_schedules()

    # student excel export command
    def _export_student_excel_command(self):
        self.clear_error("student")
        if not self.scheduler.get_schedule():
            self.show_error("student", "Bitte erst den Zeitplan generieren.")
            return
        self.export_student_schedules_excel()

    # attendance preview command
    def _preview_attendance_command(self):
        self.clear_error("attendance")
        if not self.scheduler.get_schedule():
            self.show_error("attendance", "Bitte erst den Zeitplan generieren.")
            return
        self.preview_attendance_lists()

    # attendance excel export command
    def _export_attendance_excel_command(self):
        self.clear_error("attendance")
        if not self.scheduler.get_schedule():
            self.show_error("attendance", "Bitte erst den Zeitplan generieren.")
            return
        self.export_attendance_lists_excel()

    # student schedules preview
    def preview_student_schedules(self):
        for widget in self.student_preview_frame.winfo_children():
            widget.destroy()

        self.student_preview_frame.columnconfigure(0, weight=1)
        self.student_preview_frame.columnconfigure(1, weight=4)
        self.student_preview_frame.columnconfigure(2, weight=1)
        self.student_preview_frame.columnconfigure(3, weight=1)

        overall_score = self.scheduler.calculate_overall_fulfillment_score()
        ttk.Label(
            self.student_preview_frame,
            text=f"Gesamter Erfüllungsscore: {overall_score:.1f}%",
            font=("Helvetica", 12, "bold")
        ).grid(row=0, column=0, columnspan=4, pady=(5, 20), sticky="w")

        class_schedules = {}
        
        company_to_number = {}
        for idx, company in enumerate(self.scheduler.core.companies, 1):
            company_to_number[company.name.strip()] = str(idx)
        
        for student_name, appointments in self.scheduler.get_student_schedules().items():
            student = next((s for s in self.scheduler.student_preferences if s.name == student_name), None)
            if student:
                class_name = student.student_id.split("_")[0]
                if class_name not in class_schedules:
                    class_schedules[class_name] = []
                
                schedule_data = []
                for slot_letter, time_range, company, room, wish_number in appointments:
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

        row = 1

        for class_name, students in sorted(class_schedules.items()):
            ttk.Label(
                self.student_preview_frame,
                text=f"Klasse {class_name}",
                font=("Helvetica", 11, "bold")
            ).grid(row=row, column=0, columnspan=4, pady=(20, 10), sticky="w")
            row += 1

            for student in students:
                ttk.Label(
                    self.student_preview_frame,
                    text=f"{student['name']}",
                    font=("Helvetica", 10, "bold")
                ).grid(row=row, column=0, columnspan=4, pady=(10, 5), sticky="w")
                row += 1

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
                    
                    ttk.Label(
                        self.student_preview_frame,
                        text=appointment["company"],
                        wraplength=400
                    ).grid(row=row, column=1, padx=5, pady=2, sticky="w")
                    
                    ttk.Label(
                        self.student_preview_frame,
                        text=appointment["room"],
                    ).grid(row=row, column=2, padx=5, pady=2, sticky="w")
                    
                    wish_label = ttk.Label(
                        self.student_preview_frame,
                        text=str(appointment["wish_number"]),
                    )
                    
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
            
    # attendance lists preview
    def preview_attendance_lists(self):
        self.clear_error("attendance")
        if not self.scheduler.get_schedule():
            self.show_error("attendance", "Bitte erst den Zeitplan generieren.")
            return

        for widget in self.attendance_preview_frame.winfo_children():
            widget.destroy()

        self.attendance_preview_frame.columnconfigure(0, weight=1)
        self.attendance_preview_frame.columnconfigure(1, weight=6)
        self.attendance_preview_frame.columnconfigure(2, weight=2)
        self.attendance_preview_frame.columnconfigure(3, weight=2)

        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_file:
            temp_filepath = temp_file.name
            
        if temp_filepath and self.scheduler.export_attendance_lists_pdf(temp_filepath, preview_mode=True):
            sorted_sessions = sorted(
                self.scheduler.get_schedule().items(),
                key=lambda x: (x[0][0], x[0][1]),
            )

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
                if slot_idx == -1:
                    continue
                    
                slot_letter, time_range = self.scheduler.time_slots[slot_idx]
                
                company_name = session.get_company_display_name()
                
                header_label = ttk.Label(
                    self.attendance_preview_frame,
                    text=f"{company_name}",
                    font=("Helvetica", 11, "bold")
                )
                header_label.grid(row=row, column=0, columnspan=4, pady=(20, 5), sticky="w")
                row += 1

                time_label = ttk.Label(
                    self.attendance_preview_frame,
                    text=f"Zeitfenster: {slot_letter} ({time_range}) - Raum: {session.room}",
                    font=("Helvetica", 10, "italic")
                )
                time_label.grid(row=row, column=0, columnspan=4, pady=(0, 10), sticky="w")
                row += 1

                headers = [("Nr.", 0), ("Name", 1), ("Klasse", 2), ("Anwesend", 3)]
                for header_text, col in headers:
                    header = ttk.Label(
                        self.attendance_preview_frame,
                        text=header_text,
                        font=("Helvetica", 10, "bold")
                    )
                    header.grid(row=row, column=col, padx=5, pady=5, sticky="w")
                row += 1

                if len(session.students) == 0:
                    ttk.Label(
                        self.attendance_preview_frame,
                        text="",
                    ).grid(row=row, column=0, padx=5, pady=2, sticky="w")
                    
                    warning_label = ttk.Label(
                        self.attendance_preview_frame,
                        text="Kein Schülerinteresse",
                        font=("Helvetica", 10, "bold")
                    )
                    try:
                        warning_label.configure(foreground="red")
                    except:
                        pass
                        
                    warning_label.grid(row=row, column=1, columnspan=3, padx=5, pady=10, sticky="w")
                    row += 1
                else:
                    for i, student in enumerate(sorted(session.students, key=lambda x: x["name"]), 1):
                        class_name = student["id"].split("_")[0]
                        
                        ttk.Label(
                            self.attendance_preview_frame,
                            text=str(i),
                        ).grid(row=row, column=0, padx=5, pady=2, sticky="w")
                        
                        ttk.Label(
                            self.attendance_preview_frame,
                            text=student["name"],
                            wraplength=300
                        ).grid(row=row, column=1, padx=5, pady=2, sticky="w")
                        
                        ttk.Label(
                            self.attendance_preview_frame,
                            text=class_name,
                        ).grid(row=row, column=2, padx=5, pady=2, sticky="w")
                        
                        ttk.Label(
                            self.attendance_preview_frame,
                            text="□",
                        ).grid(row=row, column=3, padx=5, pady=2, sticky="w")
                        
                        row += 1

            self.attendance_preview_frame.update_idletasks()
            self.attendance_preview_canvas.configure(
                scrollregion=self.attendance_preview_canvas.bbox("all")
            )
            
            try:
                if os.path.exists(temp_filepath):
                    os.remove(temp_filepath)
            except:
                pass 

    # student schedules pdf export
    def export_student_schedules_pdf(self):
        filepath = filedialog.asksaveasfilename(
            defaultextension=".pdf", 
            filetypes=[("PDF files", "*.pdf")]
        )
        if filepath:
            if self.scheduler.export_student_schedules_pdf(filepath):
                messagebox.showinfo("Export Erfolgreich", f"Schülerzeitpläne PDF exportiert nach {filepath}")

    # student schedules excel export
    def export_student_schedules_excel(self):
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel files", "*.xlsx")]
        )
        if filepath:
            if self.scheduler.export_student_schedules_excel(filepath):
                messagebox.showinfo("Export Erfolgreich", f"Schülerzeitpläne Excel exportiert nach {filepath}")

    # attendance lists pdf export
    def export_attendance_lists_pdf(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if filepath:
            if self.scheduler.export_attendance_lists_pdf(filepath, preview_mode=False):
                messagebox.showinfo("Export Erfolgreich", f"Anwesenheitslisten PDF exportiert nach {filepath}")

    # attendance lists excel export
    def export_attendance_lists_excel(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            if self.scheduler.export_attendance_lists_excel(filepath, preview_mode=False):
                messagebox.showinfo("Export Erfolgreich", f"Anwesenheitslisten Excel exportiert nach {filepath}")

    # room schedule export
    def export_room_schedule(self):
        self.clear_error()
        if not self.scheduler.get_schedule():
            messagebox.showerror("Exportfehler", "Bitte erst den Zeitplan generieren.")
            return
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            if hasattr(self.scheduler, 'export_room_schedule_excel') and self.scheduler.export_room_schedule_excel(filepath): 
                 messagebox.showinfo("Export Erfolgreich", f"Raumplan Excel exportiert nach {filepath}")
            else:
                 messagebox.showwarning("Nicht Implementiert", "Der Excel-Export für den Raumplan ist noch nicht implementiert.")

    # all data export
    def export_all(self):
        self.clear_error()
        if not self.scheduler.get_schedule():
            messagebox.showerror("Exportfehler", "Bitte erst den Zeitplan generieren.")
            return
            
        export_dir = filedialog.askdirectory(title="Wählen Sie einen Speicherort für alle Exporte")
        if not export_dir:
            return
            
        try:
            zip_filepath = self.scheduler.export_all_data(export_dir)
            if zip_filepath:
                self.clear_error()
                messagebox.showinfo("Information", f"Alle Exporte wurden nach {zip_filepath} exportiert.") 
        except Exception as e:
            messagebox.showerror("Error", f"Fehler beim Exportieren aller Dateien: {str(e)}") 

    # room list export
    def export_room_list(self):
        self.clear_error()
        if not self.scheduler.core.rooms:
             messagebox.showerror("Exportfehler", "Bitte zuerst die Raumliste im Import-Tab laden.")
             return
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            if self.scheduler.export_room_list_excel(filepath):
                messagebox.showinfo("Export Erfolgreich", f"Raumliste Excel exportiert nach {filepath}")