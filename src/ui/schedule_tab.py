import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading

class ScheduleTab:
    def __init__(self, parent, scheduler, app):
        self.parent = parent
        self.scheduler = scheduler
        self.app = app
        
        self.schedule_frame = ttk.Frame(parent)
        parent.add(self.schedule_frame, text="Zeitplan")

        self.schedule_controls = ttk.Frame(self.schedule_frame)
        self.schedule_controls.grid(row=0, column=0, sticky="ew", padx=15, pady=(15, 0))

        ttk.Button(
            self.schedule_controls,
            text="Zeitplan generieren",
            command=self.confirm_and_generate_schedule,
        ).grid(row=0, column=0, padx=5, pady=(0, 5), sticky="w")

        ttk.Button(
            self.schedule_controls,
            text="Zeitplan exportieren",
            command=self.export_schedule_excel,
        ).grid(row=0, column=1, padx=5, pady=(0, 5), sticky="e")

        self.schedule_controls.columnconfigure(0, weight=1)
        self.schedule_controls.columnconfigure(1, weight=0)

        self.schedule_error_label = ttk.Label(self.schedule_frame, text="", foreground="red", wraplength=1180)
        self.schedule_error_label.grid(row=1, column=0, sticky="ew", padx=15, pady=(0, 5))

        self.schedule_frame_inner = ttk.Frame(self.schedule_frame)
        self.schedule_frame_inner.grid(row=2, column=0, sticky="nsew", padx=15, pady=(0, 15))

        self.schedule_scrollbar = ttk.Scrollbar(self.schedule_frame_inner)
        self.schedule_scrollbar.grid(row=0, column=1, sticky="ns")

        self.schedule_tree = ttk.Treeview(
            self.schedule_frame_inner,
            yscrollcommand=self.schedule_scrollbar.set,
        )
        self.schedule_tree.grid(row=0, column=0, sticky="nsew")

        self.schedule_scrollbar.config(command=self.schedule_tree.yview)

        self.schedule_frame.columnconfigure(0, weight=1)
        self.schedule_frame.rowconfigure(1, weight=0)
        self.schedule_frame.rowconfigure(2, weight=1)
        self.schedule_frame_inner.columnconfigure(0, weight=1)
        self.schedule_frame_inner.rowconfigure(0, weight=1)
        
    # error message display
    def show_error(self, message):
        self.schedule_error_label.config(text=message)

    # error message clearing
    def clear_error(self):
        self.schedule_error_label.config(text="")
        
    # generation confirmation popup
    def confirm_and_generate_schedule(self):
        proceed = messagebox.askokcancel(
            "Zeitplan Generierung",
            "Die Zeitplanerstellung dauert bis zu 30 Minuten.\nFortfahren?"
        )
        if proceed:
            self.show_error("Generiere Zeitplan... Bitte warten")
            thread = threading.Thread(target=self.generate_schedule, daemon=True)
            thread.start()

    # background schedule generation
    def generate_schedule(self):
        if not self.scheduler.is_data_loaded():
            self.schedule_frame.after(0, self.show_error, "Bitte laden Sie zuerst alle Daten (Schüler, Unternehmen, Räume) im Import-Tab.")
            return

        success = self.scheduler.generate_schedule()

        self.schedule_frame.after(0, self._update_ui_after_generation, success)

    # post-generation ui update
    def _update_ui_after_generation(self, success):
        if success:
            self.update_schedule_display()
            self.clear_error()
        else:
            self.show_error("Fehler bei der Erstellung des Zeitplans. Details siehe Popup.")

    # schedule display update
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
        
        if hasattr(self, 'overall_score_label'):
            self.overall_score_label.destroy()
            delattr(self, 'overall_score_label')

        columns = ["Company"] + [slot for slot, _ in time_slots]
        self.schedule_tree["columns"] = columns
        self.schedule_tree.column("#0", width=0, stretch=tk.NO)
        self.schedule_tree.column("Company", anchor=tk.W, width=250)
        self.schedule_tree.heading("Company", text="Unternehmen", anchor=tk.W)

        for i, (slot, time_range) in enumerate(time_slots):
            self.schedule_tree.column(slot, anchor=tk.CENTER, width=150)
            self.schedule_tree.heading(slot, text=f"{slot} ({time_range})", anchor=tk.CENTER)
            
        self.schedule_tree.tag_configure("oddrow", background=self.app.oddrow_bg)
        self.schedule_tree.tag_configure("evenrow", background=self.app.evenrow_bg)
        
        schedule = self.scheduler.get_schedule()
        companies = self.scheduler.core.companies
        
        company_sessions = {}
        if schedule: 
            for (company_id, slot_idx), session in schedule.items():
                if company_id not in company_sessions:
                    company_sessions[company_id] = {}
                company_sessions[company_id][slot_idx] = session
        
        for idx, company in enumerate(companies):
            if schedule and any((company.unique_id, -1) == key for key in schedule):
                continue
                
            display_name = str(company)
            row = [display_name]
            for slot_idx, (slot_letter, time_range) in enumerate(time_slots):
                if slot_idx < company.earliest_slot or slot_idx in company.blocked_slots:
                    text = "---"
                else:
                    session = schedule.get((company.unique_id, slot_idx)) if schedule else None
                    if session:
                        text = f"Raum {session.room}"
                    else:
                        text = "---"
                row.append(text)
                
            self.schedule_tree.insert("", tk.END, values=row, tags=("evenrow" if idx % 2 == 0 else "oddrow"))
            
    # schedule pdf export
    def export_schedule_pdf(self):
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

            filepath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if not filepath:
                return

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

            title_style = ParagraphStyle("CustomTitle", parent=styles["Heading1"], fontSize=16, spaceAfter=20)
            story.append(Paragraph("Zeitplan Übersicht", title_style))

            time_slots = self.scheduler.time_slots
            headers = ["Unternehmen"] + [f"{slot} ({time})" for slot, time in time_slots]
            table_data = [headers]

            schedule = self.scheduler.get_schedule()
            companies = self.scheduler.core.companies
            
            for company in companies:
                 # Skip excluded companies
                if schedule and any((company.unique_id, -1) == key for key in schedule):
                    continue

                display_name = str(company)
                row = [display_name]
                for slot_idx, _ in enumerate(time_slots):
                    if slot_idx < company.earliest_slot or slot_idx in company.blocked_slots:
                        text = "---"
                    else:
                        session = schedule.get((company.unique_id, slot_idx)) if schedule else None
                        if session:
                            text = f"Raum {session.room}"
                        else:
                            text = "---"
                    row.append(text)
                table_data.append(row)

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
            messagebox.showinfo("Export Erfolgreich", f"Zeitplan PDF exportiert nach {filepath}")
            
        except ImportError:
            self.show_error("Fehler beim PDF-Export: ReportLab nicht installiert.")
        except Exception as e:
            self.show_error(f"Fehler beim PDF-Export des Zeitplans: {str(e)}")

    # schedule excel export
    def export_schedule_excel(self):
        self.clear_error()
        if not self.scheduler.get_schedule():
            self.show_error("Bitte erst den Zeitplan generieren!")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel files", "*.xlsx"), ("All Files", "*.*")]
        )
        if not filepath:
            return

        try:
            success = self.scheduler.export_schedule_excel(filepath)
            if success:
                self.clear_error()
                messagebox.showinfo("Export Erfolgreich", f"Zeitplan exportiert nach {filepath}")
            else:
                self.show_error("Fehler beim Exportieren des Zeitplans nach Excel. Details siehe Popup.")
            
        except Exception as e:
            self.show_error(f"Unerwarteter Fehler beim Excel-Export des Zeitplans: {str(e)}") 