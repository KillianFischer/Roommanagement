import tkinter as tk
from tkinter import ttk, filedialog

class ScheduleTab:
    def __init__(self, parent, scheduler, app):
        self.parent = parent
        self.scheduler = scheduler
        self.app = app
        
        self.schedule_frame = ttk.Frame(parent)
        parent.add(self.schedule_frame, text="Zeitplan")

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
        
    def generate_schedule(self):
        self.app.clear_error()
        if not self.scheduler.is_data_loaded():
            self.app.show_error("Bitte laden Sie zuerst alle Daten (Schülerwünsche, Unternehmen und Räume).")
            return

        if self.scheduler.generate_schedule():
            self.update_schedule_display()
        else:
            self.app.show_error("Es gab ein Problem bei der Generierung des Zeitplans. Bitte überprüfen Sie die Daten.")
            
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
        self.schedule_tree.tag_configure("oddrow", background=self.app.oddrow_bg)
        self.schedule_tree.tag_configure("evenrow", background=self.app.evenrow_bg)
        
        # Get company data from scheduler
        schedule = self.scheduler.get_schedule()
        
        # Get all companies with their fields to display properly
        companies = self.scheduler.core.companies
        
        # Group schedule by company ID (which includes field)
        company_sessions = {}
        for (company_id, slot_idx), session in schedule.items():
            if company_id not in company_sessions:
                company_sessions[company_id] = {}
            company_sessions[company_id][slot_idx] = session
        
        # Display companies in the tree
        for idx, company in enumerate(companies):
            # Skip companies that have been excluded
            if any((company.unique_id, -1) in schedule for company in companies):
                continue
                
            # Get the display name with field info
            display_name = str(company)
            
            row = [display_name]
            for slot_idx, (slot_letter, time_range) in enumerate(time_slots):
                if slot_idx < company.earliest_slot or slot_idx in company.blocked_slots:
                    text = "---"
                else:
                    session = schedule.get((company.unique_id, slot_idx))
                    if session:
                        # Show room information
                        text = f"Raum {session.room}"
                    else:
                        text = "---"
                row.append(text)
                
            # Add the row to the tree
            self.schedule_tree.insert("", tk.END, values=row, tags=("evenrow" if idx % 2 == 0 else "oddrow"))
            
    def export_schedule(self):
        self.app.clear_error()
        if not self.scheduler.get_schedule():
            self.app.show_error("Bitte erst den Zeitplan generieren!")
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
                return  # User cancelled the dialog, do nothing

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
            for (company_id, slot_idx), session in self.scheduler.schedule.items():
                if company_id not in company_sessions:
                    company_sessions[company_id] = []
                company_sessions[company_id].append((slot_idx, session))
            
            for company in self.scheduler.companies:
                display_name = str(company)
                row = [display_name]
                
                for slot_idx, _ in enumerate(time_slots):
                    if slot_idx < company.earliest_slot:
                        text = "---"
                    else:
                        session = self.scheduler.schedule.get((company.unique_id, slot_idx))
                        if session:
                            count = len(session.students)
                            capacity = session.company.capacity
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
            
            self.app.clear_error()

        except Exception as e:
            self.app.show_error(f"Fehler beim Exportieren des Zeitplans: {str(e)}") 