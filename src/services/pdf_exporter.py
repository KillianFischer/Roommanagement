from typing import List, Dict, Tuple, Optional
import os
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

from models.student import StudentPreference


class PDFExporter:
    def __init__(self):
        self.styles = getSampleStyleSheet()
        # Create custom styles
        self.styles.add(
            ParagraphStyle(
                name="CustomTitle",
                fontName="Helvetica-Bold",
                fontSize=16,
                alignment=1,
                spaceAfter=12,
            )
        )
        self.styles.add(
            ParagraphStyle(
                name="CustomSubtitle",
                fontName="Helvetica-Bold",
                fontSize=14,
                alignment=0,
                spaceAfter=6,
            )
        )
        self.styles.add(
            ParagraphStyle(
                name="CustomNormal",
                fontName="Helvetica",
                fontSize=10,
                alignment=0,
                spaceAfter=6,
            )
        )

    def export_student_schedules(self, filepath: str, schedule: Dict[Tuple[str, int], any], 
                                student_preferences: List[StudentPreference], time_slots: List[Tuple[str, str]]):
        """
        Export student schedules to a PDF file
        
        Args:
            filepath: The path to save the PDF to
            schedule: The schedule data (company name, slot) -> session
            student_preferences: List of student preferences
            time_slots: List of time slots as (letter, time range)
        """
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        
        doc = SimpleDocTemplate(filepath, pagesize=A4, 
                               rightMargin=30, leftMargin=30,
                               topMargin=30, bottomMargin=30)
        
        content = []
        
        content.append(Paragraph("Schüler Zeitpläne", self.styles["CustomTitle"]))
        content.append(Paragraph(f"Erstellt am: {datetime.now().strftime('%d.%m.%Y %H:%M')}", 
                                self.styles["CustomNormal"]))
        content.append(Spacer(1, 12))
        
        student_schedules = self._prepare_student_schedules(schedule, student_preferences)
        
        sorted_students = sorted(student_schedules.keys(), key=lambda x: x.lower())
        
        for student_name in sorted_students:
            content.append(Paragraph(f"Schüler: {student_name}", self.styles["CustomSubtitle"]))
            
            data = []
            
            headers = ["Slot", "Zeit", "Unternehmen", "Raum"]
            data.append(headers)
            
            student_data = student_schedules[student_name]
            
            for slot_idx, (slot_letter, time_range) in enumerate(time_slots):
                row = [slot_letter, time_range]
                
                if slot_idx in student_data:
                    company_name, room = student_data[slot_idx]
                    row.append(company_name)
                    row.append(room)
                else:
                    row.append("-")
                    row.append("-")
                
                data.append(row)
                
            table = Table(data, colWidths=[30, 100, 200, 100])
            
            table_style = TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 10),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
                ("BACKGROUND", (0, 1), (-1, -1), colors.beige),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 1), (-1, -1), 9),
                ("ALIGN", (0, 1), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ])
            
            table.setStyle(table_style)
            content.append(table)
            content.append(Spacer(1, 12))
            
        doc.build(content)
            
    def _prepare_student_schedules(self, schedule, student_preferences):
        """
        Create a dictionary mapping student names to their schedules
        
        Returns:
            dict: student_name -> {slot_idx: (company_name, room)}
        """
        student_schedules = {}
        
        for student in student_preferences:
            student_schedules[student.name] = {}
            
        for (company_name, slot_idx), session in schedule.items():
            if slot_idx == -1:
                continue  # Skip excluded companies
                
            for student in session.students:
                student_name = student["name"]
                if student_name not in student_schedules:
                    student_schedules[student_name] = {}
                    
                student_schedules[student_name][slot_idx] = (company_name, session.room)
                
        return student_schedules
        
    def export_company_overview(self, filepath: str, schedule: Dict[Tuple[str, int], any],
                               time_slots: List[Tuple[str, str]]):
        """
        Export company overview to a PDF file
        
        Args:
            filepath: The path to save the PDF to
            schedule: The schedule data (company name, slot) -> session
            time_slots: List of time slots as (letter, time range)
        """
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        
        doc = SimpleDocTemplate(filepath, pagesize=A4, 
                               rightMargin=30, leftMargin=30,
                               topMargin=30, bottomMargin=30)
        
        content = []
        
        content.append(Paragraph("Unternehmensübersicht", self.styles["CustomTitle"]))
        content.append(Paragraph(f"Erstellt am: {datetime.now().strftime('%d.%m.%Y %H:%M')}", 
                                self.styles["CustomNormal"]))
        content.append(Spacer(1, 12))
        
        slot_to_companies = {}
        for (company_name, slot_idx), session in schedule.items():
            if slot_idx == -1:
                continue  # Skip excluded companies
                
            if slot_idx not in slot_to_companies:
                slot_to_companies[slot_idx] = []
                
            slot_to_companies[slot_idx].append((company_name, session.room, len(session.students)))
            
        for slot_idx, (slot_letter, time_range) in enumerate(time_slots):
            if slot_idx not in slot_to_companies:
                continue
                
            content.append(Paragraph(f"Slot {slot_letter}: {time_range}", self.styles["CustomSubtitle"]))
            
            data = []
            
            headers = ["Unternehmen", "Raum", "Anzahl Schüler"]
            data.append(headers)
            
            for company_name, room, student_count in sorted(slot_to_companies[slot_idx], 
                                                          key=lambda x: x[0].lower()):
                data.append([company_name, room, student_count])
                
            table = Table(data, colWidths=[200, 100, 100])
            
            table_style = TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 10),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
                ("BACKGROUND", (0, 1), (-1, -1), colors.beige),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 1), (-1, -1), 9),
                ("ALIGN", (0, 1), (-1, -1), "LEFT"),
                ("ALIGN", (2, 1), (2, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ])
            
            table.setStyle(table_style)
            content.append(table)
            content.append(Spacer(1, 12))
            
        excluded_companies = []
        for (company_name, slot_idx), session in schedule.items():
            if slot_idx == -1:
                excluded_companies.append(company_name)
                
        if excluded_companies:
            content.append(Paragraph("Ausgeschlossene Unternehmen", self.styles["CustomSubtitle"]))
            
            for company_name in sorted(excluded_companies):
                content.append(Paragraph(f"• {company_name}: Kein Schülerinteresse", 
                                       self.styles["CustomNormal"]))
            
        doc.build(content) 