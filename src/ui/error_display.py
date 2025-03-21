import tkinter as tk
from tkinter import ttk

class ErrorDisplay:
    def __init__(self, parent_frame, row=0, column=0, columnspan=1, pady=(5, 0)):
        self.error_frame = ttk.Frame(parent_frame)
        self.error_frame.grid(row=row, column=column, sticky="ew", pady=pady, columnspan=columnspan)
        self.error_frame.grid_remove()
        
        self.error_label = ttk.Label(
            self.error_frame, 
            text="", 
            foreground="red", 
            wraplength=400
        )
        self.error_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        
        close_button = ttk.Button(
            self.error_frame, 
            text="×", 
            width=2, 
            command=self.clear
        )
        close_button.grid(row=0, column=1, sticky="e", padx=5)
    
    def show(self, message):
        error_map = {
            "Invalid input": "Ungültige Eingabe",
            "Field required": "Feld erforderlich",
            "Invalid email": "Ungültige E-Mail",
            "Password too short": "Passwort zu kurz",
            "Incorrect password": "Falsches Passwort",
            "User not found": "Nutzer nicht gefunden",
            "Server error": "Serverfehler",
            "Connection error": "Verbindungsfehler"
        }
        
        short_message = error_map.get(message, message)
        if not message in error_map and len(message) > 50:
            short_message = message[:47] + "..."
            
        self.error_label.config(text=short_message)
        self.error_frame.grid()
    
    def clear(self):
        self.error_label.config(text="")
        self.error_frame.grid_remove()