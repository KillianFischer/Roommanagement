import tkinter as tk
import os
import sys
from tkinter import ttk
import pandas as pd
from dotenv import load_dotenv
from services.scheduler import Scheduler
from ui import ErrorDisplay, ImportsTab, ScheduleTab, ExportsTab

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

load_dotenv()

class RoomManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Raumverwaltung für Berufsorientierungstag")
        self.root.geometry("1200x800")

        self.scheduler = Scheduler(error_handler=self.show_error)

        # Set up environment variables and paths
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

        # Main application frame
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky="nsew")

        # Error display
        self.error_display = ErrorDisplay(self.main_frame, row=1, column=0)
        
        # Setup UI
        self.setup_ui()
        
        # Auto import in dev mode - Execute directly instead of waiting for timer
        if self.dev_mode:
            print("Dev mode detected, performing auto-import immediately")
            self.imports_tab.auto_import_files()

    def setup_ui(self):
        """Set up the user interface with modular components"""
        # Create main notebook
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        # Create each tab using the modular UI components
        self.imports_tab = ImportsTab(self.notebook, self.scheduler, self._on_mousewheel, self)
        self.schedule_tab = ScheduleTab(self.notebook, self.scheduler, self)
        self.exports_tab = ExportsTab(self.notebook, self.scheduler, self._on_mousewheel, self)

        # Configure grid weights
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
        """Configure a treeview for previewing data"""
        tree["columns"] = columns
        tree["show"] = "headings"

        for col in columns:
            tree.heading(col, text=col, anchor="w")  # Alle Überschriften linksbündig
            tree.column(col, anchor="w", width=100)  # Alle Werte linksbündig

        # Apply the tag configurations
        tree.tag_configure("oddrow", background=self.oddrow_bg)
        tree.tag_configure("evenrow", background=self.evenrow_bg)

    def update_preview(self, tree, df, columns):
        """Update a treeview with dataframe preview"""
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

    def _on_mousewheel(self, event, canvas):
        """Handle mousewheel scrolling for canvas widgets"""
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

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
