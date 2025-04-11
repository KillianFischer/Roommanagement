import tkinter as tk
import os
import sys
from tkinter import ttk
import pandas as pd
from dotenv import load_dotenv
from services.scheduler import Scheduler
from ui import ImportsTab, ScheduleTab, ExportsTab

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

load_dotenv()

class RoomManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Raumverwaltung für Berufsorientierungstag")
        self.root.geometry("1920x1080")

        self.scheduler = Scheduler()

        self.dev_mode = os.getenv("DEV_MODE", "false").lower() == "true"
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.import_folder = os.path.join(base_dir, os.getenv("IMPORT_FOLDER", "import/"))

        if self.dev_mode and not os.path.exists(self.import_folder):
            os.makedirs(self.import_folder)

        self.setup_styles()

        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky="nsew")

        self.setup_ui()

        if self.dev_mode:
            self.imports_tab.auto_import_files()

    # user interface tabs
    def setup_ui(self):
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        self.imports_tab = ImportsTab(self.notebook, self.scheduler, self._on_mousewheel, self)
        self.schedule_tab = ScheduleTab(self.notebook, self.scheduler, self)
        self.exports_tab = ExportsTab(self.notebook, self.scheduler, self._on_mousewheel, self)

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(0, weight=1)

    # visual styles for widgets
    def setup_styles(self):
        style = ttk.Style()

        style.configure("Treeview", background="white", fieldbackground="white", foreground="black")
        style.configure("Treeview.Heading", background="#f2f2f2", foreground="black")
        style.map('Treeview', background=[('selected', '#3584e4')], foreground=[('selected', 'white')])

        style.map('Treeview', foreground=[])
        style.configure("Treeview", rowheight=25)

        self.oddrow_bg = "#f2f2f2"
        self.evenrow_bg = "white"

    # treeview widget for data preview
    def setup_preview_tree(self, tree, columns):
        tree["columns"] = columns
        tree["show"] = "headings"

        for col in columns:
            tree.heading(col, text=col, anchor="w")
            tree.column(col, anchor="w", width=100)

        tree.tag_configure("oddrow", background=self.oddrow_bg)
        tree.tag_configure("evenrow", background=self.evenrow_bg)

    # preview treeview with dataframe data
    def update_preview(self, tree, df, columns):
        for item in tree.get_children():
            tree.delete(item)

        for idx, row in df.head(6).iterrows():
            values = []
            for col in columns:
                if col in row and pd.notna(row[col]):
                    value = row[col]
                    if isinstance(value, (int, float)) and value.is_integer():
                        values.append(str(int(value)))
                    else:
                        values.append(str(value))
                else:
                    values.append("")

            tree.insert(
                "", tk.END, values=values, tags=("evenrow" if idx % 2 == 0 else "oddrow")
            )

    # mouse wheel scrolling on specified canvas
    def _on_mousewheel(self, event, canvas):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

if __name__ == "__main__":
    root = tk.Tk()
    app = RoomManagementApp(root)
    root.mainloop()
