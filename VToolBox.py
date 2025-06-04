import tkinter as tk
import logging
from MKDIR import MKDIRApp
from debuglog import show_debug_log
from tkinter import filedialog

logger = logging.getLogger(__name__)

class ToolboxWindow(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("Vincent's Toolbox")
        self.geometry("300x150")

        logger.info("Toolbox module gestart")

        mkdir_button = tk.Button(self, text="Open MKDIR", command=self.open_mkdir_window)
        mkdir_button.pack(pady=40)

        self.bind('<Alt-l>', lambda e: show_debug_log(self))

    def open_mkdir_window(self):
        logger.info("MKDIR tool geopend vanuit Toolbox")
        win = MKDIRApp(self)
        win.lift()
        win.focus_force()

    def some_dialog_method(self):
        result = filedialog.askopenfilename(parent=self)
        # ... your logic ...
        self.lift()
        self.focus_force()