import tkinter as tk
import logging
from MKDIR import MKDIRApp
from Excelwriter import ExcelWriterWindow
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
        mkdir_button.pack(pady=10)

        excel_writer_btn = tk.Button(self, text="ExcelWriter", command=self.open_excel_writer)
        excel_writer_btn.pack(pady=10)

        self.bind('<Alt-l>', lambda e: show_debug_log(self))

    def open_mkdir_window(self):
        logger.info("MKDIR tool geopend vanuit Toolbox")
        win = MKDIRApp(self)
        win.lift()
        win.focus_force()

    def open_excel_writer(self):
        logger.info("ExcelWriter tool geopend vanuit Toolbox")
        win = ExcelWriterWindow(self)
        win.lift()
        win.focus_force()
