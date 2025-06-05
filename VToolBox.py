import tkinter as tk
from tkinter import ttk
import logging
from MKDIR import MKDIRApp
from Excelwriter import ExcelWriterWindow
from debuglog import show_debug_log
from tkinter import filedialog

logger = logging.getLogger(__name__)

class ToolboxWindow(tk.Toplevel):
    """
    A Tkinter Toplevel window serving as a toolbox with buttons to launch
    various utility modules like MKDIR and ExcelWriter.
    """

    def __init__(self, master=None):
        """
        Initialize the ToolboxWindow with UI buttons and event bindings.

        Args:
            master (tk.Widget, optional): Parent widget. Defaults to None.
        """
        super().__init__(master)
        self.title("Vincent's Toolbox")
        self.geometry("300x150")
        self.attributes('-topmost', False)

        logger.info("Toolbox module gestart")

        mkdir_button = ttk.Button(self, text="Open MKDIR", command=self.open_mkdir_window)
        mkdir_button.pack(pady=10)

        excel_writer_btn = ttk.Button(self, text="ExcelWriter", command=self.open_excel_writer)
        excel_writer_btn.pack(pady=10)

        self.bind('<Alt-l>', lambda e: show_debug_log(self))

    def open_mkdir_window(self):
        """
        Open the MKDIR tool window and bring it to the front.
        """
        logger.info("MKDIR tool geopend vanuit Toolbox")
        win = MKDIRApp(self)
        win.lift()
        win.focus_force()

    def open_excel_writer(self):
        """
        Open the ExcelWriter tool window and bring it to the front.
        """
        logger.info("ExcelWriter tool geopend vanuit Toolbox")
        win = ExcelWriterWindow(self)
        win.lift()
        win.focus_force()
