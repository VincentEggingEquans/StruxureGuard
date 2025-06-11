import tkinter as tk
from tkinter import ttk
import logging
from MKDIR import MKDIRApp
from Excelwriter import ExcelWriterWindow
from debuglog import show_debug_log, TkinterLogHandler

logger = logging.getLogger(__name__)

# Attach handler only once (recommended in main or top-level window)
if not any(isinstance(h, TkinterLogHandler) for h in logging.getLogger().handlers):
    handler = TkinterLogHandler()
    handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
    logging.getLogger().addHandler(handler)
    logging.getLogger().setLevel(logging.INFO)

class ToolboxWindow(tk.Toplevel):
    """
    Toolbox window with buttons to launch utility modules like MKDIR and ExcelWriter.
    Prevents multiple instances of each tool.
    """

    def __init__(self, master=None):
        """
        Initialize the ToolboxWindow with UI buttons and event bindings.

        Args:
            master (tk.Widget, optional): Parent widget. Defaults to None.
        """
        super().__init__(master)
        self.title("Vincent's Toolbox")
        self.geometry("320x180")
        self.attributes('-topmost', False)
        self.resizable(False, False)

        logger.info("Toolbox module gestart")

        self.mkdir_window = None
        self.excel_writer_window = None

        ttk.Label(self, text="Selecteer een tool:", font=("TkDefaultFont", 11, "bold")).pack(pady=(15, 8))

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=5)

        self.mkdir_btn = ttk.Button(btn_frame, text="Open MKDIR", width=20, command=self.open_mkdir_window)
        self.mkdir_btn.grid(row=0, column=0, padx=10, pady=5)

        self.excel_btn = ttk.Button(btn_frame, text="ExcelWriter", width=20, command=self.open_excel_writer)
        self.excel_btn.grid(row=1, column=0, padx=10, pady=5)

        ttk.Button(self, text="Sluiten", command=self.destroy).pack(pady=(10, 0))

        self.bind('<Alt-l>', lambda e: show_debug_log(self))
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def open_mkdir_window(self):
        """
        Open the MKDIR tool window and bring it to the front.
        """
        if self.mkdir_window is None or not tk.Toplevel.winfo_exists(self.mkdir_window):
            logger.info("MKDIR tool geopend vanuit Toolbox")
            self.mkdir_window = MKDIRApp(self)
            self.mkdir_window.lift()
            self.mkdir_window.focus_force()
            self.mkdir_window.protocol("WM_DELETE_WINDOW", self._on_mkdir_close)
        else:
            self.mkdir_window.lift()
            self.mkdir_window.focus_force()

    def open_excel_writer(self):
        """
        Open the ExcelWriter tool window and bring it to the front.
        """
        if self.excel_writer_window is None or not tk.Toplevel.winfo_exists(self.excel_writer_window):
            logger.info("ExcelWriter tool geopend vanuit Toolbox")
            self.excel_writer_window = ExcelWriterWindow(self)
            self.excel_writer_window.lift()
            self.excel_writer_window.focus_force()
            self.excel_writer_window.protocol("WM_DELETE_WINDOW", self._on_excel_close)
        else:
            self.excel_writer_window.lift()
            self.excel_writer_window.focus_force()

    def _on_mkdir_close(self):
        """
        Handle the close event of the MKDIR window.
        """
        if self.mkdir_window:
            self.mkdir_window.destroy()
            self.mkdir_window = None

    def _on_excel_close(self):
        """
        Handle the close event of the ExcelWriter window.
        """
        if self.excel_writer_window:
            self.excel_writer_window.destroy()
            self.excel_writer_window = None

    def _on_close(self):
        """
        Handle the close event of the toolbox window.
        """
        if self.mkdir_window and tk.Toplevel.winfo_exists(self.mkdir_window):
            self.mkdir_window.destroy()
        if self.excel_writer_window and tk.Toplevel.winfo_exists(self.excel_writer_window):
            self.excel_writer_window.destroy()
        self.destroy()
