import tkinter as tk
from tkinter import ttk
import logging
from VToolBox import ToolboxWindow
from debuglog import show_debug_log, TkinterLogHandler
from RapportageGenerator import RapportageGenerator

# Set up root logger and attach TkinterLogHandler only once
logger = logging.getLogger(__name__)
if not any(isinstance(h, TkinterLogHandler) for h in logging.getLogger().handlers):
    handler = TkinterLogHandler()
    handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
    logging.getLogger().addHandler(handler)
    logging.getLogger().setLevel(logging.INFO)

class StruxureGuardApp(tk.Tk):
    """
    Main application window for StruxureGuard.
    """

    def __init__(self):
        """
        Initialize the main application window, setup UI components,
        keybindings, and logging.
        """
        super().__init__()
        self.title("StruxureGuard")
        self.geometry("400x300")

        # Rapportage Generator button
        self.report_button = ttk.Button(self, text="Rapportage Generator", command=self.open_report_generator)
        self.report_button.pack(pady=20)

        # Keybindings
        self.bind('<F8>', self.open_toolbox_window)
        self.bind('<Alt-l>', lambda e: show_debug_log(self))

        logger.info("StruxureGuard hoofdvenster gestart")

    def open_toolbox_window(self, event=None):
        """
        Open the Toolbox window.
        """
        logger.info("Toolbox venster geopend (via F8 of button)")
        win = ToolboxWindow(self)
        win.deiconify()
        win.lift()
        win.focus_force()

    def open_report_generator(self):
        """
        Open de Rapportage Generator.
        """
        logger.info("Rapportage Generator venster geopend")
        win = RapportageGenerator(self)
        win.deiconify()
        win.lift()
        win.focus_force()

if __name__ == "__main__":
    app = StruxureGuardApp()
    app.mainloop()
