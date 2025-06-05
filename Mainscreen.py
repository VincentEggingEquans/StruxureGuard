import tkinter as tk
from tkinter import ttk
import logging
from VToolBox import ToolboxWindow
from debuglog import show_debug_log, TkinterLogHandler

logger = logging.getLogger(__name__)

class StruxureGuardApp(tk.Tk):
    """
    Main application window for StruxureGuard.

    Inherits from tkinter.Tk and provides the main UI elements,
    including a hidden toolbox button activated by pressing F8,
    and a debug log window accessible with Alt+L.
    """

    def __init__(self):
        """
        Initialize the main application window, setup UI components,
        keybindings, and logging.
        """
        super().__init__()
        self.title("StruxureGuard Startpagina")
        self.geometry("400x300")

        # Create and hide the toolbox button
        self.toolbox_button = ttk.Button(self, text="Vincent's Toolbox", command=self.open_toolbox_window)
        self.toolbox_button.pack_forget()

        # Keybindings
        self.bind('<F8>', self.show_toolbox)
        self.bind('<Alt-l>', lambda e: show_debug_log(self))

        # Setup logging to GUI
        self.log_handler = TkinterLogHandler()
        self.log_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
        logging.getLogger().addHandler(self.log_handler)  # Attach to root logger
        logging.getLogger().setLevel(logging.INFO)        # Ensure INFO level logs are shown

        logger.info("StruxureGuard hoofdvenster gestart")

    def show_toolbox(self, event=None):
        """
        Show the hidden toolbox button when F8 is pressed.

        Args:
            event (tk.Event, optional): The event that triggered the method. Defaults to None.
        """
        self.toolbox_button.pack(pady=20)
        logger.info("F8 ingedrukt - Toolbox knop zichtbaar")

    def open_toolbox_window(self):
        """
        Open the Toolbox window when the toolbox button is clicked.
        """
        logger.info("Toolbox venster geopend vanuit Mainscreen")
        win = ToolboxWindow(self)
        win.lift()
        win.focus_force()

if __name__ == "__main__":
    """
    Run the StruxureGuard application.
    """
    app = StruxureGuardApp()
    app.mainloop()
