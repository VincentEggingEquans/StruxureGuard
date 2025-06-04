# mainscreen.py
import tkinter as tk
import logging
from VToolBox import ToolboxWindow
from debuglog import show_debug_log, TkinterLogHandler

logger = logging.getLogger(__name__)

class StruxureGuardApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("StruxureGuard Startpagina")
        self.geometry("400x300")

        # Create and hide the toolbox button
        self.toolbox_button = tk.Button(self, text="Vincent's Toolbox", command=self.open_toolbox_window)
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
        self.toolbox_button.pack(pady=20)
        logger.info("F8 ingedrukt - Toolbox knop zichtbaar")

    def open_toolbox_window(self):
        logger.info("Toolbox venster geopend vanuit Mainscreen")
        win = ToolboxWindow(self)
        win.lift()
        win.focus_force()

if __name__ == "__main__":
    app = StruxureGuardApp()
    app.mainloop()