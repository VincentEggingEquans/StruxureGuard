import os
import shutil
import tkinter as tk
import logging
from tkinter import filedialog, messagebox
from tkinter import ttk
from tqdm import tqdm
import threading
from debuglog import show_debug_log, TkinterLogHandler

logger = logging.getLogger(__name__)

# Attach handler only once (recommended in main or top-level window)
if not any(isinstance(h, TkinterLogHandler) for h in logging.getLogger().handlers):
    handler = TkinterLogHandler()
    handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
    logging.getLogger().addHandler(handler)
    logging.getLogger().setLevel(logging.INFO)

class MKDIRApp(tk.Toplevel):
    """
    A Tkinter Toplevel window for creating multiple directories and optionally copying a file into each.

    Features:
    - Select base directory
    - Input folder names (one per line)
    - Option to copy a selected file into each created directory
    - Progress bar to indicate creation progress
    """

    def __init__(self, master=None):
        """
        Initialize the MKDIRApp window with widgets and event bindings.

        Args:
            master (tk.Widget, optional): Parent widget. Defaults to None.
        """
        super().__init__(master)
        self.title("StruxureGuard MKDIR")
        self.attributes('-topmost', False)
        self.geometry("700x450")

        logger.info("MKDIR module gestart")

        self.bind('<Alt-l>', lambda e: show_debug_log(self))

        # Basismap selecteren
        self.base_path = tk.StringVar(value=os.getcwd())  # Default: huidige werkmap

        base_frame = tk.Frame(self)
        base_frame.pack(pady=5, padx=10)
        ttk.Label(base_frame, text="Doelmap:").pack(side='left')
        self.base_entry = tk.Entry(base_frame, textvariable=self.base_path, width=80)
        self.base_entry.pack(side='left', padx=5)
        ttk.Button(base_frame, text="Bladeren", command=self.browse_base_path).pack(side='right')

        # Mapnamen invoer
        self.label = ttk.Label(self, text="Voer mapnamen in (één per regel):")
        self.label.pack(pady=5)

        self.textbox = tk.Text(self, height=12, width=60)
        self.textbox.pack(pady=5)

        self.copy_var = tk.IntVar()
        self.copy_check = ttk.Checkbutton(self, text="Kopieer bestand naar elke map", variable=self.copy_var, command=self.toggle_file_button)
        self.copy_check.pack()

        self.file_frame = ttk.Frame(self)
        self.file_frame.pack(pady=5)

        self.file_button = ttk.Button(self.file_frame, text="Kies bestand", command=self.select_file, state=tk.DISABLED)
        self.file_button.pack(side=tk.LEFT)

        self.selected_file_label = ttk.Label(self.file_frame, text="Geen bestand geselecteerd")
        self.selected_file_label.pack(side=tk.LEFT, padx=5)

        self.progress = ttk.Progressbar(self, orient="horizontal", length=500, mode="determinate")
        self.progress.pack(pady=10)

        self.run_button = ttk.Button(self, text="Start", command=self.run)
        self.run_button.pack(pady=10)

        self.selected_file = None

    def toggle_file_button(self):
        """
        Enable or disable the 'Kies bestand' button based on the checkbox state.
        """
        state = tk.NORMAL if self.copy_var.get() else tk.DISABLED
        self.file_button.config(state=state)

    def select_file(self):
        """
        Open a file dialog for selecting a file to copy into each created directory.

        Updates the selected file label with the chosen file's name.
        """
        filepath = filedialog.askopenfilename(parent=self)
        if filepath:
            self.selected_file = filepath
            self.selected_file_label.config(text=os.path.basename(filepath))
        self.lift()
        self.focus_force()

    def browse_base_path(self):
        """
        Open a directory selection dialog to set the base path for new directories.
        """
        path = filedialog.askdirectory(parent=self)
        if path:
            self.base_path.set(path)
        self.lift()
        self.focus_force()

    def run(self):
        """
        Start the directory creation process in a separate thread.
        """
        thread = threading.Thread(target=self.create_directories)
        thread.start()

    def create_directories(self):
        """
        Create directories as specified in the textbox, optionally copying a file to each.

        Shows progress in the progress bar and displays message boxes for errors or completion.
        """
        names = self.textbox.get("1.0", tk.END).strip().splitlines()
        total = len(names)
        if total == 0:
            messagebox.showerror("Fout", "Geen mapnamen ingevoerd.")
            return

        self.progress["value"] = 0
        self.progress["maximum"] = total

        for i, name in enumerate(tqdm(names, desc="Mappen maken", unit="map")):
            try:
                dir_path = os.path.join(self.base_path.get(), name)
                os.makedirs(dir_path, exist_ok=True)
                if self.copy_var.get() and self.selected_file:
                    ext = os.path.splitext(self.selected_file)[1]
                    new_path = os.path.join(dir_path, f"{name}{ext}")
                    shutil.copy2(self.selected_file, new_path)
            except Exception as e:
                messagebox.showerror("Fout", f"Kon map of bestand niet aanmaken: {e}")
                return
            self.progress["value"] = i + 1
            self.update_idletasks()

        messagebox.showinfo("Klaar", "Alle mappen zijn aangemaakt.")
