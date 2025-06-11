import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xlwings as xw
import os
import threading
import shutil
import logging
from debuglog import show_debug_log, TkinterLogHandler

logger = logging.getLogger(__name__)

# Attach TkinterLogHandler to root logger for GUI debuglog (only once)
if not any(isinstance(h, TkinterLogHandler) for h in logging.getLogger().handlers):
    log_handler = TkinterLogHandler()
    log_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
    logging.getLogger().addHandler(log_handler)
    logging.getLogger().setLevel(logging.INFO)

class ExcelWriterWindow(tk.Toplevel):
    """
    A Tkinter Toplevel window that provides a GUI to write data into an Excel template,
    with options to select an Excel file, input various data sections, and write to the workbook
    using xlwings. Supports password-protected sheets and checkboxes handling.
    """

    def __init__(self, master=None):
        """
        Initialize the ExcelWriterWindow UI components, variables, and event bindings.

        Args:
            master (tk.Widget, optional): Parent widget. Defaults to None.
        """
        super().__init__(master)
        self.title("ExcelWriter")
        self.geometry("800x900")
        self.attributes('-topmost', False)

        logger.info("ExcelWriterWindow initialized")

        # Frame for template path + browse + password
        path_pass_frame = ttk.Frame(self)
        path_pass_frame.pack(pady=5, fill='x', padx=10)

        ttk.Label(path_pass_frame, text="Select Excel Template:").grid(row=0, column=0, sticky='w')
        self.template_path_var = tk.StringVar()
        ttk.Entry(path_pass_frame, textvariable=self.template_path_var, width=50).grid(row=1, column=0, sticky='w', padx=(0,5))
        ttk.Button(path_pass_frame, text="Browse", command=self.browse_template).grid(row=1, column=1, sticky='w')
        ttk.Label(path_pass_frame, text="Password:").grid(row=0, column=2, sticky='w', padx=(20,0))
        self.password_var = tk.StringVar()
        ttk.Entry(path_pass_frame, textvariable=self.password_var, show='*', width=20).grid(row=1, column=2, sticky='w', padx=(20,0))

        # Checkboxes for selecting which sections to write
        self.chk_vars = {
            "servers": tk.BooleanVar(value=False),
            "trendstorage": tk.BooleanVar(value=False),
            "cpu": tk.BooleanVar(value=False),
            "memory": tk.BooleanVar(value=False),
            "all_licenses": tk.BooleanVar(value=False)
        }
        chk_frame = ttk.Frame(self)
        chk_frame.pack(pady=5, padx=10, fill='x')
        chk_labels = [
            ("Write Servers data", "servers"),
            ("Write TrendStorage Geheugen data", "trendstorage"),
            ("Write CPU data", "cpu"),
            ("Write Memory data", "memory"),
            ("Write all 'Alle Benodigde licenties aanwezig' (C35 checkboxes)", "all_licenses")
        ]
        for i, (label, key) in enumerate(chk_labels):
            ttk.Checkbutton(chk_frame, text=label, variable=self.chk_vars[key]).grid(row=i, column=0, sticky='w')

        # Text areas for data input
        self.text_areas = {}
        for label in ["Servers", "TrendStorage Geheugen", "CPU", "Memory"]:
            self._add_text_area(label)

        # Configure progress bar style and widget (initially hidden)
        style = ttk.Style(self)
        style.theme_use(style.theme_use())  # Use current theme, do not force 'default'
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor=style.lookup('TFrame', 'background'),
            background='#4a90e2',
            thickness=20
        )
        self.progress = ttk.Progressbar(
            self,
            mode='indeterminate',
            style="Custom.Horizontal.TProgressbar"
        )
        self.progress.pack(pady=10, fill='x', padx=20)
        self.progress.pack_forget()

        # Button to start Excel writing
        self.write_button = ttk.Button(self, text="Write to Excel", command=self.start_edit_excel_thread)
        self.write_button.pack(pady=10)

        self.bind('<Alt-l>', lambda e: show_debug_log(self))

        logger.info("ExcelWriterWindow UI setup complete")

    def _add_text_area(self, label_text):
        """
        Add a labeled text widget for user data input.

        Args:
            label_text (str): The label text for the text area.
        """
        ttk.Label(self, text=label_text + ":").pack(pady=5)
        text_area = tk.Text(self, height=6, width=80)
        text_area.pack(pady=5, padx=10)
        self.text_areas[label_text.lower()] = text_area
        logger.debug(f"Text area added for '{label_text}'")

    def browse_template(self):
        """
        Open a file dialog for the user to select an Excel template file.
        Sets the chosen path to the template_path_var variable.
        """
        logger.info("Browse template dialog opened")
        path = filedialog.askopenfilename(
            parent=self,
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")]
        )
        if path:
            self.template_path_var.set(path)
            logger.info(f"Template selected: {path}")
        else:
            logger.info("No template selected")

    def start_edit_excel_thread(self):
        """
        Starts the Excel editing operation in a separate daemon thread,
        disables the write button, and shows the progress bar.
        """
        logger.info("Starting Excel edit thread")
        self.progress.pack(pady=10, fill='x', padx=20)
        self.progress.start(10)
        self.write_button.config(state='disabled')
        threading.Thread(target=self.edit_excel, daemon=True).start()

    def edit_excel(self):
        """
        Main logic to process inputs and write data into the Excel workbook.
        Runs in a background thread to keep the UI responsive.
        """
        logger.info("edit_excel started")
        path = self.template_path_var.get()
        if not os.path.isfile(path):
            logger.error("Invalid Excel file path")
            self.after(0, self._handle_error, "Please select a valid Excel file.")
            return

        password = self.password_var.get()
        logger.debug(f"Using password: {'*' * len(password)}")

        # Helper to extract non-empty lines from text areas
        get_lines = lambda key: [
            line.strip() for line in self.text_areas[key].get("1.0", tk.END).strip().splitlines() if line.strip()
        ]

        # Extract data lines only if respective checkbox is selected
        servers_lines = get_lines("servers") if self.chk_vars["servers"].get() else []
        trendstorage_lines = get_lines("trendstorage geheugen") if self.chk_vars["trendstorage"].get() else []
        cpu_lines = get_lines("cpu") if self.chk_vars["cpu"].get() else []
        memory_lines = get_lines("memory") if self.chk_vars["memory"].get() else []

        logger.info(f"Lines to process - Servers: {len(servers_lines)}, TrendStorage: {len(trendstorage_lines)}, CPU: {len(cpu_lines)}, Memory: {len(memory_lines)}")

        if not (servers_lines or trendstorage_lines or cpu_lines or memory_lines or self.chk_vars["all_licenses"].get()):
            logger.warning("No valid lines found or no sections selected to write")
            self.after(0, self._handle_error, "No valid lines found or no sections selected to write.")
            return

        # Prompt user for save option (overwrite or new file)
        save_option_event = threading.Event()
        self.overwrite_original = None
        def ask_save_option():
            logger.info("Prompting user for save option (overwrite or copy)")
            self.overwrite_original = messagebox.askyesnocancel(
                "Save Option",
                "Do you want to update the original template?\n\nYes: Overwrite original\nNo: Create a new edited copy"
            )
            save_option_event.set()
        self.after(0, ask_save_option)
        save_option_event.wait()

        if not self.overwrite_original:
            save_path_event = threading.Event()
            self.new_save_path = None
            def ask_save_path():
                logger.info("Prompting user for save location")
                self.new_save_path = filedialog.asksaveasfilename(
                    parent=self,
                    defaultextension=".xlsm",
                    filetypes=[("Excel Macro-Enabled Workbook", "*.xlsm")],
                    initialfile=os.path.basename(path)
                )
                save_path_event.set()
            self.after(0, ask_save_path)
            save_path_event.wait()

            if not self.new_save_path:
                logger.info("Save cancelled by user")
                self.after(0, self._cancel_save)
                return

            try:
                shutil.copy2(path, self.new_save_path)
                edit_path = self.new_save_path
                logger.info(f"Copied template to new file: {self.new_save_path}")
            except Exception as e:
                logger.error(f"Failed to copy template file: {e}")
                self.after(0, self._handle_error, f"Failed to copy file:\n{e}")
                return
        else:
            edit_path = path
            logger.info("User chose to overwrite original template")

        warning_triggered = False

        try:
            logger.info(f"Opening workbook: {edit_path}")
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False
            wb = app.books.open(edit_path)

            def get_sheet(idx):
                sheet_name = "Checklist Regelkast" if idx == 0 else f"Checklist Regelkast ({idx + 1})"
                try:
                    return wb.sheets[sheet_name], sheet_name
                except Exception as e:
                    logger.warning(f"Sheet not found: {sheet_name} - {e}")
                    return None, sheet_name

            def unprotect(sheet):
                try:
                    sheet.api.Unprotect(Password=password)
                    logger.debug(f"Unprotected sheet {sheet.name}")
                    return True
                except Exception as e:
                    logger.error(f"Could not unprotect sheet {sheet.name}: {e}")
                    self.after(0, self._handle_error, f"Failed to unprotect sheet '{sheet.name}'.\nPlease check the password or sheet protection.")
                    raise RuntimeError(f"Failed to unprotect sheet '{sheet.name}'")

            def protect(sheet):
                try:
                    sheet.api.Protect(Password=password)
                    logger.debug(f"Protected sheet {sheet.name}")
                except Exception as e:
                    logger.warning(f"Could not protect sheet {sheet.name}: {e}")

            def set_checkbox(sheet, checkbox_name, sheet_name):
                try:
                    shape = sheet.api.Shapes(checkbox_name)
                    shape.OLEFormat.Object.Value = True
                    logger.info(f"Checked {checkbox_name} on {sheet_name}")
                    return True
                except Exception as e:
                    logger.warning(f"Failed to check {checkbox_name} on {sheet_name}: {e}")
                    return False

            def process_lines(lines, cell, log_prefix, checkboxes, warn_threshold=None, value_func=None):
                """
                Process a list of lines, write data to Excel cells, and check checkboxes conditionally.

                Args:
                    lines (list): Lines of data to write.
                    cell (str): Cell address to write to.
                    log_prefix (str): Prefix for logging.
                    checkboxes (list): Checkbox names to update.
                    warn_threshold (float, optional): Threshold for warnings, if applicable.
                    value_func (callable, optional): Function to transform line value before writing.

                Returns:
                    bool: True if processing completed, False if unprotect failed.
                """
                nonlocal warning_triggered
                for idx, line in enumerate(lines):
                    logger.debug(f"Processing line {idx+1} for {log_prefix}: {line}")
                    sheet, sheet_name = get_sheet(idx)
                    if not sheet:
                        continue
                    try:
                        unprotect(sheet)
                    except RuntimeError:
                        return False
                    value = value_func(line) if value_func else line
                    try:
                        sheet.range(cell).value = value
                        logger.info(f"[{log_prefix}] Written '{value}' to {sheet_name} {cell}")
                    except Exception as e:
                        logger.error(f"[{log_prefix}] Failed to write to {sheet_name} {cell}: {e}")
                    # Checkbox logic
                    if warn_threshold is not None:
                        try:
                            percent = float(value.split('%')[0].split()[-1])
                        except Exception:
                            percent = 0
                        if percent >= warn_threshold:
                            warning_triggered = True
                            for chk in checkboxes:
                                set_checkbox(sheet, chk, sheet_name)
                    else:
                        for chk in checkboxes:
                            set_checkbox(sheet, chk, sheet_name)
                    protect(sheet)
                return True

            # Servers data processing
            if servers_lines:
                if not process_lines(
                    servers_lines,
                    "F8",
                    "Servers",
                    ["Check Box 33"]
                ):
                    return

            # TrendStorage Geheugen data processing
            if trendstorage_lines:
                if not process_lines(
                    trendstorage_lines,
                    "F8",
                    "TrendStorage Geheugen",
                    ["Check Box 37"]
                ):
                    return

            # CPU data processing (percent values)
            if cpu_lines:
                if not process_lines(
                    cpu_lines,
                    "F8",
                    "CPU",
                    ["Check Box 39"],
                    warn_threshold=75,
                    value_func=lambda line: f"{line.strip()}%"
                ):
                    return

            # Memory data processing (percent values)
            if memory_lines:
                if not process_lines(
                    memory_lines,
                    "F8",
                    "Memory",
                    ["Check Box 41"],
                    warn_threshold=75,
                    value_func=lambda line: f"{line.strip()}%"
                ):
                    return

            # 'Alle Benodigde licenties aanwezig' checkboxes (C35)
            if self.chk_vars["all_licenses"].get():
                for idx in range(6):
                    sheet, sheet_name = get_sheet(idx)
                    if not sheet:
                        continue
                    try:
                        unprotect(sheet)
                    except RuntimeError:
                        return
                    for cb_idx in range(35, 41):
                        cb_name = f"Check Box {cb_idx}"
                        set_checkbox(sheet, cb_name, sheet_name)
                    protect(sheet)

            wb.save()
            logger.info(f"Workbook saved successfully at {edit_path}")
            wb.close()
            app.quit()

            if warning_triggered:
                self.after(0, lambda: messagebox.showwarning(
                    "Warning",
                    "Some values exceeded 80%. Please review the checklist."
                ))

            self.after(0, lambda: messagebox.showinfo("Success", "Excel writing completed successfully."))

        except Exception as e:
            logger.exception("Exception during Excel writing")
            self.after(0, self._handle_error, f"An error occurred:\n{e}")
        finally:
            self.after(0, self._finalize_ui)

    def _handle_error(self, message):
        """
        Show an error message box and reset UI components.

        Args:
            message (str): Error message to display.
        """
        messagebox.showerror("Error", message)
        self._finalize_ui()

    def _cancel_save(self):
        """
        Handle user cancellation during save dialog.
        """
        messagebox.showinfo("Cancelled", "Excel writing cancelled.")
        self._finalize_ui()

    def _finalize_ui(self):
        """
        Reset UI elements after Excel operation is completed or aborted.
        """
        self.progress.stop()
        self.progress.pack_forget()
        self.write_button.config(state='normal')
        logger.info("UI reset to ready state")
