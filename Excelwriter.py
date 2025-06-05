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
    def __init__(self, master=None):
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

        # Checkboxes
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

        # Text areas
        self.text_areas = {}
        for label in ["Servers", "TrendStorage Geheugen", "CPU", "Memory"]:
            self._add_text_area(label)

        # Progress bar style (only affects the progress bar, not the rest of the window)
        style = ttk.Style(self)
        style.theme_use(style.theme_use())  # Use current theme, do not force 'default'
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor=style.lookup('TFrame', 'background'),  # Match window bg
            background='#4a90e2',  # Blue bar
            thickness=20
        )

        self.progress = ttk.Progressbar(
            self,
            mode='indeterminate',
            style="Custom.Horizontal.TProgressbar"
        )
        self.progress.pack(pady=10, fill='x', padx=20)
        self.progress.pack_forget()

        self.write_button = ttk.Button(self, text="Write to Excel", command=self.start_edit_excel_thread)
        self.write_button.pack(pady=10)

        self.bind('<Alt-l>', lambda e: show_debug_log(self))

        logger.info("ExcelWriterWindow UI setup complete")

    def _add_text_area(self, label_text):
        ttk.Label(self, text=label_text + ":").pack(pady=5)
        text_area = tk.Text(self, height=6, width=80)
        text_area.pack(pady=5, padx=10)
        self.text_areas[label_text.lower()] = text_area
        logger.debug(f"Text area added for '{label_text}'")

    def browse_template(self):
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
        logger.info("Starting Excel edit thread")
        self.progress.pack(pady=10, fill='x', padx=20)
        self.progress.start(10)
        self.write_button.config(state='disabled')
        threading.Thread(target=self.edit_excel, daemon=True).start()

    def edit_excel(self):
        logger.info("edit_excel started")
        path = self.template_path_var.get()
        if not os.path.isfile(path):
            logger.error("Invalid Excel file path")
            self.after(0, self._handle_error, "Please select a valid Excel file.")
            return

        password = self.password_var.get()
        logger.debug(f"Using password: {'*' * len(password)}")
        get_lines = lambda key: [
            line.strip() for line in self.text_areas[key].get("1.0", tk.END).strip().splitlines() if line.strip()
        ]

        # Only get lines if checkbox is checked
        servers_lines = get_lines("servers") if self.chk_vars["servers"].get() else []
        trendstorage_lines = get_lines("trendstorage geheugen") if self.chk_vars["trendstorage"].get() else []
        cpu_lines = get_lines("cpu") if self.chk_vars["cpu"].get() else []
        memory_lines = get_lines("memory") if self.chk_vars["memory"].get() else []

        logger.info(f"Lines to process - Servers: {len(servers_lines)}, TrendStorage: {len(trendstorage_lines)}, CPU: {len(cpu_lines)}, Memory: {len(memory_lines)}")

        if not (servers_lines or trendstorage_lines or cpu_lines or memory_lines or self.chk_vars["all_licenses"].get()):
            logger.warning("No valid lines found or no sections selected to write")
            self.after(0, self._handle_error, "No valid lines found or no sections selected to write.")
            return

        # Ask save option (UI thread)
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

            # Processors
            def process_lines(lines, cell, log_prefix, checkboxes, warn_threshold=None, value_func=None):
                nonlocal warning_triggered
                for idx, line in enumerate(lines):
                    logger.debug(f"Processing line {idx+1} for {log_prefix}: {line}")
                    sheet, sheet_name = get_sheet(idx)
                    if not sheet: continue
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
                            logger.warning(f"[{log_prefix}] Value {percent}% exceeds threshold {warn_threshold}% on {sheet_name}")
                            for cb in checkboxes:
                                set_checkbox(sheet, cb, sheet_name)
                        else:
                            set_checkbox(sheet, checkboxes[0], sheet_name)
                    else:
                        set_checkbox(sheet, checkboxes[0], sheet_name)
                    protect(sheet)
                return True

            # Servers
            if servers_lines:
                logger.info("Processing Servers lines")
                process_lines(servers_lines, "B2", "Servers", ["CheckBox_C35"])

            # TrendStorage
            if trendstorage_lines:
                logger.info("Processing TrendStorage lines")
                def ts_value_func(line):
                    clean = line.replace('.', '')
                    numeric_str = clean.replace(',', '.')
                    try:
                        numeric_value = float(numeric_str)
                    except ValueError:
                        numeric_value = 0.0
                    percentage = (numeric_value / 10_000_000) * 100
                    perc_int = int(round(percentage))
                    return f"{clean} van 10000000 - {perc_int}%"
                process_lines(trendstorage_lines, "J36", "TrendStorage", ["CheckBox_C36", "CheckBox_E36"], warn_threshold=80, value_func=ts_value_func)

            # CPU/Memory
            if cpu_lines or memory_lines:
                logger.info("Processing CPU/Memory lines")
                max_len = max(len(cpu_lines), len(memory_lines))
                for idx in range(max_len):
                    logger.debug(f"Processing CPU/Memory line {idx+1}")
                    sheet, sheet_name = get_sheet(idx)
                    if not sheet: continue
                    try:
                        unprotect(sheet)
                    except RuntimeError:
                        return
                    cpu_val = cpu_lines[idx] if idx < len(cpu_lines) else ""
                    mem_val = memory_lines[idx] if idx < len(memory_lines) else ""
                    def parse_pct(val):
                        try:
                            return float(val.replace('%', '').replace(',', '.').strip())
                        except Exception:
                            return 0.0
                    cpu_float = parse_pct(cpu_val)
                    mem_float = parse_pct(mem_val)
                    cpu_text = f"{cpu_float:.2f} %" if cpu_val else "0.00 %"
                    mem_text = f"{mem_float:.2f} %" if mem_val else "0.00 %"
                    combined = f"CPU: {cpu_text} Memory: {mem_text}"
                    try:
                        sheet.range("J39").value = combined
                        logger.info(f"[CPU/Memory] Written '{combined}' to {sheet_name} J39")
                    except Exception as e:
                        logger.error(f"[CPU/Memory] Failed to write to {sheet_name} J39: {e}")
                    if cpu_float > 80 or mem_float > 80:
                        warning_triggered = True
                        logger.warning(f"CPU or Memory exceeds 80% on {sheet_name}")
                        for cb in ("CheckBox_C39", "CheckBox_E39"):
                            set_checkbox(sheet, cb, sheet_name)
                    else:
                        set_checkbox(sheet, "CheckBox_C39", sheet_name)
                    protect(sheet)

            # All Licenses
            if self.chk_vars["all_licenses"].get():
                logger.info("Processing All Licenses checkboxes")
                max_idx = max(len(servers_lines), len(trendstorage_lines), len(cpu_lines), len(memory_lines), 1)
                for idx in range(max_idx):
                    sheet, sheet_name = get_sheet(idx)
                    if not sheet: continue
                    try:
                        unprotect(sheet)
                    except RuntimeError:
                        return
                    set_checkbox(sheet, "CheckBox_C35", sheet_name)
                    protect(sheet)

            # Save workbook
            try:
                wb.save()
                logger.info(f"Workbook saved at {edit_path}")
            except Exception as e:
                logger.error(f"Failed to save workbook: {e}")
                self.after(0, self._handle_error, f"Failed to save the workbook:\n{e}")
                return

            wb.close()
            app.quit()
            logger.info("Workbook closed and Excel app quit")

            if warning_triggered:
                logger.warning("Warning threshold triggered, showing warning messagebox")
                self.after(0, lambda: messagebox.showwarning("Warning", "Warning: Some values exceeded 80% threshold.\nRelevant checkboxes have been marked."))
            else:
                logger.info("Excel file successfully updated, showing info messagebox")
                self.after(0, lambda: messagebox.showinfo("Success", "Excel file successfully updated."))

        except RuntimeError:
            logger.error("RuntimeError occurred, cleaning up Excel objects")
            try:
                wb.close()
            except Exception:
                pass
            try:
                app.quit()
            except Exception:
                pass
            return
        except Exception as e:
            logger.exception("Unexpected error")
            self.after(0, self._handle_error, f"An error occurred:\n{e}")
            try:
                wb.close()
            except Exception:
                pass
            try:
                app.quit()
            except Exception:
                pass
            return
        finally:
            logger.info("Resetting UI after Excel operation")
            self.after(0, self._reset_ui)

    def _handle_error(self, message):
        logger.error(f"Error shown to user: {message}")
        messagebox.showerror("Error", message)
        self._reset_ui()

    def _cancel_save(self):
        logger.info("User cancelled save operation")
        messagebox.showinfo("Cancelled", "Save cancelled.")
        self._reset_ui()

    def _reset_ui(self):
        logger.debug("Resetting UI state")
        self.progress.stop()
        self.progress.pack_forget()
        self.write_button.config(state='normal')

if __name__ == "__main__":
    logger.info("ExcelWriterWindow started as main application")
    root = tk.Tk()
    root.withdraw()
    win = ExcelWriterWindow(root)
    win.mainloop()
