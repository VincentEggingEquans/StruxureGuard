import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xlwings as xw
import os
import threading
import shutil
import logging

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.DEBUG, format='[%(asctime)s] [%(levelname)s] %(message)s')

class ExcelWriterWindow(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("ExcelWriter")
        self.geometry("800x900")
        self.attributes('-topmost', False)

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
        self.chk_servers_var = tk.BooleanVar(value=False)
        self.chk_trendstorage_var = tk.BooleanVar(value=False)
        self.chk_cpu_var = tk.BooleanVar(value=False)
        self.chk_memory_var = tk.BooleanVar(value=False)
        self.chk_all_licenses_var = tk.BooleanVar(value=False)

        chk_frame = ttk.Frame(self)
        chk_frame.pack(pady=5, padx=10, fill='x')

        ttk.Checkbutton(chk_frame, text="Write Servers data", variable=self.chk_servers_var).grid(row=0, column=0, sticky='w')
        ttk.Checkbutton(chk_frame, text="Write TrendStorage Geheugen data", variable=self.chk_trendstorage_var).grid(row=1, column=0, sticky='w')
        ttk.Checkbutton(chk_frame, text="Write CPU data", variable=self.chk_cpu_var).grid(row=2, column=0, sticky='w')
        ttk.Checkbutton(chk_frame, text="Write Memory data", variable=self.chk_memory_var).grid(row=3, column=0, sticky='w')
        ttk.Checkbutton(chk_frame, text="Write all 'Alle Benodigde licenties aanwezig' (C35 checkboxes)", variable=self.chk_all_licenses_var).grid(row=4, column=0, sticky='w')

        # Text areas
        self.text_areas = {}
        for label in ["Servers", "TrendStorage Geheugen", "CPU", "Memory"]:
            self._add_text_area(label)

        self.progress = ttk.Progressbar(self, mode='indeterminate')
        self.progress.pack(pady=10, fill='x', padx=20)
        self.progress.pack_forget()

        self.write_button = ttk.Button(self, text="Write to Excel", command=self.start_edit_excel_thread)
        self.write_button.pack(pady=10)

    def _add_text_area(self, label_text):
        ttk.Label(self, text=label_text + ":").pack(pady=5)
        text_area = tk.Text(self, height=6, width=80)
        text_area.pack(pady=5, padx=10)
        self.text_areas[label_text.lower()] = text_area

    def browse_template(self):
        path = filedialog.askopenfilename(
            parent=self,
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")]
        )
        if path:
            self.template_path_var.set(path)

    def start_edit_excel_thread(self):
        self.progress.pack(pady=10, fill='x', padx=20)
        self.progress.start(10)
        self.write_button.config(state='disabled')
        threading.Thread(target=self.edit_excel, daemon=True).start()

    def edit_excel(self):
        path = self.template_path_var.get()
        if not os.path.isfile(path):
            self.after(0, self._handle_error, "Please select a valid Excel file.")
            return

        password = self.password_var.get()  # Get password from text box

        def get_lines(text_widget):
            content = text_widget.get("1.0", tk.END).strip()
            return [line.strip() for line in content.splitlines() if line.strip()]

        servers_lines = get_lines(self.text_areas["servers"]) if self.chk_servers_var.get() else []
        trendstorage_lines = get_lines(self.text_areas["trendstorage geheugen"]) if self.chk_trendstorage_var.get() else []
        cpu_lines = get_lines(self.text_areas["cpu"]) if self.chk_cpu_var.get() else []
        memory_lines = get_lines(self.text_areas["memory"]) if self.chk_memory_var.get() else []

        logger.debug(f"Checkbox states: Servers={self.chk_servers_var.get()}, TrendStorage={self.chk_trendstorage_var.get()}, CPU={self.chk_cpu_var.get()}, Memory={self.chk_memory_var.get()}, AllLicenses={self.chk_all_licenses_var.get()}")
        logger.debug(f"Input lines - Servers: {servers_lines}, TrendStorage: {trendstorage_lines}, CPU: {cpu_lines}, Memory: {memory_lines}")

        if not (servers_lines or trendstorage_lines or cpu_lines or memory_lines or self.chk_all_licenses_var.get()):
            self.after(0, self._handle_error, "No valid lines found or no sections selected to write.")
            return

        save_option_event = threading.Event()
        self.overwrite_original = None
        def ask_save_option():
            self.overwrite_original = messagebox.askyesno(
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
                self.after(0, self._cancel_save)
                return

            try:
                shutil.copy2(path, self.new_save_path)
                edit_path = self.new_save_path
                logger.info(f"Copied template to new file: {self.new_save_path}")
            except Exception as e:
                self.after(0, self._handle_error, f"Failed to copy file:\n{e}")
                logger.error(f"Failed to copy template file: {e}")
                return
        else:
            edit_path = path

        warning_triggered = False
        checkboxes_checked = []

        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False

            wb = app.books.open(edit_path)

            # Helper to unprotect/protect sheets with password
            def unprotect_sheet(sheet):
                try:
                    sheet.api.Unprotect(Password=password)
                    logger.debug(f"Unprotected sheet {sheet.name} with password")
                except Exception as e:
                    logger.warning(f"Could not unprotect sheet {sheet.name} with password: {e}")

            def protect_sheet(sheet):
                try:
                    sheet.api.Protect(Password=password)
                    logger.debug(f"Protected sheet {sheet.name} with password")
                except Exception as e:
                    logger.warning(f"Could not protect sheet {sheet.name} with password: {e}")

            def _set_checkbox(sheet, checkbox_name, sheet_name):
                try:
                    shape = sheet.api.Shapes(checkbox_name)
                    shape.OLEFormat.Object.Value = True
                    logger.info(f"Checked {checkbox_name} on {sheet_name}")
                    return True
                except Exception as e:
                    logger.warning(f"Failed to check {checkbox_name} on {sheet_name}: {e}")
                    return False

            self._set_checkbox = _set_checkbox

            def process_servers(lines):
                logger.debug(f"[Servers] Processing {len(lines)} lines")
                for idx, line in enumerate(lines):
                    sheet_name = "Checklist Regelkast" if idx == 0 else f"Checklist Regelkast ({idx + 1})"
                    logger.debug(f"[Servers] Target sheet: {sheet_name}, value: {line}")
                    try:
                        sheet = wb.sheets[sheet_name]
                    except Exception as e:
                        logger.warning(f"[Servers] Sheet not found: {sheet_name} - {e}")
                        continue
                    unprotect_sheet(sheet)
                    try:
                        sheet.range("B2").value = line
                        logger.info(f"[Servers] Written '{line}' to {sheet_name} B2")
                    except Exception as e:
                        logger.error(f"[Servers] Failed to write to {sheet_name} B2: {e}")
                    protect_sheet(sheet)

            def process_trendstorage(lines):
                nonlocal warning_triggered, checkboxes_checked
                logger.debug(f"[TrendStorage] Processing {len(lines)} lines")
                for idx, line in enumerate(lines):
                    sheet_name = "Checklist Regelkast" if idx == 0 else f"Checklist Regelkast ({idx + 1})"
                    logger.debug(f"[TrendStorage] Target sheet: {sheet_name}, value: {line}")
                    try:
                        sheet = wb.sheets[sheet_name]
                    except Exception as e:
                        logger.warning(f"[TrendStorage] Sheet not found: {sheet_name} - {e}")
                        continue

                    unprotect_sheet(sheet)

                    clean_line = line.replace('.', '')
                    numeric_str = clean_line.replace(',', '.')
                    try:
                        numeric_value = float(numeric_str)
                    except ValueError:
                        numeric_value = 0.0

                    percentage = (numeric_value / 10_000_000) * 100
                    perc_int = int(round(percentage))
                    value = f"{clean_line} van 10000000 - {perc_int}%"

                    logger.info(f"[TrendStorage] Writing to {sheet_name} J36: {value}")

                    try:
                        sheet.range("J36").value = value
                    except Exception as e:
                        logger.error(f"[TrendStorage] Failed to write to {sheet_name} J36: {e}")

                    checkbox_c36 = "CheckBox_C36"
                    checkbox_e36 = "CheckBox_E36"

                    if percentage >= 80:
                        warning_triggered = True
                        for cb_name in (checkbox_c36, checkbox_e36):
                            if self._set_checkbox(sheet, cb_name, sheet_name):
                                checkboxes_checked.append(f"{sheet_name}:{cb_name}")
                    else:
                        if self._set_checkbox(sheet, checkbox_c36, sheet_name):
                            checkboxes_checked.append(f"{sheet_name}:{checkbox_c36}")

                    protect_sheet(sheet)

            def process_cpu_memory_combined(cpu_lines, memory_lines):
                nonlocal warning_triggered
                max_len = max(len(cpu_lines), len(memory_lines))

                for idx in range(max_len):
                    sheet_name = "Checklist Regelkast" if idx == 0 else f"Checklist Regelkast ({idx + 1})"
                    try:
                        sheet = wb.sheets[sheet_name]
                    except Exception as e:
                        logger.warning(f"[CPU/Memory Combined] Sheet not found: {sheet_name} - {e}")
                        continue

                    unprotect_sheet(sheet)

                    cpu_val_raw = cpu_lines[idx] if idx < len(cpu_lines) else ""
                    mem_val_raw = memory_lines[idx] if idx < len(memory_lines) else ""

                    def parse_percentage(val):
                        try:
                            return float(val.replace('%', '').replace(',', '.').strip())
                        except Exception:
                            return 0.0

                    cpu_val_float = parse_percentage(cpu_val_raw)
                    mem_val_float = parse_percentage(mem_val_raw)

                    cpu_text = f"{cpu_val_float:.2f} %" if cpu_val_raw else "0.00 %"
                    mem_text = f"{mem_val_float:.2f} %" if mem_val_raw else "0.00 %"

                    combined_text = f"CPU: {cpu_text} Memory: {mem_text}"

                    try:
                        sheet.range("J39").value = combined_text
                        logger.info(f"[CPU/Memory Combined] Written '{combined_text}' to {sheet_name} J39")
                    except Exception as e:
                        logger.error(f"[CPU/Memory Combined] Failed to write to {sheet_name} J39: {e}")

                    checkbox_c39 = "CheckBox_C39"
                    checkbox_e39 = "CheckBox_E39"

                    if cpu_val_float > 80 or mem_val_float > 80:
                        warning_triggered = True
                        for cb_name in (checkbox_c39, checkbox_e39):
                            self._set_checkbox(sheet, cb_name, sheet_name)
                    else:
                        self._set_checkbox(sheet, checkbox_c39, sheet_name)

                    protect_sheet(sheet)

            def process_all_licenses():
                if not self.chk_all_licenses_var.get():
                    return
                max_idx = max(len(servers_lines), len(trendstorage_lines), len(cpu_lines), len(memory_lines), 1)
                for idx in range(max_idx):
                    sheet_name = "Checklist Regelkast" if idx == 0 else f"Checklist Regelkast ({idx + 1})"
                    try:
                        sheet = wb.sheets[sheet_name]
                    except Exception as e:
                        logger.warning(f"[All Licenses] Sheet not found: {sheet_name} - {e}")
                        continue
                    unprotect_sheet(sheet)
                    checkbox_c35 = "CheckBox_C35"
                    if self._set_checkbox(sheet, checkbox_c35, sheet_name):
                        logger.info(f"[All Licenses] Checked {checkbox_c35} on {sheet_name}")
                    protect_sheet(sheet)

            if self.chk_servers_var.get():
                process_servers(servers_lines)
            if self.chk_trendstorage_var.get():
                process_trendstorage(trendstorage_lines)
            if self.chk_cpu_var.get() or self.chk_memory_var.get():
                process_cpu_memory_combined(cpu_lines, memory_lines)
            if self.chk_all_licenses_var.get():
                process_all_licenses()

            wb.save()
            wb.close()
            app.quit()

        except Exception as e:
            self.after(0, self._handle_error, f"Failed to write to Excel file:\n{e}")
            logger.error(f"Exception during Excel writing: {e}")
            return

        def finish():
            self.progress.stop()
            self.progress.pack_forget()
            self.write_button.config(state='normal')
            if warning_triggered:
                messagebox.showwarning("Warning", "Some values exceed the 80% threshold! Please check the checked boxes.")
            else:
                messagebox.showinfo("Success", "Excel file successfully updated.")

        self.after(0, finish)

    def _handle_error(self, message):
        self.progress.stop()
        self.progress.pack_forget()
        self.write_button.config(state='normal')
        messagebox.showerror("Error", message)

    def _cancel_save(self):
        self.progress.stop()
        self.progress.pack_forget()
        self.write_button.config(state='normal')
        messagebox.showinfo("Cancelled", "Save cancelled by user.")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Hide root window
    window = ExcelWriterWindow()
    window.mainloop()
