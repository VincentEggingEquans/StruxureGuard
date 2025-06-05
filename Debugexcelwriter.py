import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xlwings as xw
import os
import threading
import shutil
import logging

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.DEBUG, format='[%(asctime)s] [%(levelname)s] %(message)s')

class ExcelEditorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Editor")
        self.geometry("800x900")

        # Checkbox variables
        self.chk_servers_var = tk.IntVar(value=1)
        self.chk_trendstorage_geheugen_var = tk.IntVar(value=1)  # Updated variable name here
        self.chk_cpu_var = tk.IntVar(value=0)
        self.chk_memory_var = tk.IntVar(value=0)
        self.chk_all_licenses_var = tk.IntVar(value=0)

        self.sections = [
            ("servers", "Servers"),
            ("trendstorage geheugen", "TrendStorage Geheugen"),
            ("cpu", "CPU"),
            ("memory", "Memory"),
        ]

        self._build_ui()

    def _build_ui(self):
        frame = ttk.Frame(self)
        frame.pack(pady=10, padx=10, fill='both', expand=True)

        ttk.Label(frame, text="Excel Template Path:").pack(anchor='w')
        path_frame = ttk.Frame(frame)
        path_frame.pack(fill='x')
        self.template_path_var = tk.StringVar()
        ttk.Entry(path_frame, textvariable=self.template_path_var, width=60).pack(side='left', padx=(0,5))
        ttk.Button(path_frame, text="Browse", command=self.browse_template).pack(side='left')

        # Checkbuttons for sections
        for key, label in self.sections:
            chk = ttk.Checkbutton(frame, text=label, variable=getattr(self, f"chk_{key.replace(' ', '_')}_var"))
            chk.pack(anchor='w')

        # Checkbox for all licenses
        chk_all = ttk.Checkbutton(frame, text="Write all 'Alle Benodigde licenties aanwezig' (C35 checkboxes)", variable=self.chk_all_licenses_var)
        chk_all.pack(anchor='w', pady=(5,10))

        # Text inputs for each section
        self.text_areas = {}
        for _, label in self.sections:
            ttk.Label(frame, text=label + ":").pack(anchor='w')
            text_area = tk.Text(frame, height=6, width=80)
            text_area.pack(pady=5)
            self.text_areas[label.lower()] = text_area

        self.progress = ttk.Progressbar(frame, mode='indeterminate')
        self.progress.pack(pady=10, fill='x')
        self.progress.pack_forget()

        self.write_button = ttk.Button(frame, text="Write to Excel", command=self.start_edit_excel_thread)
        self.write_button.pack(pady=10)

    def browse_template(self):
        path = filedialog.askopenfilename(
            parent=self,
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")]
        )
        if path:
            self.template_path_var.set(path)

    def start_edit_excel_thread(self):
        self.progress.pack(pady=10, fill='x')
        self.progress.start(10)
        self.write_button.config(state='disabled')
        threading.Thread(target=self.edit_excel, daemon=True).start()

    def edit_excel(self):
        path = self.template_path_var.get()
        if not os.path.isfile(path):
            self.after(0, self._handle_error, "Please select a valid Excel file.")
            return

        def get_lines(text_widget):
            content = text_widget.get("1.0", tk.END).strip()
            return [line.strip() for line in content.splitlines() if line.strip()]

        servers_lines = get_lines(self.text_areas["servers"]) if self.chk_servers_var.get() else []
        trendstorage_lines = get_lines(self.text_areas["trendstorage geheugen"]) if self.chk_trendstorage_geheugen_var.get() else []
        cpu_lines = get_lines(self.text_areas["cpu"]) if self.chk_cpu_var.get() else []
        memory_lines = get_lines(self.text_areas["memory"]) if self.chk_memory_var.get() else []

        logger.debug(f"Servers checkbox: {self.chk_servers_var.get()}")
        logger.debug(f"TrendStorage checkbox: {self.chk_trendstorage_geheugen_var.get()}")
        logger.debug(f"CPU checkbox: {self.chk_cpu_var.get()}")
        logger.debug(f"Memory checkbox: {self.chk_memory_var.get()}")
        logger.debug(f"All licenses checkbox: {self.chk_all_licenses_var.get()}")

        logger.debug(f"Servers input: {servers_lines}")
        logger.debug(f"TrendStorage input: {trendstorage_lines}")
        logger.debug(f"CPU input: {cpu_lines}")
        logger.debug(f"Memory input: {memory_lines}")

        if not (servers_lines or trendstorage_lines or cpu_lines or memory_lines or self.chk_all_licenses_var.get()):
            self.after(0, self._handle_error, "No valid lines found or no sections selected to write.")
            return

        # Save option prompt on main thread
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
                    try:
                        sheet.api.Unprotect()
                    except Exception:
                        logger.debug(f"[Servers] Sheet already unprotected or unprotect failed for {sheet_name}")
                    logger.info(f"[Servers] Writing to {sheet_name} B2: {line}")
                    try:
                        sheet.range("B2").value = line
                    except Exception as e:
                        logger.error(f"[Servers] Failed to write to {sheet_name} B2: {e}")

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
                        sheet.api.Unprotect()
                    except Exception:
                        logger.debug(f"[TrendStorage] Sheet already unprotected or unprotect failed for {sheet_name}")

                    sheet.range("J36").value = value

                    checkbox_c36 = "CheckBox_C36"
                    checkbox_e36 = "CheckBox_E36"

                    if percentage >= 80:
                        warning_triggered = True
                        for cb_name in (checkbox_c36, checkbox_e36):
                            if self._set_checkbox(sheet, cb_name, sheet_name):
                                checkboxes_checked.append(f"{sheet_name}:{cb_name}")

            def process_cpu(lines):
                logger.debug(f"[CPU] Processing {len(lines)} lines")
                for idx, line in enumerate(lines):
                    sheet_name = "Checklist Regelkast" if idx == 0 else f"Checklist Regelkast ({idx + 1})"
                    logger.debug(f"[CPU] Target sheet: {sheet_name}, value: {line}")
                    try:
                        sheet = wb.sheets[sheet_name]
                    except Exception as e:
                        logger.warning(f"[CPU] Sheet not found: {sheet_name} - {e}")
                        continue

                    try:
                        sheet.api.Unprotect()
                    except Exception:
                        logger.debug(f"[CPU] Sheet already unprotected or unprotect failed for {sheet_name}")

                    logger.info(f"[CPU] Writing to {sheet_name} B18: {line}")
                    try:
                        sheet.range("B18").value = line
                    except Exception as e:
                        logger.error(f"[CPU] Failed to write to {sheet_name} B18: {e}")

            def process_memory(lines):
                logger.debug(f"[Memory] Processing {len(lines)} lines")
                for idx, line in enumerate(lines):
                    sheet_name = "Checklist Regelkast" if idx == 0 else f"Checklist Regelkast ({idx + 1})"
                    logger.debug(f"[Memory] Target sheet: {sheet_name}, value: {line}")
                    try:
                        sheet = wb.sheets[sheet_name]
                    except Exception as e:
                        logger.warning(f"[Memory] Sheet not found: {sheet_name} - {e}")
                        continue

                    try:
                        sheet.api.Unprotect()
                    except Exception:
                        logger.debug(f"[Memory] Sheet already unprotected or unprotect failed for {sheet_name}")

                    logger.info(f"[Memory] Writing to {sheet_name} B21: {line}")
                    try:
                        sheet.range("B21").value = line
                    except Exception as e:
                        logger.error(f"[Memory] Failed to write to {sheet_name} B21: {e}")

            def process_all_licenses():
                logger.debug("[All Licenses] Processing checkbox C35 on all 'Checklist Regelkast' sheets")
                for sheet in wb.sheets:
                    if sheet.name.startswith("Checklist Regelkast"):
                        try:
                            sheet.api.Unprotect()
                        except Exception:
                            logger.debug(f"[All Licenses] Sheet already unprotected or unprotect failed for {sheet.name}")
                        try:
                            cb = sheet.api.CheckBoxes("CheckBox_C35")
                            cb.Value = 1  # Check it
                            logger.info(f"[All Licenses] Checked CheckBox_C35 on {sheet.name}")
                        except Exception as e:
                            logger.warning(f"[All Licenses] Could not set CheckBox_C35 on {sheet.name}: {e}")

            # Run the processing based on checkboxes
            if self.chk_servers_var.get():
                process_servers(servers_lines)
            if self.chk_trendstorage_geheugen_var.get():
                process_trendstorage(trendstorage_lines)
            if self.chk_cpu_var.get():
                process_cpu(cpu_lines)
            if self.chk_memory_var.get():
                process_memory(memory_lines)
            if self.chk_all_licenses_var.get():
                process_all_licenses()

            wb.save()
            wb.close()
            app.quit()

            msg = "Successfully wrote data to the Excel file."
            if warning_triggered:
                msg += "\n\nWARNING: Some TrendStorage values exceeded 80%, check highlighted checkboxes."
            logger.info(msg)

            self.after(0, messagebox.showinfo, "Success", msg)

        except Exception as e:
            logger.error(f"Error during Excel processing: {e}")
            self.after(0, self._handle_error, f"Failed to edit Excel file:\n{e}")
        finally:
            self.after(0, self._finish_write)

    def _set_checkbox(self, sheet, checkbox_name, sheet_name):
        try:
            cb = sheet.api.CheckBoxes(checkbox_name)
            cb.Value = 1
            logger.info(f"Checked {checkbox_name} on {sheet_name}")
            return True
        except Exception as e:
            logger.warning(f"Checkbox {checkbox_name} not found on {sheet_name}: {e}")
            return False

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

    def _finish_write(self):
        self.progress.stop()
        self.progress.pack_forget()
        self.write_button.config(state='normal')


if __name__ == "__main__":
    app = ExcelEditorApp()
    app.mainloop()
