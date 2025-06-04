import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xlwings as xw
import os
import threading
import shutil
import logging

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(message)s')

class ExcelWriterWindow(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("ExcelWriter")
        self.geometry("750x900")
        self.attributes('-topmost', True)

        # Template selection
        ttk.Label(self, text="Select Excel Template:").pack(pady=5)
        path_frame = ttk.Frame(self)
        path_frame.pack(pady=5, fill='x', padx=10)
        self.template_path_var = tk.StringVar()
        ttk.Entry(path_frame, textvariable=self.template_path_var, width=55).pack(side='left', padx=(0,5))
        ttk.Button(path_frame, text="Browse", command=self.browse_template).pack(side='left')

        # Checkboxes for toggling write options (all default OFF)
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

        # Text input areas in order: Servers, Trendstorage, CPU, Memory
        self._add_text_area("Servers:")
        self._add_text_area("TrendStorage Geheugen:")
        self._add_text_area("CPU:")
        self._add_text_area("Memory:")

        # Progressbar
        self.progress = ttk.Progressbar(self, mode='indeterminate')
        self.progress.pack(pady=10, fill='x', padx=20)
        self.progress.pack_forget()

        # Write button
        self.write_button = ttk.Button(self, text="Write to Excel", command=self.start_edit_excel_thread)
        self.write_button.pack(pady=10)

    def _add_text_area(self, label_text):
        ttk.Label(self, text=label_text).pack(pady=5)
        text_area = tk.Text(self, height=6, width=80)
        text_area.pack(pady=5, padx=10)
        setattr(self, f"{label_text.replace(' ', '_').replace(':','').lower()}_text", text_area)

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
            self._stop_progress()
            messagebox.showerror("Error", "Please select a valid Excel file.")
            return

        def get_lines(text_widget):
            content = text_widget.get("1.0", tk.END).strip()
            return [line.strip() for line in content.splitlines() if line.strip()]

        # Grab input lines only if their checkboxes are checked
        servers_lines = get_lines(self.servers_text) if self.chk_servers_var.get() else []
        trendstorage_lines = get_lines(self.trendstorage_geheugen_text) if self.chk_trendstorage_var.get() else []
        cpu_lines = get_lines(self.cpu_text) if self.chk_cpu_var.get() else []
        memory_lines = get_lines(self.memory_text) if self.chk_memory_var.get() else []

        if not (servers_lines or trendstorage_lines or cpu_lines or memory_lines or self.chk_all_licenses_var.get()):
            self._stop_progress()
            messagebox.showerror("Error", "No valid lines found or no sections selected to write.")
            return

        overwrite_original = messagebox.askyesno(
            "Save Option",
            "Do you want to update the original template?\n\n"
            "Yes: Overwrite original\nNo: Create a new edited copy"
        )

        if not overwrite_original:
            new_path = filedialog.asksaveasfilename(
                parent=self,
                defaultextension=".xlsm",
                filetypes=[("Excel Macro-Enabled Workbook", "*.xlsm")],
                initialfile=os.path.basename(path)
            )
            if not new_path:
                self._stop_progress()
                messagebox.showinfo("Cancelled", "Save cancelled.")
                return
            try:
                shutil.copy2(path, new_path)
                edit_path = new_path
                logger.info(f"Copied template to new file: {new_path}")
            except Exception as e:
                self._stop_progress()
                messagebox.showerror("Error", f"Failed to copy file:\n{e}")
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

            # Write Servers data to B2 in all relevant sheets
            def process_servers(lines):
                nonlocal checkboxes_checked
                for idx, line in enumerate(lines):
                    sheet_name = "Checklist Regelkast" if idx == 0 else f"Checklist Regelkast ({idx + 1})"
                    try:
                        sheet = wb.sheets[sheet_name]
                    except Exception:
                        logger.warning(f"Sheet not found: {sheet_name}")
                        continue
                    logger.info(f"[Servers] Writing to {sheet_name} B2: {line}")
                    sheet.range("B2").value = line

            # Write TrendStorage Geheugen data to J36 and checkbox logic
            def process_trendstorage(lines):
                nonlocal warning_triggered, checkboxes_checked
                for idx, line in enumerate(lines):
                    sheet_name = "Checklist Regelkast" if idx == 0 else f"Checklist Regelkast ({idx + 1})"
                    try:
                        sheet = wb.sheets[sheet_name]
                    except Exception:
                        logger.warning(f"Sheet not found: {sheet_name}")
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

                    logger.info(f"[TrendStorage Geheugen] Writing to {sheet_name} J36: {value}")

                    try:
                        sheet.api.Unprotect()
                    except Exception:
                        pass

                    sheet.range("J36").value = value

                    checkbox_c36 = "CheckBox_C36"
                    checkbox_e36 = "CheckBox_E36"

                    if percentage >= 80:
                        warning_triggered = True
                        for cb_name in (checkbox_c36, checkbox_e36):
                            if self._set_checkbox(sheet, cb_name, sheet_name):
                                checkboxes_checked.append(f"{cb_name} on {sheet_name}")
                    else:
                        if self._set_checkbox(sheet, checkbox_c36, sheet_name):
                            checkboxes_checked.append(f"{checkbox_c36} on {sheet_name}")

            # Write combined CPU & Memory to J39 and checkbox logic
            def process_cpu_memory(cpu_lines, memory_lines):
                nonlocal warning_triggered, checkboxes_checked

                max_lines = max(len(cpu_lines), len(memory_lines))
                for idx in range(max_lines):
                    sheet_name = "Checklist Regelkast" if idx == 0 else f"Checklist Regelkast ({idx + 1})"
                    try:
                        sheet = wb.sheets[sheet_name]
                    except Exception:
                        logger.warning(f"Sheet not found: {sheet_name}")
                        continue

                    cpu_val = 0.0
                    mem_val = 0.0

                    if idx < len(cpu_lines):
                        try:
                            cpu_val = float(cpu_lines[idx].replace(',', '.').replace('%','').strip())
                        except Exception:
                            cpu_val = 0.0
                    if idx < len(memory_lines):
                        try:
                            mem_val = float(memory_lines[idx].replace(',', '.').replace('%','').strip())
                        except Exception:
                            mem_val = 0.0

                    cpu_text = f"CPU: {cpu_val:.2f} %"
                    mem_text = f"Memory: {mem_val:.2f} %"
                    combined_text = f"{cpu_text} {mem_text}"

                    logger.info(f"[CPU/Memory] Writing to {sheet_name} J39: {combined_text}")

                    sheet.range("J39").value = combined_text

                    checkbox_c39 = "CheckBox_C39"
                    checkbox_e39 = "CheckBox_E39"

                    # Check C39 always if writing cpu or memory
                    if self._set_checkbox(sheet, checkbox_c39, sheet_name):
                        checkboxes_checked.append(f"{checkbox_c39} on {sheet_name}")

                    # If either CPU or Memory > 80%, also check E39 and warn
                    if cpu_val > 80 or mem_val > 80:
                        warning_triggered = True
                        if self._set_checkbox(sheet, checkbox_e39, sheet_name):
                            checkboxes_checked.append(f"{checkbox_e39} on {sheet_name}")

            # Check or uncheck all C35 checkboxes on all sheets if checked
            def process_all_licenses():
                nonlocal checkboxes_checked
                if not self.chk_all_licenses_var.get():
                    return
                idx = 0
                while True:
                    sheet_name = "Checklist Regelkast" if idx == 0 else f"Checklist Regelkast ({idx + 1})"
                    try:
                        sheet = wb.sheets[sheet_name]
                    except Exception:
                        break  # no more sheets matching pattern
                    checkbox_c35 = "CheckBox_C35"
                    if self._set_checkbox(sheet, checkbox_c35, sheet_name):
                        checkboxes_checked.append(f"{checkbox_c35} on {sheet_name}")
                    idx += 1

            # Process all sections
            if self.chk_servers_var.get():
                process_servers(servers_lines)

            if self.chk_trendstorage_var.get():
                process_trendstorage(trendstorage_lines)

            if self.chk_cpu_var.get() or self.chk_memory_var.get():
                process_cpu_memory(cpu_lines if self.chk_cpu_var.get() else [], memory_lines if self.chk_memory_var.get() else [])

            if self.chk_all_licenses_var.get():
                process_all_licenses()

            wb.save()
            wb.close()
            app.quit()

            logger.info(f"Checkboxes checked: {checkboxes_checked}")

            if warning_triggered:
                messagebox.showwarning("Warning", "One or more values exceeded 80%, related checkboxes have been checked!")

            self._stop_progress()
            messagebox.showinfo("Success", "Excel file has been updated successfully.")

        except Exception as e:
            self._stop_progress()
            messagebox.showerror("Error", f"An error occurred:\n{e}")
            logger.error(f"Exception during Excel editing: {e}")

    def _set_checkbox(self, sheet, checkbox_name, sheet_name):
        """
        Try to check the checkbox by its name.
        Returns True if checked successfully, False otherwise.
        """
        try:
            checkbox = sheet.api.OLEObjects(checkbox_name)
            checkbox.Object.Value = True
            logger.debug(f"Checked {checkbox_name} in {sheet_name}")
            return True
        except Exception as e:
            logger.warning(f"Could not check checkbox {checkbox_name} in {sheet_name}: {e}")
            return False

    def _stop_progress(self):
        self.progress.stop()
        self.progress.pack_forget()
        self.write_button.config(state='normal')


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    win = ExcelWriterWindow(root)
    win.mainloop()
