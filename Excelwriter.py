from tkinter import simpledialog
import shutil

def edit_excel(self):
    path = self.template_path_var.get()
    if not os.path.isfile(path):
        messagebox.showerror("Error", "Please select a valid Excel file.")
        return
    
    # Read pasted text
    text_content = self.text_area.get("1.0", tk.END).strip()
    if not text_content:
        messagebox.showerror("Error", "Please paste some text to write into the Excel sheets.")
        return
    
    lines = [line.strip() for line in text_content.splitlines() if line.strip()]

    # Ask user whether to overwrite or create new copy
    result = messagebox.askyesno(
        "Save Option",
        "Do you want to update the original template?\n\n"
        "Yes: Overwrite original\nNo: Create a new edited copy"
    )

    if not result:
        # Make new copy
        import datetime
        base, ext = os.path.splitext(path)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        new_path = f"{base}_edited_{timestamp}{ext}"
        try:
            shutil.copy2(path, new_path)
            path = new_path  # update to new copy
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create new file copy:\n{e}")
            return

    try:
        app = xw.App(visible=True)
        wb = app.books.open(path)

        for sheet in wb.sheets:
            if sheet.name.startswith("Checklist Regelkast"):
                start_row = 36
                col_j = 'J'
                col_k = 'K'

                for i, line in enumerate(lines):
                    cell_j = f"{col_j}{start_row + i}"
                    cell_k = f"{col_k}{start_row + i}"

                    text_to_write = f"{line} van 10000000"
                    sheet.range(cell_j).value = text_to_write

                    length = len(line)
                    percent = (length / 10_000_000) * 100
                    sheet.range(cell_k).value = f"{percent:.6f}%"

        wb.save()
        messagebox.showinfo("Success", f"Excel file saved:\n{path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")
    finally:
        try:
            wb.close()
            app.quit()
        except:
            pass
