import tkinter as tk
import logging

_debug_window = None
_log_text_widget = None
_log_messages = []  # Store all log messages here

def show_debug_log(master=None):
    global _debug_window, _log_text_widget
    if _debug_window is None or not tk.Toplevel.winfo_exists(_debug_window):
        _debug_window = tk.Toplevel(master)
        _debug_window.title("Debug Log")
        _debug_window.geometry("800x300")
        _log_text_widget = tk.Text(_debug_window, state='disabled')
        _log_text_widget.pack(expand=True, fill='both')
        # Insert all stored messages
        _log_text_widget.config(state='normal')
        for msg in _log_messages:
            _log_text_widget.insert(tk.END, msg + '\n')
        _log_text_widget.see(tk.END)
        _log_text_widget.config(state='disabled')
        _debug_window.protocol("WM_DELETE_WINDOW", _debug_window.destroy)
    else:
        _debug_window.lift()

def log_to_gui(message):
    global _log_text_widget, _log_messages
    _log_messages.append(message)
    if _log_text_widget:
        _log_text_widget.config(state='normal')
        _log_text_widget.insert(tk.END, message + '\n')
        _log_text_widget.see(tk.END)
        _log_text_widget.config(state='disabled')

class TkinterLogHandler(logging.Handler):
    def emit(self, record):
        msg = self.format(record)
        log_to_gui(msg)