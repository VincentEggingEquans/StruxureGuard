import tkinter as tk
from tkinter import scrolledtext
import logging
import threading

class DebugLogWindow:
    """Singleton window for displaying debug logs in a Tkinter ScrolledText widget."""
    _instance = None
    _lock = threading.Lock()

    def __init__(self):
        self.window = None
        self.text_widget = None
        self.messages = []

    @classmethod
    def instance(cls) -> "DebugLogWindow":
        with cls._lock:
            if cls._instance is None:
                cls._instance = DebugLogWindow()
            return cls._instance

    def show(self, master=None):
        if self.window is None or not tk.Toplevel.winfo_exists(self.window):
            self.window = tk.Toplevel(master)
            self.window.title("Debug Log")
            self.window.geometry("800x300")
            self.text_widget = scrolledtext.ScrolledText(self.window, state='disabled')
            self.text_widget.pack(expand=True, fill='both')
            self.window.protocol("WM_DELETE_WINDOW", self._on_close)
            self._refresh()
        else:
            self.window.lift()

    def _on_close(self):
        if self.window:
            self.window.destroy()
            self.window = None
            self.text_widget = None

    def add_message(self, message: str):
        self.messages.append(message)
        if self.text_widget:
            try:
                self.text_widget.after(0, self._append_message, message)
            except RuntimeError:
                pass

    def _append_message(self, message: str):
        if not self.text_widget:
            return
        self.text_widget.config(state='normal')
        self.text_widget.insert(tk.END, message + '\n')
        self.text_widget.see(tk.END)
        self.text_widget.config(state='disabled')

    def _refresh(self):
        if self.text_widget:
            self.text_widget.config(state='normal')
            self.text_widget.delete('1.0', tk.END)
            for msg in self.messages:
                self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.see(tk.END)
            self.text_widget.config(state='disabled')

    def clear(self):
        self.messages.clear()
        if self.text_widget:
            self.text_widget.config(state='normal')
            self.text_widget.delete('1.0', tk.END)
            self.text_widget.config(state='disabled')

def show_debug_log(master=None):
    """Show the debug log window. If already open, brings it to the front."""
    DebugLogWindow.instance().show(master)

def log_to_gui(message: str):
    """Append a log message to the debug log window."""
    DebugLogWindow.instance().add_message(message)

def clear_debug_log():
    """Clear all messages from the debug log window."""
    DebugLogWindow.instance().clear()

class TkinterLogHandler(logging.Handler):
    """Logging handler that sends log messages to the Tkinter debug log window."""
    def emit(self, record):
        msg = self.format(record)
        log_to_gui(msg)
