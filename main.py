# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, filedialog
import threading
import time
from typing import Callable
from ImportData import import_data

class ModernAutoImportApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Auto Import Data")
        self.geometry("600x400")
        self.configure(bg="#f0f0f0")

        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.configure_styles()

        self.task_completed = threading.Event()
        self.setup_ui()

    def configure_styles(self):
        self.style.configure('TFrame', background="#f0f0f0")
        self.style.configure('TLabel', background="#f0f0f0", font=("Helvetica", 12))
        self.style.configure('TEntry', font=("Helvetica", 12))
        self.style.configure('TButton', font=("Helvetica", 12, "bold"))
        self.style.configure('Header.TLabel', font=("Helvetica", 24, "bold"))
        self.style.configure('Footer.TFrame', background="#e0e0e0")

    def setup_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        header = ttk.Label(self, text="AUTO IMPORT DATA", style='Header.TLabel')
        header.grid(row=0, column=0, pady=(30, 20), sticky="n")

        main_frame = ttk.Frame(self)
        main_frame.grid(row=1, column=0, padx=40, pady=20, sticky="nsew")
        main_frame.columnconfigure(1, weight=1)

        self.excel_entry = self.create_file_input(main_frame, "Excel File:", 0, self.browse_excel)
        self.timelog_entry = self.create_file_input(main_frame, "Timelog File:", 1, self.browse_timelog)

        footer = ttk.Frame(self, style='Footer.TFrame')
        footer.grid(row=3, column=0, sticky="ews")
        footer.columnconfigure(0, weight=1)

        self.start_button = ttk.Button(footer, text='Start Import', command=self.handle_start_click, width=20)
        self.start_button.grid(row=0, column=0, pady=20)

        self.progress_bar = ttk.Progressbar(footer, orient='horizontal', mode='indeterminate', length=400)

    def create_file_input(self, parent: ttk.Frame, label: str, row: int, command: Callable) -> ttk.Entry:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=10)
        entry = ttk.Entry(parent)
        entry.grid(row=row, column=1, sticky="ew", padx=(10, 10), pady=10)
        ttk.Button(parent, text="Browse", command=command, width=10).grid(row=row, column=2, sticky="e", pady=10,
                                                                          padx=(10, 0))
        return entry

    def browse_excel(self):
        self.browse_file(self.excel_entry)

    def browse_timelog(self):
        self.browse_file(self.timelog_entry)

    def browse_file(self, entry: ttk.Entry):
        file_path = filedialog.askopenfilename(
            filetypes=[("All Files", "*.*"), ("Excel Files", "*.xlsx;*.xls"), ("Text Files", "*.txt")])
        if file_path:
            entry.delete(0, tk.END)
            entry.insert(0, file_path)

    def handle_start_click(self):
        self.toggle_ui_state(False)
        self.task_completed.clear()
        self.progress_bar.grid(row=0, column=0, pady=20, padx=40)
        self.start_button.grid_remove()

        threading.Thread(target=self.run_task, daemon=True).start()
        threading.Thread(target=self.update_progress, daemon=True).start()

    def toggle_ui_state(self, enabled: bool):
        state = 'normal' if enabled else 'disabled'
        for widget in (self.excel_entry, self.timelog_entry):
            widget.configure(state=state)

    def run_task(self):
        success = False
        try:
            import_data(self.excel_entry.get(), self.timelog_entry.get())
            success = True
        except Exception as e:
            print(f"An error occurred during import: {str(e)}")
        finally:
            self.task_completed.set()
            self.after(0, lambda: self.show_alert(success))

    def update_progress(self):
        while not self.task_completed.is_set():
            self.progress_bar['value'] = (self.progress_bar['value'] + 10) % 100
            self.update_idletasks()
            time.sleep(0.1)

        self.progress_bar.grid_remove()
        self.start_button.grid()
        self.toggle_ui_state(True)
        self.progress_bar['value'] = 0  # Reset progress bar

    def show_alert(self, success):
        alert_window = tk.Toplevel(self)
        alert_window.title("Import Result")
        alert_window.configure(bg="#f0f0f0")

        # Set window size
        window_width = 300
        window_height = 100

        # Get screen width and height
        screen_width = alert_window.winfo_screenwidth()
        screen_height = alert_window.winfo_screenheight()

        # Calculate position for center of screen
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)

        # Set the position of the window to the center of the screen
        alert_window.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

        message = "Data import completed successfully!" if success else "Data import failed. Please check the logs."
        color = "#4CAF50" if success else "#F44336"  # Green for success, Red for failure

        label = ttk.Label(alert_window, text=message, background="#f0f0f0", foreground=color,
                          font=("Helvetica", 12, "bold"))
        label.pack(expand=True)

        ok_button = ttk.Button(alert_window, text="OK", command=alert_window.destroy)
        ok_button.pack(pady=10)

        alert_window.transient(self)
        alert_window.grab_set()
        self.wait_window(alert_window)


if __name__ == "__main__":
    app = ModernAutoImportApp()
    app.mainloop()