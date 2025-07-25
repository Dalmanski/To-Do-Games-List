import tkinter as tk
from tkinter import ttk
import json
import os
import sys

if getattr(sys, 'frozen', False):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

SETTINGS_FILE = os.path.join(base_dir, "settings.json")

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_settings(settings):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=4)

def on_mode_change(event=None):
    current_settings["Mode"] = mode_var.get()
    save_settings(current_settings)

def open_settings_popup(parent=None):
    global current_settings, mode_var

    current_settings = load_settings()

    settings_win = tk.Toplevel(parent)
    settings_win.title("Settings")
    settings_win.geometry("280x110")
    settings_win.configure(bg="#1e1e1e")
    settings_win.resizable(False, False)

    frame = tk.Frame(settings_win, bg="#1e1e1e")
    frame.pack(pady=30)

    label = tk.Label(frame, text="Mode:", fg="white", bg="#1e1e1e", font=("Segoe UI", 10, "bold"))
    label.pack(side=tk.LEFT, padx=(0, 10))

    mode_var = tk.StringVar(value=current_settings.get("Mode", "simple"))
    mode_dropdown = ttk.Combobox(frame, textvariable=mode_var, values=["simple", "edit"], state="readonly", width=10)
    mode_dropdown.pack(side=tk.LEFT)

    # Apply dark theme to Combobox
    style = ttk.Style()
    style.theme_use("clam")  # Use clam theme for styling support
    style.configure("TCombobox",
                    fieldbackground="#2e2e2e",
                    background="#2e2e2e",
                    foreground="white",
                    selectforeground="white",
                    selectbackground="#2e2e2e",
                    arrowcolor="white")
    
    # Adjust dropdown list appearance
    style.map("TCombobox",
              fieldbackground=[('readonly', '#2e2e2e')],
              background=[('readonly', '#2e2e2e')],
              foreground=[('readonly', 'white')])

    mode_dropdown.bind("<<ComboboxSelected>>", on_mode_change)

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    open_settings_popup(root)
    root.mainloop()
