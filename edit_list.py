import os
import json
import tkinter as tk
from tkinter import messagebox

def open_txt_popup(parent, bg_color, list_bg, fg_color, load_from_file_callback):
    try:
        with open("settings.json", "r", encoding="utf-8") as f:
            settings = json.load(f)
        file_path = settings.get("gamelist")
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Error", "The file from settings.json does not exist.")
            return
    except Exception as e:
        messagebox.showerror("Error", f"Could not load settings.json.\n{e}")
        return

    try:
        with open(file_path, "r", encoding="utf-8") as f:
            content = f.read()
    except Exception as e:
        messagebox.showerror("Error", f"Could not open file.\n{e}")
        return

    popup = tk.Toplevel(parent)
    popup.title(f"Editing: {os.path.basename(file_path)}")
    popup.resizable(False, False)
    popup.configure(bg=bg_color)

    auto_save = tk.BooleanVar(value=False)
    warned_once = tk.BooleanVar(value=False)

    def verify_content(text):
        lines = text.splitlines()
        for line in lines:
            line = line.strip()
            if not line:
                continue
            if line.startswith("!admin "):
                line = line[7:].strip()
            if not (line.startswith('"') and line.endswith('"')):
                return False
            path = line.strip('"')
            if not os.path.exists(path):
                return False
        return True

    def save_to_file():
        text = text_widget.get("1.0", "end-1c")
        if verify_content(text):
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(text)
                messagebox.showinfo("Saved", "File saved successfully.")
                load_from_file_callback(file_path)
            except Exception as e:
                messagebox.showerror("Error", f"Could not save file.\n{e}")
        else:
            messagebox.showerror("Invalid Format", "The content has invalid formatting or missing files.")

    def on_change(event=None):
        update_line_numbers()
        if auto_save.get():
            text = text_widget.get("1.0", "end-1c")
            if verify_content(text):
                try:
                    with open(file_path, "w", encoding="utf-8") as f:
                        f.write(text)
                    load_from_file_callback(file_path)
                    if not warned_once.get():
                        messagebox.showinfo("Auto Saved", "Saved automatically!")
                        warned_once.set(True)
                except Exception:
                    pass
            else:
                warned_once.set(False)
        text_widget.edit_modified(False)

    def update_line_numbers():
        lines = text_widget.get("1.0", "end-1c").splitlines()
        line_numbers.config(state="normal")
        line_numbers.delete("1.0", "end")
        for i in range(1, len(lines) + 1):
            line_numbers.insert("end", f"{i}\n")
        line_numbers.config(state="disabled")

    def sync_scroll(*args):
        text_widget.yview(*args)
        line_numbers.yview(*args)

    main_frame = tk.Frame(popup, bg=bg_color)
    main_frame.pack(fill="both", expand=True)
    main_frame.columnconfigure(1, weight=1)
    main_frame.rowconfigure(0, weight=1)

    line_numbers = tk.Text(main_frame, width=4, padx=4, takefocus=0, border=0,
                           background="#2b2b2b", foreground="gray", state="disabled", font=("Consolas", 14))
    line_numbers.grid(row=0, column=0, sticky="ns", pady=10)

    text_widget = tk.Text(main_frame, wrap="word", bg=list_bg, fg=fg_color,
                          insertbackground="white", font=("Consolas", 14), undo=True)
    text_widget.insert("1.0", content)
    text_widget.grid(row=0, column=1, sticky="nsew", padx=(0, 10), pady=10)
    text_widget.bind("<<Modified>>", on_change)

    scrollbar = tk.Scrollbar(main_frame, command=sync_scroll)
    scrollbar.grid(row=0, column=2, sticky="ns", pady=10)
    text_widget.config(yscrollcommand=scrollbar.set)
    line_numbers.config(yscrollcommand=scrollbar.set)

    button_frame = tk.Frame(main_frame, bg=bg_color)
    button_frame.grid(row=1, column=1, sticky="ew", pady=(0, 10))

    save_btn = tk.Button(button_frame, text="Save", width=12, bg="#3cb371", fg="white", command=save_to_file)
    save_btn.pack(side="left", padx=10, pady=5)

    def toggle_auto_save():
        auto_save.set(not auto_save.get())
        warned_once.set(False)
        auto_save_btn.config(
            text=f"Auto Save: {'ON' if auto_save.get() else 'OFF'}",
            bg="#3cb371" if auto_save.get() else "#777"
        )

    auto_save_btn = tk.Button(button_frame, text="Auto Save: OFF", width=14, bg="#777", fg="white", command=toggle_auto_save)
    auto_save_btn.pack(side="left", padx=10, pady=5)

    update_line_numbers()
