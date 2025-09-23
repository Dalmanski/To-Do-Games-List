import os
import json
import tkinter as tk
from tkinter import messagebox

def open_txt_popup(parent, bg_color, list_bg, fg_color, load_from_file_callback):
    settings_path = "settings.json"
    settings = {}

    try:
        with open(settings_path, "r", encoding="utf-8") as f:
            settings = json.load(f)
        file_path = settings.get("Gamelist")
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

    auto_save = tk.BooleanVar(value=settings.get("AutoSave", False))

    def save_to_file():
        text = text_widget.get("1.0", "end-1c")
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(text)
            load_from_file_callback(file_path)
            status_label.config(text="Saved", fg="#3cb371")
            countdown_label.config(text="", fg=fg_color)
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file.\n{e}")

    countdown_seconds_default = 5
    countdown_current = 0
    countdown_job = None

    def perform_auto_save():
        text = text_widget.get("1.0", "end-1c")
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(text)
            load_from_file_callback(file_path)
            status_label.config(text="Auto Saved", fg="#3cb371")
            countdown_label.config(text="", fg=fg_color)
        except Exception:
            pass

    def countdown_tick():
        nonlocal countdown_current, countdown_job
        countdown_current -= 1
        if countdown_current > 0:
            countdown_label.config(text=f"AutoSave in: {countdown_current}s", fg="#FFD700")
            countdown_job = popup.after(1000, countdown_tick)
        else:
            countdown_label.config(text="", fg=fg_color)
            countdown_job = None
            perform_auto_save()

    def start_countdown():
        nonlocal countdown_current, countdown_job
        if not auto_save.get():
            return
        stop_countdown()
        countdown_current = countdown_seconds_default
        countdown_label.config(text=f"AutoSave in: {countdown_current}s", fg="#FFD700")
        status_label.config(text="AutoSave pending", fg="#FFD700")
        countdown_job = popup.after(1000, countdown_tick)

    def stop_countdown():
        nonlocal countdown_job
        if countdown_job:
            try:
                popup.after_cancel(countdown_job)
            except Exception:
                pass
            countdown_job = None
        countdown_label.config(text="", fg=fg_color)

    def on_change(event=None):
        update_line_numbers()
        status_label.config(text="Modified", fg="#ff7f7f")
        if auto_save.get():
            start_countdown()
        else:
            stop_countdown()
        try:
            text_widget.edit_modified(False)
        except Exception:
            pass

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

    def toggle_auto_save():
        auto_save.set(not auto_save.get())
        auto_save_btn.config(
            text=f"Auto Save: {'ON' if auto_save.get() else 'OFF'}",
            bg="#3cb371" if auto_save.get() else "#777"
        )
        try:
            settings["AutoSave"] = auto_save.get()
            with open(settings_path, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            messagebox.showerror("Error", f"Could not update settings.json\n{e}")
        if auto_save.get():
            start_countdown()
        else:
            stop_countdown()

    def on_close():
        stop_countdown()
        popup.destroy()

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

    auto_save_btn = tk.Button(
        button_frame,
        text=f"Auto Save: {'ON' if auto_save.get() else 'OFF'}",
        width=14,
        bg="#3cb371" if auto_save.get() else "#777",
        fg="white",
        command=toggle_auto_save
    )
    auto_save_btn.pack(side="left", padx=10, pady=5)

    countdown_label = tk.Label(button_frame, text="", bg=bg_color, fg=fg_color, font=("Consolas", 12))
    countdown_label.pack(side="left", padx=(5, 0))

    status_label = tk.Label(button_frame, text="", bg=bg_color, fg=fg_color, font=("Consolas", 12))
    status_label.pack(side="left", padx=(10, 0))

    update_line_numbers()

    if auto_save.get():
        start_countdown()
    else:
        countdown_label.config(text="", fg=fg_color)

    popup.protocol("WM_DELETE_WINDOW", on_close)
    popup.transient(parent)
    popup.grab_set()

    def center_popup():
        popup.update_idletasks()
        pw = popup.winfo_width()
        ph = popup.winfo_height()
        try:
            px = parent.winfo_rootx()
            py = parent.winfo_rooty()
            pw_parent = parent.winfo_width()
            ph_parent = parent.winfo_height()
            x = px + (pw_parent // 2) - (pw // 2)
            y = py + (ph_parent // 2) - (ph // 2)
        except Exception:
            screen_w = popup.winfo_screenwidth()
            screen_h = popup.winfo_screenheight()
            x = (screen_w // 2) - (pw // 2)
            y = (screen_h // 2) - (ph // 2)
        popup.geometry(f"+{x}+{y}")

    center_popup()
    return popup
