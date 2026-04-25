import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta
import os

try:
    from tkcalendar import DateEntry
except ImportError:
    DateEntry = None


def safe_float(value, default=0.0):
    try:
        return float(value)
    except Exception:
        return default


def parse_datetime(value):
    if not value:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, str):
        value = value.strip()
        if not value:
            return None
        try:
            return datetime.fromisoformat(value)
        except Exception:
            pass
        for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                return datetime.strptime(value, fmt)
            except Exception:
                continue
    return None


def serialize_datetime(value):
    if isinstance(value, datetime):
        return value.isoformat(timespec="seconds")
    return None


def center_window(window, width, height):
    window.update_idletasks()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")


class SectionDialog:
    def __init__(self, parent, app, section_idx=None):
        self.parent = parent
        self.app = app
        self.section_idx = section_idx
        self.editing = section_idx is not None
        self.section = app.sections[section_idx] if self.editing else app.make_default_section()
        self.selected_date = None

        dialog = tk.Toplevel(parent)
        dialog.title("Edit Section" if self.editing else "Add Section")
        dialog.geometry("560x860")
        dialog.resizable(False, False)
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            dialog.iconbitmap(os.path.join(base_dir, "icon.ico"))
        except Exception:
            pass

        dialog.configure(bg=app.bg_color)
        dialog.grab_set()
        center_window(dialog, 560, 860)

        self.dialog = dialog
        self.bg_color = app.bg_color
        self.fg_color = app.fg_color
        self.section_bg = app.section_bg
        self.btn_color = app.btn_color

        top_wrap = tk.Frame(dialog, bg=self.bg_color)
        top_wrap.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(top_wrap, text="Section Name:", font=("Arial", 10), bg=self.bg_color, fg=self.fg_color).pack(anchor="w")

        self.name_entry = tk.Entry(
            top_wrap,
            font=("Arial", 10),
            bg=self.section_bg,
            fg=self.fg_color,
            insertbackground=self.fg_color
        )
        self.name_entry.pack(fill="x", pady=(3, 10))
        self.name_entry.insert(0, self.section.get("section", ""))

        tk.Label(top_wrap, text="Checklist Type:", font=("Arial", 10), bg=self.bg_color, fg=self.fg_color).pack(anchor="w")

        self.checklist_type_var = tk.StringVar(value=self.section.get("checklist_type", "Temporary"))
        type_frame = tk.Frame(top_wrap, bg=self.bg_color)
        type_frame.pack(fill="x", pady=(2, 6))

        tk.Radiobutton(
            type_frame,
            text="Temporary",
            variable=self.checklist_type_var,
            value="Temporary",
            bg=self.bg_color,
            fg=self.fg_color,
            selectcolor=self.bg_color,
            activebackground=self.bg_color,
            command=lambda: self.update_visibility()
        ).pack(side="left", padx=(0, 12))

        tk.Radiobutton(
            type_frame,
            text="Permanent",
            variable=self.checklist_type_var,
            value="Permanent",
            bg=self.bg_color,
            fg=self.fg_color,
            selectcolor=self.bg_color,
            activebackground=self.bg_color,
            command=lambda: self.update_visibility()
        ).pack(side="left")

        self.schedule_var = tk.BooleanVar(value=bool(self.section.get("schedule_enabled", False)))

        self.schedule_check = tk.Checkbutton(
            top_wrap,
            text="Schedule",
            variable=self.schedule_var,
            bg=self.bg_color,
            fg=self.fg_color,
            selectcolor=self.bg_color,
            activebackground=self.bg_color,
            command=lambda: self.update_visibility()
        )

        self.schedule_frame = tk.Frame(top_wrap, bg=self.bg_color)

        self.schedule_mode_var = tk.StringVar(value="Repeat?")
        self.repeat_option = tk.Radiobutton(
            self.schedule_frame,
            text="Repeat?",
            variable=self.schedule_mode_var,
            value="Repeat?",
            bg=self.bg_color,
            fg=self.fg_color,
            selectcolor=self.bg_color,
            activebackground=self.bg_color,
            command=lambda: self.update_visibility()
        )

        self.set_date_option = tk.Radiobutton(
            self.schedule_frame,
            text="Set Date",
            variable=self.schedule_mode_var,
            value="Set Date",
            bg=self.bg_color,
            fg=self.fg_color,
            selectcolor=self.bg_color,
            activebackground=self.bg_color,
            command=lambda: self.update_visibility()
        )

        self.repeat_days_row = tk.Frame(self.schedule_frame, bg=self.bg_color)

        tk.Label(
            self.repeat_days_row,
            text="How much days should you repeated?",
            font=("Arial", 10),
            bg=self.bg_color,
            fg=self.fg_color
        ).pack(anchor="w")

        self.repeat_days_entry = tk.Entry(
            self.repeat_days_row,
            font=("Arial", 10),
            bg=self.section_bg,
            fg=self.fg_color,
            insertbackground=self.fg_color,
            width=20
        )
        self.repeat_days_entry.pack(anchor="w", pady=(3, 4))
        self.repeat_days_entry.insert(0, str(self.section.get("repeat_days", 1.0)))

        self.time_left_row = tk.Frame(self.schedule_frame, bg=self.bg_color)

        tk.Label(
            self.time_left_row,
            text="How much day left?",
            font=("Arial", 10),
            bg=self.bg_color,
            fg=self.fg_color
        ).pack(anchor="w")

        self.time_left_entry = tk.Entry(
            self.time_left_row,
            font=("Arial", 10),
            bg=self.section_bg,
            fg=self.fg_color,
            insertbackground=self.fg_color,
            width=20
        )
        self.time_left_entry.pack(anchor="w", pady=(3, 8))
        self.time_left_entry.insert(0, str(self.section.get("time_left_days", self.section.get("repeat_days", 1.0))))

        self.date_picker_row = tk.Frame(self.schedule_frame, bg=self.bg_color)

        tk.Label(
            self.date_picker_row,
            text="Select Target Date:",
            font=("Arial", 10),
            bg=self.bg_color,
            fg=self.fg_color
        ).pack(anchor="w")

        self.date_display_var = tk.StringVar(value="No date selected")

        date_button_frame = tk.Frame(self.date_picker_row, bg=self.bg_color)
        date_button_frame.pack(anchor="w", pady=(3, 0))

        tk.Button(
            date_button_frame,
            text="📅 Pick Date",
            font=("Arial", 9),
            bg=self.btn_color,
            fg="white",
            width=12,
            command=self.pick_date
        ).pack(side="left", padx=(0, 8))

        tk.Label(
            date_button_frame,
            textvariable=self.date_display_var,
            font=("Arial", 10),
            bg=self.bg_color,
            fg="#90EE90"
        ).pack(side="left", pady=(3, 8))

        tk.Label(
            top_wrap,
            text="Items (one per line):",
            font=("Arial", 10),
            bg=self.bg_color,
            fg=self.fg_color
        ).pack(anchor="w", pady=(8, 0))

        self.items_text = tk.Text(
            top_wrap,
            height=12,
            font=("Arial", 10),
            bg=self.section_bg,
            fg=self.fg_color,
            insertbackground=self.fg_color,
            wrap="word"
        )
        self.items_text.pack(fill="both", expand=True, pady=(3, 0))
        if self.section.get("items"):
            self.items_text.insert("1.0", "\n".join(item.get("text", "") for item in self.section.get("items", [])))

        button_frame = tk.Frame(dialog, bg=self.bg_color)
        button_frame.pack(fill="x", padx=10, pady=(0, 10))

        tk.Button(
            button_frame,
            text="💾 Save",
            font=("Arial", 10, "bold"),
            bg=app.btn_color,
            fg="white",
            width=14,
            command=self.on_save
        ).pack(side="left")

        tk.Button(
            button_frame,
            text="Cancel",
            font=("Arial", 10, "bold"),
            bg="#666",
            fg="white",
            width=14,
            command=self.on_close
        ).pack(side="right")

        if self.editing and self.section.get("schedule_enabled"):
            if self.section.get("repeat_enabled"):
                self.schedule_mode_var.set("Repeat?")
            else:
                self.schedule_mode_var.set("Set Date")
                if self.section.get("started_at"):
                    self.selected_date = parse_datetime(self.section.get("started_at"))
                    if self.selected_date:
                        self.date_display_var.set(self.selected_date.strftime("%Y-%m-%d"))

        self.update_visibility()
        dialog.protocol("WM_DELETE_WINDOW", self.on_close)

    def pick_date(self):
        if DateEntry is None:
            messagebox.showerror("Error", "tkcalendar is not installed. Please install it to use date picker.")
            return

        date_window = tk.Toplevel(self.dialog)
        date_window.title("Select Date")
        date_window.geometry("350x400")
        date_window.resizable(False, False)
        date_window.configure(bg=self.bg_color)
        date_window.grab_set()
        center_window(date_window, 350, 400)

        cal_frame = tk.Frame(date_window, bg=self.bg_color)
        cal_frame.pack(fill="both", expand=True, padx=10, pady=10)

        initial_date = self.selected_date if self.selected_date else datetime.now().date()

        cal = DateEntry(
            cal_frame,
            year=initial_date.year,
            month=initial_date.month,
            day=initial_date.day,
            background=self.btn_color,
            foreground="white",
            borderwidth=2,
            font=("Arial", 10)
        )
        cal.pack(pady=10)

        def confirm_date():
            self.selected_date = datetime.combine(cal.get_date(), datetime.min.time())
            self.date_display_var.set(self.selected_date.strftime("%Y-%m-%d"))
            date_window.destroy()

        tk.Button(
            cal_frame,
            text="✓ Confirm",
            font=("Arial", 10, "bold"),
            bg=self.btn_color,
            fg="white",
            command=confirm_date
        ).pack(pady=(10, 0))

    def update_visibility(self):
        is_permanent = self.checklist_type_var.get() == "Permanent"

        self.schedule_check.pack_forget()
        self.schedule_frame.pack_forget()
        self.repeat_option.pack_forget()
        self.set_date_option.pack_forget()
        self.repeat_days_row.pack_forget()
        self.time_left_row.pack_forget()
        self.date_picker_row.pack_forget()

        if is_permanent:
            self.schedule_check.pack(anchor="w", pady=(2, 2))
            if self.schedule_var.get():
                self.schedule_frame.pack(fill="x", pady=(4, 0))
                self.repeat_option.pack(anchor="w", pady=(0, 2))
                self.set_date_option.pack(anchor="w", pady=(0, 6))

                if self.schedule_mode_var.get() == "Repeat?":
                    self.repeat_days_row.pack(fill="x", pady=(2, 0))
                    self.time_left_row.pack(fill="x", pady=(2, 0))
                else:
                    self.date_picker_row.pack(fill="x", pady=(2, 0))
        else:
            self.schedule_var.set(False)
            self.selected_date = None
            self.date_display_var.set("No date selected")

    def perform_save(self):
        name = self.name_entry.get().strip()
        if not name:
            messagebox.showwarning("Empty", "Section name cannot be empty.")
            return False

        if any(
            idx != self.section_idx and s.get("section") == name
            for idx, s in enumerate(self.app.sections)
        ):
            messagebox.showwarning("Duplicate", "This section already exists.")
            return False

        checklist_type = self.checklist_type_var.get()
        schedule_enabled = checklist_type == "Permanent" and self.schedule_var.get()
        repeat_enabled = False
        repeat_days = 1.0
        time_left_days = 0.0
        started_at = None

        if schedule_enabled:
            if self.schedule_mode_var.get() == "Repeat?":
                repeat_enabled = True
                try:
                    repeat_days = float(self.repeat_days_entry.get().strip())
                    if repeat_days <= 0:
                        raise ValueError
                except Exception:
                    messagebox.showwarning("Invalid", "How much days should you repeated? must be greater than 0.")
                    return False
                try:
                    time_left_days = float(self.time_left_entry.get().strip())
                    if time_left_days < 0:
                        raise ValueError
                except Exception:
                    messagebox.showwarning("Invalid", "How much day left? must be 0 or greater.")
                    return False
                started_at = serialize_datetime(datetime.now())
            else:
                if not self.selected_date:
                    messagebox.showwarning("Invalid", "Please select a target date.")
                    return False
                time_left_days = (self.selected_date - datetime.now()).total_seconds() / 86400.0
                started_at = serialize_datetime(self.selected_date)

        raw_lines = self.items_text.get("1.0", tk.END).splitlines()
        existing_items = self.section.get("items", []) if self.editing else []
        defaults = {
            "checklist_type": checklist_type,
            "schedule_enabled": schedule_enabled,
            "repeat_enabled": repeat_enabled,
            "repeat_days": repeat_days if schedule_enabled else 1.0,
            "time_left_days": time_left_days if schedule_enabled else 0.0,
            "started_at": started_at
        }
        new_items = self.app.build_items_from_lines(raw_lines, existing_items, defaults)
        new_items = self.app.apply_section_settings_to_items(new_items, defaults)

        updated_section = {
            "section": name,
            "checklist_type": checklist_type,
            "schedule_enabled": schedule_enabled,
            "repeat_enabled": repeat_enabled,
            "repeat_days": repeat_days if schedule_enabled else 1.0,
            "time_left_days": time_left_days if schedule_enabled else 0.0,
            "started_at": started_at,
            "items": new_items
        }

        if self.editing:
            self.app.sections[self.section_idx] = updated_section
        else:
            self.app.sections.append(updated_section)

        self.app.sync_schedule_state()
        self.app.refresh_display()
        self.app.save_data()
        return True

    def on_save(self):
        if self.perform_save():
            self.on_close()

    def on_close(self):
        try:
            self.dialog.grab_release()
        except Exception:
            pass
        self.dialog.destroy()


class ItemDialog:
    def __init__(self, parent, app, section_idx, item_idx=None):
        self.parent = parent
        self.app = app
        self.section_idx = section_idx
        self.item_idx = item_idx
        self.editing = item_idx is not None
        self.section = app.sections[section_idx]
        self.item = self.section["items"][item_idx] if self.editing else app.make_default_item(self.section)
        self.selected_date = None

        dialog = tk.Toplevel(parent)
        dialog.title("Edit Item" if self.editing else "Add Item")
        dialog.geometry("560x740")
        dialog.resizable(False, False)
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            dialog.iconbitmap(os.path.join(base_dir, "icon.ico"))
        except Exception:
            pass

        dialog.configure(bg=app.bg_color)
        dialog.grab_set()
        center_window(dialog, 560, 740)

        self.dialog = dialog
        self.bg_color = app.bg_color
        self.fg_color = app.fg_color
        self.section_bg = app.section_bg
        self.btn_color = app.btn_color

        top_wrap = tk.Frame(dialog, bg=self.bg_color)
        top_wrap.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(top_wrap, text="Item Text:", font=("Arial", 10), bg=self.bg_color, fg=self.fg_color).pack(anchor="w")

        self.text = tk.Text(
            top_wrap,
            font=("Arial", 10),
            bg=self.section_bg,
            fg=self.fg_color,
            insertbackground=self.fg_color,
            wrap="word",
            height=4
        )
        self.text.pack(fill="x", pady=(3, 10))
        self.text.insert("1.0", self.item.get("text", ""))
        self.text.focus()

        tk.Label(top_wrap, text="Checklist Type:", font=("Arial", 10), bg=self.bg_color, fg=self.fg_color).pack(anchor="w")

        self.checklist_type_var = tk.StringVar(value=self.item.get("checklist_type", "Temporary"))
        type_frame = tk.Frame(top_wrap, bg=self.bg_color)
        type_frame.pack(fill="x", pady=(2, 6))

        tk.Radiobutton(
            type_frame,
            text="Temporary",
            variable=self.checklist_type_var,
            value="Temporary",
            bg=self.bg_color,
            fg=self.fg_color,
            selectcolor=self.bg_color,
            activebackground=self.bg_color,
            command=lambda: self.update_visibility()
        ).pack(side="left", padx=(0, 12))

        tk.Radiobutton(
            type_frame,
            text="Permanent",
            variable=self.checklist_type_var,
            value="Permanent",
            bg=self.bg_color,
            fg=self.fg_color,
            selectcolor=self.bg_color,
            activebackground=self.bg_color,
            command=lambda: self.update_visibility()
        ).pack(side="left")

        self.schedule_var = tk.BooleanVar(value=bool(self.item.get("schedule_enabled", False)))

        self.schedule_check = tk.Checkbutton(
            top_wrap,
            text="Schedule",
            variable=self.schedule_var,
            bg=self.bg_color,
            fg=self.fg_color,
            selectcolor=self.bg_color,
            activebackground=self.bg_color,
            command=lambda: self.update_visibility()
        )

        self.schedule_frame = tk.Frame(top_wrap, bg=self.bg_color)

        self.schedule_mode_var = tk.StringVar(value="Repeat?")
        self.repeat_option = tk.Radiobutton(
            self.schedule_frame,
            text="Repeat?",
            variable=self.schedule_mode_var,
            value="Repeat?",
            bg=self.bg_color,
            fg=self.fg_color,
            selectcolor=self.bg_color,
            activebackground=self.bg_color,
            command=lambda: self.update_visibility()
        )

        self.set_date_option = tk.Radiobutton(
            self.schedule_frame,
            text="Set Date",
            variable=self.schedule_mode_var,
            value="Set Date",
            bg=self.bg_color,
            fg=self.fg_color,
            selectcolor=self.bg_color,
            activebackground=self.bg_color,
            command=lambda: self.update_visibility()
        )

        self.repeat_days_row = tk.Frame(self.schedule_frame, bg=self.bg_color)
        tk.Label(
            self.repeat_days_row,
            text="How much days should you repeated?",
            font=("Arial", 10),
            bg=self.bg_color,
            fg=self.fg_color
        ).pack(anchor="w")
        self.repeat_days_entry = tk.Entry(
            self.repeat_days_row,
            font=("Arial", 10),
            bg=self.section_bg,
            fg=self.fg_color,
            insertbackground=self.fg_color,
            width=20
        )
        self.repeat_days_entry.pack(anchor="w", pady=(3, 4))
        self.repeat_days_entry.insert(0, str(self.item.get("repeat_days", 1.0)))

        self.time_left_row = tk.Frame(self.schedule_frame, bg=self.bg_color)
        tk.Label(
            self.time_left_row,
            text="How much day left?",
            font=("Arial", 10),
            bg=self.bg_color,
            fg=self.fg_color
        ).pack(anchor="w")
        self.time_left_entry = tk.Entry(
            self.time_left_row,
            font=("Arial", 10),
            bg=self.section_bg,
            fg=self.fg_color,
            insertbackground=self.fg_color,
            width=20
        )
        self.time_left_entry.pack(anchor="w", pady=(3, 8))
        self.time_left_entry.insert(0, str(self.item.get("time_left_days", self.item.get("repeat_days", 1.0))))

        self.date_picker_row = tk.Frame(self.schedule_frame, bg=self.bg_color)
        tk.Label(
            self.date_picker_row,
            text="Select Target Date:",
            font=("Arial", 10),
            bg=self.bg_color,
            fg=self.fg_color
        ).pack(anchor="w")

        self.date_display_var = tk.StringVar(value="No date selected")

        date_button_frame = tk.Frame(self.date_picker_row, bg=self.bg_color)
        date_button_frame.pack(anchor="w", pady=(3, 0))

        tk.Button(
            date_button_frame,
            text="📅 Pick Date",
            font=("Arial", 9),
            bg=self.btn_color,
            fg="white",
            width=12,
            command=self.pick_date
        ).pack(side="left", padx=(0, 8))

        tk.Label(
            date_button_frame,
            textvariable=self.date_display_var,
            font=("Arial", 10),
            bg=self.bg_color,
            fg="#90EE90"
        ).pack(side="left", pady=(3, 8))

        button_frame = tk.Frame(dialog, bg=self.bg_color)
        button_frame.pack(fill="x", padx=10, pady=(0, 10))

        tk.Button(
            button_frame,
            text="💾 Save",
            font=("Arial", 10, "bold"),
            bg=app.btn_color,
            fg="white",
            width=14,
            command=self.on_save
        ).pack(side="left")

        tk.Button(
            button_frame,
            text="Cancel",
            font=("Arial", 10, "bold"),
            bg="#666",
            fg="white",
            width=14,
            command=self.on_close
        ).pack(side="right")

        if self.editing and self.item.get("schedule_enabled"):
            if self.item.get("repeat_enabled"):
                self.schedule_mode_var.set("Repeat?")
            else:
                self.schedule_mode_var.set("Set Date")
                if self.item.get("started_at"):
                    self.selected_date = parse_datetime(self.item.get("started_at"))
                    if self.selected_date:
                        self.date_display_var.set(self.selected_date.strftime("%Y-%m-%d"))

        self.update_visibility()
        dialog.protocol("WM_DELETE_WINDOW", self.on_close)

    def pick_date(self):
        if DateEntry is None:
            messagebox.showerror("Error", "tkcalendar is not installed. Please install it to use date picker.")
            return

        date_window = tk.Toplevel(self.dialog)
        date_window.title("Select Date")
        date_window.geometry("350x400")
        date_window.resizable(False, False)
        date_window.configure(bg=self.bg_color)
        date_window.grab_set()
        center_window(date_window, 350, 400)

        cal_frame = tk.Frame(date_window, bg=self.bg_color)
        cal_frame.pack(fill="both", expand=True, padx=10, pady=10)

        initial_date = self.selected_date if self.selected_date else datetime.now().date()

        cal = DateEntry(
            cal_frame,
            year=initial_date.year,
            month=initial_date.month,
            day=initial_date.day,
            background=self.btn_color,
            foreground="white",
            borderwidth=2,
            font=("Arial", 10)
        )
        cal.pack(pady=10)

        def confirm_date():
            self.selected_date = datetime.combine(cal.get_date(), datetime.min.time())
            self.date_display_var.set(self.selected_date.strftime("%Y-%m-%d"))
            date_window.destroy()

        tk.Button(
            cal_frame,
            text="✓ Confirm",
            font=("Arial", 10, "bold"),
            bg=self.btn_color,
            fg="white",
            command=confirm_date
        ).pack(pady=(10, 0))

    def update_visibility(self):
        is_permanent = self.checklist_type_var.get() == "Permanent"

        self.schedule_check.pack_forget()
        self.schedule_frame.pack_forget()
        self.repeat_option.pack_forget()
        self.set_date_option.pack_forget()
        self.repeat_days_row.pack_forget()
        self.time_left_row.pack_forget()
        self.date_picker_row.pack_forget()

        if is_permanent:
            self.schedule_check.pack(anchor="w", pady=(2, 2))
            if self.schedule_var.get():
                self.schedule_frame.pack(fill="x", pady=(4, 0))
                self.repeat_option.pack(anchor="w", pady=(0, 2))
                self.set_date_option.pack(anchor="w", pady=(0, 6))

                if self.schedule_mode_var.get() == "Repeat?":
                    self.repeat_days_row.pack(fill="x", pady=(2, 0))
                    self.time_left_row.pack(fill="x", pady=(2, 0))
                else:
                    self.date_picker_row.pack(fill="x", pady=(2, 0))
        else:
            self.schedule_var.set(False)
            self.selected_date = None
            self.date_display_var.set("No date selected")

    def perform_save(self):
        value = self.text.get("1.0", tk.END).strip()
        if not value:
            messagebox.showwarning("Empty", "Item cannot be empty.")
            return False

        checklist_type = self.checklist_type_var.get()
        schedule_enabled = checklist_type == "Permanent" and self.schedule_var.get()
        repeat_enabled = False
        repeat_days = 1.0
        time_left_days = 0.0
        started_at = None

        if schedule_enabled:
            if self.schedule_mode_var.get() == "Repeat?":
                repeat_enabled = True
                try:
                    repeat_days = float(self.repeat_days_entry.get().strip())
                    if repeat_days <= 0:
                        raise ValueError
                except Exception:
                    messagebox.showwarning("Invalid", "How much days should you repeated? must be greater than 0.")
                    return False
                try:
                    time_left_days = float(self.time_left_entry.get().strip())
                    if time_left_days < 0:
                        raise ValueError
                except Exception:
                    messagebox.showwarning("Invalid", "How much day left? must be 0 or greater.")
                    return False
                started_at = serialize_datetime(datetime.now())
            else:
                if not self.selected_date:
                    messagebox.showwarning("Invalid", "Please select a target date.")
                    return False
                time_left_days = (self.selected_date - datetime.now()).total_seconds() / 86400.0
                started_at = serialize_datetime(self.selected_date)

        updated_item = {
            "text": value,
            "checked": bool(self.item.get("checked", False)),
            "checked_at": self.item.get("checked_at"),
            "checklist_type": checklist_type,
            "schedule_enabled": schedule_enabled,
            "repeat_enabled": repeat_enabled,
            "repeat_days": repeat_days if schedule_enabled else 1.0,
            "time_left_days": time_left_days if schedule_enabled else 0.0,
            "started_at": started_at
        }

        if checklist_type == "Temporary":
            updated_item["schedule_enabled"] = False
            updated_item["repeat_enabled"] = False
            updated_item["repeat_days"] = 1.0
            updated_item["time_left_days"] = 0.0
            updated_item["started_at"] = None
            updated_item["checked"] = False
            updated_item["checked_at"] = None

        if self.editing:
            self.app.sections[self.section_idx]["items"][self.item_idx] = updated_item
        else:
            self.app.sections[self.section_idx]["items"].append(updated_item)

        self.app.sync_schedule_state()
        self.app.refresh_display()
        self.app.save_data()
        return True

    def on_save(self):
        if self.perform_save():
            self.on_close()

    def on_close(self):
        try:
            self.dialog.grab_release()
        except Exception:
            pass
        self.dialog.destroy()
