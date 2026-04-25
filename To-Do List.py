import tkinter as tk
from tkinter import messagebox
from datetime import datetime, date, timedelta
import calendar
import os
import json
import sys
import importlib.util

if getattr(sys, "frozen", False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

todo_json_path = os.path.join(base_dir, "To-Do List.json")

spec = importlib.util.spec_from_file_location("modal", os.path.join(base_dir, "To-Do List Modal.py"))
modal = importlib.util.module_from_spec(spec)
spec.loader.exec_module(modal)
SectionDialog = modal.SectionDialog
ItemDialog = modal.ItemDialog


def parse_datetime(value):
    if not value:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
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


def safe_float(value, default=0.0):
    try:
        return float(value)
    except Exception:
        return default


def shorten_text(text, limit=48):
    text = str(text or "")
    if len(text) <= limit:
        return text
    return text[: limit - 1] + "…"


class HoverTooltip:
    def __init__(self, widget, text_provider):
        self.widget = widget
        self.text_provider = text_provider
        self.tipwindow = None
        self.label = None
        self.after_id = None
        self.x = 0
        self.y = 0

        widget.bind("<Enter>", self.enter, add="+")
        widget.bind("<Leave>", self.leave, add="+")
        widget.bind("<Motion>", self.motion, add="+")
        widget.bind("<Destroy>", self.destroyed, add="+")

    def enter(self, event=None):
        self.show()

    def motion(self, event=None):
        self.x = event.x_root + 16
        self.y = event.y_root + 16
        if self.tipwindow is not None:
            self.tipwindow.geometry(f"+{self.x}+{self.y}")

    def leave(self, event=None):
        self.hide()

    def destroyed(self, event=None):
        self.hide()

    def show(self):
        if self.tipwindow is not None:
            return
        text = self.text_provider()
        if not text:
            return
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_attributes("-topmost", True)
        tw.geometry(f"+{self.x}+{self.y}")
        self.label = tk.Label(
            tw,
            text=text,
            justify="left",
            background="#222222",
            foreground="white",
            relief="solid",
            borderwidth=1,
            font=("Arial", 9),
        )
        self.label.pack(ipadx=6, ipady=3)
        self._tick()

    def _tick(self):
        if self.tipwindow is None or self.label is None:
            return
        try:
            self.label.config(text=self.text_provider())
        except Exception:
            self.hide()
            return
        self.after_id = self.widget.after(1000, self._tick)

    def hide(self):
        if self.after_id is not None:
            try:
                self.widget.after_cancel(self.after_id)
            except Exception:
                pass
            self.after_id = None
        if self.tipwindow is not None:
            try:
                self.tipwindow.destroy()
            except Exception:
                pass
            self.tipwindow = None
            self.label = None


class DatePickerDialog:
    def __init__(self, parent, initial=None):
        self.parent = parent
        self.result = None

        initial_dt = parse_datetime(initial) or datetime.now()
        self.selected_date = initial_dt
        self.current_year = initial_dt.year
        self.current_month = initial_dt.month

        self.hour_var = tk.StringVar(value=f"{initial_dt.hour:02d}")
        self.minute_var = tk.StringVar(value=f"{initial_dt.minute:02d}")

        self.top = tk.Toplevel(parent)
        self.top.title("Select Date")
        self.top.resizable(False, False)
        self.top.grab_set()
        self.top.configure(bg="#1e1e1e")

        try:
            self.top.iconbitmap(os.path.join(base_dir, "icon.ico"))
        except Exception:
            pass

        header = tk.Frame(self.top, bg="#1e1e1e")
        header.pack(fill="x", padx=10, pady=(10, 6))

        tk.Button(
            header,
            text="◀",
            width=3,
            bg="#444",
            fg="white",
            activebackground="#333",
            command=self.go_prev_month,
        ).pack(side="left")

        self.title_label = tk.Label(
            header,
            text="",
            font=("Arial", 11, "bold"),
            bg="#1e1e1e",
            fg="white",
        )
        self.title_label.pack(side="left", expand=True)

        tk.Button(
            header,
            text="▶",
            width=3,
            bg="#444",
            fg="white",
            activebackground="#333",
            command=self.go_next_month,
        ).pack(side="right")

        self.calendar_frame = tk.Frame(self.top, bg="#1e1e1e")
        self.calendar_frame.pack(padx=10, pady=6)

        weekdays = tk.Frame(self.calendar_frame, bg="#1e1e1e")
        weekdays.pack(fill="x")

        for text in ["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]:
            tk.Label(
                weekdays,
                text=text,
                width=4,
                font=("Arial", 9, "bold"),
                bg="#1e1e1e",
                fg="#cfcfcf",
            ).pack(side="left", padx=1, pady=2)

        self.days_frame = tk.Frame(self.calendar_frame, bg="#1e1e1e")
        self.days_frame.pack()

        time_frame = tk.Frame(self.top, bg="#1e1e1e")
        time_frame.pack(fill="x", padx=10, pady=(6, 0))

        tk.Label(
            time_frame,
            text="Time (HH:MM)",
            font=("Arial", 10),
            bg="#1e1e1e",
            fg="white",
        ).pack(anchor="w")

        time_inner = tk.Frame(time_frame, bg="#1e1e1e")
        time_inner.pack(anchor="w", pady=(2, 0))

        tk.Entry(
            time_inner,
            textvariable=self.hour_var,
            width=4,
            font=("Arial", 10),
            bg="#2d2d2d",
            fg="white",
            insertbackground="white",
            justify="center",
        ).pack(side="left")

        tk.Label(
            time_inner,
            text=":",
            font=("Arial", 10, "bold"),
            bg="#1e1e1e",
            fg="white",
        ).pack(side="left", padx=2)

        tk.Entry(
            time_inner,
            textvariable=self.minute_var,
            width=4,
            font=("Arial", 10),
            bg="#2d2d2d",
            fg="white",
            insertbackground="white",
            justify="center",
        ).pack(side="left")

        footer = tk.Frame(self.top, bg="#1e1e1e")
        footer.pack(fill="x", padx=10, pady=10)

        tk.Button(
            footer,
            text="Today",
            width=10,
            bg="#444",
            fg="white",
            activebackground="#333",
            command=self.select_today,
        ).pack(side="left")

        tk.Button(
            footer,
            text="Cancel",
            width=10,
            bg="#666",
            fg="white",
            activebackground="#555",
            command=self.close,
        ).pack(side="right")

        self.draw_calendar()
        self.center()

        self.top.protocol("WM_DELETE_WINDOW", self.close)
        self.top.bind("<Return>", lambda e: self.confirm())
        self.top.bind("<Escape>", lambda e: self.close())

    def center(self):
        self.top.update_idletasks()
        width = 330
        height = 370
        x = (self.top.winfo_screenwidth() // 2) - (width // 2)
        y = (self.top.winfo_screenheight() // 2) - (height // 2)
        self.top.geometry(f"{width}x{height}+{x}+{y}")

    def go_prev_month(self):
        if self.current_month == 1:
            self.current_month = 12
            self.current_year -= 1
        else:
            self.current_month -= 1
        self.draw_calendar()

    def go_next_month(self):
        if self.current_month == 12:
            self.current_month = 1
            self.current_year += 1
        else:
            self.current_month += 1
        self.draw_calendar()

    def select_today(self):
        now = datetime.now()
        self.current_year = now.year
        self.current_month = now.month
        self.selected_date = now
        self.hour_var.set(f"{now.hour:02d}")
        self.minute_var.set(f"{now.minute:02d}")
        self.draw_calendar()

    def draw_calendar(self):
        for widget in self.days_frame.winfo_children():
            widget.destroy()

        self.title_label.config(text=f"{calendar.month_name[self.current_month]} {self.current_year}")

        first_weekday, last_day = calendar.monthrange(self.current_year, self.current_month)

        weeks = []
        week = []
        for _ in range(first_weekday):
            week.append(None)
        for day_num in range(1, last_day + 1):
            week.append(day_num)
            if len(week) == 7:
                weeks.append(week)
                week = []
        if week:
            while len(week) < 7:
                week.append(None)
            weeks.append(week)

        today = date.today()

        for week in weeks:
            row = tk.Frame(self.days_frame, bg="#1e1e1e")
            row.pack()
            for day_num in week:
                if day_num is None:
                    tk.Label(
                        row,
                        text=" ",
                        width=4,
                        height=2,
                        font=("Arial", 9),
                        bg="#1e1e1e",
                        fg="white",
                    ).pack(side="left", padx=1, pady=1)
                else:
                    is_selected = (
                        self.selected_date.year == self.current_year
                        and self.selected_date.month == self.current_month
                        and self.selected_date.day == day_num
                    )
                    is_today = (
                        today.year == self.current_year
                        and today.month == self.current_month
                        and today.day == day_num
                    )

                    bg = "#3c8dbc" if is_selected else "#2d2d2d"
                    if is_today and not is_selected:
                        bg = "#444"

                    tk.Button(
                        row,
                        text=str(day_num),
                        width=4,
                        height=2,
                        bg=bg,
                        fg="white",
                        activebackground="#005b29" if is_selected else "#444",
                        command=lambda d=day_num: self.pick_day(d),
                    ).pack(side="left", padx=1, pady=1)

    def pick_day(self, day_num):
        try:
            hour = int(self.hour_var.get().strip())
            minute = int(self.minute_var.get().strip())
            if not (0 <= hour <= 23 and 0 <= minute <= 59):
                raise ValueError
        except Exception:
            hour = self.selected_date.hour
            minute = self.selected_date.minute

        self.result = datetime(
            self.current_year,
            self.current_month,
            day_num,
            hour,
            minute,
            0,
            0,
        )
        self.close()

    def confirm(self):
        try:
            hour = int(self.hour_var.get().strip())
            minute = int(self.minute_var.get().strip())
            if not (0 <= hour <= 23 and 0 <= minute <= 59):
                raise ValueError
        except Exception:
            messagebox.showwarning("Invalid", "Time must be valid HH:MM.")
            return

        self.selected_date = self.selected_date.replace(
            hour=hour,
            minute=minute,
            second=0,
            microsecond=0,
        )
        self.result = self.selected_date
        self.close()

    def close(self):
        try:
            self.top.grab_release()
        except Exception:
            pass
        self.top.destroy()


class TodoListApp:
    SHARED_HEIGHT = 780

    def __init__(self, root, center=True, close_callback=None):
        self.root = root
        self.close_callback = close_callback
        self.schedule_job = None

        try:
            self.root.iconbitmap(os.path.join(base_dir, "icon.ico"))
        except Exception:
            pass

        self.root.title("To-Do List")
        self.root.geometry(f"600x{self.SHARED_HEIGHT}")
        self.root.resizable(False, False)

        self.bg_color = "#1e1e1e"
        self.fg_color = "#ffffff"
        self.list_bg = "#363636"
        self.section_bg = "#2d2d2d"
        self.btn_color = "#3c8dbc"
        self.edit_btn_color = "#005b29"

        self.default_item_bg = "#262626"
        self.checked_item_bg = "#123b22"
        self.expired_repeat_item_bg = "#423600"
        self.expired_norepeat_item_bg = "#5b0000"
        self.counting_item_bg = "#343000"

        self.root.configure(bg=self.bg_color)
        self.edit_mode = False
        self.sections = []

        self.drag_job = None
        self.drag_pending = None
        self.drag_active = False
        self.drag_kind = None
        self.drag_from = None
        self.drag_preview = None
        self.drag_preview_label = None
        self.drag_source_text = ""
        self.drag_hover_target = None

        self.load_data()
        self.sync_schedule_state()
        self.create_widgets()

        if center:
            self.center_window()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.bind_all("<B1-Motion>", self._global_drag_motion, add="+")
        self.root.bind_all("<ButtonRelease-1>", self._global_drag_release, add="+")
        self._start_schedule_refresh()

    def center_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.SHARED_HEIGHT
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def create_widgets(self):
        top_bar = tk.Frame(self.root, bg=self.bg_color)
        top_bar.pack(fill="x", padx=10, pady=(10, 0))

        tk.Label(
            top_bar,
            text="To-Do List",
            font=("Arial", 16, "bold"),
            bg=self.bg_color,
            fg=self.fg_color,
        ).pack(side="left")

        self.edit_mode_button = tk.Button(
            top_bar,
            text="✎",
            font=("Arial", 13, "bold"),
            bg="#666",
            fg="white",
            activebackground="#555",
            relief="ridge",
            bd=3,
            width=3,
            command=self.toggle_edit_mode,
        )
        self.edit_mode_button.pack(side="right")
        self.edit_mode_tooltip = HoverTooltip(self.edit_mode_button, lambda: "Edit Mode")

        control_frame = tk.Frame(self.root, bg=self.bg_color)
        control_frame.pack(fill="x", padx=10, pady=(8, 10))

        self.add_section_button = tk.Button(
            control_frame,
            text="＋ Add Section",
            font=("Arial", 10, "bold"),
            bg=self.edit_btn_color,
            fg="white",
            activebackground="#004620",
            command=lambda: self.open_section_dialog(),
        )

        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill="both", expand=True, padx=10, pady=0)

        self.canvas = tk.Canvas(main_frame, bg=self.list_bg, highlightthickness=0, height=500)
        self.canvas.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(main_frame, orient="vertical", command=self.canvas.yview)
        scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.scroll_container = tk.Frame(self.canvas, bg=self.list_bg)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scroll_container, anchor="nw")

        self.scroll_container.bind("<Configure>", self._update_scrollregion)
        self.canvas.bind("<Configure>", self._fit_canvas_width)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.refresh_display()

    def _fit_canvas_width(self, event=None):
        self.canvas.itemconfigure(self.canvas_window, width=self.canvas.winfo_width())

    def _update_scrollregion(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_mousewheel(self, event=None):
        if event is None:
            return
        bbox = self.canvas.bbox("all")
        if not bbox:
            return
        canvas_height = self.canvas.winfo_height()
        content_height = bbox[3] - bbox[1]
        if content_height <= canvas_height:
            return
        y1, y2 = self.canvas.yview()
        if event.delta > 0 and y1 <= 0:
            return
        if event.delta < 0 and y2 >= 1:
            return
        self.canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

    def _start_schedule_refresh(self):
        self._schedule_tick()

    def _schedule_tick(self):
        self.sync_schedule_state()
        self.refresh_display()
        self.schedule_job = self.root.after(30000, self._schedule_tick)

    def toggle_edit_mode(self):
        self.edit_mode = not self.edit_mode
        if self.edit_mode:
            self.edit_mode_button.config(text="✎", bg=self.edit_btn_color)
        else:
            self.edit_mode_button.config(text="✎", bg="#666")
        self._cancel_drag()
        self.refresh_display()

    def make_default_section(self):
        return {
            "section": "",
            "checklist_type": "Temporary",
            "schedule_enabled": False,
            "repeat_enabled": True,
            "repeat_days": 1.0,
            "time_left_days": 0.0,
            "started_at": None,
            "items": [],
        }

    def make_default_item(self, section=None):
        section = section or {}
        checklist_type = section.get("checklist_type", "Temporary")
        schedule_enabled = bool(section.get("schedule_enabled", False)) if checklist_type == "Permanent" else False
        repeat_enabled = bool(section.get("repeat_enabled", True)) if schedule_enabled else False
        repeat_days = safe_float(section.get("repeat_days", 1.0), 1.0) if schedule_enabled else 1.0
        time_left_days = safe_float(section.get("time_left_days", 0.0), 0.0) if schedule_enabled else 0.0
        started_at = section.get("started_at") if schedule_enabled else None
        return {
            "text": "",
            "checked": False,
            "checked_at": None,
            "checklist_type": checklist_type,
            "schedule_enabled": schedule_enabled,
            "repeat_enabled": repeat_enabled,
            "repeat_days": repeat_days,
            "time_left_days": time_left_days,
            "started_at": started_at,
        }

    def normalize_item(self, item, section_defaults=None):
        section_defaults = section_defaults or {}
        default_item = self.make_default_item(section_defaults)

        if isinstance(item, str):
            default_item["text"] = item
            return default_item

        if isinstance(item, dict):
            checklist_type = item.get("checklist_type", default_item["checklist_type"])
            checklist_type = "Permanent" if checklist_type == "Permanent" else "Temporary"

            schedule_enabled = bool(item.get("schedule_enabled", default_item["schedule_enabled"])) and checklist_type == "Permanent"
            repeat_enabled = bool(item.get("repeat_enabled", default_item["repeat_enabled"])) if schedule_enabled else False
            repeat_days = safe_float(item.get("repeat_days", item.get("repeat", default_item["repeat_days"])), 1.0)
            time_left_days = safe_float(item.get("time_left_days", item.get("time_left", default_item["time_left_days"])), 0.0)

            started_at = item.get("started_at", item.get("schedule_start_at", default_item["started_at"]))
            if isinstance(started_at, datetime):
                started_at = serialize_datetime(started_at)
            elif started_at is not None and not isinstance(started_at, str):
                started_at = None

            checked_at = item.get("checked_at")
            if isinstance(checked_at, datetime):
                checked_at = serialize_datetime(checked_at)
            elif checked_at is not None and not isinstance(checked_at, str):
                checked_at = None

            return {
                "text": item.get("text", item.get("name", "")),
                "checked": bool(item.get("checked", False)),
                "checked_at": checked_at,
                "checklist_type": checklist_type,
                "schedule_enabled": schedule_enabled,
                "repeat_enabled": repeat_enabled,
                "repeat_days": repeat_days,
                "time_left_days": time_left_days,
                "started_at": started_at,
            }

        default_item["text"] = str(item)
        return default_item

    def normalize_section(self, section):
        if not isinstance(section, dict):
            base = self.make_default_section()
            base["section"] = str(section)
            return base

        checklist_type = section.get("checklist_type", section.get("type", "Temporary"))
        checklist_type = "Permanent" if checklist_type == "Permanent" else "Temporary"

        schedule_enabled = bool(section.get("schedule_enabled", section.get("schedule", False))) and checklist_type == "Permanent"
        repeat_enabled = bool(section.get("repeat_enabled", True)) if schedule_enabled else False

        started_at = section.get("started_at", section.get("schedule_start_at"))
        if isinstance(started_at, datetime):
            started_at = serialize_datetime(started_at)
        elif started_at is not None and not isinstance(started_at, str):
            started_at = None

        repeat_days = safe_float(section.get("repeat_days", 1.0), 1.0)
        time_left_days = safe_float(section.get("time_left_days", 0.0), 0.0)

        items = section.get("items", [])
        if not isinstance(items, list):
            items = []

        defaults = {
            "checklist_type": checklist_type,
            "schedule_enabled": schedule_enabled,
            "repeat_enabled": repeat_enabled,
            "repeat_days": repeat_days,
            "time_left_days": time_left_days,
            "started_at": started_at,
        }

        return {
            "section": str(section.get("section", "")),
            "checklist_type": checklist_type,
            "schedule_enabled": schedule_enabled,
            "repeat_enabled": repeat_enabled,
            "repeat_days": repeat_days,
            "time_left_days": time_left_days,
            "started_at": started_at,
            "items": [self.normalize_item(item, defaults) for item in items],
        }

    def load_data(self):
        try:
            if os.path.exists(todo_json_path):
                with open(todo_json_path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                if isinstance(loaded, list):
                    self.sections = [self.normalize_section(section) for section in loaded]
                else:
                    self.sections = []
            else:
                self.sections = []
        except Exception as e:
            messagebox.showerror("Load Error", f"Could not load data.\n{e}")
            self.sections = []

    def save_data(self):
        try:
            with open(todo_json_path, "w", encoding="utf-8") as f:
                json.dump(self.sections, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save data.\n{e}")
            return False

    def autosave(self):
        return self.save_data()

    def on_close(self):
        self._cancel_drag()
        if self.schedule_job is not None:
            try:
                self.root.after_cancel(self.schedule_job)
            except Exception:
                pass
        self.save_data()
        if self.close_callback:
            self.close_callback()
        else:
            self.root.destroy()

    def save_and_close(self):
        self.on_close()

    def discard_changes(self):
        self.on_close()

    def get_schedule_duration_days(self, item):
        if item.get("checklist_type", "Temporary") != "Permanent":
            return None
        if not item.get("schedule_enabled", False):
            return None
        days = safe_float(item.get("time_left_days", 0.0), 0.0)
        if days <= 0 and bool(item.get("repeat_enabled", False)):
            days = safe_float(item.get("repeat_days", 0.0), 0.0)
        return days

    def get_schedule_duration_seconds(self, item):
        days = self.get_schedule_duration_days(item)
        if days is None:
            return None
        return int(max(0.0, days) * 86400)

    def sync_schedule_state(self):
        changed = False
        now = datetime.now()
        for section in self.sections:
            for item in section.get("items", []):
                if item.get("checklist_type", "Temporary") != "Permanent":
                    continue
                if not item.get("schedule_enabled", False):
                    continue

                duration_seconds = self.get_schedule_duration_seconds(item)
                if duration_seconds is None or duration_seconds <= 0:
                    continue

                started_at = parse_datetime(item.get("started_at"))
                if started_at is None:
                    item["started_at"] = serialize_datetime(now)
                    changed = True
                    continue

                elapsed = (now - started_at).total_seconds()
                if elapsed >= duration_seconds:
                    repeat_enabled = bool(item.get("repeat_enabled", False))
                    if repeat_enabled:
                        item["started_at"] = serialize_datetime(now)
                        repeat_days = safe_float(item.get("repeat_days", item.get("time_left_days", 0.0)), 0.0)
                        if repeat_days > 0 and safe_float(item.get("time_left_days", 0.0), 0.0) != repeat_days:
                            item["time_left_days"] = repeat_days
                        changed = True

                    if not self.edit_mode and item.get("checked", False):
                        item["checked"] = False
                        item["checked_at"] = None
                        changed = True

        if changed:
            self.save_data()
        return changed

    def format_remaining(self, seconds):
        if seconds is None:
            return ""
        if seconds <= 0:
            return "0d 0h left"
        days = seconds // 86400
        hours = (seconds % 86400) // 3600
        return f"{days}d {hours}h left"

    def format_remaining_full(self, seconds):
        if seconds is None:
            return ""
        if seconds <= 0:
            return "0 days, 0 hours, 0 minutes, 0 seconds left"
        days = seconds // 86400
        hours = (seconds % 86400) // 3600
        minutes = (seconds % 3600) // 60
        secs = seconds % 60
        return f"{days} days, {hours} hours, {minutes} minutes, {secs} seconds left"

    def get_item_remaining_seconds(self, item):
        if item.get("checklist_type", "Temporary") != "Permanent":
            return None
        if not item.get("schedule_enabled", False):
            return None

        duration_seconds = self.get_schedule_duration_seconds(item)
        if duration_seconds is None:
            return None

        started_at = parse_datetime(item.get("started_at"))
        if started_at is None:
            return duration_seconds

        elapsed = (datetime.now() - started_at).total_seconds()
        return max(0, int(duration_seconds - elapsed))

    def get_item_row_color(self, item):
        if self.edit_mode:
            return self.default_item_bg
        if item.get("checked", False):
            return self.checked_item_bg
        if item.get("checklist_type", "Temporary") == "Permanent" and item.get("schedule_enabled", False):
            remaining = self.get_item_remaining_seconds(item)
            if remaining is not None and remaining <= 0:
                if bool(item.get("repeat_enabled", False)):
                    return self.expired_repeat_item_bg
                return self.expired_norepeat_item_bg
            return self.counting_item_bg
        return self.default_item_bg

    def item_differs_from_section(self, item, section):
        section_type = section.get("checklist_type", "Temporary")
        section_schedule = bool(section.get("schedule_enabled", False)) if section_type == "Permanent" else False
        section_repeat_enabled = bool(section.get("repeat_enabled", True)) if section_schedule else False
        section_repeat = safe_float(section.get("repeat_days", 1.0), 1.0) if section_schedule else 1.0
        section_time_left = safe_float(section.get("time_left_days", 0.0), 0.0) if section_schedule else 0.0
        section_started = section.get("started_at") if section_schedule else None

        item_type = item.get("checklist_type", "Temporary")
        item_schedule = bool(item.get("schedule_enabled", False)) if item_type == "Permanent" else False
        item_repeat_enabled = bool(item.get("repeat_enabled", True)) if item_schedule else False
        item_repeat = safe_float(item.get("repeat_days", 1.0), 1.0) if item_schedule else 1.0
        item_time_left = safe_float(item.get("time_left_days", 0.0), 0.0) if item_schedule else 0.0
        item_started = item.get("started_at") if item_schedule else None

        return (
            item_type != section_type
            or item_schedule != section_schedule
            or item_repeat_enabled != section_repeat_enabled
            or item_repeat != section_repeat
            or item_time_left != section_time_left
            or item_started != section_started
        )

    def item_badge_text(self, item):
        parts = [item.get("checklist_type", "Temporary")]
        if item.get("checklist_type", "Temporary") == "Permanent" and item.get("schedule_enabled", False):
            parts.append("Schedule")
            if item.get("repeat_enabled", False):
                parts.append("Repeat")
        return "[" + " | ".join(parts) + "]"

    def build_items_from_lines(self, lines, existing_items=None, defaults=None):
        result = []
        existing_items = existing_items or []
        defaults = defaults or {}
        for idx, raw_line in enumerate(lines):
            line = raw_line.strip()
            if not line:
                continue
            if idx < len(existing_items):
                old = existing_items[idx]
                result.append(
                    {
                        "text": line,
                        "checked": bool(old.get("checked", False)),
                        "checked_at": old.get("checked_at"),
                        "checklist_type": old.get("checklist_type", defaults.get("checklist_type", "Temporary")),
                        "schedule_enabled": bool(old.get("schedule_enabled", defaults.get("schedule_enabled", False))),
                        "repeat_enabled": bool(old.get("repeat_enabled", defaults.get("repeat_enabled", True))),
                        "repeat_days": old.get("repeat_days", defaults.get("repeat_days", 1.0)),
                        "time_left_days": old.get("time_left_days", defaults.get("time_left_days", 0.0)),
                        "started_at": old.get("started_at", defaults.get("started_at")),
                    }
                )
            else:
                result.append(
                    {
                        "text": line,
                        "checked": False,
                        "checked_at": None,
                        "checklist_type": defaults.get("checklist_type", "Temporary"),
                        "schedule_enabled": bool(defaults.get("schedule_enabled", False)),
                        "repeat_enabled": bool(defaults.get("repeat_enabled", True)),
                        "repeat_days": defaults.get("repeat_days", 1.0),
                        "time_left_days": defaults.get("time_left_days", 0.0),
                        "started_at": defaults.get("started_at"),
                    }
                )
        return result

    def apply_section_settings_to_items(self, items, section_settings):
        updated_items = []
        checklist_type = section_settings.get("checklist_type", "Temporary")
        schedule_enabled = bool(section_settings.get("schedule_enabled", False)) if checklist_type == "Permanent" else False
        repeat_enabled = bool(section_settings.get("repeat_enabled", True)) if schedule_enabled else False
        repeat_days = safe_float(section_settings.get("repeat_days", 1.0), 1.0) if schedule_enabled else 1.0
        time_left_days = safe_float(section_settings.get("time_left_days", 0.0), 0.0) if schedule_enabled else 0.0
        started_at = section_settings.get("started_at") if schedule_enabled else None

        for item in items:
            text = item.get("text", "")
            checked = bool(item.get("checked", False))
            checked_at = item.get("checked_at")

            new_item = {
                "text": text,
                "checked": checked if checklist_type == "Permanent" else False,
                "checked_at": checked_at if checklist_type == "Permanent" else None,
                "checklist_type": checklist_type,
                "schedule_enabled": schedule_enabled,
                "repeat_enabled": repeat_enabled,
                "repeat_days": repeat_days,
                "time_left_days": time_left_days,
                "started_at": started_at,
            }

            if checklist_type == "Temporary":
                new_item["checked"] = False
                new_item["checked_at"] = None

            if checklist_type == "Permanent" and not schedule_enabled:
                new_item["started_at"] = None
                new_item["repeat_enabled"] = False

            updated_items.append(new_item)

        return updated_items

    def choose_datetime(self, entry_widget):
        initial = parse_datetime(entry_widget.get().strip()) if entry_widget.get().strip() else datetime.now()
        picker = DatePickerDialog(self.root, initial)
        self.root.wait_window(picker.top)
        if picker.result is not None:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, picker.result.strftime("%Y-%m-%d %H:%M"))

    def open_section_dialog(self, section_idx=None):
        SectionDialog(self.root, self, section_idx)

    def open_item_dialog(self, section_idx, item_idx=None):
        ItemDialog(self.root, self, section_idx, item_idx)

    def add_section_dialog(self):
        self.open_section_dialog(None)

    def add_item_to_section(self, section_idx):
        self.open_item_dialog(section_idx, None)

    def edit_item(self, section_idx, item_idx):
        self.open_item_dialog(section_idx, item_idx)

    def remove_item(self, section_idx, item_idx):
        if messagebox.askyesno("Remove", "Are you sure you want to remove this item?"):
            del self.sections[section_idx]["items"][item_idx]
            self.sync_schedule_state()
            self.save_data()
            self.refresh_display()

    def remove_section(self, section_idx):
        if messagebox.askyesno("Remove Section", f"Are you sure you want to remove '{self.sections[section_idx]['section']}'?"):
            del self.sections[section_idx]
            self.save_data()
            self.refresh_display()

    def toggle_item(self, section_idx, item_idx):
        section = self.sections[section_idx]
        item = section["items"][item_idx]

        if item.get("checklist_type", "Temporary") == "Temporary":
            if messagebox.askyesno("Finish Task", "Do you finish this task?"):
                del section["items"][item_idx]
                self.sync_schedule_state()
                self.save_data()
                self.refresh_display()
            return

        item["checked"] = not bool(item.get("checked", False))
        item["checked_at"] = serialize_datetime(datetime.now()) if item["checked"] else None
        self.sync_schedule_state()
        self.save_data()
        self.refresh_display()

    def attach_countdown_tooltip(self, widget, item):
        HoverTooltip(widget, lambda item=item: self.format_remaining_full(self.get_item_remaining_seconds(item)))

    def _bind_drag_source(self, widget, kind, section_idx, item_idx, text):
        widget.bind(
            "<ButtonPress-1>",
            lambda e, k=kind, s=section_idx, i=item_idx, t=text: self._drag_press(e, k, s, i, t),
            add="+",
        )

    def _drag_press(self, event, kind, section_idx, item_idx, text):
        if not self.edit_mode:
            return
        if getattr(event, "num", 1) != 1:
            return
        self._cancel_drag()
        self.drag_pending = {
            "kind": kind,
            "section_idx": section_idx,
            "item_idx": item_idx,
            "text": text,
        }
        self.drag_job = self.root.after(500, lambda: self._begin_drag(kind, section_idx, item_idx, text))

    def _begin_drag(self, kind, section_idx, item_idx, text):
        if not self.drag_pending:
            return
        if (
            self.drag_pending.get("kind") != kind
            or self.drag_pending.get("section_idx") != section_idx
            or self.drag_pending.get("item_idx") != item_idx
            or self.drag_pending.get("text") != text
        ):
            return
        self.drag_pending = None
        self.drag_active = True
        self.drag_kind = kind
        self.drag_from = (section_idx, item_idx)
        self.drag_source_text = text
        self.drag_hover_target = None
        self.root.configure(cursor="fleur")
        self._create_drag_preview(kind, text)
        x, y = self.root.winfo_pointerxy()
        self._move_drag_preview(x, y)
        self._update_drag_hover_target(x, y)

    def _create_drag_preview(self, kind, text):
        self._destroy_drag_preview()
        self.drag_preview = tw = tk.Toplevel(self.root)
        tw.overrideredirect(True)
        tw.wm_attributes("-topmost", True)
        try:
            tw.wm_attributes("-alpha", 0.92)
        except Exception:
            pass
        outer = tk.Frame(tw, bg="#0f0f0f", highlightthickness=2, highlightbackground="#3c8dbc")
        outer.pack(fill="both", expand=True)
        label_text = f"Move {kind}: {shorten_text(text, 52)}"
        self.drag_preview_label = tk.Label(
            outer,
            text=label_text,
            bg="#0f0f0f",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=6,
            justify="left",
        )
        self.drag_preview_label.pack()

    def _move_drag_preview(self, x_root, y_root):
        if self.drag_preview is None:
            return
        try:
            self.drag_preview.geometry(f"+{x_root + 16}+{y_root + 16}")
        except Exception:
            pass

    def _destroy_drag_preview(self):
        if self.drag_preview is not None:
            try:
                self.drag_preview.destroy()
            except Exception:
                pass
        self.drag_preview = None
        self.drag_preview_label = None

    def _cancel_drag(self):
        if self.drag_job is not None:
            try:
                self.root.after_cancel(self.drag_job)
            except Exception:
                pass
            self.drag_job = None
        self.drag_pending = None
        self.drag_active = False
        self.drag_kind = None
        self.drag_from = None
        self.drag_source_text = ""
        self.drag_hover_target = None
        self.root.configure(cursor="")
        self._destroy_drag_preview()

    def _current_drag_target(self, y_root):
        if not self.drag_active:
            return None
        if self.drag_kind == "section":
            return ("section", self._section_drop_index(y_root))
        if self.drag_kind == "item":
            source_section = self.drag_from[0] if self.drag_from is not None else 0
            item_target = self._item_drop_target(y_root, source_section)
            return ("item", item_target[0], item_target[1])
        return None

    def _update_drag_hover_target(self, x_root, y_root):
        target = self._current_drag_target(y_root)
        if target != self.drag_hover_target:
            self.drag_hover_target = target
            self.refresh_display(save=False)

    def _global_drag_motion(self, event):
        if not self.drag_active:
            return
        x_root, y_root = self.root.winfo_pointerxy()
        self._move_drag_preview(x_root, y_root)
        self._update_drag_hover_target(x_root, y_root)

    def _global_drag_release(self, event):
        if self.drag_active:
            x_root, y_root = self.root.winfo_pointerxy()
            self._finish_drag(x_root, y_root)
        else:
            self._cancel_drag()

    def _section_drop_index(self, y_root):
        headers = []
        for idx, widget in enumerate(getattr(self, "section_header_widgets", [])):
            if widget is None or not widget.winfo_exists():
                continue
            headers.append((widget.winfo_rooty() + widget.winfo_height() / 2.0, idx))
        headers.sort(key=lambda x: x[0])
        if not headers:
            return 0
        for center_y, idx in headers:
            if y_root < center_y:
                return idx
        return len(headers)

    def _item_drop_target(self, y_root, source_section=None):
        if source_section is None:
            source_section = 0

        items = []
        for entry in getattr(self, "item_widgets", []):
            widget = entry.get("widget")
            if widget is None or not widget.winfo_exists():
                continue
            if entry.get("section_idx") != source_section:
                continue
            items.append((widget.winfo_rooty() + widget.winfo_height() / 2.0, entry.get("item_idx", 0)))

        items.sort(key=lambda x: x[0])

        if not items:
            return source_section, 0

        for center_y, item_idx in items:
            if y_root < center_y:
                return source_section, item_idx
        return source_section, len(items)

    def _finish_drag(self, x_root, y_root):
        kind = self.drag_kind
        source = self.drag_from
        self._destroy_drag_preview()
        self.root.configure(cursor="")
        self.drag_active = False

        changed = False
        if kind == "section" and source is not None:
            from_idx = source[0]
            to_idx = self._section_drop_index(y_root)
            changed = self._move_section(from_idx, to_idx)
        elif kind == "item" and source is not None:
            from_section, from_item = source
            _, to_item = self._item_drop_target(y_root, from_section)
            changed = self._move_item(from_section, from_item, from_section, to_item)

        self.drag_kind = None
        self.drag_from = None
        self.drag_source_text = ""
        self.drag_pending = None
        self.drag_hover_target = None
        if self.drag_job is not None:
            try:
                self.root.after_cancel(self.drag_job)
            except Exception:
                pass
            self.drag_job = None

        if changed:
            self.sync_schedule_state()
            self.save_data()
            self.refresh_display()
        else:
            self.refresh_display()

    def _move_section(self, from_idx, to_idx):
        if from_idx < 0 or from_idx >= len(self.sections):
            return False
        to_idx = max(0, min(to_idx, len(self.sections)))
        if to_idx == from_idx:
            return False
        section = self.sections.pop(from_idx)
        if to_idx > from_idx:
            to_idx -= 1
        self.sections.insert(to_idx, section)
        return True

    def _move_item(self, from_section, from_item, to_section, to_item):
        if from_section < 0 or from_section >= len(self.sections):
            return False
        if to_section != from_section:
            return False
        if from_item < 0 or from_item >= len(self.sections[from_section].get("items", [])):
            return False

        item = self.sections[from_section]["items"].pop(from_item)
        target_items = self.sections[to_section].setdefault("items", [])

        if to_item > from_item:
            to_item -= 1

        to_item = max(0, min(to_item, len(target_items)))
        target_items.insert(to_item, item)
        return True

    def _create_drop_preview_block(self, parent, text, height=28):
        frame = tk.Frame(parent, bg="#3c8dbc", height=height, highlightthickness=2, highlightbackground="#ffffff")
        frame.pack(fill="x", pady=3)
        frame.pack_propagate(False)
        label = tk.Label(
            frame,
            text=text,
            bg="#3c8dbc",
            fg="white",
            font=("Arial", 9, "bold"),
        )
        label.pack(expand=True)
        return frame

    def refresh_display(self, save=True):
        self.sync_schedule_state()

        for widget in self.scroll_container.winfo_children():
            widget.destroy()

        self.section_header_widgets = []
        self.item_widgets = []

        section_preview_target = None
        item_preview_target = None

        if self.drag_active and self.drag_hover_target:
            if self.drag_hover_target[0] == "section":
                section_preview_target = self.drag_hover_target[1]
            elif self.drag_hover_target[0] == "item":
                item_preview_target = (self.drag_hover_target[1], self.drag_hover_target[2])

        if not self.sections:
            if self.drag_active and self.drag_kind == "section" and section_preview_target == 0:
                self._create_drop_preview_block(self.scroll_container, "Drop section here", height=32)
            tk.Label(
                self.scroll_container,
                text="No sections yet. Add one to get started!",
                font=("Arial", 12, "bold"),
                bg=self.list_bg,
                fg="#888",
            ).pack(padx=20, pady=40)
            if save:
                self.save_data()
            return

        for section_idx, section in enumerate(self.sections):
            if self.drag_active and self.drag_kind == "section" and section_preview_target == section_idx:
                self._create_drop_preview_block(self.scroll_container, "Drop section here", height=32)

            section_header_frame = tk.Frame(self.scroll_container, bg=self.section_bg)
            section_header_frame.pack(fill="x", padx=5, pady=(10, 0))
            self.section_header_widgets.append(section_header_frame)

            title_frame = tk.Frame(section_header_frame, bg=self.section_bg)
            title_frame.pack(fill="x", padx=5, pady=5)

            title_text = section.get("section", "")

            title_label = tk.Label(
                title_frame,
                text=title_text,
                font=("Arial", 12, "bold"),
                bg=self.section_bg,
                fg=self.btn_color,
                anchor="w",
                justify="left",
                wraplength=330,
            )
            title_label.pack(side="left", fill="x", expand=True, anchor="w")

            if self.edit_mode:
                meta_parts = [section.get("checklist_type", "Temporary")]
                if section.get("checklist_type", "Temporary") == "Permanent" and section.get("schedule_enabled", False):
                    meta_parts.append("Schedule")
                    if section.get("repeat_enabled", False):
                        meta_parts.append("Repeat")

                tk.Label(
                    title_frame,
                    text="[" + " | ".join(meta_parts) + "]",
                    font=("Arial", 9, "bold"),
                    bg=self.section_bg,
                    fg="#b8b8b8",
                ).pack(side="right", padx=6)

                self._bind_drag_source(title_frame, "section", section_idx, None, title_text)
                self._bind_drag_source(title_label, "section", section_idx, None, title_text)

            items_frame = tk.Frame(section_header_frame, bg=self.section_bg)
            items_frame.pack(fill="x", padx=10)

            section_items = section.get("items") or []

            if not section_items:
                if self.drag_active and self.drag_kind == "item" and item_preview_target == (section_idx, 0):
                    self._create_drop_preview_block(items_frame, "Drop item here", height=22)
                else:
                    tk.Label(
                        items_frame,
                        text="(No items)",
                        font=("Arial", 10),
                        bg=self.section_bg,
                        fg="#888",
                        anchor="w",
                        justify="left",
                    ).pack(anchor="w", pady=3)
            else:
                for item_idx, item in enumerate(section_items):
                    if self.drag_active and self.drag_kind == "item" and item_preview_target == (section_idx, item_idx):
                        self._create_drop_preview_block(items_frame, "Drop item here", height=22)

                    item_bg = self.get_item_row_color(item)
                    item_frame = tk.Frame(items_frame, bg=item_bg)
                    item_frame.pack(fill="x", pady=2)
                    self.item_widgets.append({"widget": item_frame, "section_idx": section_idx, "item_idx": item_idx})

                    if self.edit_mode:
                        text_label = tk.Label(
                            item_frame,
                            text="• " + item.get("text", ""),
                            font=("Arial", 10),
                            bg=item_bg,
                            fg=self.fg_color,
                            wraplength=250,
                            justify="left",
                            anchor="w",
                        )
                        text_label.pack(side="left", padx=10, pady=5, anchor="w", fill="x", expand=True)

                        right_box = tk.Frame(item_frame, bg=item_bg)
                        right_box.pack(side="right", padx=(2, 10), pady=2)

                        countdown = self.format_remaining(self.get_item_remaining_seconds(item))
                        if countdown:
                            countdown_label = tk.Label(
                                right_box,
                                text=countdown,
                                font=("Arial", 9, "bold"),
                                bg=item_bg,
                                fg="#dddddd",
                            )
                            countdown_label.pack(side="top", anchor="e")
                            self.attach_countdown_tooltip(countdown_label, item)

                        if self.item_differs_from_section(item, section):
                            tk.Label(
                                right_box,
                                text=self.item_badge_text(item),
                                font=("Arial", 8, "bold"),
                                bg=item_bg,
                                fg="#cccccc",
                            ).pack(side="top", anchor="e", pady=(2, 0))

                        buttons_box = tk.Frame(right_box, bg=item_bg)
                        buttons_box.pack(side="top", anchor="e", pady=(2, 0))

                        tk.Button(
                            buttons_box,
                            text="✎",
                            font=("Arial", 9),
                            bg=item_bg,
                            fg="white",
                            activebackground=item_bg,
                            width=2,
                            command=lambda sec=section_idx, itm=item_idx: self.edit_item(sec, itm),
                        ).pack(side="left", padx=2)

                        tk.Button(
                            buttons_box,
                            text="✕",
                            font=("Arial", 9),
                            bg=item_bg,
                            fg="white",
                            activebackground=item_bg,
                            width=2,
                            command=lambda sec=section_idx, itm=item_idx: self.remove_item(sec, itm),
                        ).pack(side="left", padx=2)

                        self._bind_drag_source(item_frame, "item", section_idx, item_idx, item.get("text", ""))
                        self._bind_drag_source(text_label, "item", section_idx, item_idx, item.get("text", ""))
                    else:
                        tick_symbol = "☑" if item.get("checked", False) else "☐"

                        tk.Button(
                            item_frame,
                            text=tick_symbol,
                            font=("Arial", 18, "bold"),
                            bg=item_bg,
                            fg=self.fg_color,
                            activebackground=item_bg,
                            activeforeground=self.fg_color,
                            bd=0,
                            relief="flat",
                            width=2,
                            command=lambda sec=section_idx, itm=item_idx: self.toggle_item(sec, itm),
                        ).pack(side="left", padx=(4, 4), pady=2)

                        tk.Label(
                            item_frame,
                            text=item.get("text", ""),
                            font=("Arial", 10),
                            bg=item_bg,
                            fg=self.fg_color,
                            wraplength=270,
                            justify="left",
                            anchor="w",
                        ).pack(side="left", padx=6, pady=6, anchor="w", fill="x", expand=True)

                        right_box = tk.Frame(item_frame, bg=item_bg)
                        right_box.pack(side="right", padx=(2, 10), pady=2)

                        countdown = self.format_remaining(self.get_item_remaining_seconds(item))
                        if countdown:
                            countdown_label = tk.Label(
                                right_box,
                                text=countdown,
                                font=("Arial", 9, "bold"),
                                bg=item_bg,
                                fg="#dddddd",
                                anchor="e",
                            )
                            countdown_label.pack(side="top", anchor="e")
                            self.attach_countdown_tooltip(countdown_label, item)

                if self.drag_active and self.drag_kind == "item" and item_preview_target == (section_idx, len(section_items)):
                    self._create_drop_preview_block(items_frame, "Drop item here", height=22)

            if self.edit_mode:
                footer_frame = tk.Frame(section_header_frame, bg=self.section_bg)
                footer_frame.pack(fill="x", padx=10, pady=(0, 6))

                buttons_right = tk.Frame(footer_frame, bg=self.section_bg)
                buttons_right.pack(side="right")

                remove_btn = tk.Button(
                    buttons_right,
                    text="🗑",
                    font=("Arial", 10, "bold"),
                    bg="#5B0000",
                    fg="white",
                    activebackground="#3d0000",
                    width=3,
                    command=lambda idx=section_idx: self.remove_section(idx),
                )
                remove_btn.pack(side="right", padx=(2, 0))
                HoverTooltip(remove_btn, lambda: "Remove Section")

                edit_btn = tk.Button(
                    buttons_right,
                    text="✎",
                    font=("Arial", 10, "bold"),
                    bg="#444",
                    fg="white",
                    activebackground="#333",
                    width=3,
                    command=lambda idx=section_idx: self.open_section_dialog(idx),
                )
                edit_btn.pack(side="right", padx=(2, 0))
                HoverTooltip(edit_btn, lambda: "Edit Section")

                tk.Button(
                    footer_frame,
                    text="＋ Add Item",
                    font=("Arial", 9),
                    bg=self.edit_btn_color,
                    fg="white",
                    activebackground="#004620",
                    command=lambda idx=section_idx: self.add_item_to_section(idx),
                ).pack(side="left")

                self._bind_drag_source(section_header_frame, "section", section_idx, None, title_text)

        if self.drag_active and self.drag_kind == "section" and section_preview_target == len(self.sections):
            self._create_drop_preview_block(self.scroll_container, "Drop section here", height=32)

        if self.edit_mode:
            self.add_section_button.pack(side="left")
        else:
            self.add_section_button.pack_forget()

        self.root.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        if save:
            self.save_data()


def open_todo_list_popup():
    todo_root = tk.Toplevel()
    TodoListApp(todo_root)
    todo_root.mainloop()


if __name__ == "__main__":
    root = tk.Tk()
    TodoListApp(root)
    root.mainloop()
