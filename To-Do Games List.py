import tkinter as tk
from tkinter import messagebox, filedialog
import os
import json
from PIL import Image, ImageTk, ImageDraw
import win32com.client
import win32gui
import win32con
import win32ui
import sys
import ctypes
import importlib.util
from edit_list import open_txt_popup
from help import open_help_popup
from settings import open_settings_popup
from all_app import open_all_app_popup
import re
import urllib.request
import urllib.error

if sys.platform == "win32":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except:
            pass

if getattr(sys, "frozen", False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

settings_path = os.path.join(base_dir, "settings.json")
default_filename = None

try:
    with open(settings_path, "r", encoding="utf-8") as f:
        settings = json.load(f)
        gb = settings.get("Gamelist")
        if gb:
            if os.path.isabs(gb):
                default_filename = gb
            else:
                default_filename = os.path.join(base_dir, gb)
except:
    settings = {}

def resolve_shortcut(path):
    try:
        return win32com.client.Dispatch("WScript.Shell").CreateShortcut(path).TargetPath
    except:
        return path

def fetch_steam_name(appid):
    try:
        url = f"https://store.steampowered.com/api/appdetails?appids={appid}&cc=us&l=en"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=6) as resp:
            data = json.load(resp)
        if str(appid) in data and data[str(appid)] and data[str(appid)].get("success"):
            info = data[str(appid)].get("data", {})
            name = info.get("name")
            if name:
                return name
    except:
        pass
    return None

def extract_icon(path):
    try:
        if not os.path.exists(path):
            return None
        large, _ = win32gui.ExtractIconEx(path, 0)
        if large:
            hicon = large[0]
            hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
            hbmp = win32ui.CreateBitmap()
            hbmp.CreateCompatibleBitmap(hdc, 32, 32)
            hdc = hdc.CreateCompatibleDC()
            hdc.SelectObject(hbmp)
            win32gui.DrawIconEx(hdc.GetHandleOutput(), 0, 0, hicon, 32, 32, 0, None, win32con.DI_NORMAL)
            bmpinfo = hbmp.GetInfo()
            bmpstr = hbmp.GetBitmapBits(True)
            win32gui.DestroyIcon(hicon)
            return Image.frombuffer("RGB", (bmpinfo["bmWidth"], bmpinfo["bmHeight"]), bmpstr, "raw", "BGRX", 0, 1)
    except:
        return None

def normalize_filters(value):
    if value is None:
        return []
    if isinstance(value, list):
        result = []
        for item in value:
            text = str(item).strip().lower()
            if text:
                result.append(text)
        return result
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return []
        if text.startswith("[") and text.endswith("]"):
            try:
                parsed = json.loads(text)
                if isinstance(parsed, list):
                    return normalize_filters(parsed)
            except:
                pass
        parts = re.split(r"[,\n;|]+", text)
        return [part.strip().lower() for part in parts if part.strip()]
    text = str(value).strip().lower()
    return [text] if text else []

def resolve_game_entry(game_path, is_admin, name_override=None, filters=None, filter_value=None):
    if filters is None:
        filters = filter_value if filter_value is not None else []
    filters = normalize_filters(filters)

    if not game_path:
        return {
            "name": name_override or "",
            "run": game_path,
            "real": game_path,
            "admin": is_admin,
            "filters": filters,
        }

    lower = (game_path or "").lower()
    if lower.startswith("steam://"):
        m = re.search(r"(\d+)", game_path)
        appid = m.group(1) if m else None
        name = None
        if appid:
            name = fetch_steam_name(appid)
        display_name = name if name else (f"Steam {appid}" if appid else os.path.basename(game_path))
        return {
            "name": name_override or display_name,
            "run": game_path,
            "real": game_path,
            "admin": is_admin,
            "filters": filters,
        }

    ext = os.path.splitext(game_path)[1].lower()
    resolved_path = resolve_shortcut(game_path) if ext in [".lnk", ".url"] else game_path
    try:
        display_name = os.path.splitext(os.path.basename(resolved_path or game_path))[0]
    except:
        display_name = name_override or ""

    return {
        "name": name_override or display_name,
        "run": game_path,
        "real": resolved_path,
        "admin": is_admin,
        "filters": filters,
    }

def _robust_json_load(path):
    with open(path, "r", encoding="utf-8") as f:
        raw = f.read()
    raw = raw.lstrip("\ufeff").strip()
    try:
        return json.loads(raw)
    except Exception:
        start = raw.find("[")
        end = raw.rfind("]")
        if start != -1 and end != -1 and end > start:
            chunk = raw[start:end + 1]
            chunk = re.sub(r",\s*([\]\}])", r"\1", chunk)
            chunk = re.sub(r'(?<!\\)\\(?!\\)', r'\\\\', chunk)
            return json.loads(chunk)
        objs = re.findall(r"\{(?:[^{}]|\n|\r)*?\}", raw)
        parsed = []
        for o in objs:
            o2 = re.sub(r",\s*([\}\]])", r"\1", o)
            o2 = re.sub(r'(?<!\\)\\(?!\\)', r'\\\\', o2)
            try:
                parsed.append(json.loads(o2))
            except:
                continue
        if parsed:
            return parsed
        raise ValueError("Invalid JSON")

class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tipwindow or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw,
            text=self.text,
            bg="#222222",
            fg="white",
            relief="solid",
            borderwidth=1,
            font=("Arial", 9, "normal")
        )
        label.pack(ipadx=6, ipady=3)

    def hide_tip(self, event=None):
        if self.tipwindow:
            self.tipwindow.destroy()
            self.tipwindow = None

class GamesListApp:
    GAMES_WIDTH = 580
    SHARED_HEIGHT = 820
    TODO_WIDTH = 600

    def __init__(self, root):
        self.root = root
        try:
            self.root.iconbitmap(os.path.join(base_dir, "icon.ico"))
        except:
            pass
        self.root.title("To-Do Games List")
        self.root.resizable(False, False)
        self.auto_play = tk.BooleanVar(value=settings.get("AutoPlay", True))
        self.current_index = 0
        self.games = []
        self.visible_indices = []
        self.icon_images = []
        self.item_frames = []
        self.bullet_image = self.make_bullet_image()
        self.bg_color, self.fg_color = "#1e1e1e", "#ffffff"
        self.select_bg, self.btn_color = "#345d9d", "#3c8dbc"
        self.num_list_bg, self.list_bg = "#141414", "#363636"
        self.active_filter = "All"
        self.filter_buttons = {}
        self.canvas_height = 420
        self.todo_root = None
        self.todo_app = None
        self.closing = False
        self.root.configure(bg=self.bg_color)
        self.root.geometry(f"{self.GAMES_WIDTH}x{self.SHARED_HEIGHT}")
        self.create_widgets()
        self.root.protocol("WM_DELETE_WINDOW", self.close_both_windows)
        if default_filename:
            if not os.path.exists(default_filename):
                with open(default_filename, "w", encoding="utf-8") as f:
                    json.dump([], f, indent=4)
            self.load_from_file(default_filename)
        else:
            messagebox.showinfo("No List Set", "No game list configured in settings.json.\n\nPlease create or load a list using the buttons below.")
        self.root.after(300, self.open_todo_list_auto)

    def make_bullet_image(self):
        img = Image.new("RGBA", (32, 32), (0, 0, 0, 0))
        ImageDraw.Draw(img).ellipse((10, 10, 22, 22), fill="white")
        return ImageTk.PhotoImage(img)

    def _window_size(self, window, fallback_w, fallback_h):
        window.update_idletasks()
        return fallback_w, fallback_h

    def arrange_windows(self):
        if not self.root or not self.root.winfo_exists():
            return
        self.root.update_idletasks()
        if self.todo_root and self.todo_root.winfo_exists():
            self.todo_root.update_idletasks()

        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        gap = 10

        left_w, left_h = self.GAMES_WIDTH, self.SHARED_HEIGHT
        right_w, right_h = self.TODO_WIDTH, self.SHARED_HEIGHT if self.todo_root and self.todo_root.winfo_exists() else (0, 0)

        pair_w = left_w + (gap if right_w else 0) + right_w
        x_start = max(0, (screen_w - pair_w) // 2)
        y = max(0, (screen_h - self.SHARED_HEIGHT) // 2 - 20)

        self.root.geometry(f"{left_w}x{left_h}+{x_start}+{y}")

        if self.todo_root and self.todo_root.winfo_exists():
            right_x = x_start + left_w + gap
            self.todo_root.geometry(f"{right_w}x{right_h}+{right_x}+{y}")

    def create_widgets(self):
        top = tk.Frame(self.root, bg=self.bg_color)
        top.pack(fill="x", padx=10, pady=(10, 0))

        self.list_label = tk.Label(top, font=("Arial", 12, "bold"), bg=self.bg_color, fg=self.fg_color)
        self.list_label.pack(side="left")

        right_top = tk.Frame(top, bg=self.bg_color)
        right_top.pack(side="right")

        create_btn = tk.Button(
            right_top,
            text="📄",
            font=("Arial", 12, "bold"),
            width=3,
            bg="#666",
            fg="white",
            command=self.create_list_dialog
        )
        create_btn.pack(side="left", padx=(0, 6))
        Tooltip(create_btn, "Create New List")

        todo_btn = tk.Button(
            right_top,
            text="☑",
            font=("Arial", 12, "bold"),
            width=3,
            bg="#007acc",
            fg="white",
            command=self.toggle_todo_window
        )
        todo_btn.pack(side="left")
        Tooltip(todo_btn, "To-Do List")

        self.update_list_label()

        self.filter_frame = tk.Frame(self.root, bg=self.bg_color)
        self.filter_frame.pack(fill="x", padx=10, pady=(6, 0))

        self.main_frame = tk.Frame(self.root, bg=self.bg_color)
        self.main_frame.pack(padx=10, pady=10)

        self.canvas = tk.Canvas(self.main_frame, width=520, height=self.canvas_height, bg=self.list_bg, highlightthickness=0)
        self.canvas.pack(side="left", fill="none", expand=False)

        self.scrollbar = tk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scroll_container = tk.Frame(self.canvas, bg=self.list_bg)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scroll_container, anchor="nw")

        self.scroll_container.bind("<Configure>", self._update_scrollregion)
        self.canvas.bind("<Configure>", self._fit_canvas_width)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.number_frame = tk.Frame(self.scroll_container, bg=self.num_list_bg)
        self.number_frame.pack(side="left", fill="y")
        self.scroll_frame = tk.Frame(self.scroll_container, bg=self.list_bg)
        self.scroll_frame.pack(side="left", fill="x")

        self.button_frame = tk.Frame(self.root, bg=self.bg_color)
        self.button_frame.pack(pady=10)
        for txt, cmd in [("←", self.go_left), ("▶", self.launch_game), ("→", self.go_right)]:
            tk.Button(
                self.button_frame,
                text=txt,
                font=("Arial", 28, "bold"),
                width=4,
                height=1,
                bg=self.btn_color,
                fg="white",
                activebackground="#2e6fa3",
                command=cmd
            ).pack(side="left", padx=10)

        self.auto_button = tk.Button(
            self.root,
            text="▶ Auto Play: ON",
            font=("Arial", 12, "bold"),
            bg="#3cb371",
            fg="white",
            activebackground="#2e6fa3",
            relief="ridge",
            bd=3,
            width=18,
            command=self.toggle_autoplay
        )
        self.auto_button.pack(pady=5)
        self.update_autoplay_button()

        control = tk.Frame(self.root, bg=self.bg_color)
        control.pack(pady=5)

        top_row = tk.Frame(control, bg=self.bg_color)
        top_row.pack(pady=5)
        tk.Button(top_row, text="📂 Load List (.json)", width=14, bg="#666", fg="white", command=self.load_dialog).pack(side="left", padx=5)
        tk.Button(top_row, text="➕ Add Game", width=14, bg="#005b29", fg="white", command=self.add_game_dialog).pack(side="left", padx=5)
        tk.Button(top_row, text="❓ Help", width=14, bg="#444", fg="white", command=open_help_popup).pack(side="left", padx=5)

        bottom_row = tk.Frame(control, bg=self.bg_color)
        bottom_row.pack(pady=5)
        tk.Button(bottom_row, text="✏ Edit List", width=14, bg="#666", fg="white", command=lambda: open_txt_popup(self.root, self.bg_color, self.list_bg, self.fg_color, self.load_from_file)).pack(side="left", padx=5)
        tk.Button(bottom_row, text="🗑 Remove Game", width=14, bg="#5B0000", fg="white", command=self.remove_selected_game).pack(side="left", padx=5)
        tk.Button(bottom_row, text="⚙ Settings", width=14, bg="#444", fg="white", command=open_settings_popup).pack(side="left", padx=5)

    def _fit_canvas_width(self, event=None):
        self.canvas.itemconfigure(self.canvas_window, width=self.canvas.winfo_width())

    def _update_scrollregion(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def clear_frame(self, frame):
        for widget in frame.winfo_children():
            widget.destroy()

    def add_game_widget(self, game, index):
        frame = tk.Frame(self.scroll_frame, bg=self.list_bg)
        frame.pack(fill="x", pady=5)
        self.item_frames.append(frame)
        icon = self.load_icon_image(game.get("real") or game.get("run"))
        self.icon_images.append(icon)
        tk.Label(frame, image=icon, bg=self.list_bg).pack(side="left", padx=5)
        label_text = game.get("name") or os.path.basename(game.get("run", "")) or "Unknown"
        label = tk.Label(frame, text=label_text, font=("Consolas", 16), bg=self.list_bg, fg=self.fg_color, anchor="w")
        label.pack(side="left", fill="x", expand=True)
        label.bind("<Button-1>", lambda e, idx=index: self.select_game(idx))
        frame.bind("<Button-1>", lambda e, idx=index: self.select_game(idx))

    def refresh_game_list(self):
        self.clear_frame(self.scroll_frame)
        self.clear_frame(self.number_frame)
        self.item_frames.clear()
        self.icon_images.clear()

        if self.active_filter and self.active_filter != "All":
            self.visible_indices = [i for i, g in enumerate(self.games) if self.active_filter.lower() in normalize_filters(g.get("filters", []))]
        else:
            self.visible_indices = list(range(len(self.games)))

        filtered_games = [self.games[i] for i in self.visible_indices]

        if not filtered_games:
            tk.Label(self.scroll_frame, text="Pls add new game", font=("Arial", 16, "bold"), bg=self.list_bg, fg="#888").pack(padx=30, pady=40)
            self.current_index = 0
        else:
            for idx, game in enumerate(filtered_games):
                self.add_game_widget(game, idx)
                tk.Label(self.number_frame, text=f"{idx + 1}.", font=("Consolas", 16), bg=self.num_list_bg, fg=self.fg_color, anchor="e", width=4).pack(anchor="n", pady=5)
            self.current_index = max(0, min(self.current_index, len(filtered_games) - 1))
            self.highlight_current()

        self.root.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.canvas.yview_moveto(0)

    def add_game_dialog(self):
        def on_game_selected(name, exe_path):
            entry = resolve_game_entry(exe_path, False, name_override=(name or None), filters=[])
            self.games.append(entry)
            self.current_index = len(self.visible_indices) if self.active_filter and self.active_filter != "All" else len(self.games) - 1
            self.save_games()
            self.refresh_game_list()
            self.update_filter_tabs()
        open_all_app_popup(self.root, on_game_selected)

    def remove_selected_game(self):
        if not self.games:
            messagebox.showwarning("No Game", "No game selected to remove.")
            return
        if not self.visible_indices:
            messagebox.showwarning("No Game", "No game selected to remove.")
            return

        actual_index = self.visible_indices[self.current_index]
        game = self.games[actual_index]
        if messagebox.askyesno("Remove Game", f"Are you sure you want to remove:\n\n{game.get('name','') or game.get('run','')}?"):
            del self.games[actual_index]
            self.current_index = 0
            self.save_games()
            self.refresh_game_list()
            self.update_filter_tabs()

    def create_list_dialog(self):
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")], initialdir=base_dir)
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    json.dump([], f, indent=4)
                global default_filename
                default_filename = path
                self.games.clear()
                self.current_index = 0
                self.refresh_game_list()
                self.save_games()
                self.update_list_label()
                self.update_settings_json(path)
                messagebox.showinfo("Created", f"New list created:\n{path}")
                self.update_filter_tabs()
            except Exception as e:
                messagebox.showerror("Error", f"Could not create file.\n{e}")

    def save_games(self):
        try:
            with open(default_filename, "w", encoding="utf-8") as f:
                json.dump(self.games, f, indent=4)
        except Exception as e:
            messagebox.showerror("Auto Save Error", f"Could not save automatically.\n{e}")

    def load_icon_image(self, path):
        icon_img = extract_icon(path) if path and not str(path).lower().startswith("steam://") else None
        return ImageTk.PhotoImage(icon_img.resize((32, 32), Image.LANCZOS)) if icon_img else self.bullet_image

    def _on_mousewheel(self, e):
        if not self.games:
            return
        bbox = self.canvas.bbox("all")
        if not bbox:
            return
        canvas_height = self.canvas.winfo_height()
        content_height = bbox[3] - bbox[1]
        if content_height <= canvas_height:
            return
        y1, y2 = self.canvas.yview()
        if e.delta > 0 and y1 <= 0:
            return
        if e.delta < 0 and y2 >= 1:
            return
        self.canvas.yview_scroll(-1 if e.delta > 0 else 1, "units")

    def select_game(self, index):
        self.current_index = index
        self.highlight_current()
        if self.auto_play.get():
            self.launch_game()

    def highlight_current(self):
        for i, frame in enumerate(self.item_frames):
            bg = self.select_bg if i == self.current_index else self.list_bg
            for widget in frame.winfo_children():
                try:
                    widget.configure(bg=bg)
                except:
                    pass
            try:
                frame.configure(bg=bg)
            except:
                pass

    def go_left(self):
        if self.current_index > 0:
            self.current_index -= 1
            self.highlight_current()
            if self.auto_play.get():
                self.launch_game()

    def go_right(self):
        if self.current_index < len(self.item_frames) - 1:
            self.current_index += 1
            self.highlight_current()
            if self.auto_play.get():
                self.launch_game()

    def launch_game(self):
        if not self.games or not self.visible_indices:
            return
        actual_game = None
        actual_index = self.visible_indices[self.current_index]
        if 0 <= actual_index < len(self.games):
            actual_game = self.games[actual_index]
        if not actual_game:
            return
        if actual_game.get("admin") and not messagebox.askyesno("Admin Launch", f"Launch {actual_game.get('name')} with admin?"):
            return
        try:
            os.startfile(actual_game.get("run"))
        except Exception as e:
            messagebox.showerror("Error", f"Could not run the game.\n{e}")

    def toggle_autoplay(self):
        val = not self.auto_play.get()
        self.auto_play.set(val)
        self.auto_button.config(text=f"▶ Auto Play: {'ON' if val else 'OFF'}", bg="#3cb371" if val else "#777", fg="white")
        try:
            data = {}
            if os.path.exists(settings_path):
                with open(settings_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            data["AutoPlay"] = val
            with open(settings_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            messagebox.showerror("Settings Error", f"Could not save AutoPlay setting.\n{e}")

    def update_autoplay_button(self):
        val = self.auto_play.get()
        self.auto_button.config(text=f"▶ Auto Play: {'ON' if val else 'OFF'}", bg="#3cb371" if val else "#777", fg="white")

    def open_todo_list_auto(self):
        self.open_todo_list()

    def open_todo_list(self):
        try:
            if self.todo_root and self.todo_root.winfo_exists():
                self.todo_root.deiconify()
                self.todo_root.lift()
                self.todo_root.focus_force()
                self.arrange_windows()
                return
            spec = importlib.util.spec_from_file_location("TodoList", os.path.join(base_dir, "To-Do List.py"))
            todo_module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(todo_module)
            self.todo_root = tk.Toplevel(self.root)
            self.todo_app = todo_module.TodoListApp(self.todo_root, center=False, close_callback=self.close_todo_only)
            self.todo_root.update_idletasks()
            self.arrange_windows()
            self.todo_root.lift()
        except Exception as e:
            messagebox.showerror("Error", f"Could not open To-Do List.\n{e}")

    def toggle_todo_window(self):
        if self.todo_root and self.todo_root.winfo_exists():
            if self.todo_root.state() == "normal":
                self.todo_root.withdraw()
            else:
                self.todo_root.deiconify()
                self.arrange_windows()
                self.todo_root.lift()
        else:
            self.open_todo_list()

    def close_todo_only(self):
        try:
            if self.todo_app is not None:
                self.todo_app.save_data()
        except:
            pass
        try:
            if self.todo_root and self.todo_root.winfo_exists():
                self.todo_root.destroy()
        except:
            pass
        self.todo_root = None
        self.todo_app = None

    def close_both_windows(self):
        if self.closing:
            return
        self.closing = True
        try:
            self.save_games()
        except:
            pass
        self.close_todo_only()
        try:
            if self.root and self.root.winfo_exists():
                self.root.destroy()
        except:
            pass

    def load_dialog(self):
        path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json"), ("Text Files", "*.txt")], initialdir=base_dir)
        if path:
            self.load_from_file(path)
            global default_filename
            default_filename = path
            self.update_settings_json(path)
            self.update_list_label()
            self.update_filter_tabs()

    def update_settings_json(self, path):
        try:
            settings_local = {}
            if os.path.exists(settings_path):
                with open(settings_path, "r", encoding="utf-8") as f:
                    settings_local = json.load(f)
            settings_local["Gamelist"] = os.path.basename(path)
            with open(settings_path, "w", encoding="utf-8") as f:
                json.dump(settings_local, f, indent=4)
        except Exception as e:
            messagebox.showerror("Error", f"Could not update settings.json\n{e}")

    def update_list_label(self):
        if default_filename:
            self.list_label.config(text=os.path.basename(default_filename).replace(".json", "").replace(".txt", ""))
        else:
            self.list_label.config(text="No List Selected")

    def update_filter_tabs(self):
        self.clear_frame(self.filter_frame)
        self.filter_buttons = {}
        filters = ["All"]
        seen = set()
        for g in self.games:
            for f in normalize_filters(g.get("filters", [])):
                if f and f not in seen:
                    seen.add(f)
                    filters.append(f)
        filters = ["All"] + sorted([f for f in filters if f != "All"])
        for f in filters:
            btn = tk.Button(self.filter_frame, text=f, width=12, bg="#666", fg="white", command=lambda t=f: self.set_filter(t))
            btn.pack(side="left", padx=4)
            self.filter_buttons[f] = btn
        if self.active_filter not in self.filter_buttons:
            self.active_filter = "All"
        self._update_filter_highlight()

    def _update_filter_highlight(self):
        for name, btn in self.filter_buttons.items():
            if name == (self.active_filter or "All"):
                btn.config(bg=self.select_bg, fg="white")
            else:
                btn.config(bg="#666", fg="white")

    def set_filter(self, filter_text):
        self.active_filter = filter_text
        self.current_index = 0
        self._update_filter_highlight()
        self.refresh_game_list()

    def load_from_file(self, path):
        try:
            self.games.clear()
            self.current_index = 0
            if not os.path.exists(path):
                messagebox.showwarning("Missing File", f"File not found: {path}")
                return
            ext = os.path.splitext(path)[1].lower()
            if ext == ".json":
                data = _robust_json_load(path)
                if isinstance(data, list):
                    for item in data:
                        if not isinstance(item, dict):
                            continue
                        run = item.get("run") or item.get("path") or ""
                        admin = bool(item.get("admin"))
                        name = item.get("name") or ""
                        filters = normalize_filters(item.get("filters", item.get("filter", [])))
                        if not name:
                            lower = (run or "").lower()
                            if lower.startswith("steam://"):
                                m = re.search(r"(\d+)", run)
                                appid = m.group(1) if m else None
                                if appid:
                                    name_try = fetch_steam_name(appid)
                                    if name_try:
                                        name = name_try
                                    else:
                                        name = f"Steam {appid}"
                                else:
                                    name = os.path.splitext(os.path.basename(run))[0]
                            else:
                                ext2 = os.path.splitext(run)[1].lower()
                                resolved = resolve_shortcut(run) if ext2 in [".lnk", ".url"] else run
                                name = os.path.splitext(os.path.basename(resolved or run))[0]
                        entry = resolve_game_entry(run, admin, name_override=name, filters=filters)
                        self.games.append(entry)
                else:
                    raise ValueError("JSON root is not a list")
            else:
                with open(path, "r", encoding="utf-8") as f:
                    for line in f:
                        raw = line.strip().replace('"', "")
                        if not raw:
                            continue
                        is_admin = raw.startswith("!admin ")
                        game_path = raw[7:] if is_admin else raw
                        entry = resolve_game_entry(game_path, is_admin, filters=[])
                        self.games.append(entry)
            self.refresh_game_list()
            self.update_filter_tabs()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file.\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = GamesListApp(root)
    root.mainloop()