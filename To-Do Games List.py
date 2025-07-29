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
from edit_list import open_txt_popup
from help import open_help_popup
from settings import open_settings_popup
from all_app import open_all_app_popup

if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

settings_path = os.path.join(base_dir, "settings.json")
try:
    with open(settings_path, "r", encoding="utf-8") as f:
        settings = json.load(f)
        default_filename = os.path.join(base_dir, settings.get("Gamelist", "To-Do Games List.txt"))
except Exception as e:
    messagebox.showerror("Settings Error", f"Could not load settings.json\n{e}")
    default_filename = os.path.join(base_dir, "To-Do Games List.txt")

def resolve_shortcut(path):
    try:
        return win32com.client.Dispatch("WScript.Shell").CreateShortcut(path).TargetPath
    except:
        return path

def extract_icon(path):
    try:
        if not os.path.exists(path): return None
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
            return Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), bmpstr, 'raw', 'BGRX', 0, 1)
    except:
        return None

class GamesListApp:
    def __init__(self, root):
        self.root = root
        self.root.iconbitmap(os.path.join(base_dir, "icon.ico"))
        self.root.title("To-Do Games List")
        self.root.resizable(False, False)
        self.auto_play = tk.BooleanVar(value=True)
        self.current_index = 0
        self.games, self.icon_images, self.item_frames = [], [], []
        self.bullet_image = self.make_bullet_image()

        self.bg_color, self.fg_color = "#1e1e1e", "#ffffff"
        self.select_bg, self.btn_color = "#345d9d", "#3c8dbc"
        self.num_list_bg, self.list_bg = "#141414", "#363636"

        self.root.configure(bg=self.bg_color)
        self.create_widgets()
        self.load_from_file(default_filename)
        self.center_window()

    def make_bullet_image(self):
        img = Image.new("RGBA", (32, 32), (0, 0, 0, 0))
        ImageDraw.Draw(img).ellipse((10, 10, 22, 22), fill="white")
        return ImageTk.PhotoImage(img)

    def center_window(self):
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2) - 40
        self.root.geometry(f"+{x}+{y}")

    def create_widgets(self):
        top = tk.Frame(self.root, bg=self.bg_color)
        top.pack(fill="x", padx=10, pady=(10, 0))
        self.list_label = tk.Label(top, font=("Arial", 12, "bold"), bg=self.bg_color, fg=self.fg_color)
        self.list_label.pack(side="left")
        tk.Button(top, text="Create New List", width=14, bg="#666", fg="white", command=self.create_list_dialog).pack(side="right")
        self.update_list_label()

        self.main_frame = tk.Frame(self.root, bg=self.bg_color)
        self.main_frame.pack(padx=10, pady=10)

        self.canvas = tk.Canvas(self.main_frame, width=430, height=420, bg=self.list_bg, highlightthickness=0)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar = tk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scroll_container = tk.Frame(self.canvas, bg=self.list_bg)
        self.canvas.create_window((0, 0), window=self.scroll_container, anchor="nw")
        self.scroll_container.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.number_frame = tk.Frame(self.scroll_container, bg=self.num_list_bg)
        self.number_frame.pack(side="left", fill="y")
        self.scroll_frame = tk.Frame(self.scroll_container, bg=self.list_bg)
        self.scroll_frame.pack(side="left", fill="both", expand=True)

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

        self.auto_button = tk.Button(self.root, text="Auto Play: ON", font=("Arial", 12, "bold"),
                                     bg="#3cb371", fg="white", activebackground="#2e6fa3", relief="ridge", bd=3,
                                     width=18, command=self.toggle_autoplay)
        self.auto_button.pack(pady=5)

        control = tk.Frame(self.root, bg=self.bg_color)
        control.pack(pady=5)

        top_row = tk.Frame(control, bg=self.bg_color)
        top_row.pack(pady=5)
        tk.Button(top_row, text="Load List (.txt)", width=12, bg="#666", fg="white", command=self.load_dialog).pack(side="left", padx=5)
        tk.Button(top_row, text="Add Game", width=12, bg="#005b29", fg="white", command=self.add_game_dialog).pack(side="left", padx=5)
        tk.Button(top_row, text="Help", width=12, bg="#444", fg="white", command=open_help_popup).pack(side="left", padx=5)

        bottom_row = tk.Frame(control, bg=self.bg_color)
        bottom_row.pack(pady=5)
        tk.Button(bottom_row, text="Edit List", width=12, bg="#666", fg="white", 
                  command=lambda: open_txt_popup(self.root, self.bg_color, self.list_bg, self.fg_color, self.load_from_file)).pack(side="left", padx=5)
        tk.Button(bottom_row, text="Remove Game", width=12, bg="#5B0000", fg="white", command=self.remove_selected_game).pack(side="left", padx=5)
        tk.Button(bottom_row, text="Settings", width=12, bg="#444", fg="white", command=open_settings_popup).pack(side="left", padx=5)

    def clear_frame(self, frame):
        for widget in frame.winfo_children(): widget.destroy()

    def add_game_widget(self, game, index):
        frame = tk.Frame(self.scroll_frame, bg=self.list_bg)
        frame.pack(fill="x", pady=2)
        self.item_frames.append(frame)

        icon = self.load_icon_image(game["real"])
        self.icon_images.append(icon)
        tk.Label(frame, image=icon, bg=self.list_bg).pack(side="left", padx=5)
        label = tk.Label(frame, text=game["name"], font=("Consolas", 16), bg=self.list_bg,
                         fg=self.fg_color, anchor="w")
        label.pack(side="left", fill="x", expand=True)
        label.bind("<Button-1>", lambda e, idx=index: self.select_game(idx))

    def refresh_game_list(self):
        self.clear_frame(self.scroll_frame)
        self.clear_frame(self.number_frame)
        self.item_frames.clear()
        self.icon_images.clear()

        if not self.games:
            self.empty_label = tk.Label(self.scroll_frame, text="Pls add new game", font=("Arial", 16, "bold"),
                                        bg=self.list_bg, fg="#888")
            self.empty_label.pack(padx=100, pady=180)
        else:
            for idx, game in enumerate(self.games):
                self.add_game_widget(game, idx)
                tk.Label(self.number_frame, text=f"{idx + 1}.", font=("Consolas", 16),
                        bg=self.num_list_bg, fg=self.fg_color, anchor="e", width=4).pack(anchor="n", pady=5)
            self.highlight_current()

    def add_game_dialog(self):
        def on_game_selected(name, exe_path):
            self.games.append({
                "name": name,
                "run": exe_path,
                "real": exe_path,
                "admin": False
            })
            self.current_index = len(self.games) - 1
            self.refresh_game_list()
            self.auto_save()
        open_all_app_popup(self.root, on_game_selected)

    def remove_selected_game(self):
        if not self.games:
            messagebox.showwarning("No Game", "No game selected to remove.")
            return
        game = self.games[self.current_index]
        if messagebox.askyesno("Remove Game", f"Are you sure you want to remove:\n\n{game['name']}?"):
            del self.games[self.current_index]
            self.current_index = min(self.current_index, len(self.games) - 1)
            self.refresh_game_list()
            self.auto_save()

    def create_list_dialog(self):
        path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
        if path:
            try:
                with open(path, "w", encoding="utf-8"): pass
                self.games.clear()
                self.current_index = 0
                global default_filename
                default_filename = path
                self.refresh_game_list()
                self.auto_save()
                self.update_list_label()
                messagebox.showinfo("Created", f"New list created:\n{path}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not create file.\n{e}")

    def auto_save(self):
        try:
            with open(default_filename, "w", encoding="utf-8") as f:
                for game in self.games:
                    prefix = "!admin " if game["admin"] else ""
                    f.write(f'{prefix}"{game["run"]}"\n')
        except Exception as e:
            messagebox.showerror("Auto Save Error", f"Could not save automatically.\n{e}")

    def load_icon_image(self, path):
        icon_img = extract_icon(path)
        return ImageTk.PhotoImage(icon_img.resize((32, 32), Image.LANCZOS)) if icon_img else self.bullet_image

    def _on_mousewheel(self, e):
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
                widget.configure(bg=bg)
            frame.configure(bg=bg)

    def go_left(self):
        if self.current_index > 0:
            self.current_index -= 1
            self.highlight_current()
            if self.auto_play.get():
                self.launch_game()

    def go_right(self):
        if self.current_index < len(self.games) - 1:
            self.current_index += 1
            self.highlight_current()
            if self.auto_play.get():
                self.launch_game()

    def launch_game(self):
        game = self.games[self.current_index]
        if game["admin"] and not messagebox.askyesno("Admin Launch", f"Launch {game['name']} with admin?"):
            return
        try:
            os.startfile(game["run"])
        except Exception as e:
            messagebox.showerror("Error", f"Could not run the game.\n{e}")

    def toggle_autoplay(self):
        val = not self.auto_play.get()
        self.auto_play.set(val)
        self.auto_button.config(text=f"Auto Play: {'ON' if val else 'OFF'}", bg="#3cb371" if val else "#777")

    def load_dialog(self):
        path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if path:
            self.load_from_file(path)
            global default_filename
            default_filename = path
            self.update_settings_json(path)
            self.update_list_label()

    def update_settings_json(self, path):
        try:
            settings = {}
            if os.path.exists(settings_path):
                with open(settings_path, "r", encoding="utf-8") as f:
                    settings = json.load(f)

            settings["Gamelist"] = os.path.basename(path)

            with open(settings_path, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            messagebox.showerror("Error", f"Could not update settings.json\n{e}")

    def update_list_label(self):
        self.list_label.config(text=os.path.basename(default_filename).replace(".txt", ""))

    def load_from_file(self, path):
        try:
            self.games.clear()
            if not os.path.exists(path):
                messagebox.showwarning("Missing File", f"File not found: {path}")
                return
            with open(path, "r", encoding="utf-8") as f:
                for line in f:
                    raw = line.strip().replace('"', '')
                    if not raw:
                        continue
                    is_admin = raw.startswith("!admin ")
                    game_path = raw[7:] if is_admin else raw
                    ext = os.path.splitext(game_path)[1].lower()
                    resolved_path = resolve_shortcut(game_path) if ext in [".lnk", ".url"] else game_path
                    self.games.append({
                        "name": os.path.splitext(os.path.basename(game_path))[0],
                        "run": game_path,
                        "real": resolved_path,
                        "admin": is_admin
                    })
            self.refresh_game_list()
            self.current_index = 0
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file.\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = GamesListApp(root)
    root.mainloop()
