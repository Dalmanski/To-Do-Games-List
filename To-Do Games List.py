import tkinter as tk
from tkinter import messagebox, filedialog
import os
from PIL import Image, ImageTk, ImageDraw
import win32com.client
import win32gui
import win32con
import win32ui
import sys

if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)  # Running from .exe
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))  # Running from .py
    
default_filename = os.path.join(base_dir, "To-Do Games List.txt")

def resolve_shortcut(path):
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(path)
        return shortcut.TargetPath
    except:
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
            img = Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), bmpstr, 'raw', 'BGRX', 0, 1)
            win32gui.DestroyIcon(hicon)
            return img
    except:
        return None
    return None

class GachaListApp:
    def __init__(self, root):
        self.root = root
        icon_path = os.path.join(base_dir, "icon.ico")
        self.root.iconbitmap(icon_path)
        self.root.title("To-Do Games List")
        self.current_index = 0
        self.games = []
        self.icon_images = []
        self.item_frames = []
        self.auto_play = tk.BooleanVar(value=True)
        self.bullet_image = self.make_bullet_image()

        self.bg_color = "#1e1e1e"
        self.fg_color = "#ffffff"
        self.select_bg = "#345d9d"
        self.btn_color = "#3c8dbc"
        self.list_bg = "#363636"

        self.root.configure(bg=self.bg_color)
        self.create_widgets()
        self.load_from_file(default_filename)
        self.center_window()

    def make_bullet_image(self):
        img = Image.new("RGBA", (32, 32), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.ellipse((10, 10, 22, 22), fill="white")
        return ImageTk.PhotoImage(img)

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"+{x}+{y}")

    def create_widgets(self):
        self.main_frame = tk.Frame(self.root, bg=self.bg_color)
        self.main_frame.pack(padx=10, pady=10)

        self.canvas = tk.Canvas(self.main_frame, bg=self.list_bg, highlightthickness=0, height=420)
        self.scrollbar = tk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.scroll_frame = tk.Frame(self.canvas, bg=self.list_bg)

        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.button_frame = tk.Frame(self.root, bg=self.bg_color)
        self.button_frame.pack(pady=10)

        self.left_button = tk.Button(self.button_frame, text="←", font=("Arial", 18, "bold"),
            width=5, height=2, bg=self.btn_color, fg="white", activebackground="#2e6fa3", command=self.go_left)
        self.left_button.pack(side="left", padx=10)

        self.play_button = tk.Button(self.button_frame, text="▶", font=("Arial", 18, "bold"),
            width=5, height=2, bg=self.btn_color, fg="white", activebackground="#2e6fa3", command=self.launch_game)
        self.play_button.pack(side="left", padx=10)

        self.right_button = tk.Button(self.button_frame, text="→", font=("Arial", 18, "bold"),
            width=5, height=2, bg=self.btn_color, fg="white", activebackground="#2e6fa3", command=self.go_right)
        self.right_button.pack(side="left", padx=10)

        auto_btn = tk.Button(self.root, text="Auto Play: ON", font=("Arial", 12, "bold"),
            bg="#3cb371", fg="white", activebackground="#2e6fa3", relief="ridge", bd=3, width=18,
            command=self.toggle_autoplay)
        auto_btn.pack(pady=5)
        self.auto_button = auto_btn

        # --- Control Buttons Layout (2 rows) ---
        control_frame = tk.Frame(self.root, bg=self.bg_color)
        control_frame.pack(pady=5)

        top_row = tk.Frame(control_frame, bg=self.bg_color)
        top_row.pack()
        bottom_row = tk.Frame(control_frame, bg=self.bg_color)
        bottom_row.pack(pady=5)

        tk.Button(top_row, text="Load TXT", width=12, bg="#666", fg="white", command=self.load_dialog).pack(side="left", padx=5)
        tk.Button(top_row, text="Add Game", width=12, bg="#666", fg="white", command=self.add_game_dialog).pack(side="left", padx=5)
        tk.Button(top_row, text="Delete Game", width=12, bg="#c05050", fg="white", command=self.delete_selected_game).pack(side="left", padx=5)


        tk.Button(bottom_row, text="Save TXT", width=12, bg="#666", fg="white", command=self.save_dialog).pack(side="left", padx=5)
        tk.Button(bottom_row, text="Create List", width=12, bg="#3cb371", fg="white", command=self.create_list_dialog).pack(side="left", padx=5)
        

    def add_game_widget(self, game, index):
        frame = tk.Frame(self.scroll_frame, bg=self.list_bg)
        frame.pack(fill="x", pady=2)
        self.item_frames.append(frame)

        icon = self.load_icon_image(game["real"])
        self.icon_images.append(icon)

        icon_label = tk.Label(frame, image=icon, bg=self.list_bg)
        icon_label.pack(side="left", padx=5)

        label = tk.Label(frame, text=game["name"], font=("Consolas", 16), bg=self.list_bg,
                         fg=self.fg_color, anchor="w")
        label.pack(side="left", fill="x", expand=True)
        label.bind("<Button-1>", lambda e, idx=index: self.select_game(idx))

    def create_list_dialog(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                filetypes=[("Text Files", "*.txt")])
        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write("")  # create empty txt
                self.games.clear()
                self.refresh_game_list()
                self.current_index = 0
                global default_filename
                default_filename = file_path
                messagebox.showinfo("Created", f"New list created:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not create file.\n{e}")

    def refresh_game_list(self):
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        self.item_frames.clear()
        self.icon_images.clear()
        for idx, game in enumerate(self.games):
            self.add_game_widget(game, idx)
        self.highlight_current()

    def add_game_dialog(self):
        file_path = filedialog.askopenfilename(filetypes=[("Executables or Shortcuts", "*.exe *.lnk")])
        if file_path:
            is_admin = False
            name = os.path.splitext(os.path.basename(file_path))[0]
            ext = os.path.splitext(file_path)[1].lower()
            resolved_path = resolve_shortcut(file_path) if ext == ".lnk" else file_path
            self.games.append({
                "name": name,
                "run": file_path,
                "real": resolved_path,
                "admin": is_admin
            })
            self.refresh_game_list()
            self.current_index = len(self.games) - 1
            self.highlight_current()

    def delete_selected_game(self):
        if not self.games:
            messagebox.showwarning("No Game", "No game selected to delete.")
            return
        game = self.games[self.current_index]
        confirm = messagebox.askyesno("Delete Game", f"Are you sure you want to delete:\n\n{game['name']}?")
        if confirm:
            del self.games[self.current_index]
            if self.current_index >= len(self.games):
                self.current_index = len(self.games) - 1
            self.refresh_game_list()

    def load_icon_image(self, path):
        icon_img = extract_icon(path)
        if icon_img:
            icon_img = icon_img.resize((32, 32), Image.LANCZOS)
            return ImageTk.PhotoImage(icon_img)
        else:
            return self.bullet_image

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(-1 * int(event.delta / 120), "units")

    def select_game(self, index):
        self.current_index = index
        self.highlight_current()
        if self.auto_play.get():
            self.launch_game()

    def highlight_current(self):
        for i, frame in enumerate(self.item_frames):
            bg = self.select_bg if i == self.current_index else self.list_bg
            frame.configure(bg=bg)
            for widget in frame.winfo_children():
                widget.configure(bg=bg)

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
        if game["admin"]:
            confirm = messagebox.askyesno("Admin Launch", f"This game may require admin rights:\n\n{game['name']}\n\nLaunch it?")
            if not confirm:
                return
        try:
            os.startfile(game["run"])
        except Exception as e:
            messagebox.showerror("Error", f"Could not run the game.\n{e}")

    def toggle_autoplay(self):
        self.auto_play.set(not self.auto_play.get())
        self.auto_button.config(
            text=f"Auto Play: {'ON' if self.auto_play.get() else 'OFF'}",
            bg="#3cb371" if self.auto_play.get() else "#777"
        )

    def load_dialog(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if file_path:
            self.load_from_file(file_path)

    def save_dialog(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                 filetypes=[("Text Files", "*.txt")])
        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    for game in self.games:
                        prefix = "!admin " if game["admin"] else ""
                        f.write(f'{prefix}"{game["run"]}"\n')
                messagebox.showinfo("Saved", "Game list saved successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Could not save file.\n{e}")

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
                    path = raw[7:] if is_admin else raw
                    name = os.path.splitext(os.path.basename(path))[0]
                    ext = os.path.splitext(path)[1].lower()
                    resolved_path = resolve_shortcut(path) if ext in [".lnk", ".url"] else path
                    self.games.append({
                        "name": name,
                        "run": path,
                        "real": resolved_path,
                        "admin": is_admin
                    })
            self.refresh_game_list()
            self.current_index = 0
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file.\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = GachaListApp(root)
    root.mainloop()
