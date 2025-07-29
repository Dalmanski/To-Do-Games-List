import tkinter as tk
from tkinter import messagebox
import os
import winreg
import tkinter.filedialog as filedialog

def get_installed_apps():
    apps = []
    reg_paths = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]
    hives = [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]

    for hive in hives:
        for reg_path in reg_paths:
            try:
                with winreg.OpenKey(hive, reg_path) as key:
                    for i in range(winreg.QueryInfoKey(key)[0]):
                        try:
                            subkey_name = winreg.EnumKey(key, i)
                            with winreg.OpenKey(key, subkey_name) as subkey:
                                name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                                install_values = [winreg.EnumValue(subkey, j)[0] for j in range(winreg.QueryInfoKey(subkey)[1])]
                                path = winreg.QueryValueEx(subkey, "InstallLocation")[0] if "InstallLocation" in install_values else ""
                                if path and os.path.isdir(path):
                                    apps.append((name, path))
                        except Exception:
                            continue
            except Exception:
                continue
    return sorted(apps, key=lambda x: x[0].lower())

def open_all_app_popup(master, on_select_callback):
    popup = tk.Toplevel(master)
    popup.title("Select Game or App")
    popup.geometry("500x500")
    popup.configure(bg="#1e1e1e")

    apps = get_installed_apps()
    filtered_apps = apps.copy()

    search_var = tk.StringVar()
    search_entry = tk.Entry(popup, textvariable=search_var, font=("Consolas", 12),
                            bg="#2b2b2b", fg="gray", insertbackground="white")
    search_entry.insert(0, "Search apps...")
    search_entry.pack(fill="x", padx=10, pady=(10, 0))

    def on_entry_focus_in(event):
        if search_entry.get() == "Search apps...":
            search_entry.delete(0, tk.END)
            search_entry.config(fg="white")

    def on_entry_focus_out(event):
        if not search_entry.get():
            search_entry.insert(0, "Search apps...")
            search_entry.config(fg="gray")

    search_entry.bind("<FocusIn>", on_entry_focus_in)
    search_entry.bind("<FocusOut>", on_entry_focus_out)

    list_frame = tk.Frame(popup, bg="#1e1e1e")
    list_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))

    scrollbar = tk.Scrollbar(list_frame)
    scrollbar.pack(side="right", fill="y")

    listbox = tk.Listbox(list_frame, font=("Consolas", 12), bg="#2b2b2b",
                         fg="white", selectbackground="#444", yscrollcommand=scrollbar.set)
    listbox.pack(fill="both", expand=True)
    scrollbar.config(command=listbox.yview)

    def _on_mousewheel(event):
        listbox.yview_scroll(-1 * int(event.delta / 120), "units")
    def _on_linux_scroll(event):
        listbox.yview_scroll(1 if event.num == 5 else -1, "units")

    listbox.bind("<MouseWheel>", _on_mousewheel)
    listbox.bind("<Button-4>", _on_linux_scroll) 
    listbox.bind("<Button-5>", _on_linux_scroll) 

    def update_list():
        search_term = search_var.get().lower()
        if search_term == "search apps...":
            search_term = ""
        listbox.delete(0, tk.END)
        filtered_apps.clear()
        for name, path in apps:
            if search_term in name.lower():
                filtered_apps.append((name, path))
                listbox.insert(tk.END, name)

    search_var.trace_add("write", lambda *args: update_list())
    update_list()

    def on_select():
        idx = listbox.curselection()
        if not idx:
            messagebox.showwarning("No Selection", "Please select an app.")
            return
        name = listbox.get(idx[0])
        path = next((p for n, p in apps if n == name), "")
        exe = None
        for root, _, files in os.walk(path):
            for file in files:
                if file.lower().endswith(".exe"):
                    exe = os.path.join(root, file)
                    break
            if exe:
                break
        if exe:
            on_select_callback(name, exe)
            popup.destroy()
        else:
            messagebox.showerror("Not Found", "No executable found in app folder.")

    def manual_browse():
        exe_path = filedialog.askopenfilename(
            title="Select Executable or Shortcut",
            filetypes=[("Executables and Shortcuts", "*.exe *.lnk")]
        )
        if exe_path:
            name = os.path.splitext(os.path.basename(exe_path))[0]
            on_select_callback(name, exe_path)
            popup.destroy()

    btn_frame = tk.Frame(popup, bg="#1e1e1e")
    btn_frame.pack(pady=(0, 10))

    tk.Button(btn_frame, text="Add Selected", font=("Arial", 11, "bold"),
              bg="#3c8dbc", fg="white", command=on_select).pack(side="left", padx=5)

    tk.Button(btn_frame, text="Browse for .exe", font=("Arial", 11),
              bg="#555", fg="white", command=manual_browse).pack(side="left", padx=5)

if __name__ == "__main__":
    def on_app_selected(name, exe_path):
        print(f"Selected App: {name}")
        print(f"Executable Path: {exe_path}")

    root = tk.Tk()
    root.withdraw()
    open_all_app_popup(root, on_app_selected)
    root.mainloop()
