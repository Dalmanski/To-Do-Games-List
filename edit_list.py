import os
import json
import tkinter as tk
from tkinter import messagebox, filedialog
import re
import sys

class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        self.after_id = None
        widget.bind("<Enter>", self.schedule)
        widget.bind("<Leave>", self.hide)
        widget.bind("<ButtonPress>", self.hide)

    def schedule(self, event=None):
        self.cancel()
        self.after_id = self.widget.after(450, self.show)

    def cancel(self):
        if self.after_id is not None:
            try:
                self.widget.after_cancel(self.after_id)
            except:
                pass
            self.after_id = None

    def show(self, event=None):
        if self.tipwindow or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
            self.tipwindow = tw = tk.Toplevel(self.widget)
            tw.wm_overrideredirect(True)
            tw.wm_geometry(f"+{x}+{y}")
            label = tk.Label(
                tw,
                text=self.text,
                justify="left",
                bg="#222222",
                fg="white",
                relief="solid",
                borderwidth=1,
                font=("Segoe UI", 9),
                padx=8,
                pady=4
            )
            label.pack()
        except:
            self.tipwindow = None

    def hide(self, event=None):
        self.cancel()
        if self.tipwindow:
            try:
                self.tipwindow.destroy()
            except:
                pass
            self.tipwindow = None

def open_txt_popup(parent, bg_color, list_bg, fg_color, load_from_file_callback):
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    settings_path = os.path.join(base_dir, "settings.json")
    try:
        with open(settings_path, "r", encoding="utf-8") as f:
            settings = json.load(f)
    except:
        settings = {}

    file_basename = settings.get("Gamelist")
    if not file_basename:
        messagebox.showerror("No List Set", "No game list configured in settings.json.\n\nPlease create or load a list from the main window first.")
        return

    if os.path.isabs(file_basename):
        file_path = file_basename
    else:
        file_path = os.path.join(base_dir, file_basename)

    if not os.path.exists(file_path):
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write("[]")
        except Exception as e:
            messagebox.showerror("Error", f"Could not create list file.\n{e}")
            return

    def _robust_load(path):
        with open(path, "r", encoding="utf-8") as f:
            raw = f.read()
        raw = raw.lstrip("\ufeff").strip()
        try:
            data = json.loads(raw)
            if isinstance(data, list):
                return data, raw
            raise ValueError("JSON root is not a list")
        except Exception:
            start = raw.find("[")
            end = raw.rfind("]")
            if start != -1 and end != -1 and end > start:
                chunk = raw[start:end + 1]
                chunk = re.sub(r",\s*([\]\}])", r"\1", chunk)
                chunk = re.sub(r'(?<!\\)\\(?!\\)', r'\\\\', chunk)
                try:
                    data = json.loads(chunk)
                    if isinstance(data, list):
                        return data, chunk
                except Exception:
                    pass
            objs = re.findall(r"\{(?:[^{}]|\n|\r)*?\}", raw)
            parsed = []
            for o in objs:
                o2 = re.sub(r",\s*([\}\]])", r"\1", o)
                o2 = re.sub(r'(?<!\\)\\(?!\\)', r'\\\\', o2)
                try:
                    parsed.append(json.loads(o2))
                except Exception:
                    continue
            if parsed:
                return parsed, json.dumps(parsed, indent=4)
            raise ValueError("Could not parse JSON")

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
                except Exception:
                    pass
            parts = re.split(r"[,\n;|]+", text)
            return [part.strip().lower() for part in parts if part.strip()]
        text = str(value).strip().lower()
        return [text] if text else []

    def normalize_entry(entry):
        if not isinstance(entry, dict):
            entry = {}
        entry["run"] = entry.get("run", "") or ""
        entry["name"] = entry.get("name", "") or ""
        entry["admin"] = bool(entry.get("admin", False))
        entry["filters"] = normalize_filters(entry.get("filters", []))
        entry.pop("filter", None)
        return entry

    try:
        entries, content = _robust_load(file_path)
        if not isinstance(entries, list):
            entries = []
        entries = [normalize_entry(e) for e in entries]
    except Exception as e:
        messagebox.showerror("Error", f"Could not open file or invalid JSON.\n{e}")
        return

    popup = tk.Toplevel(parent)
    popup.title(f"Editing: {os.path.basename(file_path)}")
    popup.resizable(False, False)
    popup.configure(bg=bg_color)

    left_frame = tk.Frame(popup, bg=bg_color)
    left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="n")

    tk.Label(left_frame, text="Games", bg=bg_color, fg=fg_color, font=("Consolas", 12, "bold")).pack(anchor="w")

    list_container = tk.Frame(left_frame, bg=list_bg, bd=0, relief="flat")
    list_container.pack(pady=(6, 8))

    listbox = tk.Listbox(
        list_container,
        width=38,
        height=18,
        bd=0,
        highlightthickness=0,
        activestyle="none",
        font=("Consolas", 12),
        bg=list_bg,
        fg=fg_color,
        selectbackground="#0055ff",
        selectforeground=fg_color,
        exportselection=False
    )
    listbox.pack(side="left", fill="y")

    list_scroll = tk.Scrollbar(list_container, command=listbox.yview)
    list_scroll.pack(side="right", fill="y")
    listbox.config(yscrollcommand=list_scroll.set)

    info_frame = tk.Frame(left_frame, bg=bg_color)
    info_frame.pack(fill="x")

    tk.Label(info_frame, text="Run (path):", bg=bg_color, fg=fg_color).grid(row=0, column=0, sticky="w")
    run_var = tk.StringVar()
    run_entry = tk.Entry(info_frame, textvariable=run_var, width=36)
    run_entry.grid(row=1, column=0, sticky="w")

    def browse_run():
        p = filedialog.askopenfilename(initialdir=base_dir)
        if p:
            run_var.set(p)

    browse_btn = tk.Button(info_frame, text="Browse", width=8, command=browse_run, bg="#666", fg="white")
    browse_btn.grid(row=1, column=1, padx=(6, 0), sticky="w")
    Tooltip(browse_btn, "Browse run path")

    reorder_frame = tk.Frame(info_frame, bg=bg_color)
    reorder_frame.grid(row=1, column=2, sticky="w", padx=(6, 0))

    tk.Label(info_frame, text="Game Name:", bg=bg_color, fg=fg_color).grid(row=2, column=0, sticky="w", pady=(8, 0))
    name_var = tk.StringVar()
    name_entry = tk.Entry(info_frame, textvariable=name_var, width=36)
    name_entry.grid(row=3, column=0, sticky="w")

    is_admin_var = tk.BooleanVar(value=False)
    admin_chk = tk.Checkbutton(info_frame, text="isAdmin", variable=is_admin_var, bg=bg_color, fg=fg_color, selectcolor="#333")
    admin_chk.grid(row=3, column=1, sticky="w", padx=(6, 0))

    filter_header = tk.Frame(info_frame, bg=bg_color)
    filter_header.grid(row=4, column=0, columnspan=3, sticky="we", pady=(8, 0))
    tk.Label(filter_header, text="Filter:", bg=bg_color, fg=fg_color).pack(side="left")

    filter_row_frame = tk.Frame(info_frame, bg=bg_color)
    filter_row_frame.grid(row=5, column=0, columnspan=3, sticky="we", pady=(4, 0))

    filter_canvas_container = tk.Frame(filter_row_frame, bg=bg_color)
    filter_canvas_container.pack(side="left", fill="x", expand=True)

    filter_canvas = tk.Canvas(
        filter_canvas_container,
        bg=bg_color,
        highlightthickness=0,
        bd=0,
        height=44
    )
    filter_canvas.pack(side="top", fill="x", expand=True)

    filter_hscroll = tk.Scrollbar(filter_canvas_container, orient="horizontal", command=filter_canvas.xview)
    filter_hscroll.pack(side="bottom", fill="x")
    filter_canvas.configure(xscrollcommand=filter_hscroll.set)

    filter_tags_inner = tk.Frame(filter_canvas, bg=bg_color)
    filter_tags_window = filter_canvas.create_window((0, 0), window=filter_tags_inner, anchor="nw")

    def update_filter_scrollregion(event=None):
        filter_canvas.configure(scrollregion=filter_canvas.bbox("all"))

    filter_tags_inner.bind("<Configure>", update_filter_scrollregion)

    filter_add_frame = tk.Frame(filter_row_frame, bg=bg_color)
    filter_add_frame.pack(side="left", padx=(6, 0))

    action_frame = tk.Frame(info_frame, bg=bg_color)
    action_frame.grid(row=6, column=0, columnspan=3, sticky="w", pady=(10, 0))

    status_label = tk.Label(left_frame, text="", bg=bg_color, fg=fg_color)
    status_label.pack(anchor="w", pady=(6, 0))

    center_frame = tk.Frame(popup, bg=bg_color)
    center_frame.grid(row=0, column=1, padx=10, pady=10)

    tk.Label(center_frame, text="JSON Code Editor", bg=bg_color, fg=fg_color, font=("Consolas", 12, "bold")).pack(anchor="w")

    editor_outer = tk.Frame(center_frame, bg=bg_color)
    editor_outer.pack(fill="both", expand=True, pady=(6, 0))

    line_numbers = tk.Text(
        editor_outer,
        width=4,
        padx=4,
        pady=0,
        takefocus=0,
        border=0,
        background="#2b2b2b",
        foreground="gray",
        state="disabled",
        font=("Consolas", 14),
        height=28,
        wrap="none"
    )
    line_numbers.grid(row=0, column=0, sticky="nsew", pady=0)

    text_widget = tk.Text(
        editor_outer,
        wrap="none",
        bg=list_bg,
        fg=fg_color,
        insertbackground="white",
        font=("Consolas", 14),
        undo=True,
        width=64,
        height=28,
        pady=0
    )
    text_widget.grid(row=0, column=1, sticky="nsew", padx=(0, 0), pady=0)

    def sync_scrollbar(*args):
        text_widget.yview(*args)
        line_numbers.yview(*args)

    scrollbar = tk.Scrollbar(editor_outer, command=sync_scrollbar)
    scrollbar.grid(row=0, column=2, sticky="nsew", pady=0)

    def on_scroll_text(*args):
        scrollbar.set(*args)
        line_numbers.yview_moveto(args[0])

    def on_scroll_line(*args):
        scrollbar.set(*args)
        text_widget.yview_moveto(args[0])

    text_widget.config(yscrollcommand=on_scroll_text)
    line_numbers.config(yscrollcommand=on_scroll_line)

    def sync_scroll(event=None):
        line_numbers.yview_moveto(text_widget.yview()[0])

    text_widget.bind("<MouseWheel>", sync_scroll)
    text_widget.bind("<Button-4>", sync_scroll)
    text_widget.bind("<Button-5>", sync_scroll)
    text_widget.bind("<Up>", sync_scroll)
    text_widget.bind("<Down>", sync_scroll)
    text_widget.bind("<Prior>", sync_scroll)
    text_widget.bind("<Next>", sync_scroll)

    editor_outer.grid_rowconfigure(0, weight=1)
    editor_outer.grid_columnconfigure(1, weight=1)

    editor_bottom_frame = tk.Frame(center_frame, bg=bg_color)
    editor_bottom_frame.pack(fill="x", pady=(6, 0))

    bottom_right_frame = tk.Frame(editor_bottom_frame, bg=bg_color)
    bottom_right_frame.pack(side="right")

    countdown_label = tk.Label(bottom_right_frame, text="", bg=bg_color, fg=fg_color, font=("Consolas", 11))
    countdown_label.pack(side="left", padx=(0, 10))

    top_status_label = tk.Label(bottom_right_frame, text="", bg=bg_color, fg=fg_color, font=("Consolas", 11))
    top_status_label.pack(side="left", padx=(0, 10))

    auto_save = tk.BooleanVar(value=settings.get("AutoSave", False))

    filter_vars = []
    filter_entry_widgets = []
    loading_selection = False
    updating_editor = False
    countdown_seconds_default = 5
    countdown_current = 0
    countdown_job = None

    drag_state = {
        "press_index": None,
        "dragging": False,
        "armed": False,
        "after_id": None,
        "preview": None,
        "last_x_root": 0,
        "last_y_root": 0,
        "moved": False,
        "drag_index": None,
        "drag_color": "#ff9f1a",
        "drag_fg": "#111111"
    }

    def update_line_numbers():
        lines = text_widget.get("1.0", "end-1c").splitlines()
        line_numbers.config(state="normal")
        line_numbers.delete("1.0", "end")
        if not lines:
            line_numbers.insert("end", "1")
        else:
            for i in range(1, len(lines) + 1):
                line_numbers.insert("end", f"{i}")
                if i < len(lines):
                    line_numbers.insert("end", "\n")
        line_numbers.config(state="disabled")
        line_numbers.yview_moveto(text_widget.yview()[0])

    def refresh_editor_from_entries():
        nonlocal updating_editor
        updating_editor = True
        try:
            text_widget.delete("1.0", "end")
            text_widget.insert("1.0", json.dumps(entries, indent=4))
            text_widget.edit_modified(False)
            update_line_numbers()
        finally:
            updating_editor = False

    def normalize_all_entries():
        for e in entries:
            normalize_entry(e)

    def write_entries():
        normalize_all_entries()
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(entries, f, indent=4)
        refresh_editor_from_entries()
        load_from_file_callback(file_path)

    def get_display_name(entry, idx):
        display = (entry.get("name") or "").strip()
        if display:
            return display
        run_name = os.path.splitext(os.path.basename(entry.get("run", "")))[0].strip()
        if run_name:
            return run_name
        return f"Unknown Game {idx + 1}"

    def apply_listbox_styles():
        size = listbox.size()
        for i in range(size):
            listbox.itemconfig(i, background=list_bg, foreground=fg_color, selectbackground="#0055ff", selectforeground=fg_color)
        sel = listbox.curselection()
        if sel:
            idx = sel[0]
            if 0 <= idx < size:
                listbox.itemconfig(idx, background="#0055ff", foreground=fg_color, selectbackground="#0055ff", selectforeground=fg_color)
        if drag_state["dragging"] and drag_state["drag_index"] is not None:
            idx = drag_state["drag_index"]
            if 0 <= idx < size:
                listbox.itemconfig(idx, background=drag_state["drag_color"], foreground=drag_state["drag_fg"], selectbackground=drag_state["drag_color"], selectforeground=drag_state["drag_fg"])

    def populate_listbox():
        listbox.delete(0, "end")
        for i, e in enumerate(entries):
            listbox.insert("end", f"{i + 1}. {get_display_name(e, i)}")
        apply_listbox_styles()

    def clear_filter_tags():
        for child in filter_tags_inner.winfo_children():
            child.destroy()
        filter_vars.clear()
        filter_entry_widgets.clear()
        update_filter_scrollregion()

    def get_filters_from_ui():
        filters = []
        for var in filter_vars:
            text = var.get().strip().lower()
            if text:
                filters.append(text)
        return filters

    def add_filter_button():
        for child in filter_add_frame.winfo_children():
            child.destroy()
        plus_btn = tk.Button(
            filter_add_frame,
            text="＋",
            width=3,
            bg="#666",
            fg="white",
            font=("Segoe UI Symbol", 11, "bold"),
            bd=0,
            relief="raised"
        )
        plus_btn.pack(side="left", pady=(0, 6))
        Tooltip(plus_btn, "Add filter")
        plus_btn.config(command=lambda: add_filter_tag(""))

    def add_filter_tag(value=""):
        tag_frame = tk.Frame(filter_tags_inner, bg="#3a3a3a", bd=1, relief="solid")
        tag_frame.pack(side="left", padx=(0, 6), pady=(0, 6))

        var = tk.StringVar(value=value)
        filter_vars.append(var)

        entry = tk.Entry(
            tag_frame,
            textvariable=var,
            width=14,
            bd=0,
            relief="flat",
            bg="#3a3a3a",
            fg=fg_color,
            insertbackground="white"
        )
        entry.pack(side="left", padx=(6, 2), pady=2)
        filter_entry_widgets.append(entry)

        def on_tag_change(*args):
            if not loading_selection:
                mark_modified()

        var.trace_add("write", on_tag_change)

        current_index = len(filter_vars) - 1

        def remove_self(index=current_index):
            current = get_filters_from_ui()
            if 0 <= index < len(current):
                del current[index]
            render_filter_tags(current)
            mark_modified()

        remove_btn = tk.Button(
            tag_frame,
            text="×",
            width=2,
            command=remove_self,
            bg="#5B0000",
            fg="white",
            bd=0,
            font=("Segoe UI Symbol", 10, "bold"),
            relief="raised"
        )
        remove_btn.pack(side="left", padx=(0, 4), pady=2)
        Tooltip(remove_btn, "Remove filter")
        update_filter_scrollregion()
        return entry

    def render_filter_tags(filters):
        clear_filter_tags()
        filters = normalize_filters(filters)
        if not filters:
            add_filter_tag("")
        else:
            for value in filters:
                add_filter_tag(value)
        add_filter_button()
        update_filter_scrollregion()
        filter_canvas.xview_moveto(1.0)

    def load_selected_into_ui(idx):
        nonlocal loading_selection
        loading_selection = True
        try:
            entry = entries[idx]
            run_var.set(entry.get("run", ""))
            name_var.set(entry.get("name", ""))
            is_admin_var.set(bool(entry.get("admin", False)))
            render_filter_tags(entry.get("filters", []))
        finally:
            loading_selection = False

    def on_select(event=None):
        if drag_state["dragging"]:
            return
        sel = listbox.curselection()
        if not sel:
            run_var.set("")
            name_var.set("")
            is_admin_var.set(False)
            render_filter_tags([])
            apply_listbox_styles()
            return
        load_selected_into_ui(sel[0])
        apply_listbox_styles()

    def sync_current_entry_from_ui(idx):
        entries[idx]["run"] = run_var.get()
        entries[idx]["name"] = name_var.get()
        entries[idx]["admin"] = bool(is_admin_var.get())
        entries[idx]["filters"] = get_filters_from_ui()

    def save_selected():
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("No selection", "Select a game first.")
            return
        idx = sel[0]
        sync_current_entry_from_ui(idx)
        try:
            write_entries()
            status_label.config(text="Saved selected game.", fg="#3cb371")
            populate_listbox()
            listbox.selection_set(idx)
            listbox.activate(idx)
            listbox.see(idx)
            load_selected_into_ui(idx)
            apply_listbox_styles()
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file.\n{e}")

    def save_all():
        sel = listbox.curselection()
        if sel:
            sync_current_entry_from_ui(sel[0])
        try:
            stop_countdown()
            write_entries()
            status_label.config(text="Saved all games.", fg="#3cb371")
            top_status_label.config(text="Saved", fg="#3cb371")
            populate_listbox()
            if sel:
                idx = sel[0]
                listbox.selection_clear(0, "end")
                listbox.selection_set(idx)
                listbox.activate(idx)
                listbox.see(idx)
                load_selected_into_ui(idx)
                apply_listbox_styles()
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file.\n{e}")

    def add_entry():
        new = {"run": "", "name": "", "admin": False, "filters": []}
        entries.append(new)
        try:
            write_entries()
            populate_listbox()
            listbox.selection_clear(0, "end")
            new_idx = len(entries) - 1
            listbox.selection_set(new_idx)
            listbox.activate(new_idx)
            listbox.see(new_idx)
            load_selected_into_ui(new_idx)
            apply_listbox_styles()
            status_label.config(text="Added new game.", fg="#3cb371")
        except Exception as e:
            messagebox.showerror("Error", f"Could not add game.\n{e}")

    def delete_entry():
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("No selection", "Select a game first.")
            return
        idx = sel[0]
        if not messagebox.askyesno("Delete", "Delete selected game?"):
            return
        try:
            del entries[idx]
            write_entries()
            populate_listbox()
            if entries:
                new_idx = min(idx, len(entries) - 1)
                listbox.selection_set(new_idx)
                listbox.activate(new_idx)
                listbox.see(new_idx)
                load_selected_into_ui(new_idx)
                apply_listbox_styles()
            else:
                run_var.set("")
                name_var.set("")
                is_admin_var.set(False)
                render_filter_tags([])
            status_label.config(text="Deleted selected game.", fg="#3cb371")
        except Exception as e:
            messagebox.showerror("Error", f"Could not delete game.\n{e}")

    def move_selected(offset):
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("No selection", "Select a game first.")
            return
        idx = sel[0]
        new_idx = idx + offset
        if new_idx < 0 or new_idx >= len(entries):
            return
        sync_current_entry_from_ui(idx)
        entries[idx], entries[new_idx] = entries[new_idx], entries[idx]
        try:
            write_entries()
            populate_listbox()
            listbox.selection_clear(0, "end")
            listbox.selection_set(new_idx)
            listbox.activate(new_idx)
            listbox.see(new_idx)
            load_selected_into_ui(new_idx)
            apply_listbox_styles()
            status_label.config(text="Reordered selected game.", fg="#3cb371")
        except Exception as e:
            messagebox.showerror("Error", f"Could not reorder game.\n{e}")

    def save_editor_text():
        text = text_widget.get("1.0", "end-1c")
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".json":
            try:
                parsed = json.loads(text)
            except Exception as ex:
                if not messagebox.askyesno("Invalid JSON", f"JSON parse error:\n{ex}\n\nSave raw text anyway?"):
                    return
                try:
                    with open(file_path, "w", encoding="utf-8") as f:
                        f.write(text)
                    status_label.config(text="Saved raw text.", fg="#3cb371")
                    load_from_file_callback(file_path)
                    return
                except Exception as e:
                    messagebox.showerror("Error", f"Could not save file.\n{e}")
                    return
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    json.dump(parsed, f, indent=4)
                status_label.config(text="Saved JSON from editor.", fg="#3cb371")
                load_from_file_callback(file_path)
                if isinstance(parsed, list):
                    entries[:] = [normalize_entry(e) for e in parsed]
                populate_listbox()
                refresh_editor_from_entries()
            except Exception as e:
                messagebox.showerror("Error", f"Could not save file.\n{e}")
        else:
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(text)
                status_label.config(text="Saved file.", fg="#3cb371")
                load_from_file_callback(file_path)
            except Exception as e:
                messagebox.showerror("Error", f"Could not save file.\n{e}")

    def perform_auto_save():
        sel = listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        try:
            sync_current_entry_from_ui(idx)
            write_entries()
            populate_listbox()
            listbox.selection_clear(0, "end")
            listbox.selection_set(idx)
            listbox.activate(idx)
            listbox.see(idx)
            load_selected_into_ui(idx)
            apply_listbox_styles()
            status_label.config(text="Auto saved selected game.", fg="#3cb371")
        except:
            pass

    def stop_countdown():
        nonlocal countdown_job
        if countdown_job:
            try:
                popup.after_cancel(countdown_job)
            except:
                pass
            countdown_job = None
        countdown_label.config(text="", fg=fg_color)
        top_status_label.config(text="", fg=fg_color)

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
        top_status_label.config(text="AutoSave pending", fg="#FFD700")
        countdown_job = popup.after(1000, countdown_tick)

    def mark_modified(*args):
        if loading_selection:
            return
        top_status_label.config(text="Modified", fg="#ff7f7f")
        if auto_save.get():
            start_countdown()
        else:
            stop_countdown()

    def toggle_auto_save():
        auto_save.set(not auto_save.get())
        auto_save_btn.config(text=f"AutoSave {'On' if auto_save.get() else 'Off'}", bg="#3cb371" if auto_save.get() else "#777")
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

    def create_drag_preview(index):
        destroy_drag_preview()
        text = listbox.get(index) if 0 <= index < listbox.size() else ""
        preview = tk.Toplevel(popup)
        preview.overrideredirect(True)
        preview.attributes("-topmost", True)
        try:
            preview.attributes("-alpha", 0.92)
        except:
            pass
        frame = tk.Frame(preview, bg="#1f1f1f", bd=2, relief="solid")
        frame.pack()
        label = tk.Label(
            frame,
            text=text,
            bg="#1f1f1f",
            fg="white",
            font=("Consolas", 11, "bold"),
            padx=10,
            pady=6
        )
        label.pack()
        hint = tk.Label(
            frame,
            text="Release to place",
            bg="#1f1f1f",
            fg="#cfcfcf",
            font=("Consolas", 8),
            padx=10,
            pady=4
        )
        hint.pack()
        drag_state["preview"] = preview
        move_drag_preview(drag_state["last_x_root"], drag_state["last_y_root"])

    def move_drag_preview(x_root, y_root):
        preview = drag_state.get("preview")
        if preview is None:
            return
        try:
            preview.geometry(f"+{x_root + 18}+{y_root + 18}")
        except:
            pass

    def destroy_drag_preview():
        preview = drag_state.get("preview")
        if preview is not None:
            try:
                preview.destroy()
            except:
                pass
        drag_state["preview"] = None

    def sync_before_drag():
        sel = listbox.curselection()
        if sel:
            try:
                sync_current_entry_from_ui(sel[0])
            except:
                pass

    def move_entry(from_idx, to_idx):
        if from_idx == to_idx:
            return to_idx
        item = entries.pop(from_idx)
        entries.insert(to_idx, item)
        drag_state["drag_index"] = to_idx
        populate_listbox()
        listbox.selection_clear(0, "end")
        listbox.selection_set(to_idx)
        listbox.activate(to_idx)
        listbox.see(to_idx)
        apply_listbox_styles()
        mark_modified()
        return to_idx

    def start_drag(index):
        if drag_state["dragging"]:
            return
        if drag_state["press_index"] != index:
            return
        if not (0 <= index < len(entries)):
            return
        sync_before_drag()
        drag_state["dragging"] = True
        drag_state["armed"] = False
        drag_state["moved"] = False
        drag_state["drag_index"] = index
        create_drag_preview(index)
        apply_listbox_styles()

    def arm_drag(event):
        idx = listbox.nearest(event.y)
        if idx < 0 or idx >= len(entries):
            return
        drag_state["press_index"] = idx
        drag_state["drag_index"] = idx
        drag_state["armed"] = True
        drag_state["dragging"] = False
        drag_state["moved"] = False
        drag_state["last_x_root"] = event.x_root
        drag_state["last_y_root"] = event.y_root
        if drag_state["after_id"] is not None:
            try:
                popup.after_cancel(drag_state["after_id"])
            except:
                pass
        drag_state["after_id"] = popup.after(450, lambda i=idx: start_drag(i))
        listbox.selection_clear(0, "end")
        listbox.selection_set(idx)
        listbox.activate(idx)
        listbox.see(idx)
        apply_listbox_styles()

    def drag_motion(event):
        drag_state["last_x_root"] = event.x_root
        drag_state["last_y_root"] = event.y_root
        if drag_state["dragging"]:
            move_drag_preview(event.x_root, event.y_root)
            target = listbox.nearest(event.y)
            if target < 0 or target >= len(entries):
                return
            current_idx = drag_state["drag_index"]
            if current_idx is None or not (0 <= current_idx < len(entries)):
                current_idx = drag_state["press_index"]
            if current_idx is None:
                return
            if target != current_idx:
                drag_state["moved"] = True
                current_idx = move_entry(current_idx, target)
                drag_state["press_index"] = current_idx
                drag_state["drag_index"] = current_idx

    def finish_drag(event):
        if drag_state["after_id"] is not None:
            try:
                popup.after_cancel(drag_state["after_id"])
            except:
                pass
            drag_state["after_id"] = None
        if drag_state["dragging"]:
            destroy_drag_preview()
            try:
                write_entries()
            except Exception as e:
                messagebox.showerror("Error", f"Could not save reordered list.\n{e}")
            finally:
                populate_listbox()
                if listbox.size() > 0:
                    idx = drag_state["drag_index"]
                    if idx is None:
                        idx = listbox.curselection()
                        idx = idx[0] if idx else 0
                    idx = min(max(0, idx), listbox.size() - 1)
                    listbox.selection_clear(0, "end")
                    listbox.selection_set(idx)
                    listbox.activate(idx)
                    listbox.see(idx)
                    load_selected_into_ui(idx)
                    apply_listbox_styles()
            drag_state["dragging"] = False
            drag_state["armed"] = False
            drag_state["press_index"] = None
            drag_state["drag_index"] = None
            drag_state["moved"] = False
        else:
            drag_state["armed"] = False
            drag_state["press_index"] = None
            drag_state["drag_index"] = None
            drag_state["moved"] = False

    add_btn = tk.Button(action_frame, text="Add Game", width=14, bg="#666", fg="white", command=add_entry)
    add_btn.pack(side="left", padx=(0, 6))

    del_btn = tk.Button(action_frame, text="Delete Selected Game", width=20, bg="#5B0000", fg="white", command=delete_entry)
    del_btn.pack(side="left", padx=(0, 6))

    save_sel_btn = tk.Button(action_frame, text="Save Selected Game", width=18, bg="#3cb371", fg="white", command=save_selected)
    save_sel_btn.pack(side="left", padx=(0, 6))

    save_all_btn = tk.Button(bottom_right_frame, text="Save All", width=10, bg="#666", fg="white", command=save_all)
    save_all_btn.pack(side="right", padx=(0, 6))

    auto_save_btn = tk.Button(
        bottom_right_frame,
        text=f"AutoSave {'On' if auto_save.get() else 'Off'}",
        width=14,
        bg="#3cb371" if auto_save.get() else "#777",
        fg="white",
        command=toggle_auto_save
    )
    auto_save_btn.pack(side="right", padx=(0, 6))

    up_btn = tk.Button(
        reorder_frame,
        text="Move ↑",
        width=9,
        height=1,
        bg="#666",
        fg="white",
        font=("Segoe UI", 9, "bold"),
        command=lambda: move_selected(-1),
        relief="raised",
        bd=2
    )
    up_btn.pack(side="top", pady=(0, 4))
    Tooltip(up_btn, "Move up")

    down_btn = tk.Button(
        reorder_frame,
        text="Move ↓",
        width=9,
        height=1,
        bg="#666",
        fg="white",
        font=("Segoe UI", 9, "bold"),
        command=lambda: move_selected(1),
        relief="raised",
        bd=2
    )
    down_btn.pack(side="top")
    Tooltip(down_btn, "Move down")

    run_var.trace_add("write", mark_modified)
    name_var.trace_add("write", mark_modified)
    is_admin_var.trace_add("write", mark_modified)

    listbox.bind("<<ListboxSelect>>", on_select)
    listbox.bind("<ButtonPress-1>", arm_drag)
    listbox.bind("<B1-Motion>", drag_motion)
    listbox.bind("<ButtonRelease-1>", finish_drag)

    text_widget.insert("1.0", content)
    text_widget.edit_modified(False)

    def on_change(event=None):
        if updating_editor:
            return
        if text_widget.edit_modified():
            update_line_numbers()
            top_status_label.config(text="Modified", fg="#ff7f7f")
            if auto_save.get():
                start_countdown()
            else:
                stop_countdown()
            text_widget.edit_modified(False)

    text_widget.bind("<<Modified>>", on_change)

    update_line_numbers()
    populate_listbox()

    if entries:
        listbox.selection_set(0)
        listbox.activate(0)
        listbox.see(0)
        load_selected_into_ui(0)
        apply_listbox_styles()
    else:
        render_filter_tags([])

    def close_popup():
        stop_countdown()
        destroy_drag_preview()
        popup.destroy()

    popup.protocol("WM_DELETE_WINDOW", close_popup)
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
        except:
            screen_w = popup.winfo_screenwidth()
            screen_h = popup.winfo_screenheight()
            x = (screen_w // 2) - (pw // 2)
            y = (screen_h // 2) - (ph // 2)
        popup.geometry(f"+{x}+{y}")

    center_popup()
    return popup