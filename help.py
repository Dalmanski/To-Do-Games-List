import tkinter as tk

def open_help_popup(parent=None):
    help_window = tk.Toplevel(parent)
    help_window.title("About")
    help_window.geometry("550x350")
    help_window.configure(bg="#121212")

    container = tk.Frame(help_window, bg="#1e1e1e", bd=2, relief="flat")
    container.place(relx=0.5, rely=0.5, anchor="center", width=500, height=300)

    title = tk.Label(
        container,
        text="To-Do Games List",
        font=("Segoe UI", 16, "bold"),
        bg="#1e1e1e",
        fg="#00d5ff",
        pady=10
    )
    title.pack()

    about_text = (
        "This project serves as an automatic game and app launcher.\n"
        "It opens games or applications one by one based on your list,\n"
        "so you don't need to launch them manually by finding them.\n"
        "It also helps you decide what to play next.\n\n"
        "Update: (Not final. I will recap before this)\n"
        "Date: July 25, 2025\n"
        "• Reduced code\n"
        "• Added help and settings buttons\n\n"
        "Created by Jayrald John C. Dalman."
    )

    body = tk.Label(
        container,
        text=about_text,
        font=("Segoe UI", 10),
        bg="#1e1e1e",
        fg="#f0f0f0",
        justify="left",
        wraplength=460,
        padx=15,
        pady=5
    )
    body.pack()

    help_window.mainloop()

if __name__ == "__main__":
    open_help_popup()
