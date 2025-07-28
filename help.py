import tkinter as tk

def open_help_popup(parent=None):
    help_window = tk.Toplevel(parent)
    help_window.title("About")
    help_window.geometry("550x350")
    help_window.configure(bg="#121212")

    container = tk.Frame(help_window, bg="#1e1e1e", bd=2, relief="flat")
    container.place(relx=0.5, rely=0.5, anchor="center", width=500)

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
        "so you don’t need to launch them manually or search for them.\n"
        "It also helps you decide what to play next.\n"
        "\n"
        "Update (Not final. I will revise this later):\n"
        "Date: July 25, 2025\n"
        "• Reduced code size\n"
        "• Added Help and Settings buttons\n"
        "Date: July 28, 2025\n"
        "• Adjusted text size for Left, Right, and Play buttons\n"
        "• Moved this window up by 40 pixels (Y-axis)\n"
        "• Added ZIP files so you know where it should run. Just extract it, run the .exe, and you're good to go.\n"
        "\n"
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
        pady=10
    )
    body.pack()

    help_window.mainloop()

if __name__ == "__main__":
    open_help_popup()
