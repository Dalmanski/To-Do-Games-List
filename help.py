import tkinter as tk

def open_help_popup(parent=None):
    help_window = tk.Toplevel(parent)
    help_window.title("About")
    help_window.geometry("600x350")
    help_window.configure(bg="#121212")

    container = tk.Frame(help_window, bg="#1e1e1e", bd=2, relief="flat")
    container.place(relx=0.5, rely=0.5, anchor="center", width=550, height=320)

    canvas = tk.Canvas(container, bg="#1e1e1e", highlightthickness=0)
    scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas, bg="#1e1e1e")

    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    title = tk.Label(
        scroll_frame,
        text="To-Do Games List",
        font=("Segoe UI", 16, "bold"),
        bg="#1e1e1e",
        fg="#00d5ff",
        pady=15
    )
    title.pack(anchor="center")

    about_text = (
        "This project serves as an automatic game and app launcher.\n"
        "It opens games or applications one by one based on your\n"
        "list, so you don’t need to launch them manually or search\n"
        "for them. It also helps you decide what to play next.\n"
        "\n"
        "Update (Not final. I will revise this later):\n"
        "Date: July 25, 2025\n"
        "• Reduced code size\n"
        "• Added Help and Settings buttons\n"
        "\n"
        "Date: July 28, 2025\n"
        "• Adjusted text size for Left, Right, and Play buttons\n"
        "• Moved this window up by 40 pixels (Y-axis)\n"
        "• Added ZIP files so you know where it should run. Just extract it, run the .exe, and you're good to go\n"
        "\n"
        "Date: July 29, 2025\n"
        "• Added auto search app on adding game\n"
        "• Add the label \"Pls add new game\" when the list is empty\n"
        "• Added scrollable on help\n"
        "• Added update title label list when created new list\n"
        "\n"
        "Date: July 30, 2025\n"
        "• Added auto play and save on settings.json\n"
        "• Fixed not updating txt on edit list after it create new list\n"
        "• When create new list and load txt, it will start open file on it's default file location\n"
        "\n"
        "Date: September 23, 2025\n"
        "• Added .url files (Internet Shortcut) when \"Browse on .exe\"\n"
        "• Enhance looks sharpness on file selection and in this software app\n"
        "• Use different way to save and auto-save on \"Edit List\" like countdown to 5 sec instead of check each file location to save\n"
        "• Due to Internet Shortcut on Steam file location, it can now read the Game Name from Steam App ID\n"
        "\n"
        "Created by Jayrald John C. Dalman."
    )

    body = tk.Label(
        scroll_frame,
        text=about_text,
        font=("Segoe UI", 10),
        bg="#1e1e1e",
        fg="#f0f0f0",
        justify="left",
        wraplength=460,
        padx=15,
        pady=5
    )
    body.pack(anchor="center")

    help_window.mainloop()

if __name__ == "__main__":
    open_help_popup()
