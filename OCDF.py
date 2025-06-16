import tkinter as tk
import win32com.client

def fetch_current_word_page_info():
    try:
        word_app = win32com.client.GetActiveObject("Word.Application")
        active_doc = word_app.ActiveDocument
        current_selection = word_app.Selection
        current_selection.MoveRight(Unit=1, Count=0)
        current_page = current_selection.Information(3)
        total_page_count = active_doc.ComputeStatistics(2)
        return f"{current_page}/{total_page_count}"
    except:
        return "Word not active"

def refresh_page_label():
    page_text = fetch_current_word_page_info()
    page_label.config(text=page_text)
    window.after(500, refresh_page_label)

def enable_window_dragging(frame):
    def on_start_drag(event):
        frame.drag_start_x = event.x
        frame.drag_start_y = event.y
    def while_dragging(event):
        new_x = frame.winfo_x() - frame.drag_start_x + event.x
        new_y = frame.winfo_y() - frame.drag_start_y + event.y
        frame.geometry(f"+{new_x}+{new_y}")
    frame.bind("<Button-1>", on_start_drag)
    frame.bind("<B1-Motion>", while_dragging)

window = tk.Tk()
window.overrideredirect(True)
window.attributes("-topmost", True)
window.configure(bg="black")
window.geometry("+50+50")

page_label = tk.Label(
    window,
    text="Loading...",
    font=("Consolas", 18, "bold"),
    fg="white",
    bg="black",
    padx=10,
    pady=5
)
page_label.pack()

enable_window_dragging(window)
refresh_page_label()
window.mainloop()
