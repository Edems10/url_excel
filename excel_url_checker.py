import tkinter as tk
from tkinter import filedialog, ttk
import os
import random
import threading
from excel_modify import process_excel  
import tkinter.messagebox as messagebox

CORRECT = 'leave as is'
INCORRECT = ''
ACCESS_FORBIDDEN = 'access forbidden'
CHECK_PDF = True
LOCALIZATION_TYPES = ['en-us', 'en-gb', 'en-in', 'en-ca', 'en-au']

class Tooltip:
    def __init__(self, widget, text="Information"):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.showing = False
        self.widget.bind("<Enter>", self.on_enter)
        self.widget.bind("<Leave>", self.on_leave)

    def on_enter(self, event=None):
        self.showing = True
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25

        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")

        label = tk.Label(self.tooltip, text=self.text, background="#FFFFDD", relief="solid", borderwidth=1)
        label.pack()

    def on_leave(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.showing = False

def browse_file():
    global FILE_PATH
    FILE_PATH = filedialog.askopenfilename(filetypes=[("All Files", "*.*"), ("Excel Files", "*.xlsx;*.xls")])
    update_file_status()

def clear_log_text():
    log_text.delete('1.0', tk.END)

def process_data(progress_bar):
    global FILE_PATH
    if FILE_PATH:
        clear_log_text()
        loading_window, progress_bar = show_loading_window()  # Don't need progress_bar here
        # Define a function to be run in a separate thread
        def process_data_thread():
            try:
                result = process_excel(FILE_PATH, CORRECT, INCORRECT, ACCESS_FORBIDDEN, CHECK_PDF, LOCALIZATION_TYPES, progress_bar)
                # Use tkinter's thread-safe method to update GUI
                root.after(0, lambda: log_text.insert(tk.END, f"Excel URLs outcome:\n"
                                                            f"{result[0]} are working correctly\n"
                                                            f"{result[1]} are either not accessible or not working\n"
                                                            "File was updated\n"))
            except Exception as e:
                messagebox.showerror("Error", str(e))
            finally:
                # Close loading window after processing
                root.after(0, close_loading_window, loading_window)
        
        # Start a new thread for processing data
        threading.Thread(target=process_data_thread).start()
    else:
        messagebox.showinfo("File not Selected", "Please select a file first.")



def show_loading_window():
    loading_window = tk.Toplevel()
    loading_window.title("Loading...")
    loading_window.geometry("300x100")

    loading_label = tk.Label(loading_window, text="Please wait while loading...")
    loading_label.pack(pady=5)

    progress_bar = ttk.Progressbar(loading_window, orient="horizontal", length=200, mode="determinate",maximum=100)
    progress_bar.pack(pady=5)

    return loading_window, progress_bar


def close_loading_window(loading_window):
    loading_window.destroy()

def display_popup():
    random_number = random.randint(1, 15)
    if random_number == 1:
        messagebox.showinfo("IMPORTANT", "Danielka is awesome <3!")
    elif random_number == 2:
        messagebox.showinfo("IMPORTANT", "If you think about it Danielka is Amazing")
    elif random_number == 3:
        messagebox.showinfo("IMPORTANT", "QTITO<3")
    elif random_number == 4:
        messagebox.showinfo("IMPORTANT", "!Bestito!")

def update_values(event):
    global CORRECT, INCORRECT, ACCESS_FORBIDDEN, CHECK_PDF
    CORRECT = correct_entry.get()
    INCORRECT = incorrect_entry.get()
    ACCESS_FORBIDDEN = forbidden_entry.get()
    CHECK_PDF = pdf_check_var.get()

def update_localization_types(event=None):
    global LOCALIZATION_TYPES
    input_text = localization_entry.get().strip()
    if input_text:
        LOCALIZATION_TYPES = [item.strip() for item in input_text.split(',')]
    else:
        LOCALIZATION_TYPES = []

def update_file_status():
    global FILE_PATH
    if FILE_PATH:
        if os.path.isfile(FILE_PATH):
            file_name = os.path.basename(FILE_PATH)
            file_status_label.config(text=f"{file_name}", fg="black")
        else:
            file_status_label.config(text=f"Not a file - {FILE_PATH}", fg="red")  # Red color
    else:
        file_status_label.config(text="No file selected", fg="red")  # Red color

def main():
    progress_bar = None  # Define progress_bar variable here
    global file_status_label
    global incorrect_entry
    global correct_entry
    global forbidden_entry
    global localization_entry
    global pdf_check_var
    global log_text
    global root  # Add this line to declare root as a global variable
    root = tk.Tk()  # Define root as the Tkinter root window
    root.title("Excel URL checker")
    root.geometry("600x500")  # Adjusted width to give more space

    file_status_label = tk.Label(root, text="No file selected", fg="red")
    file_status_label.pack(padx=20, pady=10)

    button_frame = tk.Frame(root, pady=20)
    button_frame.pack()

    browse_button = tk.Button(button_frame, text="Select File", command=browse_file)
    browse_button.pack(side=tk.LEFT, padx=10)

    process_button = tk.Button(button_frame, text="Check URLs", command=lambda pb=progress_bar: process_data(pb))
    process_button.pack(side=tk.LEFT, padx=10)

    settings_frame = tk.Frame(root, padx=20, pady=10)
    settings_frame.pack(fill=tk.X, expand=True)

    tk.Label(settings_frame, text="Correct:").grid(row=0, column=0, sticky=tk.W)
    global correct_entry
    correct_entry = tk.Entry(settings_frame)
    correct_entry.grid(row=0, column=1, sticky=tk.EW)
    correct_entry.insert(0, 'leave as is')

    tk.Label(settings_frame, text="Incorrect:").grid(row=1, column=0, sticky=tk.W)
    global incorrect_entry
    incorrect_entry = tk.Entry(settings_frame)
    incorrect_entry.grid(row=1, column=1, sticky=tk.EW)
    incorrect_entry.insert(0, '')

    tk.Label(settings_frame, text="Forbidden:").grid(row=2, column=0, sticky=tk.W)
    global forbidden_entry
    forbidden_entry = tk.Entry(settings_frame)
    forbidden_entry.grid(row=2, column=1, sticky=tk.EW)
    forbidden_entry.insert(0, 'access forbidden')

    tk.Label(settings_frame, text="Localization Types:").grid(row=3, column=0, sticky=tk.W)
    global localization_entry
    localization_entry = tk.Entry(settings_frame)
    localization_entry.grid(row=3, column=1, sticky=tk.EW)
    localization_entry.insert(0, ','.join(LOCALIZATION_TYPES))
    localization_entry.bind("<KeyRelease>", update_localization_types)

    global pdf_check_var
    pdf_check_var = tk.BooleanVar()
    pdf_check_var.set(True)
    pdf_check = tk.Checkbutton(settings_frame, text="Check PDF Files", variable=pdf_check_var)
    pdf_check.grid(row=4, column=0, columnspan=2, sticky=tk.W)

    global log_text
    log_text = tk.Text(root, height=10, width=60, wrap=tk.WORD)
    log_text.pack(padx=20, pady=(0, 20))

    correct_entry.bind("<KeyRelease>", update_values)
    incorrect_entry.bind("<KeyRelease>", update_values)
    forbidden_entry.bind("<KeyRelease>", update_values)

    settings_frame.columnconfigure(1, weight=1)  # This makes the second column expandable

    # Tooltip messages for settings
    tooltips = {
        correct_entry: "This will be written if the URL is working correctly",
        incorrect_entry: "This will be written if the URL is not working correctly",
        forbidden_entry: "This will be written if the URL is not accessible due to an authentication issue",
        localization_entry: "Enter comma-separated localization types (e.g., en-us, en-gb)",
    }

    for widget, tooltip_text in tooltips.items():
        Tooltip(widget, tooltip_text)

    display_popup()

    root.mainloop()

if __name__ == '__main__':
    main()
