import tkinter as tk
from tkinter import filedialog
import os
from excel_modify import process_excel  
import random
import tkinter.messagebox as messagebox

CORRECT = 'leave as is'
INCORRECT = ''
LOCALIZATION = 'remove en-us'
ACCESS_FORBIDDEN = 'access forbidden'
COMMENTS = 'DUB comment'
URL = 'link'
SLIDE = 'slide'
FILE_PATH = ""
CHECK_PDF = True


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
    FILE_PATH = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls"),("All Files","*.*")])
    update_file_status()

def clear_log_text():
    log_text.delete('1.0', tk.END)

def process_data():
    global FILE_PATH
    if FILE_PATH:
        clear_log_text()
        result = process_excel(FILE_PATH, CORRECT, INCORRECT, ACCESS_FORBIDDEN, CHECK_PDF)
        log_text.insert(tk.END, f"Excel URLs outcome:\n"
                                f"{result[0]} are working correctly\n"
                                f"{result[1]} are either not accessible or not working\n"
                                "File was updated\n")
    else:
        log_text.insert(tk.END, "Please select a file first.\n")

def display_popup():
    random_number = random.randint(1, 10)  
    if random_number == 1:  
        messagebox.showinfo("IMPORTANT", "Danielka is awesome <3!")

def update_values(event):
    global CORRECT, INCORRECT, ACCESS_FORBIDDEN, CHECK_PDF
    CORRECT = correct_entry.get()
    INCORRECT = incorrect_entry.get()
    ACCESS_FORBIDDEN = forbidden_entry.get()
    CHECK_PDF = pdf_check_var.get()

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
    global file_status_label
    global incorrect_entry
    global correct_entry
    global forbidden_entry
    global pdf_check_var
    global log_text
    
    root = tk.Tk()
    root.title("Excel URL checker")
    root.geometry("500x400")

    file_status_label = tk.Label(root, text="No file selected", fg="red")
    file_status_label.pack(padx=20, pady=10)

    button_frame = tk.Frame(root, pady=20)
    button_frame.pack()

    browse_button = tk.Button(button_frame, text="Select File", command=browse_file)
    browse_button.pack(side=tk.LEFT, padx=10)

    process_button = tk.Button(button_frame, text="Check URLs", command=process_data)
    process_button.pack(side=tk.LEFT, padx=10)

    settings_frame = tk.Frame(root, padx=20, pady=10)
    settings_frame.pack()

    tk.Label(settings_frame, text="Correct:").grid(row=0, column=0)
    correct_entry = tk.Entry(settings_frame)
    correct_entry.grid(row=0, column=1)
    correct_entry.insert(0, 'leave as is')

    tk.Label(settings_frame, text="Incorrect:").grid(row=1, column=0)
    incorrect_entry = tk.Entry(settings_frame)
    incorrect_entry.grid(row=1, column=1)
    incorrect_entry.insert(0, '')

    tk.Label(settings_frame, text="Forbidden:").grid(row=3, column=0)
    forbidden_entry = tk.Entry(settings_frame)
    forbidden_entry.grid(row=3, column=1)
    forbidden_entry.insert(0, 'access forbidden')

    pdf_check_var = tk.BooleanVar()
    pdf_check_var.set(True)
    pdf_check = tk.Checkbutton(settings_frame, text="Check PDF Files", variable=pdf_check_var)
    pdf_check.grid(row=4, columnspan=2)


    log_text = tk.Text(root, height=10, width=60, wrap=tk.WORD)
    log_text.pack(padx=20, pady=(0, 20))

    correct_entry.bind("<KeyRelease>", update_values)
    incorrect_entry.bind("<KeyRelease>", update_values)
    forbidden_entry.bind("<KeyRelease>", update_values)

    # Tooltip messages for settings
    tooltips = {
        correct_entry: "This will be written if the URL is working correctly",
        incorrect_entry: "This will be written if the URL is not working correctly",
        forbidden_entry: "This will be written if the URL is not accesible due to an authentication issue",
    }

    for widget, tooltip_text in tooltips.items():
        Tooltip(widget, tooltip_text)

    display_popup()

    root.mainloop()

if __name__ == '__main__':
    main()

