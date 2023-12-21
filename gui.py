import tkinter as tk
from tkinter import filedialog
from excel_modify import process_excel  # Import your processing logic

def browse_file():
    global FILE_PATH
    FILE_PATH = filedialog.askopenfilename()
    log_text.insert(tk.END, f"Selected File: {FILE_PATH}\n")

def clear_log_text():
    log_text.delete('1.0', tk.END)

def process_data():
    global FILE_PATH
    if FILE_PATH:
        clear_log_text()
        # Process logic here using the global variables or the values retrieved from the Entry widgets
        result = process_excel(FILE_PATH,CORRECT,INCORRECT,ACCESS_FORBIDDEN,CHECK_PDF)
        log_text.insert(tk.END, f"Excel URLs outcome:\n"
                                f"{result[0]} are working correctly\n"
                                f"{result[1]} are either not accessible or not working\n"
                                "File was updated\n")
    else:
        log_text.insert(tk.END, "Please select a file first.\n")

def update_values():
    global CORRECT, INCORRECT, ACCESS_FORBIDDEN, CHECK_PDF
    CORRECT = correct_entry.get()
    INCORRECT = incorrect_entry.get()
    ACCESS_FORBIDDEN = forbidden_entry.get()
    CHECK_PDF = pdf_check_var.get()

root = tk.Tk()
root.title("Local Application")
root.geometry("600x400")

button_frame = tk.Frame(root, pady=20)
button_frame.pack()

browse_button = tk.Button(button_frame, text="Browse", command=browse_file)
browse_button.pack(side=tk.LEFT, padx=10)

process_button = tk.Button(button_frame, text="Process", command=process_data)
process_button.pack(side=tk.RIGHT, padx=10)

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
pdf_check = tk.Checkbutton(settings_frame, text="Check PDF", variable=pdf_check_var)
pdf_check.grid(row=4, columnspan=2)

update_button = tk.Button(settings_frame, text="Update Values", command=update_values)
update_button.grid(row=5, columnspan=2, pady=10)

log_text = tk.Text(root, height=10, width=60, wrap=tk.WORD)
log_text.pack(padx=20, pady=(0, 20))

root.mainloop()
