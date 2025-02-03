import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Button, Label, Entry
import os
from functools import partial

# Import the backend functions
from main import compute

def browse_file(entry_field):
    file_path = filedialog.askopenfilename()
    if file_path:
        entry_field.delete(0, tk.END)
        entry_field.insert(0, file_path)

def browse_folder(entry_field):
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry_field.delete(0, tk.END)
        entry_field.insert(0, folder_path)

def start_processing(company_file, uan_id_file, uan_name_file, input_folder, output_folder):
    company_file = company_file.get()
    uan_id_file = uan_id_file.get()
    uan_name_file = uan_name_file.get()
    input_folder = input_folder.get()
    output_folder = output_folder.get()

    if not (company_file and uan_id_file and uan_name_file and input_folder and output_folder):
        messagebox.showerror("Error", "Please fill in all fields.")
        return

    # Ensure output folder path ends with a slash for consistency
    output_folder = os.path.abspath(output_folder).rstrip("/") + "/"

    try:
        # Pass the corrected output folder to the compute function
        compute(company_file, uan_id_file, uan_name_file, input_folder, output_folder)
        show_save_options(output_folder)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def show_save_options(output_folder):
    save_window = tk.Toplevel()
    save_window.title("Save Options")
    save_window.geometry("400x200")
    Label(save_window, text="Select the format to save:", font=("Arial", 12)).pack(pady=20)

    def save_as_excel():
        messagebox.showinfo("Success", f"Excel file saved in {output_folder}.")
        save_window.destroy()

    def save_as_pdf():
        messagebox.showinfo("Success", f"PDF file saved in {output_folder}.")
        save_window.destroy()

    Button(save_window, text="Save as Excel", command=save_as_excel).pack(pady=10)
    Button(save_window, text="Save as PDF", command=save_as_pdf).pack(pady=10)

# Main Application
root = tk.Tk()
root.title("Excel and PDF Processor")
root.geometry("600x600")
root.resizable(False, False)

# Heading
heading = Label(root, text="Excel and PDF Processor", font=("Arial", 18, "bold"))
heading.pack(pady=30)

# File selection fields
fields = [
    ("Select the Company Data File:", browse_file),
    ("Select the UAN Client File:", browse_file),
    ("Select the Candidate File:", browse_file),
    ("Select the Input Folder:", browse_folder),
    ("Select the Output Folder:", browse_folder)
]

entries = []

for label_text, browse_function in fields:
    frame = tk.Frame(root)
    frame.pack(pady=10, padx=20, fill=tk.X)

    label = Label(frame, text=label_text, anchor="w", width=25)
    label.pack(side=tk.LEFT, padx=5)

    entry = Entry(frame, width=20)
    entry.pack(side=tk.LEFT, padx=5)
    entries.append(entry)

    browse_btn = Button(frame, text="Browse", command=partial(browse_function, entry))
    browse_btn.pack(side=tk.LEFT, padx=5)

# Start Process Button
start_btn = Button(root, text="Start Process", command=lambda: start_processing(*entries), width=20)
start_btn.pack(pady=30)

root.mainloop()
