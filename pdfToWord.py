import os
import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
from tkinter import ttk

def browse_pdf_files():
    global pdf_files
    new_pdf_files = list(filedialog.askopenfilenames(filetypes=[("PDF Documents", "*.pdf")]))
    pdf_files.extend(new_pdf_files)
    for file in new_pdf_files:
        selected_files_listbox.insert(tk.END, os.path.basename(file))
    print("Selected PDF files:", pdf_files)

def browse_output_folder():
    global output_folder
    output_folder = filedialog.askdirectory()
    print("Selected output folder:", output_folder)

def convert_to_word():
    if pdf_files and output_folder:
        for file in pdf_files:
            output_file = os.path.join(output_folder, os.path.splitext(os.path.basename(file))[0] + ".docx")
            cv = Converter(file)
            cv.convert(output_file, start=0, end=None)
            cv.close()
            print("Word file saved at:", output_file)
        messagebox.showinfo("Conversion Complete", "The selected PDF documents have been converted to Word files.")
    else:
        print("Please select PDF files and output folder first.")

# Create the main window
window = tk.Tk()
window.title("PDF to Word Converter")
window.geometry("600x400")
window.configure(bg="#FFF8DC")

# Add buttons and labels
select_files_btn = tk.Button(window, text="Select PDF Documents", command=browse_pdf_files, bd=2)
select_files_btn.place(relx=0.5, rely=0.2, anchor="center")

select_output_btn = tk.Button(window, text="Output Folder", command=browse_output_folder, bd=2)
select_output_btn.place(relx=0.5, rely=0.35, anchor="center")

selected_files_listbox = tk.Listbox(window, width=50, height=5, selectmode=tk.MULTIPLE, exportselection=False)
selected_files_listbox.place(relx=0.5, rely=0.55, anchor="center")

convert_btn = tk.Button(window, text="Convert", command=convert_to_word, bd=2)
convert_btn.place(relx=0.5, rely=0.75, anchor="center")

# Initialize variables
pdf_files = []
output_folder = None

# Start the main loop
window.mainloop()