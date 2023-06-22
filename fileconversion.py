import os
from tkinter import Tk, Label, Entry, Button, OptionMenu, StringVar, filedialog
import win32com.client

from pdf2docx import Converter
from fpdf import FPDF
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from docx import Document
from PIL import Image
from pdf2image import convert_from_path


def pdf_to_docx(pdf_path, output_directory, file_name):
    docx_path = os.path.join(output_directory, f"{file_name}.docx")
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None, mode='w')
    cv.close()
    print("PDF to DOCX conversion completed.")


def docx_to_pdf(docx_path, output_directory, file_name):
    pdf_path = os.path.join(output_directory, f"{file_name}.pdf")

    # Load the DOCX file
    doc = Document(docx_path)

    # Save the DOCX as a temporary file
    temp_path = os.path.join(output_directory, "temp.docx")
    doc.save(temp_path)

    # Convert DOCX to PDF using Microsoft Word
    word_app = win32com.client.Dispatch("Word.Application")
    word_doc = word_app.Documents.Open(os.path.abspath(temp_path))
    word_doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
    word_doc.Close()
    word_app.Quit()

    # Remove the temporary DOCX file
    os.remove(temp_path)

    print("DOCX to PDF conversion completed.")


def image_to_pdf(image_path, output_directory, file_name):
    pdf_path = os.path.join(output_directory, f"{file_name}.pdf")
    image = Image.open(image_path)
    pdf = image.convert("RGB")
    pdf.save(pdf_path, "PDF", resolution=100.0, save_all=True)
    print("Image to PDF conversion completed.")


def pdf_to_image(pdf_path, output_directory):
    images = convert_from_path(pdf_path)

    # Create the output folder if it doesn't exist
    os.makedirs(output_directory, exist_ok=True)

    for i, image in enumerate(images):
        image_path = os.path.join(output_directory, f"page_{i + 1}.jpg")
        image.save(image_path, "JPEG")

    print("PDF to images conversion completed.")


def browse_input_file():
    input_path = filedialog.askopenfilename()
    input_path_entry.delete(0, "end")
    input_path_entry.insert(0, input_path)


def browse_output_directory():
    output_directory = filedialog.askdirectory()
    output_directory_entry.delete(0, "end")
    output_directory_entry.insert(0, output_directory)


def convert_file():
    # Get input values from entry fields
    conversion_type = conversion_type_var.get().lower()
    input_path = input_path_entry.get().strip('"')
    output_directory = output_directory_entry.get()
    file_name = file_name_entry.get()

    # Perform the selected conversion based on user input
    if conversion_type == "pdf2docx":
        try:
            pdf_to_docx(input_path, output_directory, file_name)
            status_label.config(text="Conversion completed.")
        except PermissionError:
            status_label.config(text="Error: Permission denied. Please choose a different output directory.")
    elif conversion_type == "docx2pdf":
        try:
            docx_to_pdf(input_path, output_directory, file_name)
            status_label.config(text="Conversion completed.")
        except PermissionError:
            status_label.config(text="Error: Permission denied. Please choose a different output directory.")
    elif conversion_type == "image2pdf":
        try:
            image_to_pdf(input_path, output_directory, file_name)
            status_label.config(text="Conversion completed.")
        except PermissionError:
            status_label.config(text="Error: Permission denied. Please choose a different output directory.")
    elif conversion_type == "pdf2image":
        try:
            pdf_to_image(input_path, output_directory)
            status_label.config(text="Conversion completed.")
        except PermissionError:
            status_label.config(text="Error: Permission denied. Please choose a different output directory.")
    else:
        status_label.config(text="Invalid conversion type selected.")


# Create the main window
window = Tk()
window.title("File Conversion")
window.geometry("400x400")

# Conversion type label and dropdown button
conversion_type_label = Label(window, text="Conversion Type:")
conversion_type_label.pack()
conversion_type_var = StringVar(window)
conversion_type_var.set("pdf2docx")  # Default conversion type
conversion_type_dropdown = OptionMenu(window, conversion_type_var, "pdf2docx", "docx2pdf", "image2pdf", "pdf2image")
conversion_type_dropdown.pack()

# Input path label, entry field, and browse button
input_path_label = Label(window, text="Input File Path:")
input_path_label.pack()
input_path_entry = Entry(window)
input_path_entry.pack()
browse_input_button = Button(window, text="Browse", command=browse_input_file)
browse_input_button.pack(anchor="center")

# Output directory label, entry field, and browse button
output_directory_label = Label(window, text="Output Directory:")
output_directory_label.pack()
output_directory_entry = Entry(window)
output_directory_entry.pack()
browse_output_button = Button(window, text="Browse", command=browse_output_directory)
browse_output_button.pack(anchor="center")

# Output file name label and entry field
file_name_label = Label(window, text="Output File Name:")
file_name_label.pack()
file_name_entry = Entry(window)
file_name_entry.pack()

# Convert button
convert_button = Button(window, text="Convert", command=convert_file)
convert_button.pack(anchor="center")

# Status label
status_label = Label(window, text="")
status_label.pack()

# Start the main loop
window.mainloop()
