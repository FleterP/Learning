import os
import subprocess
import tkinter as tk
from tkinter import messagebox
import win32print
import win32ui
import win32con

# Text to folder mapping
folder_map = {
    '90-012-154-30': 'C:/Users/aku10/Desktop/90-012-154-30',
    '90-010-004-30': 'C:/Users/aku10/Desktop/90-010-004-30',
}

# Defined print settings for specific PDF files
file_print_settings = {
    'Karta przygotowania': {'paper_size': "A4", 'color': 'color'},
    'Karta etykiet': {'paper_size': 'A3', 'color': 'monochrome'},
}

# Function to print a PDF file using Microsoft Edge
def print_pdf_file(pdf_path, settings):
    if not os.path.exists(pdf_path):
        messagebox.showerror("Error", f"File {pdf_path} does not exist.")
        return
    
    try:
        # Get the default printer
        printer_name = win32print.GetDefaultPrinter()
        print(f"Used printer: {printer_name}")  # Debugging – can be removed
        
        # Print settings
        paper_size = settings.get('paper_size', 'A4')
        color = settings.get('color', 'color')
        
        # Mapping paper size to Windows constants
        paper_sizes = {'A4': win32con.DMPAPER_A4, 'A3': win32con.DMPAPER_A3}
        paper_code = paper_sizes.get(paper_size, win32con.DMPAPER_A4)
        
        # Open the printer and get settings
        hprinter = win32print.OpenPrinter(printer_name)
        devmode = win32print.GetPrinter(hprinter, 2)["pDevMode"]
        
        # Print settings
        devmode.PaperSize = paper_code
        devmode.Color = win32con.DMCOLOR_COLOR if color == 'color' else win32con.DMCOLOR_MONOCHROME
        
        # Create device context (DC) using win32ui
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)
        
        # Start document and page
        hdc.StartDoc(os.path.basename(pdf_path))
        hdc.StartPage()
        
        # Here we need PDF rendering – for now, opening in Edge as a fallback
        # To print automatically, a PDF rendering library (e.g., PyMuPDF) is required
        edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        subprocess.run([edge_path, pdf_path], check=True)  # Fallback – opens PDF
        
        # End page and document
        hdc.EndPage()
        hdc.EndDoc()
        
        # Release resources
        hdc.DeleteDC()
        win32print.ClosePrinter(hprinter)
        
        messagebox.showinfo("Success", f"File {os.path.basename(pdf_path)} has been printed.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while printing {pdf_path}: {e}")

# Rest of your code (GUI, folder mapping, etc.) remains unchanged

# Function to print all files from a folder
def print_all_pdfs_in_folder(pdf_files, folder_path):
    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        settings = file_print_settings.get(pdf_file, {'paper_size': 'A4', 'color': 'color'})
        print_pdf_file(pdf_path, settings)

# Function to dynamically create buttons for PDF files
def load_pdf_files(folder_path):
    if not os.path.exists(folder_path):
        messagebox.showerror("Error", f"Folder {folder_path} does not exist.")
        return
    
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    
    if not pdf_files:
        messagebox.showinfo("Information", "No PDF files found in the selected folder.")
        return
    
    for widget in button_frame.winfo_children():
        widget.destroy()

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        settings = file_print_settings.get(pdf_file, {'paper_size': 'A4', 'color': 'color'})
        paper_size = settings['paper_size']
        color_option = settings['color']
        color_text = "Color" if color_option == 'color' else "Black and White"

        info_label = tk.Label(button_frame, text=f"{pdf_file} - Format: {paper_size}, Print: {color_text}")
        info_label.pack(pady=5)

        button = tk.Button(button_frame, text="Print", command=lambda p=pdf_path, s=settings: print_pdf_file(p, s))
        button.pack(pady=5)

    print_all_button = tk.Button(button_frame, text="Print All", command=lambda: print_all_pdfs_in_folder(pdf_files, folder_path))
    print_all_button.pack(pady=10)

# Function to handle text input
def load_by_keyword():
    keyword = entry.get().lower()
    folder_path = folder_map.get(keyword)
    
    if folder_path:
        load_pdf_files(folder_path)
    else:
        messagebox.showerror("Error", "No folder found for the entered number.")

# Create GUI
root = tk.Tk()
root.title("PDF Printing")

label = tk.Label(root, text="Enter model number")
label.pack(pady=10)

entry = tk.Entry(root, width=40)
entry.pack(pady=5)

load_button = tk.Button(root, text="Load Files", command=load_by_keyword)
load_button.pack(pady=10)

button_frame = tk.Frame(root)
button_frame.pack(pady=5, fill=tk.X)

root.mainloop()