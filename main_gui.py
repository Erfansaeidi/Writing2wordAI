import os
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from docx import Document
import pytesseract
from PIL import Image
import fitz  # PyMuPDF

DEFAULT_OUTPUT_FILE = "handwrite2docx.docx"

# Function to perform OCR and save the text to a Word document
def process_file(input_path, language, output_path):
    extracted_text = extract_text_from_file(input_path, language)
    
    if not extracted_text.strip():
        messagebox.showerror("Error", "No text was extracted. Please check the file.")
        return
    
    # Save to Word
    try:
        doc = Document()
        doc.add_paragraph(extracted_text)
        doc.save(output_path)

        if os.path.exists(output_path):
            messagebox.showinfo("Success", f"Processing complete!\nFile saved at:\n{output_path}")
        else:
            messagebox.showerror("Error", "Failed to save the file.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while saving: {e}")

# Function to extract text from PDF or image
def extract_text_from_file(file_path, language):
    ext = os.path.splitext(file_path)[-1].lower()

    if ext in [".png", ".jpg", ".jpeg"]:
        return extract_text_from_image(file_path, language)
    elif ext == ".pdf":
        return extract_text_from_pdf(file_path, language)
    else:
        return "Unsupported file format."

# Extract text from image using Tesseract OCR
def extract_text_from_image(image_path, language):
    try:
        img = Image.open(image_path)
        return pytesseract.image_to_string(img, lang=language)
    except Exception as e:
        return f"Error processing image: {e}"

# Extract text from PDF
def extract_text_from_pdf(pdf_path, language):
    try:
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += pytesseract.image_to_string(Image.open(page.get_pixmap().pil_tobytes()), lang=language) + "\n"
        return text
    except Exception as e:
        return f"Error processing PDF: {e}"

# Browse for input file
def browse_input(entry):
    filename = filedialog.askopenfilename(
        title="Select a file",
        filetypes=[("PDF files", "*.pdf"), ("Image files", "*.png;*.jpg;*.jpeg"), ("All files", "*.*")]
    )
    if filename:
        entry.delete(0, ttk.END)
        entry.insert(0, filename)

# Browse for output file
def browse_output(entry):
    filename = filedialog.asksaveasfilename(
        title="Save Word File",
        defaultextension=".docx",
        filetypes=[("Word Document", "*.docx")],
        initialfile=DEFAULT_OUTPUT_FILE
    )
    if filename:
        entry.delete(0, ttk.END)
        entry.insert(0, filename)

# Main GUI
def main_gui():
    global root
    root = ttk.Window(themename="minty")
    root.title("📝 Handwritten to Word Converter")
    root.geometry("700x400")
    root.minsize(500, 350)
    root.resizable(True, True)

    frame = ttk.Frame(root, padding=20)
    frame.pack(fill=BOTH, expand=True)

    # Grid Configuration
    for i in range(3):
        frame.columnconfigure(i, weight=1)
    for i in range(6):
        frame.rowconfigure(i, weight=1)

    # Title
    title_label = ttk.Label(
        frame, text="✨ Handwritten to Word Converter ✨", font=("Comic Sans MS", 18, "bold"), bootstyle="primary"
    )
    title_label.grid(row=0, column=0, columnspan=3, pady=10, sticky="ew")

    # Language selection
    languages = {"English": "eng", "Persian (فارسی)": "fas", "Spanish": "spa", "French": "fra", "German": "deu"}
    ttk.Label(frame, text="🌍 Choose Language:", font=("Arial Rounded MT Bold", 12)).grid(row=1, column=0, sticky="w", pady=5)

    language_var = ttk.StringVar(value="English")

    def clear_selection(event):
        root.after(100, lambda: language_combo.selection_clear())  # Clear selection after UI update

    language_combo = ttk.Combobox(frame, textvariable=language_var, state="readonly", font=("Arial", 10))
    language_combo['values'] = list(languages.keys())
    language_combo.bind("<<ComboboxSelected>>", clear_selection)  # Fix selection issue
    language_combo.grid(row=1, column=1, columnspan=2, sticky="ew", pady=5)

    # Input file selection
    ttk.Label(frame, text="📂 Input File:", font=("Arial Rounded MT Bold", 12)).grid(row=2, column=0, sticky="w", pady=5)

    input_entry = ttk.Entry(frame, font=("Arial", 10))
    input_entry.grid(row=2, column=1, sticky="ew", pady=5)

    input_button = ttk.Button(frame, text="Browse", bootstyle="info-outline", command=lambda: browse_input(input_entry))
    input_button.grid(row=2, column=2, sticky="ew", padx=5, pady=5)

    # Output file selection
    ttk.Label(frame, text="💾 Output File:", font=("Arial Rounded MT Bold", 12)).grid(row=3, column=0, sticky="w", pady=5)

    output_entry = ttk.Entry(frame, font=("Arial", 10))
    output_entry.grid(row=3, column=1, sticky="ew", pady=5)
    output_entry.insert(0, DEFAULT_OUTPUT_FILE)

    output_button = ttk.Button(frame, text="Browse", bootstyle="info-outline", command=lambda: browse_output(output_entry))
    output_button.grid(row=3, column=2, sticky="ew", padx=5, pady=5)

    # Process button
    def on_process():
        lang_display = language_var.get()
        lang_code = languages.get(lang_display, "eng")
        input_file = input_entry.get()
        output_file = output_entry.get()
        
        if not input_file:
            messagebox.showerror("Error", "Please select an input file.")
            return
        if not output_file:
            messagebox.showerror("Error", "Please select an output file.")
            return
        
        process_file(input_file, lang_code, output_file)

    process_button = ttk.Button(frame, text="🚀 Convert Now!", bootstyle="success", command=on_process, padding=(10, 5))
    process_button.grid(row=4, column=0, columnspan=3, pady=15, sticky="ew")

    # Copyright label
    copyright_label = ttk.Label(
        frame, text="© Erfan Saeidi", font=("Comic Sans MS", 10, "italic"), bootstyle="secondary"
    )
    copyright_label.grid(row=5, column=0, columnspan=3, pady=5, sticky="ew")

    root.mainloop()

if __name__ == "__main__":
    main_gui()
