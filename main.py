import os
from ocr import extract_text_from_image
from pdf_processing import convert_pdf_to_images
from word_generator import save_text_to_word

def process_file(input_path, output_path="output/result.docx"):
    """
    Processes an image or PDF and converts it into a Word document.

    :param input_path: Path to the input file (image or PDF)
    :param output_path: Path to save the output Word file
    """
    extracted_text = ""

    if input_path.lower().endswith((".png", ".jpg", ".jpeg")):
        extracted_text = extract_text_from_image(input_path)

    elif input_path.lower().endswith(".pdf"):
        images = convert_pdf_to_images(input_path)
        for img_path in images:
            extracted_text += extract_text_from_image(img_path) + "\n"

    else:
        raise ValueError("Unsupported file format. Use PNG, JPG, or PDF.")

    save_text_to_word(extracted_text, output_path)
    print(f"Processing complete. Text saved to {output_path}")

if __name__ == "__main__":
    test_file = "input/example.jpg"  # Change to your actual file
    process_file(test_file)
