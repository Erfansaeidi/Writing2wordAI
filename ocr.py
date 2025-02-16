# ocr.py
from PIL import Image
import pytesseract
import os

# Set Tesseract OCR path (only required for Windows)
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def extract_text_from_image(image_path, language="eng"):
    """
    Extracts text from an image using Tesseract OCR.

    :param image_path: Path to the image file
    :param language: Language code (default is 'eng')
    :return: Extracted text as a string
    """
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"File not found: {image_path}")

    img = Image.open(image_path)
    # Use the provided language for OCR
    text = pytesseract.image_to_string(img, lang=language)
    return text

if __name__ == "__main__":
    test_image = "input/example.png"
    print("Extracted Text:\n", extract_text_from_image(test_image))
