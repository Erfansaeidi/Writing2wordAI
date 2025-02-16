import fitz  # PyMuPDF
import os

def convert_pdf_to_images(pdf_path, output_folder="input/"):
    """
    Converts a PDF file into images (one per page).

    :param pdf_path: Path to the PDF file
    :param output_folder: Directory to save the images
    :return: List of image file paths
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"File not found: {pdf_path}")

    os.makedirs(output_folder, exist_ok=True)

    doc = fitz.open(pdf_path)
    image_paths = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        pix = page.get_pixmap()
        image_path = os.path.join(output_folder, f"page_{page_num + 1}.png")
        pix.save(image_path)
        image_paths.append(image_path)

    return image_paths

if __name__ == "__main__":
    # Test with an example PDF
    test_pdf = "input/example.pdf"  # Change to your actual test PDF
    images = convert_pdf_to_images(test_pdf)
    print("Generated images:", images)
