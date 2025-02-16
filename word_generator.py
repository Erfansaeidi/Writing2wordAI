from docx import Document
import os

def save_text_to_word(text, output_path):
    """
    Saves extracted text to a Word document.

    :param text: Extracted text
    :param output_path: Path to save the .docx file
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    doc = Document()
    doc.add_paragraph(text)
    doc.save(output_path)

if __name__ == "__main__":
    # Test with sample text
    test_text = "This is a test document."
    output_file = "output/result.docx"
    save_text_to_word(test_text, output_file)
    print(f"Word document saved at: {output_file}")
    