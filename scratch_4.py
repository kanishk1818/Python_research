import json
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader


def generate_pdf_from_word(data, doc_path):
    doc = Document(doc_path)
    for person in data:
        pdf_file = f"{person['First Name']}_{person['Last Name']}_details.pdf"
        c = canvas.Canvas(pdf_file, pagesize=letter)
        y_position = 750  # Initial y-position
        for paragraph in doc.paragraphs:
            text = paragraph.text
            text = text.replace('{{First Name}}', person['First Name'])
            text = text.replace('{{Last Name}}', person['Last Name'])
            text = text.replace('{{E-mail}}', person['E-mail'])
            text = text.replace('{{Date of birth}}', person['Date of birth'])
            c.drawString(100, y_position, text)
            y_position -= 20  # Adjust the vertical position for the next paragraph

        # Insert image
        image_path = person.get('Image', None)
        if image_path:
            c.drawImage(ImageReader(image_path), 100, y_position - 50, width=100, height=100)

        c.save()


if __name__ == "__main__":
    with open('D:\\LPL Python_research\\data.json') as f:
        data = json.load(f)
        generate_pdf_from_word(data, 'D:\\LPL Python_research\\TestWord.docx')
