import json
from docx import Document
from fpdf import FPDF
import base64
from PIL import Image
from io import BytesIO
import os


def generate_pdf_from_word(data, doc_path):
    doc = Document(doc_path)
    for person in data:
        pdf_file = f"{person['First Name']}_{person['Last Name']}_details.pdf"
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        for paragraph in doc.paragraphs:
            text = paragraph.text
            # Replace the placeholder text with an empty string
            if '{{Image}}' in text and person.get('Image', None):
                text = text.replace('{{Image}}', '')
            else:
                text = text.replace('{{First Name}}', person['First Name'])
                text = text.replace('{{Last Name}}', person['Last Name'])
                text = text.replace('{{E-mail}}', person['E-mail'])
                text = text.replace('{{Date of birth}}', person['Date of birth'])
            pdf.cell(200, 10, txt=text, ln=True)

        # Insert image
        image_path = person.get('Image', None)
        if image_path:
            if image_path.startswith('data:image'):
                image_data = base64.b64decode(image_path.split(',')[1])
                image = Image.open(BytesIO(image_data))
            else:
                image = Image.open(image_path)

            # Convert image to 'RGB' mode
            image = image.convert('RGB')

            image_file = f"{person['First Name']}_{person['Last Name']}_image.jpg"
            image.save(image_file)

            # Get the coordinates of the placeholder
            placeholder_x, placeholder_y = 10, pdf.get_y() + 10

            # Add the image at the position of the placeholder
            pdf.image(image_file, x=placeholder_x, y=placeholder_y, w=50)

            # Delete temporary image file
            os.remove(image_file)

        pdf.output(pdf_file)


if __name__ == "__main__":
    with open('D:\\LPL Python_research\\data.json') as f:
        data = json.load(f)
        generate_pdf_from_word(data, 'D:\\LPL Python_research\\Testv2.docx')
