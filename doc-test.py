import Source_code
from docx import Document

doc = Document("water lab bill.docx")
count = 1

doc.add_paragraph("date:" + Source_code.formatted_date)

for paragraph in doc.paragraphs:
    print(str(count) + ' ' + paragraph.text)
    count = count + 1

