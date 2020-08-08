'''
from docx import Document

doc = Document("water lab bill.docx")

for table in doc.tables:
    for row in table.rows:
        print(row)
'''

print(float(0.3) > 0.5)