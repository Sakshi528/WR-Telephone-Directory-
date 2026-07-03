from docx import Document
import re

doc = Document("uploads/Western Region Phone Directory 2025 Main_Telephone.docx")

for p in doc.paragraphs:
    text = p.text.strip()

    if re.match(r'^\d+\.', text):
        print(text)