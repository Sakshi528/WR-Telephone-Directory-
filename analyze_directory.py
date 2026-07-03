from docx import Document
import re

DOC_FILE = "uploads/Western Region Phone Directory 2025 Main_Telephone.docx"

doc = Document(DOC_FILE)

print("=" * 80)
print("HEADINGS FOUND")
print("=" * 80)

for i, p in enumerate(doc.paragraphs):
    text = p.text.strip()

    if not text:
        continue

    if re.match(r"^\d+(\.\d+)*\.?\s+", text):
        print(f"{i:5d} | {text}")