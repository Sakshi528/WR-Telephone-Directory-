from docx import Document
import re

doc = Document("uploads/Western Region Phone Directory 2025 Main.docx")

SUSPICIOUS = [
    "great",
    "achievement",
    "success",
    "team",
    "trust",
    "ambition",
    "fulfilling",
    "thinking",
    "results",
    "leadership",
    "working together",
]

for table_index, table in enumerate(doc.tables):

    for row_index, row in enumerate(table.rows):

        values = []

        for cell in row.cells:
            text = cell.text.strip()

            if text:
                values.append(text)

        if not values:
            continue

        unique = list(dict.fromkeys(values))

        if len(unique) != 1:
            continue

        text = unique[0]

        lower = text.lower()

        # motivational quotes
        if any(word in lower for word in SUSPICIOUS):
            print("\nQUOTE")
            print(f"Table {table_index}")
            print(text)

        # address lines
        elif (
            "road" in lower
            or "plot no" in lower
            or "sector" in lower
            or "fax" in lower
            or "email:" in lower
            or "dist." in lower
            or "p.o." in lower
        ):
            print("\nADDRESS")
            print(f"Table {table_index}")
            print(text)

        # phone-heavy lines
        elif len(re.findall(r"\d", text)) > 10:
            print("\nPHONE LINE")
            print(f"Table {table_index}")
            print(text)

        # very long organization candidates
        elif len(text) > 80:
            print("\nLONG TEXT")
            print(f"Table {table_index}")
            print(text)