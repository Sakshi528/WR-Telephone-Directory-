import re
import sys

from docx import Document

from app import app
from models import db
from models.directory_number import DirectoryNumber
from models.employee import Employee


DOC_FILE = "uploads/Western Region Phone Directory 2025 Main_Telephone.docx"


def clean(value):
    return re.sub(r"\s+", " ", (value or "").replace("\xa0", " ")).strip()


def normalize(value):
    return re.sub(r"[^a-z0-9]+", " ", clean(value).lower()).strip()


def add_number(name, phone_number, category):
    name = clean(name)
    phone_number = clean(phone_number)
    category = clean(category)

    if not name or not phone_number:
        return False

    existing = DirectoryNumber.query.filter_by(
        name=name,
        phone_number=phone_number
    ).first()

    if existing:
        return False

    db.session.add(
        DirectoryNumber(
            name=name,
            phone_number=phone_number,
            category=category or "Directory"
        )
    )

    return True


def import_from_word():
    document = Document(DOC_FILE)
    inserted = 0

    for table in document.tables:
        if not table.rows:
            continue

        headers = [
            normalize(cell.text)
            for cell in table.rows[0].cells
        ]

        name_index = None
        phone_indexes = []
        category = "Directory"

        for index, header in enumerate(headers):
            if header in {"name", "phone no"}:
                if header == "name":
                    name_index = index
                elif header == "phone no":
                    phone_indexes.append(index)

            if "name of hospitals" in header:
                name_index = index
                category = "Hospital"

            if "contact number" in header:
                phone_indexes.append(index)

            if "emergency number" in header:
                phone_indexes.append(index)
                category = "Hospital"

            if header in {"land line 1", "land line 2", "mobile 1", "mobile 2"}:
                phone_indexes.append(index)
                category = "Control Room"

        if name_index is None and phone_indexes and "land line 1" in headers:
            name = "WRLDC Emergency Contact"
        elif name_index is None:
            continue
        else:
            name = None

        for row in table.rows[1:]:
            values = [
                clean(cell.text)
                for cell in row.cells
            ]

            row_name = name or values[name_index]

            numbers = [
                values[index]
                for index in phone_indexes
                if index < len(values) and values[index]
            ]

            if "switch" in row_name.lower():
                if add_number(row_name, " / ".join(numbers), "Switchyard"):
                    inserted += 1
                continue

            if "control room" in row_name.lower():
                if add_number(row_name, " / ".join(numbers), "Control Room"):
                    inserted += 1
                continue

            if add_number(row_name, " / ".join(numbers), category):
                inserted += 1

    return inserted


def import_control_rooms_from_employees():
    """
    Control Rooms should not come from Employee records anymore.
    They should be imported directly from the Word document.

    Keeping this function so the rest of the code doesn't break.
    """
    return 0


def import_directory_numbers(replace=False):
    if replace:
        DirectoryNumber.query.delete()
        db.session.commit()

    inserted = import_from_word()
    inserted += import_control_rooms_from_employees()

    db.session.commit()

    return inserted


if __name__ == "__main__":
    replace = "--replace" in sys.argv

    with app.app_context():
        inserted = import_directory_numbers(
            replace=replace
        )

    print(
        f"{inserted} directory numbers imported"
    )
