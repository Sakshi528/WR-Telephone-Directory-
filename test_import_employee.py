import re
import sys

from docx import Document

from app import app
from models import db
from models.department import Department
from models.employee import Employee
from models.organization import Organization
from models.update_request import UpdateRequest


DOC_FILE = "uploads/Western Region Phone Directory 2025 Main.docx"


def clean(value):
    return re.sub(r"\s+", " ", (value or "").replace("\xa0", " ")).strip()


def normalize(value):
    return re.sub(r"[^a-z0-9]+", " ", clean(value).lower()).strip()


def unique_values(values):
    result = []

    for value in values:
        value = clean(value)

        if value and value not in result:
            result.append(value)

    return result


def joined(values):
    return " / ".join(unique_values(values))


def has_email(value):
    return "@" in value


def looks_like_header(values):
    text = " ".join(normalize(value) for value in values)

    return (
        "name" in text
        and (
            "designation" in text
            or "mobile" in text
            or "telephone" in text
        )
        and (
            "email" in text
            or "mail" in text
        )
    )


def field_for_header(header, index):
    header = normalize(header)

    if "sl" in header or header in {"s no", "sr no", "sl no", "slno", "ssss"}:
        return None

    if "station name" in header or "name of hospital" in header:
        return None

    if "name" in header:
        return "employee_name"

    if "designation" in header:
        return "designation"

    if "activity" in header:
        return "department"

    if "department" in header:
        return "department"

    if header == "region" or "region" in header:
        return "region"

    if "location" in header:
        return "location"

    if "mobile" in header:
        return "mobile_phone"

    if "email" in header or "mail" in header:
        return "email"

    if "telephone r" in header:
        return "residence_phone"

    if "telephone f" in header or header == "fax":
        return None

    if "telephone o" in header or "telephone no" in header or header == "office":
        return "office_phone"

    if header == "telephone":
        return "office_phone"

    if index == 0 and "sldc" in header and "station" in header:
        return "organization"

    return None


def get_or_create(model, field_name, value):
    value = clean(value)

    if not value:
        return None

    item = model.query.filter_by(
        **{field_name: value}
    ).first()

    if item:
        return item

    item = model(
        **{field_name: value}
    )

    db.session.add(item)
    db.session.flush()

    return item


def extract_title(table, header_index):
    print("\n" + "=" * 80)
    print("TABLE FOUND")
    
    for row_index in range(header_index - 1, -1, -1):
        
        row_values = [
            clean(cell.text)
            for cell in table.rows[row_index].cells
        ]
        print("ROW:", row_values)

        print(
        "ROW VALUES:",
        [clean(cell.text) for cell in table.rows[row_index].cells]
    )
        values = unique_values(
            cell.text for cell in table.rows[row_index].cells
        )
        print("\nTABLE HEADER FOUND")

        if len(values) != 1:
            continue

        value = values[0]
        
        print("CANDIDATE:", value)


        if len(value) < 5:
            continue

        if re.search(r"\d", value):
            continue

        if any(token in normalize(value) for token in ["address", "plot", "road", "phone", "fax", "email", "bhavan"]):
            continue

        return value

    return ""


def row_to_employee(values, header_map, fallback_organization=""):
    fields = {
        "employee_name": "",
        "designation": "",
        "department": "",
        "organization": fallback_organization,
        "region": "",
        "location": "",
        "office_phone": "",
        "residence_phone": "",
        "mobile_phone": "",
        "email": "",
    }

    for field, indexes in header_map.items():
        fields[field] = joined(
            values[index]
            for index in indexes
            if index < len(values)
        )

    if fields["organization"] == fields["employee_name"]:
        fields["organization"] = fallback_organization

    if not fields["employee_name"] or not has_email(fields["email"]):
        return None

    if fields["employee_name"].lower() in {"control room", "shift in charge"}:
        pass

    return fields


def import_employees(replace=False):
    document = Document(DOC_FILE)
    inserted = 0
    skipped = 0

    if replace:
        if UpdateRequest.query.count():
            raise RuntimeError(
                "Cannot replace employees while update requests exist."
            )

        Employee.query.delete()
        Department.query.delete()
        Organization.query.delete()
        db.session.commit()

    for table in document.tables:
        header_index = None

        for index, row in enumerate(table.rows[:8]):
            values = [
                clean(cell.text)
                for cell in row.cells
            ]

            if looks_like_header(values):
                header_index = index
                break

        if header_index is None:
            continue

        headers = [
            clean(cell.text)
            for cell in table.rows[header_index].cells
        ]

        header_map = {}

        for index, header in enumerate(headers):
            field = field_for_header(header, index)

            if field:
                header_map.setdefault(field, []).append(index)

        if "employee_name" not in header_map or "email" not in header_map:
            continue

        fallback_organization = extract_title(
            table,
            header_index
        )
        print("=" * 80)
        print("ORG:", fallback_organization)


        if fallback_organization:
            get_or_create(
                Organization,
                "organization_name",
                fallback_organization
            )

        current_organization = fallback_organization

        for row in table.rows[header_index + 1:]:
            values = [
                clean(cell.text)
                for cell in row.cells
            ]

            unique = unique_values(values)

            if len(unique) == 1 and not has_email(unique[0]):
                current_organization = unique[0]
                continue

            fields = row_to_employee(
                values,
                header_map,
                current_organization
            )

            if not fields:
                skipped += 1
                continue

            existing = Employee.query.filter_by(
                employee_name=fields["employee_name"],
                email=fields["email"]
            ).first()

            if existing:
                skipped += 1
                continue

            department = get_or_create(
                Department,
                "department_name",
                fields["department"]
            )

            organization = get_or_create(
                Organization,
                "organization_name",
                fields["organization"]
            )

            employee = Employee(
                employee_name=fields["employee_name"],
                designation=fields["designation"],
                department_id=department.id if department else None,
                organization_id=organization.id if organization else None,
                location=fields["location"],
                region=fields["region"],
                office_phone=fields["office_phone"],
                residence_phone=fields["residence_phone"],
                mobile_phone=fields["mobile_phone"],
                email=fields["email"]
            )

            db.session.add(employee)
            inserted += 1

    db.session.commit()

    return inserted, skipped


if __name__ == "__main__":
    replace = "--replace" in sys.argv

    with app.app_context():
        inserted, skipped = import_employees(
            replace=replace
        )

    print(
        f"{inserted} employees imported, {skipped} rows skipped"
    )
