from app import app
from models.organization import Organization
from models.employee import Employee
from models.directory_number import DirectoryNumber

station = input("Enter station name: ").strip()

with app.app_context():
    org = Organization.query.filter(
        Organization.organization_name.ilike(f"%{station}%")
    ).first()

    if not org:
        print("Station not found.")
        raise SystemExit

    print("\n" + "=" * 70)
    print(org.organization_name)
    print("=" * 70)

    print("\nEMPLOYEES")
    print("-" * 70)
    employees = (
        Employee.query
        .filter_by(organization_id=org.id)
        .order_by(Employee.employee_name)
        .all()
    )

    print(f"Total: {len(employees)}")
    for e in employees:
        print(f"{e.employee_name} | {e.designation}")

    print("\nDIRECTORY NUMBERS")
    print("-" * 70)

    numbers = (
        DirectoryNumber.query
        .filter_by(organization_id=org.id)
        .order_by(DirectoryNumber.category, DirectoryNumber.name)
        .all()
    )

    print(f"Total: {len(numbers)}\n")

    for n in numbers:
        print(f"Category : {n.category}")
        print(f"Name     : {n.name}")
        print(f"Phone    : {n.phone_number}")
        print(f"Email    : {n.email}")
        print("-" * 70)