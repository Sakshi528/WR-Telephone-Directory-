"""Shows all distinct organization strings in the DirectoryNumber table."""
from app import app
from models import db
from models.directory_number import DirectoryNumber

with app.app_context():
    rows = (
        db.session.query(
            DirectoryNumber.organization,
            db.func.count(DirectoryNumber.id).label("cnt"),
        )
        .group_by(DirectoryNumber.organization)
        .order_by(DirectoryNumber.organization)
        .all()
    )
    print(f"{'Count':<7} Organization string")
    print("-" * 60)
    for org, cnt in rows:
        print(f"{cnt:<7} {org}")
    print(f"\nTotal distinct org strings: {len(rows)}")
