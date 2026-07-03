"""
Run this to see all organizations and detect potential duplicates.
Usage:  python find_duplicate_orgs.py
Output: prints to screen AND saves to duplicate_orgs_report.txt
"""

import sys
from app import app
from models import db
from models.organization import Organization
from models.employee import Employee


def run():
    lines = []

    def out(text=""):
        print(text)
        lines.append(text)

    with app.app_context():
        orgs = (
            db.session.query(
                Organization.id,
                Organization.organization_name,
                Organization.address,
                db.func.count(Employee.id).label("emp_count"),
            )
            .outerjoin(Employee, Organization.id == Employee.organization_id)
            .group_by(Organization.id, Organization.organization_name, Organization.address)
            .order_by(Organization.organization_name)
            .all()
        )

        out(f"Total organizations in database: {len(orgs)}")
        out("=" * 70)
        out(f"{'ID':<6} {'Employees':<10} {'Has Address':<12} Name")
        out("-" * 70)
        for o in orgs:
            has_addr = "YES" if o.address else "no"
            out(f"{o.id:<6} {o.emp_count:<10} {has_addr:<12} {o.organization_name}")

        # ── Find potential duplicates (one name contained in another) ──────────
        out()
        out("=" * 70)
        out("POTENTIAL DUPLICATES (one name is a substring of another):")
        out("=" * 70)

        found = 0
        names = [(o.id, o.organization_name.strip().lower(), o) for o in orgs]

        seen_pairs = set()
        for i, (id_a, name_a, org_a) in enumerate(names):
            for j, (id_b, name_b, org_b) in enumerate(names):
                if i >= j:
                    continue
                pair = (min(id_a, id_b), max(id_a, id_b))
                if pair in seen_pairs:
                    continue

                # One name contained in another (min 4 chars to avoid false hits)
                if (name_b in name_a and len(name_b) >= 4) or (
                    name_a in name_b and len(name_a) >= 4
                ):
                    seen_pairs.add(pair)
                    found += 1
                    longer = org_a if len(name_a) >= len(name_b) else org_b
                    shorter = org_b if len(name_a) >= len(name_b) else org_a
                    emp_shorter = next(
                        o.emp_count for o in orgs if o.id == shorter.id
                    )
                    emp_longer = next(
                        o.emp_count for o in orgs if o.id == longer.id
                    )
                    out()
                    out(f"  KEEP   [{longer.id}] {longer.organization_name}  "
                        f"(employees: {emp_longer}, address: {'YES' if longer.address else 'no'})")
                    out(f"  REMOVE [{shorter.id}] {shorter.organization_name}  "
                        f"(employees: {emp_shorter}, address: {'YES' if shorter.address else 'no'})")

        if found == 0:
            out("  None detected by substring matching.")

        out()
        out(f"Total potential duplicate pairs found: {found}")

    with open("duplicate_orgs_report.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print("\nReport saved to duplicate_orgs_report.txt")


if __name__ == "__main__":
    run()
