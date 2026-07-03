"""
Merges confirmed duplicate organization records.

Usage:
  python merge_duplicate_orgs.py          # dry run – shows what WOULD happen
  python merge_duplicate_orgs.py --apply  # commit changes to database
"""

import sys
from app import app
from models import db
from models.organization import Organization
from models.employee import Employee

# (remove_id, keep_id, reason)
# keep_id = None means delete the record (only safe if 0 employees)
MERGES = [
    # Location-only names → full NTPC plant names
    (47,  62,   "Gandhar → NTPC Gandhar"),
    (46,  63,   "Kawas → NTPC Kawas"),
    # All-caps duplicates of SLDCs
    (8,   56,   "CHHATTISGARH STATE LOAD DESPATCH CENTRE → Chhattisgarh State Load Despatch Centre"),
    (6,   52,   "GUJARAT STATE LOAD DESPATCH CENTRE → Gujarat State Load Despatch Centre"),
    (7,   53,   "MADHYA PRADESH STATE LOAD DESPATCH CENTRE → Madhya Pradesh State Load Despatch Centre"),
    (5,   48,   "MAHARASHTRA STATE LOAD DESPATCH CENTRE → Maharashtra State Load Despatch Centre"),
    (4,   11,   "WESTERN REGION LOAD DESPATCH CENTRE → Western Region Load Despatch Centre"),
    (214, 11,   "WRLDC (abbreviation) → Western Region Load Despatch Centre"),
    # 0-employee placeholder stubs
    (43,  None, "CHHATTISGARH (0 employees, stub) → DELETE"),
    (27,  None, "GUJARAT (0 employees, stub) → DELETE"),
]

# NTPC (id=9): split 2 employees to their respective regional HQs, then delete
# Adani Power Limited (id=181): legitimate corporate entry — left untouched
NTPC_SPLITS = [
    # (employee_name, destination_org_id, destination_name)
    ("Subrata Mandal",        61, "NTPC Mumbai Head Quarter (WR-1)"),
    ("Ashwini Kumar Tripathy", 66, "NTPC Raipur Head Quarter (WR-2)"),
]
NTPC_ORG_ID = 9


def run(apply: bool) -> None:
    sep = "=" * 65
    print(sep)
    print("  Organization Duplicate Merge Script")
    print("  DRY RUN — nothing saved" if not apply else "  APPLY MODE — changes will be committed")
    print(sep)

    with app.app_context():
        total_moved = 0

        for remove_id, keep_id, reason in MERGES:
            remove_org = db.session.get(Organization, remove_id)
            if not remove_org:
                print(f"\n  [SKIP] id={remove_id} not found in database")
                continue

            emp_count = Employee.query.filter_by(organization_id=remove_id).count()

            if keep_id is not None:
                keep_org = db.session.get(Organization, keep_id)
                if not keep_org:
                    print(f"\n  [SKIP] target id={keep_id} not found in database")
                    continue

                print(f"\n  MERGE  [{remove_id:>3}] {remove_org.organization_name}")
                print(f"   INTO  [{keep_id:>3}] {keep_org.organization_name}")
                print(f"   MOVE  {emp_count} employee(s)")
                print(f"   WHY   {reason}")

                if apply:
                    Employee.query.filter_by(organization_id=remove_id).update(
                        {"organization_id": keep_id}
                    )
                    # Copy address if canonical record is missing one
                    if not keep_org.address and remove_org.address:
                        keep_org.address = remove_org.address
                        print(f"   ADDR  Copied address from removed record")
                    db.session.delete(remove_org)
                    total_moved += emp_count

            else:
                print(f"\n  DELETE [{remove_id:>3}] {remove_org.organization_name}")
                print(f"   EMP   {emp_count} (expected 0)")
                print(f"   WHY   {reason}")
                if emp_count > 0:
                    print(f"   !! SKIPPED — has employees, unsafe to delete")
                    continue
                if apply:
                    db.session.delete(remove_org)

        # ── NTPC split: move each employee to their correct regional HQ ──────
        print(f"\n{sep}")
        print("  NTPC SPLIT (id=9 → regional HQs):")
        ntpc_org = db.session.get(Organization, NTPC_ORG_ID)
        if ntpc_org:
            all_moved = True
            for emp_name, dest_id, dest_name in NTPC_SPLITS:
                emp = Employee.query.filter(
                    Employee.employee_name.ilike(emp_name),
                    Employee.organization_id == NTPC_ORG_ID,
                ).first()
                if emp:
                    print(f"\n  MOVE  {emp.employee_name} | {emp.designation}")
                    print(f"   TO   [{dest_id}] {dest_name}")
                    if apply:
                        emp.organization_id = dest_id
                else:
                    print(f"\n  [NOT FOUND] {emp_name} in NTPC (id={NTPC_ORG_ID})")
                    all_moved = False

            if apply and all_moved:
                remaining = Employee.query.filter_by(organization_id=NTPC_ORG_ID).count()
                if remaining == 0:
                    db.session.delete(ntpc_org)
                    print(f"\n  DELETE [{NTPC_ORG_ID}] NTPC (now empty)")
                else:
                    print(f"\n  [KEEP] NTPC still has {remaining} employee(s) — not deleted")
        else:
            print(f"  [SKIP] NTPC (id={NTPC_ORG_ID}) not found — already removed?")

        print(f"\n  NOTE: Adani Power Limited (id=181) kept as-is (corporate office entry)")

        if apply:
            db.session.commit()
            print(f"\n  Done. {total_moved} employee(s) reassigned across standard merges.")
        else:
            print(f"\n  [Dry run complete — rerun with --apply to commit]")
        print(sep)


if __name__ == "__main__":
    run(apply="--apply" in sys.argv)
