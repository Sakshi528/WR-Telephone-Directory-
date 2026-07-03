"""
Fixes organization name strings in the DirectoryNumber (Telephone Directory) table.

Usage:
  python fix_directory_orgs.py          # dry run
  python fix_directory_orgs.py --apply  # commit changes
"""

import sys
from app import app
from models import db
from models.directory_number import DirectoryNumber

# (old_string, new_string)  — exact match, case-sensitive
RENAMES = [
    # NTPC plant name normalization
    ("Gandhar",             "NTPC Gandhar"),
    ("Kawas",               "NTPC Kawas"),

    # All-caps SLDC duplicates → proper case
    ("CHHATTISGARH STATE LOAD DESPATCH CENTRE",  "Chhattisgarh State Load Despatch Centre"),
    ("GUJARAT STATE LOAD DESPATCH CENTRE",        "Gujarat State Load Despatch Centre"),
    ("MADHYA PRADESH STATE LOAD DESPATCH CENTRE", "Madhya Pradesh State Load Despatch Centre"),
    ("MAHARASHTRA STATE LOAD DESPATCH CENTRE",    "Maharashtra State Load Despatch Centre"),
    ("WESTERN REGION LOAD DESPATCH CENTRE",       "Western Region Load Despatch Centre"),

    # Typos
    ("ReNew Power Limite d", "ReNew Power Limited"),
    ("Damanand Diu",         "Daman and Diu"),
]

# These strings are duplicates of entries above after renaming — delete them
DELETE_ORG_STRINGS = [
    "N TPC Kawas",   # exact duplicate of Kawas (same Control Room, same phone)
]


def run(apply: bool) -> None:
    sep = "=" * 65
    print(sep)
    print("  Telephone Directory Org-String Cleanup")
    print("  DRY RUN — nothing saved" if not apply else "  APPLY MODE — saving to database")
    print(sep)

    with app.app_context():

        # ── 1. Renames ─────────────────────────────────────────────────────
        print("\n  RENAMES:")
        total_renamed = 0
        for old, new in RENAMES:
            rows = DirectoryNumber.query.filter_by(organization=old).all()
            count = len(rows)
            if count == 0:
                print(f"  [NOT FOUND] \"{old}\"")
                continue
            print(f"  {count:>2}x  \"{old}\"")
            print(f"      → \"{new}\"")
            if apply:
                for row in rows:
                    row.organization = new
                total_renamed += count

        # ── 2. Deletes (exact duplicates) ──────────────────────────────────
        print("\n  DELETES (duplicate rows):")
        total_deleted = 0
        for org_str in DELETE_ORG_STRINGS:
            rows = DirectoryNumber.query.filter_by(organization=org_str).all()
            if not rows:
                print(f"  [NOT FOUND] \"{org_str}\"")
                continue
            for row in rows:
                print(f"  DELETE id={row.id} | {row.organization} | {row.name} | {row.phone_number}")
            if apply:
                for row in rows:
                    db.session.delete(row)
                total_deleted += len(rows)

        if apply:
            db.session.commit()
            print(f"\n  Done. {total_renamed} record(s) renamed, {total_deleted} deleted.")
        else:
            print(f"\n  [Dry run — rerun with --apply to commit]")

        # ── 3. Manual review items ─────────────────────────────────────────
        print(f"\n{sep}")
        print("  NEEDS MANUAL REVIEW (not auto-fixed):\n")

        # Address-as-org entries
        address_patterns = [
            "/82/A/431/A",
            "C Wing, 402",
            "Dayapar (Kutch)",
            "Office No. 1",
        ]
        for pattern in address_patterns:
            rows = DirectoryNumber.query.filter(
                DirectoryNumber.organization.ilike(f"%{pattern}%")
            ).all()
            if rows:
                print(f"  ADDRESS used as org name ({len(rows)} record(s)):")
                for row in rows:
                    print(f"    id={row.id} | org=\"{row.organization}\"")
                    print(f"    name={row.name} | phone={row.phone_number}")
                print()

        # None organization records
        none_rows = DirectoryNumber.query.filter_by(organization=None).all()
        if none_rows:
            print(f"  NO ORGANIZATION assigned ({len(none_rows)} record(s)):")
            for row in none_rows:
                print(f"    id={row.id} | {row.name or '(no name)'} | {row.phone_number or ''} | {row.category or ''}")
            print()

        print("  NOTE: NTPC Blackstart Units control room is stored under 'Kawas'.")
        print("        After this fix it will be under 'NTPC Kawas' — search 'Kawas' to find it.")
        print(sep)


if __name__ == "__main__":
    run(apply="--apply" in sys.argv)
