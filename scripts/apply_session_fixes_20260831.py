"""Applies the fixes approved out of the 2026-08-31 final Word-vs-app
verification pass, and writes reports/session_changes_audit_20260831.csv --
same pattern as apply_session_fixes_20260827.py (read-before-write so the
audit row's old_value is the true value at write time).

Items:
  1. Lucas Joao, AE Cuncolim S/S (employee id 19492) -- mobile_phone was
     blank, Word has 8380015313. Confirmed a distinct real person/posting
     from id 19490 (AE Xeldem S/S, different mobile 8380015311, already
     correct) by parsing both Word segments directly, not just by name.
  2. Stephen Fernandes, Chief Electrical Engineer/HO (id 19474) --
     office_phone blank, Word has 0832-2224680 (single clean value).
  3. Stephen Fernandes, SE Circle II (N) (id 19475) -- office_phone blank,
     Word has 0832-2222354, 0832-2312194 (two office lines, not the
     multi-mobile ambiguity pattern -- applied as-is).
  4. Stephen Fernandes, SE Circle II (N) (id 19475) -- mobile_phone blank,
     Word has 7719012626 (single clean value).

id 19474's mobile_phone is intentionally NOT included -- Word's value
(7350611000/ 7719012626) is a multi-number cell, same ambiguous-primary-
number pattern flagged needs_manual_review throughout batch c
(reports/fix_proposal_batch_c_employee_fields.csv), so it isn't
auto-applied here either. See reports/needs_manual_review_20260831.csv.

Usage:
  python scripts/apply_session_fixes_20260831.py            # dry run, no writes
  python scripts/apply_session_fixes_20260831.py --apply    # commit + write audit CSV
"""

import csv
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from models import db
from models.employee import Employee

OUT_CSV = "reports/session_changes_audit_20260831.csv"
SESSION_DATE = "2026-08-31"

FIXES = [
    (19492, "mobile_phone", "8380015313",
     "Lucas Joao, AE Cuncolim S/S -- mobile blank in app, Word has 8380015313; "
     "confirmed distinct from id 19490 (AE Xeldem S/S, mobile 8380015311) by "
     "parsing both Word segments directly."),
    (19474, "office_phone", "0832-2224680",
     "Stephen Fernandes, Chief Electrical Engineer/HO -- office_phone blank, "
     "Word has single clean value 0832-2224680."),
    (19475, "office_phone", "0832-2222354, 0832-2312194",
     "Stephen Fernandes, SE Circle II (N) -- office_phone blank, Word has "
     "two office lines 0832-2222354, 0832-2312194; applied as-is (dual "
     "office-line pattern, not the ambiguous-mobile pattern)."),
    (19475, "mobile_phone", "7719012626",
     "Stephen Fernandes, SE Circle II (N) -- mobile_phone blank, Word has "
     "single clean value 7719012626."),
]


def main():
    apply = "--apply" in sys.argv
    audit_rows = []

    with app.app_context():
        for row_id, field, new_value, reason in FIXES:
            emp = db.session.get(Employee, row_id)
            old_value = getattr(emp, field)

            if apply:
                setattr(emp, field, new_value)

            audit_rows.append({
                "table": "employees", "row_id": row_id, "field": field,
                "old_value": old_value or "", "new_value": new_value,
                "reason": reason, "timestamp": SESSION_DATE,
            })

        if apply:
            db.session.commit()

    os.makedirs("reports", exist_ok=True)
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["table", "row_id", "field", "old_value", "new_value", "reason", "timestamp"])
        w.writeheader()
        w.writerows(audit_rows)

    print("=" * 70)
    print(f"  {'APPLIED' if apply else 'DRY RUN -- nothing written to DB'}")
    print("=" * 70)
    for r in audit_rows:
        print(f"  employees.{r['row_id']}.{r['field']}: {r['old_value']!r} -> {r['new_value']!r}")
    print(f"  Audit trail -> {OUT_CSV}")
    print("=" * 70)


if __name__ == "__main__":
    main()
