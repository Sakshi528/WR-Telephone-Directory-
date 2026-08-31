"""Applies the single approved category fix from the Word-section-vs-app-
category audit (2026-08-31): org id 6506, "Vidyut Bhavan, 3rd Floor,
Panaji, Goa-403 001" -- the populated Goa Electricity Department org
(all real staff: Stephen Fernandes, Lucas Joao, etc.) -- currently
category=Others, should be SLDC per its Word section heading, same as
every other state SLDC.

Logs to both the real audit_log table (services/audit_service.py,
same as a live Manage Organizations UI edit -- organizations.category_id
has real audit infrastructure, unlike directory_numbers) and
reports/session_changes_audit_20260831.csv, same pattern as
apply_session_fixes_20260827.py's apply_org_category_move.

Does NOT touch parent_id or org id 6297 (Goa Electricity Department, the
parent container node) -- confirmed load-bearing in the prior turn's
reference check (routes/user_routes.py:498 walks parent_id for directory-
export subtrees), left untouched by design, not an oversight.

Usage:
  python scripts/apply_goa_category_fix_20260831.py            # dry run
  python scripts/apply_goa_category_fix_20260831.py --apply    # commit + audit
"""

import csv
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from models import db
from models.organization import Organization
from models.organization_category import OrganizationCategory
from services.audit_service import log_audit_event

AUDIT_CSV = "reports/session_changes_audit_20260831.csv"
SESSION_DATE = "2026-08-31"
ORG_ID = 6506
NEW_CATEGORY_NAME = "SLDC"
REASON = (
    "Word-section-vs-app-category audit (2026-08-31): org holds all the real "
    "Goa Electricity Department staff (Stephen Fernandes, Lucas Joao, etc.) "
    "under Word's 'Goa Electricity Department' section heading, which maps to "
    "SLDC for every other state (Maharashtra/Gujarat/MP/Chhattisgarh). Was "
    "Others. The empty parent container org (id 6297, 'Goa Electricity "
    "Department') already carries category=SLDC but holds no real data; this "
    "change brings the populated org in line with it. parent_id/hierarchy "
    "untouched."
)


def main():
    apply = "--apply" in sys.argv

    with app.app_context():
        org = db.session.get(Organization, ORG_ID)
        new_category = OrganizationCategory.query.filter_by(category_name=NEW_CATEGORY_NAME).first()
        old_category = db.session.get(OrganizationCategory, org.category_id) if org.category_id else None
        old_category_name = old_category.category_name if old_category else None

        if old_category_name == NEW_CATEGORY_NAME:
            print(f"Already {NEW_CATEGORY_NAME} -- nothing to do.")
            return

        if apply:
            log_audit_event(
                module="organizations", record_type="Organization", record_id=org.id,
                action="CATEGORY_CHANGE", field_name="category_id",
                old_value=str(org.category_id) if org.category_id else None,
                new_value=str(new_category.id),
                reason=REASON,
            )
            org.category_id = new_category.id
            db.session.commit()

    audit_row = {
        "table": "organizations", "row_id": ORG_ID, "field": "category_id (via category_name)",
        "old_value": old_category_name or "", "new_value": NEW_CATEGORY_NAME,
        "reason": f"{REASON} Also logged to the real audit_log table via "
                  f"services.audit_service.log_audit_event, same as a Manage Organizations UI edit would.",
        "timestamp": SESSION_DATE,
    }

    if apply:
        file_exists = os.path.exists(AUDIT_CSV)
        with open(AUDIT_CSV, "a", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["table", "row_id", "field", "old_value", "new_value", "reason", "timestamp"])
            if not file_exists:
                w.writeheader()
            w.writerow(audit_row)

    print("=" * 70)
    print(f"  {'APPLIED' if apply else 'DRY RUN -- nothing written to DB'}")
    print("=" * 70)
    print(f"  organizations.{ORG_ID}.category_id: {old_category_name!r} -> {NEW_CATEGORY_NAME!r}")
    print(f"  Audit trail -> {AUDIT_CSV}" if apply else "")
    print("=" * 70)


if __name__ == "__main__":
    main()
