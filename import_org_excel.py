"""
Imports all organizations and sub-organizations from the Excel file into the database.

- Organization rows  → region = state, address from Excel
- SubOrganization rows → region = parent org name (so searching the parent finds sub-orgs),
                          address from Excel (state is already embedded in the address text)
- Creates any org/sub-org not already in the DB
- Matching uses fuzzy normalization to avoid duplicates from minor punctuation differences

Usage:
  python import_org_excel.py          # dry run (no DB changes)
  python import_org_excel.py --apply  # commit changes
"""

import re
import sys
import openpyxl

from app import app
from models import db
from models.organization import Organization

EXCEL_PATH = "uploads/WR_Organizations_SubOrganizations.xlsx"


def norm_key(s):
    """Normalize to lowercase alphanumeric-only key for fuzzy dedup."""
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


def clean_name(raw):
    """Strip leading serial numbers (e.g. '1. Indira Sagar' → 'Indira Sagar')."""
    return re.sub(r"^[\d\s.–-]+", "", raw).strip()


def run(apply: bool) -> None:
    sep = "=" * 65
    print(sep)
    print("  Organization / Sub-Organization Excel Import")
    print("  DRY RUN" if not apply else "  APPLY MODE — changes will be committed")
    print(sep)

    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True)
    ws = wb["Org_SubOrg_List"]
    rows = list(ws.iter_rows(values_only=True))[1:]   # skip header

    with app.app_context():
        # Build fuzzy lookup: normalized key → org object
        existing_exact = {
            o.organization_name.strip().lower(): o
            for o in Organization.query.all()
        }
        existing_fuzzy = {
            norm_key(o.organization_name): o
            for o in Organization.query.all()
        }

        created = 0
        updated = 0
        already_ok = 0

        def process(name, region_value, address, force_region=False):
            """
            name         – org/sub-org name (serial numbers already stripped)
            region_value – state for orgs; parent org name for sub-orgs
            address      – full address string from Excel
            force_region – if True, always overwrite region (used for sub-orgs)
            """
            nonlocal created, updated, already_ok

            name         = clean_name(str(name).strip()         if name         else "")
            region_value = str(region_value).strip()            if region_value else ""
            address      = str(address).strip()                 if address      else ""

            if not name:
                return

            org = existing_exact.get(name.lower()) or existing_fuzzy.get(norm_key(name))

            if org:
                changed = False
                # For sub-orgs (force_region=True), always update region so the
                # parent-org tag is set; for orgs, only fill in when empty.
                if region_value and (force_region or not org.region):
                    if apply:
                        org.region = region_value
                    changed = True
                if address and not org.address:
                    if apply:
                        org.address = address
                    changed = True

                if changed:
                    print(f"  [UPDATE] {name}")
                    updated += 1
                else:
                    already_ok += 1

            else:
                print(f"  [CREATE] {name}")
                if apply:
                    new_org = Organization(
                        organization_name=name,
                        region=region_value or None,
                        address=address or None,
                    )
                    db.session.add(new_org)
                    db.session.flush()
                    existing_exact[name.lower()] = new_org
                    existing_fuzzy[norm_key(name)] = new_org
                created += 1

        current_org_name = ""
        for row in rows:
            _, type_, org_name, sub_name, state, address = row

            if type_ == "Organization":
                current_org_name = str(org_name).strip() if org_name else ""
                process(org_name, state, address, force_region=False)

            elif type_ == "SubOrganization":
                # Use parent org name as region so sub-orgs are discoverable
                # by searching the parent (e.g. "black start", "ists", "ntpc raipur").
                process(sub_name, current_org_name, address, force_region=True)

        if apply:
            db.session.commit()

        print()
        print(sep)
        print(f"  Created  : {created}")
        print(f"  Updated  : {updated}  (region / address updated)")
        print(f"  Unchanged: {already_ok}  (already in DB)")
        if not apply:
            print(f"\n  [Dry run — rerun with --apply to commit]")
        print(sep)


if __name__ == "__main__":
    run(apply="--apply" in sys.argv)
