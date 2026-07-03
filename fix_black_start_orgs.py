"""
Fixes Black Start Facilitated Stations sub-org naming.

Problems:
  - Indira Sagar employees stored under "Black Start Facilitated Stations" ->not findable by "indira sagar"
  - Gandhi Sagar, Pench, Bargi, Birshingpur are standalone org names ->not findable by "black start"

Fix: rename all 5 sub-station orgs to "Black Start Facilitated Stations - <SubStation>"
     and update matching DirectoryNumber.organization strings.

Usage:
  python fix_black_start_orgs.py          # dry run
  python fix_black_start_orgs.py --apply  # commit
"""

import sys
from app import app
from models import db
from models.organization import Organization
from models.directory_number import DirectoryNumber

RENAMES = {
    # org_id: (old_name, new_name)
    18: ("Black Start Facilitated Stations",  "Black Start Facilitated Stations - Indira Sagar"),
    19: ("Gandhi Sagar",                       "Black Start Facilitated Stations - Gandhi Sagar"),
    20: ("Pench",                              "Black Start Facilitated Stations - Pench HPS"),
    21: ("Bargi",                              "Black Start Facilitated Stations - Bargi HPS"),
    22: ("Birshingpur",                        "Black Start Facilitated Stations - Birshingpur HPS"),
}

# DirectoryNumber.organization string patches (old ->new)
DN_PATCHES = {
    "Black Start Facilitated Stations": "Black Start Facilitated Stations - Indira Sagar",
    "Gandhi Sagar":                      "Black Start Facilitated Stations - Gandhi Sagar",
    "Pench":                             "Black Start Facilitated Stations - Pench HPS",
    "Bargi":                             "Black Start Facilitated Stations - Bargi HPS",
    "Birshingpur":                       "Black Start Facilitated Stations - Birshingpur HPS",
}

def run(apply: bool) -> None:
    print("DRY RUN" if not apply else "APPLY")

    with app.app_context():
        print("\n-- Organization renames --")
        for org_id, (old, new) in RENAMES.items():
            org = db.session.get(Organization, org_id)
            if not org:
                print(f"  SKIP id={org_id} — not found")
                continue
            if org.organization_name != old:
                print(f"  SKIP id={org_id} — name is '{org.organization_name}' (expected '{old}')")
                continue
            print(f"  [{org_id}] '{old}' ->'{new}'")
            if apply:
                org.organization_name = new

        print("\n-- DirectoryNumber.organization patches --")
        for old, new in DN_PATCHES.items():
            matches = DirectoryNumber.query.filter_by(organization=old).all()
            for dn in matches:
                print(f"  id={dn.id} | {dn.name} | '{old}' ->'{new}'")
                if apply:
                    dn.organization = new
            if not matches:
                print(f"  (no match for '{old}')")

        if apply:
            db.session.commit()
            print("\nDone — changes committed.")
        else:
            print("\n[Dry run — rerun with --apply to commit]")

if __name__ == "__main__":
    run(apply="--apply" in sys.argv)
