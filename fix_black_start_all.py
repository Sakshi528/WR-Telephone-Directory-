"""
Renames ALL Black Start Facilitated Stations sub-orgs to include state name.

Format: "Black Start Facilitated Stations - <State> - <Station>"

Also updates matching DirectoryNumber.organization strings.

Usage:
  python fix_black_start_all.py          # dry run
  python fix_black_start_all.py --apply  # commit
"""

import sys
from app import app
from models import db
from models.organization import Organization
from models.directory_number import DirectoryNumber

# org_id: (current_name, new_name)
RENAMES = {
    # Madhya Pradesh (already partially fixed - add state)
    18: ("Black Start Facilitated Stations - Indira Sagar",
         "Black Start Facilitated Stations - Madhya Pradesh - Indira Sagar"),
    19: ("Black Start Facilitated Stations - Gandhi Sagar",
         "Black Start Facilitated Stations - Madhya Pradesh - Gandhi Sagar"),
    20: ("Black Start Facilitated Stations - Pench HPS",
         "Black Start Facilitated Stations - Madhya Pradesh - Pench HPS"),
    21: ("Black Start Facilitated Stations - Bargi HPS",
         "Black Start Facilitated Stations - Madhya Pradesh - Bargi HPS"),
    22: ("Black Start Facilitated Stations - Birshingpur HPS",
         "Black Start Facilitated Stations - Madhya Pradesh - Birshingpur HPS"),

    # Madhya Pradesh (new)
    23: ("Omkareshwar",  "Black Start Facilitated Stations - Madhya Pradesh - Omkareshwar"),
    24: ("Tons",         "Black Start Facilitated Stations - Madhya Pradesh - Tons"),
    25: ("Manikheda",   "Black Start Facilitated Stations - Madhya Pradesh - Manikheda"),
    26: ("Rajghat",     "Black Start Facilitated Stations - Madhya Pradesh - Rajghat"),

    # Gujarat
    28: ("UKAI",        "Black Start Facilitated Stations - Gujarat - UKAI"),
    29: ("SSNNL",       "Black Start Facilitated Stations - Gujarat - SSNNL"),
    30: ("Kadana",      "Black Start Facilitated Stations - Gujarat - Kadana"),
    31: ("GIPCL 1 & 2", "Black Start Facilitated Stations - Gujarat - GIPCL 1 & 2"),
    32: ("Dhuvaran",    "Black Start Facilitated Stations - Gujarat - Dhuvaran"),
    33: ("Sugen",       "Black Start Facilitated Stations - Gujarat - Sugen"),
    34: ("Mini Hydro",  "Black Start Facilitated Stations - Gujarat - Mini Hydro"),

    # Maharashtra
    35: ("MAHARASHTRA 17. Khopoli & Bhivpuri",
         "Black Start Facilitated Stations - Maharashtra - Khopoli & Bhivpuri"),
    36: ("Bhira",        "Black Start Facilitated Stations - Maharashtra - Bhira"),
    37: ("Koyna, Pophali","Black Start Facilitated Stations - Maharashtra - Koyna-Pophali"),
    38: ("Ghatghar",     "Black Start Facilitated Stations - Maharashtra - Ghatghar"),
    39: ("Trombay",      "Black Start Facilitated Stations - Maharashtra - Trombay"),
    40: ("RGPPL",        "Black Start Facilitated Stations - Maharashtra - RGPPL"),
    41: ("Paithan",      "Black Start Facilitated Stations - Maharashtra - Paithan"),
    42: ("Uran",         "Black Start Facilitated Stations - Maharashtra - Uran"),

    # Chhattisgarh
    44: ("Hasdeo Bango", "Black Start Facilitated Stations - Chhattisgarh - Hasdeo Bango"),
}

# DirectoryNumber.organization string patches (old -> new)
DN_PATCHES = {
    "Black Start Facilitated Stations - Indira Sagar":   "Black Start Facilitated Stations - Madhya Pradesh - Indira Sagar",
    "Black Start Facilitated Stations - Gandhi Sagar":   "Black Start Facilitated Stations - Madhya Pradesh - Gandhi Sagar",
    "Black Start Facilitated Stations - Pench HPS":      "Black Start Facilitated Stations - Madhya Pradesh - Pench HPS",
    "Black Start Facilitated Stations - Bargi HPS":      "Black Start Facilitated Stations - Madhya Pradesh - Bargi HPS",
    "Black Start Facilitated Stations - Birshingpur HPS":"Black Start Facilitated Stations - Madhya Pradesh - Birshingpur HPS",
    "Omkareshwar":   "Black Start Facilitated Stations - Madhya Pradesh - Omkareshwar",
    "Tons":          "Black Start Facilitated Stations - Madhya Pradesh - Tons",
    "Manikheda":     "Black Start Facilitated Stations - Madhya Pradesh - Manikheda",
    "Rajghat":       "Black Start Facilitated Stations - Madhya Pradesh - Rajghat",
    "UKAI":          "Black Start Facilitated Stations - Gujarat - UKAI",
    "SSNNL":         "Black Start Facilitated Stations - Gujarat - SSNNL",
    "Kadana":        "Black Start Facilitated Stations - Gujarat - Kadana",
    "GIPCL 1 & 2":  "Black Start Facilitated Stations - Gujarat - GIPCL 1 & 2",
    "Dhuvaran":      "Black Start Facilitated Stations - Gujarat - Dhuvaran",
    "Sugen":         "Black Start Facilitated Stations - Gujarat - Sugen",
    "Mini Hydro":    "Black Start Facilitated Stations - Gujarat - Mini Hydro",
    "MAHARASHTRA 17. Khopoli & Bhivpuri": "Black Start Facilitated Stations - Maharashtra - Khopoli & Bhivpuri",
    "Bhira":         "Black Start Facilitated Stations - Maharashtra - Bhira",
    "Koyna, Pophali":"Black Start Facilitated Stations - Maharashtra - Koyna-Pophali",
    "Ghatghar":      "Black Start Facilitated Stations - Maharashtra - Ghatghar",
    "Trombay":       "Black Start Facilitated Stations - Maharashtra - Trombay",
    "RGPPL":         "Black Start Facilitated Stations - Maharashtra - RGPPL",
    "Paithan":       "Black Start Facilitated Stations - Maharashtra - Paithan",
    "Uran":          "Black Start Facilitated Stations - Maharashtra - Uran",
    "Hasdeo Bango":  "Black Start Facilitated Stations - Chhattisgarh - Hasdeo Bango",
}


def run(apply: bool) -> None:
    print("DRY RUN\n" if not apply else "APPLY\n")

    with app.app_context():
        print("-- Organization renames --")
        for org_id, (expected, new_name) in RENAMES.items():
            org = db.session.get(Organization, org_id)
            if not org:
                print(f"  SKIP id={org_id} -- not found")
                continue
            if org.organization_name != expected:
                print(f"  MISMATCH id={org_id} -- actual: '{org.organization_name}'")
                continue
            print(f"  [{org_id}] '{expected}' -> '{new_name}'")
            if apply:
                org.organization_name = new_name

        print("\n-- DirectoryNumber.organization patches --")
        for old, new in DN_PATCHES.items():
            matches = DirectoryNumber.query.filter_by(organization=old).all()
            for dn in matches:
                print(f"  id={dn.id} | {dn.name} | '{old}' -> '{new}'")
                if apply:
                    dn.organization = new
            if not matches:
                print(f"  (no match) '{old}'")

        if apply:
            db.session.commit()
            print("\nDone -- all changes committed.")
        else:
            print("\n[Dry run -- rerun with --apply to commit]")


if __name__ == "__main__":
    run(apply="--apply" in sys.argv)
