"""
Fixes the remaining DirectoryNumber records:
  - 4 entries that had an address stored as the org name
  - 25 entries with no org assigned
Run: python fix_remaining_orgs.py --apply
"""
import sys
from app import app
from models import db
from models.directory_number import DirectoryNumber

# (record_id, correct_org_name)
FIXES = [
    # ── Address-as-org → correct company name ─────────────────────────────
    (70,  "INOX Green Energy Services Limited"),
    (76,  "Continuum Power Trading Private Limited"),
    (89,  "Srijan Energy Systems Private Limited"),
    (180, "KSK Mahanadi Power Company Limited (JSW Energy)"),

    # ── None org → WRLDC ──────────────────────────────────────────────────
    (211, "Western Region Load Despatch Centre"),

    # ── None org → Hospitals ──────────────────────────────────────────────
    (212, "Hospitals"),
    (213, "Hospitals"),
    (214, "Hospitals"),
    (215, "Hospitals"),
    (216, "Hospitals"),
    (217, "Hospitals"),
    (218, "Hospitals"),
    (219, "Hospitals"),
    (221, "Hospitals"),
    (222, "Hospitals"),
    (223, "Hospitals"),
    (224, "Hospitals"),
    (225, "Hospitals"),

    # ── None org → Forest Department ──────────────────────────────────────
    (226, "Forest Department"),
    (227, "Forest Department"),

    # ── None org → Emergency & Helpline Numbers ───────────────────────────
    (228, "Emergency & Helpline Numbers"),
    (229, "Emergency & Helpline Numbers"),
    (230, "Emergency & Helpline Numbers"),
    (231, "Emergency & Helpline Numbers"),
    (232, "Emergency & Helpline Numbers"),
    (234, "Emergency & Helpline Numbers"),
    (235, "Emergency & Helpline Numbers"),
    (236, "Emergency & Helpline Numbers"),
]

apply = "--apply" in sys.argv

with app.app_context():
    print("DRY RUN\n" if not apply else "APPLYING\n")
    fixed = 0
    for record_id, org_name in FIXES:
        row = db.session.get(DirectoryNumber, record_id)
        if not row:
            print(f"  [NOT FOUND] id={record_id}")
            continue
        old = row.organization or "(none)"
        print(f"  id={record_id:<4}  {old[:45]:<45}  →  {org_name}")
        if apply:
            row.organization = org_name
            fixed += 1

    if apply:
        db.session.commit()
        print(f"\nDone. {fixed} record(s) updated.")
    else:
        print(f"\n{len(FIXES)} record(s) would be updated. Run with --apply to commit.")
