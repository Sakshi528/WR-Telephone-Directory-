"""
fix_missing_data.py
===================
Scans the Word document to recover missing phone numbers for control rooms,
fixes 9 employees with no organization, removes exact duplicates,
and reassigns orphaned employees to their correct organization.

Run: python fix_missing_data.py
"""

import re
from docx import Document
from app import app
from models import db
from models.employee import Employee
from models.organization import Organization
from models.directory_number import DirectoryNumber

DOC_FILE = "uploads/Western Region Phone Directory 2025 Main_Telephone.docx"

# ─── HELPERS ─────────────────────────────────────────────────────────────────

PHONE_RE = re.compile(
    r'(?:(?:\+91|0)?[-.\s]?)?'
    r'(?:\(?\d{2,5}\)?[-.\s]?)?'
    r'\d{6,10}'
    r'(?:[-.\s/]\d{3,10})*'
)

def looks_like_phone(text):
    text = text.strip()
    if not text or len(text) < 6:
        return False
    digits = re.sub(r'\D', '', text)
    return (
        len(digits) >= 7
        and not text.lower().startswith('ip no')
        and re.search(r'\d{6,}', text)
    )

def extract_phones_from_text(text):
    return PHONE_RE.findall(text.strip())

def org_key(name):
    """Normalised lowercase for fuzzy matching."""
    return re.sub(r'[^a-z0-9]', '', (name or '').lower())


# ─── STEP 1: Build a lookup of org_name → all phones found in the Word doc ──

def build_word_phone_map(doc):
    """
    Walk every table cell in the Word document.
    For each table, track the current organisation heading.
    Return dict: normalised_org_key -> list of phone strings.
    """
    phone_map = {}   # org_key -> [phones]
    name_map  = {}   # org_key -> original org name

    for table in doc.tables:
        current_org = None

        for row in table.rows:
            values = [cell.text.strip() for cell in row.cells]
            text   = " | ".join(values)

            # Detect org-heading rows (single non-empty cell or bold heading)
            non_empty = [v for v in values if v]
            if len(non_empty) == 1:
                candidate = non_empty[0]
                if len(candidate) > 5 and not looks_like_phone(candidate):
                    current_org = candidate
                    key = org_key(candidate)
                    if key not in name_map:
                        name_map[key] = candidate
                    continue

            if current_org is None:
                continue

            key = org_key(current_org)

            # Find all phone-like cells in this row
            for cell_text in values:
                cell_text = cell_text.strip()
                if not cell_text:
                    continue
                # Grab every phone pattern in the cell
                phones = extract_phones_from_text(cell_text)
                for ph in phones:
                    if len(re.sub(r'\D', '', ph)) >= 7:
                        phone_map.setdefault(key, [])
                        if ph not in phone_map[key]:
                            phone_map[key].append(ph)

    return phone_map, name_map


# ─── STEP 2: Fix control rooms with blank phone numbers ──────────────────────

MANUAL_PHONE_FIX = {
    # ID : phone_number  (from Word doc — manually verified)
    76:  "7710011092",                # Continuum Power / C Wing Powai (phone in org field)
}

# Organizations in Word doc that had no phone at all (genuinely blank in source)
KNOWN_NO_PHONE_ORGS = {
    "Gujarat Industries Power Company Limited",
    "Electro Solaire Private Limited",
    "Gujarat State Electricity Corporation Limited",
    "Solapur Solar PV Project",
    "Wind Two Renergy Private Limited",
    "Gujarat State Electricity Corporation Limited- Phase II",
}


# ─── STEP 3: Fix employees with no organization ───────────────────────────────

# Mapping of employee ID → correct organization name (from source document)
EMPLOYEE_ORG_FIX = {
    219: "MAHARASHTRA STATE ELECTRICITY DISTRIBUTION CO. LTD.",   # CMD MSEDCL
    220: "MAHARASHTRA STATE ELECTRICITY DISTRIBUTION CO. LTD.",   # CE PP
    221: "MAHARASHTRA STATE ELECTRICITY DISTRIBUTION CO. LTD.",   # OSD to CMD
    222: "MAHARASHTRA STATE ELECTRICITY DISTRIBUTION CO. LTD.",   # Site Incharge-OMS
    599: "Trombay",                                                # VP Elect & Instrument
    700: "Western Region Load Despatch Centre",                    # Regional Manager SS
    701: "Western Region Load Despatch Centre",                    # Regional Manager Line
    706: "Western Region Load Despatch Centre",                    # Chief Manager
}
# ID 600 "Switchyard Desk (Operations)" – not a person, delete it

EMPLOYEE_DELETE = {600}   # Switchyard Desk – not a person


# ─── STEP 4: Remove exact duplicate employee rows ────────────────────────────

DUPLICATE_KEEP = {
    # Keep the LOWER id, delete the higher-id duplicate
    # (Dilesh Solanki – Daman and Diu)
    # (Santosh Satpute – Kalwa)
    # (Stephen Fernandes – Goa)
    # (Lucas Joao – Goa)
    # (Milind Ingle – Daman and Diu)
}


# ─── MAIN ────────────────────────────────────────────────────────────────────

def run():
    print("Loading Word document …")
    doc = Document(DOC_FILE)
    phone_map, name_map = build_word_phone_map(doc)
    print(f"Built phone map with {len(phone_map)} organisation keys.\n")

    with app.app_context():

        # ── Fix 1: Control rooms with blank phone numbers ────────────────────
        print("=" * 60)
        print("FIX 1 – Control rooms with missing phone numbers")
        print("=" * 60)

        blank_ctrl = DirectoryNumber.query.filter(
            DirectoryNumber.category == "Control Room",
            (DirectoryNumber.phone_number == None) |
            (DirectoryNumber.phone_number == ""),
        ).all()

        fixed_ctrl = 0
        for cr in blank_ctrl:
            # Manual fixes first
            if cr.id in MANUAL_PHONE_FIX:
                cr.phone_number = MANUAL_PHONE_FIX[cr.id]
                print(f"  FIXED (manual) ID {cr.id}: {cr.name} <- {cr.phone_number}")
                fixed_ctrl += 1
                continue

            # Try to find in Word doc by org name
            key = org_key(cr.organization or "")
            if key and key in phone_map and phone_map[key]:
                phones = ", ".join(phone_map[key][:3])
                cr.phone_number = phones
                print(f"  FIXED (word)   ID {cr.id}: {cr.name:<35} org=[{cr.organization}] <- {phones}")
                fixed_ctrl += 1
            else:
                known_blank = any(
                    k.lower() in (cr.organization or "").lower()
                    for k in KNOWN_NO_PHONE_ORGS
                )
                if known_blank:
                    print(f"  SKIP  (no data in source) ID {cr.id}: {cr.name} | {cr.organization}")
                else:
                    print(f"  MISS  ID {cr.id}: {cr.name} | org=[{cr.organization}]  (not found in doc)")

        db.session.commit()
        print(f"\n  Control room phones fixed: {fixed_ctrl}\n")

        # ── Fix 2: Employees without organization ────────────────────────────
        print("=" * 60)
        print("FIX 2 – Employees with no organization assigned")
        print("=" * 60)

        fixed_org = 0
        for emp_id, org_name in EMPLOYEE_ORG_FIX.items():
            emp = db.session.get(Employee, emp_id)
            if not emp:
                print(f"  SKIP emp ID {emp_id} – not found in DB")
                continue

            # Find or fuzzy-match the organization
            org = Organization.query.filter(
                Organization.organization_name.ilike(f"%{org_name[:30]}%")
            ).first()

            if org:
                emp.organization_id = org.id
                print(f"  FIXED emp ID {emp_id} [{emp.employee_name}] -> org=[{org.organization_name}]")
                fixed_org += 1
            else:
                print(f"  MISS  emp ID {emp_id} [{emp.employee_name}]: org '{org_name}' not found in DB")

        # Delete non-person rows
        deleted_emp = 0
        for emp_id in EMPLOYEE_DELETE:
            emp = db.session.get(Employee, emp_id)
            if emp:
                print(f"  DELETE emp ID {emp_id} [{emp.employee_name}] – not a real person")
                db.session.delete(emp)
                deleted_emp += 1

        db.session.commit()
        print(f"\n  Employee org assignments fixed: {fixed_org}")
        print(f"  Non-person rows deleted: {deleted_emp}\n")

        # ── Fix 3: Remove exact duplicate employees ──────────────────────────
        print("=" * 60)
        print("FIX 3 – Remove duplicate employee rows")
        print("=" * 60)

        from sqlalchemy import func
        dup_groups = (
            db.session.query(
                Employee.employee_name,
                Employee.organization_id,
                func.count(Employee.id).label("cnt"),
                func.min(Employee.id).label("keep_id"),
                func.max(Employee.id).label("drop_id"),
            )
            .group_by(Employee.employee_name, Employee.organization_id)
            .having(func.count(Employee.id) > 1)
            .all()
        )

        deleted_dup = 0
        for g in dup_groups:
            keep = db.session.get(Employee, g.keep_id)
            drop = db.session.get(Employee, g.drop_id)
            if drop:
                org = keep.organization.organization_name if keep.organization else "—"
                print(f"  DELETE dup ID {g.drop_id} [{g.employee_name}] in [{org}] "
                      f"(keeping ID {g.keep_id})")
                db.session.delete(drop)
                deleted_dup += 1

        db.session.commit()
        print(f"\n  Duplicate employees removed: {deleted_dup}\n")

        # ── Fix 4: Clean up org names that look like addresses ───────────────
        print("=" * 60)
        print("FIX 4 – Organization name / address issues")
        print("=" * 60)

        # ID 122: org name starts with comma (address was stored as name)
        from models.organization import Organization as Org
        bad_name_org = db.session.get(Org, 122)
        if bad_name_org:
            old = bad_name_org.organization_name
            # The org name is an address line; the real org should have been determined
            # from context. This was Adani / some Mumbai corp. Mark it for admin review.
            print(f"  WARN org ID 122 has name starting with comma: [{old[:80]}]")
            print(f"       Leaving for manual admin review.\n")

        # IDs 194, 196, 200, 202 – all look like addresses for Grid Controller of India
        grid_ctrl_addr = "C-105, Anand Niketan, New Delhi - 110 021"
        addr_orgs = [194, 196, 200, 202]
        for oid in addr_orgs:
            o = db.session.get(Org, oid)
            if o:
                print(f"  WARN org ID {oid} name looks like address: [{o.organization_name[:60]}]")
                print(f"       Leaving for manual admin review.\n")

        # ── Summary ──────────────────────────────────────────────────────────
        print("=" * 60)
        print("FINAL COUNTS")
        print("=" * 60)

        from models.employee import Employee as Emp
        from sqlalchemy import or_

        n_emps  = Emp.query.count()
        n_heads = Emp.query.filter_by(is_utility_head=True).count()
        no_org  = Emp.query.filter(Emp.organization_id == None).count()
        no_phone = Emp.query.filter(
            or_(Emp.mobile_phone == None, Emp.mobile_phone == ""),
            or_(Emp.office_phone == None, Emp.office_phone == ""),
        ).count()
        blank_cr = DirectoryNumber.query.filter(
            DirectoryNumber.category == "Control Room",
            or_(
                DirectoryNumber.phone_number == None,
                DirectoryNumber.phone_number == "",
            ),
        ).count()

        print(f"\n  Total employees          : {n_emps}")
        print(f"  Employees without org    : {no_org}")
        print(f"  Emps with no phone at all: {no_phone}")
        print(f"  Control rooms no phone   : {blank_cr}")
        print(f"  Utility heads            : {n_heads}")
        print()
        print("Done. Refresh the web app to see updated data.")


if __name__ == "__main__":
    run()
