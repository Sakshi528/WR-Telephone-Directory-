"""
Imports all organizations, employees, control rooms, and switchyards
from 'uploads/WR_DB_Ready_Final_Verified.xlsx' into the database.

Replaces all existing orgs, employees, CRs, and switchyards.
WRLDC employees are NOT touched — run import_employees_excel.py separately.

Delete + insert run inside a single database transaction: if anything fails
partway through, the whole thing is rolled back rather than leaving the
database with data deleted and nothing reinserted.

Usage:
  python import_from_excel_db.py           # dry run
  python import_from_excel_db.py --apply   # commit to DB
"""

import re
import sys
import openpyxl
from collections import defaultdict

from app import app
from models import db
from models.organization import Organization
from models.employee import Employee
from models.department import Department
from models.directory_number import DirectoryNumber
from models.emergency_contact import EmergencyContact
from utils.logger import get_logger

logger = get_logger(__name__)

EXCEL_PATH = "uploads/WR_DB_Ready_Final_Verified.xlsx"

# Top-level org owned by import_employees_excel.py — never touched by this script.
WRLDC_ORG_NAME = "Western Region Load Despatch Centre"

# emp_id → suborg name overrides (Excel has wrong or missing suborg_id for these)
MANUAL_SUBORG_MAP = {
    "State Load Despatch Centre, MSETCL": list(range(72, 87)) + [88],
    "Maharashtra State Electricity Distribution Co. Ltd.": [90, 91, 92],
    # NTPC Gadarwara/Khargone employees have org_id=12 but suborg_id=None in Excel
    "NTPC Gadarwara": [226, 227, 228, 229, 230],
    "NTPC Khargone":  [231, 232, 233, 234],
    # emp_id 87 (Vaishali Pazare) has suborg_id=35 in Excel (State Load Despatch
    # Centre, MSETCL / Kalwa, Thane-Belapur), which is wrong — verified against
    # 'uploads/Western Region Phone Directory 2025 Main_Telephone.docx', she is
    # listed under Area Load Despatch Centre, MSETCL, Ambazari, Nagpur (suborg_id=36).
    "Area Load Despatch Centre, MSETCL,": [87],
}
# emp_ids to skip entirely (junk: phone-as-name, headers, hospitals, emergency lines)
# emp_ids 87 and 89 were previously skipped here; verified against the Word
# directory to be genuine employees (Vaishali Pazare, Dinesh Agrawal) and removed —
# see MANUAL_SUBORG_MAP above for 87's suborg correction; 89's existing suborg_id
# (37, Maharashtra State Electricity Distribution Co. Ltd.) was already correct.
SKIP_EMP_IDS = set(range(754, 800)) | {791, 793, 797}

# Switchyard rows with broken org/suborg links in Excel — manually corrected.
# For these rows the cr_label col holds the phone number and desig holds the email;
# all standard phone/email cols are NULL.  key = Excel switchyard id (rid).
MANUAL_SW_MAP = {
    13: ("SASAN Power Limited", "Switchyard"),
    14: ("SASAN Power Limited", "Corporate Office"),
}


def clean(v):
    return re.sub(r"\s+", " ", str(v or "").replace("\xa0", " ")).strip()


def strip_serial(name):
    """Remove leading serial numbers: '1. Indira Sagar' → 'Indira Sagar'."""
    return re.sub(r"^\d+\.\s*", "", name).strip()


_STATE_PREFIXES = [
    "MAHARASHTRA", "GUJARAT", "MADHYA PRADESH", "CHHATTISGARH",
    "GOA", "RAJASTHAN", "UTTAR PRADESH", "KARNATAKA", "NTPC",
]

def normalize_suborg_name(raw):
    """Strip leading state/org prefix then serial: 'MAHARASHTRA 17. Khopoli & Bhivpuri' → 'Khopoli & Bhivpuri'."""
    s = raw.strip()
    for prefix in _STATE_PREFIXES:
        if s.upper().startswith(prefix):
            s = s[len(prefix):].strip()
            break
    return strip_serial(s)


def phones_joined(*vals):
    return " / ".join(clean(v) for v in vals if clean(v))


def run(apply: bool):
    logger.info("=" * 70)
    logger.info("WR_DB_Ready Excel Import")
    logger.info("DRY RUN" if not apply else "APPLY MODE — changes will be committed")
    logger.info("=" * 70)

    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True)

    def sheet_rows(name):
        return list(wb[name].iter_rows(values_only=True))[1:]

    orgs_rows   = sheet_rows("organizations")
    suborg_rows = sheet_rows("sub_organizations")
    emp_rows    = sheet_rows("employees")
    phone_rows  = sheet_rows("phone_numbers")
    email_rows  = sheet_rows("email_addresses")
    cr_rows     = sheet_rows("control_rooms")
    sw_rows     = sheet_rows("switchyards")

    # ── Phone/email lookups ──────────────────────────────────────────────────
    emp_phones = defaultdict(lambda: defaultdict(list))
    for _, emp_id, ptype, number in phone_rows:
        if emp_id and ptype and number:
            emp_phones[emp_id][clean(ptype)].append(clean(number))

    emp_emails = defaultdict(list)
    for _, emp_id, email in email_rows:
        if emp_id and email:
            emp_emails[emp_id].append(clean(email))

    logger.info(
        "Source data: %d organizations, %d sub-organizations, %d employees, "
        "%d control rooms, %d switchyards",
        len(orgs_rows), len(suborg_rows), len(emp_rows), len(cr_rows), len(sw_rows),
    )

    if not apply:
        logger.info("Dry run complete — rerun with --apply to commit")
        return

    with app.app_context():
        try:
            # ── DELETE existing KMP data ─────────────────────────────────────
            logger.info("Deleting existing KMP data...")
            kmp_ids = [e.id for e in Employee.query.filter_by(is_kmp=True).all()]
            if kmp_ids:
                EmergencyContact.query.filter(
                    EmergencyContact.employee_id.in_(kmp_ids)
                ).delete(synchronize_session=False)
            Employee.query.filter_by(is_kmp=True).delete(synchronize_session=False)
            DirectoryNumber.query.filter(
                DirectoryNumber.category.in_(["Control Room", "Switchyard"])
            ).delete()

            wrldc_org = Organization.query.filter_by(organization_name=WRLDC_ORG_NAME).first()
            org_delete_q = Organization.query
            if wrldc_org:
                org_delete_q = org_delete_q.filter(Organization.id != wrldc_org.id)
            org_delete_q.delete(synchronize_session=False)
            logger.info("Existing KMP data deleted (not yet committed).")

            # ── Build parent-org name map ────────────────────────────────────
            parent_map = {}
            for row in orgs_rows:
                oid, name, address, state = row
                parent_map[oid] = clean(name)

            # ── INSERT Organizations ─────────────────────────────────────────
            name_to_org = {}

            def get_or_insert_org(name, address, region, parent_id=None):
                if name in name_to_org:
                    return name_to_org[name]
                o = Organization(
                    organization_name=name,
                    address=address or None,
                    region=region or None,
                    parent_id=parent_id,
                )
                db.session.add(o)
                db.session.flush()
                name_to_org[name] = o
                return o

            org_obj_by_parent_id = {}
            for row in orgs_rows:
                oid, name, address, state = row
                name = clean(name)
                if not name:
                    continue
                # Parent orgs have no parent_id (top-level)
                o = get_or_insert_org(name, clean(address), clean(state))
                org_obj_by_parent_id[oid] = o

            org_obj_by_suborg_id = {}
            suborg_name_map = {}   # suborg_id → clean name (for CR matching)
            for row in suborg_rows:
                sid, parent_oid, name, address, state = row
                name = strip_serial(clean(name))   # strip "1. " prefixes
                if not name:
                    continue
                parent_name = parent_map.get(parent_oid, "")
                parent_org  = org_obj_by_parent_id.get(parent_oid)
                o = get_or_insert_org(
                    name,
                    clean(address),
                    parent_name or clean(state),   # region unchanged (backward compat)
                    parent_id=parent_org.id if parent_org else None,
                )
                org_obj_by_suborg_id[sid] = o
                suborg_name_map[sid] = name

            db.session.flush()
            logger.info(
                "Created %d parent orgs + %d sub-orgs",
                len(org_obj_by_parent_id), len(org_obj_by_suborg_id),
            )

            # ── Build reverse lookup: suborg clean name → org obj (for manual map)
            suborg_name_to_obj = {o.organization_name: o for o in org_obj_by_suborg_id.values()}

            # ── Pre-compute manual emp_id → org_obj ─────────────────────────
            manual_emp_map = {}
            for suborg_name, emp_ids in MANUAL_SUBORG_MAP.items():
                org_obj = suborg_name_to_obj.get(suborg_name)
                if org_obj:
                    for eid in emp_ids:
                        manual_emp_map[eid] = org_obj
                else:
                    logger.warning("suborg not found for manual map: %s", suborg_name)

            # Build sorted list of known org names (longest first) for text matching.
            # Must be built here — before employee import — so designation-based
            # sub-org resolution works for employees whose suborg_id is missing/wrong.
            known_org_names = sorted(name_to_org.keys(), key=len, reverse=True)

            def org_by_name_in(text):
                """Return the longest known org name found anywhere in text."""
                t = text.lower()
                for known in known_org_names:
                    if known.lower() in t:
                        return known
                return None

            # ── INSERT Employees (KMP) ───────────────────────────────────────
            emp_inserted = 0
            emp_skipped  = 0
            for row in emp_rows:
                emp_id, org_id, suborg_id, name, designation = row
                name = clean(name)

                if emp_id in SKIP_EMP_IDS or not name:
                    emp_skipped += 1
                    continue

                # Resolve org:
                # 1. Manual override map (highest priority)
                # 2. suborg_id FK from Excel (when the ID exists in sub_organizations)
                # 3. Designation text — catches employees whose suborg_id is wrong/missing
                #    but whose designation names a recognisable station (e.g. "EE Pench HPS")
                # 4. Parent org fallback (org_id FK)
                desig_text = clean(designation)
                if emp_id in manual_emp_map:
                    org_obj = manual_emp_map[emp_id]
                elif suborg_id and suborg_id in org_obj_by_suborg_id:
                    org_obj = org_obj_by_suborg_id[suborg_id]
                elif desig_text and org_by_name_in(desig_text):
                    org_obj = name_to_org.get(org_by_name_in(desig_text))
                elif org_id and org_id in org_obj_by_parent_id:
                    org_obj = org_obj_by_parent_id[org_id]
                else:
                    emp_skipped += 1
                    continue

                phones = emp_phones.get(emp_id, {})
                emails = emp_emails.get(emp_id, [])

                emp = Employee(
                    employee_name   = name,
                    designation     = clean(designation) or None,
                    organization_id = org_obj.id,
                    office_phone    = " / ".join(phones.get("office", [])) or None,
                    mobile_phone    = " / ".join(phones.get("mobile", [])) or None,
                    residence_phone = " / ".join(phones.get("residence", [])) or None,
                    email           = emails[0] if emails else None,
                    is_kmp          = True,
                )
                db.session.add(emp)
                emp_inserted += 1

            db.session.flush()
            logger.info("Inserted %d KMP employees (%d skipped)", emp_inserted, emp_skipped)

            # ── INSERT Control Rooms & Switchyards ───────────────────────────
            def insert_directory_numbers(rows, category):
                count = 0
                for row in rows:
                    (rid, org_id, suborg_id, org_name, suborg_name,
                     label, desig, addr, state,
                     p1, p2, p3, p4, fax, m1, m2, e1, e2) = row

                    # Manual override: rows where Excel has wrong org and phone/email
                    # are stored in non-standard columns (cr_label and designation).
                    if category == "Switchyard" and rid in MANUAL_SW_MAP:
                        sw_org_name, sw_label = MANUAL_SW_MAP[rid]
                        dn_org_obj = name_to_org.get(sw_org_name)
                        dn = DirectoryNumber(
                            name            = sw_label,
                            organization    = sw_org_name,
                            organization_id = dn_org_obj.id if dn_org_obj else None,
                            phone_number    = clean(label) or None,   # phone is in cr_label col
                            email           = clean(desig) or None,   # email is in designation col
                            category        = category,
                        )
                        db.session.add(dn)
                        count += 1
                        continue

                    # Resolve display org: suborg table → name-in-text match → parent org
                    if suborg_id and suborg_id in suborg_name_map:
                        display_org = suborg_name_map[suborg_id]
                    elif clean(suborg_name):
                        normalized = normalize_suborg_name(clean(suborg_name))
                        matched = org_by_name_in(normalized) or org_by_name_in(clean(suborg_name))
                        display_org = matched if matched else clean(org_name)
                    else:
                        # Try extracting sub-org from the label (e.g. "Ukai Thermal Control room" → "UKAI")
                        matched = org_by_name_in(clean(label)) if label else None
                        display_org = matched if matched else clean(org_name)

                    phone = phones_joined(p1, p2, p3, p4, m1, m2, fax)
                    email = " / ".join(filter(None, [clean(e1), clean(e2)])) or None

                    # Resolve the org FK from the name we just determined
                    dn_org_obj = name_to_org.get(display_org)

                    dn = DirectoryNumber(
                        name            = clean(label) or category,
                        organization    = display_org,             # TEXT kept (backward compat)
                        organization_id = dn_org_obj.id if dn_org_obj else None,
                        phone_number    = phone or None,
                        email           = email,
                        category        = category,
                    )
                    db.session.add(dn)
                    count += 1
                return count

            cr_count = insert_directory_numbers(cr_rows, "Control Room")
            sw_count = insert_directory_numbers(sw_rows, "Switchyard")
            logger.info("Inserted %d control rooms, %d switchyards", cr_count, sw_count)

            db.session.commit()

            logger.info("=" * 70)
            logger.info(
                "DONE — Orgs: %d, KMP employees: %d, Control Rooms: %d, Switchyards: %d",
                len(name_to_org), emp_inserted, cr_count, sw_count,
            )
            logger.info("=" * 70)

        except Exception:
            logger.exception("Import failed — rolling back all changes.")
            db.session.rollback()
            raise


if __name__ == "__main__":
    run(apply="--apply" in sys.argv)
