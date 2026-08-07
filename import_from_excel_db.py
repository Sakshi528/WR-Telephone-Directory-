"""Imports organizations, employees, control rooms, and switchyards from
'WR_DB_Ready_Final_Verified_v3.xlsx' into the database. Replaces all
existing orgs/employees/CRs/switchyards; WRLDC employees are not touched
(run import_employees_excel.py separately). Delete + insert run inside one
transaction, rolled back whole on failure.

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
from models.import_batch import ImportBatch
from models.organization_category import OrganizationCategory
from utils.logger import get_logger

logger = get_logger(__name__)

# Was 'uploads/WR_DB_Ready_Final_Verified.xlsx' -- that file no longer
# exists (superseded by the validation/correction/normalization pipeline
# that produced v3; see scripts/generate_validation_report.py,
# apply_corrections.py, normalize_formatting.py). Running with the old
# path raised FileNotFoundError on every invocation.
EXCEL_PATH = "WR_DB_Ready_Final_Verified_v3.xlsx"

# Top-level org owned by import_employees_excel.py — never touched by this script.
WRLDC_ORG_NAME = "Western Region Load Despatch Centre"

# emp_id → suborg name overrides (Excel has wrong or missing suborg_id for these)
MANUAL_SUBORG_MAP = {
    "State Load Despatch Centre, MSETCL": list(range(72, 87)) + [88],
    "Maharashtra State Electricity Distribution Co. Ltd.": [90, 91, 92],
    "NTPC Gadarwara": [226, 227, 228, 229, 230],
    "NTPC Khargone":  [231, 232, 233, 234],
    # emp_id 87 (Vaishali Pazare): Excel has suborg_id=35, but Word places her
    # under Area Load Despatch Centre, MSETCL, Ambazari, Nagpur (suborg_id=36)
    "Area Load Despatch Centre, MSETCL,": [87],
}
# junk emp_ids to skip entirely (phone-as-name, headers, hospitals, emergency lines)
SKIP_EMP_IDS = set(range(754, 800)) | {791, 793, 797}


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

    # phone_numbers/email_addresses now carry a reference_type column (added
    # for the wider correction pipeline, which also stores control_room/kmp/
    # hospital/utility_head/switchyard numbers here) -- was a plain (id, emp_id,
    # type, number) tuple when this script was first written, so unpacking
    # broke (ValueError: too many values to unpack) and, if merely re-padded
    # without filtering, would have silently cross-merged unrelated entities'
    # phones/emails whenever their reference_id happened to collide with an
    # employee_id. This script only ever joins onto the `employees` sheet, so
    # only 'employee'-typed rows are relevant here.
    emp_phones = defaultdict(lambda: defaultdict(list))
    for _, ref_type, emp_id, ptype, number in phone_rows:
        if ref_type == "employee" and emp_id and ptype and number:
            emp_phones[emp_id][clean(ptype)].append(clean(number))

    emp_emails = defaultdict(list)
    for _, ref_type, emp_id, email in email_rows:
        if ref_type == "employee" and emp_id and email:
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

            # organizations.category_id is NOT NULL (added after this script was
            # first written, for the Organization Type feature) -- this script
            # has no source of Type data, so newly (re)inserted orgs default to
            # the catch-all 'Others' category, same fallback
            # admin_routes._default_category_id() uses for the manual Add
            # Organization form. Re-run scripts/reclassify_organization_types.py
            # afterward to restore real classifications -- this import alone
            # cannot know them.
            others_category = OrganizationCategory.query.filter_by(category_name="Others").first()
            if not others_category:
                raise RuntimeError(
                    "organization_categories has no 'Others' row -- run "
                    "migrations 009-012 before importing."
                )

            parent_map = {}
            for row in orgs_rows:
                oid, name, address, state = row
                parent_map[oid] = clean(name)

            name_to_org = {}
            # The workbook's own organizations sheet includes an entry named
            # WRLDC_ORG_NAME (org_id 25) -- the preserved live WRLDC row above
            # was excluded from deletion, but nothing stopped this loop from
            # then trying to INSERT a second row with the same name, violating
            # organizations.organization_name's unique constraint. Seed
            # name_to_org with the preserved row so get_or_insert_org reuses
            # it instead, matching the docstring's own stated intent that this
            # script never touches WRLDC.
            if wrldc_org:
                name_to_org[WRLDC_ORG_NAME] = wrldc_org

            def get_or_insert_org(name, address, region, parent_id=None, state=None):
                if name in name_to_org:
                    return name_to_org[name]
                o = Organization(
                    organization_name=name,
                    address=address or None,
                    region=region or None,
                    state=state or None,
                    parent_id=parent_id,
                    category_id=others_category.id,
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
                o = get_or_insert_org(name, clean(address), clean(state), state=clean(state))
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
                    state=clean(state),
                )
                org_obj_by_suborg_id[sid] = o
                suborg_name_map[sid] = name

            db.session.flush()
            logger.info(
                "Created %d parent orgs + %d sub-orgs",
                len(org_obj_by_parent_id), len(org_obj_by_suborg_id),
            )

            suborg_name_to_obj = {o.organization_name: o for o in org_obj_by_suborg_id.values()}

            manual_emp_map = {}
            for suborg_name, emp_ids in MANUAL_SUBORG_MAP.items():
                org_obj = suborg_name_to_obj.get(suborg_name)
                if org_obj:
                    for eid in emp_ids:
                        manual_emp_map[eid] = org_obj
                else:
                    logger.warning("suborg not found for manual map: %s", suborg_name)

            # built before employee import so designation-based resolution works below
            known_org_names = sorted(name_to_org.keys(), key=len, reverse=True)

            def org_by_name_in(text):
                """Return the longest known org name found anywhere in text."""
                t = text.lower()
                for known in known_org_names:
                    if known.lower() in t:
                        return known
                return None

            # a blank-suborg_id row sharing a (name, designation) key with a
            # valid-suborg_id row is a shared multi-station role (e.g. "Head of
            # the Station" repeated across sites) -- skip rather than misattach
            key_has_valid_suborg = defaultdict(bool)
            key_has_blank_suborg = defaultdict(bool)
            for _eid, _oid, _sid, _dept, _name, _desig in emp_rows:
                _name = clean(_name)
                if not _name:
                    continue
                key = (_name.lower(), clean(_desig).lower())
                if _sid and _sid in org_obj_by_suborg_id:
                    key_has_valid_suborg[key] = True
                elif _sid is None:
                    key_has_blank_suborg[key] = True
            ambiguous_multi_station_keys = {
                key for key in key_has_blank_suborg
                if key_has_valid_suborg[key]
            }

            emp_inserted = 0
            emp_skipped  = 0
            emp_skipped_ambiguous = 0
            for row in emp_rows:
                emp_id, org_id, suborg_id, _department, name, designation = row
                name = clean(name)

                if emp_id in SKIP_EMP_IDS or not name:
                    emp_skipped += 1
                    continue

                # Resolve org, in priority order: valid suborg_id from Excel (a
                # repair_workbook.py pass can fix this, so it outranks the manual
                # map below) -> manual override map -> designation text naming a
                # known station -> skip if ambiguous shared multi-station role ->
                # parent org fallback.
                desig_text = clean(designation)
                key = (name.lower(), desig_text.lower())
                if suborg_id and suborg_id in org_obj_by_suborg_id:
                    org_obj = org_obj_by_suborg_id[suborg_id]
                elif emp_id in manual_emp_map:
                    org_obj = manual_emp_map[emp_id]
                elif desig_text and org_by_name_in(desig_text):
                    org_obj = name_to_org.get(org_by_name_in(desig_text))
                elif suborg_id is None and key in ambiguous_multi_station_keys:
                    emp_skipped_ambiguous += 1
                    continue
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
            logger.info(
                "Inserted %d KMP employees (%d skipped, %d skipped as ambiguous "
                "shared multi-station roles)",
                emp_inserted, emp_skipped, emp_skipped_ambiguous,
            )

            # detects rows where suborg_id conflicts with the row's own free-text
            # suborg_name -- a confirmed non-constant ID-shift error in this sheet
            suborg_id_by_name_lower = {v.lower(): k for k, v in suborg_name_map.items()}

            def insert_directory_numbers(rows, category):
                # control_rooms (12 cols) carries a designation + a second
                # phone/mobile slot; switchyards (9 cols) doesn't -- was
                # unpacked as one 18-column shape for both, which no longer
                # matches either sheet (ValueError: not enough values to
                # unpack on every row).
                count = 0
                for row in rows:
                    if category == "Control Room":
                        (rid, org_id, suborg_id, org_name, suborg_name,
                         label, desig, p1, p2, m1, m2, e1) = row
                    else:
                        (rid, org_id, suborg_id, org_name, suborg_name,
                         label, p1, m1, e1) = row
                        p2 = m2 = None

                    # Resolve display org: exact text match (if it disagrees with a
                    # present suborg_id) → suborg table → name-in-text match → parent org
                    own_text_sid = suborg_id_by_name_lower.get(clean(suborg_name).strip().lower())
                    if own_text_sid is not None and own_text_sid != suborg_id:
                        display_org = suborg_name_map[own_text_sid]
                    elif suborg_id and suborg_id in suborg_name_map:
                        display_org = suborg_name_map[suborg_id]
                    elif clean(suborg_name):
                        normalized = normalize_suborg_name(clean(suborg_name))
                        matched = org_by_name_in(normalized) or org_by_name_in(clean(suborg_name))
                        display_org = matched if matched else clean(org_name)
                    else:
                        # Try extracting sub-org from the label (e.g. "Ukai Thermal Control room" → "UKAI")
                        matched = org_by_name_in(clean(label)) if label else None
                        display_org = matched if matched else clean(org_name)

                    phone = phones_joined(p1, p2, m1, m2)
                    email = clean(e1) or None

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

            # First persisted record of this run -- Archive > Import History.
            # This whole pipeline is delete-then-reinsert, not a row-level
            # diff, so "updated" isn't a meaningful count here.
            db.session.add(ImportBatch(
                workbook_name=EXCEL_PATH, mode="APPLIED", status="SUCCESS",
                inserted_count=len(name_to_org) + emp_inserted + cr_count + sw_count,
                updated_count=0, skipped_count=emp_skipped + emp_skipped_ambiguous,
                summary=(
                    f"Orgs: {len(name_to_org)}, KMP employees: {emp_inserted} "
                    f"({emp_skipped} skipped, {emp_skipped_ambiguous} ambiguous), "
                    f"Control Rooms: {cr_count}, Switchyards: {sw_count}"
                ),
            ))
            db.session.commit()

        except Exception as exc:
            logger.exception("Import failed — rolling back all changes.")
            db.session.rollback()
            db.session.add(ImportBatch(
                workbook_name=EXCEL_PATH, mode="APPLIED", status="FAILED",
                summary=f"Import failed: {exc}",
            ))
            db.session.commit()
            raise


if __name__ == "__main__":
    run(apply="--apply" in sys.argv)
