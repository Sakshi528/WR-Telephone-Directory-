import logging
import re
import sys
from datetime import datetime

from docx import Document

from app import app
from models import db
from models.department import Department
from models.directory_number import DirectoryNumber
from models.employee import Employee
from models.organization import Organization
from models.update_request import UpdateRequest


DOC_FILE = "uploads/Western Region Phone Directory 2025 Main_Telephone.docx"
LOG_FILE = "import_skipped.log"

# ─── LOGGING ─────────────────────────────────────────────────────────────────

logger = logging.getLogger("import_skip")
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler(LOG_FILE, mode="w", encoding="utf-8")
fh.setFormatter(logging.Formatter("%(message)s"))
logger.addHandler(fh)
ch = logging.StreamHandler(sys.stdout)
ch.setFormatter(logging.Formatter("%(message)s"))
logger.addHandler(ch)

def log_skip(t, r, reason, values):
    preview = " | ".join(str(v)[:40] for v in values if v)[:120]
    logger.info(f"[T{t:3d}|R{r:3d}] SKIP: {reason:<40} | {preview}")

def log_insert(t, r, kind, name):
    logger.debug(f"[T{t:3d}|R{r:3d}] {kind:<15} -> {name}")


# ─── TEXT UTILITIES ──────────────────────────────────────────────────────────

def clean(value):
    return re.sub(r"\s+", " ", (value or "").replace("\xa0", " ")).strip()

def normalize(value):
    return re.sub(r"[^a-z0-9]+", " ", clean(value).lower()).strip()

def unique_cell_texts(row):
    seen, result = set(), []
    for cell in row.cells:
        if id(cell) not in seen:
            seen.add(id(cell))
            result.append(clean(cell.text))
    return result

def joined(values):
    seen, result = [], []
    for v in values:
        v = clean(v)
        if v and v not in seen:
            seen.append(v)
            result.append(v)
    return " / ".join(result)

def has_email(value):
    return "@" in value

def para_text(el):
    texts, seen_t = [], set()
    for t in el.iter():
        if t.tag.endswith("}t") and t.text:
            txt = t.text.strip()
            if txt and txt not in seen_t:
                seen_t.add(txt)
                texts.append(txt)
    return clean(" ".join(texts))

def looks_like_phone(value):
    """True if value is mostly digits/phone-like."""
    digits = re.sub(r"[^0-9]", "", value)
    return len(digits) >= 6 and len(digits) >= len(value) * 0.4


# ─── ROW CLASSIFICATION ──────────────────────────────────────────────────────

SWITCHYARD_PATTERNS = [
    "switchyard", "switch yard", "hvdc switch",
    "400 kv switch", "400kv gis",
]

CONTROL_ROOM_PATTERNS = [
    "control room", "control  room", "controlroom",
    "aldc control", "sldc control", "rldc control",
    "sldc dnh", "s/s cr",
    "rtamc", "cpcc", "hotline",
    "shift manager office", "control room office",
    "power house", "ukai thermal control",
    "hydro control room", "mini hydro control",
    "plant operations control room",
    "central control room",
    "sic control room", "ccr control room",
    "process control room", "forecasting & scheduling control room",
    "rumsl unit", "site control room",
    "apml tiroda", "cgpl central",
    "sce, rbph", "sce, chph", "emc, control room",
    "sterlite operation control room",
]

def is_switchyard_row(name_value):
    low = name_value.lower().strip()
    return any(p in low for p in SWITCHYARD_PATTERNS)

def is_control_room_row(name_value):
    low = name_value.lower().strip()
    return any(p in low for p in CONTROL_ROOM_PATTERNS)


def extract_cr_phone_email(values, header_map, fields):
    phones = []
    emails = []

    for v in values:
        v = clean(v)

        if not v:
            continue

        if "@" in v:
            emails.append(v)
            continue

        digits = re.sub(r"\D", "", v)
        if len(digits) >= 6:
            phones.append(v)

    return " / ".join(phones), " / ".join(emails)


def is_address_line(text):
    stripped = text.strip()
    if re.match(r"^\d[\d\s]{4,9}$", stripped):
        return True
    low = text.lower()

    # Skip city-at-end heuristic if the string contains a 3+ letter ALL-CAPS
    # abbreviation — that signals an org name like "POWERGRID WRTS-I, Nagpur"
    # or "State Load Despatch Centre, MSETCL" rather than a plain address.
    has_org_abbrev = bool(re.search(r"\b[A-Z]{3,}\b", text))

    # State / city name at end of string → address component, not org name
    # (only when no org abbreviation present)
    if not has_org_abbrev and re.search(
        r"\b(gujarat|maharashtra|madhya pradesh|chhattisgarh|rajasthan|"
        r"goa|haryana|nagpur|surat|navi mumbai|mumbai|pune|thane|"
        r"ahmedabad|vadodara|indore|bhopal|raipur|gurugram|gurgaon|"
        r"delhi|kolkata|bengaluru|bangalore|hyderabad|kutch|khandwa|"
        r"jabalpur|korba|bharuch|ratnagiri|aurangabad|satara|raigad|"
        r"palghar|kevadiya|mandsaur|seoni|sirmour|tapi|mahisagar|"
        r"kheda|andheri|boisar|punasa|belapur|vasai|virar|kalyan|"
        r"dombivli|ulhasnagar|bhiwandi|vapi)\s*[,.]?\s*$",
        low,
    ):
        return True

    # State/city name followed by dash/hyphen + pin code is always an address
    if re.search(
        r"\b(gujarat|maharashtra|madhya pradesh|chhattisgarh|rajasthan|"
        r"goa|haryana|karnataka|telangana|andhra pradesh|uttar pradesh|"
        r"odisha|west bengal|kerala|delhi|new delhi|chandigarh)\s*[-–]\s*\d{3}",
        low,
    ):
        return True

    SIGNALS = [
        "p.o.", "p.o ", "dist.", "dist:", " dist ", "pin :", "pin code", "stpp", "stps",
        "village", "sector", "floor", "bhawan",
        "scope complex", "lodhi road", "ph no", "fax no", "email:",
        "bandra", "midc", "opp ", "near ", "plot no",
        "s g highway", "sg highway", "shantigram", "khodiyar",
        "race course", "race cource",
        "vidyut bhavan", "shakti bhawan", "danganiya",
        "132 kv", "400 kv", "765 kv", "substation compound",
        "road no", "road,", "nagar,", "nagar.",
        "jubilee hills", "pincode", ", india",
        "mob:", "mob :", "powai", "orchard avenue",
        "po:", "po :", "p.o:",
        " 401", " 495", " 486", " 382", " 390", " 380",
        " 452", " 413", " 487", " 492", " 451", " 496", " 482",
        " 391", " 394", " 393", " 396", " 395",
        " 400 0", " 400 07",
    ]
    return any(s in low for s in SIGNALS)


def is_junk_paragraph(text):
    low = text.lower()
    JUNK = [
        "overcoming barriers to performance",
        "in union there is strength",
        "working together for a great",
        "teamwork does not tolerate",
        "neither gold nor diamonds",
        "we = power", "city control- 100",
        "city fire control", "ambulance- 102",
        "helpline number- 1926",
    ]
    return any(j in low for j in JUNK)


# ─── ORG NAME CLEANING ───────────────────────────────────────────────────────

def clean_org_name(raw):
    raw = re.sub(r"^[\d\s.–-]+", "", raw).strip()
    # Fix OCR spaces inside abbreviations: "N TPC" → "NTPC", "NTPC G adarwara" → "NTPC Gadarwara"
    raw = re.sub(r"\b([A-Z])\s+([A-Z]{2,})\b", r"\1\2", raw)  # "N TPC" → "NTPC"
    raw = re.sub(r"\b([A-Z])\s+([a-z])", r"\1\2", raw)         # "G adarwara" → "Gadarwara"
    raw = re.sub(r"\b([A-Z][a-z]+)\s+([a-z]{2,})",
                 lambda m: m.group(1) + m.group(2), raw)
    return raw.strip()


def build_table_org_map(document):
    body = list(document.element.body)
    table_els = [el for el in body if el.tag.endswith("}tbl")]
    result = {}

    for t_idx, tbl_el in enumerate(table_els):
        pos = next((i for i, el in enumerate(body) if el is tbl_el), None)
        if pos is None:
            continue
        org_name, address = "", ""
        for i in range(pos - 1, max(pos - 50, -1), -1):
            el = body[i]
            if el.tag.endswith("}tbl"):
                break
            if not el.tag.endswith("}p"):
                continue
            raw = para_text(el)
            if not raw or len(raw) < 3:
                continue
            if is_junk_paragraph(raw):
                continue
            if is_address_line(raw):
                if not address:
                    address = raw
            else:
                if not org_name:
                    org_name = clean_org_name(raw)
            if org_name and address:
                break
        result[t_idx] = {"org": org_name, "address": address}
    return result


def org_from_table_row0(table, header_index):
    for row_idx in range(header_index - 1, -1, -1):
        vals = unique_cell_texts(table.rows[row_idx])
        non_empty = [v for v in vals if v]
        if len(non_empty) == 1:
            c = non_empty[0]
            if (not is_address_line(c) and not is_junk_paragraph(c)
                    and len(c) > 4 and not re.match(r"^[\d\s.]+$", c)):
                return clean_org_name(c)
    return ""


# ─── HEADER / FIELD DETECTION ────────────────────────────────────────────────

def looks_like_header(values):
    text = " ".join(normalize(v) for v in values)
    return "name" in text and any(
        k in text for k in ("designation", "mobile", "telephone",
                            "email", "mail", "land line", "contact number",
                            "activity")
    )


def field_for_header(header, index):
    h = normalize(header)
    if re.match(r"^(sl|sr|s)\s*(no|num)", h) or h in {"ssss", "nos"}:
        return None
    if "station name" in h or "name of hospital" in h:
        return None
    if "activity" in h or "department" in h:
        return "department"
    if "name" in h:
        return "employee_name"
    if "designation" in h:
        return "designation"
    if h == "region" or h.startswith("region"):
        return "region"
    if "location" in h:
        return "location"
    if "mobile" in h or "contact number" in h or "emergency number" in h:
        return "mobile_phone"
    if "email" in h or "mail" in h:
        return "email"
    if "telephone r" in h or "residence" in h or "land line 2" in h:
        return "residence_phone"
    if "telephone f" in h or h == "fax":
        return None
    if ("telephone o" in h or "telephone no" in h or "land line 1" in h
            or h in ("office", "telephone")):
        return "office_phone"
    return None


# ─── DB HELPERS ──────────────────────────────────────────────────────────────

def get_or_create_org(org_name, address=""):
    org_name = clean(org_name)
    if not org_name:
        return None
    item = Organization.query.filter_by(organization_name=org_name).first()
    if item:
        if address and not getattr(item, "address", None):
            try:
                item.address = clean(address)
                db.session.flush()
            except Exception:
                pass
        return item
    kwargs = {"organization_name": org_name}
    try:
        kwargs["address"] = clean(address)
    except Exception:
        pass
    item = Organization(**kwargs)
    db.session.add(item)
    db.session.flush()
    return item


def get_or_create(model, field_name, value):
    value = clean(value)
    if not value:
        return None
    item = model.query.filter_by(**{field_name: value}).first()
    if item:
        return item
    item = model(**{field_name: value})
    db.session.add(item)
    db.session.flush()
    return item


# ─── EXTRACT FIELDS ──────────────────────────────────────────────────────────

def extract_fields(values, header_map, fallback_organization=""):
    fields = {
        "employee_name": "", "designation": "", "department": "",
        "organization": fallback_organization, "region": "", "location": "",
        "office_phone": "", "residence_phone": "", "mobile_phone": "", "email": "",
    }
    for field, indexes in header_map.items():
        fields[field] = joined(
            values[index] for index in indexes if index < len(values)
        )
    if fields["organization"] == fields["employee_name"]:
        fields["organization"] = fallback_organization
    return fields


def insert_control_room(t_idx, r_idx, name, phone, email, organization,
                        dry_run, inserted_count):
    if dry_run:
        log_insert(t_idx, r_idx, "CONTROL ROOM",
                   f"{name} | {phone[:30]} | [{organization}]")
        return inserted_count + 1

    existing = DirectoryNumber.query.filter_by(
        name=name,
        phone_number=phone,
        organization=organization
    ).first()
    if existing:
        log_skip(t_idx, r_idx, "duplicate control room",
                 [name, phone, organization])
        return inserted_count

    dn = DirectoryNumber(
        name=name,
        phone_number=phone,
        category="Control Room",
        organization=organization,
    )
    try:
        dn.email = email
    except Exception:
        pass
    db.session.add(dn)
    log_insert(t_idx, r_idx, "CONTROL ROOM",
               f"{name} | {phone[:30]} | [{organization}]")
    return inserted_count + 1


def insert_switchyard(t_idx, r_idx, name, phone, email, organization,
                      dry_run, inserted_count):
    if dry_run:
        log_insert(t_idx, r_idx, "SWITCHYARD",
                   f"{name} | {phone[:30]} | [{organization}]")
        return inserted_count + 1

    existing = DirectoryNumber.query.filter_by(
        name=name,
        phone_number=phone,
        organization=organization
    ).first()
    if existing:
        log_skip(t_idx, r_idx, "duplicate switchyard",
                 [name, phone, organization])
        return inserted_count

    dn = DirectoryNumber(
        name=name,
        phone_number=phone,
        category="Switchyard",
        organization=organization,
    )
    try:
        dn.email = email
    except Exception:
        pass
    db.session.add(dn)
    log_insert(t_idx, r_idx, "SWITCHYARD",
               f"{name} | {phone[:30]} | [{organization}]")
    return inserted_count + 1


# ─── MAIN IMPORT ─────────────────────────────────────────────────────────────

def import_employees(replace=False, dry_run=False):
    document = Document(DOC_FILE)
    emp_inserted = 0
    cr_inserted = 0
    sw_inserted = 0

    skip_counts = {
        "empty row": 0,
        "no header in table": 0,
        "no name column": 0,
        "sub-section heading": 0,
        "blank name cell": 0,
        "duplicate employee": 0,
        "junk/separator": 0,
    }

    if replace and not dry_run:
        if UpdateRequest.query.filter_by(status="Pending").count():
            raise RuntimeError("Cannot replace while pending update requests exist.")
        Employee.query.delete()
        DirectoryNumber.query.filter_by(category="Control Room").delete()
        DirectoryNumber.query.filter_by(category="Switchyard").delete()
        Department.query.delete()
        Organization.query.delete()
        db.session.commit()

    table_org_map = build_table_org_map(document)

    logger.info("=" * 80)
    logger.info(f"Import: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} "
                f"| Mode: {'DRY RUN' if dry_run else ('REPLACE' if replace else 'NORMAL')}")
    logger.info("=" * 80)

    for t_idx, table in enumerate(document.tables):

        # ── find header row ──────────────────────────────────────────────────
        header_index = None
        for idx, row in enumerate(table.rows[:8]):
            if looks_like_header(unique_cell_texts(row)):
                header_index = idx
                break

        if header_index is None:
            for r_idx, row in enumerate(table.rows):
                values = unique_cell_texts(row)
                if any(values):
                    log_skip(t_idx, r_idx, "no header found in table", values)
                    skip_counts["no header in table"] += 1
            continue

        # ── build header map ─────────────────────────────────────────────────
        headers = unique_cell_texts(table.rows[header_index])
        header_map = {}
        for idx, header in enumerate(headers):
            field = field_for_header(header, idx)
            if field:
                header_map.setdefault(field, []).append(idx)

        if "employee_name" not in header_map:
            for r_idx, row in enumerate(table.rows):
                values = unique_cell_texts(row)
                if any(values):
                    log_skip(t_idx, r_idx, "no name column in header", values)
                    skip_counts["no name column"] += 1
            continue

        # ── resolve org + address ────────────────────────────────────────────
        # Prefer org name from inside the table (row0 before header) over the
        # external body paragraph, because multi-org tables embed the specific
        # org name as a merged cell row above their own header.
        meta = table_org_map.get(t_idx, {})
        fallback_org = org_from_table_row0(table, header_index) or meta.get("org", "")
        org_address = meta.get("address", "")

        if fallback_org and not dry_run:
            get_or_create_org(fallback_org, org_address)

        current_organization = fallback_org
        current_address = org_address

        # ── data rows ────────────────────────────────────────────────────────
        for r_idx, row in enumerate(table.rows[header_index + 1:],
                                    start=header_index + 1):
            values = unique_cell_texts(row)

            if not any(values):
                log_skip(t_idx, r_idx, "empty row", values)
                skip_counts["empty row"] += 1
                continue

            non_empty = [v for v in values if v]

            # ── sub-section heading ──────────────────────────────────────────
            if len(non_empty) == 1 and not has_email(non_empty[0]):
                heading = non_empty[0]
                low = heading.lower()
                if (not is_address_line(heading)
                        and not is_junk_paragraph(heading)
                        and not re.match(r"^[\d\s.]+$", heading)
                        and not re.match(r"^(sl|sr)\s*no", low)):
                    new_org = clean_org_name(heading)
                    if new_org:
                        current_organization = new_org
                        current_address = ""
                        if not dry_run:
                            get_or_create_org(current_organization, "")
                    log_skip(t_idx, r_idx,
                             f"sub-heading → org='{current_organization}'",
                             values)
                    skip_counts["sub-section heading"] += 1
                else:
                    log_skip(t_idx, r_idx, "junk/separator", values)
                    skip_counts["junk/separator"] += 1
                continue

            # ── also handle multi-header rows (re-detect header) ─────────────
            if looks_like_header(values):
                # Update header map from this new header row
                header_map = {}
                for idx, header in enumerate(values):
                    field = field_for_header(header, idx)
                    if field:
                        header_map.setdefault(field, []).append(idx)
                log_skip(t_idx, r_idx, "sub-header row (updated map)", values)
                skip_counts["sub-section heading"] += 1
                continue

            # ── get name cell ────────────────────────────────────────────────
            name_indexes = header_map.get("employee_name", [])
            name_cell = (values[name_indexes[0]]
                         if name_indexes and name_indexes[0] < len(values)
                         else "")

            # Skip blank name
            if not name_cell:
                log_skip(t_idx, r_idx, "blank name cell", values)
                skip_counts["blank name cell"] += 1
                continue

            fields = extract_fields(values, header_map, current_organization)
            org_obj = None if dry_run else get_or_create_org(
                current_organization, current_address
            )

            # ── CONTROL ROOM → DirectoryNumber ──────────────────────────────
            # Check CR before switchyard: a cell like "400KV GIS Control room"
            # matches both patterns — "control room" should take priority.
            cr_name = (name_cell if is_control_room_row(name_cell)
                       else next((v for v in non_empty if is_control_room_row(v)), None))
            if cr_name:
                phone, email = extract_cr_phone_email(values, header_map, fields)
                cr_inserted = insert_control_room(
                    t_idx, r_idx,
                    name=cr_name,
                    phone=phone,
                    email=email,
                    organization=current_organization,
                    dry_run=dry_run,
                    inserted_count=cr_inserted
                )
                continue

            # ── SWITCHYARD → DirectoryNumber ────────────────────────────────
            # Also scan all cells in case the name is in col 0 (serial-number col)
            sw_name = (name_cell if is_switchyard_row(name_cell)
                       else next((v for v in non_empty if is_switchyard_row(v)), None))
            if sw_name:
                phone, email = extract_cr_phone_email(values, header_map, fields)
                sw_inserted = insert_switchyard(
                    t_idx, r_idx,
                    name=sw_name,
                    phone=phone,
                    email=email,
                    organization=current_organization,
                    dry_run=dry_run,
                    inserted_count=sw_inserted
                )
                continue

            # ── EMPLOYEE ─────────────────────────────────────────────────────
            if not fields["employee_name"]:
                log_skip(t_idx, r_idx, "blank name cell", values)
                skip_counts["blank name cell"] += 1
                continue

            if not dry_run:
                existing = Employee.query.filter_by(
                    employee_name=fields["employee_name"],
                    email=fields["email"],
                ).first()
                if existing:
                    log_skip(t_idx, r_idx, "duplicate employee",
                             [fields["employee_name"], fields["email"]])
                    skip_counts["duplicate employee"] += 1
                    continue

                department = get_or_create(
                    Department, "department_name", fields["department"]
                )
                employee = Employee(
                    employee_name=fields["employee_name"],
                    designation=fields["designation"],
                    department_id=department.id if department else None,
                    organization_id=org_obj.id if org_obj else None,
                    location=fields["location"],
                    region=fields["region"],
                    office_phone=fields["office_phone"],
                    residence_phone=fields["residence_phone"],
                    mobile_phone=fields["mobile_phone"],
                    email=fields["email"],
                )
                db.session.add(employee)

            log_insert(t_idx, r_idx, "EMPLOYEE",
                       f"{fields['employee_name']} [{current_organization}]")
            emp_inserted += 1

    if not dry_run:
        db.session.commit()

    # ── Summary ──────────────────────────────────────────────────────────────
    logger.info("")
    logger.info("=" * 80)
    logger.info("SUMMARY")
    logger.info(f"  Employees inserted    : {emp_inserted}")
    logger.info(f"  Control rooms inserted: {cr_inserted}")
    logger.info(f"  Switchyards inserted  : {sw_inserted}")
    total_skipped = sum(skip_counts.values())
    logger.info(f"\n  SKIPPED BREAKDOWN:")
    for reason, count in skip_counts.items():
        if count:
            logger.info(f"    {reason:<40}: {count}")
    logger.info(f"    {'TOTAL':<40}: {total_skipped}")
    logger.info(f"\n  Log: {LOG_FILE}")
    logger.info("=" * 80)

    return emp_inserted, cr_inserted, sw_inserted, skip_counts


# ─── ENTRY POINT ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    replace  = "--replace"  in sys.argv
    dry_run  = "--dry-run"  in sys.argv

    with app.app_context():
        emp_ins, cr_ins, sw_ins, skip_counts = import_employees(
            replace=replace, dry_run=dry_run
        )

    total_skipped = sum(skip_counts.values())
    print(f"\n[OK]  {emp_ins} employees imported")
    print(f"[CR]  {cr_ins} control room entries imported")
    print(f"[SW]  {sw_ins} switchyard entries imported")
    print(f"[--]  {total_skipped} rows skipped  (see {LOG_FILE} for details)")
