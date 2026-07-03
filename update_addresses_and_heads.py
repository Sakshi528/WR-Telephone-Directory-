"""
update_addresses_and_heads.py
=============================================================
1. Updates organization addresses from the Word document.
2. Marks the most-senior employee per organization as utility head,
   using a designation-priority ranking so only true heads are marked.
=============================================================
Run: python update_addresses_and_heads.py
"""

import re
import sys

from docx import Document

from app import app
from models import db
from models.employee import Employee
from models.organization import Organization
from sqlalchemy import text


DOC_FILE = "uploads/Western Region Phone Directory 2025 Main_Telephone.docx"


# ── Designation seniority scoring ────────────────────────────────────────────
# Higher score = more senior.  Score 0 = not senior enough to be a utility head.

DESIGNATION_TIERS = [
    (100, ["chairman", "cmd", "c.m.d"]),
    (90,  ["executive director", "director general",
            " ed ", "ed(", "ed-", "ed,", "(ed)", "/ed"]),
    (85,  ["principal secretary", "secretary general"]),
    (80,  ["managing director", " md,", "m.d.", " md ", "(md)",
            "secretary (power)", "secretary power",
            "secretary (energy)", "member (power)", "member secretary"]),
    (75,  ["ceo", "c.e.o", "coo", "c.o.o"]),
    (70,  ["president", "vice president", " vp ", "commissioner"]),
    (65,  ["regional executive director", "red "]),
    (60,  ["head of station", "head of the station", "station head", "station director",
            "plant head", "plant head &", "head o&m", "head of project",
            "hop ", "hop,", "business head", "head - pscc", "head bhira",
            "chief mundra", "chief -", "chief – ", "head of plant"]),
    (55,  ["general manager", "chief general manager", "sr. general manager",
            "group vice president", "group vp", "senior vice president",
            "sr. vice president", "national head"]),
    (50,  [" gm,", " gm(", " gm ", "cgm", "dgm", "agm"]),
    (45,  ["chief engineer", "chief manager", "chief - hydros", "chief superintendent"]),
    (40,  ["additional chief", "addl. c.e", "addl.ce", "addl. chief"]),
    (35,  ["director", "red(", "red-"]),
    (30,  ["superintending engineer", "superintendent of engineer"]),
    (25,  ["se,", "se ", " se(", "c.e ", "ce(", "ce ", " ce,", "c.e(", "c.e.",
            "ee(", "ee,", " ee ", "eee"]),
    (20,  ["executive engineer", "addl. ee", "addl.ee"]),
    (15,  ["sr. manager", "senior manager", "deputy manager", "dy. manager",
            "dy manager", "project manager", "plant manager", "sub divisional"]),
    (10,  ["manager", "head-commercial"]),
]

# Designations that should NEVER be utility head regardless of order
EXCLUDED_DESIGNATIONS = [
    "computer technician", "technician", "junior engineer", "j.e",
    "assistant engineer", "a.e", "team lead", "2nd in-charge", "trainee",
    "site engineer", "resident support engineer", "intern", "operator",
    "officer-", "scientific officer", "station officer", "data entry",
    "programmer", "accounts officer", "jo ", "jo,",
]


def designation_score(desig):
    """Return seniority score for a designation string. 0 means not a utility head."""
    if not desig or not desig.strip():
        return 0
    low = desig.lower().strip()

    for excl in EXCLUDED_DESIGNATIONS:
        if excl in low:
            return 0

    for score, keywords in DESIGNATION_TIERS:
        for kw in keywords:
            if kw in low:
                return score

    return 0


# ── Text helpers ──────────────────────────────────────────────────────────────

def clean(v):
    return re.sub(r"\s+", " ", (v or "").replace("\xa0", " ")).strip()


def para_text(el):
    texts, seen_t = [], set()
    for t in el.iter():
        if t.tag.endswith("}t") and t.text:
            txt = t.text.strip()
            if txt and txt not in seen_t:
                seen_t.add(txt)
                texts.append(txt)
    return clean(" ".join(texts))


def unique_cells(row):
    seen, result = set(), []
    for cell in row.cells:
        if id(cell) not in seen:
            seen.add(id(cell))
            result.append(clean(cell.text))
    return result


def is_junk(text):
    low = text.lower()
    return any(j in low for j in [
        "overcoming barriers", "in union there is strength",
        "working together", "teamwork does not", "we = power",
    ])


def is_address(text):
    low = text.lower()
    return any(s in low for s in [
        "p.o.", "p.o ", "dist", "pin :", "pin:", "stpp", "stps",
        "village", "sector", "complex", "floor", "bhawan",
        "midc", "opp ", "near ", "plot no", "s g highway",
        "road,", "nagar,", "nagar.", "substation compound",
        "shakti bhawan", "vidyut bhavan", "132 kv", "400 kv",
        " 401", " 495", " 486", " 382", " 390", " 380",
        " 452", " 413", " 487", " 492", " 451", " 496", " 482",
        " 403", " 396", " 413", " 411", " 444", " 445",
        "ph no", "fax no", "andheri", "bandra", "ahmedabad",
        "nagpur", "mumbai", "pune", "surat", "vadodara",
        "jabalpur", "raipur", "bhopal", "indore",
    ])


def looks_like_header(values):
    text = " ".join(v.lower() for v in values)
    return "name" in text and any(
        k in text for k in ("designation", "mobile", "telephone", "email", "mail")
    )


def normalize_name(name):
    name = re.sub(r"^[\d\s.–\-]+", "", name)
    return re.sub(r"\s+", " ", name.lower()).strip()


# ── Address extraction ────────────────────────────────────────────────────────

def para_address_for_table(body, tbl_el):
    pos = next((i for i, el in enumerate(body) if el is tbl_el), None)
    if pos is None:
        return "", ""
    org_name, address = "", ""
    for i in range(pos - 1, max(pos - 25, -1), -1):
        el = body[i]
        if el.tag.endswith("}tbl"):
            break
        if not el.tag.endswith("}p"):
            continue
        raw = para_text(el)
        if not raw or len(raw) < 4 or is_junk(raw):
            continue
        if is_address(raw):
            address = address or raw
        else:
            org_name = org_name or raw
        if org_name and address:
            break
    return org_name, address


def table_row_address(table):
    header_index = None
    for idx, row in enumerate(table.rows[:8]):
        if looks_like_header(unique_cells(row)):
            header_index = idx
            break
    if header_index is None:
        return "", ""

    org_name, address = "", ""
    for row in table.rows[:header_index]:
        vals = unique_cells(row)
        non_empty = [v for v in vals if v]
        if not non_empty:
            continue
        unique_vals = list(dict.fromkeys(non_empty))
        combined = " ".join(unique_vals)
        if is_address(combined):
            if not address:
                address = combined
        elif not is_junk(combined) and len(combined) > 4:
            if not org_name:
                org_name = combined
    return org_name, address


# ── Main ──────────────────────────────────────────────────────────────────────

def run():
    doc = Document(DOC_FILE)
    body = list(doc.element.body)
    table_els = [el for el in body if el.tag.endswith("}tbl")]

    with app.app_context():

        # ── 1. Update organization addresses ──────────────────────────────────
        all_orgs = Organization.query.all()
        org_lookup = {
            normalize_name(o.organization_name): o
            for o in all_orgs if o.organization_name
        }

        addr_updated = 0

        for t_idx, tbl_el in enumerate(table_els):
            table = doc.tables[t_idx]
            p_org, p_addr = para_address_for_table(body, tbl_el)
            r_org, r_addr = table_row_address(table)

            raw_org  = p_org  or r_org
            raw_addr = p_addr or r_addr

            if not raw_addr or not raw_org:
                continue

            key = normalize_name(raw_org)
            org_obj = org_lookup.get(key)

            if org_obj is None:
                for stored_key, stored_org in org_lookup.items():
                    if key and (key in stored_key or stored_key in key):
                        org_obj = stored_org
                        break

            if org_obj and not org_obj.address:
                org_obj.address = clean(raw_addr)
                addr_updated += 1

        db.session.commit()
        with_addr = Organization.query.filter(
            Organization.address.isnot(None),
            Organization.address != ""
        ).count()
        print(f"Addresses updated: {addr_updated}  "
              f"| Organizations with address: {with_addr}/{Organization.query.count()}")

        # ── 2. Mark utility heads by designation priority ─────────────────────
        # Reset all flags
        Employee.query.update({"is_utility_head": False})
        db.session.flush()

        # Get all employees grouped by organization
        employees = (
            Employee.query
            .filter(Employee.organization_id.isnot(None))
            .order_by(Employee.organization_id, Employee.id)
            .all()
        )

        # Group by org, score each employee's designation
        from collections import defaultdict
        by_org = defaultdict(list)
        for emp in employees:
            by_org[emp.organization_id].append(emp)

        head_ids = []
        skipped_orgs = 0

        for org_id, org_employees in by_org.items():
            best_emp  = None
            best_score = 0

            for emp in org_employees:
                score = designation_score(emp.designation)
                if score > best_score:
                    best_score = score
                    best_emp   = emp

            if best_emp and best_score > 0:
                head_ids.append(best_emp.id)
            else:
                skipped_orgs += 1

        if head_ids:
            db.session.execute(
                text("UPDATE employees SET is_utility_head = TRUE WHERE id = ANY(:ids)"),
                {"ids": head_ids}
            )
        db.session.commit()

        total_heads = Employee.query.filter_by(is_utility_head=True).count()
        print(f"Utility heads marked: {total_heads}  "
              f"| Orgs with no qualifying head: {skipped_orgs}")

        # Print sample
        heads = (
            Employee.query
            .filter_by(is_utility_head=True)
            .join(Organization)
            .order_by(Employee.id)
            .limit(20)
            .all()
        )
        print("\nSample utility heads (first 20):")
        for e in heads:
            org = e.organization.organization_name if e.organization else "—"
            score = designation_score(e.designation)
            print(f"  [{score:3d}] {e.employee_name:35s} | "
                  f"{(e.designation or ''):30s} | {org[:45]}")


if __name__ == "__main__":
    run()
