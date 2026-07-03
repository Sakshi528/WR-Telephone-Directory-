"""
fix_data_quality.py
============================================================
1. Fix organization names that were incorrectly extracted
   (address stored as org name, spacing errors, Unicode garbage).
2. Re-run utility head marking after org name corrections.
============================================================
Run: python fix_data_quality.py
"""

import re
from app import app
from models import db
from models.organization import Organization
from models.employee import Employee
from sqlalchemy import text, or_


# ── Org name corrections ─────────────────────────────────────────────────────
# Format: (org_id, correct_name, address_if_known)
ORG_FIXES = [
    # Address was extracted as org name
    (81,  "INOX Green Energy Services Limited",
          "Dayapar (Kutch), Gujarat - 370625"),
    (87,  "Continuum Power Trading Private Limited",
          "C Wing, 402/404, Delphi Building, Orchard Avenue, Powai, Mumbai - 400076"),
    (167, "KSK Mahanadi Power Company Limited (JSW Energy)",
          "/82/A/431/A, Road No 22, Jubilee Hills, Hyderabad - 500033"),
    # Spacing errors
    (59,  "Daman and Diu", ""),
    (63,  "NTPC Kawas", ""),
    # Unicode en-dash (–) and garbled chars (�) -> plain hyphen / space
    # These will be handled by the bulk_clean pass below
]

# ── Bulk Unicode cleanup patterns ────────────────────────────────────────────
# Replace en-dash (–), em-dash (—), and replacement char (?) in org names
UNICODE_REPLACEMENTS = [
    ("–", "-"),   # en-dash
    ("—", "-"),   # em-dash
    ("�", ""),    # replacement char
    ("  ", " "),       # double spaces
]


def clean_org_name(name):
    if not name:
        return name
    for old, new in UNICODE_REPLACEMENTS:
        name = name.replace(old, new)
    return name.strip()


def run():
    with app.app_context():

        # ── 1. Targeted org name corrections ─────────────────────────────────
        fixed = 0
        for org_id, correct_name, address in ORG_FIXES:
            org = db.session.get(Organization, org_id)
            if org:
                old = org.organization_name
                org.organization_name = correct_name
                if address and not org.address:
                    org.address = address
                elif address:
                    org.address = address
                print(f"  Fixed org {org_id}: [{old[:50]}] -> [{correct_name}]")
                fixed += 1

        # ── 2. Bulk Unicode cleanup ───────────────────────────────────────────
        bulk_fixed = 0
        for org in Organization.query.all():
            cleaned = clean_org_name(org.organization_name)
            if cleaned != org.organization_name:
                print(f"  Unicode fix {org.id}: [{org.organization_name[:60]}] -> [{cleaned[:60]}]")
                org.organization_name = cleaned
                bulk_fixed += 1

        db.session.commit()
        print(f"\nTargeted fixes: {fixed}  |  Unicode cleanup: {bulk_fixed}")

        # ── 3. Re-run utility head marking after renames ─────────────────────
        # Designation scoring (copy from update_addresses_and_heads.py logic)
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
                   "dy manager", "project manager", "plant manager"]),
            (10,  ["manager", "head-commercial"]),
        ]
        EXCLUDED = [
            "computer technician", "technician", "junior engineer", "j.e",
            "assistant engineer", "a.e", "team lead", "2nd in-charge", "trainee",
            "site engineer", "resident support engineer", "intern", "operator",
            "scientific officer", "data entry", "programmer", "accounts officer",
        ]

        def score(desig):
            if not desig or not desig.strip():
                return 0
            low = desig.lower().strip()
            for ex in EXCLUDED:
                if ex in low:
                    return 0
            for s, kws in DESIGNATION_TIERS:
                for kw in kws:
                    if kw in low:
                        return s
            return 0

        Employee.query.update({"is_utility_head": False})
        db.session.flush()

        employees = (
            Employee.query
            .filter(Employee.organization_id.isnot(None))
            .filter(Employee.employee_name.isnot(None))
            .filter(Employee.employee_name != "")
            .order_by(Employee.organization_id, Employee.id)
            .all()
        )

        from collections import defaultdict
        by_org = defaultdict(list)
        for emp in employees:
            by_org[emp.organization_id].append(emp)

        head_ids = []
        for org_id, org_emps in by_org.items():
            best, best_score = None, 0
            for emp in org_emps:
                s = score(emp.designation)
                if s > best_score:
                    best_score, best = s, emp
            if best and best_score > 0:
                head_ids.append(best.id)

        if head_ids:
            db.session.execute(
                text("UPDATE employees SET is_utility_head = TRUE WHERE id = ANY(:ids)"),
                {"ids": head_ids}
            )
        db.session.commit()

        total_heads = Employee.query.filter_by(is_utility_head=True).count()
        print(f"Utility heads re-marked: {total_heads}")

        # ── 4. Fix employees with empty names ────────────────────────────────
        blank = Employee.query.filter(
            (Employee.employee_name == None) | (Employee.employee_name == "")
        ).all()
        print(f"\nEmployees with blank names: {len(blank)}")
        for e in blank[:10]:
            print(f"  ID:{e.id} org_id:{e.organization_id} desig:[{e.designation}]")

        # Remove employees with no name and no designation (pure junk rows)
        removed = 0
        for e in blank:
            if not e.designation or not e.designation.strip():
                db.session.delete(e)
                removed += 1
        db.session.commit()
        print(f"Removed {removed} nameless/designationless employee rows")


if __name__ == "__main__":
    run()
