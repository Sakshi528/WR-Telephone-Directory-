"""Full pass: for every Organization, compares its current Category +
Subcategory (Organization.region) against what the Word directory's
top-level (Heading 1) section implies.

Read-only -- writes reports/full_category_subcategory_bifurcation.csv
only, no DB writes. Reuses the same section-parsing/org-matching
machinery as scripts/word_section_vs_app_category.py (see that file's
docstring for the full lineage), but reports on ALL 220 organizations,
not just the 102 with a clean single-section match -- including orgs
under an ambiguous/catch-all section and orgs the Word doc never
mentions in its body tables at all.

Usage:
  python scripts/full_category_subcategory_bifurcation.py
"""

import csv
import os
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from models.organization import Organization
from models.organization_category import OrganizationCategory

from scripts.word_section_vs_app_category import (
    WORD_PATH, NON_ORG_SECTIONS, SECTION_TO_CATEGORY, AMBIGUOUS_SECTIONS,
    parse_sections,
)

OUT_CSV = "reports/full_category_subcategory_bifurcation.csv"


def main():
    with app.app_context():
        orgs = Organization.query.order_by(Organization.organization_name).all()
        org_id_by_name_lower = {o.organization_name.strip().lower(): o.id for o in orgs}
        cat_name_by_id = {c.id: c.category_name for c in OrganizationCategory.query.all()}

    segments = parse_sections(WORD_PATH, org_id_by_name_lower)

    with app.app_context():
        sections_by_org = defaultdict(set)
        for h1, station_name, employees, cr, sw in segments:
            if not (employees or cr or sw):
                continue
            if h1 is None or h1 in NON_ORG_SECTIONS:
                continue
            from scripts.full_word_vs_app_field_diff import match_to_org
            org_id = match_to_org(station_name, org_id_by_name_lower)
            if org_id is None:
                continue
            sections_by_org[org_id].add(h1)

        rows = []
        for org in orgs:
            h1_set = sections_by_org.get(org.id, set())
            current_cat = cat_name_by_id.get(org.category_id, "") if org.category_id else "(none)"

            if len(h1_set) == 0:
                word_section = "(not found in Word body tables)"
                implied_cat = ""
                note = "no Word section match"
            elif len(h1_set) > 1:
                word_section = " | ".join(sorted(h1_set))
                implied_cat = ""
                note = "spans multiple Word sections"
            else:
                (only_h1,) = h1_set
                word_section = only_h1
                if only_h1 in SECTION_TO_CATEGORY:
                    implied_cat = SECTION_TO_CATEGORY[only_h1]
                    note = "match" if implied_cat == current_cat else "MISMATCH"
                elif only_h1 in AMBIGUOUS_SECTIONS:
                    implied_cat = ""
                    note = "ambiguous section: " + AMBIGUOUS_SECTIONS[only_h1]
                else:
                    implied_cat = ""
                    note = "unmapped section"

            rows.append({
                "org_id": org.id,
                "organization_name": org.organization_name,
                "current_category": current_cat,
                "current_subcategory_region": org.region or "",
                "word_h1_section": word_section,
                "implied_category": implied_cat,
                "note": note,
            })

    os.makedirs("reports", exist_ok=True)
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)

    mismatches = [r for r in rows if r["note"] == "MISMATCH"]
    no_match = [r for r in rows if r["note"] == "no Word section match"]
    ambiguous = [r for r in rows if r["note"].startswith("ambiguous")]
    unmapped = [r for r in rows if r["note"] == "unmapped section"]
    multi = [r for r in rows if r["note"] == "spans multiple Word sections"]
    clean_match = [r for r in rows if r["note"] == "match"]

    print("=" * 70)
    print(f"  Total organizations                         : {len(rows)}  -> {OUT_CSV}")
    print(f"    - clean match                              : {len(clean_match)}")
    print(f"    - MISMATCH (category should change)        : {len(mismatches)}")
    print(f"    - no Word section match at all              : {len(no_match)}")
    print(f"    - ambiguous/catch-all section               : {len(ambiguous)}")
    print(f"    - unmapped section (no category rule yet)   : {len(unmapped)}")
    print(f"    - spans multiple Word sections               : {len(multi)}")
    print("=" * 70)
    if mismatches:
        print("MISMATCHES:")
        for r in mismatches:
            print(f"  {r['organization_name']}: {r['current_category']} -> {r['implied_category']}  (section: {r['word_h1_section']})")


if __name__ == "__main__":
    main()
