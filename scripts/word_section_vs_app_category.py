"""Cross-checks each Organization's current app category against which
top-level (Heading 1) numbered section of the Word directory it falls
under -- e.g. orgs under "Nuclear Power Stations in Western Region" should
be category=Nuclear, orgs under "Other Transmission Licensees in Western
Region" should be category=Transmission Utility.

Read-only -- writes reports/word_section_vs_app_category.csv only, no DB
writes.

Reuses the existing station-to-org matching machinery unchanged:
  - parse_table_records / iter_block_items / clean / strip_leading_number
    (compare_word_excel.py, full_word_vs_app_field_diff.py) for turning
    Word tables into (station_name, employees, cr, sw) segments
  - match_to_org (full_word_vs_app_field_diff.py) for resolving a station
    name to an Organization id, including its substring-match fallback

The only new logic here is tracking which Heading 1 paragraph most
recently preceded each segment -- parse_word_with_address in
full_word_vs_app_field_diff.py deliberately treats every heading level the
same way (it needs the immediate/nearest heading as the station name), so
it doesn't expose the enclosing top-level section. This script does a
parallel single walk of the same document, calling the same
parse_table_records for each table, but tracking current_h1 (last Heading
1 text) separately from current_heading (nearest heading of any level,
same as before) -- so it augments, not replaces, the existing per-station
matching.

Usage:
  python scripts/word_section_vs_app_category.py
"""

import csv
import os
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import docx
from docx.text.paragraph import Paragraph

from app import app
from models.organization import Organization
from models.organization_category import OrganizationCategory

from scripts.compare_word_excel import clean, strip_leading_number, iter_block_items
from scripts.full_word_vs_app_field_diff import parse_table_records, match_to_org

WORD_PATH = "uploads/Western Region Phone Directory 2026 Main_TD.docx"
OUT_CSV = "reports/word_section_vs_app_category.csv"

# Non-org front/back matter Heading 1 sections -- not organizational
# groupings at all, so out of scope for a category cross-check.
NON_ORG_SECTIONS = {
    "Synopsis of Important Telephone Numbers",
    "Disaster Management Contact Details",
    "24. Hotline Nos- Orange Directory",
    "25. Important Vendors in WRLDC",
    "26. Emergency Contact Numbers",
}

# Word Heading-1 section -> implied app category. Sections not listed here
# are intentionally left unmapped (ambiguous / catch-all / no single
# category fits) and reported separately, never guessed.
SECTION_TO_CATEGORY = {
    "Western Region Load Despatch Centre": "RLDC",
    "NLDC and RLDCs": "RLDC",
    "Western Regional Power Committee": "Regional Power Committee",
    "Maharashtra State Load Despatch Centre": "SLDC",
    "Gujarat State Load Despatch Centre": "SLDC",
    "Madhya Pradesh State Load Despatch Centre": "SLDC",
    "Chhattisgarh State Load Despatch Centre": "SLDC",
    "Goa Electricity Department": "SLDC",
    "Daman and Diu": "SLDC",
    "Dadra and Nagar Haveli": "SLDC",
    "NTPC Mumbai Head Quarter (WR-1)": "Thermal",
    "NTPC Raipur Head Quarter (WR-2)": "Thermal",
    "Nuclear Power Stations in Western Region": "Nuclear",
    # RE Generators category exists but was deliberately merged into
    # Generation Company earlier this session (74 orgs, 2026-08-27) --
    # mapping here follows that decision rather than the now-empty
    # RE Generators category.
    "ISTS Connected RE Generators in Western Region": "Generation Company",
    "Other Transmission Licensees in Western Region": "Transmission Utility",
}

# Sections deliberately left unmapped, with the reason -- shown to the
# user, never silently guessed.
AMBIGUOUS_SECTIONS = {
    "Black Start Facilitated Stations": "cross-cutting flag, not a category -- these are Thermal/Hydel/Nuclear stations already listed under their real section elsewhere in the document",
    "Qualified Coordinating Agency": "a functional role (QCA), not one of the existing app categories",
    "Other Generating Stations in Western Region": "catch-all 'Other' bucket -- likely a mix of Thermal/Hydel/Generation Company, can't assign a single category without inspecting each org",
    "Other Buyers in Western Region": "catch-all 'Other' bucket, no clear single-category match",
    "Other Users in Western Region": "catch-all 'Other' bucket, no clear single-category match",
    "RRAS Provider in Western Region": "functional role (ancillary services provider), not one of the existing app categories",
}


def parse_sections(path, org_id_by_name_lower):
    doc = docx.Document(path)
    current_h1 = None
    current_heading = None
    all_segments = []  # (h1, station_name, employees, cr, sw)

    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            text = clean(item.text)
            if item.style.name == "Heading 1" and text:
                current_h1 = text
                current_heading = strip_leading_number(text)
            elif item.style.name.startswith("Heading") and text:
                current_heading = strip_leading_number(text)
            continue

        segments = parse_table_records(item, current_heading, org_id_by_name_lower)
        for name, _name_raw, employees, cr, sw in segments:
            all_segments.append((current_h1, name, employees, cr, sw))

    return all_segments


def main():
    with app.app_context():
        orgs = Organization.query.all()
        org_id_by_name_lower = {o.organization_name.strip().lower(): o.id for o in orgs}
        org_by_id = {o.id: o for o in orgs}
        cat_name_by_id = {c.id: c.category_name for c in OrganizationCategory.query.all()}

    segments = parse_sections(WORD_PATH, org_id_by_name_lower)

    with app.app_context():
        sections_by_org = defaultdict(set)
        unmatched_sections = defaultdict(int)

        for h1, station_name, employees, cr, sw in segments:
            if not (employees or cr or sw):
                continue
            if h1 is None or h1 in NON_ORG_SECTIONS:
                continue
            org_id = match_to_org(station_name, org_id_by_name_lower)
            if org_id is None:
                unmatched_sections[(h1, station_name)] += 1
                continue
            sections_by_org[org_id].add(h1)

        rows = []
        ambiguous_multi_section = []
        ambiguous_no_category_map = []
        matches = 0
        mismatches = 0
        skipped_no_word_section = 0

        for org in orgs:
            h1_set = sections_by_org.get(org.id, set())
            app_cat = cat_name_by_id.get(org.category_id, "(none)")

            if not h1_set:
                skipped_no_word_section += 1
                continue

            if len(h1_set) > 1:
                ambiguous_multi_section.append({
                    "organization": org.organization_name,
                    "word_sections": "; ".join(sorted(h1_set)),
                    "app_current_category": app_cat,
                })
                continue

            word_section = next(iter(h1_set))

            if word_section in AMBIGUOUS_SECTIONS:
                ambiguous_no_category_map.append({
                    "organization": org.organization_name,
                    "word_section": word_section,
                    "app_current_category": app_cat,
                    "reason": AMBIGUOUS_SECTIONS[word_section],
                })
                continue

            implied_cat = SECTION_TO_CATEGORY.get(word_section)
            if implied_cat is None:
                ambiguous_no_category_map.append({
                    "organization": org.organization_name,
                    "word_section": word_section,
                    "app_current_category": app_cat,
                    "reason": "Word section not in the confirmed heading-to-category mapping",
                })
                continue

            match = "yes" if implied_cat == app_cat else "no"
            if match == "yes":
                matches += 1
            else:
                mismatches += 1
            rows.append({
                "organization": org.organization_name,
                "word_section": word_section,
                "word_implied_category": implied_cat,
                "app_current_category": app_cat,
                "match": match,
            })

    fieldnames = ["organization", "word_section", "word_implied_category", "app_current_category", "match"]
    os.makedirs("reports", exist_ok=True)
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)

    print("=" * 70)
    print("  Word Section vs App Category -- audit only, nothing changed")
    print("=" * 70)
    print(f"  Organizations compared (clean 1:1 section match) : {len(rows)}  -> {OUT_CSV}")
    print(f"    - match=yes  : {matches}")
    print(f"    - match=no   : {mismatches}")
    print(f"  Organizations under an ambiguous/unmapped section : {len(ambiguous_no_category_map)}")
    print(f"  Organizations under multiple Word H1 sections     : {len(ambiguous_multi_section)}")
    print(f"  Organizations with no matched Word section at all : {skipped_no_word_section}")
    print(f"  Word station segments with no app org match       : {len(unmatched_sections)}")
    print("=" * 70)

    if mismatches:
        from collections import Counter
        by_cat_pair = Counter((r["word_implied_category"], r["app_current_category"]) for r in rows if r["match"] == "no")
        print("  Mismatch breakdown (word_implied -> app_current):")
        for (wc, ac), n in sorted(by_cat_pair.items(), key=lambda x: -x[1]):
            print(f"    {wc:<20} -> {ac:<25} {n}")
        print("=" * 70)

    with open("reports/word_section_vs_app_category_ambiguous.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["organization", "word_section", "app_current_category", "reason"])
        w.writeheader()
        w.writerows(ambiguous_no_category_map)

    with open("reports/word_section_vs_app_category_multi_section.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["organization", "word_sections", "app_current_category"])
        w.writeheader()
        w.writerows(ambiguous_multi_section)

    print(f"  Ambiguous/unmapped -> reports/word_section_vs_app_category_ambiguous.csv")
    print(f"  Multi-section       -> reports/word_section_vs_app_category_multi_section.csv")
    print("=" * 70)


if __name__ == "__main__":
    main()
