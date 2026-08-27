"""Pre-fix-script deliverable, requested as a follow-up to
full_data_quality_audit.py before any fix script is written. Read-only
throughout -- writes proposal CSVs and a console summary only, changes
nothing in the database.

1. Splits address_in_name into two confidence tiers (reuses the exact same
   signal detector as full_data_quality_audit.py, so the split is derived
   from the real per-row signals, not a re-parse of the notes text):
     - high_confidence: PIN-code pattern, long parenthetical, 2+ commas, or
       a "Backup"/"State Load Despatch Centre" phrase hit
     - low_confidence: ONLY a bare city/state-name hit with nothing else
       (e.g. "Gujarat State Electricity Corporation Limited" -- likely a
       real company name, not address contamination)
   Any row carrying the (rare, currently unused) "unusually long name"
   signal alongside nothing else is conservatively treated as high, since
   it isn't the bare-city-hint case the low tier is defined as.

2. numbered_prefix: NOT a fix proposal -- prints a confirmation of whether
   any export/grouping/display logic in routes/user_routes.py (or the PDF/
   Excel generators, or the admin listing) sorts by or otherwise depends on
   the leading number in DirectoryNumber.name.

3. duplicated_phone fix proposals, split by the Phase 2 origin bucket for
   each row's station (from reports/word_vs_app_cleanliness.csv):
     - import_bugs: station's likely_origin == "introduced_during_import"
     - word_divergence: likely_origin in ("originated_in_word", "mixed")
     - no_word_match: station has duplicated_phone in the app but never
       matched a Word segment at all, so origin can't be determined either
       way -- not requested as a named bucket, but silently dropping these
       would hide them, so they get their own small file instead.
   Cleaned value: split on the same [/;,] separators as the detector, drop
   exact repeat segments while preserving first-seen order, rejoin with
   " / ". Only touches exact-duplicate segments -- doesn't invent, correct,
   or drop any other segment (e.g. a bare "NA" placeholder segment is left
   alone; that's a different issue, not this one's job).

Usage:
  python scripts/build_fix_proposals.py
"""

import csv
import os
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from models.directory_number import DirectoryNumber

from scripts.full_data_quality_audit import (
    address_in_name_signals,
    classify_category,
    SEGMENT_SPLIT_RE,
)

AUDIT_CSV = "reports/full_data_quality_audit.csv"
CLEANLINESS_CSV = "reports/word_vs_app_cleanliness.csv"

ADDR_HIGH_CSV = "reports/address_in_name_high_confidence.csv"
ADDR_LOW_CSV = "reports/address_in_name_low_confidence.csv"

DUP_IMPORT_CSV = "reports/fix_proposal_duplicated_phone_import_bugs.csv"
DUP_WORD_DIVERGENCE_CSV = "reports/fix_proposal_duplicated_phone_word_divergence.csv"
DUP_NO_WORD_MATCH_CSV = "reports/fix_proposal_duplicated_phone_no_word_match.csv"


# ─── 1. address_in_name confidence split ────────────────────────────────────

def is_bare_city_hit(signals):
    return len(signals) == 1 and signals[0].startswith("city/state name")


def split_address_in_name(row_data):
    high, low = [], []
    for r in row_data:
        signals = address_in_name_signals(r["name"])
        if not signals:
            continue
        tier = "low_confidence" if is_bare_city_hit(signals) else "high_confidence"
        out = {
            "id": r["id"], "station": r["station"], "category": r["category"],
            "raw_value": r["name"], "signals": "; ".join(signals),
        }
        (low if tier == "low_confidence" else high).append(out)

    fieldnames = ["id", "station", "category", "raw_value", "signals"]
    for path, rows in ((ADDR_HIGH_CSV, high), (ADDR_LOW_CSV, low)):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader()
            w.writerows(rows)
    return high, low


# ─── 2. numbered_prefix sort-order confirmation (no fix, no CSV) ──────────

NUMBERED_PREFIX_FINDING = """
numbered_prefix -- sort-order confirmation (no fix proposed):
  Checked routes/user_routes.py (telephone_directory page, PDF/Excel export
  snapshot builder, docx per-org export, search/autosuggest) and
  services/pdf_generator.py / services/excel_generator.py. Found exactly one
  place that orders by DirectoryNumber.name: the admin listing/search query
  in routes/admin_routes.py (order_by(DirectoryNumber.organization,
  DirectoryNumber.name)) -- a plain alphabetical secondary sort, not a
  numbering scheme anyone designed around.
  Every other grouping/display path (telephone_directory page, PDF export,
  Excel export, docx org export) orders by Organization, not by name, and
  within an organization simply iterates DB query order -- no ORDER BY on
  name at all.
  Verified all 13 flagged rows are the ONLY DirectoryNumber row for their
  (organization_id, category) -- there is no sibling row anywhere for the
  leading number to be ordering against. Combined with the missing-sort-
  dependency finding above, the leading number ("2. Gandhi Sagar", "7.
  Tons", ...) does not appear to be functionally load-bearing anywhere in
  the app; it matches the same Word-section-heading-leaked-into-the-name-
  field pattern already documented in reports/naming_artifacts.csv for
  other stations. Still holding off on a fix per your instruction -- this
  is a confirmation, not a proposal.
"""


# ─── 3. duplicated_phone fix proposals, split by origin bucket ────────────

def dedupe_preserve_order(segments):
    seen = set()
    out = []
    for s in segments:
        if s not in seen:
            seen.add(s)
            out.append(s)
    return out


def propose_phone_value(raw_value):
    segments = [s.strip() for s in SEGMENT_SPLIT_RE.split(raw_value) if s.strip()]
    return " / ".join(dedupe_preserve_order(segments))


def load_station_origin(field="duplicated_phone"):
    with open(CLEANLINESS_CSV, encoding="utf-8") as f:
        return {r["station"]: r["likely_origin"] for r in csv.DictReader(f)}


def propose_duplicated_phone_fixes(row_data):
    origin_by_station = load_station_origin()

    dup_rows = [r for r in row_data if "duplicated_phone" in r["issue_types"]]

    import_bugs, word_divergence, no_word_match = [], [], []
    for r in dup_rows:
        origin = origin_by_station.get(r["station"])
        proposed = propose_phone_value(r["phone_number"])
        out = {
            "id": r["id"], "station": r["station"], "category": r["category"],
            "raw_value": r["phone_number"], "proposed_value": proposed,
            "likely_origin": origin or "no_word_match",
        }
        if origin == "introduced_during_import":
            import_bugs.append(out)
        elif origin in ("originated_in_word", "mixed"):
            word_divergence.append(out)
        else:
            no_word_match.append(out)

    fieldnames = ["id", "station", "category", "raw_value", "proposed_value", "likely_origin"]
    for path, rows in (
        (DUP_IMPORT_CSV, import_bugs),
        (DUP_WORD_DIVERGENCE_CSV, word_divergence),
        (DUP_NO_WORD_MATCH_CSV, no_word_match),
    ):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader()
            w.writerows(rows)

    return import_bugs, word_divergence, no_word_match


# ─── main ────────────────────────────────────────────────────────────────

def load_row_data():
    with app.app_context():
        rows = DirectoryNumber.query.all()
        return [
            {
                "id": dn.id,
                "station": dn.organization_obj.organization_name if dn.organization_obj else (dn.organization or ""),
                "category": classify_category(dn.category),
                "name": dn.name or "",
                "phone_number": dn.phone_number or "",
            }
            for dn in rows
        ]


def attach_issue_types(row_data):
    """Reads full_data_quality_audit.csv (already-computed Phase 1 output)
    to know which issue types apply to each id, instead of recomputing
    duplicated_phone detection separately -- single source of truth."""
    issues_by_id = defaultdict(set)
    with open(AUDIT_CSV, encoding="utf-8") as f:
        for r in csv.DictReader(f):
            for issue in r["issue_type(s)"].split("; "):
                issues_by_id[int(r["id"])].add(issue)
    for r in row_data:
        r["issue_types"] = issues_by_id.get(r["id"], set())
    return row_data


def main():
    row_data = attach_issue_types(load_row_data())

    high, low = split_address_in_name(row_data)

    print(NUMBERED_PREFIX_FINDING)

    import_bugs, word_divergence, no_word_match = propose_duplicated_phone_fixes(row_data)

    print("=" * 70)
    print("  Fix-Proposal Prep -- Updated Counts (nothing written to the DB)")
    print("=" * 70)
    print("  address_in_name split:")
    print(f"    - high_confidence : {len(high)}  -> {ADDR_HIGH_CSV}")
    print(f"    - low_confidence  : {len(low)}  -> {ADDR_LOW_CSV}  (held for manual review, excluded from fix-proposal scope)")
    print()
    print("  duplicated_phone split by Word<->app origin (station-level, from word_vs_app_cleanliness.csv):")
    print(f"    - import_bugs (introduced_during_import) : {len(import_bugs)} rows / "
          f"{len({r['station'] for r in import_bugs})} stations -> {DUP_IMPORT_CSV}")
    print(f"    - word_divergence (originated_in_word + mixed) : {len(word_divergence)} rows / "
          f"{len({r['station'] for r in word_divergence})} stations -> {DUP_WORD_DIVERGENCE_CSV}")
    print(f"    - no_word_match (station never matched Word at all -- origin unknown) : "
          f"{len(no_word_match)} rows / {len({r['station'] for r in no_word_match})} stations -> {DUP_NO_WORD_MATCH_CSV}")
    print("=" * 70)


if __name__ == "__main__":
    main()
