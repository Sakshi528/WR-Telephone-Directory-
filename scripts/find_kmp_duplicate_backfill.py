"""Scope-check + proposal for the KMP/non-KMP duplicate-record pattern
found at WRLDC (org 619): a KMP-flagged Employee row missing office_phone/
mobile_phone, with a name-matching NON-KMP row in the SAME organization
that has the missing data. Searched within-org only (not across the whole
table) -- a same-org match is far more defensible than a bare name match
against a common Indian name at an unrelated organization, which would
risk false positives.

Read-only -- writes reports/kmp_duplicate_backfill.csv only.

Matching:
  - exact match: normalized name (letters only, lowercased) identical
  - fuzzy match: difflib ratio >= 0.82 on normalized names, only when no
    exact match exists (same threshold family as
    filter_word_excel_artifacts.py's existing 0.88, slightly relaxed
    since these are shorter person names not org names)

Confidence:
  - HIGH: exact name match, exactly one candidate duplicate, no
    designation conflict
  - NEEDS_REVIEW: fuzzy (non-exact) name match, OR more than one
    candidate duplicate, OR a designation conflict (both sides have a
    non-blank designation that doesn't match even after normalizing case/
    whitespace, and neither is a substring of the other)

Only proposes values for fields that are NULL/blank on the KMP side and
non-blank on the candidate -- never overwrites existing KMP data.

Usage:
  python scripts/find_kmp_duplicate_backfill.py
"""

import csv
import os
import re
import sys
from collections import Counter, defaultdict
from difflib import SequenceMatcher

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from models.employee import Employee
from models.organization import Organization

OUT_CSV = "reports/kmp_duplicate_backfill.csv"
FUZZY_THRESHOLD = 0.82

PHONE_FIELDS = ["office_phone", "mobile_phone"]

# Explicitly excluded per instruction -- confirmed genuine external gap,
# no duplicate exists anywhere for these (same treatment as
# Administrative Heads id 36).
EXCLUDED_IDS = {19210, 19211, 19212, 19216, 19260, 19268, 19276}


def norm_name(name):
    return re.sub(r"[^a-z]", "", (name or "").lower())


def norm_designation(d):
    return re.sub(r"\s+", " ", (d or "").strip().lower())


def designation_conflicts(a, b):
    a, b = norm_designation(a), norm_designation(b)
    if not a or not b:
        return False
    if a == b:
        return False
    return a not in b and b not in a


def main():
    with app.app_context():
        all_emps = Employee.query.all()
        org_name_by_id = {o.id: o.organization_name for o in Organization.query.all()}

    by_org = defaultdict(list)
    for e in all_emps:
        if e.organization_id:
            by_org[e.organization_id].append(e)

    rows = []
    org_row_counts = Counter()
    org_kmp_employee_counts = defaultdict(set)

    for org_id, emps in by_org.items():
        kmp = [e for e in emps if e.is_kmp and e.id not in EXCLUDED_IDS]
        non_kmp = [e for e in emps if not e.is_kmp]
        if not kmp or not non_kmp:
            continue

        non_kmp_by_exact = defaultdict(list)
        for e in non_kmp:
            non_kmp_by_exact[norm_name(e.employee_name)].append(e)

        for k in kmp:
            missing = [f for f in PHONE_FIELDS if not getattr(k, f)]
            if not missing:
                continue

            k_norm = norm_name(k.employee_name)
            exact_candidates = non_kmp_by_exact.get(k_norm, [])

            if exact_candidates:
                candidates, is_exact = exact_candidates, True
            else:
                scored = []
                if len(k_norm) >= 4:
                    for e in non_kmp:
                        ratio = SequenceMatcher(None, k_norm, norm_name(e.employee_name)).ratio()
                        if ratio >= FUZZY_THRESHOLD:
                            scored.append((ratio, e))
                scored.sort(key=lambda x: -x[0])
                candidates = [e for _, e in scored]
                is_exact = False

            if not candidates:
                continue  # no duplicate found in this org -- out of scope, not proposed

            best = candidates[0]
            conflicting = any(designation_conflicts(k.designation, c.designation) for c in candidates)
            multiple = len(candidates) > 1

            confidence = "HIGH" if (is_exact and not multiple and not conflicting) else "NEEDS_REVIEW"

            note_parts = []
            if not is_exact:
                note_parts.append(f"fuzzy name match ({SequenceMatcher(None, k_norm, norm_name(best.employee_name)).ratio():.2f})")
            if multiple:
                note_parts.append(f"{len(candidates)} candidate duplicates found")
            if conflicting:
                note_parts.append(f"designation conflict: KMP={k.designation!r} vs duplicate={best.designation!r}")

            for field in missing:
                val = getattr(best, field)
                if not val:
                    continue
                rows.append({
                    "kmp_employee_id": k.id, "name": k.employee_name,
                    "org": org_name_by_id.get(org_id, str(org_id)),
                    "field": field, "current_value": "", "proposed_value": val,
                    "duplicate_employee_id": best.id, "confidence": confidence,
                    "notes": "; ".join(note_parts) if note_parts else "",
                })
                org_row_counts[org_name_by_id.get(org_id, str(org_id))] += 1
                org_kmp_employee_counts[org_name_by_id.get(org_id, str(org_id))].add(k.id)

    fieldnames = ["kmp_employee_id", "name", "org", "field", "current_value",
                  "proposed_value", "duplicate_employee_id", "confidence", "notes"]
    os.makedirs("reports", exist_ok=True)
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)

    print("=" * 70)
    print(f"  KMP/non-KMP duplicate backfill scope -- {len(org_row_counts)} organizations affected")
    print("=" * 70)
    for org, count in sorted(org_row_counts.items(), key=lambda x: -x[1]):
        print(f"    {org:<55} {count:>4} field-rows, {len(org_kmp_employee_counts[org]):>3} employees")
    print("-" * 70)
    print(f"  Total rows: {len(rows)} -> {OUT_CSV}")
    conf_counts = Counter(r["confidence"] for r in rows)
    print(f"  Confidence: {dict(conf_counts)}")
    print("=" * 70)


if __name__ == "__main__":
    main()
