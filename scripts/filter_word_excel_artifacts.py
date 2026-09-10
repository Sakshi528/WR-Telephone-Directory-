"""Splits reports/word_vs_excel_2026.csv into three buckets, reusing
scripts/filter_reconciliation_artifacts.py's per-(Station, Category) pairing
logic unchanged (same three naming-artifact patterns, same CSV schema) --
see that module for the a/b/c pattern definitions.

This adds a fourth bucket on top: DATA QUALITY ISSUES. After the exact
naming-artifact pass, some Missing/Extra rows within a group are near-but-
not-exact matches of each other -- e.g. "Avikaran Solar India Private
Limited" vs "Avikiran Solar India Private Limited", a one-letter spelling
difference, not a punctuation-only one, so Pattern C's exact-after-
normalize check correctly leaves it unpaired. These are the same real-world
entity with a typo/formatting slip in one source, not a genuinely missing
or extra record -- remediation is "fix the record," not "add the missing
entry" or "note as a real gap," so they get their own bucket rather than
being lumped into either of the other two.

Fuzzy matching is difflib.SequenceMatcher (stdlib, no new dependency) on
normalized names, ratio >= FUZZY_THRESHOLD, applied only within the same
(Station, Category) group and only to rows the exact pass left unmatched --
greedy best-match pairing, one match per row, so it can't chain unrelated
names together through a low-confidence bridge match.

Usage:
  python scripts/filter_word_excel_artifacts.py
"""

import csv
import os
import sys
from collections import defaultdict
from difflib import SequenceMatcher

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from scripts.filter_reconciliation_artifacts import classify_group, normalize

INPUT_CSV = "reports/word_vs_excel_2026.csv"
REAL_CSV = "reports/word_excel_real_discrepancies.csv"
ARTIFACTS_CSV = "reports/word_excel_naming_artifacts.csv"
DATA_QUALITY_CSV = "reports/word_excel_data_quality_issues.csv"

FUZZY_THRESHOLD = 0.88
FUZZY_MIN_LENGTH = 4  # skip very short names -- coincidental high ratios


def find_data_quality_pairs(rows):
    """rows: the leftover (non-artifact) rows for ONE (Station, Category)
    group. Returns (data_quality_rows, remaining_real_rows)."""
    extras = [r for r in rows if r["Issue"] == "Extra"]
    missing = [r for r in rows if r["Issue"] == "Missing"]
    others = [r for r in rows if r["Issue"] not in ("Extra", "Missing")]

    used_extra_idx = set()
    used_missing_idx = set()
    dq_pairs = []

    for mi, m in enumerate(missing):
        m_norm = normalize(m["Name"])
        if len(m_norm) < FUZZY_MIN_LENGTH:
            continue
        best_ei, best_ratio = None, 0.0
        for ei, e in enumerate(extras):
            if ei in used_extra_idx:
                continue
            e_norm = normalize(e["Name"])
            if len(e_norm) < FUZZY_MIN_LENGTH:
                continue
            ratio = SequenceMatcher(None, m_norm, e_norm).ratio()
            if ratio > best_ratio:
                best_ei, best_ratio = ei, ratio
        if best_ei is not None and best_ratio >= FUZZY_THRESHOLD:
            used_extra_idx.add(best_ei)
            used_missing_idx.add(mi)
            dq_pairs.append({**m, "Details": m["Details"] + f" [near-duplicate, similarity {best_ratio:.2f}]"})
            dq_pairs.append({**extras[best_ei], "Details": extras[best_ei]["Details"] + f" [near-duplicate, similarity {best_ratio:.2f}]"})

    remaining = (
        [e for i, e in enumerate(extras) if i not in used_extra_idx]
        + [m for i, m in enumerate(missing) if i not in used_missing_idx]
        + others
    )
    return dq_pairs, remaining


def main():
    with open(INPUT_CSV, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    groups = defaultdict(list)
    for r in rows:
        groups[(r["Station"], r["Category"])].append(r)

    artifacts, data_quality, real = [], [], []
    for (station, category), group_rows in groups.items():
        a, leftover = classify_group(station, category, group_rows)
        artifacts.extend(a)
        dq, remaining_real = find_data_quality_pairs(leftover)
        data_quality.extend(dq)
        real.extend(remaining_real)

    fieldnames = ["Station", "Category", "Issue", "Name", "Details"]

    for path, out_rows in ((REAL_CSV, real), (ARTIFACTS_CSV, artifacts), (DATA_QUALITY_CSV, data_quality)):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader()
            w.writerows(out_rows)

    total_out = len(real) + len(artifacts) + len(data_quality)
    print("=" * 70)
    print("  Word <-> Excel Naming-Artifact / Data-Quality Filter Summary")
    print("=" * 70)
    print(f"  Input rows              : {len(rows)}")
    print(f"  -> Real discrepancies   : {len(real)}")
    print(f"  -> Naming artifacts     : {len(artifacts)}")
    print(f"  -> Data quality issues  : {len(data_quality)}")
    print(f"  Check: {len(real)} + {len(artifacts)} + {len(data_quality)} = {total_out} "
          f"({'matches input' if total_out == len(rows) else 'MISMATCH vs input ' + str(len(rows))})")
    print("=" * 70)

    by_cat = defaultdict(lambda: [0, 0, 0])
    for r in real:
        by_cat[r["Category"]][0] += 1
    for r in artifacts:
        by_cat[r["Category"]][1] += 1
    for r in data_quality:
        by_cat[r["Category"]][2] += 1
    print(f"  {'Category':<15} {'Real':>6} {'Artifact':>9} {'DataQual':>9}")
    for cat, (real_n, art_n, dq_n) in sorted(by_cat.items()):
        print(f"  {cat:<15} {real_n:>6} {art_n:>9} {dq_n:>9}")
    print("=" * 70)
    print(f"  Written: {REAL_CSV}")
    print(f"  Written: {ARTIFACTS_CSV}")
    print(f"  Written: {DATA_QUALITY_CSV}")


if __name__ == "__main__":
    main()
