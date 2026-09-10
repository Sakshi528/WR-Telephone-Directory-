"""Splits reports/directory_reconciliation_report.csv into genuine gaps and
naming artifacts. Read-only against the CSV; writes two new CSVs.

Three artifact patterns, all scoped to Control Room / Switchyard rows only
(Employee rows are never filtered -- there is no "generic label" concept for
a person's name):

  A/B) An Extra row's name, once a leading section number is stripped
       ("18. Bhira" -> "Bhira"), equals the station's own name, paired with
       a Missing row whose name is the generic literal "Control Room" --
       same record, database stores the org name, Word just says the label.
  C)   An Extra/Missing pair whose names are identical once case,
       punctuation, and whitespace are normalized away (e.g. "PH-3, 4" vs
       "PH-3 , 4", or "Control Room 1" vs "Control Room-1").

Pairing is per (Station, Category) group, matched greedily by name so a
group with N extras and N missing rows (e.g. Dhuvaran's 3 control rooms) is
only fully classified as an artifact if every row finds a partner --
leftover, unmatched rows in a partially-matching group stay as real
discrepancies rather than being swept in with the matched ones.

Usage:
  python scripts/filter_reconciliation_artifacts.py
"""

import csv
import re
from collections import defaultdict

INPUT_CSV = "reports/directory_reconciliation_report.csv"
REAL_CSV = "reports/real_discrepancies.csv"
ARTIFACTS_CSV = "reports/naming_artifacts.csv"

LEADING_NUMBER_RE = re.compile(r"^\d+(\.\d+)*\.?\s*")
NON_ALNUM_RE = re.compile(r"[^a-z0-9]+")


def normalize(name, strip_leading_number=False):
    name = (name or "").strip()
    if strip_leading_number:
        name = LEADING_NUMBER_RE.sub("", name)
    name = name.lower()
    name = NON_ALNUM_RE.sub(" ", name)
    return re.sub(r"\s+", " ", name).strip()


def is_generic_control_room_label(name):
    return normalize(name) == "control room"


def classify_group(station, category, rows):
    """rows: list of CSV row dicts sharing this (Station, Category).
    Returns (artifact_rows, real_rows)."""
    if category not in ("Control Room", "Switchyard"):
        return [], rows

    extras = [r for r in rows if r["Issue"] == "Extra"]
    missing = [r for r in rows if r["Issue"] == "Missing"]
    others = [r for r in rows if r["Issue"] not in ("Extra", "Missing")]

    station_norm = normalize(station)
    used_extra_idx = set()
    used_missing_idx = set()
    artifact_pairs = []

    for mi, m in enumerate(missing):
        m_norm = normalize(m["Name"])
        match_ei = None

        # Pattern C: identical once normalized (punctuation/case/spacing only)
        for ei, e in enumerate(extras):
            if ei in used_extra_idx:
                continue
            if normalize(e["Name"]) == m_norm and m_norm != "":
                match_ei = ei
                break

        # Pattern A/B: generic "Control Room" missing label, paired against
        # an extra whose name (leading number stripped) is the org's own name
        if match_ei is None and category == "Control Room" and is_generic_control_room_label(m["Name"]):
            for ei, e in enumerate(extras):
                if ei in used_extra_idx:
                    continue
                if normalize(e["Name"], strip_leading_number=True) == station_norm and station_norm != "":
                    match_ei = ei
                    break

        if match_ei is not None:
            used_extra_idx.add(match_ei)
            used_missing_idx.add(mi)
            artifact_pairs.append(missing[mi])
            artifact_pairs.append(extras[match_ei])

    real_rows = (
        [e for i, e in enumerate(extras) if i not in used_extra_idx]
        + [m for i, m in enumerate(missing) if i not in used_missing_idx]
        + others
    )
    return artifact_pairs, real_rows


def main():
    with open(INPUT_CSV, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    groups = defaultdict(list)
    for r in rows:
        groups[(r["Station"], r["Category"])].append(r)

    artifacts, real = [], []
    for (station, category), group_rows in groups.items():
        a, r = classify_group(station, category, group_rows)
        artifacts.extend(a)
        real.extend(r)

    fieldnames = ["Station", "Category", "Issue", "Name", "Details"]

    with open(REAL_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(real)

    with open(ARTIFACTS_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(artifacts)

    print("=" * 70)
    print("  Naming-Artifact Filter Summary")
    print("=" * 70)
    print(f"  Input rows              : {len(rows)}")
    print(f"  -> Real discrepancies   : {len(real)}")
    print(f"  -> Naming artifacts     : {len(artifacts)}")
    print(f"  Check: {len(real)} + {len(artifacts)} = {len(real) + len(artifacts)} (should equal input)")
    print("=" * 70)

    by_cat = defaultdict(lambda: [0, 0])
    for r in real:
        by_cat[r["Category"]][0] += 1
    for r in artifacts:
        by_cat[r["Category"]][1] += 1
    print(f"  {'Category':<15} {'Real':>6} {'Artifact':>9}")
    for cat, (real_n, art_n) in sorted(by_cat.items()):
        print(f"  {cat:<15} {real_n:>6} {art_n:>9}")
    print("=" * 70)
    print(f"  Written: {REAL_CSV}")
    print(f"  Written: {ARTIFACTS_CSV}")


if __name__ == "__main__":
    main()
