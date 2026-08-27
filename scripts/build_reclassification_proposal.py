"""Dry-run proposal for the org-category reclassification requested on top
of scripts/reclassify_organization_types.py's existing 12-type taxonomy.
Read-only -- writes reports/reclassification_proposal.csv only, no writes
to organizations.category_id or organization_categories.

Four rules, applied in this priority order per org:

  1. AMBIGUOUS -- name contains both "distribution" and "transmission"
     (can't tell which was meant).
  2. AMBIGUOUS -- name matches a name-based rule (distribution/transmission)
     but the org's current_category is anything other than "Others" --
     i.e. it already carries a specific, presumably hand-confirmed
     classification (including State SLDC / RE Generators, which
     reclassify_organization_types.py's own docstring says were hand-
     confirmed by direct DB inspection) that this rule would silently
     override. Flagged for review instead of auto-applied, per instruction.
  3. name contains "distribution" -> DISCOM (only when current_category ==
     "Others", i.e. currently unclassified).
  4. name contains "transmission" -> Transmission Utility (same condition).
     NOTE: this is a materially broader rule than
     reclassify_organization_types.py's existing Transmission Utility rule,
     which deliberately matches only the institutional phrase "state
     electricity transmission" -- not bare "transmission" -- specifically
     to exclude ~26 privately-owned TBCB transmission SPVs (e.g. "...
     Transmission Limited" / "... Transco Ltd") that the prior script kept
     in Others on purpose. Restricting rule 3/4 to current_category ==
     "Others" only limits the blast radius to genuinely-unclassified orgs;
     it does not by itself resolve that policy conflict -- see the summary
     printed at the end for how many of those private SPVs this rule would
     now pull in.
  5. category == "State SLDC" -> label rename to "SLDC" (same orgs, no
     category re-assignment -- just the category_name text changes).
  6. category in ("Generation Company", "RE Generators") -> merged to
     "Generation Company" (RE Generators orgs move; Generation Company
     orgs are listed as already-there for completeness).

Only orgs with an actual proposed change (or an ambiguous flag) are
written -- orgs untouched by any rule are omitted from the CSV, not
silently included as no-ops.

Usage:
  python scripts/build_reclassification_proposal.py
"""

import csv
import os
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from models.organization import Organization
from models.organization_category import OrganizationCategory

OUT_CSV = "reports/reclassification_proposal.csv"


def build():
    with app.app_context():
        orgs = Organization.query.all()
        cat_name_by_id = {c.id: c.category_name for c in OrganizationCategory.query.all()}

        rows = []
        counts = Counter()

        for org in orgs:
            name_lower = (org.organization_name or "").lower()
            current_category = cat_name_by_id.get(org.category_id, "(none)")

            has_distribution = "distribution" in name_lower
            has_transmission = "transmission" in name_lower

            row = {
                "org_id": org.id,
                "org_name": org.organization_name,
                "current_category": current_category,
                "proposed_category": "",
                "matched_rule": "",
                "flag": "",
            }

            if has_distribution and has_transmission:
                row["matched_rule"] = "name contains both 'distribution' and 'transmission'"
                row["flag"] = "AMBIGUOUS: name matches both DISCOM and Transmission Utility patterns"
                rows.append(row)
                counts["ambiguous_both_keywords"] += 1
                continue

            name_rule = None
            if has_distribution:
                name_rule = ("DISCOM", "name contains 'distribution' -> DISCOM")
            elif has_transmission:
                name_rule = ("Transmission Utility", "name contains 'transmission' -> Transmission Utility")

            if name_rule is not None:
                if current_category == name_rule[0]:
                    # Already sitting in the category this rule would propose --
                    # not a conflict, just redundant confirmation. Skip entirely
                    # rather than flag as ambiguous or list as a no-op change.
                    counts["already_matches_name_rule"] += 1
                    continue
                if current_category != "Others":
                    row["matched_rule"] = name_rule[1]
                    row["flag"] = (
                        f"AMBIGUOUS: name matches this rule but org is already classified "
                        f"'{current_category}' (a DIFFERENT category than the rule proposes), "
                        f"not 'Others' -- rule would override an existing assignment"
                    )
                    rows.append(row)
                    counts["ambiguous_overrides_existing"] += 1
                    continue
                else:
                    row["proposed_category"] = name_rule[0]
                    row["matched_rule"] = name_rule[1]
                    rows.append(row)
                    counts[f"applied_{name_rule[0]}"] += 1
                    continue

            if current_category == "State SLDC":
                row["proposed_category"] = "SLDC"
                row["matched_rule"] = "category = State SLDC -> renamed to SLDC"
                rows.append(row)
                counts["applied_SLDC_rename"] += 1
                continue

            if current_category == "RE Generators":
                row["proposed_category"] = "Generation Company"
                row["matched_rule"] = "category in (Generation Company, RE Generators) -> merged to Generation Company"
                rows.append(row)
                counts["applied_Generation_merge_from_RE"] += 1
                continue

            if current_category == "Generation Company":
                row["proposed_category"] = "Generation Company"
                row["matched_rule"] = "category in (Generation Company, RE Generators) -> merged to Generation Company (already here)"
                rows.append(row)
                counts["already_Generation_Company"] += 1
                continue

    fieldnames = ["org_id", "org_name", "current_category", "proposed_category", "matched_rule", "flag"]
    os.makedirs("reports", exist_ok=True)
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)

    print("=" * 70)
    print("  Org-Category Reclassification Proposal -- counts (nothing written to the DB)")
    print("=" * 70)
    print(f"  Total organizations scanned : {len(orgs)}")
    print(f"  Total rows written          : {len(rows)}  -> {OUT_CSV}")
    print("-" * 70)
    for key, count in sorted(counts.items()):
        print(f"    - {key:<40}: {count}")
    print("=" * 70)
    return rows, counts


if __name__ == "__main__":
    build()
