"""Two read-only deliverables for the IPP-removal / CPSU-population request:

1. reports/ipp_removal_impact.csv -- every place "IPP" is referenced,
   found by grepping the codebase and inspecting the two FKs that point at
   organization_categories.id, checked against live data (not assumed from
   the migration files alone).

2. reports/cpsu_population_proposal.csv -- candidate orgs for the CPSU
   category. NOT a blind keyword sweep: reclassify_organization_types.py's
   own docstring defines CPSU as "central-PSU entities whose function isn't
   generation/transmission/distribution" (its own example: POSOCO's NLDC,
   NTPC's trading subsidiary) -- i.e. the ownership-based bucket is meant
   for entities WITHOUT a more specific functional category, not a
   replacement for the functional categories (Thermal/Hydel/RE Generators/
   Transmission Utility) that already correctly describe what an org does.
   So every NTPC/SJVN match is shown, but only the ones with no operational
   function of their own (trading subsidiaries, testing labs) are proposed
   as CONFIRM; actual power stations and Power Grid itself (function IS
   transmission) are shown as EXCLUDED-BY-DESIGN, with the reasoning
   spelled out, not just silently dropped. Power Grid is listed as
   USER-CONFIRMED per explicit instruction even though it conflicts with
   that design principle -- flagged, not silently applied either way.

No writes -- proposal only.

Usage:
  python scripts/build_ipp_cpsu_proposal.py
"""

import csv
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from models.organization import Organization
from models.organization_category import OrganizationCategory
from models.email_group_filter import EmailGroupFilter

IPP_IMPACT_CSV = "reports/ipp_removal_impact.csv"
CPSU_PROPOSAL_CSV = "reports/cpsu_population_proposal.csv"


def build_ipp_impact():
    rows = []
    with app.app_context():
        ipp_cat = OrganizationCategory.query.filter_by(category_name="IPP").first()
        ipp_id = ipp_cat.id if ipp_cat else None

        org_count = Organization.query.filter_by(category_id=ipp_id).count() if ipp_id else 0
        rows.append({
            "reference_type": "FK: organizations.category_id (ON DELETE SET NULL)",
            "location": "organizations table, live data",
            "detail": f"{org_count} organizations currently reference IPP (category id {ipp_id})",
            "risk_if_deleted": "None -- 0 rows to null out" if org_count == 0 else f"{org_count} orgs would have category_id set to NULL",
        })

        filter_count = EmailGroupFilter.query.filter_by(category_id=ipp_id).count() if ipp_id else 0
        rows.append({
            "reference_type": "FK: email_group_filters.category_id (ON DELETE CASCADE)",
            "location": "email_group_filters table, live data",
            "detail": f"{filter_count} filter rows currently reference IPP "
                      f"(table has {EmailGroupFilter.query.count()} rows total, all categories)",
            "risk_if_deleted": "None -- 0 filter rows exist" if filter_count == 0 else f"{filter_count} filter rows would be CASCADE-DELETED, possibly breaking a saved dynamic group",
        })

    rows.append({
        "reference_type": "Hardcoded dict: TELEPHONE_DIRECTORY_CATEGORY_DESCRIPTIONS",
        "location": "routes/user_routes.py ~line 106",
        "detail": "'IPP': 'Independent power producers' -- presentational blurb, keyed by "
                  "category_name, looked up via .get(name, '') against LIVE OrganizationCategory "
                  "rows only (the calling code iterates the DB table, never this dict directly)",
        "risk_if_deleted": "None -- becomes an unreachable dead dict entry, no error, no code change required (optional cleanup only)",
    })

    rows.append({
        "reference_type": "Hardcoded list: CATEGORY_SLUGS -> STANDARD_GROUPS",
        "location": "services/email_distribution_service.py ~line 197",
        "detail": "('ipp', 'IPP', 'IPPs') -- creates a permanent 'IPPs' tab/slug in the Email "
                  "Distribution browse UI, resolved via resolve_category_contacts('IPP', ...) "
                  "which joins Organization to OrganizationCategory by name",
        "risk_if_deleted": "None functionally -- already shows 0 results today (0 orgs are IPP) and "
                  "will continue to show 0 results after deletion (the join just won't match). "
                  "BUT the dead 'IPPs' tab itself stays visible in the UI unless this tuple is also "
                  "removed -- that's a separate, optional code edit, your call whether you want it.",
    })

    rows.append({
        "reference_type": "Seed migration: migrations/009_organization_types.sql",
        "location": "migrations/009_organization_types.sql line 18",
        "detail": "INSERT ... ('IPP', NULL) ... ON CONFLICT (category_name) DO NOTHING",
        "risk_if_deleted": "IMPORTANT: deleting the live row is NOT permanent against a full "
                  "migration replay -- ON CONFLICT DO NOTHING only skips the insert when a row "
                  "ALREADY exists; if this row is deleted and migrations are ever re-run from "
                  "scratch (fresh DB / CI setup), 'IPP' would silently come back. A durable "
                  "removal needs a new migration that explicitly deletes it, not just a live DELETE.",
    })

    os.makedirs("reports", exist_ok=True)
    with open(IPP_IMPACT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["reference_type", "location", "detail", "risk_if_deleted"])
        w.writeheader()
        w.writerows(rows)
    return rows


CENTRAL_PSU_KEYWORDS = [
    "ntpc", "nhpc", "sjvn", "posoco", "pgcil", "power finance", "rural electrification",
    "thdc", "damodar valley", "bhel", "neepco", "satluj jal", "nlc india", "nuclear power corporation",
]

# Orgs with a real operational function of their own -- correctly excluded
# from CPSU by the category's own design (see module docstring), listed
# explicitly rather than silently dropped.
FUNCTIONAL_EXCLUSIONS = {
    6309: "Power Grid Corporation Of India Limited -- function IS transmission; Transmission "
          "Utility is the more specific, already-correct category per CPSU's own design principle.",
    6312: "POWERGRID WRTS-I, Nagpur -- same reasoning as Power Grid itself.",
    6313: "POWERGRID WRTS-II, Vadodara -- same reasoning as Power Grid itself.",
}
NTPC_STATION_NOTE = (
    "NTPC generating station -- reclassify_organization_types.py deliberately classified NTPC's "
    "WR stations by fuel type (Thermal), not by CPSU ownership (see NTPC_THERMAL_IDS in that "
    "script's docstring) -- moving to CPSU would undo that deliberate, more specific classification."
)
CENTRAL_PSU_RE_NOTE = (
    "Central-PSU renewable-generation subsidiary -- function is RE generation; RE Generators is "
    "the correct functional category, same reasoning as the Thermal stations above."
)


def build_cpsu_proposal():
    rows = []
    with app.app_context():
        cat_name = {c.id: c.category_name for c in OrganizationCategory.query.all()}
        orgs = {o.id: o for o in Organization.query.all()}

        # Power Grid -- user-confirmed, flagged as a policy override
        for org_id in (6309, 6312, 6313):
            o = orgs[org_id]
            rows.append({
                "org_id": o.id, "org_name": o.organization_name,
                "current_category": cat_name.get(o.category_id), "proposed_category": "CPSU",
                "status": "USER_CONFIRMED_OVERRIDE",
                "rationale": FUNCTIONAL_EXCLUSIONS[org_id] + " Proposed anyway because you "
                             "explicitly confirmed this one directly -- flagged so the conflict "
                             "with CPSU's own stated design principle is visible before you apply it.",
            })

        matched_ids = set()
        for o in orgs.values():
            name_lower = o.organization_name.lower()
            if any(kw in name_lower for kw in CENTRAL_PSU_KEYWORDS):
                matched_ids.add(o.id)
        matched_ids -= {6309, 6312, 6313}

        for org_id in sorted(matched_ids):
            o = orgs[org_id]
            current = cat_name.get(o.category_id)
            name_lower = o.organization_name.lower()

            if current == "Thermal":
                rows.append({
                    "org_id": o.id, "org_name": o.organization_name, "current_category": current,
                    "proposed_category": "(no change)", "status": "EXCLUDED_BY_DESIGN",
                    "rationale": NTPC_STATION_NOTE,
                })
            elif current == "RE Generators":
                rows.append({
                    "org_id": o.id, "org_name": o.organization_name, "current_category": current,
                    "proposed_category": "(no change)", "status": "EXCLUDED_BY_DESIGN",
                    "rationale": CENTRAL_PSU_RE_NOTE,
                })
            elif o.organization_name == "NTPC Vidyut Vyapar Nigam Limited":
                rows.append({
                    "org_id": o.id, "org_name": o.organization_name, "current_category": current,
                    "proposed_category": "CPSU", "status": "CONFIRM",
                    "rationale": "NTPC's trading subsidiary -- no generation/transmission/"
                                 "distribution function of its own; this is the exact example "
                                 "reclassify_organization_types.py's own docstring names for CPSU.",
                })
            elif o.organization_name == "NTPC":
                rows.append({
                    "org_id": o.id, "org_name": o.organization_name, "current_category": current,
                    "proposed_category": "CPSU", "status": "AMBIGUOUS",
                    "rationale": "Bare 'NTPC' entry (currently Thermal) -- unclear whether this "
                                 "represents the parent corporation (CPSU-appropriate) or an "
                                 "unspecific generic station entry (Thermal-appropriate, like its "
                                 "named stations). Needs your call, not inferred.",
                })
            elif "sail power company" in name_lower:
                rows.append({
                    "org_id": o.id, "org_name": o.organization_name, "current_category": current,
                    "proposed_category": "CPSU?", "status": "AMBIGUOUS",
                    "rationale": "NTPC-SAIL joint venture -- IS a generation company by function "
                                 "(name says 'Power Company'), and only partially NTPC-owned (JV "
                                 "with SAIL), so it doesn't cleanly fit CPSU's 'not "
                                 "generation/transmission/distribution' definition either. Could "
                                 "arguably be Generation Company instead -- your call.",
                })
            else:
                rows.append({
                    "org_id": o.id, "org_name": o.organization_name, "current_category": current,
                    "proposed_category": "?", "status": "NEEDS_REVIEW",
                    "rationale": "Matched a central-PSU keyword but doesn't fit the patterns above -- review manually.",
                })

        # Government/DAE-linked entities found while eyeballing the Others
        # bucket -- structurally different from a corporate PSU (a
        # government department / research institute / JV testing lab, not
        # an incorporated PSU), so listed as low-confidence, not proposed.
        for org_id, note in {
            6496: "Heavy Water Board -- Dept. of Atomic Energy government body, not an "
                  "incorporated PSU in the usual sense.",
            6495: "BARC Facility -- Bhabha Atomic Research Centre, a DAE research institution, "
                  "not an incorporated PSU.",
            6498: "National High Power Test Laboratory -- a joint testing-lab entity promoted by "
                  "several central PSUs collectively (CPRI/NTPC/PGCIL/DVC/NPCIL), not itself a "
                  "single central PSU.",
        }.items():
            o = orgs[org_id]
            rows.append({
                "org_id": o.id, "org_name": o.organization_name,
                "current_category": cat_name.get(o.category_id), "proposed_category": "CPSU?",
                "status": "LOW_CONFIDENCE", "rationale": note,
            })

    os.makedirs("reports", exist_ok=True)
    with open(CPSU_PROPOSAL_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["org_id", "org_name", "current_category", "proposed_category", "status", "rationale"])
        w.writeheader()
        w.writerows(rows)
    return rows


def main():
    ipp_rows = build_ipp_impact()
    cpsu_rows = build_cpsu_proposal()

    print("=" * 70)
    print(f"  IPP removal impact -- {len(ipp_rows)} references found -> {IPP_IMPACT_CSV}")
    print("=" * 70)
    print(f"  CPSU population proposal -- {len(cpsu_rows)} rows -> {CPSU_PROPOSAL_CSV}")
    from collections import Counter
    for status, count in Counter(r["status"] for r in cpsu_rows).items():
        print(f"    - {status}: {count}")
    print("=" * 70)


if __name__ == "__main__":
    main()
