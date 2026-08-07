"""
Two independent, source-grounded passes over every Organization, for the
cascading Type -> State -> Organization dropdowns:

STATE -- read from WR_DB_Ready_Final_Verified_v3.xlsx (project root), the
workbook confirmed to be this app's authoritative organization data
source. Its `organizations`/`sub_organizations` sheets have a `state`
column that matches 1:1 by name against the live Organization table. Only
set where the workbook has a value (~44% coverage) -- never invented.

TYPE -- remaps Organization.category_id to the 12-type taxonomy (Transmission
Utility/RLDC/State SLDC/CPSU/IPP/RE Generators/DISCOM/Generation Company/
Thermal/Hydel/Nuclear/Others, seeded by migrations/009 + 010 + 011 + 012).
No file in this project has an explicit Organization
Type column, so Type is derived from structured signals actually present in
the data -- organization name, the existing `region` grouping label, and
the source workbook's `sub_organizations.address` text -- rather than a
blind keyword classifier or a memorized list of company facts. Rules, in
priority order (see resolve_type / FUEL_KEYWORDS):

  1. State SLDC / DISCOM -- small explicit org-id sets (STATE_SLDC_IDS /
     DISCOM_IDS below), hand-confirmed by direct DB inspection, the same
     size and rigor as classify_organizations.py's own precedent.
  2. Transmission Utility (merges the earlier separate CTU/STU types) --
     name contains "state electricity transmission" (an institutional
     phrase; only genuine state transmission utilities are named this way
     -- the ~26 privately-owned TBCB transmission SPVs in this dataset use
     generic corporate names like "... Transmission Limited"/"...
     Transco Ltd" instead, which this phrase does NOT match, so they
     correctly stay Others rather than being misclassified as
     government-owned).
  3. Nuclear -- region == the Nuclear Power Stations grouping.
  4. RLDC -- WRLDC itself, or name contains "Regional Load Despatch
     Centre" (covers NRLDC/ERLDC/SRLDC/NERLDC). POSOCO's *National* Load
     Despatch Centre is deliberately excluded (National != Regional).
  5. RE Generators -- region == the ISTS-connected RE Generators grouping
     (per classify_organizations.py's own hand review, none of these are
     PSU-owned), or name contains "renewable energy". Split out from IPP
     as its own type -- IPP is currently unpopulated by any rule (no
     source signal confidently asserts *non-RE* private ownership for the
     remaining generator companies; see Generation Company below), left
     available for manual classification via Manage Organizations > Edit.
  6. Thermal / Hydel -- keyword match (FUEL_KEYWORDS) against the
     organization's own name plus its matched address text from the
     workbook's sub_organizations sheet -- e.g. "... HPS"/"Hydro"/"NHDC"/
     "Pumped Storage" -> Hydel, "... TPS"/"STPP"/"CCPP"/"Gas & Power" ->
     Thermal. These are standard, unambiguous Indian power-sector station
     abbreviations already present in the source data, not guesses.
     NTPC_THERMAL_IDS covers NTPC's Western Region stations explicitly --
     NTPC has zero hydro/nuclear assets in this region, but none of its
     station names contain a fuel-type keyword, so this is the one
     narrow, well-documented exception handled by explicit id rather than
     keyword.
  7. Generation Company -- region == one of GENERATING_STATION_REGIONS
     ("Other Generating Stations in Western Region", "Black Start
     Facilitated Stations" -- the source data's own section groupings for
     generating stations, a structural fact, not a guess), or name
     contains "power generation" (e.g. Madhya Pradesh Power Generation
     Company Limited). Used only when no more specific fuel-type keyword
     matched in step 6 -- i.e. "this is confirmed to be a generating
     station/company, fuel type unconfirmed" rather than leaving a known
     generator in the generic Others bucket.
  8. CPSU -- a small explicit set (CPSU_IDS) for central-PSU entities
     whose function isn't generation/transmission/distribution (POSOCO's
     NLDC, NTPC's trading subsidiary) -- i.e. exactly the "does not
     already belong to a more appropriate operational category" case.
  9. Everything else -> Others: private transmission SPVs (no CTU/STU
     evidence), government/statutory bodies, section-header rows, and
     anything else with no structural signal anywhere in the source data.
     Guessing further (e.g. assuming a private company's fuel type or
     ownership from general knowledge rather than the uploaded files)
     risks silently corrupting real classification data -- left for an
     admin to confirm via Manage Organizations > Edit instead, now much
     faster with the cascading
     Type -> State -> Organization picker this data feeds.

After remapping every organization, deletes any now-fully-unreferenced old
category rows and writes
reports/organization_type_and_state_<timestamp>.csv listing every
organization's old category, new type, derived state, and a NEEDS_REVIEW
flag for anything defaulted to Others.

Usage:
  python scripts/reclassify_organization_types.py          # dry run
  python scripts/reclassify_organization_types.py --apply  # commit
"""

import csv
import os
import re
import sys
from datetime import datetime

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from models import db
from models.organization import Organization
from models.organization_category import OrganizationCategory

WORKBOOK_PATH = "WR_DB_Ready_Final_Verified_v3.xlsx"

# RLDC is deliberately excluded -- it's unchanged between the old and new
# taxonomy (reused as-is by resolve_type(), never renamed), so it's not an
# "old, now-obsolete" category to delete. "Transmission Utility" is
# likewise excluded despite being an original-taxonomy name -- it was
# deleted once already (nothing referenced it) and migrations/012 re-adds
# it fresh as the CTU+STU merge target, so it must NOT be swept up here as
# "old" or this cleanup step would delete the very row just assigned to.
OLD_CATEGORIES = [
    "Generator", "RE Generator", "Distribution Company",
    "SLDC", "Central Utility", "Government", "Private Utility", "Other",
    "CTU", "STU",  # merged into Transmission Utility -- see resolve_type
]

NUCLEAR_REGION = "Nuclear Power Stations in Western Region"
RE_GENERATOR_REGION = "ISTS Connected RE Generators in Western Region"
# The source data's own section groupings for generating stations -- used
# by the Generation Company fallback rule (see resolve_type).
GENERATING_STATION_REGIONS = {
    "Other Generating Stations in Western Region",
    "Black Start Facilitated Stations",
}

# Explicit org-id allow-lists -- hand-confirmed by direct DB inspection,
# used where a reliable name/region/keyword rule can't reach but the
# classification is well-documented and unambiguous. Kept deliberately
# small (a handful of entries each), same rigor as
# classify_organizations.py's own precedent -- not a large hardcoded
# name list standing in for a missing rule.
STATE_SLDC_IDS = {5155, 5156, 5157, 5158, 5159, 5160, 5161, 5202, 5203}
DISCOM_IDS = {5204, 5355}
# NTPC's Western Region generating stations -- all thermal (coal/gas), zero
# hydro/nuclear assets in this region -- but none of their station names
# contain a fuel-type keyword, so FUEL_KEYWORDS can't reach them.
NTPC_THERMAL_IDS = {5162, 5163, 5207, 5208, 5209, 5210, 5211, 5212, 5213, 5214, 5215, 5216}
# Central PSU entities whose function is grid operations/trading, not
# generation/transmission/distribution -- the "no more appropriate
# operational category" case.
CPSU_IDS = {5171, 5357}  # POSOCO National Load Despatch Centre; NTPC Vidyut Vyapar Nigam (trading arm)

# Industrial consumers grouped under the RE Generators region by the
# original import (they connect via the same ISTS pooling infrastructure)
# but are not themselves generators -- classify_organizations.py's earlier
# hand review already identified these specifically as "Private Utility"
# (industrial consumers), not generators. The new taxonomy has no
# industrial-consumer bucket, so these fall through to Others rather than
# being mislabeled RE Generators.
NON_GENERATOR_CONSUMER_IDS = {5278, 5279, 5280, 5290, 5291}  # Ambuja Cements x2, Adani Hazira Port, Adani Ports & SEZ x2

# (keyword, resulting type) -- checked as whole-word, case-insensitive
# matches against an organization's own name plus its matched
# sub_organizations.address text from the source workbook. Order matters:
# first match wins.
FUEL_KEYWORDS = [
    (r"\bnhdc\b", "Hydel"),              # Narmada Hydroelectric Development Corporation
    (r"pumped storage", "Hydel"),
    (r"\bhydro\b", "Hydel"),             # "Hydro Power Station", "Mini Hydro"
    (r"\bhps\b", "Hydel"),               # standard "<name> HPS" abbreviation
    (r"\bstpp\b", "Thermal"),            # Super Thermal Power Project
    (r"\bccpp\b", "Thermal"),            # Combined Cycle Power Plant (gas-thermal)
    (r"\bthermal\b", "Thermal"),
    (r"gas\s*&\s*power", "Thermal"),
    (r"gas tps", "Thermal"),
    (r"\btps\b", "Thermal"),             # standard "<name> TPS"/"...MTPS" abbreviation
]


def _normalize(name):
    return re.sub(r"\s+", " ", (name or "").strip()).lower()


def load_from_workbook():
    """Returns (state_map, address_map): both {normalized_name: value},
    read from the organizations/sub_organizations sheets. address_map only
    covers sub_organizations (individual stations/entities), used as the
    fuel-type keyword source for FUEL_KEYWORDS."""
    wb = openpyxl.load_workbook(WORKBOOK_PATH, read_only=True, data_only=True)
    state_map, address_map = {}, {}
    for sheet_name in ("organizations", "sub_organizations"):
        ws = wb[sheet_name]
        rows = ws.iter_rows(values_only=True)
        header = [str(h or "").strip().lower() for h in next(rows)]
        name_col = next(i for i, h in enumerate(header) if h.startswith("name"))
        state_col = next(i for i, h in enumerate(header) if h.startswith("state"))
        addr_col = next((i for i, h in enumerate(header) if h.startswith("address")), None)
        for row in rows:
            name = row[name_col]
            if not name:
                continue
            key = _normalize(name)
            state = row[state_col]
            if state:
                state_map[key] = str(state).strip()
            if addr_col is not None and row[addr_col]:
                address_map[key] = str(row[addr_col])
    return state_map, address_map


def resolve_fuel_type(org, address_map):
    """None if no keyword matches, or if BOTH Thermal and Hydel keywords
    match (e.g. "UKAI Power Station (Thermal + Hydro)" -- a genuinely
    mixed-fuel station per the source text itself) -- picking one in that
    case would silently discard information the file is actually telling
    us, so it falls through to Others instead."""
    text = f"{org.organization_name} {address_map.get(_normalize(org.organization_name), '')}".lower()
    matched_types = {
        fuel_type for pattern, fuel_type in FUEL_KEYWORDS
        if re.search(pattern, text, re.IGNORECASE)
    }
    if len(matched_types) == 1:
        return matched_types.pop()
    return None


def resolve_type(org, address_map):
    """Keyed only on org.id / org.region / org.organization_name and the
    workbook-derived address_map -- none of which this script mutates --
    so this is safe to re-run any number of times regardless of what
    category an organization currently holds."""
    if org.id in STATE_SLDC_IDS:
        return "State SLDC"
    if org.id in DISCOM_IDS:
        return "DISCOM"
    if "state electricity transmission" in _normalize(org.organization_name):
        # CTU and STU are merged into a single Transmission Utility type --
        # this is currently the only confidently-identifiable entity in
        # either (no POWERGRID/CTU entity exists in this dataset).
        return "Transmission Utility"
    if org.region == NUCLEAR_REGION:
        return "Nuclear"
    name_norm = _normalize(org.organization_name)
    if name_norm == "western region load despatch centre" or \
            "regional load despatch centre" in name_norm:
        return "RLDC"
    if org.id not in NON_GENERATOR_CONSUMER_IDS and \
            (org.region == RE_GENERATOR_REGION or "renewable energy" in name_norm):
        return "RE Generators"
    if org.id in NTPC_THERMAL_IDS:
        return "Thermal"
    fuel_type = resolve_fuel_type(org, address_map)
    if fuel_type:
        return fuel_type
    if org.region in GENERATING_STATION_REGIONS or "power generation" in name_norm:
        # Confirmed generating station/company (the source data's own
        # section grouping says so), fuel type just not determinable from
        # any keyword -- a more specific classification than Others
        # without guessing which fuel it actually is.
        return "Generation Company"
    if org.id in CPSU_IDS:
        return "CPSU"
    return "Others"


def run(apply: bool) -> None:
    print("DRY RUN\n" if not apply else "APPLY\n")
    state_map, address_map = load_from_workbook()
    print(f"Loaded {len(state_map)} state mappings and {len(address_map)} address "
          f"entries from {WORKBOOK_PATH}\n")

    with app.app_context():
        categories_by_name = {c.category_name: c for c in OrganizationCategory.query.all()}

        rows_for_report = []
        type_changed = 0
        state_changed = 0
        defaulted = 0

        for org in Organization.query.order_by(Organization.id).all():
            old_category_name = org.category.category_name if org.category else None
            new_type_name = resolve_type(org, address_map)
            target = categories_by_name[new_type_name]
            needs_review = new_type_name == "Others"

            derived_state = state_map.get(_normalize(org.organization_name))

            rows_for_report.append({
                "org_id": org.id,
                "organization_name": org.organization_name,
                "old_category": old_category_name or "(uncategorized)",
                "new_type": new_type_name,
                "derived_state": derived_state or "",
                "needs_review": "YES" if needs_review else "",
            })

            if org.category_id != target.id:
                print(f"  [{org.id}] {org.organization_name[:55]:<55} "
                      f"{(old_category_name or '(uncategorized)'):<22} -> {new_type_name}")
                if apply:
                    org.category_id = target.id
                type_changed += 1
                if needs_review:
                    defaulted += 1

            if derived_state and org.state != derived_state:
                state_changed += 1
                if apply:
                    org.state = derived_state

        report_dir = "reports"
        os.makedirs(report_dir, exist_ok=True)
        report_path = os.path.join(
            report_dir,
            f"organization_type_and_state_{datetime.now():%Y%m%d_%H%M}.csv",
        )
        with open(report_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=[
                "org_id", "organization_name", "old_category", "new_type",
                "derived_state", "needs_review",
            ])
            writer.writeheader()
            writer.writerows(rows_for_report)
        print(f"\nReport written: {report_path}")

        if apply:
            old_ids = [c.id for name, c in categories_by_name.items() if name in OLD_CATEGORIES]
            still_referenced = Organization.query.filter(Organization.category_id.in_(old_ids)).count()
            if still_referenced == 0 and old_ids:
                OrganizationCategory.query.filter(OrganizationCategory.id.in_(old_ids)).delete(
                    synchronize_session=False
                )
                print(f"Deleted {len(old_ids)} old (now-unreferenced) category rows.")
            elif old_ids:
                print(f"NOT deleting old categories -- {still_referenced} organization(s) "
                      f"still reference them.")

            db.session.commit()
            print(f"\nDone -- {type_changed} organizations retyped "
                  f"({defaulted} defaulted to 'Others' for manual review), "
                  f"{state_changed} states set.")
        else:
            print(f"\n[Dry run -- {type_changed} orgs would be retyped "
                  f"({defaulted} defaulting to 'Others'), {state_changed} states "
                  f"would be set. Rerun with --apply]")


if __name__ == "__main__":
    run(apply="--apply" in sys.argv)
