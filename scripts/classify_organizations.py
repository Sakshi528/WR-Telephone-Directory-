"""
Backfills Organization.category_id for the Email Distribution Lists module
(see migrations/003_email_distribution_lists.sql). Every organization in the
DB was inspected by name and assigned one of the ten organization_categories
seeded by that migration. Section-header rows that aren't real
organizations (e.g. "Other Generating Stations in Western Region", used as
parent_id groupings) and a few genuinely ambiguous entries (QCA roles) are
left as 'Other' -- an admin can reclassify any of these later via
Manage Organizations > Edit.

This is a one-off, hand-reviewed mapping (same approach as
set_org_regions.py), not a keyword heuristic -- a regex like "%NTPC%" would
misfile "NTPC Vidyut Vyapar Nigam Limited" (a PSU trading arm) as a
Generator alongside NTPC's actual power stations.

Usage:
  python scripts/classify_organizations.py          # dry run
  python scripts/classify_organizations.py --apply  # commit
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from models import db
from models.organization import Organization
from models.organization_category import OrganizationCategory

# org_id -> organization_categories.category_name
ORG_CATEGORIES = {

    # ── RLDC / NLDC ──────────────────────────────────────────────────────────
    619:  "RLDC",   # Western Region Load Despatch Centre (WRLDC itself)
    5171: "Central Utility",  # POSOCO National Load Despatch Centre (NLDC)
    5172: "RLDC",   # Northern Regional Load Despatch Centre
    5173: "RLDC",   # Eastern Regional Load Despatch Centre
    5174: "RLDC",   # Southern Regional Load Despatch Centre
    5175: "RLDC",   # North Eastern Regional Load Despatch Centre

    # ── Government / statutory bodies ───────────────────────────────────────
    5153: "Government",  # Western Regional Power Committee
    5205: "Government",  # MP Power Management Co. (state holding/bulk purchase agency)
    5353: "Government",  # BARC Facility
    5354: "Government",  # Heavy Water Board
    5356: "Government",  # National High Power Test Laboratory

    # ── SLDCs ────────────────────────────────────────────────────────────────
    5155: "SLDC",  # Maharashtra SLDC
    5156: "SLDC",  # Gujarat SLDC
    5157: "SLDC",  # Madhya Pradesh SLDC
    5158: "SLDC",  # Chhattisgarh SLDC
    5159: "SLDC",  # Goa Electricity Department
    5160: "SLDC",  # Daman and Diu
    5161: "SLDC",  # Dadra and Nagar Haveli
    5202: "SLDC",  # State Load Despatch Centre, MSETCL
    5203: "SLDC",  # Area Load Despatch Centre, MSETCL

    # ── Transmission Utility ─────────────────────────────────────────────────
    5201: "Transmission Utility",  # Maharashtra State Electricity Transmission Co.
    5325: "Transmission Utility",  # AMNS (Essar) Power Transmission Co.
    5326: "Transmission Utility",  # India Grid Trust (IndiGrid)
    5327: "Transmission Utility",  # Bhopal-Dhule Transmission Co.
    5328: "Transmission Utility",  # Jabalpur Transmission Co.
    5329: "Transmission Utility",  # RAPP Transmission Co.
    5330: "Transmission Utility",  # Raichur Sholapur Transmission Co.
    5331: "Transmission Utility",  # Odisha Generation Phase-II transmission
    5332: "Transmission Utility",  # Khargone Transmission
    5334: "Transmission Utility",  # Sterlite Power Transmission
    5335: "Transmission Utility",  # Lakadia Vadodara Transmission Project
    5336: "Transmission Utility",  # Mumbai Urja Marg
    5337: "Transmission Utility",  # Goa Tamnar Transmission Project
    5338: "Transmission Utility",  # Western Transco Power
    5339: "Transmission Utility",  # Warora Kurnool Transmission
    5340: "Transmission Utility",  # Kallam Transmission
    5341: "Transmission Utility",  # Khavda II-A Transmission
    5342: "Transmission Utility",  # Sipat Transmission
    5343: "Transmission Utility",  # WRSS XXI(A) Transco
    5344: "Transmission Utility",  # Rajgarh Transmission
    5345: "Transmission Utility",  # KPS1 Transmission
    5346: "Transmission Utility",  # Chhattisgarh-WR Transmission
    5347: "Transmission Utility",  # Khavda-Bhuj Transmission
    5348: "Transmission Utility",  # Lakadia Banaskantha Transco
    5349: "Transmission Utility",  # Western Transmission Gujarat
    5350: "Transmission Utility",  # Raipur-Rajnandgaon-Warora Transmission
    5351: "Transmission Utility",  # Jam Khambaliya Transco

    # ── Distribution Company ────────────────────────────────────────────────
    5204: "Distribution Company",  # Maharashtra State Electricity Distribution Co.
    5355: "Distribution Company",  # Tata Power Mumbai (distribution licence area)

    # ── Central Utility ──────────────────────────────────────────────────────
    5333: "Central Utility",  # Central Operations Control Room ENOC
    5357: "Central Utility",  # NTPC Vidyut Vyapar Nigam Limited (PSU trading arm)

    # ── Generator (thermal / hydro / nuclear) ───────────────────────────────
    5162: "Generator",  # NTPC Mumbai HQ (WR-1)
    5163: "Generator",  # NTPC Raipur HQ (WR-2)
    5176: "Generator", 5177: "Generator", 5178: "Generator", 5179: "Generator",
    5180: "Generator", 5181: "Generator", 5182: "Generator", 5183: "Generator",
    5184: "Generator", 5185: "Generator", 5186: "Generator", 5187: "Generator",
    5188: "Generator", 5189: "Generator", 5190: "Generator", 5191: "Generator",
    5192: "Generator", 5193: "Generator", 5194: "Generator", 5195: "Generator",
    5196: "Generator", 5197: "Generator", 5198: "Generator", 5199: "Generator",
    5200: "Generator",  # hydro / thermal stations (Indira Sagar ... Hasdeo Bango)
    5206: "Generator",  # MP Power Generation Company
    5207: "Generator", 5208: "Generator", 5209: "Generator", 5210: "Generator",
    5211: "Generator", 5212: "Generator", 5213: "Generator", 5214: "Generator",
    5215: "Generator", 5216: "Generator",  # NTPC generating stations
    5217: "Generator", 5218: "Generator", 5219: "Generator", 5220: "Generator",
    # Tarapur / Kakrapar Atomic Power Stations
    5232: "Generator",  # Gujarat Industries Power Co. (GIPCL)
    5240: "Generator",  # Gujarat State Electricity Corp.
    5255: "Generator",  # Gujarat State Electricity Corp. - Phase II
    5263: "Generator",  # Torrent Power Limited (generating station connection)
    5296: "Generator",  # ACB (India) Limited
    5297: "Generator", 5298: "Generator", 5299: "Generator",  # Chakabura/SCPL/MCCPL thermal
    5300: "Generator",  # Adani Power (Mundra)
    5301: "Generator",  # Adani Power Tiroda
    5302: "Generator",  # Bharat Aluminium Co. (BALCO captive power)
    5303: "Generator",  # Tata Power Mundra (MTPS)
    5304: "Generator",  # DB Power
    5305: "Generator",  # DGEN Mega Power Project
    5306: "Generator",  # Dhariwal Infrastructure
    5307: "Generator",  # Dongamouha CPP
    5308: "Generator",  # Mahan Energen (Essar Power MP)
    5309: "Generator",  # GMR Warora Energy
    5310: "Generator",  # Jaypee Nigrie STPP
    5311: "Generator",  # Jhabua Power
    5312: "Generator",  # Jindal Power
    5313: "Generator",  # KSK Mahanadi (JSW Energy)
    5314: "Generator",  # LANCO Amarkantak
    5315: "Generator",  # MB Power (MP)
    5316: "Generator",  # NSPCL (NTPC-SAIL)
    5317: "Generator",  # Raigarh Energy Generation
    5318: "Generator",  # Adani Power - Raipur
    5319: "Generator",  # Ratnagiri Gas & Power
    5320: "Generator",  # RKM Powergen
    5321: "Generator",  # SASAN Power
    5322: "Generator",  # SKS Power Generation (CG)
    5323: "Generator",  # TRN Energy
    5324: "Generator",  # Adani Power Limited

    # ── RE Generator ─────────────────────────────────────────────────────────
    5221: "RE Generator", 5222: "RE Generator", 5223: "RE Generator",
    5224: "RE Generator", 5225: "RE Generator", 5226: "RE Generator",
    5227: "RE Generator", 5228: "RE Generator", 5229: "RE Generator",
    5230: "RE Generator", 5231: "RE Generator", 5233: "RE Generator",
    5234: "RE Generator", 5235: "RE Generator", 5236: "RE Generator",
    5237: "RE Generator", 5238: "RE Generator", 5239: "RE Generator",
    5241: "RE Generator", 5242: "RE Generator", 5243: "RE Generator",
    5244: "RE Generator", 5245: "RE Generator", 5246: "RE Generator",
    5247: "RE Generator", 5248: "RE Generator", 5249: "RE Generator",
    5250: "RE Generator", 5251: "RE Generator", 5252: "RE Generator",
    5253: "RE Generator", 5254: "RE Generator", 5256: "RE Generator",
    5257: "RE Generator", 5258: "RE Generator", 5259: "RE Generator",
    5260: "RE Generator", 5261: "RE Generator", 5262: "RE Generator",
    5264: "RE Generator", 5265: "RE Generator", 5266: "RE Generator",
    5267: "RE Generator", 5268: "RE Generator", 5269: "RE Generator",
    5270: "RE Generator", 5271: "RE Generator", 5272: "RE Generator",
    5273: "RE Generator", 5274: "RE Generator", 5275: "RE Generator",
    5276: "RE Generator", 5277: "RE Generator",
    5281: "RE Generator", 5282: "RE Generator", 5283: "RE Generator",
    5284: "RE Generator", 5285: "RE Generator", 5286: "RE Generator",
    5287: "RE Generator", 5288: "RE Generator", 5289: "RE Generator",
    5295: "RE Generator",  # Adani Renewable Energy Holding Five

    # ── Private Utility (industrial consumers / non-power private entities) ─
    5278: "Private Utility",  # Ambuja Cements - PSS-4
    5279: "Private Utility",  # Ambuja Cements - PSS-3
    5280: "Private Utility",  # Adani Hazira Port - PSS3
    5290: "Private Utility",  # Adani Ports & SEZ - PSS4
    5291: "Private Utility",  # Adani Ports & SEZ - PSS3
    5293: "Private Utility",  # M/s Manikaran Analytics Limited
    5352: "Private Utility",  # ArcelorMittal Nippon Steel India Limited

    # ── Other (section-header / grouping rows and genuinely ambiguous roles) ─
    5152: "Other",  # "NLDC and RLDCs" -- section header (parent_id grouping)
    5154: "Other",  # "Black Start Facilitated Stations" -- section header
    5164: "Other",  # "Nuclear Power Stations in Western Region" -- section header
    5165: "Other",  # "ISTS Connected RE Generators in Western Region" -- section header
    5166: "Other",  # "Qualified Coordinating Agency" -- role description, not an org
    5167: "Other",  # "Other Generating Stations in Western Region" -- section header
    5168: "Other",  # "Other Transmission Licensees in Western Region" -- section header
    5169: "Other",  # "Other Buyers in Western Region" -- section header
    5170: "Other",  # "Other Users in Western Region" -- section header
    5292: "Other",  # QCA for Bhuj-I ISTS Pooling station
    5294: "Other",  # QCA for REGSs of PSS1 & PSS2 at KPS-1 ISTS Pooling station
}


def run(apply: bool) -> None:
    print("DRY RUN\n" if not apply else "APPLY\n")
    with app.app_context():
        categories_by_name = {
            c.category_name: c.id for c in OrganizationCategory.query.all()
        }
        other_id = categories_by_name["Other"]

        changed = 0
        defaulted = 0
        for org in Organization.query.order_by(Organization.id).all():
            category_name = ORG_CATEGORIES.get(org.id)
            target_id = categories_by_name[category_name] if category_name else other_id

            if org.category_id == target_id:
                continue

            label = category_name or "Other (unmapped, review later)"
            print(f"  [{org.id}] {org.organization_name[:60]:<60} -> {label}")
            if apply:
                org.category_id = target_id
            changed += 1
            if not category_name:
                defaulted += 1

        if apply:
            db.session.commit()
            print(f"\nDone -- {changed} organizations categorized "
                  f"({defaulted} defaulted to 'Other' for manual review).")
        else:
            print(f"\n[Dry run -- {changed} orgs would be updated "
                  f"({defaulted} defaulting to 'Other'). Rerun with --apply]")


if __name__ == "__main__":
    run(apply="--apply" in sys.argv)
