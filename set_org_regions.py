"""
Sets Organization.region to the state (Gujarat / Maharashtra / Madhya Pradesh /
Chhattisgarh / Goa / etc.) for every identifiable station.

This enables state-based searching in both the Employee Directory
(which already searches Organization.region) and, after a code change to
directory_number_search_filter, the Telephone Directory too.

Usage:
  python set_org_regions.py          # dry run
  python set_org_regions.py --apply  # commit
"""

import sys
from app import app
from models import db
from models.organization import Organization

# org_id -> state
ORG_STATES = {

    # ── WRLDC / WRPC (Maharashtra - Mumbai) ───────────────────────────────
    11:  "Maharashtra",   # WRLDC
    17:  "Maharashtra",   # Western Regional Power Committee

    # ── POWERGRID ──────────────────────────────────────────────────────────
    10:  "Maharashtra",   # POWERGRID WRTS-I, Nagpur

    # ── SLDCs ──────────────────────────────────────────────────────────────
    48:  "Maharashtra",   # Maharashtra SLDC
    49:  "Maharashtra",   # State Load Despatch Centre, MSETCL
    50:  "Maharashtra",   # Kalwa, Thane-Belapur
    51:  "Maharashtra",   # Ambazari, Nagpur
    52:  "Gujarat",       # Gujarat SLDC
    53:  "Madhya Pradesh",  # MP SLDC
    56:  "Chhattisgarh",  # Chhattisgarh SLDC
    57:  "Chhattisgarh",  # Backup SLDC CSPTCL Bhilai
    58:  "Goa",           # Goa Electricity Department
    59:  "Daman and Diu",
    60:  "Dadra and Nagar Haveli",
    228: "Maharashtra",   # MSEDC

    # ── NTPC WR-1 (Mumbai HQ) ─────────────────────────────────────────────
    61:  "Maharashtra",   # NTPC Mumbai HQ (WR-1) - office in Mumbai
    62:  "Gujarat",       # NTPC Gandhar (Bharuch, Gujarat)
    63:  "Gujarat",       # NTPC Kawas (Surat, Gujarat)
    64:  "Maharashtra",   # NTPC Mauda (Nagpur, Maharashtra)
    65:  "Maharashtra",   # NTPC Solapur (Solapur, Maharashtra)

    # ── NTPC WR-2 (Raipur HQ) ─────────────────────────────────────────────
    66:  "Chhattisgarh",  # NTPC Raipur HQ (Nava Raipur, CG)
    67:  "Madhya Pradesh",# NTPC Gadarwara (Narsinghpur, MP)
    68:  "Madhya Pradesh",# NTPC Khargone (Khandwa, MP)
    69:  "Chhattisgarh",  # NTPC Korba (Korba, CG)
    70:  "Chhattisgarh",  # NTPC Lara (Raigarh, CG)
    71:  "Chhattisgarh",  # NTPC Sipat (Bilaspur, CG)
    72:  "Madhya Pradesh",# NTPC Vindhyachal (Waidhan, MP)

    # ── Nuclear ────────────────────────────────────────────────────────────
    73:  "Maharashtra",   # Tarapur 1&2 (Palghar, Maharashtra)
    74:  "Maharashtra",   # Tarapur 3&4 (Palghar, Maharashtra)
    75:  "Gujarat",       # Kakrapar 1&2 (Surat, Gujarat)
    76:  "Gujarat",       # Kakrapar 3&4 (Surat, Gujarat)
    208: "Maharashtra",   # BARC Facility (Tarapur)
    209: "Maharashtra",   # Heavy Water Board (Mumbai)

    # ── RE Generators ─────────────────────────────────────────────────────
    77:  "Madhya Pradesh",# Athena Jaipur Solar (Rewa, MP)
    78:  "Gujarat",       # Adani Wind Energy Kutch One (Kutch, Gujarat)
    79:  "Madhya Pradesh",# Arinsun Clean Energy (Rewa, MP)
    81:  "Gujarat",       # INOX Green Energy (Dayapar, Kutch, Gujarat)
    83:  "Gujarat",       # Ostro Kutch Wind (Kutch, Gujarat)
    84:  "Rajasthan",     # ReNew Power (Rajasthan projects)
    86:  "Andhra Pradesh",# ReNew Wind Energy (AP2)
    87:  "Maharashtra",   # Continuum Power Trading (Mumbai)
    88:  "Gujarat",       # Gujarat Industries Power (Vadodara, Gujarat)
    89:  "Rajasthan",     # Electro Solaire (Rajasthan solar)
    90:  "Maharashtra",   # Tata Power Renewable (Mumbai HQ)
    91:  "Gujarat",       # NTPC Kawas Solar PV (Kawas, Gujarat)
    92:  "Gujarat",       # NTPC Jhanor-Gandhar Solar (Bharuch, Gujarat)
    93:  "Rajasthan",     # Ostro Energy (Rajasthan)
    94:  "Gujarat",       # Adani Wind Energy Kutch Three (Kutch, Gujarat)
    95:  "Gujarat",       # Adani Wind Energy Kutch Five (Kutch, Gujarat)
    96:  "Gujarat",       # GSECL (Vadodara, Gujarat)
    97:  "Rajasthan",     # Avikaran Solar (Delhi corp, Rajasthan project)
    98:  "Maharashtra",   # Powerica (Mumbai/Trombay)
    99:  "Rajasthan",     # Sitac Kabini Renewables
    100: "Madhya Pradesh",# Srijan Energy (Ratlam, MP)
    101: "Maharashtra",   # Solapur Solar PV (Solapur, Maharashtra)
    102: "Madhya Pradesh",# Adani Wind Energy MP One (Ratlam, MP)
    103: "Gujarat",       # Adani Wind Energy Kutch Four (Kutch, Gujarat)
    104: "Gujarat",       # Apraava Energy (Ahmedabad, Gujarat)
    105: "Rajasthan",     # Alfanar Netra Wind (Rajasthan)
    107: "Gujarat",       # Wind One Renergy (Vadodara, Gujarat)
    108: "Gujarat",       # Wind Two Renergy (Ahmedabad, Gujarat)
    109: "Gujarat",       # Wind Three Renergy (Vadodara, Gujarat)
    110: "Gujarat",       # Wind Five Renergy (Vadodara, Gujarat)
    111: "Gujarat",       # GSECL Phase II (Vadodara, Gujarat)
    113: "Madhya Pradesh",# NTPC RE Shajapur Unit-1 (Shajapur, MP)
    114: "Madhya Pradesh",# NTPC RE Shajapur Unit-2 (Shajapur, MP)
    115: "Gujarat",       # NTPC RE Dayapar WEP (Dayapar, Kutch, Gujarat)
    116: "Rajasthan",     # TP Saurya Unit-1 (Rajasthan solar)
    117: "Rajasthan",     # TP Saurya Unit-2 (Rajasthan solar)
    119: "Gujarat",       # Torrent Power (Ahmedabad, Gujarat)
    160: "Gujarat",       # D GEN Mega Power (Dahej, Bharuch, Gujarat)

    # Adani PSS entries (Mundra / Khavda, Gujarat)
    126: "Gujarat", 127: "Gujarat", 128: "Gujarat", 129: "Gujarat",
    130: "Gujarat", 131: "Gujarat", 132: "Gujarat", 133: "Gujarat",
    134: "Gujarat", 135: "Gujarat", 136: "Gujarat", 137: "Gujarat",
    138: "Gujarat", 139: "Gujarat", 140: "Gujarat", 141: "Gujarat",
    142: "Gujarat", 143: "Rajasthan", 144: "Rajasthan",
    145: "Gujarat", 146: "Gujarat", 147: "Gujarat",
    149: "Gujarat",  # ADANI RE HOLDING FIVE (Bhuj PSS)

    # ── Other Generating Stations ──────────────────────────────────────────
    150: "Chhattisgarh",  # ACB (India) - Korba, CG
    154: "Gujarat",       # Adani Power Mundra (Mundra, Kutch, Gujarat)
    155: "Maharashtra",   # ADANI POWER TIRODA (Tiroda, Gondia, Maharashtra)
    156: "Chhattisgarh",  # BALCO (Korba, CG)
    157: "Gujarat",       # Tata Power Mundra (Mundra, Kutch, Gujarat)
    158: "Chhattisgarh",  # DB Power (Janjgir-Champa, CG)
    160: "Gujarat",       # D GEN Mega Power (Dahej, Gujarat)
    161: "Maharashtra",   # Dhariwal Infrastructure (Chandrapur, MH)
    162: "Chhattisgarh",  # Dongamouha CPP (Raigarh, CG)
    163: "Madhya Pradesh",# Mahan Energen - Essar Power MP (Singrauli, MP)
    164: "Maharashtra",   # GMR Warora Energy (Warora, Chandrapur, MH)
    165: "Madhya Pradesh",# Jhabua Power (Seoni, MP)
    166: "Chhattisgarh",  # Jindal Power (Tamnar, Raigarh, CG)
    167: "Chhattisgarh",  # KSK Mahanadi (Chhattisgarh - plant, HQ Hyderabad)
    168: "Chhattisgarh",  # LANCO Amarkantak (Korba, CG)
    169: "Madhya Pradesh",# MB Power (Anupur, MP)
    170: "Chhattisgarh",  # NTPC-SAIL NSPCL (Bhilai, CG)
    171: "Chhattisgarh",  # Raigarh Energy Generation (Korba, CG)
    172: "Chhattisgarh",  # Adani Power Raipur (Raikheda, Raipur, CG)
    173: "Maharashtra",   # Ratnagiri Gas & Power (Ratnagiri, MH)
    175: "Madhya Pradesh",# SASAN Power (Singrauli, MP)
    177: "Chhattisgarh",  # SKS Power Generation CG
    178: "Madhya Pradesh",# SSP (Indore, MP from address)
    179: "Chhattisgarh",  # TRN Energy (Raigarh, CG)
    181: "Gujarat",       # Adani Power Ltd (Mundra, Gujarat HQ)
    182: "Gujarat",       # AMNS Power Transmission (Hazira, Surat, Gujarat)
    207: "Gujarat",       # ArcelorMittal Nippon Steel (Hazira, Surat, Gujarat)

    # ── Transmission ──────────────────────────────────────────────────────
    191: "Maharashtra",   # Mumbai Urja Marg Limited
    192: "Goa",           # Goa Tamnar Transmission
    204: "Gujarat",       # Western Transmission Gujarat
    205: "Maharashtra",   # Raipur-Rajnandgaon-Warora (Mumbai HQ)
    211: "Madhya Pradesh",# National High Power Test Lab (Bina, Sagar, MP)
    222: "Maharashtra",   # SIEMENS REMC (Thane, MH)

    # ── Misc ──────────────────────────────────────────────────────────────
    54:  "Madhya Pradesh",  # MP POWER MANAGEMENT CO.
    55:  "Madhya Pradesh",  # MP Power Generation Company
}


def run(apply: bool) -> None:
    print("DRY RUN\n" if not apply else "APPLY\n")
    with app.app_context():
        changed = 0
        for org_id, state in sorted(ORG_STATES.items()):
            org = db.session.get(Organization, org_id)
            if not org:
                print(f"  SKIP id={org_id} -- not found")
                continue
            if org.region == state:
                continue   # already set correctly
            print(f"  [{org_id}] {org.organization_name[:55]:<55} region -> {state}")
            if apply:
                org.region = state
            changed += 1

        if apply:
            db.session.commit()
            print(f"\nDone -- {changed} organizations updated.")
        else:
            print(f"\n[Dry run -- {changed} orgs would be updated. Rerun with --apply]")


if __name__ == "__main__":
    run(apply="--apply" in sys.argv)
