"""Session fixes (2026-09-22) -- category/subcategory cleanup.

Follow-up to scripts/full_category_subcategory_bifurcation.py. That report's
automated Word-section matcher only cleanly resolves 98/216 organizations;
everything else needs per-organization judgment (many are legitimate
Word-doc-section "parent" organizations used as tree nodes in the Telephone
Directory -- e.g. "Black Start Facilitated Stations", "Other Generating
Stations In Western Region" -- and must NOT be deleted or recategorized,
since their children are already correctly split across Thermal/Hydel/
Generation Company/Nuclear).

This script applies only the fixes that were individually verified this
session (DB inspection + public-source lookup where noted):

1. M/S Adani Renewable Energy Holding Five Limited -- was parented under the
   "Qualified Coordinating Agency" node with subcategory "Qualified
   Coordinating Agency", but it is a Generation Company (RE) like every one
   of its ~72 siblings. Re-parented under "ISTS Connected RE Generators In
   Western Region" with matching subcategory text.

2. NTPC Mumbai Head Quarter (WR-1) / NTPC Raipur Head Quarter (WR-2) -- each
   had its own state name in the subcategory (region) field instead of a
   blank/self value, inconsistent with every plant that correctly points its
   subcategory at one of these two HQ names. Cleared to blank.

3. Madhya Pradesh Power Generation Company Limited -- subcategory field said
   "Madhya Pradesh State Load Despatch Centre" (it physically sits near the
   SLDC contact block in the Word doc, but it is a Genco, not an SLDC
   sub-unit). Cleared to blank.

4. Tata Power Mumbai -- was in the catch-all "Others" category. Verified via
   public source (tatapower.com/energy-solutions/distribution/mumbai) that
   Tata Power's Mumbai operation is a distribution licensee serving Colaba to
   Mira-Bhayander. Recategorized to DISCOM (state = Maharashtra).

5. "NTPC" (org id from a stray import row, subcategory "NTPC") -- has 0
   children, 0 employees, 0 directory numbers, 0 administrative heads, and no
   plain "NTPC" heading/line exists anywhere in the current source Word
   document (uploads/Western Region Phone Directory 2026 Main_TD.docx) --
   confirmed by directly scanning the document's paragraphs. This is a
   leftover, dependent-free duplicate from an earlier import generation and
   is deleted.

Explicitly NOT touched (verified correct or correctly ambiguous):
- The 7 legitimate Word-section "parent" organizations that hold real
  children (Black Start Facilitated Stations, Other Buyers In Western
  Region, Other Generating Stations In Western Region, Other Users In
  Western Region, Qualified Coordinating Agency, ISTS Connected RE
  Generators In Western Region, Nuclear Power Stations In Western Region).
  These ARE the Telephone Directory's tree nodes -- deleting them would
  orphan their children in the UI.
- BARC Facility, Heavy Water Board, Arcelormittal Nippon Steel India
  Limited, National High Power Test Laboratory, NTPC Vidyut Vyapar Nigam
  Limited, M/S Manikaran Analytics Limited, QCA For Bhuj-I ISTS Pooling
  Station, QCA For Regss Of PSS1 & PSS2 At 765/400 kV KPS-1 ISTS Pooling
  Station, MP Power Management CO. LTD. -- none of these are power
  utilities/generators; "Others" is the correct fit given the app's
  controlled 10-category taxonomy has no Buyer/Trading/Testing-Lab
  category. Left as-is deliberately, not an oversight.

Usage:
  python scripts/apply_session_fixes_20260922_category_cleanup.py           (dry run, prints only)
  python scripts/apply_session_fixes_20260922_category_cleanup.py --apply   (writes to DB)
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from models import db
from models.organization import Organization
from models.organization_category import OrganizationCategory


def main():
    apply = "--apply" in sys.argv

    with app.app_context():
        cat_id = {c.category_name: c.id for c in OrganizationCategory.query.all()}

        def get(name):
            org = Organization.query.filter_by(organization_name=name).first()
            if org is None:
                print(f"  !! NOT FOUND: {name!r} -- skipping")
            return org

        changes = []

        # 1. Adani RE Holding Five -- re-parent + fix subcategory
        adani = get("M/S Adani Renewable Energy Holding Five Limited")
        re_parent = get("ISTS Connected RE Generators In Western Region")
        if adani and re_parent:
            changes.append((
                adani,
                f"parent_id {adani.parent_id} -> {re_parent.id}; "
                f"region {adani.region!r} -> {re_parent.organization_name!r}",
                lambda o=adani, p=re_parent: (setattr(o, "parent_id", p.id),
                                               setattr(o, "region", p.organization_name)),
            ))

        # 2. NTPC HQ subcategory cleanup
        for hq_name in ("NTPC Mumbai Head Quarter (WR-1)", "NTPC Raipur Head Quarter (WR-2)"):
            hq = get(hq_name)
            if hq:
                changes.append((
                    hq, f"region {hq.region!r} -> ''",
                    lambda o=hq: setattr(o, "region", ""),
                ))

        # 3. MP Power Generation Co subcategory cleanup
        mppgcl = get("Madhya Pradesh Power Generation Company Limited")
        if mppgcl:
            changes.append((
                mppgcl, f"region {mppgcl.region!r} -> ''",
                lambda o=mppgcl: setattr(o, "region", ""),
            ))

        # 4. Tata Power Mumbai -> DISCOM
        tpm = get("Tata Power Mumbai")
        if tpm and "DISCOM" in cat_id:
            old_cat = tpm.category_id
            changes.append((
                tpm, f"category_id {old_cat} -> {cat_id['DISCOM']} (DISCOM); state -> Maharashtra",
                lambda o=tpm: (setattr(o, "category_id", cat_id["DISCOM"]),
                               setattr(o, "state", "Maharashtra")),
            ))

        # 5. Delete orphan "NTPC" duplicate (0 children, 0 contacts, confirmed no longer
        #    present as a heading/line in the current source Word document)
        ntpc_stray = Organization.query.filter_by(organization_name="NTPC").first()
        ntpc_delete = None
        if ntpc_stray:
            no_children = Organization.query.filter_by(parent_id=ntpc_stray.id).count() == 0
            from models.employee import Employee
            from models.directory_number import DirectoryNumber
            from models.administrative_head import AdministrativeHead
            no_emp = Employee.query.filter_by(organization_id=ntpc_stray.id).count() == 0
            no_dn = DirectoryNumber.query.filter_by(organization_id=ntpc_stray.id).count() == 0
            no_ah = AdministrativeHead.query.filter_by(organization_id=ntpc_stray.id).count() == 0
            if no_children and no_emp and no_dn and no_ah:
                ntpc_delete = ntpc_stray
            else:
                print("  !! 'NTPC' now has dependents -- NOT deleting, re-check manually")

        print(f"{'APPLYING' if apply else 'DRY RUN -- would apply'} {len(changes)} update(s)"
              f"{' + 1 delete' if ntpc_delete else ''}:\n")
        for org, desc, fn in changes:
            print(f"  [{org.id}] {org.organization_name}: {desc}")
            if apply:
                fn()
        if ntpc_delete:
            print(f"  [{ntpc_delete.id}] DELETE 'NTPC' (dependent-free duplicate)")
            if apply:
                db.session.delete(ntpc_delete)

        if apply:
            db.session.commit()
            print("\nCommitted.")
        else:
            print("\nNo changes written -- re-run with --apply to commit.")


if __name__ == "__main__":
    main()
