"""Merge the top-level Thermal and Hydel categories into Generation Company
as Subcategory values (2026-09-22).

Background: the source Word directory (uploads/Western Region Phone
Directory 2026 Main_TD.docx) has no "Thermal Power Stations" or "Hydel
Power Stations" heading of its own -- confirmed by scanning every Heading 1
paragraph in the document. Every thermal and hydel plant sits inside the
single generic "Other Generating Stations in Western Region" section (or
under the cross-cutting "Black Start Facilitated Stations" flag), and was
only later split into separate Thermal/Hydel/Nuclear top-level app
categories by migration 010. "Nuclear Power Stations in Western Region" IS
its own Heading 1 in the Word document, so Nuclear is left as its own
top-level category -- only Thermal and Hydel are merged here, per the GM's
feedback that Thermal/Hydel should be a Subcategory of Generation Company,
not a Category of their own.

What this does:
  1. Every organization currently in category Thermal or Hydel moves to
     category Generation Company, with its Subcategory (Organization.region)
     set to "Thermal" or "Hydel" respectively.
     Safe to overwrite `region` this way: for every affected organization,
     the previous `region` value (e.g. "Black Start Facilitated Stations",
     "NTPC Raipur Head Quarter (WR-2)", "Other Generating Stations In
     Western Region") is ALSO stored independently via `parent_id` (the
     real tree relationship the Telephone Directory page renders from) --
     confirmed by inspecting all 65 affected rows before running this
     script. No grouping information is lost.
  2. The now-empty Thermal / Hydel rows are deleted from
     organization_categories (0 email_group_filters referenced either,
     confirmed before deleting).
  3. Runs the app's own subcategory sync logic (same as the admin
     "Sync from live data" button, routes/admin_routes.py sync_subcategories)
     to add the new Thermal/Hydel suggestions under Generation Company and
     drop now-stale suggestions left over from the old categories.

Usage:
  python scripts/merge_thermal_hydel_into_generation_company.py           (dry run)
  python scripts/merge_thermal_hydel_into_generation_company.py --apply   (writes to DB)
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from models import db
from models.organization import Organization
from models.organization_category import OrganizationCategory
from models.organization_subcategory import OrganizationSubcategory
from sqlalchemy import func


def main():
    apply = "--apply" in sys.argv

    with app.app_context():
        gen_co = OrganizationCategory.query.filter_by(category_name="Generation Company").first()
        thermal = OrganizationCategory.query.filter_by(category_name="Thermal").first()
        hydel = OrganizationCategory.query.filter_by(category_name="Hydel").first()
        if not (gen_co and thermal and hydel):
            print("!! Expected categories not found -- aborting"); return

        thermal_orgs = Organization.query.filter_by(category_id=thermal.id).all()
        hydel_orgs = Organization.query.filter_by(category_id=hydel.id).all()

        print(f"{'APPLYING' if apply else 'DRY RUN -- would move'}: "
              f"{len(thermal_orgs)} Thermal + {len(hydel_orgs)} Hydel org(s) "
              f"-> Generation Company (subcategory 'Thermal' / 'Hydel')\n")

        for org in thermal_orgs:
            print(f"  [{org.id}] {org.organization_name}: region {org.region!r} -> 'Thermal'")
            if apply:
                org.category_id = gen_co.id
                org.region = "Thermal"
        for org in hydel_orgs:
            print(f"  [{org.id}] {org.organization_name}: region {org.region!r} -> 'Hydel'")
            if apply:
                org.category_id = gen_co.id
                org.region = "Hydel"

        # Safety re-check right before deleting the categories
        stray_thermal = Organization.query.filter_by(category_id=thermal.id).count()
        stray_hydel = Organization.query.filter_by(category_id=hydel.id).count()
        from models.email_group_filter import EmailGroupFilter
        efilters = EmailGroupFilter.query.filter(
            EmailGroupFilter.category_id.in_([thermal.id, hydel.id])
        ).count()

        if apply:
            db.session.flush()
            stray_thermal = Organization.query.filter_by(category_id=thermal.id).count()
            stray_hydel = Organization.query.filter_by(category_id=hydel.id).count()

        if stray_thermal or stray_hydel or efilters:
            print(f"\n!! Not deleting categories -- still referenced "
                  f"(orgs: {stray_thermal + stray_hydel}, email filters: {efilters})")
            if apply:
                db.session.commit()
            return

        print(f"\n{'Deleting' if apply else 'Would delete'} now-empty categories: Thermal, Hydel")
        if apply:
            # Raw DELETE, not ORM db.session.delete(): the ORM's default cascade
            # tries to null out organization_subcategories.category_id first,
            # which violates that column's NOT NULL constraint. The DB's own
            # ON DELETE CASCADE (migration 018) handles any leftover
            # subcategory suggestion rows correctly.
            db.session.execute(
                OrganizationCategory.__table__.delete().where(
                    OrganizationCategory.id.in_([thermal.id, hydel.id])
                )
            )
            db.session.commit()

            # Reuse the app's own subcategory sync logic (routes/admin_routes.py
            # sync_subcategories) so Thermal/Hydel become proper suggestions under
            # Generation Company and stale ones from the old categories are dropped.
            added = (
                db.session.query(Organization.category_id, Organization.region)
                .filter(Organization.category_id.isnot(None), Organization.region.isnot(None))
                .filter(func.btrim(Organization.region) != "")
                .distinct()
                .all()
            )
            added_count = 0
            for category_id, region in added:
                exists = OrganizationSubcategory.query.filter_by(
                    category_id=category_id, subcategory_name=region
                ).first()
                if not exists:
                    db.session.add(OrganizationSubcategory(category_id=category_id, subcategory_name=region))
                    added_count += 1
            removed_count = 0
            for sub in OrganizationSubcategory.query.all():
                still_used = Organization.query.filter_by(
                    category_id=sub.category_id, region=sub.subcategory_name
                ).first()
                if not still_used:
                    db.session.delete(sub)
                    removed_count += 1
            db.session.commit()
            print(f"Subcategory sync: added {added_count}, removed {removed_count} stale suggestion(s).")
            print("\nCommitted.")
        else:
            print("\nNo changes written -- re-run with --apply to commit.")


if __name__ == "__main__":
    main()
