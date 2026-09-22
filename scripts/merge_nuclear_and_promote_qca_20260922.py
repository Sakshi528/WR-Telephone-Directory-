"""Second category-restructure pass (2026-09-22), per GM feedback relayed
through the admin: reduce the top-level Organisation Category list down to
CPSU, DISCOM, SLDC, Transmission Utility, QCA, Others, Generation Company,
Regional Power Committee, RLDC (RLDC and Regional Power Committee confirmed
kept -- they're grid-operator bodies, not utility/generator/buyer types, so
forcing them into any of the 7 named categories would misclassify them,
including WRLDC itself).

Two changes:

1. Nuclear -> merged into Generation Company as subcategory "Nuclear", same
   pattern already applied to Thermal/Hydel/RE Generators. 5 organizations
   (2 Kakrapar units, 2 Tarapur units, + the "Nuclear Power Stations In
   Western Region" parent/tree node). 0 email_group_filters reference the
   Nuclear category_id -- confirmed safe to delete after moving.

2. QCA promoted from a sub-grouping inside "Others" to its own top-level
   category. The existing "Qualified Coordinating Agency" organization (a
   parent/tree node, 3 real children: the two QCA pooling-station entries
   and M/S Manikaran Analytics Limited) and its children move from Others
   to the new QCA category.

Usage:
  python scripts/merge_nuclear_and_promote_qca_20260922.py           (dry run)
  python scripts/merge_nuclear_and_promote_qca_20260922.py --apply   (writes to DB)
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


def sync_subcategories():
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


def main():
    apply = "--apply" in sys.argv

    with app.app_context():
        gen_co = OrganizationCategory.query.filter_by(category_name="Generation Company").first()
        nuclear = OrganizationCategory.query.filter_by(category_name="Nuclear").first()
        others = OrganizationCategory.query.filter_by(category_name="Others").first()
        qca_cat = OrganizationCategory.query.filter_by(category_name="QCA").first()

        if not (gen_co and nuclear and others):
            print("!! Expected categories not found -- aborting"); return

        # --- 1. Nuclear -> Generation Company / subcategory "Nuclear" ---
        nuclear_orgs = Organization.query.filter_by(category_id=nuclear.id).all()
        print(f"{'APPLYING' if apply else 'DRY RUN -- would move'}: "
              f"{len(nuclear_orgs)} Nuclear org(s) -> Generation Company (subcategory 'Nuclear')\n")
        for org in nuclear_orgs:
            print(f"  [{org.id}] {org.organization_name}: region {org.region!r} -> 'Nuclear'")
            if apply:
                org.category_id = gen_co.id
                org.region = "Nuclear"

        # --- 2. Promote QCA ---
        if not qca_cat:
            print(f"\n{'Creating' if apply else 'Would create'} category 'QCA'")
            if apply:
                qca_cat = OrganizationCategory(category_name="QCA",
                                                description="Qualified Coordinating Agency",
                                                is_state_based=False)
                db.session.add(qca_cat)
                db.session.flush()

        qca_parent = Organization.query.filter_by(organization_name="Qualified Coordinating Agency").first()
        if qca_parent:
            qca_children = Organization.query.filter_by(parent_id=qca_parent.id).all()
            print(f"\n{'Moving' if apply else 'Would move'} QCA parent + "
                  f"{len(qca_children)} child org(s) from Others -> QCA:")
            for org in [qca_parent] + qca_children:
                print(f"  [{org.id}] {org.organization_name}")
                if apply:
                    org.category_id = qca_cat.id

        if apply:
            db.session.commit()

            # Delete now-empty Nuclear category (raw DELETE -- see the Thermal/Hydel
            # merge script for why ORM db.session.delete() fails here)
            stray = Organization.query.filter_by(category_id=nuclear.id).count()
            from models.email_group_filter import EmailGroupFilter
            efilters = EmailGroupFilter.query.filter_by(category_id=nuclear.id).count()
            if stray == 0 and efilters == 0:
                db.session.execute(
                    OrganizationCategory.__table__.delete().where(OrganizationCategory.id == nuclear.id)
                )
                db.session.commit()
                print("\nDeleted now-empty category: Nuclear")
            else:
                print(f"\n!! Not deleting Nuclear -- still referenced (orgs: {stray}, filters: {efilters})")

            sync_subcategories()
            print("\nCommitted.")
        else:
            print("\nNo changes written -- re-run with --apply to commit.")


if __name__ == "__main__":
    main()
