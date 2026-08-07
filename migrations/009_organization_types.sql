-- Migration: Organization Type taxonomy (cascading Organization Type ->
-- Organization dropdowns). Reuses the existing organization_categories
-- table / Organization.category_id column -- "Organization Type" is a UI
-- label for the same concept, not a new schema field.
--
-- Additive only: inserts the 7 new category rows. Does NOT delete the old
-- 10 categories or reassign any organization -- see
-- scripts/reclassify_organization_types.py for that (mirrors the existing
-- seed-migration / assignment-script split used for the original taxonomy:
-- 003_email_distribution_lists.sql + scripts/classify_organizations.py).
--
-- Safe to run multiple times.

INSERT INTO organization_categories (category_name, description) VALUES
    ('CTU', NULL),
    ('STU', NULL),
    ('State SLDC', NULL),
    ('IPP', NULL),
    ('CPSU', NULL),
    ('DISCOM', NULL),
    ('Others', NULL)
ON CONFLICT (category_name) DO NOTHING;
