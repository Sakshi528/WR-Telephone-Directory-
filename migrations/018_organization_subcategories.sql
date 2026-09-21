-- Migration: Organization Subcategories -- a curated, per-Category list of
-- reusable Subcategory values (Organization.region), so an admin can
-- register a new subcategory (for a future station/organization) through
-- the UI instead of needing a code change/migration, and the Add/Edit
-- Organization form can suggest previously-used values.
-- Safe to run multiple times.

CREATE TABLE IF NOT EXISTS organization_subcategories (
    id SERIAL PRIMARY KEY,
    category_id INTEGER NOT NULL REFERENCES organization_categories(id) ON DELETE CASCADE,
    subcategory_name VARCHAR(255) NOT NULL
);

DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM pg_constraint WHERE conname = 'uq_subcategory_per_category'
    ) THEN
        ALTER TABLE organization_subcategories
            ADD CONSTRAINT uq_subcategory_per_category UNIQUE (category_id, subcategory_name);
    END IF;
END $$;

CREATE INDEX IF NOT EXISTS idx_organization_subcategories_category ON organization_subcategories(category_id);

-- Backfill: seed one subcategory row per category from the distinct
-- Organization.region values already in use, so existing data shows up as
-- suggestions immediately instead of an empty list.
INSERT INTO organization_subcategories (category_id, subcategory_name)
SELECT DISTINCT o.category_id, o.region
FROM organizations o
WHERE o.category_id IS NOT NULL
  AND o.region IS NOT NULL
  AND btrim(o.region) <> ''
ON CONFLICT (category_id, subcategory_name) DO NOTHING;
