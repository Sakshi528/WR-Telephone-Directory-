-- Migration: State level + richer Organization Type taxonomy for the
-- cascading Type -> State -> Organization dropdowns.
--
-- Adds a real `state` column (Organization.region already holds a
-- different concept -- the parent import-grouping label -- confirmed by
-- inspection, so it can't be repurposed). Adds `is_state_based` on
-- organization_categories to drive whether the State dropdown appears for
-- a given Type. Adds the 3 category rows the earlier 7-value taxonomy
-- (migration 009) didn't cover -- RLDC already exists from the original
-- 10-category seed (003_email_distribution_lists.sql).
--
-- Safe to run multiple times.

ALTER TABLE organizations ADD COLUMN IF NOT EXISTS state VARCHAR(100);

ALTER TABLE organization_categories
    ADD COLUMN IF NOT EXISTS is_state_based BOOLEAN NOT NULL DEFAULT FALSE;

INSERT INTO organization_categories (category_name, description) VALUES
    ('Thermal', NULL),
    ('Hydel', NULL),
    ('Nuclear', NULL)
ON CONFLICT (category_name) DO NOTHING;

UPDATE organization_categories
SET is_state_based = TRUE
WHERE category_name IN ('State SLDC', 'STU', 'DISCOM');
