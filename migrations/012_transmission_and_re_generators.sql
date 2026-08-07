-- Migration: merges CTU + STU into a single "Transmission Utility" type,
-- and splits RE Generators out of IPP as its own type. Both changes are
-- additive here -- migrating existing organizations off CTU/STU/IPP onto
-- the new rows, and retiring the now-unused CTU/STU category rows, is
-- scripts/reclassify_organization_types.py's job (same seed-migration /
-- assignment-script split used throughout this taxonomy's evolution).
--
-- Safe to run multiple times.

INSERT INTO organization_categories (category_name, description) VALUES
    ('Transmission Utility', NULL),
    ('RE Generators', NULL)
ON CONFLICT (category_name) DO NOTHING;
