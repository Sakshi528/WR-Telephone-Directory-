-- Migration: removes the unused 'RE Generators' organization category.
-- All 74 RE Generator organizations were deliberately merged into
-- 'Generation Company' on 2026-08-27 (see reports/reclassification_proposal.csv),
-- leaving this category empty (0 organizations) ever since -- confirmed by
-- the admin re-reviewing the Categories page and asking for it to be
-- removed the same way the empty 'IPP' category was (migration 015).
--
-- Confirmed safe before writing this migration: 0 organizations.category_id
-- and 0 email_group_filters.category_id reference it. Its one
-- organization_subcategories row ("ISTS Connected RE Generators In Western
-- Region") cascade-deletes via that table's ON DELETE CASCADE FK
-- (migration 018) -- no separate cleanup needed.
--
-- Without this migration, deleting the live row alone would not be
-- durable against a fresh migration replay (009's seed insert uses
-- ON CONFLICT DO NOTHING, so it would silently come back).
--
-- Safe to run multiple times (no-op if already deleted).

DELETE FROM organization_categories WHERE category_name = 'RE Generators';
