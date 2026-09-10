-- Migration: adds organizations.utility_head_excluded, letting an admin
-- explicitly say "this organization has no Utility Head" instead of the
-- app always auto-resolving one from the most senior designation
-- (utils.designation_rank.resolve_utility_head). Without this flag, an
-- organization with only one employee could never actually have its
-- Utility Head "removed" -- clearing Employee.is_utility_head just made
-- auto-resolution immediately pick that same employee again, since
-- resolve_utility_head always returns someone when the employee list is
-- non-empty.
--
-- Safe to run multiple times (IF NOT EXISTS guards the column add).

ALTER TABLE organizations
    ADD COLUMN IF NOT EXISTS utility_head_excluded BOOLEAN NOT NULL DEFAULT FALSE;
