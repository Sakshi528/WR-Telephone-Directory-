-- Migration: adds switch_yard, control_room, and ip_address to
-- directory_numbers, so a Control Room/Switchyard entry can be explicitly
-- flagged and carry a network IP (in addition to the existing free-text
-- `category` column, which continues to drive Control Room/Switchyard
-- grouping everywhere else in the app -- these are additive fields only).
--
-- Safe to run multiple times.

ALTER TABLE directory_numbers
    ADD COLUMN IF NOT EXISTS switch_yard BOOLEAN NOT NULL DEFAULT FALSE,
    ADD COLUMN IF NOT EXISTS control_room BOOLEAN NOT NULL DEFAULT FALSE,
    ADD COLUMN IF NOT EXISTS ip_address VARCHAR(45);
