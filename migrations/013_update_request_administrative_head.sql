-- Migration: adds administrative_head_id to update_requests, so the public
-- "Request Update" flow (already available for Employee and DirectoryNumber
-- records) can also be used for Administrative Head records.
--
-- Safe to run multiple times.

ALTER TABLE update_requests
    ADD COLUMN IF NOT EXISTS administrative_head_id INTEGER
        REFERENCES administrative_heads(id);
