-- v2 Migration: Run this once against the telephone_directory database
-- NOTE: MySQL 8.0 does NOT support ADD COLUMN (that is MariaDB syntax).
-- If a column already exists, MySQL will raise error 1060 (duplicate column) — just skip it.

USE telephone_directory;

-- 1. Add utility head flag to employees
ALTER TABLE employees
    ADD COLUMN is_utility_head BOOLEAN NOT NULL DEFAULT FALSE;

-- 2. Emergency contacts table (family/personal contacts of Grid India employees)
CREATE TABLE IF NOT EXISTS emergency_contacts (
    id          INT AUTO_INCREMENT PRIMARY KEY,
    employee_id INT NOT NULL,
    contact_name VARCHAR(255) NOT NULL,
    relation    VARCHAR(100),
    phone       VARCHAR(50),
    FOREIGN KEY (employee_id)
        REFERENCES employees(id)
        ON DELETE CASCADE
);

-- 3. Add new generic fields to update_requests (old specific columns are kept)
ALTER TABLE update_requests
    ADD COLUMN requested_by   VARCHAR(255),
    ADD COLUMN department     VARCHAR(255),
    ADD COLUMN contact_number VARCHAR(50),
    ADD COLUMN field_name     VARCHAR(100),
    ADD COLUMN new_value      TEXT;

-- 4. Make user_id nullable so anonymous submissions are allowed
ALTER TABLE update_requests
    MODIFY COLUMN user_id INT NULL;

-- 5. Add email and organization columns to directory_numbers if missing
--    (these were added in the earlier migration; safe to re-run)
ALTER TABLE directory_numbers
    ADD COLUMN email        TEXT,
    ADD COLUMN organization VARCHAR(255);
