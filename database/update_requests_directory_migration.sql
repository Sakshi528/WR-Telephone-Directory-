USE telephone_directory;

ALTER TABLE update_requests
    MODIFY employee_id INT NULL;

ALTER TABLE update_requests
    ADD COLUMN directory_number_id INT NULL;

ALTER TABLE update_requests
    ADD COLUMN requested_directory_number VARCHAR(100) NULL;

ALTER TABLE update_requests
    ADD COLUMN request_type VARCHAR(20) DEFAULT 'employee';

ALTER TABLE update_requests
    ADD CONSTRAINT fk_update_requests_directory_number
    FOREIGN KEY (directory_number_id)
    REFERENCES directory_numbers(id)
    ON DELETE CASCADE;
