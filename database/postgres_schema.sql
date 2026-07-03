-- PostgreSQL schema for Telephone Directory
-- Run once on a fresh database: psql -U postgres -d telephone_directory -f postgres_schema.sql

CREATE TABLE IF NOT EXISTS organizations (
    id                SERIAL PRIMARY KEY,
    organization_name VARCHAR(500) UNIQUE,
    region            VARCHAR(255),
    address           TEXT
);

CREATE TABLE IF NOT EXISTS departments (
    id              SERIAL PRIMARY KEY,
    department_name VARCHAR(100) UNIQUE NOT NULL
);

CREATE TABLE IF NOT EXISTS users (
    id         SERIAL PRIMARY KEY,
    username   VARCHAR(100) NOT NULL,
    email      VARCHAR(255) UNIQUE NOT NULL,
    password   VARCHAR(255) NOT NULL,
    role       VARCHAR(20)  DEFAULT 'user',
    status     VARCHAR(20)  DEFAULT 'active',
    created_at TIMESTAMP    DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS employees (
    id              SERIAL PRIMARY KEY,
    employee_name   VARCHAR(255) NOT NULL,
    designation     VARCHAR(255),
    organization_id INT REFERENCES organizations(id) ON DELETE SET NULL,
    department_id   INT REFERENCES departments(id)   ON DELETE SET NULL,
    location        VARCHAR(255),
    region          VARCHAR(255),
    office_phone    TEXT,
    residence_phone TEXT,
    mobile_phone    TEXT,
    email           VARCHAR(255),
    is_utility_head BOOLEAN NOT NULL DEFAULT FALSE,
    created_at      TIMESTAMP DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS directory_numbers (
    id           SERIAL PRIMARY KEY,
    name         VARCHAR(255) NOT NULL,
    phone_number TEXT,
    email        TEXT,
    organization VARCHAR(255),
    category     VARCHAR(100)
);

CREATE TABLE IF NOT EXISTS emergency_contacts (
    id           SERIAL PRIMARY KEY,
    employee_id  INT NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
    contact_name VARCHAR(255) NOT NULL,
    relation     VARCHAR(100),
    phone        VARCHAR(50)
);

CREATE TABLE IF NOT EXISTS update_requests (
    id                        SERIAL PRIMARY KEY,
    user_id                   INT REFERENCES users(id) ON DELETE CASCADE,
    employee_id               INT REFERENCES employees(id) ON DELETE CASCADE,
    directory_number_id       INT REFERENCES directory_numbers(id) ON DELETE CASCADE,
    -- Legacy specific fields (kept for backward compatibility)
    requested_mobile          VARCHAR(50),
    requested_office_phone    VARCHAR(100),
    requested_email           VARCHAR(255),
    requested_directory_number VARCHAR(100),
    -- New generic fields
    requested_by              VARCHAR(255),
    department                VARCHAR(255),
    contact_number            VARCHAR(50),
    field_name                VARCHAR(100),
    new_value                 TEXT,
    reason                    TEXT,
    request_type              VARCHAR(20) DEFAULT 'employee',
    status                    VARCHAR(20) DEFAULT 'Pending',
    request_date              TIMESTAMP   DEFAULT NOW()
);
