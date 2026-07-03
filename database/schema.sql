CREATE DATABASE IF NOT EXISTS telephone_directory;

USE telephone_directory;


CREATE TABLE IF NOT EXISTS organizations(
    id INT AUTO_INCREMENT PRIMARY KEY,
    organization_name VARCHAR(500) UNIQUE,
    region VARCHAR(255),
    address TEXT
);


CREATE TABLE IF NOT EXISTS departments(
    id INT AUTO_INCREMENT PRIMARY KEY,
    department_name VARCHAR(100) UNIQUE
);


CREATE TABLE IF NOT EXISTS users(
    id INT AUTO_INCREMENT PRIMARY KEY,

    username VARCHAR(100) NOT NULL,

    email VARCHAR(255) UNIQUE NOT NULL,

    password VARCHAR(255) NOT NULL,

    role ENUM('admin','user')
    DEFAULT 'user',

    status ENUM('active','inactive')
    DEFAULT 'active',

    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);


CREATE TABLE IF NOT EXISTS employees(
    id INT AUTO_INCREMENT PRIMARY KEY,

    employee_name VARCHAR(255) NOT NULL,

    designation VARCHAR(255),

    organization_id INT,

    department_id INT,

    location VARCHAR(255),

    region VARCHAR(255),

    office_phone TEXT,

    residence_phone TEXT,

    mobile_phone TEXT,

    email VARCHAR(255),

    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

    FOREIGN KEY (organization_id)
    REFERENCES organizations(id)
    ON DELETE SET NULL,

    FOREIGN KEY (department_id)
    REFERENCES departments(id)
    ON DELETE SET NULL
);


CREATE TABLE IF NOT EXISTS directory_numbers(
    id INT AUTO_INCREMENT PRIMARY KEY,

    name VARCHAR(255) NOT NULL,

    phone_number TEXT,

    category VARCHAR(100)
);


CREATE TABLE IF NOT EXISTS update_requests(
    id INT AUTO_INCREMENT PRIMARY KEY,

    user_id INT NOT NULL,

    employee_id INT,

    directory_number_id INT,

    requested_mobile VARCHAR(50),

    requested_office_phone VARCHAR(100),

    requested_email VARCHAR(255),

    requested_directory_number VARCHAR(100),

    reason TEXT,

    request_type VARCHAR(20) DEFAULT 'employee',

    status ENUM(
        'Pending',
        'Approved',
        'Rejected'
    ) DEFAULT 'Pending',

    request_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

    FOREIGN KEY (user_id)
    REFERENCES users(id)
    ON DELETE CASCADE,

    FOREIGN KEY (employee_id)
    REFERENCES employees(id)
    ON DELETE CASCADE,

    FOREIGN KEY (directory_number_id)
    REFERENCES directory_numbers(id)
    ON DELETE CASCADE
);
