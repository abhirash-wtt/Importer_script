-- DDO++ Attendance tables for the sidecar API (NocoBase 0.9.4-alpha.2 / PostgreSQL).
-- NocoBase collections can later map 1:1 onto these public.att_* tables.
-- The sidecar writes here; NocoBase reads the same database.

CREATE TABLE IF NOT EXISTS att_locations (
    id BIGSERIAL PRIMARY KEY,
    location_code VARCHAR(20) NOT NULL UNIQUE,
    location_name VARCHAR(100) NOT NULL UNIQUE,
    is_active BOOLEAN NOT NULL DEFAULT TRUE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

INSERT INTO att_locations (location_code, location_name)
VALUES
    ('AGRA', 'Agra'),
    ('NOIDA', 'Noida'),
    ('HYD', 'Hyderabad')
ON CONFLICT (location_code) DO NOTHING;

CREATE TABLE IF NOT EXISTS att_departments (
    id BIGSERIAL PRIMARY KEY,
    department_code VARCHAR(50) UNIQUE,
    department_name VARCHAR(150) NOT NULL UNIQUE,
    is_active BOOLEAN NOT NULL DEFAULT TRUE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS att_employees (
    id BIGSERIAL PRIMARY KEY,
    employee_code VARCHAR(50) NOT NULL,
    employee_name VARCHAR(200) NOT NULL DEFAULT '',
    department_id BIGINT REFERENCES att_departments(id),
    location_id BIGINT NOT NULL REFERENCES att_locations(id),
    biometric_code VARCHAR(100),
    identity_status VARCHAR(30) NOT NULL DEFAULT 'CODED',
    is_active BOOLEAN NOT NULL DEFAULT TRUE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    CONSTRAINT uq_att_employees_code_location UNIQUE (location_id, employee_code)
);

CREATE TABLE IF NOT EXISTS att_attendance_statuses (
    id BIGSERIAL PRIMARY KEY,
    code VARCHAR(30) NOT NULL UNIQUE,
    display_name VARCHAR(100) NOT NULL,
    category VARCHAR(30) NOT NULL,
    is_half_day BOOLEAN NOT NULL DEFAULT FALSE,
    hr_definition_required BOOLEAN NOT NULL DEFAULT FALSE,
    is_active BOOLEAN NOT NULL DEFAULT TRUE,
    description TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS att_import_batches (
    id BIGSERIAL PRIMARY KEY,
    batch_uid VARCHAR(64) NOT NULL UNIQUE,
    location_id BIGINT NOT NULL REFERENCES att_locations(id),
    source_file VARCHAR(500) NOT NULL DEFAULT '',
    file_hash VARCHAR(128),
    report_from DATE,
    report_to DATE,
    employee_count INTEGER DEFAULT 0,
    records_processed INTEGER DEFAULT 0,
    records_inserted INTEGER DEFAULT 0,
    records_updated INTEGER DEFAULT 0,
    records_failed INTEGER DEFAULT 0,
    unknown_status_count INTEGER DEFAULT 0,
    data_quality_issue_count INTEGER DEFAULT 0,
    status VARCHAR(30) NOT NULL DEFAULT 'RUNNING',
    error_message TEXT,
    failures JSONB NOT NULL DEFAULT '[]'::jsonb,
    unknown_statuses JSONB NOT NULL DEFAULT '[]'::jsonb,
    data_quality_issues JSONB NOT NULL DEFAULT '[]'::jsonb,
    started_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    completed_at TIMESTAMPTZ,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS att_attendance (
    id BIGSERIAL PRIMARY KEY,
    employee_id BIGINT NOT NULL REFERENCES att_employees(id),
    location_id BIGINT NOT NULL REFERENCES att_locations(id),
    attendance_date DATE NOT NULL,
    status VARCHAR(30) NOT NULL,
    in_time TIME,
    out_time TIME,
    total_minutes INTEGER,
    total_duration INTERVAL,
    source_file VARCHAR(500),
    import_batch_id BIGINT REFERENCES att_import_batches(id),
    punch_expected BOOLEAN NOT NULL DEFAULT FALSE,
    imported_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    CONSTRAINT uq_employee_attendance_date UNIQUE (employee_id, attendance_date)
);

CREATE INDEX IF NOT EXISTS idx_attendance_location ON att_attendance (location_id);
CREATE INDEX IF NOT EXISTS idx_attendance_location_date ON att_attendance (location_id, attendance_date);
CREATE INDEX IF NOT EXISTS idx_attendance_employee ON att_attendance (employee_id);
CREATE INDEX IF NOT EXISTS idx_attendance_date ON att_attendance (attendance_date);
CREATE INDEX IF NOT EXISTS idx_employees_location ON att_employees (location_id);
CREATE INDEX IF NOT EXISTS idx_batches_location ON att_import_batches (location_id);
