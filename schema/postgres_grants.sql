-- Dedicated sidecar role for production. Run as a Postgres superuser / NocoBase owner AFTER att_* tables exist.
-- Replace CHANGE_ME and adjust the database name if NocoBase does not use "nocobase".

DO $$
BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'ddo_att_importer') THEN
    CREATE ROLE ddo_att_importer LOGIN PASSWORD 'CHANGE_ME';
  END IF;
END
$$;

GRANT CONNECT ON DATABASE nocobase TO ddo_att_importer;
GRANT USAGE ON SCHEMA public TO ddo_att_importer;

GRANT SELECT, INSERT, UPDATE ON
  att_locations,
  att_departments,
  att_employees,
  att_attendance_statuses,
  att_import_batches,
  att_attendance
TO ddo_att_importer;

GRANT USAGE, SELECT ON ALL SEQUENCES IN SCHEMA public TO ddo_att_importer;

ALTER DEFAULT PRIVILEGES IN SCHEMA public
  GRANT USAGE, SELECT ON SEQUENCES TO ddo_att_importer;
