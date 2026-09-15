"use strict";

/**
 * PostgreSQL bridge for the DDO++ attendance sidecar.
 * Shares the NocoBase 0.9.4-alpha.2 database (public.att_* tables).
 */

const fs = require("fs");
const path = require("path");

let pool = null;

function log(level, message, extra) {
  const line = extra ? `${message} ${JSON.stringify(extra)}` : message;
  console.log(`${new Date().toISOString()} ${level} ${line}`);
}

function sslConfig() {
  const mode = String(process.env.PGSSLMODE || process.env.DATABASE_SSL || "").toLowerCase();
  if (!mode || mode === "disable" || mode === "false" || mode === "0") return undefined;
  if (mode === "no-verify") return { rejectUnauthorized: false };
  if (
    mode === "require" ||
    mode === "true" ||
    mode === "1" ||
    mode === "prefer" ||
    mode === "verify-ca" ||
    mode === "verify-full"
  ) {
    return { rejectUnauthorized: process.env.PGSSL_REJECT_UNAUTHORIZED !== "false" };
  }
  return undefined;
}

function configFromEnv() {
  const ssl = sslConfig();
  if (process.env.DATABASE_URL) {
    const config = { connectionString: process.env.DATABASE_URL };
    if (ssl) config.ssl = ssl;
    return config;
  }
  const host = process.env.DB_HOST || process.env.DATABASE_HOST || process.env.PGHOST;
  const database = process.env.DB_DATABASE || process.env.DATABASE_DATABASE || process.env.PGDATABASE;
  const user = process.env.DB_USER || process.env.DATABASE_USER || process.env.PGUSER;
  const password = process.env.DB_PASSWORD || process.env.DATABASE_PASSWORD || process.env.PGPASSWORD;
  if (!host || !database || !user) return null;
  const config = {
    host,
    port: Number(process.env.DB_PORT || process.env.DATABASE_PORT || process.env.PGPORT || 5432),
    database,
    user,
    password: password || "",
  };
  if (ssl) config.ssl = ssl;
  return config;
}

function redactedConfig() {
  const config = configFromEnv();
  if (!config) return { configured: false };
  if (config.connectionString) {
    let host = "";
    let database = "";
    try {
      const parsed = new URL(config.connectionString.replace(/^postgres(ql)?:/i, "http:"));
      host = parsed.hostname;
      database = parsed.pathname.replace(/^\//, "");
    } catch {
      host = "(unparsed)";
    }
    return { configured: true, via: "DATABASE_URL", host, database, ssl: Boolean(config.ssl) };
  }
  return {
    configured: true,
    via: "discrete",
    host: config.host,
    port: config.port,
    database: config.database,
    user: config.user,
    ssl: Boolean(config.ssl),
  };
}

function isEnabled() {
  return Boolean(configFromEnv());
}

async function connect() {
  if (pool) return pool;
  const config = configFromEnv();
  if (!config) {
    throw new Error(
      "PostgreSQL is not configured. Set DATABASE_URL or DB_HOST/DB_DATABASE/DB_USER/DB_PASSWORD (NocoBase 0.9 env names also work)."
    );
  }
  const { Pool } = require("pg");
  pool = new Pool({
    ...config,
    max: Number(process.env.DB_POOL_SIZE || 10),
    idleTimeoutMillis: 30_000,
    connectionTimeoutMillis: Number(process.env.DB_CONNECT_TIMEOUT_MS || 10_000),
  });
  pool.on("error", (err) => log("ERROR", `PostgreSQL pool error: ${err.message}`));
  return pool;
}

async function ping() {
  const client = await (await connect()).connect();
  try {
    await client.query("SELECT 1");
    return true;
  } finally {
    client.release();
  }
}

async function ensureSchema(statusMaster) {
  const client = await (await connect()).connect();
  try {
    const sqlPath = path.join(__dirname, "schema", "postgres_baseline.sql");
    await client.query(fs.readFileSync(sqlPath, "utf8"));
    await client.query(`
      ALTER TABLE att_employees ADD COLUMN IF NOT EXISTS identity_status VARCHAR(30) NOT NULL DEFAULT 'CODED';
      ALTER TABLE att_import_batches ADD COLUMN IF NOT EXISTS batch_uid VARCHAR(64);
      ALTER TABLE att_import_batches ADD COLUMN IF NOT EXISTS unknown_status_count INTEGER DEFAULT 0;
      ALTER TABLE att_import_batches ADD COLUMN IF NOT EXISTS data_quality_issue_count INTEGER DEFAULT 0;
      ALTER TABLE att_import_batches ADD COLUMN IF NOT EXISTS failures JSONB NOT NULL DEFAULT '[]'::jsonb;
      ALTER TABLE att_import_batches ADD COLUMN IF NOT EXISTS unknown_statuses JSONB NOT NULL DEFAULT '[]'::jsonb;
      ALTER TABLE att_import_batches ADD COLUMN IF NOT EXISTS data_quality_issues JSONB NOT NULL DEFAULT '[]'::jsonb;
      ALTER TABLE att_attendance_statuses ADD COLUMN IF NOT EXISTS hr_definition_required BOOLEAN NOT NULL DEFAULT FALSE;
      ALTER TABLE att_attendance ADD COLUMN IF NOT EXISTS punch_expected BOOLEAN NOT NULL DEFAULT FALSE;
    `);
    await client.query(`
      CREATE UNIQUE INDEX IF NOT EXISTS uq_att_employees_code_location
        ON att_employees (location_id, employee_code);
      CREATE UNIQUE INDEX IF NOT EXISTS uq_att_import_batches_uid
        ON att_import_batches (batch_uid);
    `);
    for (const row of statusMaster) {
      await client.query(
        `INSERT INTO att_attendance_statuses
           (code, display_name, category, is_half_day, hr_definition_required, is_active, updated_at)
         VALUES ($1, $2, $3, $4, $5, TRUE, NOW())
         ON CONFLICT (code) DO UPDATE SET
           display_name = EXCLUDED.display_name,
           category = EXCLUDED.category,
           is_half_day = EXCLUDED.is_half_day,
           hr_definition_required = EXCLUDED.hr_definition_required,
           updated_at = NOW()`,
        [row.code, row.display_name, row.category, row.is_half_day, row.hr_definition_required]
      );
    }
    log("INFO", "PostgreSQL attendance schema is ready");
  } finally {
    client.release();
  }
}

function asTime(value) {
  if (!value) return null;
  const text = String(value).trim();
  return /^\d{1,2}:\d{2}(:\d{2})?$/.test(text) ? text : null;
}

async function persistImport({
  locationCode,
  batchId,
  sourceFile,
  fileHash,
  reportFrom,
  reportTo,
  employeeCount,
  records,
  recordsProcessed,
  failures,
  unknownStatuses,
  dataQualityIssues,
  startedAt,
}) {
  const client = await (await connect()).connect();
  let inserted = 0;
  let updated = 0;
  try {
    await client.query("BEGIN");

    const loc = await client.query(
      "SELECT id, location_name FROM att_locations WHERE location_code = $1 AND is_active = TRUE",
      [locationCode]
    );
    if (!loc.rowCount) {
      const error = new Error(`Unknown or inactive location_code: ${locationCode}`);
      error.statusCode = 400;
      throw error;
    }
    const locationId = loc.rows[0].id;

    const batchIns = await client.query(
      `INSERT INTO att_import_batches (
         batch_uid, location_id, source_file, file_hash, report_from, report_to,
         employee_count, status, started_at
       ) VALUES ($1, $2, $3, $4, $5, $6, $7, 'RUNNING', $8)
       ON CONFLICT (batch_uid) DO UPDATE SET
         source_file = EXCLUDED.source_file,
         file_hash = EXCLUDED.file_hash,
         status = 'RUNNING',
         updated_at = NOW()
       RETURNING id`,
      [batchId, locationId, sourceFile || "", fileHash || null, reportFrom, reportTo, employeeCount, startedAt]
    );
    const importBatchId = batchIns.rows[0].id;

    for (const record of records) {
      const emp = await client.query(
        `INSERT INTO att_employees (
           employee_code, employee_name, location_id, identity_status, updated_at
         ) VALUES ($1, $2, $3, $4, NOW())
         ON CONFLICT (location_id, employee_code) DO UPDATE SET
           employee_name = CASE
             WHEN EXCLUDED.employee_name <> '' THEN EXCLUDED.employee_name
             ELSE att_employees.employee_name
           END,
           identity_status = EXCLUDED.identity_status,
           updated_at = NOW()
         RETURNING id`,
        [record.employee_code, record.employee_name || "", locationId, record.identity_status || "CODED"]
      );
      const employeeId = emp.rows[0].id;

      const minutes = record.total_minutes == null || Number.isNaN(Number(record.total_minutes))
        ? null
        : Number(record.total_minutes);
      const upsert = await client.query(
        `INSERT INTO att_attendance (
           employee_id, location_id, attendance_date, status, in_time, out_time,
           total_minutes, total_duration, source_file, import_batch_id, punch_expected, imported_at, updated_at
         ) VALUES ($1, $2, $3, $4, $5::time, $6::time, $7, CASE WHEN $7 IS NULL THEN NULL ELSE make_interval(mins => $7) END, $8, $9, $10, $11, NOW())
         ON CONFLICT (employee_id, attendance_date) DO UPDATE SET
           status = EXCLUDED.status,
           in_time = EXCLUDED.in_time,
           out_time = EXCLUDED.out_time,
           total_minutes = EXCLUDED.total_minutes,
           total_duration = EXCLUDED.total_duration,
           source_file = EXCLUDED.source_file,
           import_batch_id = EXCLUDED.import_batch_id,
           punch_expected = EXCLUDED.punch_expected,
           updated_at = NOW()
         RETURNING (xmax = 0) AS inserted`,
        [
          employeeId,
          locationId,
          record.attendance_date,
          record.status,
          asTime(record.in_time),
          asTime(record.out_time),
          minutes,
          sourceFile || null,
          importBatchId,
          Boolean(record.punch_expected),
          startedAt,
        ]
      );
      if (upsert.rows[0].inserted) inserted += 1;
      else updated += 1;
    }

    const failed = failures.length;
    let status = "SUCCESS";
    if (failed && !inserted && !updated) status = "FAILED";
    else if (failed || dataQualityIssues.length) status = "PARTIAL";

    await client.query(
      `UPDATE att_import_batches SET
         records_processed = $2,
         records_inserted = $3,
         records_updated = $4,
         records_failed = $5,
         unknown_status_count = $6,
         data_quality_issue_count = $7,
         status = $8,
         error_message = $9,
         failures = $10::jsonb,
         unknown_statuses = $11::jsonb,
         data_quality_issues = $12::jsonb,
         completed_at = NOW(),
         updated_at = NOW()
       WHERE id = $1`,
      [
        importBatchId,
        recordsProcessed ?? (records.length + failed),
        inserted,
        updated,
        failed,
        unknownStatuses.length,
        dataQualityIssues.length,
        status,
        status === "FAILED" ? "All records failed validation" : null,
        JSON.stringify(failures),
        JSON.stringify(unknownStatuses),
        JSON.stringify(dataQualityIssues),
      ]
    );

    await client.query("COMMIT");
    return {
      ok: status !== "FAILED",
      batch_id: batchId,
      status,
      location_code: locationCode,
      records_processed: recordsProcessed ?? (records.length + failed),
      records_inserted: inserted,
      records_updated: updated,
      records_failed: failed,
      unknown_status_count: unknownStatuses.length,
      data_quality_issue_count: dataQualityIssues.length,
      failures,
    };
  } catch (err) {
    try {
      await client.query("ROLLBACK");
    } catch {
      /* ignore */
    }
    throw err;
  } finally {
    client.release();
  }
}

const ATTENDANCE_SELECT = `
  SELECT
    a.id,
    e.employee_code,
    e.employee_name,
    l.location_code,
    l.location_name,
    to_char(a.attendance_date, 'YYYY-MM-DD') AS attendance_date,
    a.status,
    s.display_name AS status_display,
    COALESCE(s.category, 'OTHER') AS status_category,
    (s.id IS NOT NULL) AS status_known,
    COALESCE(s.hr_definition_required, TRUE) AS hr_definition_required,
    COALESCE(s.is_half_day, FALSE) AS is_half_day,
    to_char(a.in_time, 'HH24:MI') AS in_time,
    to_char(a.out_time, 'HH24:MI') AS out_time,
    a.total_minutes,
    CASE WHEN a.total_minutes IS NULL THEN ''
         ELSE (a.total_minutes / 60)::text || ':' || lpad((a.total_minutes % 60)::text, 2, '0')
    END AS working_hours,
    e.identity_status,
    a.source_file,
    b.batch_uid AS import_batch_id,
    a.punch_expected,
    a.imported_at,
    a.updated_at
  FROM att_attendance a
  JOIN att_employees e ON e.id = a.employee_id
  JOIN att_locations l ON l.id = a.location_id
  LEFT JOIN att_attendance_statuses s ON s.code = a.status
  LEFT JOIN att_import_batches b ON b.id = a.import_batch_id
`;

function attendanceFilters(query) {
  const locationCode = (query.get("location_code") || "").trim().toUpperCase();
  const employeeCode = (query.get("employee_code") || "").trim();
  const from = query.get("from") || "";
  const to = query.get("to") || "";
  const status = (query.get("status") || "").trim().toUpperCase();
  const clauses = [];
  const params = [];
  if (locationCode && locationCode !== "ALL") {
    params.push(locationCode);
    clauses.push(`l.location_code = $${params.length}`);
  }
  if (employeeCode) {
    params.push(employeeCode);
    clauses.push(`e.employee_code = $${params.length}`);
  }
  if (from) {
    params.push(from);
    clauses.push(`a.attendance_date >= $${params.length}`);
  }
  if (to) {
    params.push(to);
    clauses.push(`a.attendance_date <= $${params.length}`);
  }
  if (status) {
    params.push(status);
    clauses.push(`upper(a.status) = $${params.length}`);
  }
  const where = clauses.length ? `WHERE ${clauses.join(" AND ")}` : "";
  return { where, params };
}

async function listAttendance(query) {
  const { where, params } = attendanceFilters(query);
  const result = await (await connect()).query(
    `${ATTENDANCE_SELECT} ${where}
     ORDER BY a.attendance_date, l.location_code, e.employee_code`,
    params
  );
  return result.rows;
}

async function listBatches() {
  const result = await (await connect()).query(
    `SELECT
       b.batch_uid AS batch_id,
       l.location_code,
       l.location_name,
       b.source_file,
       b.file_hash,
       to_char(b.report_from, 'YYYY-MM-DD') AS report_from,
       to_char(b.report_to, 'YYYY-MM-DD') AS report_to,
       b.employee_count,
       b.records_processed,
       b.records_inserted,
       b.records_updated,
       b.records_failed,
       b.unknown_status_count,
       b.data_quality_issue_count,
       b.status,
       b.error_message,
       b.failures,
       b.unknown_statuses,
       b.data_quality_issues,
       b.started_at,
       b.completed_at
     FROM att_import_batches b
     JOIN att_locations l ON l.id = b.location_id
     ORDER BY b.completed_at DESC NULLS LAST, b.id DESC`
  );
  return result.rows;
}

async function listEmployees(locationCode) {
  const params = [];
  let where = "";
  if (locationCode && locationCode !== "ALL") {
    params.push(locationCode);
    where = "WHERE l.location_code = $1";
  }
  const result = await (await connect()).query(
    `SELECT
       e.employee_code,
       e.employee_name,
       l.location_code,
       l.location_name,
       e.identity_status,
       e.updated_at
     FROM att_employees e
     JOIN att_locations l ON l.id = e.location_id
     ${where}
     ORDER BY e.employee_name, e.employee_code`,
    params
  );
  return result.rows;
}

async function listExceptions() {
  const attendance = await listAttendance(new URLSearchParams());
  const batches = await listBatches();
  const employees = await listEmployees("");
  return { attendance, batches, employees };
}

async function close() {
  if (pool) {
    await pool.end();
    pool = null;
  }
}

module.exports = {
  isEnabled,
  configFromEnv,
  redactedConfig,
  connect,
  ping,
  ensureSchema,
  persistImport,
  listAttendance,
  listBatches,
  listEmployees,
  listExceptions,
  close,
};
