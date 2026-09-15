"use strict";

require("../env");
const db = require("../db");

function log(level, message, extra) {
  const line = extra ? `${message} ${JSON.stringify(extra)}` : message;
  console.log(`${new Date().toISOString()} ${level} ${line}`);
}

const REQUIRED_TABLES = [
  "att_locations",
  "att_departments",
  "att_employees",
  "att_attendance_statuses",
  "att_import_batches",
  "att_attendance",
];

async function main() {
  const summary = db.redactedConfig();
  if (!summary.configured) {
    log("ERROR", "PostgreSQL is not configured. Set DATABASE_URL or DB_HOST/DB_DATABASE/DB_USER/DB_PASSWORD.");
    process.exitCode = 1;
    return;
  }
  log("INFO", "Checking NocoBase Postgres", summary);
  await db.connect();
  await db.ping();
  const client = await (await db.connect()).connect();
  try {
    const info = await client.query(
      "SELECT current_database() AS database, current_user AS user, version() AS version"
    );
    const row = info.rows[0];
    log("INFO", "Connected", {
      database: row.database,
      user: row.user,
      version: String(row.version).split(",")[0],
    });
    const tables = await client.query(
      `SELECT table_name
         FROM information_schema.tables
        WHERE table_schema = 'public'
          AND table_name = ANY($1::text[])
        ORDER BY table_name`,
      [REQUIRED_TABLES]
    );
    const found = new Set(tables.rows.map((item) => item.table_name));
    const missing = REQUIRED_TABLES.filter((name) => !found.has(name));
    log("INFO", "Attendance tables", { present: [...found], missing });
    if (missing.length) {
      log("WARN", "Run npm run apply-schema to create missing att_* tables.");
      process.exitCode = 2;
      return;
    }
    const locations = await client.query(
      "SELECT location_code FROM att_locations ORDER BY location_code"
    );
    log("INFO", "Locations", { codes: locations.rows.map((item) => item.location_code) });
    log("INFO", "Phase 1 Postgres check passed.");
  } finally {
    client.release();
    await db.close();
  }
}

main().catch((err) => {
  log("ERROR", err.message || String(err));
  process.exitCode = 1;
});
