"use strict";

require("../env");
const db = require("../db");
const { STATUS_MASTER } = require("../status_master");

function log(level, message, extra) {
  const line = extra ? `${message} ${JSON.stringify(extra)}` : message;
  console.log(`${new Date().toISOString()} ${level} ${line}`);
}

async function main() {
  if (!db.isEnabled()) {
    log("ERROR", "PostgreSQL is not configured. Set DATABASE_URL first.");
    process.exitCode = 1;
    return;
  }
  log("INFO", "Applying attendance schema", db.redactedConfig());
  await db.connect();
  await db.ensureSchema(STATUS_MASTER);
  log("INFO", "Phase 2 schema apply complete. att_* tables are ready.");
  await db.close();
}

main().catch((err) => {
  log("ERROR", err.message || String(err));
  process.exitCode = 1;
});
