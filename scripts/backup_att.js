"use strict";

require("../env");
const { spawnSync } = require("child_process");
const fs = require("fs");
const path = require("path");
const db = require("../db");

function log(level, message, extra) {
  const line = extra ? `${message} ${JSON.stringify(extra)}` : message;
  console.log(`${new Date().toISOString()} ${level} ${line}`);
}

const TABLES = [
  "att_locations",
  "att_departments",
  "att_employees",
  "att_attendance_statuses",
  "att_import_batches",
  "att_attendance",
];

function stamp() {
  return new Date().toISOString().replace(/[:.]/g, "-");
}

async function jsonBackup(outFile) {
  const client = await (await db.connect()).connect();
  try {
    const dump = { dumped_at: new Date().toISOString(), tables: {} };
    for (const table of TABLES) {
      const result = await client.query(`SELECT * FROM ${table}`);
      dump.tables[table] = result.rows;
    }
    fs.writeFileSync(outFile, JSON.stringify(dump, null, 2), "utf8");
  } finally {
    client.release();
    await db.close();
  }
}

async function main() {
  if (!db.isEnabled()) {
    throw new Error("PostgreSQL is not configured. Set DATABASE_URL first.");
  }
  const outDir = path.join(__dirname, "..", "backups");
  fs.mkdirSync(outDir, { recursive: true });
  const base = path.join(outDir, `att_${stamp()}`);
  const pgDump = spawnSync("pg_dump", ["--version"], { encoding: "utf8" });
  if (pgDump.status === 0 && process.env.DATABASE_URL) {
    const file = `${base}.sql`;
    const dumped = spawnSync(
      "pg_dump",
      ["--data-only", "--no-owner", ...TABLES.flatMap((table) => ["-t", table]), "-f", file, process.env.DATABASE_URL],
      { encoding: "utf8" }
    );
    if (dumped.status === 0) {
      log("INFO", "pg_dump backup written", { file });
      return;
    }
    log("WARN", "pg_dump failed; writing JSON backup", { error: dumped.stderr || dumped.error });
  }
  const file = `${base}.json`;
  await jsonBackup(file);
  log("INFO", "JSON backup written", { file });
}

main().catch((err) => {
  log("ERROR", err.message || String(err));
  process.exitCode = 1;
});
