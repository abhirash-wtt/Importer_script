"use strict";

/**
 * Registers existing public.att_* tables as NocoBase 0.9 collections.
 * Does not create a second copy of the tables. Safe to re-run.
 *
 * After this script: open Collection manager, confirm att_attendance,
 * then point page /admin/vb0faior0eg at that collection.
 */

require("../env");
const crypto = require("crypto");
const db = require("../db");

function log(level, message, extra) {
  const line = extra ? `${message} ${JSON.stringify(extra)}` : message;
  console.log(`${new Date().toISOString()} ${level} ${line}`);
}

function uid(seed) {
  return `a${crypto.createHash("sha1").update(seed).digest("hex").slice(0, 12)}`;
}

const COLLECTIONS = [
  {
    name: "att_locations",
    title: "Attendance Locations",
    fields: [
      { name: "id", type: "bigInt", interface: "integer", primaryKey: true, autoIncrement: true, uiTitle: "ID" },
      { name: "location_code", type: "string", interface: "input", uiTitle: "Location code" },
      { name: "location_name", type: "string", interface: "input", uiTitle: "Location name" },
      { name: "is_active", type: "boolean", interface: "checkbox", uiTitle: "Active" },
      { name: "created_at", type: "date", interface: "createdAt", uiTitle: "Created at" },
      { name: "updated_at", type: "date", interface: "updatedAt", uiTitle: "Updated at" },
    ],
  },
  {
    name: "att_departments",
    title: "Attendance Departments",
    fields: [
      { name: "id", type: "bigInt", interface: "integer", primaryKey: true, autoIncrement: true, uiTitle: "ID" },
      { name: "department_code", type: "string", interface: "input", uiTitle: "Department code" },
      { name: "department_name", type: "string", interface: "input", uiTitle: "Department name" },
      { name: "is_active", type: "boolean", interface: "checkbox", uiTitle: "Active" },
      { name: "created_at", type: "date", interface: "createdAt", uiTitle: "Created at" },
      { name: "updated_at", type: "date", interface: "updatedAt", uiTitle: "Updated at" },
    ],
  },
  {
    name: "att_employees",
    title: "Attendance Employees",
    fields: [
      { name: "id", type: "bigInt", interface: "integer", primaryKey: true, autoIncrement: true, uiTitle: "ID" },
      { name: "employee_code", type: "string", interface: "input", uiTitle: "Employee code" },
      { name: "employee_name", type: "string", interface: "input", uiTitle: "Employee name" },
      { name: "location_id", type: "bigInt", interface: "integer", uiTitle: "Location ID" },
      { name: "department_id", type: "bigInt", interface: "integer", uiTitle: "Department ID" },
      { name: "biometric_code", type: "string", interface: "input", uiTitle: "Biometric code" },
      { name: "identity_status", type: "string", interface: "input", uiTitle: "Identity status" },
      { name: "is_active", type: "boolean", interface: "checkbox", uiTitle: "Active" },
      { name: "created_at", type: "date", interface: "createdAt", uiTitle: "Created at" },
      { name: "updated_at", type: "date", interface: "updatedAt", uiTitle: "Updated at" },
    ],
  },
  {
    name: "att_attendance_statuses",
    title: "Attendance Statuses",
    fields: [
      { name: "id", type: "bigInt", interface: "integer", primaryKey: true, autoIncrement: true, uiTitle: "ID" },
      { name: "code", type: "string", interface: "input", uiTitle: "Code" },
      { name: "display_name", type: "string", interface: "input", uiTitle: "Display name" },
      { name: "category", type: "string", interface: "input", uiTitle: "Category" },
      { name: "is_half_day", type: "boolean", interface: "checkbox", uiTitle: "Half day" },
      { name: "hr_definition_required", type: "boolean", interface: "checkbox", uiTitle: "HR definition required" },
      { name: "is_active", type: "boolean", interface: "checkbox", uiTitle: "Active" },
      { name: "description", type: "text", interface: "textarea", uiTitle: "Description" },
      { name: "created_at", type: "date", interface: "createdAt", uiTitle: "Created at" },
      { name: "updated_at", type: "date", interface: "updatedAt", uiTitle: "Updated at" },
    ],
  },
  {
    name: "att_import_batches",
    title: "Attendance Import Batches",
    fields: [
      { name: "id", type: "bigInt", interface: "integer", primaryKey: true, autoIncrement: true, uiTitle: "ID" },
      { name: "batch_uid", type: "string", interface: "input", uiTitle: "Batch id" },
      { name: "location_id", type: "bigInt", interface: "integer", uiTitle: "Location ID" },
      { name: "source_file", type: "string", interface: "input", uiTitle: "Source file" },
      { name: "file_hash", type: "string", interface: "input", uiTitle: "File hash" },
      { name: "report_from", type: "date", interface: "date", uiTitle: "Report from" },
      { name: "report_to", type: "date", interface: "date", uiTitle: "Report to" },
      { name: "employee_count", type: "integer", interface: "integer", uiTitle: "Employees" },
      { name: "records_processed", type: "integer", interface: "integer", uiTitle: "Processed" },
      { name: "records_inserted", type: "integer", interface: "integer", uiTitle: "Inserted" },
      { name: "records_updated", type: "integer", interface: "integer", uiTitle: "Updated" },
      { name: "records_failed", type: "integer", interface: "integer", uiTitle: "Failed" },
      { name: "status", type: "string", interface: "input", uiTitle: "Status" },
      { name: "error_message", type: "text", interface: "textarea", uiTitle: "Error" },
      { name: "started_at", type: "date", interface: "datetime", uiTitle: "Started at" },
      { name: "completed_at", type: "date", interface: "datetime", uiTitle: "Completed at" },
      { name: "created_at", type: "date", interface: "createdAt", uiTitle: "Created at" },
      { name: "updated_at", type: "date", interface: "updatedAt", uiTitle: "Updated at" },
    ],
  },
  {
    name: "att_attendance",
    title: "Attendance",
    fields: [
      { name: "id", type: "bigInt", interface: "integer", primaryKey: true, autoIncrement: true, uiTitle: "ID" },
      { name: "employee_id", type: "bigInt", interface: "integer", uiTitle: "Employee ID" },
      { name: "location_id", type: "bigInt", interface: "integer", uiTitle: "Location ID" },
      { name: "attendance_date", type: "date", interface: "date", uiTitle: "Date" },
      { name: "status", type: "string", interface: "input", uiTitle: "Status" },
      { name: "in_time", type: "time", interface: "time", uiTitle: "In time" },
      { name: "out_time", type: "time", interface: "time", uiTitle: "Out time" },
      { name: "total_minutes", type: "integer", interface: "integer", uiTitle: "Total minutes" },
      { name: "source_file", type: "string", interface: "input", uiTitle: "Source file" },
      { name: "import_batch_id", type: "bigInt", interface: "integer", uiTitle: "Import batch ID" },
      { name: "punch_expected", type: "boolean", interface: "checkbox", uiTitle: "Punch expected" },
      { name: "imported_at", type: "date", interface: "datetime", uiTitle: "Imported at" },
      { name: "created_at", type: "date", interface: "createdAt", uiTitle: "Created at" },
      { name: "updated_at", type: "date", interface: "updatedAt", uiTitle: "Updated at" },
    ],
  },
];

async function columnSet(client, table) {
  const result = await client.query(
    `SELECT column_name
       FROM information_schema.columns
      WHERE table_schema = 'public' AND table_name = $1`,
    [table]
  );
  return new Set(result.rows.map((row) => row.column_name));
}

function pick(row, allowed) {
  const out = {};
  for (const key of allowed) {
    if (row[key] !== undefined) out[key] = row[key];
  }
  return out;
}

async function insertRow(client, table, columns, row) {
  const payload = pick(row, columns);
  const keys = Object.keys(payload);
  if (!keys.length) return;
  const values = keys.map((_, idx) => `$${idx + 1}`);
  await client.query(
    `INSERT INTO ${table} (${keys.map((key) => `"${key}"`).join(", ")})
     VALUES (${values.join(", ")})`,
    keys.map((key) => payload[key])
  );
}

function collectionRow(spec, columns) {
  const now = new Date();
  const options = {
    underscored: true,
    timestamps: true,
    createdAt: "created_at",
    updatedAt: "updated_at",
    filterTargetKey: "id",
  };
  return {
    key: uid(`collection:${spec.name}`),
    name: spec.name,
    title: spec.title,
    inherit: false,
    hidden: false,
    sortable: false,
    description: "DDO++ attendance sidecar table. Do not recreate; sidecar writes this table.",
    options: columns.includes("options") ? options : undefined,
    createdAt: now,
    updatedAt: now,
    created_at: now,
    updated_at: now,
  };
}

function fieldRow(collectionName, field, columns) {
  const now = new Date();
  const uiSchema = {
    type: field.type === "boolean" ? "boolean" : "string",
    title: field.uiTitle,
    "x-component": field.interface === "textarea" ? "Input.TextArea" : "Input",
  };
  const options = {
    primaryKey: Boolean(field.primaryKey),
    autoIncrement: Boolean(field.autoIncrement),
  };
  return {
    key: uid(`field:${collectionName}:${field.name}`),
    name: field.name,
    type: field.type,
    interface: field.interface,
    collectionName: collectionName,
    collection_name: collectionName,
    parentKey: null,
    reverseKey: null,
    uiSchema: columns.includes("uiSchema") ? uiSchema : undefined,
    options: columns.includes("options") ? options : undefined,
    createdAt: now,
    updatedAt: now,
    created_at: now,
    updated_at: now,
  };
}

async function main() {
  if (!db.isEnabled()) {
    log("ERROR", "PostgreSQL is not configured. Set DATABASE_URL first.");
    process.exitCode = 1;
    return;
  }
  const client = await (await db.connect()).connect();
  try {
    const collectionCols = [...(await columnSet(client, "collections"))];
    if (!collectionCols.length) {
      log("ERROR", "NocoBase collections table was not found. Is this the NocoBase database?");
      process.exitCode = 1;
      return;
    }
    const fieldCols = [...(await columnSet(client, "fields"))];
    if (!fieldCols.length) {
      log("ERROR", "NocoBase fields table was not found.");
      process.exitCode = 1;
      return;
    }
    log("INFO", "NocoBase metadata tables", {
      collections_columns: collectionCols,
      fields_columns: fieldCols,
    });

    const nameColumn = fieldCols.includes("collectionName")
      ? "collectionName"
      : fieldCols.includes("collection_name")
        ? "collection_name"
        : null;
    if (!nameColumn) {
      log("ERROR", "Could not find collectionName column on fields table", { fieldCols });
      process.exitCode = 1;
      return;
    }

    for (const spec of COLLECTIONS) {
      const existing = await client.query("SELECT name FROM collections WHERE name = $1", [spec.name]);
      if (!existing.rowCount) {
        await insertRow(client, "collections", collectionCols, collectionRow(spec, collectionCols));
        log("INFO", `Created collection ${spec.name}`);
      } else {
        log("INFO", `Collection already exists: ${spec.name}`);
      }

      for (const field of spec.fields) {
        const found = await client.query(
          `SELECT name FROM fields WHERE "${nameColumn}" = $1 AND name = $2`,
          [spec.name, field.name]
        );
        if (found.rowCount) continue;
        await insertRow(client, "fields", fieldCols, fieldRow(spec.name, field, fieldCols));
        log("INFO", `  added field ${spec.name}.${field.name}`);
      }
    }
    log("INFO", "Phase 3 collection registration complete.");
    log("INFO", "Restart NocoBase, then open Collection manager and bind page vb0faior0eg to att_attendance.");
  } finally {
    client.release();
    await db.close();
  }
}

main().catch((err) => {
  log("ERROR", err.message || String(err));
  process.exitCode = 1;
});
