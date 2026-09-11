"use strict";

/**
 * DDO++ Attendance sidecar API — receiving service for office importer POSTs.
 *
 * Deploy this on ddoplusnodeapi.walkingtree.tech (with db.js + schema/).
 *
 * POST /api/attendance/import
 * GET  /api/attendance
 * GET  /api/import-history
 * GET  /api/dashboard
 * GET  /api/employees
 * GET  /api/exceptions
 * GET  /api/statuses
 * GET  /api/locations
 * GET  /health
 * UI    /
 *
 * Run:  node attendance_api.js
 */

const fs = require("fs");
const http = require("http");
const https = require("https");
const path = require("path");
const { randomUUID } = require("crypto");
const { URL } = require("url");
const db = require("./db");

const BASE_DIR = __dirname;
loadEnvFile(path.join(BASE_DIR, ".env"));

const ALLOWED_LOCATIONS = {
  AGRA: "Agra",
  NOIDA: "Noida",
  HYD: "Hyderabad",
};

const STATUS_MASTER = [
  { code: "P", display_name: "Present", category: "PRESENT", is_half_day: false, hr_definition_required: false },
  { code: "A", display_name: "Absent", category: "ABSENT", is_half_day: false, hr_definition_required: false },
  { code: "WO", display_name: "Week Off", category: "WEEK_OFF", is_half_day: false, hr_definition_required: false },
  { code: "H", display_name: "Holiday", category: "HOLIDAY", is_half_day: false, hr_definition_required: false },
  { code: "CL", display_name: "Casual Leave", category: "LEAVE", is_half_day: false, hr_definition_required: false },
  { code: "EL", display_name: "Earned Leave", category: "LEAVE", is_half_day: false, hr_definition_required: false },
  { code: "WFH", display_name: "Work From Home", category: "WFH", is_half_day: false, hr_definition_required: false },
  { code: "LOP", display_name: "Loss of Pay", category: "ABSENT", is_half_day: false, hr_definition_required: false },
  { code: "AEL", display_name: "AEL", category: "OTHER", is_half_day: false, hr_definition_required: true },
  { code: "CO", display_name: "Comp Off", category: "OTHER", is_half_day: false, hr_definition_required: true },
  { code: "HP", display_name: "HP", category: "OTHER", is_half_day: false, hr_definition_required: true },
  { code: "CLP", display_name: "CLP", category: "LEAVE", is_half_day: false, hr_definition_required: true },
  { code: "WOP", display_name: "WOP", category: "OTHER", is_half_day: false, hr_definition_required: true },
  { code: "ELP", display_name: "ELP", category: "LEAVE", is_half_day: false, hr_definition_required: true },
  { code: "½CL", display_name: "Half Casual Leave", category: "LEAVE", is_half_day: true, hr_definition_required: true },
  { code: "½CLP", display_name: "Half CLP", category: "LEAVE", is_half_day: true, hr_definition_required: true },
  { code: "½EL", display_name: "Half Earned Leave", category: "LEAVE", is_half_day: true, hr_definition_required: true },
  { code: "½ELP", display_name: "Half ELP", category: "LEAVE", is_half_day: true, hr_definition_required: true },
  { code: "½WFHP", display_name: "Half WFH", category: "WFH", is_half_day: true, hr_definition_required: true },
  { code: "½AELP", display_name: "Half AELP", category: "OTHER", is_half_day: true, hr_definition_required: true },
  { code: "WFHP", display_name: "WFHP", category: "WFH", is_half_day: false, hr_definition_required: true },
];

function isProduction() {
  return String(process.env.NODE_ENV || "").toLowerCase() === "production";
}

function loadEnvFile(filePath) {
  if (!fs.existsSync(filePath)) return;
  for (const raw of fs.readFileSync(filePath, "utf8").split(/\r?\n/)) {
    const line = raw.trim();
    if (!line || line.startsWith("#") || !line.includes("=")) continue;
    const eq = line.indexOf("=");
    const key = line.slice(0, eq).trim();
    const value = line.slice(eq + 1).trim().replace(/^['"]|['"]$/g, "");
    if (!(key in process.env)) process.env[key] = value;
  }
}

const STATUS_INDEX = new Map(STATUS_MASTER.map((row) => [normalizeStatusKey(row.code), row]));
const PUBLIC_DIR = path.join(BASE_DIR, "public");
const DATA_DIR = process.env.DDO_DATA_DIR || path.join(BASE_DIR, "api_data");
const PORT = Number(process.env.DDO_API_PORT || 3000);
const BIND = (process.env.DDO_API_BIND || "0.0.0.0").trim();
const POC_TOKEN = "dev-ddo-attendance-token";
const TOKEN = (process.env.DDO_API_TOKEN || (isProduction() ? "" : POC_TOKEN)).trim();
const TLS_CERT = (process.env.DDO_TLS_CERT || "").trim();
const TLS_KEY = (process.env.DDO_TLS_KEY || "").trim();
const MAX_BODY_BYTES = 10 * 1024 * 1024;
const ALLOW_IPS = String(process.env.DDO_API_ALLOW_IPS || "")
  .split(",")
  .map((item) => item.trim())
  .filter(Boolean);
const TRUST_PROXY =
  process.env.DDO_TRUST_PROXY === "0" || process.env.DDO_TRUST_PROXY === "false"
    ? false
    : isProduction() || process.env.DDO_TRUST_PROXY === "1" || process.env.DDO_TRUST_PROXY === "true";
const CORS_ORIGIN = process.env.DDO_CORS_ORIGIN || "*";

const files = {
  attendance: path.join(DATA_DIR, "attendance.json"),
  batches: path.join(DATA_DIR, "batches.json"),
  employees: path.join(DATA_DIR, "employees.json"),
};

if (!isProduction()) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

function log(level, message, extra) {
  const line = extra ? `${message} ${JSON.stringify(extra)}` : message;
  console.log(`${new Date().toISOString()} ${level} ${line}`);
}

function readStore(filePath) {
  if (!fs.existsSync(filePath)) return {};
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function writeStore(filePath, data) {
  fs.writeFileSync(filePath, JSON.stringify(data, null, 2), "utf8");
}

function sendJson(res, statusCode, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body),
    "Cache-Control": "no-store",
    "Access-Control-Allow-Origin": CORS_ORIGIN,
    "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Location-Code",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  });
  res.end(body);
}

function readBearerToken(req) {
  const header = req.headers.authorization || "";
  const match = header.match(/^Bearer\s+(.+)$/i);
  return match ? match[1].trim() : "";
}

function requireAuth(req, res) {
  const provided = readBearerToken(req);
  if (!provided || provided !== TOKEN) {
    sendJson(res, 401, { ok: false, error: "Unauthorized" });
    return false;
  }
  return true;
}

function normalizeIp(ip) {
  if (!ip) return "";
  const value = String(ip).trim();
  if (value.startsWith("::ffff:")) return value.slice(7);
  return value;
}

function clientIp(req) {
  if (TRUST_PROXY) {
    const forwarded = req.headers["x-forwarded-for"];
    if (forwarded) return normalizeIp(String(forwarded).split(",")[0]);
    if (req.headers["x-real-ip"]) return normalizeIp(req.headers["x-real-ip"]);
  }
  return normalizeIp(req.socket && req.socket.remoteAddress);
}

function parseIpv4(ip) {
  const parts = String(ip).split(".").map(Number);
  if (parts.length !== 4 || parts.some((n) => Number.isNaN(n) || n < 0 || n > 255)) return null;
  return ((parts[0] << 24) >>> 0) + (parts[1] << 16) + (parts[2] << 8) + parts[3];
}

function ipMatches(ip, rule) {
  if (ip === rule) return true;
  if (!rule.includes("/")) return false;
  const [range, bitsRaw] = rule.split("/");
  const ipNum = parseIpv4(ip);
  const rangeNum = parseIpv4(range);
  const bits = Number(bitsRaw);
  if (ipNum == null || rangeNum == null || Number.isNaN(bits) || bits < 0 || bits > 32) return false;
  const mask = bits === 0 ? 0 : (~0 << (32 - bits)) >>> 0;
  return (ipNum & mask) === (rangeNum & mask);
}

function allowImportIp(req, res) {
  if (!ALLOW_IPS.length) return true;
  const ip = clientIp(req);
  if (ALLOW_IPS.some((rule) => ipMatches(ip, rule))) return true;
  log("WARN", "Import POST rejected by IP allow-list", { ip });
  sendJson(res, 403, { ok: false, error: "Forbidden" });
  return false;
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > MAX_BODY_BYTES) {
        reject(Object.assign(new Error("Payload too large"), { statusCode: 413 }));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

function isValidDate(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(value || "") && !Number.isNaN(Date.parse(`${value}T00:00:00Z`));
}

function attendanceKey(locationCode, employeeCode, attendanceDate) {
  return `${locationCode}|${employeeCode}|${attendanceDate}`;
}

function normalizeStatusKey(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/1\/2/g, "½")
    .replace(/\s+/g, "");
}

function lookupStatus(code) {
  const key = normalizeStatusKey(code);
  return STATUS_INDEX.get(key) || null;
}

function minutesToHours(minutes) {
  if (minutes == null || Number.isNaN(Number(minutes))) return "";
  const total = Number(minutes);
  const hh = Math.floor(total / 60);
  const mm = String(total % 60).padStart(2, "0");
  return `${hh}:${mm}`;
}

function normalizeRecord(raw, locationCode, index) {
  const employeeCode = String(raw.employee_code || "").trim();
  const attendanceDate = String(raw.attendance_date || raw.date || "").trim();
  const status = String(raw.status || "").trim();
  const failures = [];

  if (!employeeCode) {
    failures.push({ index, reason: "Unknown employee", employee_code: "" });
  }
  if (!isValidDate(attendanceDate)) {
    failures.push({ index, employee_code: employeeCode, reason: "Invalid attendance_date" });
  }
  if (!status) {
    failures.push({
      index,
      employee_code: employeeCode,
      attendance_date: attendanceDate,
      reason: "Missing status",
    });
  }

  const statusMeta = lookupStatus(status);
  return {
    failures,
    record: {
      employee_code: employeeCode,
      employee_name: String(raw.employee_name || "").trim(),
      location_code: locationCode,
      attendance_date: attendanceDate,
      status,
      status_display: statusMeta ? statusMeta.display_name : status,
      status_category: statusMeta ? statusMeta.category : "OTHER",
      status_known: Boolean(statusMeta),
      hr_definition_required: statusMeta ? statusMeta.hr_definition_required : true,
      is_half_day: statusMeta ? statusMeta.is_half_day : /½|1\/2/.test(status),
      in_time: raw.in_time || null,
      out_time: raw.out_time || null,
      total_minutes:
        raw.total_minutes == null || raw.total_minutes === "" ? null : Number(raw.total_minutes),
      working_hours: minutesToHours(
        raw.total_minutes == null || raw.total_minutes === "" ? null : Number(raw.total_minutes)
      ),
      identity_status: raw.identity_status || (String(employeeCode).startsWith("UNMAPPED:") ? "UNMAPPED_NAME" : "CODED"),
      punch_expected: Boolean(raw.punch_expected),
      source_file: null,
      import_batch_id: null,
      imported_at: null,
      updated_at: null,
    },
  };
}

function importAttendance(payload, headerLocation) {
  const locationCode = String(payload.location_code || "").trim().toUpperCase();
  if (!ALLOWED_LOCATIONS[locationCode]) {
    const error = new Error("Invalid location_code. Use AGRA, NOIDA, or HYD.");
    error.statusCode = 400;
    throw error;
  }
  if (headerLocation && String(headerLocation).toUpperCase() !== locationCode) {
    const error = new Error("X-Location-Code does not match location_code");
    error.statusCode = 400;
    throw error;
  }
  if (!Array.isArray(payload.records)) {
    const error = new Error("records must be an array");
    error.statusCode = 400;
    throw error;
  }

  const batchId = String(payload.batch_id || randomUUID());
  const sourceFile = String(payload.source_file || "");
  const fileHash = String(payload.file_hash || "");
  const importedAt = new Date().toISOString();
  const failures = Array.isArray(payload.failures) ? [...payload.failures] : [];
  const unknownStatuses = [];
  const dataQualityIssues = Array.isArray(payload.data_quality_issues)
    ? [...payload.data_quality_issues]
    : [];
  const validRecords = [];

  payload.records.forEach((raw, index) => {
    const { record, failures: recordFailures } = normalizeRecord(raw, locationCode, index);
    if (recordFailures.length) {
      failures.push(...recordFailures);
      return;
    }
    if (!record.status_known) {
      unknownStatuses.push({
        employee_code: record.employee_code,
        attendance_date: record.attendance_date,
        status: record.status,
      });
    }
    record.source_file = sourceFile;
    record.import_batch_id = batchId;
    record.imported_at = importedAt;
    record.updated_at = importedAt;
    validRecords.push(record);
  });

  const employeeCount =
    payload.employee_count ||
    new Set(validRecords.map((row) => String(row.employee_code || "").trim()).filter(Boolean)).size;

  if (db.isEnabled()) {
    return db.persistImport({
      locationCode,
      batchId,
      sourceFile,
      fileHash,
      reportFrom: payload.report_from || null,
      reportTo: payload.report_to || null,
      employeeCount,
      records: validRecords,
      recordsProcessed: payload.records.length,
      failures,
      unknownStatuses,
      dataQualityIssues,
      startedAt: importedAt,
    });
  }

  if (isProduction()) {
    const error = new Error("PostgreSQL is required in production. JSON file store is disabled.");
    error.statusCode = 503;
    throw error;
  }

  return persistImportJson({
    locationCode,
    batchId,
    sourceFile,
    fileHash,
    payload,
    validRecords,
    failures,
    unknownStatuses,
    dataQualityIssues,
    importedAt,
  });
}

function persistImportJson({
  locationCode,
  batchId,
  sourceFile,
  fileHash,
  payload,
  validRecords,
  failures,
  unknownStatuses,
  dataQualityIssues,
  importedAt,
}) {
  const attendance = readStore(files.attendance);
  const employees = readStore(files.employees);
  const batches = readStore(files.batches);
  let inserted = 0;
  let updated = 0;

  if (fileHash) {
    const duplicate = Object.values(batches).find(
      (row) => row.file_hash === fileHash && row.location_code === locationCode && row.status === "SUCCESS"
    );
    if (duplicate) {
      log("INFO", "Duplicate file hash detected; reprocessing with UPSERT", {
        file_hash: fileHash,
        previous_batch_id: duplicate.batch_id,
      });
    }
  }

  for (const record of validRecords) {
    const key = attendanceKey(locationCode, record.employee_code, record.attendance_date);
    const existing = attendance[key];
    if (existing) {
      record.imported_at = existing.imported_at;
      updated += 1;
    } else {
      inserted += 1;
    }
    attendance[key] = record;
    const employeeKey = `${locationCode}|${record.employee_code}`;
    employees[employeeKey] = {
      employee_code: record.employee_code,
      employee_name: record.employee_name || (employees[employeeKey] || {}).employee_name || "",
      location_code: locationCode,
      location_name: ALLOWED_LOCATIONS[locationCode],
      identity_status: record.identity_status,
      updated_at: importedAt,
    };
  }

  const processed = payload.records.length;
  const failed = failures.length;
  let status = "SUCCESS";
  if (failed && !inserted && !updated) status = "FAILED";
  else if (failed || dataQualityIssues.length) status = "PARTIAL";

  batches[batchId] = {
    batch_id: batchId,
    location_code: locationCode,
    location_name: ALLOWED_LOCATIONS[locationCode],
    source_file: sourceFile,
    file_hash: fileHash,
    report_from: payload.report_from || null,
    report_to: payload.report_to || null,
    employee_count:
      payload.employee_count ||
      new Set(payload.records.map((row) => String(row.employee_code || "").trim()).filter(Boolean)).size,
    records_processed: processed,
    records_inserted: inserted,
    records_updated: updated,
    records_failed: failed,
    unknown_status_count: unknownStatuses.length,
    data_quality_issue_count: dataQualityIssues.length,
    status,
    error_message: status === "FAILED" ? "All records failed validation" : null,
    started_at: importedAt,
    completed_at: importedAt,
    failures,
    unknown_statuses: unknownStatuses,
    data_quality_issues: dataQualityIssues,
  };

  writeStore(files.attendance, attendance);
  writeStore(files.employees, employees);
  writeStore(files.batches, batches);

  return {
    ok: status !== "FAILED",
    batch_id: batchId,
    status,
    location_code: locationCode,
    records_processed: processed,
    records_inserted: inserted,
    records_updated: updated,
    records_failed: failed,
    unknown_status_count: unknownStatuses.length,
    data_quality_issue_count: dataQualityIssues.length,
    failures,
  };
}

async function filterAttendance(query) {
  if (db.isEnabled()) {
    return db.listAttendance(query);
  }
  const attendance = Object.values(readStore(files.attendance));
  const locationCode = (query.get("location_code") || "").trim().toUpperCase();
  const employeeCode = (query.get("employee_code") || "").trim();
  const from = query.get("from") || "";
  const to = query.get("to") || "";
  const status = (query.get("status") || "").trim().toUpperCase();
  return attendance
    .filter((row) => !locationCode || locationCode === "ALL" || row.location_code === locationCode)
    .filter((row) => !employeeCode || row.employee_code === employeeCode)
    .filter((row) => !from || row.attendance_date >= from)
    .filter((row) => !to || row.attendance_date <= to)
    .filter((row) => !status || normalizeStatusKey(row.status) === status)
    .sort(
      (a, b) =>
        a.attendance_date.localeCompare(b.attendance_date) ||
        a.location_code.localeCompare(b.location_code) ||
        a.employee_code.localeCompare(b.employee_code)
    );
}

function emptyMetrics() {
  return {
    total_employees: 0,
    present: 0,
    absent: 0,
    leave: 0,
    week_off: 0,
    wfh: 0,
    holiday: 0,
    other: 0,
    missing_punches: 0,
    working_minutes: 0,
  };
}

function accumulate(metrics, row, employeeSet) {
  employeeSet.add(`${row.location_code}|${row.employee_code}`);
  const category = row.status_category || lookupStatus(row.status)?.category || "OTHER";
  if (category === "PRESENT") metrics.present += 1;
  else if (category === "ABSENT") metrics.absent += 1;
  else if (category === "LEAVE") metrics.leave += 1;
  else if (category === "WEEK_OFF") metrics.week_off += 1;
  else if (category === "WFH") metrics.wfh += 1;
  else if (category === "HOLIDAY") metrics.holiday += 1;
  else metrics.other += 1;
  if (row.punch_expected && (category === "PRESENT" || category === "WFH") && (!row.in_time || !row.out_time)) {
    metrics.missing_punches += 1;
  }
  if (row.total_minutes) metrics.working_minutes += Number(row.total_minutes) || 0;
}

async function dashboard(query) {
  const rows = await filterAttendance(query);
  const locationCode = (query.get("location_code") || "ALL").trim().toUpperCase() || "ALL";
  const consolidated = emptyMetrics();
  const employees = new Set();
  const byLocation = {};
  for (const code of Object.keys(ALLOWED_LOCATIONS)) {
    byLocation[code] = { location_code: code, location_name: ALLOWED_LOCATIONS[code], ...emptyMetrics(), _employees: new Set() };
  }
  for (const row of rows) {
    accumulate(consolidated, row, employees);
    const loc = byLocation[row.location_code];
    if (loc) accumulate(loc, row, loc._employees);
  }
  consolidated.total_employees = employees.size;
  const locations = Object.values(byLocation).map((loc) => {
    const { _employees, ...rest } = loc;
    rest.total_employees = _employees.size;
    rest.working_hours = minutesToHours(rest.working_minutes);
    return rest;
  });
  return {
    location_code: locationCode,
    from: query.get("from") || null,
    to: query.get("to") || null,
    record_count: rows.length,
    ...consolidated,
    working_hours: minutesToHours(consolidated.working_minutes),
    by_location: locations,
  };
}

async function listExceptions() {
  if (db.isEnabled()) {
    const { attendance: rows, batches, employees } = await db.listExceptions();
    return buildExceptions(rows, batches, employees);
  }
  const rows = Object.values(readStore(files.attendance));
  const batches = Object.values(readStore(files.batches));
  const employees = Object.values(readStore(files.employees));
  return buildExceptions(rows, batches, employees);
}

function buildExceptions(rows, batches, employees) {
  const missingPunches = rows
    .filter((row) => row.punch_expected && (row.status_category === "PRESENT" || row.status_category === "WFH") && (!row.in_time || !row.out_time))
    .map((row) => ({
      type: "MISSING_PUNCH",
      employee_code: row.employee_code,
      employee_name: row.employee_name,
      location_code: row.location_code,
      attendance_date: row.attendance_date,
      status: row.status,
      in_time: row.in_time,
      out_time: row.out_time,
    }));
  const unknownStatuses = rows
    .filter((row) => !row.status_known)
    .map((row) => ({
      type: "UNKNOWN_STATUS",
      employee_code: row.employee_code,
      employee_name: row.employee_name,
      location_code: row.location_code,
      attendance_date: row.attendance_date,
      status: row.status,
    }));
  const unmappedEmployees = employees
    .filter((row) => row.identity_status === "UNMAPPED_NAME")
    .map((row) => ({
      type: "UNMAPPED_EMPLOYEE",
      employee_code: row.employee_code,
      employee_name: row.employee_name,
      location_code: row.location_code,
    }));
  const batchIssues = batches.flatMap((batch) =>
    (batch.data_quality_issues || []).map((issue) => ({
      type: "DATA_QUALITY",
      batch_id: batch.batch_id,
      location_code: batch.location_code,
      ...issue,
    }))
  );
  return {
    missing_punches: missingPunches,
    unknown_statuses: unknownStatuses,
    unmapped_employees: unmappedEmployees,
    data_quality_issues: batchIssues,
  };
}

function contentType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  return (
    {
      ".html": "text/html; charset=utf-8",
      ".css": "text/css; charset=utf-8",
      ".js": "application/javascript; charset=utf-8",
      ".json": "application/json; charset=utf-8",
      ".svg": "image/svg+xml",
      ".ico": "image/x-icon",
    }[ext] || "application/octet-stream"
  );
}

function servePublic(req, res, url) {
  let relative = url.pathname === "/" ? "/index.html" : url.pathname;
  relative = path.normalize(relative).replace(/^(\.\.[/\\])+/, "");
  const filePath = path.join(PUBLIC_DIR, relative);
  if (!filePath.startsWith(PUBLIC_DIR)) {
    sendJson(res, 403, { ok: false, error: "Forbidden" });
    return true;
  }
  if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) return false;
  const body = fs.readFileSync(filePath);
  res.writeHead(200, { "Content-Type": contentType(filePath), "Cache-Control": "no-store" });
  res.end(body);
  return true;
}

async function handleRequest(req, res) {
  const url = new URL(req.url, `http://${req.headers.host || "localhost"}`);
  const route = `${req.method} ${url.pathname}`;

  try {
    if (req.method === "OPTIONS") {
      res.writeHead(204, {
        "Access-Control-Allow-Origin": CORS_ORIGIN,
        "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Location-Code",
        "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      });
      res.end();
      return;
    }

    if (route === "GET /health") {
      let database = { enabled: db.isEnabled(), ok: false };
      if (db.isEnabled()) {
        try {
          await db.ping();
          database = { enabled: true, ok: true };
        } catch (err) {
          database = { enabled: true, ok: false, error: err.message };
        }
      }
      sendJson(res, database.enabled && !database.ok ? 503 : 200, {
        ok: !database.enabled || database.ok,
        service: "ddo-attendance-api",
        module: "DDO++ Attendance sidecar",
        store: database.enabled ? "postgres" : "json",
        production: isProduction(),
        database,
        postgres: typeof db.redactedConfig === "function" ? db.redactedConfig() : { configured: db.isEnabled() },
      });
      return;
    }

    if (req.method === "GET" && servePublic(req, res, url)) return;

    if (!requireAuth(req, res)) return;

    if (route === "POST /api/attendance/import") {
      if (!allowImportIp(req, res)) return;
      const raw = await readBody(req);
      let payload;
      try {
        payload = JSON.parse(raw || "{}");
      } catch {
        sendJson(res, 400, { ok: false, error: "Invalid JSON body" });
        return;
      }
      const result = await importAttendance(payload, req.headers["x-location-code"]);
      log("INFO", "Import received", {
        batch_id: result.batch_id,
        location_code: result.location_code,
        status: result.status,
        records_processed: result.records_processed,
        records_inserted: result.records_inserted,
        records_updated: result.records_updated,
        records_failed: result.records_failed,
      });
      sendJson(res, result.ok ? 200 : 400, result);
      return;
    }

    if (route === "GET /api/import-history") {
      const batches = db.isEnabled()
        ? await db.listBatches()
        : Object.values(readStore(files.batches)).sort((a, b) =>
            String(b.completed_at).localeCompare(String(a.completed_at))
          );
      sendJson(res, 200, { ok: true, batches });
      return;
    }

    if (route === "GET /api/attendance") {
      const rows = await filterAttendance(url.searchParams);
      sendJson(res, 200, { ok: true, count: rows.length, records: rows });
      return;
    }

    if (route === "GET /api/dashboard") {
      sendJson(res, 200, { ok: true, dashboard: await dashboard(url.searchParams) });
      return;
    }

    if (route === "GET /api/employees") {
      const locationCode = (url.searchParams.get("location_code") || "").trim().toUpperCase();
      const employees = db.isEnabled()
        ? await db.listEmployees(locationCode)
        : Object.values(readStore(files.employees))
            .filter((row) => !locationCode || locationCode === "ALL" || row.location_code === locationCode)
            .sort((a, b) => a.employee_name.localeCompare(b.employee_name));
      sendJson(res, 200, { ok: true, count: employees.length, employees });
      return;
    }

    if (route === "GET /api/exceptions") {
      sendJson(res, 200, { ok: true, ...(await listExceptions()) });
      return;
    }

    if (route === "GET /api/statuses") {
      sendJson(res, 200, { ok: true, statuses: STATUS_MASTER });
      return;
    }

    if (route === "GET /api/locations") {
      sendJson(res, 200, {
        ok: true,
        locations: [
          { location_code: "ALL", location_name: "All Locations" },
          ...Object.entries(ALLOWED_LOCATIONS).map(([location_code, location_name]) => ({
            location_code,
            location_name,
          })),
        ],
      });
      return;
    }

    sendJson(res, 404, { ok: false, error: "Not found" });
  } catch (err) {
    const statusCode = err.statusCode || 500;
    log("ERROR", err.message);
    sendJson(res, statusCode, { ok: false, error: err.message });
  }
}

function createServer() {
  if (TLS_CERT && TLS_KEY) {
    return https.createServer(
      { cert: fs.readFileSync(TLS_CERT), key: fs.readFileSync(TLS_KEY) },
      handleRequest
    );
  }
  return http.createServer(handleRequest);
}

const useHttps = Boolean(TLS_CERT && TLS_KEY);
const server = createServer();

async function start() {
  if (isProduction() && !db.isEnabled()) {
    throw new Error("NODE_ENV=production requires DATABASE_URL or DB_HOST/DB_DATABASE/DB_USER/DB_PASSWORD.");
  }
  if (!TOKEN) {
    throw new Error("DDO_API_TOKEN is not configured.");
  }
  if (isProduction() && TOKEN === POC_TOKEN) {
    log("WARN", "Using POC token dev-ddo-attendance-token in production. Rotate DDO_API_TOKEN when you can.");
  }

  if (db.isEnabled()) {
    await db.connect();
    await db.ensureSchema(STATUS_MASTER);
    log("INFO", "Attendance store: PostgreSQL (NocoBase/DDO database)", db.redactedConfig());
  } else {
    log("WARN", "DATABASE_URL / DB_HOST not set; using JSON files. Set Postgres env to persist for DDO++.");
  }
  server.listen(PORT, BIND, () => {
    const scheme = useHttps ? "https" : "http";
    log("INFO", `DDO++ attendance API listening on ${scheme}://${BIND}:${PORT}`);
    log("INFO", "POST /api/attendance/import  GET /api/dashboard");
    if (ALLOW_IPS.length) {
      log("INFO", `Import allow-list enabled (${ALLOW_IPS.length} entries)`);
    }
    if (!useHttps) {
      log("WARN", "TLS cert/key not configured; using HTTP. Put Nginx/IIS in front for HTTPS.");
    }
  });
}

start().catch((err) => {
  log("ERROR", `Failed to start attendance API: ${err.message}`);
  process.exit(1);
});
