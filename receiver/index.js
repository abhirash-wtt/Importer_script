"use strict";

/**
 * Office importer payload verification.
 * No database, no .env, no schema — use DDO backend auth and config.
 */

const { randomUUID } = require("crypto");
const { STATUS_MASTER, lookupStatus, normalizeStatusKey } = require("./statuses");

const ALLOWED_LOCATIONS = {
  AGRA: "Agra",
  NOIDA: "Noida",
  HYD: "Hyderabad",
};

function isValidDate(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(value || "") && !Number.isNaN(Date.parse(`${value}T00:00:00Z`));
}

function minutesToHours(minutes) {
  if (minutes == null || Number.isNaN(Number(minutes))) return "";
  const total = Number(minutes);
  const hh = Math.floor(total / 60);
  const mm = String(total % 60).padStart(2, "0");
  return `${hh}:${mm}`;
}

function httpError(statusCode, message) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
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
      identity_status:
        raw.identity_status ||
        (String(employeeCode).startsWith("UNMAPPED:") ? "UNMAPPED_NAME" : "CODED"),
      punch_expected: Boolean(raw.punch_expected),
    },
  };
}

/**
 * Verify the JSON the office scheduler POSTs.
 * Does not write to a database.
 *
 * @param {object} payload
 * @param {string} [locationHeader] X-Location-Code
 */
function verifyImport(payload, locationHeader) {
  const body = payload && typeof payload === "object" ? payload : {};
  const locationCode = String(body.location_code || "").trim().toUpperCase();

  if (!ALLOWED_LOCATIONS[locationCode]) {
    throw httpError(400, "Invalid location_code. Use AGRA, NOIDA, or HYD.");
  }
  if (locationHeader && String(locationHeader).toUpperCase() !== locationCode) {
    throw httpError(400, "X-Location-Code does not match location_code");
  }
  if (!Array.isArray(body.records)) {
    throw httpError(400, "records must be an array");
  }

  const batchId = String(body.batch_id || randomUUID());
  const sourceFile = String(body.source_file || "");
  const importedAt = new Date().toISOString();
  const failures = Array.isArray(body.failures) ? [...body.failures] : [];
  const unknownStatuses = [];
  const dataQualityIssues = Array.isArray(body.data_quality_issues)
    ? [...body.data_quality_issues]
    : [];
  const validRecords = [];

  body.records.forEach((raw, index) => {
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

  const failed = failures.length;
  let status = "SUCCESS";
  if (failed && !validRecords.length) status = "FAILED";
  else if (failed || dataQualityIssues.length || unknownStatuses.length) status = "PARTIAL";

  return {
    ok: status !== "FAILED",
    batch_id: batchId,
    status,
    location_code: locationCode,
    location_name: ALLOWED_LOCATIONS[locationCode],
    source_file: sourceFile,
    file_hash: String(body.file_hash || ""),
    report_from: body.report_from || null,
    report_to: body.report_to || null,
    records_processed: body.records.length,
    records_inserted: validRecords.length,
    records_updated: 0,
    records_failed: failed,
    unknown_status_count: unknownStatuses.length,
    data_quality_issue_count: dataQualityIssues.length,
    failures,
    unknown_statuses: unknownStatuses,
    data_quality_issues: dataQualityIssues,
    records: validRecords,
  };
}

function toClientResponse(result) {
  const {
    records,
    unknown_statuses,
    data_quality_issues,
    location_name,
    source_file,
    file_hash,
    report_from,
    report_to,
    ...client
  } = result;
  return client;
}

/**
 * Express router. Mount behind DDO's existing auth.
 *
 *   const { createAttendanceImportRouter } = require("./receiver");
 *   app.use("/api/attendance", existingAuth, createAttendanceImportRouter());
 *
 * Optional persist(verified) — save with DDO's own DB/models. If omitted, this
 * only verifies and returns the importer's expected JSON.
 */
function createAttendanceImportRouter(options = {}) {
  const express = require("express");
  const router = express.Router();
  const persist = options.persist;

  router.post("/import", async (req, res, next) => {
    try {
      const verified = verifyImport(req.body, req.get("X-Location-Code"));
      if (typeof persist === "function") {
        const saved = await persist(verified, req);
        if (saved && typeof saved === "object") {
          Object.assign(verified, saved);
        }
      }
      res.status(verified.ok ? 200 : 400).json(toClientResponse(verified));
    } catch (err) {
      if (err.statusCode) {
        res.status(err.statusCode).json({ ok: false, error: err.message });
        return;
      }
      next(err);
    }
  });

  return router;
}

module.exports = {
  ALLOWED_LOCATIONS,
  STATUS_MASTER,
  normalizeStatusKey,
  lookupStatus,
  verifyImport,
  toClientResponse,
  createAttendanceImportRouter,
};
