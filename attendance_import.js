"use strict";

/**
 * DDO++ local attendance import agent (POC).
 *
 * Same package at every site. Location is configuration, not Excel content:
 *   LOCATION_CODE=AGRA | NOIDA | HYD
 *
 *   node attendance_import.js
 *   node attendance_import.js --location NOIDA "C:\\path\\report.xls"
 */

const crypto = require("crypto");
const fs = require("fs");
const http = require("http");
const https = require("https");
const path = require("path");
const { URL } = require("url");
const XLSX = require("xlsx");

const BASE_DIR = __dirname;
loadEnvFile(path.join(BASE_DIR, ".env"));

const ALLOWED_LOCATIONS = new Set(["AGRA", "NOIDA", "HYD"]);
const WEEKDAY_CODES = { M: 0, T: 1, W: 2, TH: 3, F: 4, ST: 5, S: 6 };
const ROW_LABELS = {
  STATUS: "status",
  INTIME: "intime",
  OUTTIME: "outtime",
  TOTAL: "total",
};
const DEPARTMENT_NAMES = new Set(
  ["Delivery", "TA", "Sales & BD", "Marketing", "HR", "F&A", "Management", "IT"].map((n) =>
    n.toLowerCase()
  )
);

const CONFIG = {
  location: (process.env.LOCATION_CODE || "").trim().toUpperCase(),
  apiEndpoint: (process.env.DDO_API_ENDPOINT || "http://127.0.0.1:3000/api/attendance/import").trim(),
  apiToken: (process.env.DDO_API_TOKEN || "dev-ddo-attendance-token").trim(),
  apiTimeoutMs: Number(process.env.DDO_API_TIMEOUT_SECONDS || 30) * 1000,
  apiMaxRetries: Number(process.env.DDO_API_MAX_RETRIES || 3),
  inputDir: path.join(BASE_DIR, "input"),
  processedDir: path.join(BASE_DIR, "processed"),
  failedDir: path.join(BASE_DIR, "failed"),
  logsDir: path.join(BASE_DIR, "logs"),
  outputDir: path.join(BASE_DIR, "output"),
};

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

function log(level, message, extra) {
  const line = extra ? `${message} ${JSON.stringify(extra)}` : message;
  const text = `${new Date().toISOString()} ${level} ${line}`;
  console.log(text);
  fs.mkdirSync(CONFIG.logsDir, { recursive: true });
  fs.appendFileSync(path.join(CONFIG.logsDir, "importer.log"), `${text}\n`, "utf8");
}

function cellText(value) {
  if (value == null || value === "") return "";
  if (typeof value === "number" && Number.isInteger(value)) return String(value);
  return String(value).replace(/\s+/g, " ").trim();
}

function parseDayHeader(header) {
  const match = String(header || "")
    .trim()
    .match(/^(\d{1,2})\s*([A-Za-z]{1,2})$/);
  if (!match) return null;
  const weekday = match[2].toUpperCase();
  if (!(weekday in WEEKDAY_CODES)) return null;
  return { day: Number(match[1]), weekday, header: String(header).trim() };
}

function inferYearMonth(fileName, reportText) {
  const months = {
    january: 1, february: 2, march: 3, april: 4, may: 5, june: 6,
    july: 7, august: 8, september: 9, october: 10, november: 11, december: 12,
    jan: 1, feb: 2, mar: 3, apr: 4, jun: 6, jul: 7, aug: 8,
    sep: 9, oct: 10, nov: 11, dec: 12,
  };
  const haystack = `${fileName} ${reportText || ""}`;
  const yearMatch = haystack.match(/\b(20\d{2})\b/);
  const today = new Date();
  let year = yearMatch ? Number(yearMatch[1]) : today.getFullYear();
  let month = today.getMonth() + 1;
  for (const [name, number] of Object.entries(months)) {
    if (new RegExp(`\\b${name}\\b`, "i").test(haystack)) {
      month = number;
      break;
    }
  }
  if (!yearMatch && month > today.getMonth() + 1) year -= 1;
  return { year, month };
}

function mapHeadersToDates(dayHeaders, year, month) {
  const parsed = dayHeaders.map(parseDayHeader).filter(Boolean);
  if (!parsed.length) return [];
  let currentYear = year;
  let currentMonth = parsed[0].day > 20 ? month - 1 : month;
  if (currentMonth < 1) {
    currentMonth = 12;
    currentYear -= 1;
  }
  const dates = [];
  let lastDay = parsed[0].day;
  for (const item of parsed) {
    if (dates.length && item.day < lastDay) {
      currentMonth += 1;
      if (currentMonth > 12) {
        currentMonth = 1;
        currentYear += 1;
      }
    }
    dates.push(new Date(Date.UTC(currentYear, currentMonth - 1, item.day)));
    lastDay = item.day;
  }
  return dates;
}

function isoDate(d) {
  return d.toISOString().slice(0, 10);
}

function normalizeStatus(raw) {
  return cellText(raw);
}

function parseTime(value) {
  const text = cellText(value);
  if (!text) return null;
  const match = text.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) return null;
  const hh = String(Number(match[1])).padStart(2, "0");
  return `${hh}:${match[2]}${match[3] ? `:${match[3]}` : ""}`;
}

function durationToMinutes(value) {
  const text = cellText(value);
  if (!text) return null;
  const match = text.match(/^(\d{1,3}):(\d{2})(?::(\d{2}))?$/);
  if (!match) return null;
  return Number(match[1]) * 60 + Number(match[2]);
}

function rowLabel(row) {
  for (let i = 0; i < Math.min(row.length, 6); i += 1) {
    const value = cellText(row[i]).toLowerCase().replace(/[\s._-]/g, "");
    if (value === "status") return ROW_LABELS.STATUS;
    if (value === "intime") return ROW_LABELS.INTIME;
    if (value === "outtime") return ROW_LABELS.OUTTIME;
    if (value === "total") return ROW_LABELS.TOTAL;
  }
  return null;
}

function looksLikeDepartment(row, headerOffset) {
  const name = cellText(row[headerOffset]);
  if (!name) return false;
  if (DEPARTMENT_NAMES.has(name.toLowerCase())) return true;
  return /^department\b/i.test(name) && !parseDayHeader(row[headerOffset + 1]);
}

function selectSheet(workbook) {
  let best = { name: workbook.SheetNames[0], rows: [] };
  let bestScore = -1;
  for (const name of workbook.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: "", raw: false });
    const joined = rows.slice(0, 8).flat().map(cellText).join(" ").toLowerCase();
    let score = 0;
    if (name.toLowerCase().includes("basicworkduration")) score += 50;
    if (joined.includes("emp. code") || joined.includes("emp code")) score += 20;
    if (rows.some((row) => row.some((cell) => parseDayHeader(cell)))) score += 10;
    if (rows.some((row) => rowLabel(row))) score += 15;
    score += Math.min(rows.length, 200) / 10;
    if (score > bestScore) {
      bestScore = score;
      best = { name, rows };
    }
  }
  return best;
}

function extractReportPeriod(fileName, rows) {
  const headerRowIndex = rows.findIndex((row) => row.some((cell) => parseDayHeader(cell)));
  const headers = headerRowIndex >= 0 ? rows[headerRowIndex] : [];
  const dayHeaders = headers.filter((cell) => parseDayHeader(cell));
  const reportText = rows.slice(0, 8).flat().map(cellText).join(" ");
  const { year, month } = inferYearMonth(fileName, reportText);
  const dates = mapHeadersToDates(dayHeaders, year, month);
  const headerOffset = Math.max(
    headers.findIndex((cell) => parseDayHeader(cell)) - 1,
    0
  );
  return {
    label: dates.length ? `${dates[0].toLocaleString("en-GB", { month: "long", year: "numeric", timeZone: "UTC" })}` : "",
    year,
    month,
    start: dates[0] ? isoDate(dates[0]) : null,
    end: dates.length ? isoDate(dates[dates.length - 1]) : null,
    dates,
    dayHeaders,
    headerOffset,
    headerRowIndex: headerRowIndex < 0 ? 0 : headerRowIndex,
  };
}

function daySlice(row, headerOffset, datesLength) {
  return row.slice(headerOffset + 1, headerOffset + 1 + datesLength).map(cellText);
}

function extractPunchRepBlocks(rows, period) {
  const blocks = [];
  let i = period.headerRowIndex + 1;
  while (i < rows.length) {
    const row = rows[i];
    if (!row || !row.some((cell) => cellText(cell))) {
      i += 1;
      continue;
    }
    if (looksLikeDepartment(row, period.headerOffset)) {
      i += 1;
      continue;
    }
    const label = rowLabel(row);
    if (label) {
      i += 1;
      continue;
    }

    const code = cellText(row[0]);
    const name = cellText(row[period.headerOffset]) || cellText(row[1]);
    const grouped = { status: daySlice(row, period.headerOffset, period.dates.length) };
    let j = i + 1;
    while (j < rows.length && rowLabel(rows[j])) {
      grouped[rowLabel(rows[j])] = daySlice(rows[j], period.headerOffset, period.dates.length);
      j += 1;
    }
    if (grouped.status || grouped.intime || grouped.outtime || grouped.total) {
      blocks.push({ code, name, grouped });
      i = j;
      continue;
    }
    i += 1;
  }
  return blocks;
}

function extractGridBlocks(rows, period) {
  const blocks = [];
  for (const row of rows.slice(period.headerRowIndex + 1)) {
    if (!row || !row.some((cell) => cellText(cell))) continue;
    if (looksLikeDepartment(row, period.headerOffset)) continue;
    if (rowLabel(row)) continue;
    const first = cellText(row[0]);
    const second = cellText(row[1]);
    const hasCodeHeader = rows[period.headerRowIndex]
      .map(cellText)
      .some((h) => /emp\.?\s*code/i.test(h));
    let code = "";
    let name = "";
    if (hasCodeHeader) {
      code = first;
      name = second || cellText(row[period.headerOffset]);
    } else {
      name = first || cellText(row[period.headerOffset]);
      code = /^\d{3,}$/.test(first) ? first : "";
    }
    if (!code && !name) continue;
    blocks.push({
      code,
      name,
      grouped: { status: daySlice(row, period.headerOffset, period.dates.length) },
    });
  }
  return blocks;
}

function extractEmployeeBlocks(rows, period) {
  const hasPunchRows = rows.some((row) => rowLabel(row));
  return hasPunchRows ? extractPunchRepBlocks(rows, period) : extractGridBlocks(rows, period);
}

function fileHash(filePath) {
  return crypto.createHash("sha256").update(fs.readFileSync(filePath)).digest("hex");
}

function inferLocation(fileName, explicit) {
  if (explicit) return explicit.toUpperCase();
  const lower = fileName.toLowerCase();
  if (lower.includes("agra")) return "AGRA";
  if (lower.includes("noida")) return "NOIDA";
  if (lower.includes("hyd") || lower.includes("hyderabad")) return "HYD";
  return (CONFIG.location || "").toUpperCase();
}

function validateLocation(location) {
  if (!ALLOWED_LOCATIONS.has(location)) {
    throw new Error(`Invalid location '${location || ""}'. Use AGRA, NOIDA, or HYD.`);
  }
  return location;
}

function transformDay(dayDate, grouped, employee, location, sourceFile) {
  const status = normalizeStatus(grouped.status);
  return {
    employee_code: employee.code,
    employee_name: employee.name,
    location_code: location,
    attendance_date: isoDate(dayDate),
    status,
    in_time: parseTime(grouped.intime),
    out_time: parseTime(grouped.outtime),
    total_minutes: durationToMinutes(grouped.total),
    source_file: sourceFile,
    identity_status: employee.identityStatus,
    punch_expected: Boolean(grouped.intime || grouped.outtime || grouped.total),
  };
}

function moveFile(filePath, targetDir) {
  fs.mkdirSync(targetDir, { recursive: true });
  const source = path.basename(filePath);
  let destination = path.join(targetDir, source);
  if (fs.existsSync(destination)) {
    const stamp = new Date().toISOString().replace(/[-:.TZ]/g, "").slice(0, 15);
    destination = path.join(targetDir, `${path.parse(source).name}_${stamp}${path.parse(source).ext}`);
  }
  fs.renameSync(filePath, destination);
  log("INFO", `Moved ${source} -> ${destination}`);
  return destination;
}

function isAllowedEndpoint(endpoint) {
  const parsed = new URL(endpoint);
  const host = (parsed.hostname || "").toLowerCase();
  if (parsed.protocol === "https:") return true;
  return parsed.protocol === "http:" && ["127.0.0.1", "localhost", "::1"].includes(host);
}

function postJson(endpoint, token, locationCode, payload, timeoutMs) {
  const body = Buffer.from(JSON.stringify(payload));
  const parsed = new URL(endpoint);
  const lib = parsed.protocol === "https:" ? https : http;
  const options = {
    method: "POST",
    hostname: parsed.hostname,
    port: parsed.port,
    path: `${parsed.pathname}${parsed.search}`,
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
      Authorization: `Bearer ${token}`,
      "X-Location-Code": locationCode,
      "Content-Length": body.length,
    },
    timeout: timeoutMs,
  };
  return new Promise((resolve, reject) => {
    const req = lib.request(options, (res) => {
      const chunks = [];
      res.on("data", (chunk) => chunks.push(chunk));
      res.on("end", () => {
        const text = Buffer.concat(chunks).toString("utf8");
        resolve({ statusCode: res.statusCode, text });
      });
    });
    req.on("error", reject);
    req.on("timeout", () => {
      req.destroy();
      reject(new Error("Timeout"));
    });
    req.write(body);
    req.end();
  });
}

async function sendToApi(batch) {
  if (!isAllowedEndpoint(CONFIG.apiEndpoint)) {
    throw new Error(`DDO++ API endpoint must use HTTPS (http only for localhost): ${CONFIG.apiEndpoint}`);
  }
  if (!CONFIG.apiToken) throw new Error("DDO_API_TOKEN is not configured");

  const payload = {
    location_code: batch.location,
    report_from: batch.period.start,
    report_to: batch.period.end,
    source_file: batch.fileName,
    file_hash: batch.fileHash,
    batch_id: batch.id,
    employee_count: batch.employeeCount,
    records: batch.records,
    failures: batch.failures,
    data_quality_issues: batch.dataQualityIssues,
  };

  let lastError = "Unknown API error";
  for (let attempt = 1; attempt <= CONFIG.apiMaxRetries; attempt += 1) {
    try {
      const { statusCode, text } = await postJson(
        CONFIG.apiEndpoint,
        CONFIG.apiToken,
        batch.location,
        payload,
        CONFIG.apiTimeoutMs
      );
      let parsed = {};
      try {
        parsed = JSON.parse(text || "{}");
      } catch {
        parsed = { raw: text };
      }
      if (statusCode >= 200 && statusCode < 300) {
        batch.apiResult = { ok: true, status_code: statusCode, attempt, response: parsed };
        log("INFO", `DDO++ API accepted batch ${batch.id}`, parsed);
        return parsed;
      }
      lastError = `HTTP ${statusCode}: ${text.slice(0, 400)}`;
      log("ERROR", lastError);
      if (statusCode < 500 && statusCode !== 429) break;
    } catch (err) {
      lastError = err.message || String(err);
      log("ERROR", `API attempt ${attempt}/${CONFIG.apiMaxRetries} failed: ${lastError}`);
    }
    if (attempt < CONFIG.apiMaxRetries) {
      const delay = Math.min(2 ** (attempt - 1), 8) * 1000;
      await new Promise((resolve) => setTimeout(resolve, delay));
    }
  }
  throw new Error(`Failed to send attendance JSON to DDO++ API: ${lastError}`);
}

function writeBatchOutput(batch) {
  fs.mkdirSync(CONFIG.outputDir, { recursive: true });
  const summary = {
    batch_id: batch.id,
    file_name: batch.fileName,
    location: batch.location,
    status: batch.status,
    error: batch.error,
    period: { label: batch.period.label, start: batch.period.start, end: batch.period.end },
    employee_count: batch.employeeCount,
    attendance_upserted: batch.records.length,
    employees_failed: batch.failures.length,
    data_quality_issues: batch.dataQualityIssues.length,
    failures: batch.failures,
    issues: batch.dataQualityIssues,
    api_result: batch.apiResult,
  };
  fs.writeFileSync(
    path.join(CONFIG.outputDir, `batch_${batch.id}_summary.json`),
    JSON.stringify(summary, null, 2),
    "utf8"
  );
  fs.writeFileSync(
    path.join(CONFIG.outputDir, `batch_${batch.id}_attendance.json`),
    JSON.stringify(batch.records, null, 2),
    "utf8"
  );
}

function processFile(filePath, locationArg) {
  const fileName = path.basename(filePath);
  const location = validateLocation(inferLocation(fileName, locationArg));
  const batch = {
    id: crypto.randomUUID().slice(0, 8),
    fileName,
    fileHash: fileHash(filePath),
    location,
    status: "RUNNING",
    error: null,
    period: {},
    records: [],
    failures: [],
    dataQualityIssues: [],
    employeeCount: 0,
    apiResult: {},
  };
  log("INFO", `Created import batch ${batch.id}`, { fileName, location });

  try {
    const workbook = XLSX.readFile(filePath, { raw: false });
    const sheet = selectSheet(workbook);
    const period = extractReportPeriod(fileName, sheet.rows);
    batch.period = period;
    log("INFO", `Report period: ${period.start} to ${period.end} (${sheet.name})`);

    const blocks = extractEmployeeBlocks(sheet.rows, period);
    batch.employeeCount = blocks.length;
    log("INFO", `Found ${blocks.length} employee blocks in ${sheet.name}`);

    for (const block of blocks) {
      if (!block.code && !block.name) {
        batch.failures.push({
          employee_code: "",
          employee_name: "",
          reason: "Unknown employee",
        });
        continue;
      }
      const identityStatus = block.code ? "CODED" : "UNMAPPED_NAME";
      const code =
        block.code ||
        `UNMAPPED:${block.name.toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "")}`;
      if (!block.code) {
        batch.dataQualityIssues.push({
          employee_code: code,
          employee_name: block.name,
          date: null,
          reason: "Missing employee_code; name-based identity used",
        });
        log("WARN", `Missing employee_code for ${block.name}; using ${code}`);
      }
      const employee = { code, name: block.name, identityStatus };
      for (let idx = 0; idx < period.dates.length; idx += 1) {
        const grouped = {
          status: (block.grouped.status || [])[idx] || "",
          intime: (block.grouped.intime || [])[idx] || "",
          outtime: (block.grouped.outtime || [])[idx] || "",
          total: (block.grouped.total || [])[idx] || "",
        };
        if (!grouped.status) {
          batch.dataQualityIssues.push({
            employee_code: employee.code,
            employee_name: employee.name,
            date: isoDate(period.dates[idx]),
            reason: "Missing attendance status",
          });
          continue;
        }
        const record = transformDay(period.dates[idx], grouped, employee, location, fileName);
        if (
          record.punch_expected &&
          (record.status.toUpperCase() === "P" || record.status.toUpperCase() === "WFH") &&
          (!record.in_time || !record.out_time)
        ) {
          batch.dataQualityIssues.push({
            employee_code: employee.code,
            employee_name: employee.name,
            date: record.attendance_date,
            reason: !record.in_time && !record.out_time ? "Missing InTime and OutTime" : !record.in_time ? "Missing InTime" : "Missing OutTime",
          });
        }
        batch.records.push(record);
      }
    }

    batch.status = batch.failures.length && !batch.records.length ? "FAILED" : "SUCCESS";
    writeBatchOutput(batch);
    return sendToApi(batch).then(() => {
      writeBatchOutput(batch);
      moveFile(filePath, CONFIG.processedDir);
      return batch;
    });
  } catch (err) {
    batch.status = "FAILED";
    batch.error = err.message || String(err);
    log("ERROR", `Import failed for ${fileName}: ${batch.error}`);
    writeBatchOutput(batch);
    try {
      moveFile(filePath, CONFIG.failedDir);
    } catch (moveErr) {
      log("ERROR", `Could not move failed file: ${moveErr.message}`);
    }
    return Promise.reject(err);
  }
}

function parseArgs(argv) {
  const args = { location: "", files: [] };
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i] === "--location" || argv[i] === "-l") {
      args.location = argv[i + 1] || "";
      i += 1;
    } else {
      args.files.push(argv[i]);
    }
  }
  return args;
}

async function main() {
  fs.mkdirSync(CONFIG.inputDir, { recursive: true });
  fs.mkdirSync(CONFIG.processedDir, { recursive: true });
  fs.mkdirSync(CONFIG.failedDir, { recursive: true });
  fs.mkdirSync(CONFIG.logsDir, { recursive: true });

  const args = parseArgs(process.argv.slice(2));
  let files = args.files.filter((file) => fs.existsSync(file));
  if (!files.length) {
    files = fs
      .readdirSync(CONFIG.inputDir)
      .filter((name) => /\.xls[x]?$/i.test(name))
      .map((name) => path.join(CONFIG.inputDir, name));
  }
  if (!files.length) {
    log("ERROR", "No input file was provided and input/ contains no .xls/.xlsx reports.");
    process.exitCode = 1;
    return;
  }

  for (const file of files) {
    await processFile(file, args.location);
  }
}

if (require.main === module) {
  main().catch((err) => {
    log("ERROR", err.message || String(err));
    process.exitCode = 1;
  });
}

module.exports = { processFile, extractEmployeeBlocks, extractReportPeriod };
