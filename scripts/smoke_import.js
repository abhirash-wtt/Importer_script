"use strict";

require("../env");
const crypto = require("crypto");
const http = require("http");
const https = require("https");
const { URL } = require("url");

function log(level, message, extra) {
  const line = extra ? `${message} ${JSON.stringify(extra)}` : message;
  console.log(`${new Date().toISOString()} ${level} ${line}`);
}

const endpoint = (process.env.DDO_API_ENDPOINT || "http://127.0.0.1:3000/api/attendance/import").trim();
const token = (process.env.DDO_API_TOKEN || "dev-ddo-attendance-token").trim();
const location = (process.env.LOCATION_CODE || "NOIDA").trim().toUpperCase();

function request(method, urlString, body) {
  const parsed = new URL(urlString);
  const lib = parsed.protocol === "https:" ? https : http;
  const payload = body ? Buffer.from(JSON.stringify(body)) : null;
  const headers = { Accept: "application/json" };
  if (token) headers.Authorization = `Bearer ${token}`;
  if (payload) {
    headers["Content-Type"] = "application/json";
    headers["Content-Length"] = payload.length;
    headers["X-Location-Code"] = location;
  }
  return new Promise((resolve, reject) => {
    const req = lib.request(
      {
        method,
        hostname: parsed.hostname,
        port: parsed.port,
        path: `${parsed.pathname}${parsed.search}`,
        headers,
        timeout: 20000,
      },
      (res) => {
        const chunks = [];
        res.on("data", (chunk) => chunks.push(chunk));
        res.on("end", () => {
          const text = Buffer.concat(chunks).toString("utf8");
          let parsedBody = {};
          try {
            parsedBody = JSON.parse(text || "{}");
          } catch {
            parsedBody = { raw: text };
          }
          resolve({ statusCode: res.statusCode, body: parsedBody });
        });
      }
    );
    req.on("error", reject);
    req.on("timeout", () => {
      req.destroy();
      reject(new Error("Timeout"));
    });
    if (payload) req.write(payload);
    req.end();
  });
}

function originFrom(urlString) {
  const parsed = new URL(urlString);
  return `${parsed.protocol}//${parsed.host}`;
}

function payload(batchId) {
  return {
    location_code: location,
    report_from: "2000-01-01",
    report_to: "2000-01-01",
    source_file: "smoke_import.xls",
    file_hash: "smoke-hash-phase6",
    batch_id: batchId,
    employee_count: 1,
    records: [
      {
        employee_code: "SMOKE_DDO_ATT",
        employee_name: "Smoke Test",
        attendance_date: "2000-01-01",
        status: "P",
        in_time: "09:00",
        out_time: "18:00",
        total_minutes: 540,
      },
    ],
  };
}

async function main() {
  if (endpoint.toLowerCase().includes("/admin/")) {
    throw new Error("DDO_API_ENDPOINT points at /admin/ — use /api/attendance/import");
  }
  const origin = originFrom(endpoint);
  log("INFO", "Phase 6 smoke import", { endpoint, location });

  const health = await request("GET", `${origin}/health`);
  if (health.statusCode !== 200 || !health.body.ok) {
    throw new Error(`/health failed: HTTP ${health.statusCode} ${JSON.stringify(health.body)}`);
  }
  log("INFO", "Health ok", { store: health.body.store, production: health.body.production });

  const noToken = await new Promise((resolve, reject) => {
    const parsed = new URL(endpoint);
    const lib = parsed.protocol === "https:" ? https : http;
    const body = Buffer.from(JSON.stringify(payload("smoke-denied")));
    const req = lib.request(
      {
        method: "POST",
        hostname: parsed.hostname,
        port: parsed.port,
        path: `${parsed.pathname}${parsed.search}`,
        headers: {
          "Content-Type": "application/json",
          "Content-Length": body.length,
        },
        timeout: 15000,
      },
      (res) => {
        resolve(res.statusCode);
        res.resume();
      }
    );
    req.on("error", reject);
    req.write(body);
    req.end();
  });
  if (noToken !== 401) {
    throw new Error(`Expected 401 without token, got ${noToken}`);
  }
  log("INFO", "Auth check passed (401 without token)");

  const first = await request("POST", endpoint, payload(`smoke-${crypto.randomUUID().slice(0, 8)}`));
  if (first.statusCode < 200 || first.statusCode >= 300 || !first.body.ok) {
    throw new Error(`First import failed: HTTP ${first.statusCode} ${JSON.stringify(first.body)}`);
  }
  log("INFO", "First import", {
    inserted: first.body.records_inserted,
    updated: first.body.records_updated,
  });

  const second = await request("POST", endpoint, payload(`smoke-${crypto.randomUUID().slice(0, 8)}`));
  if (second.statusCode < 200 || second.statusCode >= 300 || !second.body.ok) {
    throw new Error(`Second import failed: HTTP ${second.statusCode} ${JSON.stringify(second.body)}`);
  }
  log("INFO", "Second import (idempotency)", {
    inserted: second.body.records_inserted,
    updated: second.body.records_updated,
  });
  if (Number(second.body.records_inserted) !== 0 || Number(second.body.records_updated) < 1) {
    throw new Error("Idempotency failed: second POST should update the same employee+date, not insert a new row.");
  }

  const listed = await request(
    "GET",
    `${origin}/api/attendance?employee_code=SMOKE_DDO_ATT&from=2000-01-01&to=2000-01-01`
  );
  if (listed.statusCode !== 200 || (listed.body.count || 0) !== 1) {
    throw new Error(`Expected 1 smoke row in GET /api/attendance, got ${JSON.stringify(listed.body)}`);
  }
  log("INFO", "Phase 6 passed. One smoke row exists for SMOKE_DDO_ATT on 2000-01-01.");
  log("INFO", "Confirm the same row on https://ddoplus.walkingtree.tech/admin/vb0faior0eg after collections are mapped.");
}

main().catch((err) => {
  log("ERROR", err.message || String(err));
  process.exitCode = 1;
});
