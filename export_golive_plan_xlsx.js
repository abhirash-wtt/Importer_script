"use strict";

const path = require("path");
const ExcelJS = require("exceljs");

const OUT_FILE = path.join(__dirname, "Attendance_GoLive_Plan_Sep2026.xlsx");

const HEADER_FILL = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FF1F4E79" },
};
const HEADER_FONT = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
const BODY_FONT = { name: "Calibri", size: 11 };
const TITLE_FONT = { name: "Calibri", size: 18, bold: true, color: { argb: "FF1F4E79" } };
const SUBTITLE_FONT = { name: "Calibri", size: 12, color: { argb: "FF333333" } };
const WRAP = { wrapText: true, vertical: "middle" };
const THIN = {
  top: { style: "thin", color: { argb: "FFB0B0B0" } },
  left: { style: "thin", color: { argb: "FFB0B0B0" } },
  bottom: { style: "thin", color: { argb: "FFB0B0B0" } },
  right: { style: "thin", color: { argb: "FFB0B0B0" } },
};
const ALT_FILL = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFF2F2F2" },
};

const STATUS_FILLS = {
  "Not Started": { type: "pattern", pattern: "solid", fgColor: { argb: "FFFCE4D6" } },
  Open: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFCE4D6" } },
  Pending: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFCE4D6" } },
  High: { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8CBAD" } },
  Medium: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } },
  Low: { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2EFDA" } },
  Critical: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } },
};

function applyHeaderRow(sheet, rowNumber, columnCount) {
  const row = sheet.getRow(rowNumber);
  row.height = 22;
  for (let col = 1; col <= columnCount; col += 1) {
    const cell = row.getCell(col);
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.alignment = { ...WRAP, horizontal: "center" };
    cell.border = THIN;
  }
}

function applyBody(sheet, startRow, endRow, columnCount) {
  for (let r = startRow; r <= endRow; r += 1) {
    const row = sheet.getRow(r);
    row.alignment = WRAP;
    row.font = BODY_FONT;
    for (let c = 1; c <= columnCount; c += 1) {
      const cell = row.getCell(c);
      cell.border = THIN;
      if ((r - startRow) % 2 === 1) cell.fill = ALT_FILL;
    }
  }
}

function styleSheet(sheet, headerRow, headers, widths) {
  sheet.views = [{ state: "frozen", ySplit: headerRow, xSplit: 0 }];
  headers.forEach((header, i) => {
    sheet.getColumn(i + 1).width = widths[i];
  });
  applyHeaderRow(sheet, headerRow, headers.length);
  const last = sheet.lastRow ? sheet.lastRow.number : headerRow;
  if (last > headerRow) applyBody(sheet, headerRow + 1, last, headers.length);
  sheet.autoFilter = {
    from: { row: headerRow, column: 1 },
    to: { row: last, column: headers.length },
  };
}

function addTableSheet(wb, name, title, subtitle, headers, rows, widths) {
  const sheet = wb.addWorksheet(name, { views: [{ showGridLines: false }] });
  sheet.mergeCells(1, 1, 1, headers.length);
  sheet.getCell("A1").value = title;
  sheet.getCell("A1").font = TITLE_FONT;
  sheet.getCell("A1").alignment = { vertical: "middle" };
  sheet.getRow(1).height = 28;

  sheet.mergeCells(2, 1, 2, headers.length);
  sheet.getCell("A2").value = subtitle;
  sheet.getCell("A2").font = SUBTITLE_FONT;
  sheet.getRow(2).height = 20;

  sheet.addRow(headers);
  rows.forEach((row) => sheet.addRow(row));
  styleSheet(sheet, 3, headers, widths);

  const statusCol = headers.findIndex((h) =>
    ["Status", "Priority", "Likelihood", "Impact"].includes(h)
  );
  if (statusCol >= 0) {
    for (let r = 4; r <= sheet.lastRow.number; r += 1) {
      const cell = sheet.getRow(r).getCell(statusCol + 1);
      const fill = STATUS_FILLS[String(cell.value || "")];
      if (fill) cell.fill = fill;
    }
  }
  return sheet;
}

function coverSheet(wb) {
  const sheet = wb.addWorksheet("01_Cover_Summary", {
    views: [{ showGridLines: false }],
  });
  sheet.getColumn(1).width = 28;
  sheet.getColumn(2).width = 72;
  sheet.getColumn(3).width = 28;
  sheet.getColumn(4).width = 28;

  sheet.mergeCells("A1:D1");
  sheet.getCell("A1").value = "DDO++ Attendance — Server Deploy & Local Scheduler";
  sheet.getCell("A1").font = TITLE_FONT;
  sheet.getRow(1).height = 32;

  sheet.mergeCells("A2:D2");
  sheet.getCell("A2").value =
    "Project plan · 15 working days · Thu 10 Sep 2026 to Wed 30 Sep 2026 · Prepared 9 Sep 2026";
  sheet.getCell("A2").font = SUBTITLE_FONT;

  const metaHeaders = ["Field", "Value"];
  const meta = [
    ["Document Name", "Attendance Go-Live Plan"],
    ["Programme", "DDO++ Attendance Module"],
    ["Plan Start Date", "10-Sep-2026"],
    ["Plan Finish Date", "30-Sep-2026"],
    ["Target Go-Live Date", "30-Sep-2026"],
    ["Hypercare Window", "30-Sep-2026 to 02-Oct-2026"],
    ["Working Days", 15],
    ["Sites in Scope", "Noida, Agra, Hyderabad"],
    ["Primary Delivery Model", "1 developer FTE + part-time Infra and HR"],
    ["Central Component", "attendance_api.js on shared server (HTTPS + PostgreSQL)"],
    ["Local Component", "attendance_import.js via Windows Task Scheduler"],
    ["Out of Scope", "Auto-export from biometric machine; NocoBase UI replacement"],
    ["Document Owner", "Project Manager"],
    ["Status", "Baseline for Day-1 kickoff"],
  ];

  sheet.getCell("A4").value = "Document Control";
  sheet.getCell("A4").font = { name: "Calibri", size: 14, bold: true, color: { argb: "FF1F4E79" } };

  sheet.addRow([]);
  const headerRow = sheet.addRow(metaHeaders);
  headerRow.eachCell((cell) => {
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.border = THIN;
  });
  meta.forEach((pair, i) => {
    const row = sheet.addRow(pair);
    row.font = BODY_FONT;
    row.alignment = WRAP;
    row.eachCell((cell) => {
      cell.border = THIN;
      if (i % 2 === 1) cell.fill = ALT_FILL;
    });
  });

  const kpiStart = sheet.lastRow.number + 2;
  sheet.getCell(`A${kpiStart}`).value = "Plan Snapshot";
  sheet.getCell(`A${kpiStart}`).font = {
    name: "Calibri",
    size: 14,
    bold: true,
    color: { argb: "FF1F4E79" },
  };

  const kpiHeaders = [
    "Workstream",
    "Objective",
    "Window",
    "Success Measure",
  ];
  const kpiRows = [
    [
      "A — Central Server",
      "Run attendance API as a managed HTTPS service writing to NocoBase PostgreSQL",
      "10–16 Sep 2026",
      "/health and import succeed from an office PC over HTTPS",
    ],
    [
      "B — Local Scheduler",
      "Office inbox drop auto-imports every 10 minutes without a manual command",
      "17–25 Sep 2026",
      "Noida real file appears on dashboard within 10 minutes of drop",
    ],
    [
      "C — Rollout & UAT",
      "Agra + Hyderabad agents live; scheduled import matches last manual file",
      "24–30 Sep 2026",
      "Go / No-Go signed on 29 Sep; cutover 30 Sep",
    ],
  ];
  const kh = sheet.addRow([]);
  void kh;
  const kh2 = sheet.addRow(kpiHeaders);
  kh2.eachCell((cell) => {
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.border = THIN;
    cell.alignment = { wrapText: true, horizontal: "center", vertical: "middle" };
  });
  kpiRows.forEach((row, i) => {
    const r = sheet.addRow(row);
    r.font = BODY_FONT;
    r.alignment = WRAP;
    r.height = 36;
    r.eachCell((cell) => {
      cell.border = THIN;
      if (i % 2 === 1) cell.fill = ALT_FILL;
    });
  });

  const decStart = sheet.lastRow.number + 2;
  sheet.getCell(`A${decStart}`).value = "Day-1 Decisions (must lock 10 Sep)";
  sheet.getCell(`A${decStart}`).font = {
    name: "Calibri",
    size: 14,
    bold: true,
    color: { argb: "FF1F4E79" },
  };
  const decHeaders = ["Decision ID", "Topic", "Recommended Option", "Rejected Option", "Rationale"];
  const decRows = [
    [
      "DEC-01",
      "Office importer runtime",
      "Node attendance_import.js",
      "Python importer_script.py as production agent",
      "Node agent already has inbox / processed / failed / logs; one runtime with the API",
    ],
    [
      "DEC-02",
      "Local trigger",
      "Windows Task Scheduler every 10 minutes, 08:00–20:00 IST, Mon–Sat",
      "Long-running file-watcher Windows service",
      "Survives reboot; office IT can support; no extra service to babysit",
    ],
    [
      "DEC-03",
      "API hosting",
      "Sidecar on the same host as NocoBase PostgreSQL",
      "Separate database or JSON-file store",
      "Dockerfile already exists; avoids a second source of truth",
    ],
  ];
  sheet.addRow([]);
  const dh = sheet.addRow(decHeaders);
  dh.eachCell((cell) => {
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.border = THIN;
    cell.alignment = { wrapText: true, horizontal: "center", vertical: "middle" };
  });
  decRows.forEach((row, i) => {
    const r = sheet.addRow(row);
    r.font = BODY_FONT;
    r.alignment = WRAP;
    r.height = 42;
    r.eachCell((cell) => {
      cell.border = THIN;
      if (i % 2 === 1) cell.fill = ALT_FILL;
    });
  });

  sheet.getColumn(1).width = 26;
  sheet.getColumn(2).width = 28;
  sheet.getColumn(3).width = 48;
  sheet.getColumn(4).width = 42;
  sheet.getColumn(5).width = 52;
}

const MASTER_HEADERS = [
  "Task ID",
  "Workstream",
  "Phase",
  "Task Name",
  "Description",
  "Owner Role",
  "Support Role",
  "Start Date",
  "Finish Date",
  "Duration (Hours)",
  "Predecessor Task ID",
  "Priority",
  "Status",
  "Site / Location",
  "Acceptance Criteria",
  "Notes",
];

const MASTER_WIDTHS = [12, 18, 22, 42, 48, 16, 16, 14, 14, 16, 20, 12, 14, 18, 48, 36];

const MASTER_ROWS = [
  ["A0.1", "A — Central Server", "Kickoff", "Kickoff and freeze delivery decisions", "Confirm server OS/host, Postgres access, three office PCs, who drops Excel, and import cadence.", "Project Manager", "Dev, Infra, HR", "10-Sep-2026", "10-Sep-2026", 8, "—", "Critical", "Not Started", "All sites", "Host, token owner, folder path, and 10-minute cadence signed", "Day-1 gate"],
  ["A1.1", "A — Central Server", "API Hardening", "Bind API beyond localhost and place reverse proxy", "Listen on 0.0.0.0 or Unix socket; Nginx/IIS in front of attendance_api.js.", "Developer", "Infra", "11-Sep-2026", "11-Sep-2026", 3, "A0.1", "Critical", "Not Started", "Central server", "curl from another machine hits /health", "Currently bound to 127.0.0.1:3000"],
  ["A1.2", "A — Central Server", "API Hardening", "Enable TLS and HTTP to HTTPS redirect", "Install certificates; keep HTTP only on localhost.", "Infra", "Developer", "15-Sep-2026", "15-Sep-2026", 3, "A1.1", "Critical", "Not Started", "Central server", "Importer POST succeeds on https://<host>/api/attendance/import", "Blocks office scheduler"],
  ["A1.3", "A — Central Server", "API Hardening", "Rotate production API token", "Replace POC token; store DDO_API_TOKEN in server env/secrets only.", "Developer", "Infra", "11-Sep-2026", "14-Sep-2026", 1, "A1.1", "High", "Not Started", "Central server", "Old token returns 401; new token is not in source control", "POC value is dev-ddo-attendance-token"],
  ["A1.4", "A — Central Server", "API Hardening", "Require PostgreSQL in production", "Refuse JSON-file fallback when NODE_ENV=production; require DATABASE_URL.", "Developer", "Infra", "11-Sep-2026", "14-Sep-2026", 2, "A0.1", "High", "Not Started", "Central server", "Missing DB env causes non-zero exit", "Must share NocoBase database"],
  ["A1.5", "A — Central Server", "API Hardening", "Structured logs, rotation, and health check", "Stdout + file logs; unauthenticated /health for load balancer; process manager restart.", "Developer", "Infra", "14-Sep-2026", "14-Sep-2026", 2, "A1.1", "High", "Not Started", "Central server", "Crash auto-restarts; /health is 200", "PM2, systemd, or NSSM"],
  ["A1.6", "A — Central Server", "API Hardening", "Allow-list office IPs or VPN on import POST", "Restrict POST /api/attendance/import to office networks.", "Infra", "Developer", "15-Sep-2026", "15-Sep-2026", 2, "A1.2", "High", "Not Started", "Central server", "Non-office IP is rejected at proxy or app", "Noida, Agra, Hyd egress IPs needed"],
  ["A1.7", "A — Central Server", "API Hardening", "Smoke import of last Noida batch", "Replay last successful Noida file; confirm UPSERT does not duplicate rows.", "Developer", "HR", "16-Sep-2026", "16-Sep-2026", 3, "A1.4", "High", "Not Started", "Noida", "Dashboard counts match POC; re-import updates only", "Use existing output batch as reference"],
  ["A2.1", "A — Central Server", "Host Deploy", "Provision VM or reuse NocoBase host", "Node 18+, git, firewall port 443, SSH access.", "Infra", "Developer", "14-Sep-2026", "14-Sep-2026", 3, "A0.1", "Critical", "Not Started", "Central server", "SSH in; node -v is 18 or higher", "Prefer same host as Postgres"],
  ["A2.2", "A — Central Server", "Host Deploy", "Deploy API via Docker/compose or PM2/systemd", "Do not use a manual npm start session as the run model.", "Developer", "Infra", "14-Sep-2026", "14-Sep-2026", 3, "A2.1", "Critical", "Not Started", "Central server", "Service starts on boot", "Dockerfile already in repo"],
  ["A2.3", "A — Central Server", "Host Deploy", "Connect NocoBase Postgres and apply schema", "Apply schema/postgres_baseline.sql; confirm att_* tables.", "Developer", "Infra", "14-Sep-2026", "15-Sep-2026", 2, "A2.2", "Critical", "Not Started", "Central server", "API log says PostgreSQL store", "Same DB as NocoBase 0.9"],
  ["A2.4", "A — Central Server", "Host Deploy", "Nightly backup of att_* and restore drill", "pg_dump of attendance tables; restore once onto a copy database.", "Infra", "Developer", "15-Sep-2026", "16-Sep-2026", 2, "A2.3", "High", "Not Started", "Central server", "Restore tested on a copy database", "Include in go-live evidence"],
  ["A2.5", "A — Central Server", "Host Deploy", "Write one-page server runbook", "Start/stop, logs, token rotate, rollback git tag.", "Developer", "Infra", "16-Sep-2026", "16-Sep-2026", 2, "A2.2", "Medium", "Not Started", "Central server", "Infra can restart without the developer on the call", "Keep to one page"],
  ["A3.1", "A — Central Server", "Server Acceptance", "Server acceptance pack", "Run /health, UI, POST import, Postgres persist, and idempotency checks.", "Developer", "Infra, HR", "16-Sep-2026", "16-Sep-2026", 4, "A1.7, A2.4", "Critical", "Not Started", "Central server", "All five server checks pass (see Server_Acceptance sheet)", "Hard gate before Noida pilot"],
  ["B1.1", "B — Local Scheduler", "Agent Behaviour", "Empty inbox exits success (code 0)", "Task Scheduler must not mark the job failed when inbox is empty.", "Developer", "Office IT", "16-Sep-2026", "17-Sep-2026", 1, "A0.1", "Critical", "Not Started", "All sites", "No files → log INFO, exit 0", "Current importer exits 1 on empty input/"],
  ["B1.2", "B — Local Scheduler", "Agent Behaviour", "Ignore temp Excel locks and in-copy files", "Skip ~$*.xls and files still being written.", "Developer", "—", "17-Sep-2026", "17-Sep-2026", 2, "B1.1", "High", "Not Started", "All sites", "Open/saving file is skipped until next 10-minute tick", "Common HR failure mode"],
  ["B1.3", "B — Local Scheduler", "Agent Behaviour", "Single-instance lock file", "Prevent two Task Scheduler ticks importing the same report.", "Developer", "—", "17-Sep-2026", "17-Sep-2026", 2, "B1.1", "High", "Not Started", "All sites", "Second process exits 0 with already running", "Required for 10-minute cadence"],
  ["B1.4", "B — Local Scheduler", "Agent Behaviour", "Skip already-processed content by file hash", "Re-drop of same file is UPSERT/log, not a crash.", "Developer", "—", "18-Sep-2026", "18-Sep-2026", 2, "B1.3", "Medium", "Not Started", "All sites", "Same file twice logged as duplicate or update", "API already UPSERTs by employee+date"],
  ["B1.5", "B — Local Scheduler", "Agent Behaviour", "PowerShell wrapper and last-run stamp", "scheduler.ps1 selects files, uses LOCATION_CODE, writes last-run timestamp.", "Developer", "Office IT", "17-Sep-2026", "18-Sep-2026", 3, "B1.1", "High", "Not Started", "All sites", "IT can run scheduler.ps1 by hand and see a one-line result", "Registered in Task Scheduler"],
  ["B1.6", "B — Local Scheduler", "Agent Behaviour", "Failure alert when failed folder or API down", "Log plus optional SMTP/Teams webhook within one tick.", "Developer", "Infra", "18-Sep-2026", "18-Sep-2026", 2, "B1.5", "High", "Not Started", "All sites", "Forced API-down test is visible within 10 minutes", "Do not email on empty inbox"],
  ["B2.1", "B — Local Scheduler", "Package & Pilot", "Build office install package and HR card", "Zip: Node note, .env.example, Task Scheduler XML, one-page HR drop-folder card.", "Developer", "Office IT", "21-Sep-2026", "21-Sep-2026", 3, "B1.6", "High", "Not Started", "All sites", "A non-developer follows the card and completes a dry run", "Python script excluded from zip"],
  ["B2.2", "B — Local Scheduler", "Package & Pilot", "Install Noida agent", "LOCATION_CODE=NOIDA; DDO_API_ENDPOINT = production HTTPS URL.", "Office IT", "Developer, HR", "22-Sep-2026", "22-Sep-2026", 3, "A3.1, B2.1", "Critical", "Not Started", "Noida", "Test drop appears on server dashboard within 10 minutes", "Pilot site"],
  ["B2.3", "B — Local Scheduler", "Package & Pilot", "Noida soak with real file and re-import", "One real monthly/weekly file plus re-import UPSERT check.", "HR", "Developer", "23-Sep-2026", "23-Sep-2026", 4, "B2.2", "Critical", "Not Started", "Noida", "HR confirms counts; no duplicate employees", "Do not skip soak if crashing the timeline"],
  ["B2.4", "B — Local Scheduler", "Package & Pilot", "Train one HR user on drop and failure folder", "User completes a drop without development in the room.", "HR", "Office IT", "23-Sep-2026", "23-Sep-2026", 2, "B2.2", "High", "Not Started", "Noida", "User finds failed/ without calling development", "Repeat at Agra and Hyd"],
  ["B3.1", "B — Local Scheduler", "Site Rollout", "Install Agra agent", "Same package; LOCATION_CODE=AGRA.", "Office IT", "Developer, HR", "24-Sep-2026", "24-Sep-2026", 3, "B2.3", "High", "Not Started", "Agra", "Agra file lands under Agra in All Locations view", "Do not share a PC across offices"],
  ["B3.2", "B — Local Scheduler", "Site Rollout", "Install Hyderabad agent", "LOCATION_CODE=HYD; sample file acceptable if no live export yet.", "Office IT", "Developer, HR", "25-Sep-2026", "25-Sep-2026", 3, "B2.3", "Medium", "Not Started", "Hyderabad", "Agent installed; sample or live file proves the path", "Live data may follow later"],
  ["C1.1", "C — Rollout & UAT", "UAT & Cutover", "Cross-site UAT versus last manual file", "Compare scheduled import to last manual import; own unknown-employee list.", "HR", "Developer", "28-Sep-2026", "28-Sep-2026", 8, "B3.1, B3.2", "Critical", "Not Started", "All sites", "Exception list assigned to named HR owner", "Unknown codes already seen on Noida POC"],
  ["C1.2", "C — Rollout & UAT", "UAT & Cutover", "Go-live rehearsal and rollback drill", "Disable Task Scheduler; roll API to previous git tag once.", "Project Manager", "Dev, Infra", "29-Sep-2026", "29-Sep-2026", 8, "C1.1", "Critical", "Not Started", "All sites", "Written go / no-go on 29 Sep", "No-go slips to next working day"],
  ["C1.3", "C — Rollout & UAT", "UAT & Cutover", "Cutover and hypercare day 1", "Production live; same-day response on failed/ folders through 2 Oct.", "Project Manager", "Dev, HR, Infra", "30-Sep-2026", "30-Sep-2026", 8, "C1.2", "Critical", "Not Started", "All sites", "Live traffic on scheduled path; hypercare roster named", "Hypercare 30 Sep–2 Oct"],
];

const CALENDAR_HEADERS = [
  "Day Number",
  "Calendar Date",
  "Day of Week",
  "Working Day (Y/N)",
  "Focus Area",
  "Workstream",
  "Planned Output",
  "Planned Hours",
  "Primary Owner",
  "Support Roles",
  "Milestone (Y/N)",
];

const CALENDAR_ROWS = [
  ["D1", "10-Sep-2026", "Thursday", "Y", "Kickoff and freeze decisions", "A + B", "Host, token owner, folder path, cadence signed", 8, "Project Manager", "Dev, Infra, HR", "Y"],
  ["D2", "11-Sep-2026", "Friday", "Y", "API bind, token, refuse JSON in prod", "A — Central Server", "Hardened build on a branch", 8, "Developer", "Infra", "N"],
  ["—", "12-Sep-2026", "Saturday", "N", "Non-working", "—", "—", 0, "—", "—", "N"],
  ["—", "13-Sep-2026", "Sunday", "N", "Non-working", "—", "—", 0, "—", "—", "N"],
  ["D3", "14-Sep-2026", "Monday", "Y", "Server provision and first deploy", "A — Central Server", "API up on host; /health green", 8, "Infra", "Developer", "Y"],
  ["D4", "15-Sep-2026", "Tuesday", "Y", "TLS, firewall, database backups", "A — Central Server", "HTTPS import from a non-local client", 8, "Infra", "Developer", "Y"],
  ["D5", "16-Sep-2026", "Wednesday", "Y", "Server UAT, runbook, start agent lock/empty-inbox", "A then B", "Server accepted; scheduler design locked", 8, "Developer", "Infra, HR", "Y"],
  ["D6", "17-Sep-2026", "Thursday", "Y", "Wrapper, lock, skip temp Excel files", "B — Local Scheduler", "scheduler.ps1 works locally", 8, "Developer", "Office IT", "N"],
  ["D7", "18-Sep-2026", "Friday", "Y", "Hash skip and failure alert", "B — Local Scheduler", "Forced-fail test visible in failed/", 8, "Developer", "Infra", "N"],
  ["—", "19-Sep-2026", "Saturday", "N", "Non-working", "—", "—", 0, "—", "—", "N"],
  ["—", "20-Sep-2026", "Sunday", "N", "Non-working", "—", "—", 0, "—", "—", "N"],
  ["D8", "21-Sep-2026", "Monday", "Y", "Package and install card", "B — Local Scheduler", "Drop-in zip for office IT", 8, "Developer", "Office IT", "N"],
  ["D9", "22-Sep-2026", "Tuesday", "Y", "Noida pilot install", "B — Local Scheduler", "First scheduled import in production DB", 8, "Office IT", "Developer, HR", "Y"],
  ["D10", "23-Sep-2026", "Wednesday", "Y", "Noida soak and HR training", "B — Local Scheduler", "Pilot sign-off or punch-list", 8, "HR", "Developer", "Y"],
  ["D11", "24-Sep-2026", "Thursday", "Y", "Agra rollout", "C — Rollout & UAT", "Agra LOCATION_CODE live", 6, "Office IT", "Developer, HR", "N"],
  ["D12", "25-Sep-2026", "Friday", "Y", "Hyderabad rollout", "C — Rollout & UAT", "Three agents installed", 6, "Office IT", "Developer, HR", "N"],
  ["—", "26-Sep-2026", "Saturday", "N", "Non-working / optional soak", "B — Local Scheduler", "Leave Noida job running", 0, "—", "—", "N"],
  ["—", "27-Sep-2026", "Sunday", "N", "Non-working / optional soak", "B — Local Scheduler", "Leave Noida job running", 0, "—", "—", "N"],
  ["D13", "28-Sep-2026", "Monday", "Y", "Cross-site UAT vs last manual file", "C — Rollout & UAT", "Unknown-employee list owned by HR", 8, "HR", "Developer", "Y"],
  ["D14", "29-Sep-2026", "Tuesday", "Y", "Go-live rehearsal and rollback drill", "C — Rollout & UAT", "Written go / no-go", 8, "Project Manager", "Dev, Infra", "Y"],
  ["D15", "30-Sep-2026", "Wednesday", "Y", "Cutover and hypercare day 1", "C — Rollout & UAT", "Live; 3-day hypercare into 1–2 Oct", 8, "Project Manager", "Dev, HR, Infra", "Y"],
];

const RAID_HEADERS = [
  "RAID ID",
  "Type",
  "Title",
  "Description",
  "Likelihood",
  "Impact",
  "Risk Score",
  "Mitigation / Response",
  "Owner Role",
  "Due Date",
  "Status",
];

const RAID_ROWS = [
  ["R-01", "Risk", "API still bound to localhost", "Office schedulers appear to run but data never reaches the server.", "High", "High", "High", "Noida pilot is blocked until /health works from that PC.", "Developer", "16-Sep-2026", "Open"],
  ["R-02", "Risk", "No route from office PC to server:443", "Site cannot import.", "Medium", "High", "High", "Day-1 network test; document VPN fallback.", "Infra", "10-Sep-2026", "Open"],
  ["R-03", "Risk", "Excel file dropped while still open", "Partial read lands in failed/.", "High", "Medium", "High", "Skip locked and ~$ files; next tick retries.", "Developer", "17-Sep-2026", "Open"],
  ["R-04", "Risk", "Unknown employee codes", "Rows skipped; HR thinks import failed. Already seen on Noida POC.", "High", "High", "High", "Dashboard exceptions plus HR master-data cleanup on D13.", "HR", "28-Sep-2026", "Open"],
  ["R-05", "Risk", "Two importers confuse operators", "Python and Node both exist; wrong script misses inbox.", "Medium", "Medium", "Medium", "Python labelled lab-only; offices receive Node zip only.", "Project Manager", "21-Sep-2026", "Open"],
  ["A-01", "Assumption", "NocoBase Postgres reachable from API host", "No persist if DATABASE_URL is wrong.", "—", "High", "High", "Confirm DATABASE_URL on D1; deploy sidecar on same host.", "Infra", "10-Sep-2026", "Open"],
  ["A-02", "Assumption", "One LOCATION_CODE per machine", "Wrong location tagging if a PC is shared across offices.", "—", "High", "High", "Do not share one PC across offices.", "Office IT", "22-Sep-2026", "Open"],
  ["D-01", "Dependency", "TLS certificate and DNS name", "Blocks D4 HTTPS work.", "—", "High", "High", "If public cert delayed, use internal CA — do not skip TLS on WAN.", "Infra", "15-Sep-2026", "Open"],
  ["D-02", "Dependency", "Named HR owner for unknown employees", "UAT cannot close.", "—", "High", "High", "Name an HR owner on D1.", "Project Manager", "10-Sep-2026", "Open"],
  ["D-03", "Dependency", "Office egress IPs or VPN", "Cannot complete import allow-list.", "—", "Medium", "Medium", "Collect Noida, Agra, Hyd public IPs during kickoff.", "Infra", "10-Sep-2026", "Open"],
];

const GOLIVE_HEADERS = [
  "Item No.",
  "Category",
  "Go-Live Criterion",
  "Owner Role",
  "Evidence Required",
  "Status",
  "Sign-off Date",
  "Sign-off Name",
];

const GOLIVE_ROWS = [
  ["GL-01", "Infrastructure", "API /health and UI reachable over HTTPS from Noida, Agra, and Hyd PCs", "Infra", "Screenshot or curl log from each office PC", "Pending", "", ""],
  ["GL-02", "Security", "Production token in use; POC token disabled", "Developer", "401 response with old token; secrets store entry", "Pending", "", ""],
  ["GL-03", "Data", "PostgreSQL is the store; JSON api_data is not used in production", "Developer", "API startup log line showing PostgreSQL store", "Pending", "", ""],
  ["GL-04", "Resilience", "Restore drill of att_* completed once", "Infra", "Restore log on copy database", "Pending", "", ""],
  ["GL-05", "Functional", "Noida scheduled import of a real file matches last manual import (UPSERT, not extra rows)", "HR", "Dashboard counts + batch summary", "Pending", "", ""],
  ["GL-06", "Rollout", "Agra and Hyd agents installed; Hyd may use a sample file", "Office IT", "Task Scheduler screenshot per site", "Pending", "", ""],
  ["GL-07", "Training", "HR can drop a file and find failed/ without calling development", "HR", "Training completion note", "Pending", "", ""],
  ["GL-08", "Master Data", "Unknown-employee list assigned to named HR owner", "HR", "Named owner on RAID D-02", "Pending", "", ""],
  ["GL-09", "Rollback", "Disable Task Scheduler; previous API git tag documented", "Developer", "Runbook section with tag name", "Pending", "", ""],
  ["GL-10", "Governance", "PM + HR written go / no-go on 29 Sep", "Project Manager", "Email or signed row on this sheet", "Pending", "", ""],
];

const RESOURCE_HEADERS = [
  "Role",
  "Calendar Days Required",
  "Estimated Hours",
  "Responsibility",
  "RACI — Responsible",
  "RACI — Accountable",
  "RACI — Consulted",
  "RACI — Informed",
];

const RESOURCE_ROWS = [
  ["Developer", "D1–D15 (all 15 working days)", 112, "Harden API, scheduler wrapper, package, sit with first site", "A1–B2, C1.2 technical", "Project Manager", "Infra, HR", "Office IT"],
  ["Infra / Server Owner", "D1, D3–D5, D8, D14", 24, "VM, TLS, firewall, backups, DNS", "A2.1, A1.2, A1.6, A2.4", "Project Manager", "Developer", "HR"],
  ["Office IT (each site)", "D9, D11, D12", 12, "Local Node install and Task Scheduler", "B2.2, B3.1, B3.2", "Project Manager", "Developer", "HR"],
  ["HR (Noida lead)", "D9–D10, D13–D15", 20, "Drop files, confirm dashboard, unknown-employee list", "B2.3, B2.4, C1.1", "Project Manager", "Developer", "Office IT"],
  ["HR (Agra / Hyderabad)", "D11–D13", 8, "Repeat drop-folder training", "B3.1, B3.2 support", "HR Noida lead", "Office IT", "Project Manager"],
  ["Project Manager", "D1, D13–D15", 16, "Decisions, RAID, go / no-go, hypercare roster", "A0.1, C1.2, C1.3, GL-10", "Sponsor", "All roles", "Sponsor"],
];

const FOLDER_HEADERS = [
  "Folder Name",
  "Path Convention",
  "Who Writes",
  "Who Reads",
  "Purpose",
];

const FOLDER_ROWS = [
  ["inbox", "<install>\\inbox\\", "HR", "Scheduler", "Drop .xls / .xlsx reports here"],
  ["processing", "<install>\\processing\\", "Scheduler", "Scheduler", "File moved here while import runs (lock)"],
  ["processed", "<install>\\processed\\", "Scheduler", "IT / HR (audit)", "Success archive"],
  ["failed", "<install>\\failed\\", "Scheduler", "HR / IT", "Parse or API failure; do not retry forever"],
  ["logs", "<install>\\logs\\importer.log", "Scheduler", "IT", "Rolling operational log"],
];

const ACCEPT_HEADERS = [
  "Check ID",
  "Workstream",
  "Check Name",
  "Pass Rule",
  "Status",
  "Tested Date",
  "Tester Name",
];

const ACCEPT_ROWS = [
  ["SA-01", "A — Central Server", "/health", "HTTP 200 from inside the VPC and from one office PC", "Pending", "", ""],
  ["SA-02", "A — Central Server", "UI /", "Loads over HTTPS; location filter still works", "Pending", "", ""],
  ["SA-03", "A — Central Server", "POST /api/attendance/import", "Bearer + location header accepted; 401 without token", "Pending", "", ""],
  ["SA-04", "A — Central Server", "PostgreSQL persist", "New batch visible in att_import_batches and att_attendance", "Pending", "", ""],
  ["SA-05", "A — Central Server", "Idempotency", "Second POST of same file updates rows; count unchanged", "Pending", "", ""],
  ["SB-01", "B — Local Scheduler", "Empty inbox", "Exit code 0; INFO log; Task Scheduler shows Success", "Pending", "", ""],
  ["SB-02", "B — Local Scheduler", "Locked Excel skip", "Open file skipped; imported on next tick after close", "Pending", "", ""],
  ["SB-03", "B — Local Scheduler", "Single instance", "Overlapping tick does not double-import", "Pending", "", ""],
  ["SB-04", "B — Local Scheduler", "Noida 10-minute SLA", "Dropped file visible on dashboard within 10 minutes", "Pending", "", ""],
  ["SB-05", "B — Local Scheduler", "Location tagging", "Agra file does not appear as Noida", "Pending", "", ""],
];

async function main() {
  const wb = new ExcelJS.Workbook();
  wb.creator = "DDO++ Attendance Project Manager";
  wb.created = new Date(2026, 8, 9);
  wb.modified = new Date(2026, 8, 9);
  wb.title = "Attendance Go-Live Plan — Server and Local Scheduler";
  wb.description =
    "15 working day plan to deploy attendance_api.js on a shared server and schedule attendance_import.js on office PCs.";

  coverSheet(wb);

  addTableSheet(
    wb,
    "02_Master_Task_Register",
    "Master Task Register",
    "All delivery tasks · filter by Workstream, Owner Role, Status, or Site / Location",
    MASTER_HEADERS,
    MASTER_ROWS,
    MASTER_WIDTHS
  );

  addTableSheet(
    wb,
    "03_Delivery_Calendar",
    "Delivery Calendar",
    "10 Sep 2026 – 30 Sep 2026 · IST · weekends off · 15 working days",
    CALENDAR_HEADERS,
    CALENDAR_ROWS,
    [12, 14, 14, 16, 42, 22, 42, 14, 18, 20, 16]
  );

  addTableSheet(
    wb,
    "04_RAID_Log",
    "RAID Log",
    "Risks, Assumptions, Issues, Dependencies — review at kickoff and every Friday",
    RAID_HEADERS,
    RAID_ROWS,
    [12, 14, 32, 48, 14, 12, 12, 48, 16, 14, 12]
  );

  addTableSheet(
    wb,
    "05_Go_Live_Checklist",
    "Go-Live Checklist",
    "Cut over on 30 Sep 2026 only if every row is Done · hypercare 30 Sep–2 Oct",
    GOLIVE_HEADERS,
    GOLIVE_ROWS,
    [12, 16, 72, 16, 40, 12, 14, 16]
  );

  addTableSheet(
    wb,
    "06_Resource_RACI",
    "Resource Plan and RACI",
    "Who is needed when · hours are estimates for planning, not timesheets",
    RESOURCE_HEADERS,
    RESOURCE_ROWS,
    [24, 32, 16, 48, 24, 18, 18, 16]
  );

  addTableSheet(
    wb,
    "07_Folder_Contract",
    "Local Folder Contract",
    "Same layout on every office PC · one LOCATION_CODE per machine",
    FOLDER_HEADERS,
    FOLDER_ROWS,
    [16, 28, 14, 18, 48]
  );

  addTableSheet(
    wb,
    "08_Acceptance_Checks",
    "Acceptance Checks",
    "Server pack on 16 Sep · Scheduler pack during Noida soak 22–23 Sep",
    ACCEPT_HEADERS,
    ACCEPT_ROWS,
    [12, 22, 28, 56, 12, 14, 16]
  );

  await wb.xlsx.writeFile(OUT_FILE);
  console.log(`Wrote ${OUT_FILE}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
