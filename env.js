"use strict";

const fs = require("fs");
const path = require("path");

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

function loadProjectEnv() {
  loadEnvFile(path.join(__dirname, ".env"));
}

function isProduction() {
  return String(process.env.NODE_ENV || "").toLowerCase() === "production";
}

loadProjectEnv();

module.exports = { loadEnvFile, loadProjectEnv, isProduction };
