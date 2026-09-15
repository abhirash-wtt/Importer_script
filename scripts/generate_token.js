"use strict";

const crypto = require("crypto");
const token = crypto.randomBytes(32).toString("hex");
console.log(token);
console.log("");
console.log("Set DDO_API_TOKEN to this value on the server and every office PC.");
console.log("Do not commit it. The old POC token is dev-ddo-attendance-token.");
