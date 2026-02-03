// Extended process shim for browser that includes pid
// just-bash requires process.pid for BASHPID

var process = require("process/browser");

// Add missing properties needed by just-bash
process.pid = 1;
process.ppid = 0;
process.platform = "linux";
process.arch = "x64";

module.exports = process;
