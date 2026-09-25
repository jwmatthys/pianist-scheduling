const { existsSync } = require("node:fs");
const { spawnSync } = require("node:child_process");
const path = require("node:path");

const backendPath = path.resolve(__dirname, "..", "..", "backend");
const venvPython = process.platform === "win32"
  ? path.join(backendPath, ".venv", "Scripts", "python.exe")
  : path.join(backendPath, ".venv", "bin", "python");
const python = process.env.PYTHON ?? (existsSync(venvPython) ? venvPython : "python");
const result = spawnSync(python, [
  "-m", "PyInstaller",
  "--noconfirm",
  "--clean",
  "--onefile",
  "--name", "pianist-scheduling-api",
  "--collect-all", "pandas",
  "--collect-all", "openpyxl",
  "desktop_server.py",
], { cwd: backendPath, stdio: "inherit" });

if (result.error?.code === "ENOENT") {
  console.error(`Python executable not found: ${python}`);
  process.exit(1);
}

if (result.status !== 0) {
  console.error(`\nInstall backend build dependencies with:\n  ${python} -m pip install -r requirements.txt`);
  process.exit(result.status ?? 1);
}