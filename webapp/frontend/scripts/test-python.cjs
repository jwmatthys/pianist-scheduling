const { existsSync } = require("node:fs");
const { spawnSync } = require("node:child_process");
const path = require("node:path");

const frontendPath = path.resolve(__dirname, "..");
const repoPath = path.resolve(frontendPath, "..", "..");
const backendPath = path.join(repoPath, "webapp", "backend");
const venvPython = process.platform === "win32"
  ? path.join(backendPath, ".venv", "Scripts", "python.exe")
  : path.join(backendPath, ".venv", "bin", "python");
const python = process.env.PYTHON
  ?? (existsSync(venvPython) ? venvPython : process.platform === "win32" ? "python" : "python3");
const pythonPath = [backendPath, process.env.PYTHONPATH].filter(Boolean).join(path.delimiter);
const result = spawnSync(python, [
  "-m", "unittest",
  "discover",
  "-s", path.join("webapp", "backend", "tests"),
  "-v",
], {
  cwd: repoPath,
  env: { ...process.env, PYTHONPATH: pythonPath },
  stdio: "inherit",
});

if (result.error?.code === "ENOENT") {
  console.error(`Python executable not found: ${python}`);
  process.exit(1);
}

if (result.status !== 0) process.exit(result.status ?? 1);
