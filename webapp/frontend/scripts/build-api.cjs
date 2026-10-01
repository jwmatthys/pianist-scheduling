const { existsSync } = require("node:fs");
const { spawnSync } = require("node:child_process");
const path = require("node:path");

const tauriSidecar = process.argv.includes("--tauri");
let executableName = "pianist-scheduling-api";

if (tauriSidecar) {
  const rustc = process.env.RUSTC ?? "rustc";
  const result = spawnSync(rustc, ["--print", "host-tuple"], { encoding: "utf8" });
  if (result.error || result.status !== 0) {
    console.error(`Unable to determine the Rust host triple with ${rustc}.`);
    process.exit(1);
  }

  const hostTriple = result.stdout.trim();
  const targetTriple = process.env.TAURI_ENV_TARGET_TRIPLE ?? hostTriple;
  if (targetTriple !== hostTriple) {
    console.error(
      `PyInstaller must build natively: requested target ${targetTriple}, host is ${hostTriple}.`
    );
    process.exit(1);
  }
  executableName = `pianist-scheduling-api-${targetTriple}`;
}

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
  "--name", executableName,
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