const { app, BrowserWindow, dialog } = require("electron");
const { spawn } = require("node:child_process");
const net = require("node:net");
const path = require("node:path");

let apiProcess;

function reservePort() {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.unref();
    server.on("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const { port } = server.address();
      server.close((error) => (error ? reject(error) : resolve(port)));
    });
  });
}

function waitForApi(port, timeoutMs = 10000) {
  const endpoint = `http://127.0.0.1:${port}/api/health`;
  const deadline = Date.now() + timeoutMs;
  return new Promise((resolve, reject) => {
    const check = async () => {
      try {
        const response = await fetch(endpoint);
        if (response.ok) return resolve();
      } catch {
        // The bundled service is still starting.
      }
      if (Date.now() >= deadline) return reject(new Error("The scheduling service did not start."));
      setTimeout(check, 150);
    };
    check();
  });
}

function startApi(port) {
  const executable = process.platform === "win32" ? "pianist-scheduling-api.exe" : "pianist-scheduling-api";
  const servicePath = app.isPackaged
    ? path.join(process.resourcesPath, "api", executable)
    : path.join(__dirname, "..", "..", "backend", "dist", executable);
  const databasePath = path.join(app.getPath("userData"), "pianist_scheduling.db");
  apiProcess = spawn(servicePath, ["--port", String(port)], {
    env: { ...process.env, PIANIST_SCHEDULING_DB_PATH: databasePath },
    windowsHide: true,
  });
  apiProcess.on("error", (error) => {
    dialog.showErrorBox("Unable to start Pianist Scheduling", error.message);
  });
}

async function createWindow() {
  const port = await reservePort();
  startApi(port);
  await waitForApi(port);

  const window = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 1024,
    minHeight: 700,
    webPreferences: { contextIsolation: true, nodeIntegration: false },
  });
  await window.loadFile(path.join(__dirname, "..", "dist", "index.html"), {
    query: { apiBase: `http://127.0.0.1:${port}` },
  });
}

app.whenReady().then(createWindow).catch((error) => {
  dialog.showErrorBox("Unable to start Pianist Scheduling", error.message);
  app.quit();
});

app.on("window-all-closed", () => app.quit());
app.on("before-quit", () => apiProcess?.kill());