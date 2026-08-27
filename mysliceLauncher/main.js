const { app, BrowserWindow } = require("electron");
const { exec } = require("child_process");
const path = require("path");

const PROTOCOL = "mysliceLTS";
const PS_SCRIPT_PATH = app.isPackaged
  ? path.join(process.resourcesPath, "myslice.ps1")
  : path.join(__dirname, "myslice.ps1");

let isQuitting = false;
let psChildren = 0;

const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
} else {
  boot();
}

function boot() {
  app.on("second-instance", (_, commandLine) => {
    handleProtocolInvocation(commandLine);
  });

  app.whenReady().then(() => {
    app.setAsDefaultProtocolClient(PROTOCOL);

    app.on("open-url", (event, url) => {
      event.preventDefault();
      handleUrl(url);
    });

    handleProtocolInvocation(process.argv);
  });

  app.on("before-quit", (event) => {
    if (psChildren > 0) {
      event.preventDefault();
      isQuitting = false;
      return;
    }
    isQuitting = true;
  });

  app.on("window-all-closed", () => {
    if (psChildren === 0) {
      maybeQuit();
    }
  });
}

function maybeQuit() {
  if (psChildren > 0 || isQuitting) return;
  isQuitting = true;
  app.quit();
}

function createFocusStealer() {
  const win = new BrowserWindow({
    width: 200,
    height: 200,
    show: true,
    frame: false,
    transparent: true,
    alwaysOnTop: true,
    focusable: true,
    skipTaskbar: true,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
  });

  win.loadURL("about:blank");

  win.once("focus", () => {
    setTimeout(() => {
      if (!win.isDestroyed()) win.hide();
    }, 100);
  });

  setTimeout(() => {
    if (!win.isDestroyed()) win.hide();
  }, 400);
}

function extractUrl(args) {
  if (!Array.isArray(args)) return null;

  for (const arg of args) {
    if (typeof arg === "string") {
      const lower = arg.toLowerCase();
      if (
        lower.startsWith(`${PROTOCOL.toLowerCase()}://`) ||
        lower.startsWith(`${PROTOCOL.toLowerCase()}:/`)
      ) {
        return arg;
      }
    }
  }
  return null;
}

function cleanUrl(rawUrl) {
  if (typeof rawUrl !== "string") return "";
  return rawUrl
    .replace(new RegExp(`^${PROTOCOL}://`, "i"), "")
    .replace(new RegExp(`^${PROTOCOL}:/`, "i"), "")
    .replace(/^\/+/, "");
}

function escapePsDoubleQuoted(value) {
  return String(value).replace(/"/g, '`"');
}

function runPowerShell(cleanedUrl) {
  const cmd = `powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File "${PS_SCRIPT_PATH}" -url "${escapePsDoubleQuoted(cleanedUrl)}"`;

  psChildren += 1;
  let settled = false;
  const finish = () => {
    if (settled) return;
    settled = true;
    psChildren = Math.max(0, psChildren - 1);
    console.log("PowerShell exited. Remaining children:", psChildren);
    maybeQuit();
  };

  exec(cmd, { windowsHide: true }, (error, stdout, stderr) => {
    if (error) console.error("PS Error:", error);
    if (stderr) console.error("PS Stderr:", stderr);
    if (stdout) console.log("PS Stdout:", stdout);
    finish();
  });
}

function handleUrl(rawUrl) {
  if (!rawUrl) {
    if (psChildren === 0) maybeQuit();
    return;
  }

  createFocusStealer();

  const cleaned = cleanUrl(rawUrl);
  console.log("Received:", rawUrl);
  console.log("Cleaned URL passed to PS:", cleaned);

  runPowerShell(cleaned);
}

function handleProtocolInvocation(argv) {
  const url = extractUrl(argv);
  if (url) {
    handleUrl(url);
  } else if (psChildren === 0) {
    maybeQuit();
  }
}
