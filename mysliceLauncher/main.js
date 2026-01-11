// main.js - Robust hidden protocol launcher with forced browser blur
const { app, BrowserWindow } = require('electron');
const { exec } = require('child_process');
const path = require('path');

const PROTOCOL = 'mysliceLTS';
// const PS_SCRIPT_PATH = 'D:\\mysliceLauncher\\myslice.ps1'; // Update if needed
const PS_SCRIPT_PATH = 'C:\\ProgramData\\myslice\\mysliceLTS\\launcher\\resources\\myslice.ps1'; // Update if needed

let isQuitting = false;

// Prevent multiple instances
const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
  return;
}

// Second instance → forward the protocol URL
app.on('second-instance', (_, commandLine) => {
  if (isQuitting) return;
  handleProtocolInvocation(commandLine);
});

// Create a tiny temporary window to steal focus → triggers browser blur
function createFocusStealer() {
  const win = new BrowserWindow({
    width: 200,
    height: 200,
    show: true,              // Must be visible briefly to steal focus
    frame: false,
    transparent: true,
    alwaysOnTop: true,
    focusable: true,
    skipTaskbar: true,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true
    }
  });

  win.loadURL('about:blank');

  // Hide as soon as it gains focus (browser registers blur)
  win.once('focus', () => {
    setTimeout(() => {
      if (!win.isDestroyed()) win.hide();
    }, 100); // 100ms is reliable
  });

  // Safety fallback
  setTimeout(() => {
    if (!win.isDestroyed()) win.hide();
  }, 400);
}

// Safely extract the protocol URL from args
function extractUrl(args) {
  if (!Array.isArray(args)) return null;

  for (const arg of args) {
    if (typeof arg === 'string') {
      const lower = arg.toLowerCase();
      if (lower.startsWith(`${PROTOCOL.toLowerCase()}://`) ||
          lower.startsWith(`${PROTOCOL.toLowerCase()}:/`)) {
        return arg; // return original (with correct case)
      }
    }
  }
  return null;
}

// Clean the URL (remove protocol and leading slashes)
function cleanUrl(rawUrl) {
  if (typeof rawUrl !== 'string') return '';
  let url = rawUrl
    .replace(new RegExp(`^${PROTOCOL}://`, 'i'), '')
    .replace(new RegExp(`^${PROTOCOL}:/`, 'i'), '')
    .replace(/^\/+/, ''); // remove leading slashes
  return url;
}

// Execute PowerShell script with cleaned URL
function runPowerShell(cleanedUrl) {
  const cmd = `powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File "${PS_SCRIPT_PATH}" -url "${cleanedUrl}"`;

  exec(cmd, { windowsHide: true }, (error, stdout, stderr) => {
    if (error) console.error('PS Error:', error);
    if (stderr) console.error('PS Stderr:', stderr);
    if (stdout) console.log('PS Stdout:', stdout);

    // Quit shortly after script finishes
    setTimeout(() => app.quit(), 500);
  });
}

// Main handler for protocol URLs
function handleUrl(rawUrl) {
  if (!rawUrl) {
    app.quit();
    return;
  }

  // Force browser blur by stealing focus briefly
  createFocusStealer();

  const cleaned = cleanUrl(rawUrl);
  console.log('Received:', rawUrl);
  console.log('Cleaned URL passed to PS:', cleaned);

  runPowerShell(cleaned);
}

// Process command line or open-url event
function handleProtocolInvocation(argv) {
  const url = extractUrl(argv);
  if (url) {
    handleUrl(url);
  } else {
    app.quit(); // No protocol → just exit silently
  }
}

// App ready
app.whenReady().then(() => {
  // Register protocol client
  app.setAsDefaultProtocolClient(PROTOCOL);

  // Handle macOS/open-url (and some Windows cases)
  app.on('open-url', (event, url) => {
    event.preventDefault();
    handleUrl(url);
  });

  // Handle initial launch (Windows passes URL in argv)
  handleProtocolInvocation(process.argv);
});

// Cleanup
app.on('before-quit', () => {
  isQuitting = true;
});

app.on('window-all-closed', () => {
  // Don't quit on macOS automatically, but we override anyway
  if (process.platform !== 'darwin') app.quit();
});