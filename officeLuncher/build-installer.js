const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");
const { ensureIcons } = require("../scripts/ensure-icon");

ensureIcons();

const unpacked = path.join(__dirname, "..", "mysliceLauncher", "dist", "win-unpacked");
const exe = path.join(unpacked, "mysliceLTS.exe");
const staging = path.resolve(__dirname, "dist", "prepackaged");
const iconIco = path.join(__dirname, "build", "icon.ico");

if (!fs.existsSync(exe)) {
  console.error("mysliceLauncher is not built yet.");
  console.error("Run this first:");
  console.error("  cd ../mysliceLauncher && npm install && npm run build");
  process.exit(1);
}

function getElectronVersion() {
  const versionFile = path.join(unpacked, "version");
  if (fs.existsSync(versionFile)) {
    const fromDist = fs.readFileSync(versionFile, "utf8").trim();
    if (fromDist) return fromDist;
  }

  const electronPkg = path.join(
    __dirname,
    "..",
    "mysliceLauncher",
    "node_modules",
    "electron",
    "package.json"
  );
  if (fs.existsSync(electronPkg)) {
    return JSON.parse(fs.readFileSync(electronPkg, "utf8")).version;
  }

  return "31.3.1";
}

function stampExeIcon(exePath, icoPath) {
  if (!fs.existsSync(exePath) || !fs.existsSync(icoPath)) return;

  const localAppData = process.env.LOCALAPPDATA || "";
  const rceditCandidates = [
    path.join(localAppData, "electron-builder", "Cache", "winCodeSign", "rcedit-x64.exe"),
    path.join(localAppData, "electron-builder", "Cache", "winCodeSign", "rcedit.exe"),
    path.join(__dirname, "node_modules", "rcedit", "bin", "rcedit.exe"),
  ];
  const rcedit = rceditCandidates.find((file) => fs.existsSync(file));
  if (!rcedit) {
    console.log("rcedit not found; setup EXE will still use the MySlice icon.");
    return;
  }

  const stamped = spawnSync(rcedit, [exePath, "--set-icon", icoPath], {
    stdio: "inherit",
    windowsHide: true,
  });
  if (stamped.status === 0) {
    console.log("Applied MySlice icon to", exePath);
  }
}

fs.rmSync(staging, { recursive: true, force: true });
fs.mkdirSync(path.dirname(staging), { recursive: true });
fs.cpSync(unpacked, staging, { recursive: true });
stampExeIcon(path.join(staging, "mysliceLTS.exe"), iconIco);

const resources = path.join(staging, "resources");
fs.mkdirSync(resources, { recursive: true });
fs.copyFileSync(path.join(__dirname, "manifest.xml"), path.join(resources, "manifest.xml"));
fs.copyFileSync(path.join(__dirname, "install.ps1"), path.join(resources, "install.ps1"));
fs.copyFileSync(path.join(__dirname, "uninstall.ps1"), path.join(resources, "uninstall.ps1"));

const electronVersion = getElectronVersion();
console.log("Packaging launcher with Electron", electronVersion);

const result = spawnSync(
  "npx",
  [
    "electron-builder",
    "--prepackaged",
    staging,
    "--publish",
    "never",
    `--config.electronVersion=${electronVersion}`,
  ],
  {
    cwd: __dirname,
    stdio: "inherit",
    shell: true,
  }
);

process.exit(result.status === null ? 1 : result.status);
