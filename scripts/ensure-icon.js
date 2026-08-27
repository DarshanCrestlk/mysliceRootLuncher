const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");
const { pngToIco, writeIconFiles } = require("./png-to-ico");

const TARGETS = [
  path.join(__dirname, "..", "officeLuncher", "build"),
  path.join(__dirname, "..", "mysliceLauncher", "build"),
];

function findSourcePng() {
  const home = process.env.USERPROFILE || process.env.HOME || "";
  const candidates = [
    path.join(
      home,
      ".cursor",
      "projects",
      "d-myslice-add-in-myslice-windows",
      "assets",
      "myslice-icon.png"
    ),
    path.join(__dirname, "..", "officeLuncher", "build", "icon.png"),
    path.join(__dirname, "..", "mysliceLauncher", "build", "icon.png"),
  ];
  return candidates.find((file) => fs.existsSync(file));
}

function generatePngWithPowerShell(outPng) {
  const ps1 = path.join(__dirname, "generate-icon.ps1");
  const result = spawnSync(
    "powershell.exe",
    ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps1, "-OutFile", outPng],
    { stdio: "inherit", windowsHide: true }
  );
  if (result.status !== 0 || !fs.existsSync(outPng)) {
    throw new Error("Failed to generate MySlice icon.png");
  }
}

function ensureIcons() {
  let pngPath = findSourcePng();
  const stagingPng = path.join(__dirname, "..", "officeLuncher", "build", "icon.png");

  if (!pngPath) {
    fs.mkdirSync(path.dirname(stagingPng), { recursive: true });
    generatePngWithPowerShell(stagingPng);
    pngPath = stagingPng;
  }

  const pngBuffer = fs.readFileSync(pngPath);
  for (const dir of TARGETS) {
    writeIconFiles(dir, pngBuffer);
  }
}

if (require.main === module) {
  try {
    ensureIcons();
  } catch (err) {
    console.error(err.message || err);
    process.exit(1);
  }
}

module.exports = { ensureIcons };
