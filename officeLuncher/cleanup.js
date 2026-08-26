const { execFileSync } = require("child_process");
const fs = require("fs");
const os = require("os");
const path = require("path");

const BASE_DIR = path.join(process.env.PROGRAMDATA || "C:\\ProgramData", "myslice", "mysliceLTS");
const MYSLICE_ROOT = path.join(process.env.PROGRAMDATA || "C:\\ProgramData", "myslice");
const CATALOG_GUID = "c77550fc-0d50-495e-be1a-8695539e5d54";

function runTempPS(script, errorMsg) {
  const file = path.join(os.tmpdir(), `myslice_cleanup_${Date.now()}.ps1`);
  try {
    fs.writeFileSync(file, "\uFEFF" + script, { encoding: "utf8" });
    execFileSync(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File", file],
      { windowsHide: true }
    );
    return true;
  } catch (err) {
    console.error(errorMsg, err.message || err);
    return false;
  } finally {
    if (fs.existsSync(file)) fs.unlinkSync(file);
  }
}

function stopLauncher() {
  try {
    execFileSync("taskkill", ["/F", "/IM", "mysliceLTS.exe", "/T"], {
      windowsHide: true,
      stdio: "ignore",
    });
  } catch (_err) {
    // not running
  }

  runTempPS(
    `
Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
  Where-Object { $_.Name -match 'powershell' -and $_.CommandLine -and ($_.CommandLine -like '*myslice.ps1*') } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
`,
    "Stop myslice.ps1 failed"
  );
  return true;
}

function removeShare() {
  return runTempPS(
    `
$share = Get-SmbShare -Name "mysliceLTS" -ErrorAction SilentlyContinue
if ($share) { Remove-SmbShare -Name "mysliceLTS" -Force }
`,
    "SMB share remove failed"
  );
}

function removeProtocol() {
  return runTempPS(
    `
$baseKey = "Registry::HKEY_CLASSES_ROOT\\mysliceLTS"
if (Test-Path $baseKey) { Remove-Item -Path $baseKey -Recurse -Force }
`,
    "Protocol remove failed"
  );
}

function removeTrustedCatalog() {
  return runTempPS(
    `
$guid = "${CATALOG_GUID}"
$paths = @(
  "HKCU:\\Software\\Policies\\Microsoft\\Office\\16.0\\WEF\\TrustedCatalogs\\{$guid}",
  "HKLM:\\Software\\Policies\\Microsoft\\Office\\16.0\\WEF\\TrustedCatalogs\\{$guid}",
  "HKCU:\\Software\\Microsoft\\Office\\16.0\\WEF\\TrustedCatalogs\\{$guid}"
)
foreach ($p in $paths) {
  if (Test-Path $p) { Remove-Item -Path $p -Recurse -Force }
}
`,
    "Trusted catalog remove failed"
  );
}

function removeProgramData() {
  try {
    if (fs.existsSync(BASE_DIR)) {
      fs.rmSync(BASE_DIR, { recursive: true, force: true });
    }
    if (fs.existsSync(MYSLICE_ROOT) && fs.readdirSync(MYSLICE_ROOT).length === 0) {
      fs.rmSync(MYSLICE_ROOT, { recursive: true, force: true });
    }
    return true;
  } catch (err) {
    console.error("ProgramData remove failed:", err.message || err);
    return false;
  }
}

function removeMySlice() {
  stopLauncher();

  const failedSteps = [];
  if (!removeShare()) failedSteps.push("Network share");
  if (!removeProtocol()) failedSteps.push("Protocol mysliceLTS://");
  if (!removeTrustedCatalog()) failedSteps.push("Office trusted catalog");
  if (!removeProgramData()) failedSteps.push("ProgramData files");

  return {
    ok: failedSteps.length === 0,
    failedSteps,
  };
}

module.exports = {
  removeMySlice,
};
