const { app, dialog } = require("electron");
const { execFileSync } = require("child_process");
const path = require("path");
const fs = require("fs");
const os = require("os");

// ---------------- CONFIG ----------------
const BASE_DIR = path.join(process.env.PROGRAMDATA, "myslice", "mysliceLTS");
const MANIFEST_DIR = path.join(BASE_DIR, "manifest");
const MANIFEST_TARGET = path.join(MANIFEST_DIR, "manifest.xml");
const MYSLICE_EXE_TARGET = path.join(BASE_DIR, "launcher", "mysliceLTS.exe");

// ---------------- HELPERS ----------------

// Copy manifest.xml to ProgramData
function setupManifest() {
  try {
    const sourceManifest = app.isPackaged
      ? path.join(process.resourcesPath, "manifest.xml")
      : path.join(__dirname, "manifest.xml");

    if (!fs.existsSync(sourceManifest)) {
      console.error("Source manifest.xml not found:", sourceManifest);
      return false;
    }

    fs.mkdirSync(MANIFEST_DIR, { recursive: true });
    fs.copyFileSync(sourceManifest, MANIFEST_TARGET);
    console.log("Manifest deployed →", MANIFEST_TARGET);
    return true;
  } catch (err) {
    console.error("Manifest setup failed:", err.message);
    return false;
  }
}

// Copy myslice.exe to ProgramData
function setupMYSliceExe() {
  try {
    const sourceFolder = app.isPackaged
      ? path.join(process.resourcesPath, "mysliceLauncher")
      : path.join(__dirname, "..", "mysliceLauncher", "dist", "win-unpacked");

    const targetFolder = path.join(BASE_DIR, "launcher");

    if (!fs.existsSync(sourceFolder)) {
      console.error("Source launcher folder not found:", sourceFolder);
      return false;
    }

    fs.mkdirSync(targetFolder, { recursive: true });
    fs.cpSync(sourceFolder, targetFolder, { recursive: true });
    console.log("Launcher folder deployed →", targetFolder);
    return true;
  } catch (err) {
    console.error("Launcher setup failed:", err.message);
    return false;
  }
}

// ---------------- NETWORK SHARE ----------------
function runNetworkShareScript() {
  const embeddedPS = `
$basePath = "C:\\ProgramData\\myslice\\mysliceLTS"
$folderPath = "$basePath\\manifest"
$shareName = "mysliceLTS"

if (-not (Test-Path $basePath)) { New-Item -Path $basePath -ItemType Directory | Out-Null }
if (-not (Test-Path $folderPath)) { New-Item -Path $folderPath -ItemType Directory | Out-Null }

if (Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue) {
    Remove-SmbShare -Name $shareName -Force
}

New-SmbShare -Name $shareName -Path $folderPath -FullAccess "Everyone" -Description "MySlice LTS Share" | Out-Null

$acl = Get-Acl $folderPath
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone","FullControl","ContainerInherit,ObjectInherit","None","Allow")
$acl.SetAccessRule($rule)
Set-Acl $folderPath $acl
`;

  const tempPsFile = path.join(os.tmpdir(), `MySlice_Share_${Date.now()}.ps1`);

  try {
    fs.writeFileSync(tempPsFile, "\uFEFF" + embeddedPS, { encoding: "utf8" });
    execFileSync(
      "powershell.exe",
      [
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-WindowStyle",
        "Hidden",
        "-File",
        tempPsFile,
      ],
      {
        stdio: "inherit",
        windowsHide: true,
      }
    );
    console.log("Network share setup completed.");
    return true;
  } catch (err) {
    console.error("Network share failed:", err.message || err);
    return false;
  } finally {
    if (fs.existsSync(tempPsFile)) fs.unlinkSync(tempPsFile);
  }
}

// ---------------- OFFICE TRUSTED CATALOG ----------------
function addRegistryTrustedCatalog() {
  const embeddedPS = `
$desktopName = $env:COMPUTERNAME

$regContent = @"
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\\\\Software\\\\Microsoft\\\\Office\\\\16.0\\\\WEF\\\\TrustedCatalogs\\\\{c77550fc-0d50-495e-be1a-8695539e5d54}]
\\"Id\\"=\\"{c77550fc-0d50-495e-be1a-8695539e5d54}\\"
\\"Url\\"=\\"\\\\\\\\$desktopName\\\\mysliceLTS\\"
\\"Flags\\"=dword:00000001
"@

$regFilePath = "$env:TEMP\\MySlice_Trusted.reg"
Set-Content -Path $regFilePath -Value $regContent -Encoding Unicode
Start-Process regedit.exe -ArgumentList "/s", $regFilePath -Wait
Remove-Item $regFilePath -Force
`;

  return runTempPS(embeddedPS, "Trusted catalog failed");
}

// ---------------- CUSTOM PROTOCOL ----------------
function addCustomUrlProtocol() {
  const exePath = MYSLICE_EXE_TARGET;

  const embeddedPS = `
$baseKey = "Registry::HKEY_CLASSES_ROOT\\mysliceLTS"
if (!(Test-Path $baseKey)) { New-Item -Path $baseKey -Force | Out-Null }

Set-Item -Path $baseKey -Value "URL:MySlice LTS Protocol"
New-ItemProperty -Path $baseKey -Name "URL Protocol" -Value "" -PropertyType String -Force | Out-Null

$iconKey = "$baseKey\\DefaultIcon"
if (!(Test-Path $iconKey)) { New-Item -Path $iconKey -Force | Out-Null }
Set-Item -Path $iconKey -Value '"${exePath}",0'

$commandKey = "$baseKey\\shell\\open\\command"
if (!(Test-Path $commandKey)) { New-Item -Path $commandKey -Force | Out-Null }
Set-Item -Path $commandKey -Value '"${exePath}" "%1"'
`;

  return runTempPS(embeddedPS, "Protocol registration failed");
}

// ---------------- UTIL ----------------
function runTempPS(script, errorMsg) {
  const file = path.join(os.tmpdir(), `myslice_${Date.now()}.ps1`);
  try {
    fs.writeFileSync(file, script, "utf8");
    execFileSync(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", file],
      {
        windowsHide: true,
      }
    );
    return true;
  } catch (err) {
    console.error(errorMsg, err);
    return false;
  } finally {
    if (fs.existsSync(file)) fs.unlinkSync(file);
  }
}

// ---------------- APP ENTRY ----------------
app.whenReady().then(() => {
  console.log("MySlice LTS Installer – Starting setup...\n");

  const exeOk = setupMYSliceExe();
  const protocolOk = exeOk && addCustomUrlProtocol();
  const manifestOk = setupManifest();
  const shareOk = runNetworkShareScript();
  const trustOk = addRegistryTrustedCatalog();

  const allSuccess = exeOk && protocolOk && manifestOk && shareOk && trustOk;

  const failedSteps = [];
  if (!exeOk) failedSteps.push("MySlice.exe deployment");
  if (!protocolOk) failedSteps.push("Protocol registration");
  if (!manifestOk) failedSteps.push("Manifest setup");
  if (!shareOk) failedSteps.push("Network share setup");
  if (!trustOk) failedSteps.push("Office trusted catalog");

  if (allSuccess) {
    dialog.showMessageBoxSync(null, {
      type: "info",
      title: "MySlice LTS – Setup Complete",
      message:
        "All steps completed successfully.\n\n" +
        "• Files deployed\n" +
        "• Network share ready\n" +
        "• Office trust enabled\n" +
        "• Protocol registered\n\n" +
        "Your app should now work!",
    });
  } else {
    dialog.showErrorBox(
      "Setup Issue",
      `Failed steps:\n• ${failedSteps.join(
        "\n• "
      )}\n\nPlease ensure you run the installer as Administrator.`
    );
  }

  app.quit();
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
