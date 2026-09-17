$ErrorActionPreference = "Continue"
$failed = New-Object System.Collections.Generic.List[string]

$baseDir = Join-Path $env:PROGRAMDATA "myslice\mysliceLTS"
$mysliceRoot = Join-Path $env:PROGRAMDATA "myslice"
$catalogGuid = "c77550fc-0d50-495e-be1a-8695539e5d54"
$addInId = "7ac86ae0-404b-43c2-b9d9-e6c178dc4b94"

function Write-Step($name, [scriptblock]$action) {
    try {
        $ErrorActionPreference = "Stop"
        & $action
        Write-Host "OK  $name"
        return $true
    }
    catch {
        Write-Host "FAIL  $name : $($_.Exception.Message)"
        [void]$failed.Add($name)
        return $false
    }
}

Write-Host "MySlice LTS uninstall starting..."

Write-Host "Stopping launcher..."
cmd /c "taskkill /F /IM mysliceLTS.exe /T >nul 2>&1"
Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match "powershell" -and $_.CommandLine -and ($_.CommandLine -like "*myslice.ps1*") } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }

Write-Step "Network share" {
    $share = Get-SmbShare -Name "mysliceLTS" -ErrorAction SilentlyContinue
    if ($share) {
        Remove-SmbShare -Name "mysliceLTS" -Force
    }
}

Write-Step "Protocol mysliceLTS://" {
    $baseKey = "Registry::HKEY_CLASSES_ROOT\mysliceLTS"
    if (Test-Path $baseKey) {
        Remove-Item -Path $baseKey -Recurse -Force
    }
}

Write-Step "Office trusted catalog" {
    $paths = @(
        "HKCU:\Software\Policies\Microsoft\Office\16.0\WEF\TrustedCatalogs\{$catalogGuid}",
        "HKLM:\Software\Policies\Microsoft\Office\16.0\WEF\TrustedCatalogs\{$catalogGuid}",
        "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{$catalogGuid}",
        "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\$addInId"
    )
    foreach ($p in $paths) {
        if (Test-Path $p) {
            Remove-Item -Path $p -Recurse -Force
        }
    }
}

Write-Step "ProgramData files" {
    if (Test-Path $baseDir) {
        Remove-Item -LiteralPath $baseDir -Recurse -Force
    }
    if ((Test-Path $mysliceRoot) -and (@(Get-ChildItem -LiteralPath $mysliceRoot -Force -ErrorAction SilentlyContinue).Count -eq 0)) {
        Remove-Item -LiteralPath $mysliceRoot -Force -ErrorAction SilentlyContinue
    }
}

if ($failed.Count -gt 0) {
    Write-Host "Failed steps:"
    $failed | ForEach-Object { Write-Host " - $_" }
    Write-Host "Run uninstall as Administrator and close Word/Excel."
    exit 1
}

Write-Host "MySlice LTS uninstall complete."
exit 0
