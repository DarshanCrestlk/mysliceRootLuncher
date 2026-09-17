param(
    [Parameter(Mandatory = $true)]
    [string]$ExePath,

    [Parameter(Mandatory = $true)]
    [string]$ManifestSource,

    [string]$LauncherSource = ""
)

$ErrorActionPreference = "Stop"
$failed = New-Object System.Collections.Generic.List[string]
$script:ExePath = $ExePath

$baseDir = Join-Path $env:PROGRAMDATA "myslice\mysliceLTS"
$manifestDir = Join-Path $baseDir "manifest"
$manifestTarget = Join-Path $manifestDir "manifest.xml"
$launcherDir = Join-Path $baseDir "launcher"

function Write-Step($name, [scriptblock]$action) {
    try {
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

Write-Host "MySlice LTS setup starting..."

if ($LauncherSource) {
    Write-Step "MySlice.exe deployment" {
        if (-not (Test-Path $LauncherSource)) {
            throw "Source launcher folder not found: $LauncherSource"
        }
        New-Item -ItemType Directory -Force -Path $launcherDir | Out-Null
        Copy-Item -Path (Join-Path $LauncherSource "*") -Destination $launcherDir -Recurse -Force
    }
    $script:ExePath = Join-Path $launcherDir "mysliceLTS.exe"
}

Write-Step "Manifest setup" {
    if (-not (Test-Path $ManifestSource)) {
        throw "Source manifest.xml not found: $ManifestSource"
    }
    New-Item -ItemType Directory -Force -Path $manifestDir | Out-Null
    Copy-Item -LiteralPath $ManifestSource -Destination $manifestTarget -Force
}

$exeOk = Write-Step "Launcher path" {
    if (-not (Test-Path $script:ExePath)) {
        throw "mysliceLTS.exe not found: $script:ExePath"
    }
}

if ($exeOk) {
    Write-Step "Protocol registration" {
        $baseKey = "Registry::HKEY_CLASSES_ROOT\mysliceLTS"
        if (-not (Test-Path $baseKey)) {
            New-Item -Path $baseKey -Force | Out-Null
        }
        Set-Item -Path $baseKey -Value "URL:MySlice LTS Protocol"
        New-ItemProperty -Path $baseKey -Name "URL Protocol" -Value "" -PropertyType String -Force | Out-Null

        $iconKey = "$baseKey\DefaultIcon"
        if (-not (Test-Path $iconKey)) {
            New-Item -Path $iconKey -Force | Out-Null
        }
        Set-Item -Path $iconKey -Value "`"$($script:ExePath)`",0"

        $commandKey = "$baseKey\shell\open\command"
        if (-not (Test-Path $commandKey)) {
            New-Item -Path $commandKey -Force | Out-Null
        }
        Set-Item -Path $commandKey -Value "`"$($script:ExePath)`" `"%1`""
    }
}

Write-Step "Network share setup" {
    $shareName = "mysliceLTS"
    if (-not (Test-Path $baseDir)) {
        New-Item -ItemType Directory -Force -Path $baseDir | Out-Null
    }
    if (-not (Test-Path $manifestDir)) {
        New-Item -ItemType Directory -Force -Path $manifestDir | Out-Null
    }

    if (Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue) {
        Remove-SmbShare -Name $shareName -Force
    }

    New-SmbShare -Name $shareName -Path $manifestDir -FullAccess "Everyone" -Description "MySlice LTS Share" | Out-Null

    $acl = Get-Acl $manifestDir
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
        "Everyone",
        "FullControl",
        "ContainerInherit,ObjectInherit",
        "None",
        "Allow"
    )
    $acl.SetAccessRule($rule)
    Set-Acl $manifestDir $acl
}

Write-Step "Office trusted catalog" {
    $desktopName = $env:COMPUTERNAME
    $shareName = "mysliceLTS"
    $guid = "c77550fc-0d50-495e-be1a-8695539e5d54"
    $addInId = "7ac86ae0-404b-43c2-b9d9-e6c178dc4b94"
    $regContent = @"
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\16.0\WEF\TrustedCatalogs\{$guid}]
"Id"="{$guid}"
"Url"="\\\\$desktopName\$shareName"
"Flags"=dword:00000001

[HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\$addInId]
"UseDirectDebugger"=dword:00000000
"UseWebDebugger"=dword:00000000
"@
    $regFilePath = Join-Path $env:TEMP "MySlice_Trusted_Catalog.reg"
    $regContent | Out-File -FilePath $regFilePath -Encoding Unicode -Force
    try {
        Start-Process regedit.exe -ArgumentList "/s", "`"$regFilePath`"" -Wait -NoNewWindow
    }
    finally {
        Remove-Item $regFilePath -Force -ErrorAction SilentlyContinue
    }
}

Write-Step "Office add-in debugger off" {
    $addInId = "7ac86ae0-404b-43c2-b9d9-e6c178dc4b94"
    $key = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\$addInId"
    if (-not (Test-Path $key)) {
        New-Item -Path $key -Force | Out-Null
    }
    New-ItemProperty -Path $key -Name "UseDirectDebugger" -Value 0 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $key -Name "UseWebDebugger" -Value 0 -PropertyType DWord -Force | Out-Null
}

if ($failed.Count -gt 0) {
    Write-Host "Failed steps:"
    $failed | ForEach-Object { Write-Host " - $_" }
    Write-Host "Run the installer as Administrator."
    exit 1
}

Write-Host "MySlice LTS setup complete."
exit 0
