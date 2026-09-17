param (
    [string]$url
)

function Disable-OfficeAddinWebViewDebugger {
    $addInId = "7ac86ae0-404b-43c2-b9d9-e6c178dc4b94"
    $key = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\$addInId"
    try {
        if (-not (Test-Path $key)) {
            New-Item -Path $key -Force | Out-Null
        }
        New-ItemProperty -Path $key -Name "UseDirectDebugger" -Value 0 -PropertyType DWord -Force | Out-Null
        New-ItemProperty -Path $key -Name "UseWebDebugger" -Value 0 -PropertyType DWord -Force | Out-Null
    } catch {
        Write-Host "Could not disable Office add-in debugger: $_"
    }
}

Disable-OfficeAddinWebViewDebugger

function Get-QueryValue {
    param(
        [string]$RawUrl,
        [string]$Key
    )
    if ([string]::IsNullOrWhiteSpace($RawUrl)) { return $null }
    $match = [regex]::Match($RawUrl, "[?&]$Key=([^?&]*)")
    if (-not $match.Success) { return $null }
    return [System.Net.WebUtility]::UrlDecode($match.Groups[1].Value.Trim())
}

function Get-DownloadUrl {
    param([string]$RawUrl)
    $download = $RawUrl
    foreach ($key in @("ext", "file_id", "origin", "mode", "p")) {
        $download = [regex]::Replace($download, "[?&]$key=[^?&]*", "")
    }
    return $download.TrimEnd("?", "&")
}

function Get-SafeBaseName {
    param([string]$DownloadUrl)
    $pathPart = $DownloadUrl
    try {
        $uri = [Uri]$DownloadUrl
        if ($uri.IsAbsoluteUri) { $pathPart = $uri.AbsolutePath }
    } catch {}

    $leaf = [System.IO.Path]::GetFileNameWithoutExtension($pathPart)
    if ($leaf) { $leaf = [System.Net.WebUtility]::UrlDecode($leaf) }
    $leaf = ($leaf -replace "^\d+\.*", "").Trim()
    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) {
        if ($null -ne $leaf) { $leaf = $leaf.Replace([string]$ch, "") }
    }
    $leaf = $leaf -replace "\s+", "_"
    if ([string]::IsNullOrWhiteSpace($leaf)) { $leaf = "document" }
    if ($leaf.Length -gt 50) { $leaf = $leaf.Substring(0, 50) }
    return $leaf
}

function Get-SafeSlug {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $slug = $Value.Trim().ToLower()
    if ($slug -notmatch '^[a-z0-9]+(-[a-z0-9]+)*$') { return $null }
    return $slug
}

function Get-NewProcessId {
    param(
        [string]$Name,
        [int[]]$ExistingIds
    )
    Start-Sleep -Milliseconds 800
    try {
        $all = @(Get-Process -Name $Name -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $new = @($all | Where-Object { $ExistingIds -notcontains $_ })
        if ($new.Count -gt 0) { return $new[0] }
    } catch {}
    return $null
}

function Remove-TempFileSafe {
    param([string]$Path)
    if ($Path -and (Test-Path $Path)) {
        Remove-Item $Path -Force -ErrorAction SilentlyContinue
    }
}

if ($url -notmatch '^https://') {
    $url = $url -replace '^https//', 'https://'
}

$fileId = Get-QueryValue -RawUrl $url -Key 'file_id'
$ext = Get-QueryValue -RawUrl $url -Key 'ext'
$permissions = Get-QueryValue -RawUrl $url -Key 'p'
$origin = Get-SafeSlug (Get-QueryValue -RawUrl $url -Key 'origin')
$downloadUrl = Get-DownloadUrl $url
$cleanedFileName = Get-SafeBaseName $downloadUrl

$meta = "$fileId+$(Get-Date -Format 'yyyyMMddHHmmss')"
if ($origin) { $meta = "$meta+$origin" }

if ($permissions) {
    $tempFile = Join-Path -Path $env:TEMP -ChildPath "$cleanedFileName.$ext($meta),$permissions"
} else {
    $tempFile = Join-Path -Path $env:TEMP -ChildPath "$cleanedFileName.$ext($meta)"
}

Write-Host "ext=$ext origin=$origin"
Write-Host "downloadUrl=$downloadUrl"
Write-Host "tempFile=$tempFile"

Invoke-WebRequest -Uri $downloadUrl -OutFile $tempFile -UseBasicParsing

if ($ext -eq "docx") {
    $existingIds = @()
    try { $existingIds = @(Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) } catch {}

    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $wordPid = Get-NewProcessId -Name "WINWORD" -ExistingIds $existingIds
    $document = $null

    try {
        $document = $word.Documents.Open($tempFile, [ref]$false, [ref]$false, [ref]$false)
        $selection = $word.Selection
        $selection.Font.Hidden = $true
        $selection.TypeText(" ")

        while ($true) {
            Start-Sleep -Milliseconds 200
            try {
                if ($document.Windows.Count -eq 0) { break }
            } catch {
                break
            }
        }
    } catch {
        Write-Host "Word error: $_"
    } finally {
        try {
            if ($null -ne $document) {
                try {
                    if ($document.Windows.Count -gt 0) { $document.Close($false) }
                } catch {}
            }
        } catch {}

        $shouldQuit = $false
        try {
            if ($null -ne $word -and $word.Documents.Count -eq 0) { $shouldQuit = $true }
        } catch {
            $shouldQuit = $true
        }

        if ($shouldQuit) {
            try { if ($null -ne $word) { $word.Quit() } } catch {}
        } else {
            Write-Host "Word still has other documents; not quitting this instance. pid=$wordPid"
        }

        Remove-TempFileSafe $tempFile
        exit
    }
}

elseif ($ext -eq "xlsx") {
    $existingIds = @()
    try { $existingIds = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) } catch {}

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excelPid = Get-NewProcessId -Name "EXCEL" -ExistingIds $existingIds
    $workbook = $null

    try {
        if (-not (Test-Path $tempFile)) { throw "Temp file not found - $tempFile" }
        $workbook = $excel.Workbooks.Open($tempFile)
        Write-Host "Excel opened pid=$excelPid"

        while ($true) {
            Start-Sleep -Milliseconds 400
            try {
                $null = $workbook.Name
                if ($workbook.Windows.Count -eq 0) { break }
            } catch {
                break
            }
        }
    } catch {
        Write-Host "Excel error: $_"
    } finally {
        try {
            if ($null -ne $workbook) {
                try { $workbook.Close($false) } catch {}
                try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
            }
        } catch {}

        $empty = $false
        try {
            if ($null -eq $excel -or $excel.Workbooks.Count -eq 0) { $empty = $true }
        } catch {
            $empty = $true
        }

        if ($empty) {
            try {
                if ($null -ne $excel) {
                    $excel.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                }
            } catch {}
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()

            if ($excelPid) {
                Start-Sleep -Milliseconds 800
                try {
                    $proc = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
                    if ($proc -and -not $proc.HasExited) {
                        Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
                    }
                } catch {}
            }
        } else {
            Write-Host "Excel still has other workbooks; not quitting. pid=$excelPid"
        }

        Remove-TempFileSafe $tempFile
        exit
    }
}

elseif ($ext -eq "pptx") {
    $existingIds = @()
    try { $existingIds = @(Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) } catch {}

    $powerpoint = New-Object -ComObject PowerPoint.Application
    $powerpoint.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $pptPid = Get-NewProcessId -Name "POWERPNT" -ExistingIds $existingIds
    $presentation = $null

    try {
        $presentation = $powerpoint.Presentations.Open($tempFile, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoTrue)
        while ($true) {
            Start-Sleep -Milliseconds 200
            try {
                if ($presentation.Windows.Count -eq 0) { break }
            } catch { break }
        }
    } catch {
        Write-Host "PowerPoint error: $_"
    } finally {
        try { if ($null -ne $presentation) { $presentation.Close() } } catch {}

        $empty = $false
        try {
            if ($null -eq $powerpoint -or $powerpoint.Presentations.Count -eq 0) { $empty = $true }
        } catch { $empty = $true }

        if ($empty) {
            try { if ($null -ne $powerpoint) { $powerpoint.Quit() } } catch {}
        } else {
            Write-Host "PowerPoint still has other presentations; not quitting. pid=$pptPid"
        }

        Remove-TempFileSafe $tempFile
        exit
    }
}
else {
    Write-Host "Unsupported file type: $ext"
    Remove-TempFileSafe $tempFile
}
