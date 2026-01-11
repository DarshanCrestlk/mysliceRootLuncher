param (
    [string]$url
)

# Ensure URL starts with https://
if ($url -notmatch '^https://') {
    $url = $url -replace '^https//', 'https://'
}

# Extract file_id from the URL
$fileId = ($url -split 'file_id=')[1] -split '\?end=' | Select-Object -First 1
$ext = ($url -split 'ext=')[1] -split '\?file_id=' | Select-Object -First 1
$permissions = ($url -split 'p=')[1] -split '\?ext=' | Select-Object -First 1

Write-Host "---------------"$ext
Write-Host "---------------"$permissions

# Extract the original filename from the URL and clean it up
$originalFileName = [System.IO.Path]::GetFileName($url)
$cleanedFileName = $originalFileName -replace '^\d+\.*', ''       # Remove numeric prefixes
$cleanedFileName = $cleanedFileName -replace '\..*$', ''          # Remove everything after the first period
$cleanedFileName = [System.Net.WebUtility]::UrlDecode($cleanedFileName.Trim())  # Decode URL-encoded characters and trim whitespace

# Define the temp file path with cleaned filename and .docx extension
if ($permissions) {
    $tempFile = Join-Path -Path $env:TEMP -ChildPath "$cleanedFileName.$ext($fileId+$(Get-Date -Format 'yyyyMMddHHmmss')),$permissions"
    Write-Host "------"$tempFile
} else {
    $tempFile = Join-Path -Path $env:TEMP -ChildPath "$cleanedFileName.$ext($fileId+$(Get-Date -Format 'yyyyMMddHHmmss'))"
}

# Download the file to a temporary location
Invoke-WebRequest -Uri $url -OutFile $tempFile

# Word Document Processing
if ($ext -eq "docx") {
    # Create a new Word application object
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true

    try {
        $document = $word.Documents.Open($tempFile, [ref]$false, [ref]$false, [ref]$false)
        $selection = $word.Selection
        $selection.Font.Hidden = $true
        $selection.TypeText(" ")

        Write-Host "Document is open in Word. Press Ctrl+C in the console to close the script."

        while ($document.Windows.Count -gt 0) {
            Start-Sleep -Milliseconds 100
        }
    }
    catch {
        Write-Host "An error occurred: $_"
    }
    finally {
        try {
            if ($document -ne $null -and $document.Windows.Count -gt 0) {
                $document.Close($false)
            }
        } catch {
            Write-Host "Document was already closed or disconnected."
        }

        try {
            if ($word -ne $null -and $word.Visible -eq $true) {
                $word.Quit()
            }
        } catch {
            Write-Host "Word was already closed or disconnected."
        }

        if (Test-Path $tempFile) {
            Remove-Item $tempFile -Force
        }

        exit
    }
}

# Excel Workbook Processing - FIXED VERSION
elseif ($ext -eq "xlsx") {
    # Get existing Excel processes before starting
    $existingProcesses = @()
    try {
        $existingProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id
    } catch {}
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    
    $workbook = $null
    $newProcessId = $null

    try {
        # Find the new Excel process
        Start-Sleep -Milliseconds 1000
        $allProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id
        $newProcesses = $allProcesses | Where-Object { $_ -notin $existingProcesses }
        if ($newProcesses) {
            $newProcessId = $newProcesses[0]
            Write-Host "Excel process ID: $newProcessId"
        }
        
        # Try opening the file
        if (Test-Path $tempFile) {
            $workbook = $excel.Workbooks.Open($tempFile)
            Write-Host "Excel file opened successfully. Close Excel window to continue..."
        } else {
            Write-Host "Error: Temp file not found - $tempFile"
            throw "File not found"
        }

        # Wait for Excel to be closed - simplified approach
        while ($true) {
            Start-Sleep -Milliseconds 500
            try {
                # Test if Excel COM object is still valid
                $count = $excel.Workbooks.Count
                if ($count -eq 0) {
                    Write-Host "All workbooks closed."
                    break
                }
            } catch {
                # Excel COM object is no longer valid (Excel was closed)
                Write-Host "Excel was closed."
                break
            }
        }

    } catch {
        Write-Host "Error opening Excel file: $_"
    } finally {
        # Cleanup
        Write-Host "Starting cleanup..."
        
        try {
            if ($workbook -ne $null) {
                $workbook.Close($false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            }
        } catch {
            Write-Host "Workbook cleanup error: $_"
        }

        try {
            if ($excel -ne $null) {
                $excel.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
        } catch {
            Write-Host "Excel cleanup error: $_"
        }

        # Force garbage collection
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        
        # Kill the Excel process if still running
        if ($newProcessId) {
            Start-Sleep -Milliseconds 1000
            try {
                $process = Get-Process -Id $newProcessId -ErrorAction SilentlyContinue
                if ($process -and !$process.HasExited) {
                    Write-Host "Force killing Excel process $newProcessId"
                    Stop-Process -Id $newProcessId -Force
                }
            } catch {
                Write-Host "Process cleanup: $_"
            }
        }

        # Clean up temp file
        if (Test-Path $tempFile) {
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
    
    Write-Host "Excel processing completed."
    exit
}

# PowerPoint Presentation Processing
elseif ($ext -eq "pptx") {
    $powerpoint = New-Object -ComObject PowerPoint.Application
    $powerpoint.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

    try {
        $presentation = $powerpoint.Presentations.Open($tempFile, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoTrue)

        Write-Host "Presentation is open in PowerPoint. Press Ctrl+C in the console to close the script."

        while ($presentation.Windows.Count -gt 0) {
            Start-Sleep -Milliseconds 100
        }
    }
    catch {
        Write-Host "An error occurred: $_"
    }
    finally {
        if ($presentation -ne $null) { $presentation.Close() }
        if ($powerpoint -ne $null) { $powerpoint.Quit() }
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force }
        exit
    }
}
else {
    Write-Host "Unsupported file type: $ext"
}