# Monday 9:30 AM Forecast Data Sync Automation (Improved with Debug Features)
param(
    [switch]$Debug = $false
)
$ErrorActionPreference = "Stop"

# Paths
$downloadsPath = "$env:USERPROFILE\Downloads"
$forecastingPath = "C:\Users\usuario\OneDrive - talentorecruiting.com\Forecasting"
$dataPath = "$forecastingPath\Data"
$oldPath = "$forecastingPath\Old"
$today = Get-Date -Format "yyyy-MM-dd"
$forecastToolFile = "$forecastingPath\Forecast Tool$today.xlsx"
$logFile = "$forecastingPath\sync_log_$today.txt"

# Ensure folders exist
if (-not (Test-Path $dataPath)) { New-Item -ItemType Directory -Path $dataPath -Force | Out-Null }
if (-not (Test-Path $oldPath)) { New-Item -ItemType Directory -Path $oldPath -Force | Out-Null }

function Write-Log {
    param([string]$message, [switch]$Error = $false)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    $logMessage | Tee-Object -FilePath $logFile -Append
    if ($Error) {
        Write-Host $logMessage -ForegroundColor Red
    } else {
        Write-Host $logMessage
    }
}

function Find-FileWithDebug {
    param([string]$path, [string]$filter, [string]$description)
    Write-Log "Searching for $description in: $path"
    if (-not (Test-Path $path)) {
        Write-Log "ERROR: Path does not exist: $path" -Error
        return $null
    }

    $files = Get-ChildItem $path -Filter $filter 2>$null | Sort-Object LastWriteTime -Descending
    if ($files.Count -gt 0) {
        Write-Log "Found $($files.Count) file(s) matching '$filter'"
        if ($Debug) {
            $files | ForEach-Object { Write-Log "  - $($_.Name) (Modified: $($_.LastWriteTime))" }
        }
        return $files[0]
    } else {
        Write-Log "WARNING: No files found matching '$filter'" -Error
        if ($Debug) {
            Write-Log "Files in $path:"
            Get-ChildItem $path 2>$null | Select-Object -First 10 | ForEach-Object {
                Write-Log "  - $($_.Name)"
            }
        }
        return $null
    }
}

Write-Log "===== STARTING FORECAST DATA SYNC ====="
if ($Debug) { Write-Log "DEBUG MODE ENABLED" }

$excel = $null
try {
    # ================= STEP 0 =================
    Write-Log "STEP 0 - Preparing Forecast Tool..."

    # Find, backup, and rename Forecast Tool file
    if (-not (Test-Path $forecastingPath)) {
        Write-Log "ERROR: Forecasting path does not exist: $forecastingPath" -Error
        exit 1
    }

    $existingFiles = Get-ChildItem $forecastingPath -Filter "Forecast Tool*.xlsx" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending

    if ($existingFiles.Count -eq 0) {
        Write-Log "ERROR: No Forecast Tool file found in $forecastingPath" -Error
        exit 1
    }

    $sourceFile = $existingFiles[0].FullName
    $sourceFileName = $existingFiles[0].Name

    # Create backup with timestamp
    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $backupName = $sourceFileName -replace "\.xlsx$", "_backup_$timestamp.xlsx"
    $backupPath = "$oldPath\$backupName"
    Copy-Item $sourceFile $backupPath -Force
    Write-Log "Backup created: $backupName"

    # Rename to today's date if needed
    if ($sourceFile -ne $forecastToolFile) {
        Rename-Item $sourceFile "Forecast Tool$today.xlsx" -Force
        Write-Log "File renamed to: Forecast Tool$today.xlsx"
    }

    # Open Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    Write-Log "Excel application opened"

    # ================= STEP 1 =================
    Write-Log "STEP 1 - Processing Export file..."
    $exportFile = Find-FileWithDebug $downloadsPath "export*.csv" "Export file"

    if ($exportFile) {
        Write-Log "Opening: $($exportFile.Name)"
        $exportWB = $excel.Workbooks.Open($exportFile.FullName, 0)
        $exportWS = $exportWB.ActiveSheet
        Write-Log "Deleting header row and processing data..."
        $exportWS.Rows(1).Delete()
        $usedRange = $exportWS.UsedRange
        $values = $usedRange.Value2

        $savedPath = "$dataPath\$($exportFile.BaseName)-processed.xlsx"
        $exportWB.SaveAs($savedPath, 51)
        $exportWB.Close($false)
        Write-Log "Export file processed and saved to Data folder"

        Write-Log "Updating Forecast Tool with export data..."
        $forecastWB = $excel.Workbooks.Open($forecastToolFile, 0)
        $forecastDataWS = $forecastWB.Sheets("Forecast data")
        $forecastDataWS.Cells.Clear()

        if ($values -ne $null) {
            Write-Log "Pasting data ($($usedRange.Rows.Count) rows, $($usedRange.Columns.Count) columns)"
            $forecastDataWS.Range("A1").Resize(
                $usedRange.Rows.Count,
                $usedRange.Columns.Count
            ).Value2 = $values
        }
        $forecastWB.Save()
        $forecastWB.Close()
    }
    else {
        Write-Log "No export file found - skipping export processing"
    }

    # ================= STEP 2 =================
    Write-Log "STEP 2 - Processing Placement Activity Report..."
    $placementFile = Find-FileWithDebug $downloadsPath "Placement Activity Report*" "Placement Activity Report"

    if ($placementFile) {
        Write-Log "Opening: $($placementFile.Name)"
        $placementWB = $excel.Workbooks.Open($placementFile.FullName, 0)
        $placementWS = $placementWB.ActiveSheet
        Write-Log "Deleting header row and column B..."
        $placementWS.Rows(1).Delete()
        $placementWS.Columns("B").Delete()
        $usedRange = $placementWS.UsedRange
        $values = $usedRange.Value2

        $savedPath = "$dataPath\$($placementFile.BaseName)-processed.xlsx"
        $placementWB.SaveAs($savedPath, 51)
        $placementWB.Close($false)
        Write-Log "Placement file processed and saved to Data folder"

        Write-Log "Updating Forecast Tool with placement data..."
        $forecastWB = $excel.Workbooks.Open($forecastToolFile, 0)
        $placementDataWS = $forecastWB.Sheets("Placement Data")
        $placementDataWS.Range("B:ZZ").Clear()

        if ($values -ne $null) {
            Write-Log "Pasting data ($($usedRange.Rows.Count) rows, $($usedRange.Columns.Count) columns)"
            $placementDataWS.Range("B1").Resize(
                $usedRange.Rows.Count,
                $usedRange.Columns.Count
            ).Value2 = $values
        }
        $forecastWB.Save()
        $forecastWB.Close()
    }
    else {
        Write-Log "No Placement file found - skipping placement processing"
    }

    # ================= STEP 3 & 4 =================
    Write-Log "STEP 3 & 4 - Running Sync Logic..."
    $forecastWB = $excel.Workbooks.Open($forecastToolFile, 0)

    # ---------- STEP 3 ----------
    Write-Log "STEP 3 - Syncing Forecast data..."
    $forecastDataWS = $forecastWB.Sheets("Forecast data")
    $forecastWS = $forecastWB.Sheets("Forecast")
    $dataLastRow = $forecastDataWS.UsedRange.Rows.Count
    $forecastLastRow = $forecastWS.UsedRange.Rows.Count

    $dataValues = @{}
    $dataRows = @{}
    for ($i = 2; $i -le $dataLastRow; $i++) {
        $value = $forecastDataWS.Cells($i, 1).Value2
        if ($value) {
            $key = ([string]$value).Trim()
            if ($key -ne "") {
                $dataValues[$key] = $true
                $row = @()
                $lastCol = $forecastDataWS.Cells($i, $forecastDataWS.Columns.Count).End(-4159).Column
                for ($j = 1; $j -le $lastCol; $j++) {
                    $row += $forecastDataWS.Cells($i, $j).Value2
                }
                $dataRows[$key] = $row
            }
        }
    }

    $forecastValues = @{}
    for ($i = 5; $i -le $forecastLastRow; $i++) {
        $value = $forecastWS.Cells($i, 3).Value2
        if ($value) {
            $forecastValues[([string]$value).Trim()] = $i
        }
    }

    $nextRow = 5
    while ($forecastWS.Cells($nextRow, 3).Value2) { $nextRow++ }

    $newRecordsAdded = 0
    foreach ($key in $dataValues.Keys) {
        if (-not $forecastValues.ContainsKey($key)) {
            $rowData = $dataRows[$key]
            for ($j = 0; $j -lt $rowData.Count; $j++) {
                $forecastWS.Cells($nextRow, $j + 3).Value2 = $rowData[$j]
            }
            $nextRow++
            $newRecordsAdded++
        }
    }
    Write-Log "Added $newRecordsAdded new records to Forecast"

    $rowsHighlighted = 0
    for ($i = 5; $i -le $forecastLastRow; $i++) {
        $value = $forecastWS.Cells($i, 3).Value2
        if ($value -and -not $dataValues.ContainsKey(([string]$value).Trim())) {
            $forecastWS.Rows($i).Interior.Color = 65535  # Yellow
            $rowsHighlighted++
        }
    }
    Write-Log "Highlighted $rowsHighlighted inactive records (yellow)"

    # ---------- STEP 4 ----------
    Write-Log "STEP 4 - Syncing Placement data..."
    $placementDataWS = $forecastWB.Sheets("Placement Data")
    $placedWS = $forecastWB.Sheets("Placed")
    $placementLastRow = $placementDataWS.UsedRange.Rows.Count
    $placedLastRow = $placedWS.UsedRange.Rows.Count

    $placementValues = @{}
    $placementRows = @{}
    for ($i = 2; $i -le $placementLastRow; $i++) {
        $value = $placementDataWS.Cells($i, 3).Value2
        if ($value) {
            $key = ([string]$value).Trim()
            $placementValues[$key] = $true
            $row = @()
            $lastCol = $placementDataWS.Cells($i, $placementDataWS.Columns.Count).End(-4159).Column
            for ($j = 1; $j -le $lastCol; $j++) {
                $row += $placementDataWS.Cells($i, $j).Value2
            }
            $placementRows[$key] = $row
        }
    }

    $placedValues = @{}
    for ($i = 4; $i -le $placedLastRow; $i++) {
        $value = $placedWS.Cells($i, 3).Value2
        if ($value) {
            $placedValues[([string]$value).Trim()] = $i
        }
    }

    $nextPlacedRow = 4
    while ($placedWS.Cells($nextPlacedRow, 3).Value2) { $nextPlacedRow++ }

    $placedRecordsAdded = 0
    foreach ($key in $placementValues.Keys) {
        if (-not $placedValues.ContainsKey($key)) {
            $rowData = $placementRows[$key]
            for ($j = 0; $j -lt $rowData.Count; $j++) {
                $placedWS.Cells($nextPlacedRow, $j + 1).Value2 = $rowData[$j]
            }
            $nextPlacedRow++
            $placedRecordsAdded++
        }
    }
    Write-Log "Added $placedRecordsAdded new records to Placed"

    $forecastWB.Save()
    $forecastWB.Close()
    Write-Log "Forecast Tool saved and closed"

    # Cleanup
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Log "===== PROCESS COMPLETED SUCCESSFULLY ====="
    exit 0
}
catch {
    Write-Log "ERROR: $($_.Exception.Message)" -Error
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Error
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    exit 1
}
