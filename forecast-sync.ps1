# Monday 9:30 AM Forecast Data Sync Automation
param()
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
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $message" | Tee-Object -FilePath $logFile -Append
}
Write-Log "===== STARTING FORECAST DATA SYNC ====="
$excel = $null
try {
    # ================= STEP 0 =================
    Write-Log "STEP 0 - Preparing Forecast Tool..."
    $existingFiles = Get-ChildItem $forecastingPath -Filter "Forecast Tool*.xlsx" |
        Sort-Object LastWriteTime -Descending
    if ($existingFiles.Count -eq 0) {
        Write-Log "ERROR: No Forecast Tool file found."
        exit 1
    }
    $sourceFile = $existingFiles[0].FullName
    $sourceFileName = $existingFiles[0].Name
    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $backupName = $sourceFileName -replace "\.xlsx$", "_backup_$timestamp.xlsx"
    Copy-Item $sourceFile "$oldPath\$backupName" -Force
    Write-Log "Backup created."
    if ($sourceFile -ne $forecastToolFile) {
        Rename-Item $sourceFile "Forecast Tool$today.xlsx" -Force
        Write-Log "File renamed to today's date."
    }
    # Open Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    # ================= STEP 1 =================
    Write-Log "STEP 1 - Processing Export file..."
    $exportFile = Get-ChildItem $downloadsPath -Filter "export*.csv" |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($exportFile) {
        $exportWB = $excel.Workbooks.Open($exportFile.FullName, 0)
        $exportWS = $exportWB.ActiveSheet
        $exportWS.Rows(1).Delete()
        $usedRange = $exportWS.UsedRange
        $values = $usedRange.Value2
        $savedPath = "$dataPath\$($exportFile.BaseName)-processed.xlsx"
        $exportWB.SaveAs($savedPath, 51)
        $exportWB.Close($false)
        Write-Log "Export file processed."
        $forecastWB = $excel.Workbooks.Open($forecastToolFile, 0)
        $forecastDataWS = $forecastWB.Sheets("Forecast data")
        $forecastDataWS.Cells.Clear()
        if ($values -ne $null) {
            $forecastDataWS.Range("A1").Resize(
                $usedRange.Rows.Count,
                $usedRange.Columns.Count
            ).Value2 = $values
        }
        $forecastWB.Save()
        $forecastWB.Close()
    }
    else {
        Write-Log "No export file found."
    }
    # ================= STEP 2 =================
    Write-Log "STEP 2 - Processing Placement Activity Report..."
    $placementFile = Get-ChildItem $downloadsPath -Filter "Placement Activity Report*" |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($placementFile) {
        $placementWB = $excel.Workbooks.Open($placementFile.FullName, 0)
        $placementWS = $placementWB.ActiveSheet
        $placementWS.Rows(1).Delete()
        $placementWS.Columns("B").Delete()
        $usedRange = $placementWS.UsedRange
        $values = $usedRange.Value2
        $savedPath = "$dataPath\$($placementFile.BaseName)-processed.xlsx"
        $placementWB.SaveAs($savedPath, 51)
        $placementWB.Close($false)
        Write-Log "Placement file processed."
        $forecastWB = $excel.Workbooks.Open($forecastToolFile, 0)
        $placementDataWS = $forecastWB.Sheets("Placement Data")
        $placementDataWS.Range("B:ZZ").Clear()
        if ($values -ne $null) {
            $placementDataWS.Range("B1").Resize(
                $usedRange.Rows.Count,
                $usedRange.Columns.Count
            ).Value2 = $values
        }
        $forecastWB.Save()
        $forecastWB.Close()
    }
    else {
        Write-Log "No Placement file found."
    }
    # ================= STEP 3 & 4 =================
    Write-Log "STEP 3 & 4 - Running Sync Logic..."
    $forecastWB = $excel.Workbooks.Open($forecastToolFile, 0)
    # ---------- STEP 3 ----------
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
    foreach ($key in $dataValues.Keys) {
        if (-not $forecastValues.ContainsKey($key)) {
            $rowData = $dataRows[$key]
            for ($j = 0; $j -lt $rowData.Count; $j++) {
                $forecastWS.Cells($nextRow, $j + 3).Value2 = $rowData[$j]
            }
            $nextRow++
        }
    }
    for ($i = 5; $i -le $forecastLastRow; $i++) {
        $value = $forecastWS.Cells($i, 3).Value2
        if ($value -and -not $dataValues.ContainsKey(([string]$value).Trim())) {
            $forecastWS.Rows($i).Interior.Color = 65535
        }
    }
    # ---------- STEP 4 ----------
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
    foreach ($key in $placementValues.Keys) {
        if (-not $placedValues.ContainsKey($key)) {
            $rowData = $placementRows[$key]
            for ($j = 0; $j -lt $rowData.Count; $j++) {
                $placedWS.Cells($nextPlacedRow, $j + 1).Value2 = $rowData[$j]
            }
            $nextPlacedRow++
        }
    }
    for ($i = 4; $i -le $placedLastRow; $i++) {
        $value = $placedWS.Cells($i, 3).Value2
        if ($value -and -not $placementValues.ContainsKey(([string]$value).Trim())) {
            $placedWS.Rows($i).Interior.Color = 65535
        }
    }
    $forecastWB.Save()
    $forecastWB.Close()
    # Cleanup
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Log "===== PROCESS COMPLETED SUCCESSFULLY ====="
    exit 0
}
catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    exit 1
}
