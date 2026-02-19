# Forecast Update Automation Script
# Runs Monday at 9:30 AM to process and update forecast files

# Configuration
$downloadsPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads")
$outputPath = "C:\Users\usuario\OneDrive - talentorecruiting.com\Forecasting\Data"
$forecastToolPath = "C:\Users\usuario\OneDrive - talentorecruiting.com\Forecasting"

# Log function
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] $Message"
    Add-Content -Path "$outputPath\automation.log" -Value "[$timestamp] $Message"
}

try {
    Write-Log "Starting Forecast Automation..."

    # Find source files
    Write-Log "Looking for source files in $downloadsPath"
    $csvMatches = Get-ChildItem -Path $downloadsPath -Filter "export*" -File | Where-Object { $_.Extension -eq ".csv" } | Sort-Object LastWriteTime -Descending
    if ($csvMatches.Count -gt 1) {
        Write-Log "WARNING: Multiple CSV files found matching 'export*' - using most recent: $($csvMatches[0].Name)"
        foreach ($f in $csvMatches) { Write-Log "  Found: $($f.Name) (modified $($f.LastWriteTime))" }
    }
    $csvFile = $csvMatches | Select-Object -First 1

    $excelMatches = Get-ChildItem -Path $downloadsPath -Filter "placement activity*" -File | Where-Object { $_.Extension -eq ".xlsx" -or $_.Extension -eq ".xls" } | Sort-Object LastWriteTime -Descending
    if ($excelMatches.Count -gt 1) {
        Write-Log "WARNING: Multiple Excel files found matching 'placement activity*' - using most recent: $($excelMatches[0].Name)"
        foreach ($f in $excelMatches) { Write-Log "  Found: $($f.Name) (modified $($f.LastWriteTime))" }
    }
    $excelFile = $excelMatches | Select-Object -First 1

    if (-not $csvFile) {
        throw "CSV file (export*) not found in Downloads"
    }
    if (-not $excelFile) {
        throw "Excel file (placement activity*) not found in Downloads"
    }

    Write-Log "Found CSV: $($csvFile.Name)"
    Write-Log "Found Excel: $($excelFile.Name)"

    # Create Excel COM object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # ========== PROCESS CSV FILE ==========
    Write-Log "Processing CSV file..."

    # Create temporary Excel workbook for CSV data
    $dateStamp = Get-Date -Format "yyyy-MM-dd"
    $csvExcelPath = Join-Path -Path $outputPath -ChildPath "export_processed_$dateStamp.xlsx"
    $csvWorkbook = $excel.Workbooks.Add()
    $csvSheet = $csvWorkbook.Sheets.Item(1)

    # Write CSV content to Excel using proper CSV parsing (handles commas inside fields)
    $csvData = Import-Csv -Path $csvFile.FullName
    $row = 1
    foreach ($record in $csvData) {
        $col = 1
        foreach ($property in $record.PSObject.Properties) {
            $csvSheet.Cells.Item($row, $col) = $property.Value
            $col++
        }
        $row++
    }

    $csvWorkbook.SaveAs($csvExcelPath, 51) # 51 = xlsx format
    $csvWorkbook.Close($false)
    Write-Log "CSV processed and saved to $csvExcelPath"

    # ========== PROCESS EXCEL FILE ==========
    Write-Log "Processing Excel file..."
    $excelWorkbook = $excel.Workbooks.Open($excelFile.FullName)
    $excelSheet = $excelWorkbook.Sheets.Item(1)

    # Delete first row
    $excelSheet.Rows.Item(1).Delete()

    # Delete column B
    $excelSheet.Columns.Item(2).Delete()

    # Delete rows where column A starts with "Total", "Totals", "Department", "Grand Total", or "Grand Totals" (case-insensitive, except header)
    $lastRow = $excelSheet.UsedRange.Rows.Count
    for ($i = $lastRow; $i -ge 2; $i--) {
        $cellValue = [string]$excelSheet.Cells.Item($i, 1).Value2
        if ($cellValue -imatch "^(totals?|department|grand totals?)\b") {
            $excelSheet.Rows.Item($i).Delete()
        }
    }

    $excelProcessedPath = Join-Path -Path $outputPath -ChildPath "placement_activity_processed_$dateStamp.xlsx"
    $excelWorkbook.SaveAs($excelProcessedPath, 51)
    $excelWorkbook.Close($false)
    Write-Log "Excel processed and saved to $excelProcessedPath"

    # ========== FIND AND OPEN FORECAST TOOL ==========
    Write-Log "Looking for Forecast Tool file..."
    $forecastToolFile = Get-ChildItem -Path $forecastToolPath -Filter "Forecast Tool*" -File | Where-Object { $_.Extension -eq ".xlsx" -or $_.Extension -eq ".xls" } | Select-Object -First 1

    if (-not $forecastToolFile) {
        throw "Forecast Tool file not found in $forecastToolPath"
    }

    Write-Log "Found Forecast Tool: $($forecastToolFile.Name)"
    $forecastWorkbook = $excel.Workbooks.Open($forecastToolFile.FullName)

    # ========== PASTE DATA INTO FORECAST TOOL ==========
    Write-Log "Updating Forecast Tool tabs..."

    # Get sheet names (should be "forecast data" and "placement data")
    $sheetNames = @()
    foreach ($sheet in $forecastWorkbook.Sheets) {
        $sheetNames += $sheet.Name
    }
    Write-Log "Available sheets: $($sheetNames -join ', ')"

    # Paste CSV data to "forecast data" tab
    if ($sheetNames -contains "forecast data") {
        $forecastDataSheet = $forecastWorkbook.Sheets.Item("forecast data")
        Write-Log "Clearing existing data from 'forecast data' tab..."
        $forecastDataSheet.Cells.Clear()

        Write-Log "Pasting CSV data to 'forecast data' tab..."
        $csvWorkbook = $excel.Workbooks.Open($csvExcelPath)
        $csvSheet = $csvWorkbook.Sheets.Item(1)
        $csvSheet.UsedRange.Copy()

        $forecastDataSheet.Cells.Item(1, 1).PasteSpecial([Microsoft.Office.Interop.Excel.XlPasteType]::xlPasteAll)

        $csvWorkbook.Close($false)
        Write-Log "CSV data pasted to forecast data tab"
    }

    # Paste Excel data to "placement data" tab
    if ($sheetNames -contains "placement data") {
        $placementDataSheet = $forecastWorkbook.Sheets.Item("placement data")
        Write-Log "Clearing existing data from 'placement data' tab..."
        $placementDataSheet.Cells.Clear()

        Write-Log "Pasting Excel data to 'placement data' tab..."
        $excelWorkbook = $excel.Workbooks.Open($excelProcessedPath)
        $excelSheet = $excelWorkbook.Sheets.Item(1)
        $excelSheet.UsedRange.Copy()

        $placementDataSheet.Cells.Item(1, 1).PasteSpecial([Microsoft.Office.Interop.Excel.XlPasteType]::xlPasteAll)

        $excelWorkbook.Close($false)
        Write-Log "Excel data pasted to placement data tab"
    }

    # ========== COMPARISON: placement data vs placed ==========
    if (($sheetNames -contains "placement data") -and ($sheetNames -contains "placed")) {
        Write-Log "Comparing placement data with placed tab..."
        $placementDataSheet = $forecastWorkbook.Sheets.Item("placement data")
        $placedSheet = $forecastWorkbook.Sheets.Item("placed")

        # Get data from placement data (column B is key)
        $placementLastRow = $placementDataSheet.UsedRange.Rows.Count
        $placementKeys = @{}
        for ($i = 2; $i -le $placementLastRow; $i++) {
            $key = $placementDataSheet.Cells.Item($i, 2).Value2
            if ($key) {
                $placementKeys[$key] = $i
            }
        }

        # Get data from placed (column C is key)
        $placedLastRow = $placedSheet.UsedRange.Rows.Count
        $placedKeys = @{}
        for ($i = 3; $i -le $placedLastRow; $i++) { # Skip rows 1 and 2
            $key = $placedSheet.Cells.Item($i, 3).Value2
            if ($key) {
                $placedKeys[$key] = $i
            }
        }

        # Find new entries in placement data
        $newCount = 0
        foreach ($key in $placementKeys.Keys) {
            if (-not $placedKeys.ContainsKey($key)) {
                # Find first empty row in placed tab (skip rows 1-2), check column B since data starts there
                $emptyRow = 3
                while ($placedSheet.Cells.Item($emptyRow, 2).Value2) {
                    $emptyRow++
                }

                # Copy columns A-AI (1-35) from placement data into columns B-AJ (2-36) on placed tab
                for ($col = 1; $col -le 35; $col++) {
                    $value = $placementDataSheet.Cells.Item($placementKeys[$key], $col).Value2
                    $placedSheet.Cells.Item($emptyRow, $col + 1) = $value
                }

                # Copy formulas from columns AK:CO (37-93) from the last populated row above
                $formulaSourceRow = $emptyRow - 1
                if ($formulaSourceRow -ge 3) {
                    $sourceRange = $placedSheet.Range(
                        $placedSheet.Cells.Item($formulaSourceRow, 37),
                        $placedSheet.Cells.Item($formulaSourceRow, 93)
                    )
                    $destRange = $placedSheet.Range(
                        $placedSheet.Cells.Item($emptyRow, 37),
                        $placedSheet.Cells.Item($emptyRow, 93)
                    )
                    $sourceRange.Copy($destRange)
                    Write-Log "Copied formulas (AK:CO) from row $formulaSourceRow to row $emptyRow"
                }

                $newCount++
                Write-Log "Added new entry to placed tab (row $emptyRow): $key"
            }
        }
        Write-Log "Total new entries added to placed: $newCount"
    }

    # ========== COMPARISON: forecast data vs forecast ==========
    if (($sheetNames -contains "forecast data") -and ($sheetNames -contains "forecast")) {
        Write-Log "Comparing forecast data with forecast tab..."
        $forecastDataSheet = $forecastWorkbook.Sheets.Item("forecast data")
        $forecastSheet = $forecastWorkbook.Sheets.Item("forecast")

        # Get data from forecast data (column A is key)
        $forecastDataLastRow = $forecastDataSheet.UsedRange.Rows.Count
        $forecastDataKeys = @{}
        for ($i = 1; $i -le $forecastDataLastRow; $i++) {
            $key = $forecastDataSheet.Cells.Item($i, 1).Value2
            if ($key) {
                $forecastDataKeys[$key] = $i
            }
        }

        # Get data from forecast (column C is key)
        $forecastLastRow = $forecastSheet.UsedRange.Rows.Count
        $forecastKeys = @{}
        for ($i = 4; $i -le $forecastLastRow; $i++) { # Skip first three rows
            $key = $forecastSheet.Cells.Item($i, 3).Value2
            if ($key) {
                $forecastKeys[$key] = $i
            }
        }

        # Find new entries in forecast data
        $addedCount = 0
        $highlightedCount = 0
        foreach ($key in $forecastDataKeys.Keys) {
            if (-not $forecastKeys.ContainsKey($key)) {
                # Find first empty row in forecast tab (skip rows 1-3), check column C since data starts there
                $emptyRow = 4
                while ($forecastSheet.Cells.Item($emptyRow, 3).Value2) {
                    $emptyRow++
                }

                # Copy columns A-AC (1-29) from forecast data into columns C-AE (3-31) on forecast tab
                for ($col = 1; $col -le 29; $col++) {
                    $value = $forecastDataSheet.Cells.Item($forecastDataKeys[$key], $col).Value2
                    $forecastSheet.Cells.Item($emptyRow, $col + 2) = $value
                }

                # Copy formulas from columns AI:CA (35-79) from the last populated row above
                $formulaSourceRow = $emptyRow - 1
                if ($formulaSourceRow -ge 4) {
                    $sourceRange = $forecastSheet.Range(
                        $forecastSheet.Cells.Item($formulaSourceRow, 35),
                        $forecastSheet.Cells.Item($formulaSourceRow, 79)
                    )
                    $destRange = $forecastSheet.Range(
                        $forecastSheet.Cells.Item($emptyRow, 35),
                        $forecastSheet.Cells.Item($emptyRow, 79)
                    )
                    $sourceRange.Copy($destRange)
                    Write-Log "Copied formulas (AI:CA) from row $formulaSourceRow to row $emptyRow"
                }

                $addedCount++
                Write-Log "Added new entry to forecast tab (row $emptyRow): $key"
            }
        }

        # Find entries in forecast that are not in forecast data and highlight them
        foreach ($key in $forecastKeys.Keys) {
            if (-not $forecastDataKeys.ContainsKey($key)) {
                $rowNum = $forecastKeys[$key]
                # Highlight entire row yellow
                $forecastSheet.Rows.Item($rowNum).Interior.Color = 65535 # Yellow
                $highlightedCount++
                Write-Log "Highlighted missing entry in forecast tab (row $rowNum): $key"
            }
        }

        Write-Log "Total new entries added to forecast: $addedCount"
        Write-Log "Total entries highlighted in yellow: $highlightedCount"
    }

    # ========== UPDATE LOG LOG TAB ==========
    Write-Log "Updating 'log log' tab..."
    try {
        $ErrorActionPreference = "Stop"

        # Step 1 -- count active jobs
        Write-Log "  [LL-1] Counting active jobs. sheetNames contains 'forecast': $($sheetNames -contains 'forecast')"
        $totalActiveJobs = 0
        if ($sheetNames -contains "forecast") {
            Write-Log "  [LL-2] Opening forecast sheet..."
            $fSheet = $forecastWorkbook.Sheets.Item("forecast")
            $fLastRow = $fSheet.UsedRange.Rows.Count
            Write-Log "  [LL-3] Forecast last row: $fLastRow"
            for ($i = 4; $i -le $fLastRow; $i++) {
                $cellVal = $fSheet.Cells.Item($i, 3).Value2
                $rowColor = $fSheet.Rows.Item($i).Interior.Color
                if ($cellVal -and $rowColor -ne 65535) { $totalActiveJobs++ }
            }
        }
        Write-Log "  [LL-4] totalActiveJobs=$totalActiveJobs"

        # Step 2 -- count placements
        Write-Log "  [LL-5] Counting placements. sheetNames contains 'placed': $($sheetNames -contains 'placed')"
        $totalPlacements = 0
        if ($sheetNames -contains "placed") {
            Write-Log "  [LL-6] Opening placed sheet..."
            $pSheet = $forecastWorkbook.Sheets.Item("placed")
            $pLastRow = $pSheet.UsedRange.Rows.Count
            Write-Log "  [LL-7] Placed last row: $pLastRow"
            for ($i = 3; $i -le $pLastRow; $i++) {
                if ($pSheet.Cells.Item($i, 2).Value2) { $totalPlacements++ }
            }
        }
        Write-Log "  [LL-8] totalPlacements=$totalPlacements"

        # Step 3 -- find log log sheet
        Write-Log "  [LL-9] sheetNames: $($sheetNames -join ', ')"
        Write-Log "  [LL-10] sheetNames contains 'log log': $($sheetNames -contains 'log log')"
        if (-not ($sheetNames -contains "log log")) {
            Write-Log "  WARNING: 'log log' tab not found in sheetNames -- skipping log update"
        } else {
            Write-Log "  [LL-11] Opening log log sheet..."
            $logSheet = $forecastWorkbook.Sheets.Item("log log")
            Write-Log "  [LL-12] ListObjects count: $($logSheet.ListObjects.Count)"
            for ($li = 1; $li -le $logSheet.ListObjects.Count; $li++) {
                $lo = $logSheet.ListObjects.Item($li)
                Write-Log "  Table[$li]: '$($lo.Name)' at col $($lo.Range.Column)"
            }

            Write-Log "  [LL-13] Looking up CumulativeData table..."
            $logListObj1 = $null
            $logListObj2 = $null
            try { $logListObj1 = $logSheet.ListObjects.Item("CumulativeData") } catch { Write-Log "  CumulativeData lookup error: $_" }
            Write-Log "  [LL-14] Looking up LastRunData table..."
            try { $logListObj2 = $logSheet.ListObjects.Item("LastRunData")    } catch { Write-Log "  LastRunData lookup error: $_" }
            Write-Log "  [LL-15] CumulativeData found: $($null -ne $logListObj1) | LastRunData found: $($null -ne $logListObj2)"

            # ---- Accumulating table: add row at top, renumber IDs ----
            if ($logListObj1) {
                Write-Log "  [LL-16] Adding row to CumulativeData at position 1..."
                $newRow1 = $logListObj1.ListRows.Add(1)
                Write-Log "  [LL-17] Writing cells..."
                $newRow1.Range.Cells.Item(1, 1) = 0
                $newRow1.Range.Cells.Item(1, 2) = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                $newRow1.Range.Cells.Item(1, 3) = if ($addedCount)      { $addedCount }      else { 0 }
                $newRow1.Range.Cells.Item(1, 4) = if ($newCount)        { $newCount }        else { 0 }
                $newRow1.Range.Cells.Item(1, 5) = $totalActiveJobs
                $newRow1.Range.Cells.Item(1, 6) = $totalPlacements
                $newRow1.Range.Cells.Item(1, 7) = if ($highlightedCount) { $highlightedCount } else { 0 }
                Write-Log "  [LL-18] Renumbering IDs..."
                $totalRows = $logListObj1.ListRows.Count
                for ($r = 1; $r -le $totalRows; $r++) {
                    $logListObj1.ListRows.Item($r).Range.Cells.Item(1, 1) = $totalRows - ($r - 1)
                }
                Write-Log "  [LL-19] CumulativeData done: $totalRows total entries"
            } else {
                Write-Log "  WARNING: 'CumulativeData' table not found in log log tab"
            }

            # ---- Latest-only table: clear all rows, write single row ----
            if ($logListObj2) {
                Write-Log "  [LL-20] Clearing LastRunData rows..."
                while ($logListObj2.ListRows.Count -gt 0) { $logListObj2.ListRows.Item(1).Delete() }
                Write-Log "  [LL-21] Adding new row to LastRunData..."
                $newRow2 = $logListObj2.ListRows.Add()
                $newRow2.Range.Cells.Item(1, 1) = 1
                $newRow2.Range.Cells.Item(1, 2) = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                $newRow2.Range.Cells.Item(1, 3) = if ($addedCount)      { $addedCount }      else { 0 }
                $newRow2.Range.Cells.Item(1, 4) = if ($newCount)        { $newCount }        else { 0 }
                $newRow2.Range.Cells.Item(1, 5) = $totalActiveJobs
                $newRow2.Range.Cells.Item(1, 6) = $totalPlacements
                $newRow2.Range.Cells.Item(1, 7) = if ($highlightedCount) { $highlightedCount } else { 0 }
                Write-Log "  [LL-22] LastRunData written successfully"
            } else {
                Write-Log "  WARNING: 'LastRunData' table not found in log log tab"
            }
        }
    } catch {
        Write-Log "  ERROR in log log section: $_"
        Write-Log "  Stack trace: $($_.ScriptStackTrace)"
    } finally {
        $ErrorActionPreference = "Continue"
    }

    # ========== SAVE FORECAST TOOL ==========
    Write-Log "Saving Forecast Tool file..."
    $forecastSavePath = Join-Path $forecastToolPath "Forecast Tool $dateStamp.xlsx"
    $alreadyNamed = ($forecastToolFile.FullName -ieq $forecastSavePath)

    if ($alreadyNamed) {
        # File is already named with today's date -- save in place
        $forecastWorkbook.Save()
        Write-Log "Forecast Tool saved in place (already named 'Forecast Tool $dateStamp.xlsx')"
    } else {
        $forecastWorkbook.SaveAs($forecastSavePath, 51)
        Write-Log "Forecast Tool saved as 'Forecast Tool $dateStamp.xlsx'"
    }
    $forecastWorkbook.Close($false)

    # Copy dated file to Reports folder
    $reportsPath = "C:\Users\usuario\OneDrive - talentorecruiting.com\Documentos\Reports\ForecastReports"
    if (-not (Test-Path $reportsPath)) {
        New-Item -ItemType Directory -Path $reportsPath | Out-Null
        Write-Log "Created Reports folder at $reportsPath"
    }
    Copy-Item -Path $forecastSavePath -Destination (Join-Path $reportsPath "Forecast Tool.xlsx") -Force
    Write-Log "Copied to Reports folder as 'Forecast Tool.xlsx'"

    # Move the original file to the Old subfolder (only if it had a different name)
    if (-not $alreadyNamed) {
        $oldFolderPath = Join-Path $forecastToolPath "Old"
        if (-not (Test-Path $oldFolderPath)) {
            New-Item -ItemType Directory -Path $oldFolderPath | Out-Null
            Write-Log "Created 'Old' folder at $oldFolderPath"
        }
        $oldDestPath = Join-Path $oldFolderPath $forecastToolFile.Name
        Move-Item -Path $forecastToolFile.FullName -Destination $oldDestPath -Force
        Write-Log "Moved '$($forecastToolFile.Name)' to Old folder"
    }

    Write-Log "Automation completed successfully!"

} catch {
    Write-Log "ERROR: $_"
    Write-Error $_
} finally {
    # Clean up
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    Write-Log "Script ended"
}
