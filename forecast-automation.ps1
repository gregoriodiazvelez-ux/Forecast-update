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
    $dateStamp = Get-Date -Format "yyyy-MM-dd"

    # Find source files
    Write-Log "Looking for source files in $downloadsPath"
    $csvFile = Get-ChildItem -Path $downloadsPath -Filter "export*" -File | Where-Object { $_.Extension -eq ".csv" } | Select-Object -First 1
    $excelFile = Get-ChildItem -Path $downloadsPath -Filter "placement activity*" -File | Where-Object { $_.Extension -eq ".xlsx" -or $_.Extension -eq ".xls" } | Select-Object -First 1

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

    # Delete rows where column A starts with "Total", "Totals", or "Department" (case-insensitive, except header)
    $lastRow = $excelSheet.UsedRange.Rows.Count
    for ($i = $lastRow; $i -ge 2; $i--) {
        $cellValue = [string]$excelSheet.Cells.Item($i, 1).Value2
        if ($cellValue -imatch "^(totals?|department)\b") {
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
        Write-Log "Pasting CSV data to 'forecast data' tab..."
        $csvWorkbook = $excel.Workbooks.Open($csvExcelPath)
        $csvSheet = $csvWorkbook.Sheets.Item(1)
        $csvSheet.UsedRange.Copy()

        $forecastDataSheet = $forecastWorkbook.Sheets.Item("forecast data")
        $forecastDataSheet.Cells.Item(1, 1).PasteSpecial([Microsoft.Office.Interop.Excel.XlPasteType]::xlPasteAll)

        $csvWorkbook.Close($false)
        Write-Log "CSV data pasted to forecast data tab"
    }

    # Paste Excel data to "placement data" tab
    if ($sheetNames -contains "placement data") {
        Write-Log "Pasting Excel data to 'placement data' tab..."
        $excelWorkbook = $excel.Workbooks.Open($excelProcessedPath)
        $excelSheet = $excelWorkbook.Sheets.Item(1)
        $excelSheet.UsedRange.Copy()

        $placementDataSheet = $forecastWorkbook.Sheets.Item("placement data")
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
                # Find first empty row in placed tab (skip rows 1-2)
                $emptyRow = 3
                while ($placedSheet.Cells.Item($emptyRow, 1).Value2) {
                    $emptyRow++
                }

                # Copy entire row from placement data to placed
                for ($col = 1; $col -le 26; $col++) { # A-Z columns
                    $value = $placementDataSheet.Cells.Item($placementKeys[$key], $col).Value2
                    $placedSheet.Cells.Item($emptyRow, $col) = $value
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

    # ========== SAVE FORECAST TOOL ==========
    Write-Log "Saving Forecast Tool file..."
    $forecastWorkbook.Save()
    $forecastWorkbook.Close($false)
    Write-Log "Forecast Tool saved successfully"

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
