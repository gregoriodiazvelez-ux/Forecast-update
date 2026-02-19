# Forecast Update Automation Script
# Runs Monday at 9:30 AM to process and update forecast files

# Configuration
$downloadsPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads")
$outputPath = "C:\Users\usuario\OneDrive - talentorecruiting.com\Forecasting\Data"
$forecastToolPath = "C:\Users\usuario\OneDrive - talentorecruiting.com\Forecasting"

# Email Configuration
# Setup: https://portal.azure.com → Azure AD → App registrations → New registration
#   → Certificates & secrets → New client secret
#   → API permissions → Add → Microsoft Graph → Application → Mail.Send → Grant admin consent
# Load credentials from local config file (not stored in git)
$configFile = Join-Path $PSScriptRoot "config.ps1"
if (Test-Path $configFile) {
    . $configFile
} else {
    throw "config.ps1 not found. Copy config.ps1.example to config.ps1 and fill in your values."
}

# Log function
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] $Message"
    Add-Content -Path "$outputPath\automation.log" -Value "[$timestamp] $Message"
}

# Email notification function
function Send-EmailNotification {
    param(
        [int]$NewPlacements,
        [int]$NewForecastJobs,
        [int]$HighlightedJobs,
        [bool]$Success = $true,
        [string]$ErrorMessage = ""
    )
    try {
        $runDate = Get-Date -Format "MMMM dd, yyyy"
        $runTime = Get-Date -Format "hh:mm tt"

        if ($Success) {
            $subject = "Forecast Update Summary - $runDate"

            # Calculate SVG bar heights (chart baseline at y=175, max chart height 140px)
            $maxVal = [Math]::Max([Math]::Max($NewPlacements, $NewForecastJobs), [Math]::Max($HighlightedJobs, 1))
            $base   = 175
            $chartH = 140
            $bar1H  = [int][Math]::Round(($NewPlacements  / $maxVal) * $chartH)
            $bar2H  = [int][Math]::Round(($NewForecastJobs / $maxVal) * $chartH)
            $bar3H  = [int][Math]::Round(($HighlightedJobs / $maxVal) * $chartH)
            $bar1Y  = $base - $bar1H
            $bar2Y  = $base - $bar2H
            $bar3Y  = $base - $bar3H
            $bar1TY = $bar1Y - 8
            $bar2TY = $bar2Y - 8
            $bar3TY = $bar3Y - 8

            $body = @"
<!DOCTYPE html>
<html>
<body style="margin:0; padding:20px; background-color:#f0f4f0; font-family:Arial,sans-serif;">
  <div style="max-width:580px; margin:0 auto; background:white; border-radius:12px; overflow:hidden; box-shadow:0 4px 20px rgba(0,0,0,0.12);">

    <!-- Header -->
    <div style="background:linear-gradient(135deg,#1B5E20,#388E3C); padding:28px 30px; text-align:center;">
      <div style="font-size:22px; font-weight:bold; color:white; letter-spacing:2px;">TALENTO RECRUITING</div>
      <div style="font-size:14px; color:#A5D6A7; margin-top:6px;">Forecast Automation Report</div>
      <div style="font-size:12px; color:#C8E6C9; margin-top:4px;">$runDate &#8226; $runTime</div>
    </div>

    <!-- Bar Chart -->
    <div style="padding:28px 30px 10px;">
      <p style="text-align:center; color:#666; font-size:13px; margin:0 0 20px;">Weekly update completed successfully. Here are this week's results:</p>
      <svg width="100%" viewBox="0 0 500 210" xmlns="http://www.w3.org/2000/svg">
        <!-- Grid lines -->
        <line x1="50" y1="175" x2="470" y2="175" stroke="#ccc" stroke-width="1.5"/>
        <line x1="50" y1="140" x2="470" y2="140" stroke="#eee" stroke-width="1" stroke-dasharray="4,3"/>
        <line x1="50" y1="105" x2="470" y2="105" stroke="#eee" stroke-width="1" stroke-dasharray="4,3"/>
        <line x1="50" y1="70"  x2="470" y2="70"  stroke="#eee" stroke-width="1" stroke-dasharray="4,3"/>
        <line x1="50" y1="35"  x2="470" y2="35"  stroke="#eee" stroke-width="1" stroke-dasharray="4,3"/>
        <!-- Bar 1: New Placements (dark green) -->
        <rect x="80" y="$bar1Y" width="90" height="$bar1H" fill="#2E7D32" rx="5"/>
        <text x="125" y="$bar1TY" text-anchor="middle" fill="#1B5E20" font-size="15" font-weight="bold" font-family="Arial">$NewPlacements</text>
        <text x="125" y="195" text-anchor="middle" fill="#555" font-size="11" font-family="Arial">New Placements</text>
        <!-- Bar 2: New Forecast Jobs (light green) -->
        <rect x="205" y="$bar2Y" width="90" height="$bar2H" fill="#66BB6A" rx="5"/>
        <text x="250" y="$bar2TY" text-anchor="middle" fill="#2E7D32" font-size="15" font-weight="bold" font-family="Arial">$NewForecastJobs</text>
        <text x="250" y="195" text-anchor="middle" fill="#555" font-size="11" font-family="Arial">New Jobs Added</text>
        <!-- Bar 3: Flagged Yellow -->
        <rect x="330" y="$bar3Y" width="90" height="$bar3H" fill="#FDD835" rx="5"/>
        <text x="375" y="$bar3TY" text-anchor="middle" fill="#F57F17" font-size="15" font-weight="bold" font-family="Arial">$HighlightedJobs</text>
        <text x="375" y="195" text-anchor="middle" fill="#555" font-size="11" font-family="Arial">Flagged Yellow</text>
      </svg>
    </div>

    <!-- Stat Cards -->
    <div style="padding:0 30px 28px;">
      <table width="100%" cellpadding="0" cellspacing="8" style="border-collapse:separate;">
        <tr>
          <td style="background:#E8F5E9; border-radius:10px; padding:16px 12px; text-align:center; width:33%;">
            <div style="font-size:36px; font-weight:bold; color:#2E7D32; line-height:1;">$NewPlacements</div>
            <div style="font-size:10px; color:#4CAF50; margin-top:5px; letter-spacing:1px;">NEW PLACEMENTS</div>
          </td>
          <td style="background:#F1F8E9; border-radius:10px; padding:16px 12px; text-align:center; width:33%;">
            <div style="font-size:36px; font-weight:bold; color:#388E3C; line-height:1;">$NewForecastJobs</div>
            <div style="font-size:10px; color:#7CB342; margin-top:5px; letter-spacing:1px;">NEW JOBS ADDED</div>
          </td>
          <td style="background:#FFFDE7; border-radius:10px; padding:16px 12px; text-align:center; width:33%;">
            <div style="font-size:36px; font-weight:bold; color:#F9A825; line-height:1;">$HighlightedJobs</div>
            <div style="font-size:10px; color:#FFA000; margin-top:5px; letter-spacing:1px;">FLAGGED YELLOW</div>
          </td>
        </tr>
      </table>
    </div>

    <!-- Footer -->
    <div style="background:#F5F5F5; padding:14px 30px; text-align:center; border-top:1px solid #E8E8E8;">
      <p style="color:#BDBDBD; font-size:11px; margin:0;">Automated by Forecast Update Automation &#8226; Talento Recruiting</p>
    </div>
  </div>
</body>
</html>
"@
        } else {
            $subject = "Forecast Update FAILED - $runDate"
            $body = @"
<!DOCTYPE html>
<html>
<body style="margin:0; padding:20px; background-color:#fce4e4; font-family:Arial,sans-serif;">
  <div style="max-width:580px; margin:0 auto; background:white; border-radius:12px; overflow:hidden; box-shadow:0 4px 20px rgba(0,0,0,0.12);">
    <div style="background:linear-gradient(135deg,#B71C1C,#C62828); padding:28px 30px; text-align:center;">
      <div style="font-size:22px; font-weight:bold; color:white; letter-spacing:2px;">TALENTO RECRUITING</div>
      <div style="font-size:14px; color:#EF9A9A; margin-top:6px;">Forecast Automation Report</div>
      <div style="font-size:12px; color:#FFCDD2; margin-top:4px;">$runDate &#8226; $runTime</div>
    </div>
    <div style="padding:30px;">
      <h2 style="color:#C62828; margin:0 0 16px;">Automation Failed</h2>
      <p style="color:#555; margin:0 0 12px;">The weekly forecast update encountered an error and could not complete.</p>
      <div style="background:#FFF3F3; border-left:4px solid #C62828; padding:14px 16px; border-radius:4px; font-size:13px; color:#333; word-break:break-word;">$ErrorMessage</div>
      <p style="color:#777; font-size:13px; margin:16px 0 0;">Please check the automation log in the Forecasting\Data folder for full details.</p>
    </div>
    <div style="background:#F5F5F5; padding:14px 30px; text-align:center; border-top:1px solid #E8E8E8;">
      <p style="color:#BDBDBD; font-size:11px; margin:0;">Automated by Forecast Update Automation &#8226; Talento Recruiting</p>
    </div>
  </div>
</body>
</html>
"@
        }

        # Send via Microsoft Graph API — works in scheduled tasks, no SMTP AUTH needed
        $tokenResponse = Invoke-RestMethod `
            -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
            -Method Post `
            -Body @{
                client_id     = $clientId
                client_secret = $clientSecret
                scope         = "https://graph.microsoft.com/.default"
                grant_type    = "client_credentials"
            }

        $toRecipients = if ($emailTo -is [array]) {
            $emailTo | ForEach-Object { @{ emailAddress = @{ address = $_ } } }
        } else {
            @(@{ emailAddress = @{ address = $emailTo } })
        }

        $payload = @{
            message = @{
                subject      = $subject
                body         = @{ contentType = "HTML"; content = $body }
                toRecipients = $toRecipients
            }
        } | ConvertTo-Json -Depth 10

        Invoke-RestMethod `
            -Uri "https://graph.microsoft.com/v1.0/users/$emailFrom/sendMail" `
            -Method Post `
            -Headers @{ Authorization = "Bearer $($tokenResponse.access_token)" } `
            -ContentType "application/json" `
            -Body $payload

        Write-Log "Email notification sent to $emailTo"
    } catch {
        Write-Log "WARNING: Failed to send email notification: $_"
    }
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

    # ========== SAVE FORECAST TOOL ==========
    Write-Log "Saving Forecast Tool file..."
    $forecastWorkbook.Save()
    $forecastWorkbook.Close($false)
    Write-Log "Forecast Tool saved successfully"

    Write-Log "Automation completed successfully!"

    # Send summary email
    Send-EmailNotification -NewPlacements $newCount -NewForecastJobs $addedCount -HighlightedJobs $highlightedCount

} catch {
    Write-Log "ERROR: $_"
    Write-Error $_
    # Send failure notification email
    Send-EmailNotification -NewPlacements 0 -NewForecastJobs 0 -HighlightedJobs 0 -Success $false -ErrorMessage $_.ToString()
} finally {
    # Clean up
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    Write-Log "Script ended"
}
