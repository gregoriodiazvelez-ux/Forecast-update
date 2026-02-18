# Forecast Automation Setup Instructions

This automation processes Excel and CSV files from your Downloads folder and updates your Forecast Tool file.

## What the Automation Does

### File Processing
1. **CSV File** (export*):
   - Removes first row
   - Converts to Excel format
   - Saves as `export_processed.xlsx`

2. **Excel File** (placement activity*):
   - Removes first row
   - Deletes column B
   - Removes rows starting with "total", "totals", or "department" (keeps header)
   - Saves as `placement_activity_processed.xlsx`

### Forecast Tool Updates
- Pastes processed CSV data → "forecast data" tab
- Pastes processed Excel data → "placement data" tab
- Compares and adds new entries:
  - **placement data** (col B) ↔ **placed** tab (col C): Adds new entries starting at row 3
  - **forecast data** (col A) ↔ **forecast** tab (col C): Adds new entries starting at row 4
- Highlights missing entries in **forecast** tab with yellow color

### Output
- Processed files saved to: `C:\Users\usuario\OneDrive - talentorecruiting.com\Forecasting\Data`
- Execution log: `automation.log` in the Data folder

## Setup Steps

### Step 1: Update Script Paths
Edit `forecast-automation.ps1` if your paths are different:
```powershell
$downloadsPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads")  # Usually fine
$outputPath = "C:\Users\usuario\OneDrive - talentorecruiting.com\Forecasting\Data"
$forecastToolPath = "C:\Users\usuario\OneDrive - talentorecruiting.com\Forecasting"
```

### Step 2: Edit Setup Script
Edit `setup-scheduled-task.ps1` and update this line with the FULL path to `forecast-automation.ps1`:
```powershell
$scriptPath = "C:\path\to\forecast-automation.ps1"  # Change this!
```

Example:
```powershell
$scriptPath = "C:\Users\usuario\Downloads\forecast-automation.ps1"
```

### Step 3: Run Setup as Administrator
1. Open PowerShell **as Administrator**
2. Run this command:
```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
.\setup-scheduled-task.ps1
```

You should see:
```
SUCCESS! Scheduled task created:
  Task Name: Forecast Update Automation
  Schedule: Every Monday at 09:30
  Script: C:\path\to\forecast-automation.ps1

The task will run automatically every Monday at 9:30 AM
```

### Step 4: Verify Task Creation
Open Task Scheduler:
- Press `Win + R`, type `taskschd.msc`, press Enter
- Or search for "Task Scheduler"
- Look for "Forecast Update Automation" in the task list

## File Requirements

Your files must be in this exact location each week:
- **CSV File**: `Downloads\export*` (e.g., export_2024-02-19.csv)
- **Excel File**: `Downloads\placement activity*` (e.g., placement activity_2024-02-19.xlsx)

## Running Manually

To test or run the automation anytime without waiting for Monday:

```powershell
.\forecast-automation.ps1
```

**Important**: Run from PowerShell as Administrator for proper Excel COM operations.

## Troubleshooting

### Task won't run
- Ensure script path in `setup-scheduled-task.ps1` is correct (full path, not relative)
- Check that the account has permission to access all file paths
- Look at Task Scheduler History tab for error messages

### Excel files locked error
- Ensure no one has the Forecast Tool file open
- Make sure Excel isn't running in the background
- Check that antivirus isn't locking files during processing

### Files not found
- Verify CSV and Excel files are in Downloads folder
- Check file names start with "export" and "placement activity"
- Use Get-ChildItem in PowerShell to verify:
```powershell
Get-ChildItem "$env:USERPROFILE\Downloads\export*"
Get-ChildItem "$env:USERPROFILE\Downloads\placement activity*"
```

### Check logs
Look at `C:\Users\usuario\OneDrive - talentorecruiting.com\Forecasting\Data\automation.log` for detailed execution information

## Important Notes

- **Excel must be installed** on the machine for this to work
- **Run as Administrator** for proper permissions
- The script creates Excel COM objects - these are automatically cleaned up
- If automation fails, processed files may be left in the Data folder for debugging
- Always keep a backup of your Forecast Tool file before first run

## Modifying the Schedule

To change the day/time:
1. Open `setup-scheduled-task.ps1`
2. Modify these lines:
```powershell
$time = "09:30"              # Change time (24-hour format)
$daysOfWeek = "Monday"       # Change day
```
3. Re-run the setup script as Administrator

Options for `$daysOfWeek`: Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday
