# Setup script to create a Windows Scheduled Task for the Forecast Automation
# Run this script as Administrator to set up the scheduled task

# Configuration
$scriptPath = "C:\path\to\forecast-automation.ps1" # Update this with actual path
$taskName = "Forecast Update Automation"
$taskDescription = "Automated forecast and placement data update - Runs every Monday at 9:30 AM"
$time = "09:30"
$daysOfWeek = "Monday"

Write-Host "Setting up Scheduled Task for Forecast Automation..."
Write-Host ""

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Error "This script must be run as Administrator!"
    exit 1
}

# Verify script exists
if (-not (Test-Path $scriptPath)) {
    Write-Error "Script not found at: $scriptPath"
    Write-Host "Please update the `$scriptPath variable with the correct path to forecast-automation.ps1"
    exit 1
}

# Remove existing task if it exists
$existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "Removing existing scheduled task..."
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

# Create trigger (Weekly - Monday at 9:30 AM)
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $daysOfWeek -At $time

# Create action
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""

# Create principal (run as SYSTEM or current user)
$principal = New-ScheduledTaskPrincipal -UserId "$env:COMPUTERNAME\$env:USERNAME" -LogonType Interactive -RunLevel Highest

# Register the task
try {
    Register-ScheduledTask -TaskName $taskName -Trigger $trigger -Action $action -Principal $principal -Description $taskDescription -Force
    Write-Host ""
    Write-Host "SUCCESS! Scheduled task created:" -ForegroundColor Green
    Write-Host "  Task Name: $taskName"
    Write-Host "  Schedule: Every $daysOfWeek at $time"
    Write-Host "  Script: $scriptPath"
    Write-Host ""
    Write-Host "The task will run automatically every Monday at 9:30 AM"
    Write-Host ""
    Write-Host "To verify the task:"
    Write-Host "  taskschd.msc  (or search 'Task Scheduler')"
    Write-Host ""
} catch {
    Write-Error "Failed to create scheduled task: $_"
    exit 1
}
