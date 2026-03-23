# ═══════════════════════════════════════════════════════════════
#  SETUP_SCHEDULE.ps1
#  Sets up the scheduled task for the morning meeting automation:
#
#  AShop_MorningMeeting_Update : 7:45am Mon-Fri
#    → runs update script, reads PPT + Downtime Excel, pushes to GitHub
#    → StartWhenAvailable = True (runs on first login if missed)
#
#  Power BI does NOT need to be open. All KPI data comes from PPT.
#
#  Run this ONCE from PowerShell (as yourself, no admin needed).
#  Right-click → "Run with PowerShell"
# ═══════════════════════════════════════════════════════════════

$dashFolder = "C:\Users\NYOUSIF\Desktop\AShop_Dashboard"
$user       = $env:USERNAME

# ── Remove old PBI launch task if it exists (no longer needed) ───────────────
$oldTask = Get-ScheduledTask -TaskName "AShop_LaunchPBI" -ErrorAction SilentlyContinue
if ($oldTask) {
    Unregister-ScheduledTask -TaskName "AShop_LaunchPBI" -Confirm:$false
    Write-Host "Removed old task: AShop_LaunchPBI (Power BI no longer needed)"
}

# ── Task: Run morning meeting update at 7:45am ────────────────────────────────
$action   = New-ScheduledTaskAction `
              -Execute "$dashFolder\RUN_MORNING_MEETING.bat"
$trigger  = New-ScheduledTaskTrigger `
              -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At "7:45AM"
$settings = New-ScheduledTaskSettingsSet `
              -ExecutionTimeLimit (New-TimeSpan -Minutes 15) `
              -StartWhenAvailable
$principal = New-ScheduledTaskPrincipal `
              -UserId $user -LogonType Interactive -RunLevel Limited

Register-ScheduledTask `
    -TaskName  "AShop_MorningMeeting_Update" `
    -Action    $action `
    -Trigger   $trigger `
    -Settings  $settings `
    -Principal $principal `
    -Force

Write-Host ""
Write-Host "Task registered: AShop_MorningMeeting_Update  (7:45am Mon-Fri)"
Write-Host "  StartWhenAvailable = True"
Write-Host "  If computer is off at 7:45, the task runs on next login automatically."
Write-Host ""

# ── Confirm ───────────────────────────────────────────────────────────────────
Get-ScheduledTask | Where-Object { $_.TaskName -like "AShop_*" } |
    Select-Object TaskName, State |
    Format-Table -AutoSize
