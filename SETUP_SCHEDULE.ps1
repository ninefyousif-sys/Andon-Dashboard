# ═══════════════════════════════════════════════════════════════════
#  SETUP_SCHEDULE.ps1
#  Sets up TWO scheduled tasks for the morning meeting automation:
#
#  Task 1 — AShop_LaunchPBI       : 7:20am  → opens both PBI files
#  Task 2 — AShop_MorningMeeting  : 7:45am  → runs update script
#
#  Run this ONCE from PowerShell (as yourself, no admin needed).
#  Right-click → "Run with PowerShell"
# ═══════════════════════════════════════════════════════════════════

$dashFolder = "C:\Users\NYOUSIF\Desktop\AShop_Dashboard"
$user       = $env:USERNAME

# ── Task 1: Launch PBI at 7:20am ─────────────────────────────────────────────
$action1  = New-ScheduledTaskAction `
              -Execute "$dashFolder\LAUNCH_PBI.bat"
$trigger1 = New-ScheduledTaskTrigger `
              -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At "7:20AM"
$settings1 = New-ScheduledTaskSettingsSet `
              -ExecutionTimeLimit (New-TimeSpan -Minutes 2) `
              -StartWhenAvailable
$principal = New-ScheduledTaskPrincipal `
              -UserId $user -LogonType Interactive -RunLevel Limited

Register-ScheduledTask `
    -TaskName  "AShop_LaunchPBI" `
    -Action    $action1 `
    -Trigger   $trigger1 `
    -Settings  $settings1 `
    -Principal $principal `
    -Force

Write-Host "Task 1 registered: AShop_LaunchPBI  (7:20am Mon-Fri)"

# ── Task 2: Run update at 7:45am ──────────────────────────────────────────────
$action2  = New-ScheduledTaskAction `
              -Execute "$dashFolder\RUN_MORNING_MEETING.bat"
$trigger2 = New-ScheduledTaskTrigger `
              -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At "7:45AM"
$settings2 = New-ScheduledTaskSettingsSet `
              -ExecutionTimeLimit (New-TimeSpan -Minutes 15) `
              -StartWhenAvailable

Register-ScheduledTask `
    -TaskName  "AShop_MorningMeeting_Update" `
    -Action    $action2 `
    -Trigger   $trigger2 `
    -Settings  $settings2 `
    -Principal $principal `
    -Force

Write-Host "Task 2 registered: AShop_MorningMeeting_Update  (7:45am Mon-Fri)"

# ── Confirm both tasks ────────────────────────────────────────────────────────
Write-Host ""
Write-Host "Scheduled tasks:"
Get-ScheduledTask | Where-Object { $_.TaskName -like "AShop_*" } |
    Select-Object TaskName, State |
    Format-Table -AutoSize
