$action = New-ScheduledTaskAction -Execute "C:\Users\NYOUSIF\Desktop\AShop_Dashboard\RUN_MORNING_MEETING.bat"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At "7:45AM"
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -StartWhenAvailable
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName "AShop_MorningMeeting_Update" -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force
Write-Host "Done"
Get-ScheduledTask -TaskName "AShop_MorningMeeting_Update" | Select-Object TaskName, State
