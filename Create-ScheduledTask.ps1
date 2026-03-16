# Run as Administrator
# Save as: C:\Automation\Create-ScheduledTask.ps1

$TaskName = "Freshservice-AD-Automation"
$BatchPath = "C:\Automation\Run-Automation.bat"

# Remove old task if exists
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

# Create Task Action
$Action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$BatchPath`""

# Run every 1 minute
$Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
            -RepetitionInterval (New-TimeSpan -Minutes 1)

# Task settings
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries `
    -StartWhenAvailable -MultipleInstances IgnoreNew

# Run using domain automation account
Register-ScheduledTask -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -RunLevel Highest `
    -User "DESICREW\svcAutomation" `
    -Password (Read-Host "Enter svcAutomation Password" -AsSecureString)

Write-Host "Task created successfully and will run every 1 min as svcAutomation." -ForegroundColor Green
