@echo off
cd /d C:\Automation

REM ----- Prevent multiple executions -----
if exist C:\Automation\automation.lock (
    echo [%date% %time%] Scheduler attempted to run but previous run still processing >> C:\Automation\schedule.log
    exit /b
)

echo [%date% %time%] Starting automation >> C:\Automation\schedule.log
type nul > C:\Automation\automation.lock

powershell.exe -ExecutionPolicy Bypass -File "C:\Automation\Run-All-Automations.ps1" -AutoRun >> C:\Automation\schedule.log 2>&1

del C:\Automation\automation.lock
echo [%date% %time%] Automation completed >> C:\Automation\schedule.log
