# Save as: C:\Automation\Run-All-Automations.ps1
# UPDATED VERSION - Includes Domain Unlock Automation + Mutex Lock
# - Supports AutoRun mode for Task Scheduler
# - Supports manual menu mode
# - Uses a global named mutex to prevent overlapping runs
# - Safe cleanup with try/finally

param(
    [switch]$AutoRun = $false  # Parameter for scheduled task
)

# Name of the global mutex used to prevent multiple instances
$mutexName = "Global\FreshserviceAutomationMutex"
$mutex = $null

try {
    # ==============================================================
    # Prevent multiple instances from running simultaneously
    # ==============================================================

    $createdNew = $false
    $mutex = New-Object System.Threading.Mutex($false, $mutexName, [ref]$createdNew)

    if (-not $mutex.WaitOne(0, $false)) {
        Write-Host "Another instance is already running. Exiting..." -ForegroundColor Yellow
        return
    }

    Write-Host "MASTER AUTOMATION RUNNER" -ForegroundColor Cyan
    Write-Host "Started: $(Get-Date)" -ForegroundColor White

    # ==============================================================
    # MAIN FUNCTIONS
    # ==============================================================

    function Run-AllAutomations {
        Write-Host "RUNNING ALL AUTOMATIONS" -ForegroundColor Cyan
        
        # 1. AD User Creation
        Write-Host "1. RUNNING AD USER CREATION AUTOMATION..." -ForegroundColor Yellow
        try {
            & "C:\Automation\Automate-ADUserCreation.ps1"
            Write-Host "AD User Creation Completed" -ForegroundColor Green
        } catch {
            Write-Host "AD User Creation Failed: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # 2. Password Reset
        Write-Host "2. RUNNING PASSWORD RESET AUTOMATION..." -ForegroundColor Yellow
        try {
            & "C:\Automation\Automate-PasswordReset.ps1"
            Write-Host "Password Reset Completed" -ForegroundColor Green
        } catch {
            Write-Host "Password Reset Failed: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # 3. User Offboarding
        Write-Host "3. RUNNING USER OFFBOARDING AUTOMATION..." -ForegroundColor Yellow
        try {
            & "C:\Automation\Automate-UserOffboarding.ps1"
            Write-Host "User Offboarding Completed" -ForegroundColor Green
        } catch {
            Write-Host "User Offboarding Failed: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # 4. Domain Unlock
        Write-Host "4. RUNNING DOMAIN UNLOCK AUTOMATION..." -ForegroundColor Yellow
        try {
            & "C:\Automation\Automate-DomainUnlock.ps1"
            Write-Host "Domain Unlock Completed" -ForegroundColor Green
        } catch {
            Write-Host "Domain Unlock Failed: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # 5. Generate dashboard AFTER running all automations
        Write-Host "GENERATING DASHBOARD..." -ForegroundColor Yellow
        try {
            & "C:\Automation\Generate-Dashboard.ps1"
            Write-Host "Dashboard Updated" -ForegroundColor Green
        } catch {
            Write-Host "Dashboard Update Failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        Write-Host "ALL AUTOMATIONS COMPLETED: $(Get-Date)" -ForegroundColor Green
    }

    function Show-Menu {
        Write-Host "SELECT AN OPTION:" -ForegroundColor Yellow
        Write-Host "  1. Run All Automations" -ForegroundColor Green
        Write-Host "  2. Run AD User Creation Only" -ForegroundColor White
        Write-Host "  3. Run Password Reset Only" -ForegroundColor White
        Write-Host "  4. Run User Offboarding Only" -ForegroundColor White
        Write-Host "  5. Run Domain Unlock Only" -ForegroundColor White
        Write-Host "  6. Show Current Status" -ForegroundColor White
        Write-Host "  7. Exit" -ForegroundColor White
    }

    # ==============================================================
    # ENTRY POINT
    # ==============================================================

    if ($AutoRun) {
        # Called from Task Scheduler
        Write-Host "AUTO-RUN MODE: Running all automations automatically..." -ForegroundColor Cyan
        Run-AllAutomations
    } else {
        # Manual mode with menu
        Write-Host "MANUAL MODE: Showing menu..." -ForegroundColor Cyan
        
        do {
            Show-Menu
            $choice = Read-Host "Enter your choice [1]"
            
            if ([string]::IsNullOrWhiteSpace($choice)) { $choice = "1" }
            
            switch ($choice) {
                "1" { 
                    Run-AllAutomations
                    break
                }
                "2" { 
                    Write-Host "RUNNING AD USER CREATION ONLY..." -ForegroundColor Cyan
                    try {
                        & "C:\Automation\Automate-ADUserCreation.ps1"
                    } catch {
                        Write-Host "AD User Creation Failed: $($_.Exception.Message)" -ForegroundColor Red
                    }
                    break
                }
                "3" { 
                    Write-Host "RUNNING PASSWORD RESET ONLY..." -ForegroundColor Cyan
                    try {
                        & "C:\Automation\Automate-PasswordReset.ps1"
                    } catch {
                        Write-Host "Password Reset Failed: $($_.Exception.Message)" -ForegroundColor Red
                    }
                    break
                }
                "4" { 
                    Write-Host "RUNNING USER OFFBOARDING ONLY..." -ForegroundColor Cyan
                    try {
                        & "C:\Automation\Automate-UserOffboarding.ps1"
                    } catch {
                        Write-Host "User Offboarding Failed: $($_.Exception.Message)" -ForegroundColor Red
                    }
                    break
                }
                "5" { 
                    Write-Host "RUNNING DOMAIN UNLOCK ONLY..." -ForegroundColor Cyan
                    try {
                        & "C:\Automation\Automate-DomainUnlock.ps1"
                    } catch {
                        Write-Host "Domain Unlock Failed: $($_.Exception.Message)" -ForegroundColor Red
                    }
                    break
                }
                "6" { 
                    Write-Host "CURRENT STATUS:" -ForegroundColor Cyan
                    Write-Host "  All automations are ready" -ForegroundColor White
                    Write-Host "  Last run (local time): $(Get-Date)" -ForegroundColor White
                    Write-Host "  Scheduled task name: Freshservice-AD-Automation" -ForegroundColor White
                    Write-Host "  Schedule: Configured in Windows Task Scheduler" -ForegroundColor White
                    Write-Host "  Available Automations:" -ForegroundColor White
                    Write-Host "    - AD User Creation" -ForegroundColor Gray
                    Write-Host "    - Password Reset" -ForegroundColor Gray
                    Write-Host "    - User Offboarding" -ForegroundColor Gray
                    Write-Host "    - Domain Unlock" -ForegroundColor Gray
                    Write-Host "  Environment: Desicrew" -ForegroundColor Gray
                    Write-Host "  Domain: desicrew.in" -ForegroundColor Gray
                    Write-Host "  Freshservice: desicrew.freshservice.com" -ForegroundColor Gray
                    break
                }
                "7" { 
                    Write-Host "Exiting..." -ForegroundColor Yellow
                    break
                }
                default {
                    Write-Host "Invalid choice. Please select 1-7." -ForegroundColor Red
                }
            }
            
            if ($choice -ne "7") {
                $continue = Read-Host "Run another operation? (Y/N) [N]"
                if ($continue -notlike "Y*" -and $continue -notlike "y*") {
                    Write-Host "Exiting..." -ForegroundColor Yellow
                    break
                }
            }
            
        } while ($choice -ne "7")
    }

} finally {
    # ==============================================================
    # Release mutex at end (even if there was an error)
    # ==============================================================

    if ($mutex -ne $null) {
        try {
            $mutex.ReleaseMutex()
        } catch {
            # Ignore if already released or not owned
        }
        $mutex.Dispose()
    }

    Write-Host "MASTER AUTOMATION RUNNER COMPLETED: $(Get-Date)" -ForegroundColor Cyan
}
