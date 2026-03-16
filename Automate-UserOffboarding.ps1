# ===================================================================
# AD USER OFFBOARD AUTOMATION
# Fully Automatic Offboarding + Smart OU Handling
# ===================================================================

Import-Module ActiveDirectory
. "$PSScriptRoot\Dashboard-Helper.ps1"

# -------------------- CONFIGURATION -------------------------------

$FreshserviceDomain = "desicrew"
$FreshserviceAPIKey = "QlRPOby1GBAPxBAthIAW"
$FreshserviceURL = "https://$FreshserviceDomain.freshservice.com"

$OffboardOU = "OU=Offboarded Users,DC=desicrew,DC=in"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "AD USER OFFBOARD AUTOMATION STARTED: $(Get-Date)" -ForegroundColor White
Write-Host "==================================================" -ForegroundColor Cyan

# -------------------- ENSURE OFFBOARD OU EXISTS --------------------

try {
    $exists = Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$OffboardOU'" -ErrorAction Stop
    Write-Host "✔ Offboard OU found: $OffboardOU" -ForegroundColor Green
}
catch {
    Write-Host "⚠ Offboard OU missing — creating..." -ForegroundColor Yellow
    New-ADOrganizationalUnit -Name "Offboarded Users" -Path "DC=desicrew,DC=in" -ProtectedFromAccidentalDeletion $false
    Write-Host "✔ Offboard OU created." -ForegroundColor Green
}

# ===================================================================
# FUNCTION: FETCH FRESHSERVICE TICKETS
# ===================================================================
function Get-FreshserviceTickets {
    try {
        $auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($FreshserviceAPIKey + ":X")))

        $headers = @{
            "Authorization" = "Basic $auth"
            "Content-Type"  = "application/json"
        }

        $uri = "$FreshserviceURL/api/v2/tickets?include=requester"
        return (Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 30).tickets
    }
    catch {
        Write-Host "❌ Failed Freshservice API connection: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# ===================================================================
# FUNCTION: FILTER OFFBOARD TICKETS
# ===================================================================
function Get-OffboardTickets {
    param($AllTickets)

    $patterns = @(
        "*offboard*",
        "*remove user*",
        "*disable account*",
        "*separation*",
        "*exit process*",
        "*user left*",
        "*employee left*",
        "*resigned*",
        "*deactivate user*"
    )

    $result = @()

    foreach ($ticket in $AllTickets) {
        if ($ticket.status -ne 2) { continue }  # Only OPEN tickets

        $subject = $ticket.subject.ToLower()
        $desc = $ticket.description.ToLower()

        foreach ($p in $patterns) {
            if ($subject -like $p -or $desc -like $p) {
                Write-Host "MATCHED Offboard Ticket: #$($ticket.id)" -ForegroundColor Yellow
                $result += $ticket
                break
            }
        }
    }

    return $result
}

# ===================================================================
# FUNCTION: EXTRACT USERNAME FROM TICKET
# ===================================================================
function Extract-OffboardUser {
    param($Ticket)

    $desc = $Ticket.description

    # Username:
    if ($desc -match "(?i)username\s*[:=]\s*([A-Za-z0-9._-]+)") {
        return $matches[1]
    }

    # Email:
    if ($desc -match "(?i)email\s*[:=]\s*([A-Za-z0-9._%+-]+)") {
        return $matches[1].Split("@")[0]  # Convert email → username
    }

    # Name → Try first part match in AD
    if ($desc -match "(?i)name\s*[:=]\s*([A-Za-z]+(?:\s+[A-Za-z]+)*)") {
        $fullName = $matches[1]
        Write-Host "Searching AD for DisplayName: $fullName" -ForegroundColor Cyan

        $user = Get-ADUser -Filter "DisplayName -eq '$fullName'" -ErrorAction SilentlyContinue
        if ($user) { return $user.SamAccountName }
    }

    Write-Host "❌ Unable to extract username for ticket #$($Ticket.id)" -ForegroundColor Red
    return $null
}

# ===================================================================
# FUNCTION: OFFBOARD AD USER
# ===================================================================
function Offboard-ADUser {
    param($userSam, $Ticket)

    try {
        $user = Get-ADUser -Filter "SamAccountName -eq '$userSam'" -Properties * -ErrorAction Stop

        Write-Host "✔ User Located: $($user.SamAccountName)" -ForegroundColor Green

        # ---------------------------------------------------------------
        # Remove from all groups except Domain Users
        # ---------------------------------------------------------------
        Write-Host "Removing AD groups..." -ForegroundColor Yellow
        $groups = Get-ADPrincipalGroupMembership $user | Where-Object { $_.Name -ne "Domain Users" }

        foreach ($g in $groups) {
            Remove-ADGroupMember -Identity $g.Name -Members $userSam -Confirm:$false -ErrorAction SilentlyContinue
        }

        # ---------------------------------------------------------------
        # Disable Account
        # ---------------------------------------------------------------
        Disable-ADAccount -Identity $userSam
        Write-Host "✔ Account Disabled" -ForegroundColor Yellow

        # ---------------------------------------------------------------
        # Set Temporary Disabled Password
        # ---------------------------------------------------------------
        $tempPass = "Disable@$(Get-Random -Minimum 1000 -Maximum 9999)"
        Set-ADAccountPassword -Identity $userSam -NewPassword (ConvertTo-SecureString $tempPass -AsPlainText -Force) -Reset
        Write-Host "✔ Password Reset" -ForegroundColor Yellow

        # ---------------------------------------------------------------
        # Prepare object for moving (remove protected flag)
        # ---------------------------------------------------------------
        Write-Host "Removing protection flag..." -ForegroundColor DarkYellow
        Set-ADObject -Identity $user.DistinguishedName -ProtectedFromAccidentalDeletion $false -ErrorAction SilentlyContinue

        # ---------------------------------------------------------------
        # MOVE USER TO OFFBOARD OU
        # ---------------------------------------------------------------
        Write-Host "Moving user to Offboard OU..." -ForegroundColor Yellow

        Move-ADObject -Identity $user.DistinguishedName -TargetPath $OffboardOU -ErrorAction Stop

        Write-Host "✔ User moved to: $OffboardOU" -ForegroundColor Green

        return @{
            Success      = $true
            Username     = $userSam
            DisplayName  = $user.DisplayName
            TempPassword = $tempPass
            TicketID     = $Ticket.id
        }
    }
    catch {
        Write-Host "❌ Offboarding FAILED: $($_.Exception.Message)" -ForegroundColor Red

        return @{
            Success  = $false
            Error    = $_.Exception.Message
            TicketID = $Ticket.id
        }
    }
}

# ===================================================================
# FUNCTION: UPDATE FRESHSERVICE TICKET
# ===================================================================
function Update-FreshserviceOffboard {
    param($Result)

    try {
        $auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($FreshserviceAPIKey + ":X")))

        $headers = @{
            "Authorization" = "Basic $auth"
            "Content-Type"  = "application/json"
        }

        $note = @"
🛑 **ACCOUNT OFFBOARDING COMPLETED**

**Employee:** $($Result.DisplayName)  
**Username:** $($Result.Username)  
**Temporary Disabled Password:** $($Result.TempPassword)  
**Status:** Disabled  
**AD Groups:** Removed  
**Moved to:** Offboarded Users OU  

Ticket auto-closed by automation.

Regards,  
DC IT Helpdesk
"@

        # Add note
        Invoke-RestMethod -Uri "$FreshserviceURL/api/v2/tickets/$($Result.TicketID)/notes" `
            -Method POST -Headers $headers -Body (@{ body = $note; private = $false } | ConvertTo-Json)

        # Close Ticket
        Invoke-RestMethod -Uri "$FreshserviceURL/api/v2/tickets/$($Result.TicketID)" `
            -Method PUT -Headers $headers -Body (@{ status = 5 } | ConvertTo-Json)

        Write-Host "✔ Freshservice Updated + Ticket Closed" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Freshservice Update Failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# ===================================================================
# MAIN EXECUTION
# ===================================================================

$tickets = Get-FreshserviceTickets

if (-not $tickets) {
    Write-Host "❌ Cannot continue - Freshservice API failed" -ForegroundColor Red
    exit
}

$offboardTickets = Get-OffboardTickets -AllTickets $tickets

foreach ($ticket in $offboardTickets) {

    Write-Host "`nProcessing Ticket #$($ticket.id)" -ForegroundColor Cyan

    $userSam = Extract-OffboardUser -Ticket $ticket

    if ($null -eq $userSam) {
        Write-Host "❌ Skipping ticket — No valid username found." -ForegroundColor Red
        continue
    }

    $result = Offboard-ADUser -userSam $userSam -Ticket $ticket

    if ($result.Success) {
        Update-FreshserviceOffboard -Result $result
        Send-DashboardLog -Service "Offboarding" -Status "success" -Message "Offboarded user: $($result.Username)"
    }
    else {
        Send-DashboardLog -Service "Offboarding" -Status "error" -Message "Failed to offboard: $($userSam)"
    }
}

& "$PSScriptRoot\Generate-Dashboard.ps1"

Write-Host "`n============== OFFBOARDING COMPLETED ==============" -ForegroundColor Cyan
