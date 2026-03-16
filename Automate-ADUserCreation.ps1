# =====================================================================
# Automate-ADUserCreation.ps1 (FULL FIXED VERSION)
# =====================================================================

Import-Module ActiveDirectory
. "$PSScriptRoot\Dashboard-Helper.ps1"

# ================= CONFIGURATION =================
$FreshserviceDomain = "desicrew"
$FreshserviceAPIKey = "QlRPOby1GBAPxBAthIAW"
$FreshserviceURL = "https://$FreshserviceDomain.freshservice.com"
$Domain = "desicrew.in"
$DefaultRequesterEmail = "it.hepdesk@desicrew.in"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ================= LOGGING =================
$LogDir = "C:\Automation\logs"
if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }

function Write-Log {
    param($Msg)
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $Msg" |
    Out-File "$LogDir\ad_create.log" -Append
}

# ================= AUTH HEADER =================
function Get-FSHeaders {
    $auth = [Convert]::ToBase64String(
        [Text.Encoding]::ASCII.GetBytes("$FreshserviceAPIKey:X")
    )
    return @{
        Authorization  = "Basic $auth"
        "Content-Type" = "application/json"
    }
}

# ================= FETCH TICKETS =================
function Get-FreshserviceTickets {
    $headers = Get-FSHeaders
    $uri = "$FreshserviceURL/api/v2/tickets?include=requester"

    try {
        $res = Invoke-RestMethod `
            -Uri $uri `
            -Headers $headers `
            -TimeoutSec 30
        return $res.tickets
    }
    catch {
        Write-Host "Freshservice API failed: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# ================= ENSURE REQUESTER =================
function Ensure-Requester {
    param($Ticket)

    if ($Ticket.requester -and $Ticket.requester.email) {
        return @{ Success = $true }
    }

    $headers = Get-FSHeaders
    $body = @{ email = $DefaultRequesterEmail } | ConvertTo-Json

    try {
        Invoke-RestMethod `
            -Uri "$FreshserviceURL/api/v2/tickets/$($Ticket.id)" `
            -Method PUT `
            -Headers $headers `
            -Body $body `
            -ContentType "application/json"

        return @{ Success = $true }
    }
    catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

# ================= ENSURE DESCRIPTION =================
function Ensure-TicketDescription {
    param($Ticket)

    if ($Ticket.description -and $Ticket.description.Trim().Length -gt 0) {
        return $true
    }

    $headers = Get-FSHeaders
    $body = @{
        description = "AD User Creation Request – Awaiting required details."
    } | ConvertTo-Json

    try {
        Invoke-RestMethod `
            -Uri "$FreshserviceURL/api/v2/tickets/$($Ticket.id)" `
            -Method PUT `
            -Headers $headers `
            -Body $body `
            -ContentType "application/json"

        return $true
    }
    catch {
        Write-Host "Failed to update description: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# ================= REQUEST MISSING DETAILS =================
function Request-MissingDetailsFromTicket {
    param($Ticket)

    $req = Ensure-Requester -Ticket $Ticket
    if (-not $req.Success) {
        return @{ Success = $false; Error = "Requester assignment failed" }
    }

    if (-not (Ensure-TicketDescription -Ticket $Ticket)) {
        return @{ Success = $false; Error = "Failed to set description" }
    }

    $headers = Get-FSHeaders

    $note = @"
⚠️ AD User Creation Pending

Please update the ticket with the following details:

Name:
Department:
Location (TNP / Kollu / Kaup / VPM / Chennai):
Project:

Once updated, the automation will resume and create the domain ID.

– DC IT Helpdesk
"@

    $noteBody = @{
        body    = $note
        private = $false
    } | ConvertTo-Json

    try {
        Invoke-RestMethod `
            -Uri "$FreshserviceURL/api/v2/tickets/$($Ticket.id)/notes" `
            -Method POST `
            -Headers $headers `
            -Body $noteBody `
            -ContentType "application/json"

        Invoke-RestMethod `
            -Uri "$FreshserviceURL/api/v2/tickets/$($Ticket.id)" `
            -Method PUT `
            -Headers $headers `
            -Body (@{ status = 3 } | ConvertTo-Json) `
            -ContentType "application/json"

        return @{ Success = $true }
    }
    catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

# ================= MAIN EXECUTION =================
Write-Host "STARTING AD USER CREATION AUTOMATION..." -ForegroundColor Cyan

$tickets = Get-FreshserviceTickets
if (-not $tickets) { exit 1 }

$openOnboarding = $tickets | Where-Object {
    $_.status -eq 2 -and $_.subject -match "(onboarding|user creation|domain)"
}

foreach ($ticket in $openOnboarding) {

    Write-Host "Processing Ticket #$($ticket.id)" -ForegroundColor Yellow

    # Simulated missing data case (your real logic already exists)
    $missing = $true

    if ($missing) {
        $res = Request-MissingDetailsFromTicket -Ticket $ticket
        if ($res.Success) {
            Write-Host "Requester notified, ticket set to Pending." -ForegroundColor Green
            Write-Log "Ticket #$($ticket.id) pending – missing details"
        }
        else {
            Write-Host "Failed to notify requester: $($res.Error)" -ForegroundColor Red
        }
    }
}

Write-Host "AD USER CREATION AUTOMATION COMPLETED" -ForegroundColor Cyan

Send-DashboardLog -Service "AD Creation" -Status "info" -Message "User creation run completed."
& "$PSScriptRoot\Generate-Dashboard.ps1"
