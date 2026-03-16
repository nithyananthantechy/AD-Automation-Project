# ======================================================================
#  CLEAN AD DASHBOARD GENERATOR (PUSH TO NEW API)
# ======================================================================

$DashboardAPI = "http://localhost:3030/api/update"
$UpdateToken = "desicrew-update-token"

Write-Host "`n🎨 GENERATING AD DASHBOARD" -ForegroundColor Cyan
Write-Host "=========================================="

# AD Statistics
try {
    $allUsers = Get-ADUser -Filter *
    $totalUsers = $allUsers.Count
    $enabledUsers = ($allUsers | Where-Object Enabled).Count
    $disabledUsers = $totalUsers - $enabledUsers
    $locked = (Search-ADAccount -LockedOut).Count
}
catch {
    Write-Host "❌ Failed to fetch AD stats: $($_.Exception.Message)" -ForegroundColor Red
    $totalUsers = 0; $enabledUsers = 0; $disabledUsers = 0; $locked = 0
}

# System Info
$serverName = $env:COMPUTERNAME
$uptime = (Get-Date) - (Get-CimInstance Win32_OperatingSystem).LastBootUpTime |
ForEach-Object { "{0}d {1}h {2}m" -f $_.Days, $_.Hours, $_.Minutes }

# ----------------------------------------------------------------------
# PUSH DATA TO API
# ----------------------------------------------------------------------

$headers = @{
    "Content-Type"   = "application/json"
    "x-update-token" = $UpdateToken
}

# Push Stats
$statsPayload = @{
    type    = "stats"
    payload = @{
        serverName    = $serverName
        uptime        = $uptime
        totalUsers    = $totalUsers
        enabledUsers  = $enabledUsers
        disabledUsers = $disabledUsers
        lockedUsers   = $locked
    }
}

try {
    Invoke-RestMethod -Uri $DashboardAPI -Method Post -Headers $headers -Body ($statsPayload | ConvertTo-Json)
    Write-Host "✔ AD Stats pushed to Dashboard" -ForegroundColor Green
}
catch {
    Write-Host "❌ Failed to push stats: $($_.Exception.Message)" -ForegroundColor Red
}

# Push Dashboard Generation Log
$logPayload = @{
    type    = "log"
    payload = @{
        service = "Dashboard"
        status  = "success"
        message = "Dashboard stats updated successfully"
    }
}

try {
    Invoke-RestMethod -Uri $DashboardAPI -Method Post -Headers $headers -Body ($logPayload | ConvertTo-Json)
    Write-Host "✔ Dashboard Log pushed" -ForegroundColor Green
}
catch {
    Write-Host "❌ Failed to push log: $($_.Exception.Message)" -ForegroundColor Red
}
