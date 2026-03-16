# ======================================================================
#  DASHBOARD INTEGRATION HELPER
# ======================================================================

$DashboardAPI = "http://localhost:3030/api/update"
$UpdateToken = "desicrew-update-token"

function Send-DashboardLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Service,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet("success", "error", "warning", "info")]
        [string]$Status,
        
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $headers = @{
        "Content-Type"   = "application/json"
        "x-update-token" = $UpdateToken
    }

    $payload = @{
        type    = "log"
        payload = @{
            service = $Service
            status  = $Status
            message = $Message
        }
    }

    try {
        Invoke-RestMethod -Uri $DashboardAPI -Method Post -Headers $headers -Body ($payload | ConvertTo-Json)
        Write-Host "✔ Dashboard updated: [$Service] $Message" -ForegroundColor Gray
    }
    catch {
        Write-Host "⚠ Dashboard update failed: $($_.Exception.Message)" -ForegroundColor DarkGray
    }
}

Export-ModuleMember -Function Send-DashboardLog
