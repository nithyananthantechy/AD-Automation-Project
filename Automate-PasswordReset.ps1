# Save as: C:\Automation\Automate-PasswordReset.ps1
# UPDATED VERSION - Closes tickets after completion
# ADDED MORE SUBJECT PATTERNS AND IT MANAGER APPROVAL WORKFLOW
# UPDATED: Skip application (Zoho/Gmail/O365/etc.) password resets – only reset AD/domain passwords

Import-Module ActiveDirectory
. "$PSScriptRoot\Dashboard-Helper.ps1"

# CONFIGURATION
$FreshserviceDomain = "desicrew"
$FreshserviceAPIKey = "QlRPOby1GBAPxBAthIAW"
$FreshserviceURL = "https://$FreshserviceDomain.freshservice.com"
$Domain = "desicrew.in"
$ITManagerEmail = "balachandran@desicrew.in"

# IT Support Team Employee IDs - Require Manager Approval
$ITSupportTeam = @(
    "DK0298",
    "DC2531",
    "DC5000",
    "DC2037",
    "DC4944",
    "DC1651",
    "DC1928"
)

# APPLICATION PASSWORD KEYWORDS (tickets with these are NOT AD password resets)
$AppPasswordKeywords = @(
    "zoho",
    "gmail",
    "google workspace",
    "google account",
    "office 365",
    "microsoft 365",
    "m365",
    "o365",
    "outlook.com",
    "outlook",
    "webmail",
    "slack",
    "teams",
    "zoom"
)

# FIX: Force TLS 1.2 for secure connection
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Host "PASSWORD RESET AUTOMATION" -ForegroundColor Cyan
Write-Host "Started: $(Get-Date)" -ForegroundColor White

# FUNCTION: Check if employee is in IT Support Team
function Test-ITSupportTeamMember {
    param($EmployeeID)

    $EmployeeID = $EmployeeID.ToUpper()
    return $ITSupportTeam -contains $EmployeeID
}

# FUNCTION: Send Approval Request to IT Manager
function Send-ITManagerApproval {
    param($Ticket, $EmployeeID, $EmployeeName)

    try {
        Write-Host "IT SUPPORT TEAM MEMBER DETECTED: $EmployeeID" -ForegroundColor Yellow
        Write-Host "Sending approval request to IT Manager: $ITManagerEmail" -ForegroundColor Cyan

        # Prepare approval request details
        $ticketLink = "$FreshserviceURL/helpdesk/tickets/$($Ticket.id)"
        $currentTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # Create approval request in Freshservice
        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($FreshserviceAPIKey + ":X")))
        $headers = @{
            "Authorization" = "Basic $base64AuthInfo"
            "Content-Type"  = "application/json"
        }

        # Add approval required note to ticket
        $approvalNote = @"
🔐 **APPROVAL REQUIRED - IT SUPPORT TEAM MEMBER**

**Request Details:**
- Ticket ID: $($Ticket.id)
- Employee ID: $EmployeeID
- Employee Name: $EmployeeName
- Request Type: Password Reset
- Requested By: IT Support Team Member
- Date/Time: $currentTime

**Approval Required From:** $ITManagerEmail

**Next Steps:**
1. IT Manager must review this request
2. If approved, IT Manager should reply with "APPROVED"
3. If denied, IT Manager should reply with "DENIED"

**Ticket Link:** $ticketLink

**Note:** This is an automated approval request. Password reset will proceed only after explicit approval.
"@

        $noteBody = @{
            body    = $approvalNote
            private = $false
        } | ConvertTo-Json

        $noteUri = "$FreshserviceURL/api/v2/tickets/$($Ticket.id)/notes"
        $noteResponse = Invoke-RestMethod -Uri $noteUri -Method Post -Headers $headers -Body $noteBody

        Write-Host "Approval request note added to ticket" -ForegroundColor Green

        # Set ticket to pending approval status (status 3 = Pending)
        $body = @{
            status   = 3
            priority = 2  # High priority
        } | ConvertTo-Json

        $uri = "$FreshserviceURL/api/v2/tickets/$($Ticket.id)"
        $response = Invoke-RestMethod -Uri $uri -Method Put -Headers $headers -Body $body

        Write-Host "Ticket status set to PENDING APPROVAL" -ForegroundColor Yellow
        Write-Host "Waiting for IT Manager approval..." -ForegroundColor Yellow

        # Check for existing approval in ticket conversation
        $approvalGranted = Check-ExistingApproval -Ticket $Ticket

        if ($approvalGranted) {
            Write-Host "Existing approval found! Proceeding with password reset..." -ForegroundColor Green
            return $true
        }

        Write-Host "Approval workflow initiated. Ticket #$($Ticket.id) awaits IT Manager approval." -ForegroundColor Cyan
        return $false

    }
    catch {
        Write-Host "Failed to send approval request: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# FUNCTION: Check for existing approval in ticket conversation
function Check-ExistingApproval {
    param($Ticket)

    try {
        Write-Host "Checking for existing approvals in ticket conversation..." -ForegroundColor Gray

        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($FreshserviceAPIKey + ":X")))
        $headers = @{
            "Authorization" = "Basic $base64AuthInfo"
            "Content-Type"  = "application/json"
        }

        # Get ticket conversation
        $conversationUri = "$FreshserviceURL/api/v2/tickets/$($Ticket.id)/conversations"
        $response = Invoke-RestMethod -Uri $conversationUri -Headers $headers

        # Check each conversation entry for approval
        foreach ($conversation in $response.conversations) {
            if ($conversation.body_text -and $conversation.body_text -match "(?i)approved|yes|go ahead|proceed|reset") {
                # Check if it's from IT Manager email
                if ($conversation.from_email -eq $ITManagerEmail) {
                    Write-Host "Found approval from IT Manager: $($conversation.body_text)" -ForegroundColor Green
                    return $true
                }
            }
        }

        return $false

    }
    catch {
        Write-Host "Error checking conversation: $($_.Exception.Message)" -ForegroundColor Yellow
        return $false
    }
}

# FUNCTION: Monitor for approval
function Wait-ForApproval {
    param($Ticket, $TimeoutMinutes = 30)

    Write-Host "Monitoring for IT Manager approval (timeout: $TimeoutMinutes minutes)..." -ForegroundColor Yellow

    $startTime = Get-Date
    $timeout = $startTime.AddMinutes($TimeoutMinutes)

    while ((Get-Date) -lt $timeout) {
        $elapsed = [int]((Get-Date) - $startTime).TotalMinutes
        Write-Host "Elapsed time: $elapsed minutes" -ForegroundColor Gray

        $approved = Check-ExistingApproval -Ticket $Ticket

        if ($approved) {
            Write-Host "APPROVAL RECEIVED! Proceeding with password reset..." -ForegroundColor Green
            return $true
        }

        Write-Host "Approval not yet received. Waiting 1 minute..." -ForegroundColor Yellow
        Start-Sleep -Seconds 60

        # Update ticket with waiting status
        if ($elapsed % 5 -eq 0) {
            Update-WaitingStatus -Ticket $Ticket -ElapsedMinutes $elapsed
        }
    }

    Write-Host "Approval timeout reached. Ticket remains pending." -ForegroundColor Red
    return $false
}

# FUNCTION: Update waiting status
function Update-WaitingStatus {
    param($Ticket, $ElapsedMinutes)

    try {
        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($FreshserviceAPIKey + ":X")))
        $headers = @{
            "Authorization" = "Basic $base64AuthInfo"
            "Content-Type"  = "application/json"
        }

        $statusNote = @"
⏳ **Awaiting IT Manager Approval** - $ElapsedMinutes minutes elapsed

Still waiting for approval from $ITManagerEmail for IT Support Team member password reset.

Please review and approve by replying "APPROVED" to this ticket.
"@

        $privateNote = @{
            body    = $statusNote
            private = $true
        } | ConvertTo-Json

        $noteUri = "$FreshserviceURL/api/v2/tickets/$($Ticket.id)/notes"
        Invoke-RestMethod -Uri $noteUri -Method Post -Headers $headers -Body $privateNote -ErrorAction SilentlyContinue

    }
    catch {
        # Silent fail for status updates
    }
}

# FUNCTION: Process IT Support Team Member with Approval
function Process-ITSupportMember {
    param($Ticket, $EmployeeID, $EmployeeName)

    Write-Host "=== IT SUPPORT TEAM MEMBER WORKFLOW ===" -ForegroundColor Yellow

    # Check for existing approval first
    $existingApproval = Check-ExistingApproval -Ticket $Ticket

    if ($existingApproval) {
        Write-Host "Existing approval found. Proceeding with password reset..." -ForegroundColor Green
        return $true
    }

    # Send approval request
    $approvalSent = Send-ITManagerApproval -Ticket $Ticket -EmployeeID $EmployeeID -EmployeeName $EmployeeName

    if (-not $approvalSent) {
        Write-Host "Failed to send approval request. Ticket requires manual processing." -ForegroundColor Red
        return $false
    }

    # Wait for approval
    $approved = Wait-ForApproval -Ticket $Ticket -TimeoutMinutes 30

    if (-not $approved) {
        Write-Host "Approval not received within timeout period." -ForegroundColor Red
        Write-Host "Ticket #$($Ticket.id) requires manual intervention." -ForegroundColor Yellow

        # Add timeout note
        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($FreshserviceAPIKey + ":X")))
        $headers = @{
            "Authorization" = "Basic $base64AuthInfo"
            "Content-Type"  = "application/json"
        }

        $timeoutNote = @"
⏰ **APPROVAL TIMEOUT**

The approval request for IT Support Team member $EmployeeID has timed out (30 minutes).

**Required Action:**
1. IT Manager ($ITManagerEmail) must manually review and approve
2. Reply with "APPROVED" to proceed with automated password reset
3. Or process manually via Active Directory

Ticket remains in pending state for manual handling.
"@

        $noteBody = @{
            body    = $timeoutNote
            private = $false
        } | ConvertTo-Json

        $noteUri = "$FreshserviceURL/api/v2/tickets/$($Ticket.id)/notes"
        Invoke-RestMethod -Uri $noteUri -Method Post -Headers $headers -Body $noteBody

        return $false
    }

    return $true
}

# FUNCTION: Get Freshservice Tickets
function Get-FreshserviceTickets {
    try {
        Write-Host "Connecting to Freshservice..." -ForegroundColor Yellow

        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($FreshserviceAPIKey + ":X")))
        $headers = @{
            "Authorization" = "Basic $base64AuthInfo"
            "Content-Type"  = "application/json"
        }

        $uri = "$FreshserviceURL/api/v2/tickets?include=requester"
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 30

        Write-Host "API Connection Successful!" -ForegroundColor Green
        Write-Host "Found $($response.tickets.Count) total tickets" -ForegroundColor White

        return $response.tickets

    }
    catch {
        Write-Host "API Connection Failed: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# FUNCTION: Get Password Reset Tickets - UPDATED WITH APP SKIP FILTER
function Get-PasswordResetTickets {
    param($AllTickets)

    Write-Host "Checking for Password Reset tickets..." -ForegroundColor Yellow

    $openTickets = @()
    foreach ($ticket in $AllTickets) {
        if ($ticket.status -eq 2) {
            $openTickets += $ticket
        }
    }

    Write-Host "Found $($openTickets.Count) OPEN tickets" -ForegroundColor Green

    if ($openTickets.Count -gt 0) {
        Write-Host "Open Tickets:" -ForegroundColor White
        foreach ($ticket in $openTickets) {
            Write-Host "ID: $($ticket.id) | Subject: $($ticket.subject)" -ForegroundColor Gray
        }
    }
    else {
        Write-Host "No open tickets found" -ForegroundColor Yellow
        return @()
    }

    $passwordResetTickets = @()
    foreach ($ticket in $openTickets) {
        $subject = $ticket.subject.ToLower()
        $description = if ($ticket.description) { $ticket.description.ToLower() } else { "" }

        $subjectPatterns = @(
            "*password reset*",
            "*reset password*",
            "*forgot password*",
            "*password help*",
            "*password change*",
            "*change password*",
            "*password expired*",
            "*expired password*",
            "*reset - *",
            "*password - *",
            "*forgot - *",
            "*change - *"
        )

        $descriptionPatterns = @(
            "*password reset*",
            "*reset password*",
            "*forgot password*",
            "*password help*",
            "*password change*"
        )

        $isPasswordResetTicket = $false

        # Check subject patterns
        foreach ($pattern in $subjectPatterns) {
            if ($subject -like $pattern) {
                $isPasswordResetTicket = $true
                Write-Host "Ticket #$($ticket.id) matches subject pattern: $pattern" -ForegroundColor Green
                break
            }
        }

        # Check description patterns if subject didn't match
        if (-not $isPasswordResetTicket) {
            foreach ($pattern in $descriptionPatterns) {
                if ($description -like $pattern) {
                    $isPasswordResetTicket = $true
                    Write-Host "Ticket #$($ticket.id) matches description pattern: $pattern" -ForegroundColor Green
                    break
                }
            }
        }

        # Additional check for simple patterns like "Reset - DC5365"
        if (-not $isPasswordResetTicket) {
            # Check for patterns like "Reset - DC5365" or "Password - DC5365"
            if ($subject -match "^(reset|password|forgot|change)\s*[-:]\s*\w+\d+") {
                $isPasswordResetTicket = $true
                Write-Host "Ticket #$($ticket.id) matches simple pattern: $($ticket.subject)" -ForegroundColor Green
            }
        }

        if ($isPasswordResetTicket) {
            # 🔒 EXTRA SAFETY: Skip application-specific password resets (Zoho, Gmail, O365, etc.)
            $isAppReset = $false
            $matchedKeyword = $null

            foreach ($k in $AppPasswordKeywords) {
                if ($subject -like "*$k*" -or $description -like "*$k*") {
                    $isAppReset = $true
                    $matchedKeyword = $k
                    break
                }
            }

            if ($isAppReset) {
                Write-Host "Ticket #$($ticket.id) is for APPLICATION password ('$matchedKeyword') → SKIPPING AD reset" -ForegroundColor Yellow
                continue   # ⛔ Do NOT add to $passwordResetTickets → no AD action
            }

            Write-Host "Ticket #$($ticket.id) identified as AD Password Reset ticket" -ForegroundColor Cyan
            $passwordResetTickets += $ticket
        }
    }

    if ($passwordResetTickets.Count -gt 0) {
        Write-Host "Found $($passwordResetTickets.Count) matching AD Password Reset tickets" -ForegroundColor Green
        Write-Host "Tickets to process:" -ForegroundColor Cyan
        foreach ($ticket in $passwordResetTickets) {
            Write-Host "ID $($ticket.id): $($ticket.subject)" -ForegroundColor Cyan
        }
        return , $passwordResetTickets
    }
    else {
        Write-Host "No matching AD Password Reset tickets found" -ForegroundColor Yellow
        Write-Host "Note: Application password tickets (Zoho / Gmail / O365 / etc.) are skipped and must be handled separately." -ForegroundColor Yellow
        return @()
    }
}

# FUNCTION: Reset AD User Password - UPDATED WITH APP CHECK + APPROVAL
function Reset-ADUserPassword {
    param($Ticket)

    try {
        Write-Host "PROCESSING PASSWORD RESET TICKET: $($Ticket.id)" -ForegroundColor Cyan
        Write-Host "Subject: $($Ticket.subject)" -ForegroundColor White
        Write-Host "Description: $($Ticket.description)" -ForegroundColor Gray

        # EXTRA SAFETY: Do NOT touch AD password for application (Zoho/Gmail/O365/etc.) reset tickets
        $subLower = $Ticket.subject.ToLower()
        $descLower = if ($Ticket.description) { $Ticket.description.ToLower() } else { "" }
        $isAppResetHere = $false
        $matchedAppKey = $null

        foreach ($k in $AppPasswordKeywords) {
            if ($subLower -like "*$k*" -or $descLower -like "*$k*") {
                $isAppResetHere = $true
                $matchedAppKey = $k
                break
            }
        }

        if ($isAppResetHere) {
            Write-Host "Detected application password reset ('$matchedAppKey') – NO AD password reset will be done." -ForegroundColor Yellow
            return @{
                Success = $false
                Error   = "Application ($matchedAppKey) password reset request – AD password not changed"
            }
        }

        $employeeID = $null

        # EXTENSIVE PATTERNS FOR EMPLOYEE ID EXTRACTION
        $patterns = @(
            # Standard patterns with dash/hyphen
            "Password Reset\s*[-:]\s*(\w+\d+)",
            "Reset Password\s*[-:]\s*(\w+\d+)",
            "Reset Password for\s*(\w+\d+)",
            "Password Reset for\s*(\w+\d+)",
            "Forgot Password\s*[-:]\s*(\w+\d+)",
            "Password Help\s*[-:]\s*(\w+\d+)",
            "Password Change\s*[-:]\s*(\w+\d+)",
            "Change Password\s*[-:]\s*(\w+\d+)",
            "Password Expired\s*[-:]\s*(\w+\d+)",

            # Simple patterns like "Reset - DC5365" or "Password - DC5365"
            "^(Reset|Password|Forgot|Change)\s*[-:]\s*(\w+\d+)",

            # Patterns with parentheses or brackets
            "Password Reset.*?\[(\w+\d+)\]",
            "Password Reset.*?\((\w+\d+)\)",
            "Reset Password.*?\[(\w+\d+)\]",
            "Reset Password.*?\((\w+\d+)\)",

            # Generic patterns anywhere in text
            "(\bDC\d+\b)",  # Specific for DC employee IDs
            "(\bDK\d+\b)",  # Specific for DK employee IDs
            "(\b[A-Z]{2}\d+\b)",  # General pattern: 2 letters followed by numbers

            # Patterns with spaces instead of dashes
            "Password Reset\s+(\w+\d+)",
            "Reset Password\s+(\w+\d+)",
            "Reset\s+(\w+\d+)",
            "Password\s+(\w+\d+)",

            # Partial matches
            ".*?(\w+\d+).*?password.*",
            ".*?password.*?(\w+\d+).*",
            ".*?(\w+\d+).*?reset.*",
            ".*?reset.*?(\w+\d+).*"
        )

        # First try to extract from subject
        $subjectText = $Ticket.subject
        Write-Host "Analyzing subject: $subjectText" -ForegroundColor Gray

        foreach ($pattern in $patterns) {
            if ($subjectText -match $pattern) {
                # For patterns with multiple capture groups, get the last non-empty match
                $matches.Keys | Where-Object { $_ -notmatch '^\d+$' } | ForEach-Object {
                    if ($matches[$_] -match '^\w+\d+$') {
                        $employeeID = $matches[$_]
                    }
                }

                # If we didn't get from named groups, try numbered groups
                if (-not $employeeID) {
                    for ($i = 1; $i -lt $matches.Count; $i++) {
                        if ($matches[$i] -match '^\w+\d+$') {
                            $employeeID = $matches[$i]
                            break
                        }
                    }
                }

                if ($employeeID) {
                    Write-Host "Parsed Employee ID from subject: $employeeID (Pattern: $pattern)" -ForegroundColor Green
                    break
                }
            }
        }

        # If not found in subject, try description
        if (-not $employeeID -and $Ticket.description) {
            $descriptionText = $Ticket.description
            Write-Host "Analyzing description for Employee ID..." -ForegroundColor Gray

            foreach ($pattern in $patterns) {
                if ($descriptionText -match $pattern) {
                    # For patterns with multiple capture groups, get the last non-empty match
                    $matches.Keys | Where-Object { $_ -notmatch '^\d+$' } | ForEach-Object {
                        if ($matches[$_] -match '^\w+\d+$') {
                            $employeeID = $matches[$_]
                        }
                    }

                    # If we didn't get from named groups, try numbered groups
                    if (-not $employeeID) {
                        for ($i = 1; $i -lt $matches.Count; $i++) {
                            if ($matches[$i] -match '^\w+\d+$') {
                                $employeeID = $matches[$i]
                                break
                            }
                        }
                    }

                    if ($employeeID) {
                        Write-Host "Parsed Employee ID from description: $employeeID (Pattern: $pattern)" -ForegroundColor Green
                        break
                    }
                }
            }
        }

        # Last resort: Look for any employee ID pattern in the entire text
        if (-not $employeeID) {
            $combinedText = "$subjectText $($Ticket.description)"
            if ($combinedText -match '\b([A-Z]{2}\d+)\b') {
                $employeeID = $matches[1]
                Write-Host "Found Employee ID in text: $employeeID" -ForegroundColor Yellow
            }
        }

        if (-not $employeeID) {
            Write-Host "Could not extract Employee ID from ticket" -ForegroundColor Yellow
            Write-Host "SUPPORTED TICKET FORMATS:" -ForegroundColor Cyan
            Write-Host "  - Password Reset - DC5365" -ForegroundColor White
            Write-Host "  - Reset Password for DC5365" -ForegroundColor White
            Write-Host "  - Password Reset for DC5365" -ForegroundColor White
            Write-Host "  - Forgot Password - DC5365" -ForegroundColor White
            Write-Host "  - Password Help - DC5365" -ForegroundColor White
            Write-Host "  - Reset - DC5365" -ForegroundColor White
            Write-Host "  - Password - DC5365" -ForegroundColor White
            Write-Host "  - Forgot - DC5365" -ForegroundColor White
            Write-Host "  - Change - DC5365" -ForegroundColor White
            Write-Host "  - Password Reset [DC5365]" -ForegroundColor White
            Write-Host "  - Reset Password (DC5365)" -ForegroundColor White
            return @{Success = $false; Error = "Employee ID pattern not found" }
        }

        # Validate Employee ID format
        if ($employeeID -notmatch '^\w+\d+$') {
            Write-Host "Invalid Employee ID format: $employeeID" -ForegroundColor Red
            return @{Success = $false; Error = "Invalid Employee ID format" }
        }

        # Normalize to uppercase
        $employeeID = $employeeID.ToUpper()

        Write-Host "Looking up employee in Active Directory: $employeeID" -ForegroundColor Yellow

        # Find user by Employee ID (SamAccountName)
        $user = Get-ADUser -Filter "SamAccountName -eq '$employeeID'" -ErrorAction SilentlyContinue

        if (-not $user) {
            Write-Host "Employee '$employeeID' not found in AD" -ForegroundColor Red
            # Try alternative search methods
            Write-Host "Trying alternative search methods..." -ForegroundColor Yellow

            # Search by EmployeeNumber if exists
            $user = Get-ADUser -Filter "EmployeeNumber -eq '$employeeID'" -ErrorAction SilentlyContinue

            if (-not $user) {
                # Search in DisplayName
                $user = Get-ADUser -Filter "DisplayName -like '*$employeeID*'" -ErrorAction SilentlyContinue
            }

            if (-not $user) {
                return @{Success = $false; Error = "Employee not found" }
            }
        }

        $userDetails = Get-ADUser -Identity $user.SamAccountName -Properties DisplayName, Department, EmailAddress
        Write-Host "EMPLOYEE DETAILS:" -ForegroundColor White
        Write-Host "  Name: $($userDetails.DisplayName)" -ForegroundColor Gray
        Write-Host "  Department: $($userDetails.Department)" -ForegroundColor Gray
        Write-Host "  Email: $($userDetails.EmailAddress)" -ForegroundColor Gray

        # CHECK IF IT SUPPORT TEAM MEMBER
        $isITSupportMember = Test-ITSupportTeamMember -EmployeeID $employeeID

        if ($isITSupportMember) {
            Write-Host "⚠️  IT SUPPORT TEAM MEMBER DETECTED - APPROVAL REQUIRED" -ForegroundColor Red
            Write-Host "Manager Approval Required: $ITManagerEmail" -ForegroundColor Yellow

            # Process IT Support member with approval workflow
            $approved = Process-ITSupportMember -Ticket $Ticket -EmployeeID $EmployeeID -EmployeeName $userDetails.DisplayName

            if (-not $approved) {
                Write-Host "Approval not received. Ticket pending IT Manager review." -ForegroundColor Yellow
                return @{
                    Success         = $false
                    Error           = "IT Manager approval required"
                    ApprovalPending = $true
                    TicketID        = $Ticket.id
                }
            }

            Write-Host "✅ IT MANAGER APPROVAL RECEIVED - Proceeding with password reset..." -ForegroundColor Green
        }

        # Generate stronger password that meets domain requirements
        $newPassword = "TempPass" + (Get-Random -Minimum 10000 -Maximum 99999) + "!"
        $securePassword = ConvertTo-SecureString -String $newPassword -AsPlainText -Force

        Write-Host "Resetting password for Employee ID: $employeeID..." -ForegroundColor Yellow

        # Reset password with stronger requirements
        Set-ADAccountPassword -Identity $user.SamAccountName -NewPassword $securePassword -Reset
        Set-ADUser -Identity $user.SamAccountName -ChangePasswordAtLogon $true

        Write-Host "PASSWORD RESET SUCCESSFULLY!" -ForegroundColor Green
        Write-Host "  Employee ID: $employeeID" -ForegroundColor White
        Write-Host "  New Temporary Password: $newPassword" -ForegroundColor White
        Write-Host "  User must change password at next login" -ForegroundColor White

        return @{
            Success           = $true
            EmployeeID        = $employeeID
            NewPassword       = $newPassword
            DisplayName       = $userDetails.DisplayName
            Department        = $userDetails.Department
            TicketID          = $Ticket.id
            IsITSupport       = $isITSupportMember
            ITManagerApproved = $isITSupportMember
        }

    }
    catch {
        Write-Host "Failed to reset password: $($_.Exception.Message)" -ForegroundColor Red
        return @{Success = $false; Error = $_.Exception.Message }
    }
}

# FUNCTION: Update Freshservice Ticket - UPDATED TO CLOSE TICKETS
function Update-FreshserviceTicket {
    param($Result)

    try {
        Write-Host "Updating Freshservice ticket #$($Result.TicketID)..." -ForegroundColor Yellow

        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($FreshserviceAPIKey + ":X")))
        $headers = @{
            "Authorization" = "Basic $base64AuthInfo"
            "Content-Type"  = "application/json"
        }

        $approvalNote = if ($Result.ITManagerApproved) { " (Approved by IT Manager)" } else { "" }

        $note = @"
Thank you for reaching out to DC IT Helpdesk.

PASSWORD RESET COMPLETED SUCCESSFULLY$approvalNote
- Employee ID: $($Result.EmployeeID)
- Employee Name: $($Result.DisplayName)
- Department: $($Result.Department)
- New Temporary Password: $($Result.NewPassword)
- Action: User must change password at next login

Please log in with the temporary password above, then immediately change your password.

**Security Reminder:**
1. Do not share your password with anyone
2. Create a strong, unique password
3. Do not reuse old passwords

Ticket auto-closed by Password Reset Automation.

Best regards,
DC IT Helpdesk
"@

        $noteBody = @{
            body    = $note
            private = $false
        } | ConvertTo-Json

        $noteUri = "$FreshserviceURL/api/v2/tickets/$($Result.TicketID)/notes"
        $noteResponse = Invoke-RestMethod -Uri $noteUri -Method Post -Headers $headers -Body $noteBody

        Write-Host "Completion note added to ticket" -ForegroundColor Green

        # UPDATE: Change status to 5 (Closed) instead of 3 (Pending)
        $body = @{
            status = 5  # Closed status
        } | ConvertTo-Json

        $uri = "$FreshserviceURL/api/v2/tickets/$($Result.TicketID)"
        $response = Invoke-RestMethod -Uri $uri -Method Put -Headers $headers -Body $body

        Write-Host "Ticket status updated to Closed" -ForegroundColor Green
        Write-Host "Automation completed successfully!" -ForegroundColor Green

        Send-DashboardLog -Service "Password Reset" -Status "success" -Message "Reset password for: $($Result.EmployeeID)"
    }
    catch {
        Send-DashboardLog -Service "Password Reset" -Status "error" -Message "Failed to update ticket: $($_.Exception.Message)"
        Write-Host "Could not update ticket: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# MAIN EXECUTION
Write-Host "PASSWORD RESET AUTOMATION WITH IT MANAGER APPROVAL WORKFLOW" -ForegroundColor Cyan
# ... (rest of main execution)

# Add at the very end:
Send-DashboardLog -Service "Password Reset" -Status "info" -Message "Password reset run completed."
& "$PSScriptRoot\Generate-Dashboard.ps1"
Write-Host "IT Support Team Members requiring approval:" -ForegroundColor Yellow
foreach ($itMember in $ITSupportTeam) {
    Write-Host "  - $itMember" -ForegroundColor White
}
Write-Host "IT Manager: $ITManagerEmail" -ForegroundColor Green

Write-Host "STARTING PASSWORD RESET AUTOMATION..." -ForegroundColor Yellow

# Test network connectivity first
Write-Host "Testing network connectivity..." -ForegroundColor Gray
try {
    Test-NetConnection -ComputerName "$FreshserviceDomain.freshservice.com" -Port 443 -InformationLevel Quiet
    Write-Host "Network connectivity: OK" -ForegroundColor Green
}
catch {
    Write-Host "Network connectivity: FAILED" -ForegroundColor Red
}

$allTickets = Get-FreshserviceTickets

if (-not $allTickets) {
    Write-Host "Cannot proceed - API connection failed" -ForegroundColor Red
    Write-Host "Possible solutions:" -ForegroundColor Yellow
    Write-Host "1. Check internet connectivity" -ForegroundColor White
    Write-Host "2. Verify Freshservice domain: $FreshserviceDomain" -ForegroundColor White
    Write-Host "3. Check API key permissions" -ForegroundColor White
    Write-Host "4. Verify TLS 1.2 is enabled on this system" -ForegroundColor White
    exit 1
}

Write-Host "Checking for Password Reset tickets..." -ForegroundColor Cyan
$passwordResetTickets = Get-PasswordResetTickets -AllTickets $allTickets

$ticketsArray = @($passwordResetTickets)

Write-Host "After array conversion:" -ForegroundColor Magenta
Write-Host "ticketsArray count: $($ticketsArray.Count)" -ForegroundColor Magenta

if ($ticketsArray.Count -gt 0) {
    Write-Host "Found $($ticketsArray.Count) Password Reset tickets to process!" -ForegroundColor Green

    Write-Host "Processing tickets..." -ForegroundColor Green

    foreach ($ticket in $ticketsArray) {
        Write-Host "PROCESSING TICKET #$($ticket.id)" -ForegroundColor Green
        Write-Host "Subject: $($ticket.subject)" -ForegroundColor White

        $result = Reset-ADUserPassword -Ticket $ticket

        if ($result.Success) {
            Update-FreshserviceTicket -Result $result
            Write-Host "PASSWORD RESET COMPLETED SUCCESSFULLY!" -ForegroundColor Green
            Write-Host "Reset password for Employee ID: $($result.EmployeeID)" -ForegroundColor White

            if ($result.IsITSupport) {
                Write-Host "IT Manager Approval: RECEIVED" -ForegroundColor Green
            }

            Write-Host "Ticket #$($result.TicketID) has been CLOSED" -ForegroundColor Green
            break
        }
        elseif ($result.ApprovalPending) {
            Write-Host "IT Manager approval pending for ticket #$($ticket.id)" -ForegroundColor Yellow
            Write-Host "Ticket remains open for IT Manager review" -ForegroundColor Cyan
            # Continue to next ticket instead of breaking
        }
        else {
            Write-Host "Failed to process ticket #$($ticket.id): $($result.Error)" -ForegroundColor Red
        }
    }
}
else {
    Write-Host "No AD Password Reset tickets found to process" -ForegroundColor Red
    Write-Host "ACCEPTED TICKET FORMATS (for AD reset):" -ForegroundColor Cyan
    Write-Host "  - Password Reset - DC5365" -ForegroundColor White
    Write-Host "  - Reset Password for DC5365" -ForegroundColor White
    Write-Host "  - Password Reset for DC5365" -ForegroundColor White
    Write-Host "  - Forgot Password - DC5365" -ForegroundColor White
    Write-Host "  - Password Help - DC5365" -ForegroundColor White
    Write-Host "  - Reset - DC5365" -ForegroundColor White
    Write-Host "  - Password - DC5365" -ForegroundColor White
    Write-Host "  - Forgot - DC5365" -ForegroundColor White
    Write-Host "  - Change - DC5365" -ForegroundColor White
    Write-Host "  - Password Reset [DC5365]" -ForegroundColor White
    Write-Host "  - Reset Password (DC5365)" -ForegroundColor White
    Write-Host "NOTE: Tickets mentioning Zoho / Gmail / O365 / Outlook / etc. are treated as application passwords and will NOT reset AD." -ForegroundColor Yellow
}

Write-Host "PASSWORD RESET AUTOMATION COMPLETED: $(Get-Date)" -ForegroundColor Cyan
