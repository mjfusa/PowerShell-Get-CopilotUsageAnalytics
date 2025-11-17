# Comprehensive Copilot Usage Analytics
# This script provides real usage tracking for Microsoft Copilot and Agents

param(
    [int]$DaysBack = 30,
    [switch]$DetailedAnalysis,
    [switch]$ExportAll,
    [int]$WindowDays = 7  # Size of date window for chunked queries
)

Import-Module ExchangeOnlineManagement -Force

#region Helper Functions

function Invoke-WithRetry {
    <#
    .SYNOPSIS
    Executes a script block with retry logic to handle transient failures and throttling.
    
    .DESCRIPTION
    Retries failed operations with exponential backoff and random jitter to avoid overwhelming the service.
    #>
    param(
        [Parameter(Mandatory)] 
        [scriptblock]$Script,
        [int]$MaxAttempts = 4,
        [int]$BaseDelaySeconds = 3
    )
    
    for ($i = 1; $i -le $MaxAttempts; $i++) {
        try {
            return & $Script
        } catch {
            $err = $_
            # Surface the message and any server IDs if present
            Write-Warning ("Attempt {0}/{1} failed: {2}" -f $i, $MaxAttempts, $err.Exception.Message)
            if ($err.Exception.Data) { 
                $err.Exception.Data | Out-String | Write-Verbose 
            }
            
            if ($i -eq $MaxAttempts) { 
                throw 
            }
            
            # Exponential backoff with random jitter
            $delay = $BaseDelaySeconds * $i + (Get-Random -Minimum 0 -Maximum 3)
            Write-Host "   Waiting $delay seconds before retry..." -ForegroundColor Gray
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-AuditLogsWindowed {
    <#
    .SYNOPSIS
    Retrieves audit logs in time-windowed chunks to avoid throttling and timeouts.
    
    .DESCRIPTION
    Processes a large date range by breaking it into smaller windows and using retry logic.
    #>
    param(
        [Parameter(Mandatory)]
        [DateTime]$StartDate,
        [Parameter(Mandatory)]
        [DateTime]$EndDate,
        [string]$Operations,
        [string]$FreeText,
        [int]$ResultSize = 500,
        [int]$WindowDays = 7
    )
    
    $allResults = @()
    $cursor = $StartDate
    
    while ($cursor -lt $EndDate) {
        $windowEnd = $cursor.AddDays($WindowDays)
        if ($windowEnd -gt $EndDate) { 
            $windowEnd = $EndDate 
        }
        
        Write-Verbose ("Processing window: {0:yyyy-MM-dd} → {1:yyyy-MM-dd}" -f $cursor, $windowEnd)
        
        try {
            $result = Invoke-WithRetry {
                $params = @{
                    StartDate = $cursor
                    EndDate = $windowEnd
                    ResultSize = $ResultSize
                    ErrorAction = 'Stop'
                }
                
                if ($Operations) {
                    $params['Operations'] = $Operations
                }
                
                if ($FreeText) {
                    $params['FreeText'] = $FreeText
                }
                
                Search-UnifiedAuditLog @params
            }
            
            if ($result) {
                $allResults += $result
                Write-Verbose "   Retrieved $($result.Count) records from this window"
            }
        } catch {
            Write-Warning ("Failed to retrieve data for window {0:yyyy-MM-dd} → {1:yyyy-MM-dd}: {2}" -f $cursor, $windowEnd, $_.Exception.Message)
        }
        
        $cursor = $windowEnd
        
        # Small delay between windows to be courteous to the service
        if ($cursor -lt $EndDate) {
            Start-Sleep -Milliseconds 500
        }
    }
    
    return $allResults
}

#endregion

Write-Host "=== Microsoft Copilot & Agent Usage Analytics ===" -ForegroundColor Green
Write-Host "Period: Last $DaysBack days (using $WindowDays-day windows)" -ForegroundColor Gray

try {
    Connect-ExchangeOnline -ShowProgress $false

    # Check if user has required permissions
    Write-Host "`nChecking required permissions..." -ForegroundColor Cyan
    
    # Required roles for accessing audit logs (minimum permissions needed)
    $requiredRoles = @(
        "View-Only Audit Logs (minimum required)",
        "Audit Logs", 
        "Compliance Administrator",
        "Security Administrator",
        "Security Reader",
        "Global Reader (includes audit log access)"
    )
    
    $additionalInfo = @"

MINIMUM REQUIRED PERMISSIONS:
• View-Only Audit Logs role assignment
• OR any role that includes audit log read permissions

NOTES:
• Global Administrator is NOT required for this script
• The script only reads audit logs, no administrative actions
• Exchange Online PowerShell connection is required
• Audit logging must be enabled in your tenant
"@
    
    try {
        # Get current user's role assignments
        $currentUser = Get-ConnectionInformation | Select-Object -First 1
        if ($currentUser) {
            Write-Host "   Connected as: $($currentUser.UserPrincipalName)" -ForegroundColor Gray
            
            # Test audit log access by attempting a simple query with retry
            $testAudit = Invoke-WithRetry {
                Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) -ResultSize 1 -ErrorAction Stop
            }
            Write-Host "   ✓ Audit log access confirmed" -ForegroundColor Green
        } else {
            throw "Unable to determine current user connection"
        }
    } catch {
        Write-Host "   ✗ ERROR: Unable to access audit logs" -ForegroundColor Red
        Write-Host "`n   Required permissions (any ONE of these):" -ForegroundColor Yellow
        foreach ($role in $requiredRoles) {
            Write-Host "   • $role" -ForegroundColor White
        }
        Write-Host $additionalInfo -ForegroundColor Gray
        Write-Host "`n   Troubleshooting steps:" -ForegroundColor Yellow
        Write-Host "   1. Contact your administrator to assign the required role" -ForegroundColor White
        Write-Host "   2. Verify you're connected with the correct account" -ForegroundColor White
        Write-Host "   3. Check that audit logging is enabled in your organization" -ForegroundColor White
        Write-Host "`n   Current error: $($_.Exception.Message)" -ForegroundColor Gray
        
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        exit 1
    }

    $startDate = (Get-Date).AddDays(-$DaysBack)
    $endDate = Get-Date
    
    Write-Host "\n1. Microsoft Copilot Direct Usage..." -ForegroundColor Yellow
    
    # Get Copilot interactions using windowed approach
    $copilotEvents = Get-AuditLogsWindowed -StartDate $startDate -EndDate $endDate -Operations "CopilotInteraction" -ResultSize 5000 -WindowDays $WindowDays
    
    # Initialize array to store Copilot agent information
    $copilotAgentNames = @()
    
    if ($copilotEvents -and $copilotEvents.Count -gt 0) {
        Write-Host "   ✓ Found $($copilotEvents.Count) Copilot interactions" -ForegroundColor Green
        
        # Extract agent names from Copilot interactions
        foreach ($event in $copilotEvents) {
            try {
                if ($event.AuditData) {
                    $auditObj = $event.AuditData | ConvertFrom-Json -ErrorAction SilentlyContinue
                    if ($auditObj) {
                        $agentName = $null
                        # Look for agent/app identifiers in Copilot interactions
                        if ($auditObj.AppName) { $agentName = $auditObj.AppName }
                        elseif ($auditObj.AgentName) { $agentName = $auditObj.AgentName }
                        elseif ($auditObj.ApplicationName) { $agentName = $auditObj.ApplicationName }
                        elseif ($auditObj.Workload) { $agentName = "Copilot for $($auditObj.Workload)" }
                        elseif ($auditObj.AppType) { $agentName = $auditObj.AppType }
                        
                        if ($agentName) {
                            $copilotAgentNames += [PSCustomObject]@{
                                AgentName = $agentName
                                User = $event.UserIds
                                Operation = $event.Operations
                                Timestamp = $event.CreationDate
                            }
                        }
                    }
                }
            } catch {
                # Continue processing even if JSON parsing fails
            }
        }
        
        # Analyze usage patterns
        $userUsage = $copilotEvents | Group-Object UserIds | Sort-Object Count -Descending
        $dailyUsage = $copilotEvents | Group-Object {$_.CreationDate.Date} | Sort-Object Name
        
        Write-Host "   Top Copilot Users:" -ForegroundColor Cyan
        $userUsage | Select-Object -First 10 | ForEach-Object {
            Write-Host "     $($_.Name): $($_.Count) interactions" -ForegroundColor White
        }
        
        Write-Host "   Daily Usage Trend:" -ForegroundColor Cyan
        $dailyUsage | ForEach-Object {
            $date = [DateTime]::Parse($_.Name).ToString("MM/dd")
            Write-Host "     $date`: $($_.Count) interactions" -ForegroundColor White
        }
        
        if ($ExportAll) {
            $copilotPath = ".\CopilotInteractions-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
            $copilotEvents | Export-Csv -Path $copilotPath -NoTypeInformation
            Write-Host "   ✓ Copilot interactions exported to: $copilotPath" -ForegroundColor Green
        }
    } else {
        Write-Host "   • No Copilot interactions found" -ForegroundColor Gray
    }
    
    Write-Host "2. Bot/Agent Development Activity..." -ForegroundColor Yellow
    
    $botOperations = @(
        "BotCreate", "BotUpdateOperation-BotPublish", "BotUpdateOperation-BotNameUpdate",
        "BotUpdateOperation-BotIconUpdate", "BotUpdateOperation-BotAuthUpdate", 
        "BotComponentCreate", "BotComponentUpdate", "BotComponentDelete"
    )
    
    $allBotEvents = @()
    $botAgentNames = @()
    
    foreach ($op in $botOperations) {
        Write-Host "   Searching for operation: $op" -ForegroundColor Gray
        
        $events = Get-AuditLogsWindowed -StartDate $startDate -EndDate $endDate -Operations $op -ResultSize 1000 -WindowDays $WindowDays
        
        if ($events -and $events.Count -gt 0) {
            $allBotEvents += $events
            Write-Host "   ✓ $($op): $($events.Count) events" -ForegroundColor Green
            
            # Extract bot/agent names from bot operations
            foreach ($event in $events) {
                try {
                    if ($event.AuditData) {
                        $auditObj = $event.AuditData | ConvertFrom-Json -ErrorAction SilentlyContinue
                        if ($auditObj) {
                            $agentName = $null
                            if ($auditObj.BotName) { $agentName = $auditObj.BotName }
                            elseif ($auditObj.AgentName) { $agentName = $auditObj.AgentName }
                            elseif ($auditObj.Name) { $agentName = $auditObj.Name }
                            elseif ($auditObj.DisplayName) { $agentName = $auditObj.DisplayName }
                            elseif ($event.ObjectId) { $agentName = $event.ObjectId }
                            
                            if ($agentName) {
                                $botAgentNames += [PSCustomObject]@{
                                    AgentName = $agentName
                                    User = $event.UserIds
                                    Operation = $event.Operations
                                    Timestamp = $event.CreationDate
                                }
                            }
                        }
                    }
                } catch {
                    # Continue processing even if JSON parsing fails
                }
            }
        }
    }
    
    if ($allBotEvents -and $allBotEvents.Count -gt 0) {
        Write-Host "   Total Bot/Agent Events: $($allBotEvents.Count)" -ForegroundColor White
        
        $botUsers = $allBotEvents | Group-Object UserIds | Sort-Object Count -Descending
        Write-Host "   Bot Developers:" -ForegroundColor Cyan
        $botUsers | ForEach-Object {
            Write-Host "     $($_.Name): $($_.Count) bot operations" -ForegroundColor White
        }
    } else {
        Write-Host "   • No bot/agent development activity found" -ForegroundColor Gray
    }
    
    Write-Host "`n3. SharePoint Agent Activity (Comprehensive Search)..." -ForegroundColor Yellow
    
    # Look for SharePoint activities that might be agent-enhanced
    $spAIActivities = @()
    
    # Enhanced search for file activities with agent indicators (same as Find-CopilotAuditEvents.ps1)
    $fileOps = @("FileAccessed", "FileModified", "FileCreated", "FileCopied", "FileDownloaded", "PageViewed", "SearchQueryPerformed")
    foreach ($fileOp in $fileOps) {
        Write-Host "   Searching SharePoint operation: $fileOp" -ForegroundColor Gray
        
        try {
            $events = Get-AuditLogsWindowed -StartDate $startDate -EndDate $endDate -Operations $fileOp -ResultSize 500 -WindowDays $WindowDays
            
            if ($events -and $events.Count -gt 0) {
                # Filter for potential agent-related activities (expanded criteria)
                $agentEvents = $events | Where-Object { 
                    $_.AuditData -like "*agent*" -or 
                    $_.AuditData -like "*copilot*" -or
                    $_.AuditData -like "*bot*" -or
                    $_.ObjectId -like "*.agent*" -or
                    $_.UserIds -like "*agent*" -or
                    $_.UserIds -like "*copilot*" -or
                    $_.AuditData -like "*AI*" -or
                    $_.AuditData -like "*assistant*" -or
                    $_.AuditData -like "*intelligent*" -or
                    ($_.AuditData -like "*CorrelationId*" -and $_.AuditData -like "*smart*")
                }
                
                if ($agentEvents -and $agentEvents.Count -gt 0) {
                    $spAIActivities += $agentEvents
                    Write-Host "     ✓ Found $($agentEvents.Count) potential agent-related $fileOp events" -ForegroundColor Green
                } else {
                    Write-Host "     • No agent-related $fileOp events found" -ForegroundColor Gray
                }
            }
        } catch {
            Write-Host "     ⚠ Error searching SharePoint operation $fileOp`: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    # Additional keyword-based search for AI/ML/Copilot terms in SharePoint activities
    Write-Host "   Searching for AI/ML/Copilot keywords in SharePoint activities..." -ForegroundColor Gray
    
    $searchTerms = @("copilot", "agent", "AI", "assistant", "bot", "intelligent", "smart", "generate", "compose")
    
    foreach ($term in $searchTerms) {
        try {
            Write-Host "     Searching for keyword: $term" -ForegroundColor DarkGray
            
            # Search in SharePoint-specific operations for our keywords
            $keywordResults = Get-AuditLogsWindowed -StartDate $startDate -EndDate $endDate -FreeText $term -ResultSize 100 -WindowDays $WindowDays
            
            if ($keywordResults -and $keywordResults.Count -gt 0) {
                # Filter for SharePoint operations only
                $spKeywordResults = $keywordResults | Where-Object {
                    $_.Operations -like "*File*" -or 
                    $_.Operations -like "*Page*" -or 
                    $_.Operations -like "*Search*" -or
                    $_.Operations -like "*SharePoint*" -or
                    $_.AuditData -like "*sharepoint*"
                }
                
                if ($spKeywordResults -and $spKeywordResults.Count -gt 0) {
                    $spAIActivities += $spKeywordResults
                    Write-Host "     ✓ Found $($spKeywordResults.Count) SharePoint events containing '$term'" -ForegroundColor Green
                }
            }
        } catch {
            Write-Host "     ⚠ Error searching for keyword $term`: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    # Remove duplicates from SharePoint activities and analyze
    $uniqueSPActivities = $spAIActivities | Sort-Object Identity -Unique
    
    if ($uniqueSPActivities -and $uniqueSPActivities.Count -gt 0) {
        Write-Host "   ✓ Total unique SharePoint agent activities: $($uniqueSPActivities.Count)" -ForegroundColor Green
        
        # Analyze SharePoint agent usage patterns
        $spOperationGroups = $uniqueSPActivities | Group-Object Operations | Sort-Object Count -Descending
        Write-Host "   SharePoint Agent Activity by Operation:" -ForegroundColor Cyan
        $spOperationGroups | Select-Object -First 10 | ForEach-Object {
            Write-Host "     $($_.Name): $($_.Count) events" -ForegroundColor White
        }
        
        $spUserGroups = $uniqueSPActivities | Group-Object UserIds | Sort-Object Count -Descending
        Write-Host "   Top SharePoint Agent Users:" -ForegroundColor Cyan
        $spUserGroups | Select-Object -First 5 | ForEach-Object {
            Write-Host "     $($_.Name): $($_.Count) agent activities" -ForegroundColor White
        }
        
        # Extract and analyze agent names from SharePoint activities
        $spAgentNames = @()
        foreach ($activity in $uniqueSPActivities) {
            try {
                if ($activity.AuditData) {
                    $auditObj = $activity.AuditData | ConvertFrom-Json -ErrorAction SilentlyContinue
                    if ($auditObj) {
                        # Look for agent name in various properties
                        $agentName = $null
                        if ($auditObj.AgentName) { $agentName = $auditObj.AgentName }
                        elseif ($auditObj.ApplicationName -and $auditObj.ApplicationName -like "*agent*") { $agentName = $auditObj.ApplicationName }
                        elseif ($auditObj.UserAgent -and $auditObj.UserAgent -like "*agent*") { $agentName = $auditObj.UserAgent }
                        elseif ($auditObj.SourceFileName -and $auditObj.SourceFileName -like "*.agent*") { $agentName = $auditObj.SourceFileName }
                        elseif ($activity.ObjectId -and $activity.ObjectId -like "*.agent*") { $agentName = $activity.ObjectId }
                        
                        if ($agentName) {
                            $spAgentNames += [PSCustomObject]@{
                                AgentName = $agentName
                                User = $activity.UserIds
                                Operation = $activity.Operations
                                Timestamp = $activity.CreationDate
                            }
                        }
                    }
                }
            } catch {
                # Continue processing even if JSON parsing fails
            }
        }
        
        # Group and display agent names
        if ($spAgentNames -and $spAgentNames.Count -gt 0) {
            $agentGroups = $spAgentNames | Group-Object AgentName | Sort-Object Count -Descending
            Write-Host "   Top SharePoint Agents and Users:" -ForegroundColor Cyan
            $agentGroups | Select-Object -First 10 | ForEach-Object {
                $agentName = $_.Name
                $users = ($_.Group.User | Sort-Object -Unique) -join ', '
                Write-Host "     $agentName`: $($_.Count) activities (Users: $users)" -ForegroundColor White
            }
            
            # Also show by user-agent combination
            $userAgentGroups = $spAgentNames | Group-Object User, AgentName | Sort-Object Count -Descending
            Write-Host "   Top User-SharePoint Agent Combinations:" -ForegroundColor Cyan
            $userAgentGroups | Select-Object -First 10 | ForEach-Object {
                $userAgent = $_.Name -replace ', ', ' → '
                Write-Host "     $userAgent`: $($_.Count) activities" -ForegroundColor White
            }
        } else {
            Write-Host "   • No specific agent names detected in audit data" -ForegroundColor Gray
        }
        
        # Store unique results for final summary
        $spAIActivities = $uniqueSPActivities
    } else {
        Write-Host "   • No SharePoint agent activities detected" -ForegroundColor Gray
        $spAIActivities = @()
    }
    
    # Combine all non-SharePoint agent data
    $allCopilotAgents = @()
    if ($copilotAgentNames) { $allCopilotAgents += $copilotAgentNames }
    if ($botAgentNames) { $allCopilotAgents += $botAgentNames }
    
    # Top Copilot Agent Users (excluding SharePoint)
    if ($allCopilotAgents -and $allCopilotAgents.Count -gt 0) {
        Write-Host "\nTop Copilot Agent Users:" -ForegroundColor Cyan
        $topCopilotAgentUsers = $allCopilotAgents | Group-Object User | Sort-Object Count -Descending | Select-Object -First 5
        $topCopilotAgentUsers | ForEach-Object {
            Write-Host "  • $($_.Name): $($_.Count) agent interactions" -ForegroundColor White
        }
    }
    
    # Top Copilot Agents and Users (excluding SharePoint)
    if ($allCopilotAgents -and $allCopilotAgents.Count -gt 0) {
        Write-Host "\nTop Copilot Agents and Users:" -ForegroundColor Cyan
        $topCopilotAgentCombos = $allCopilotAgents | Group-Object AgentName, User | Sort-Object Count -Descending | Select-Object -First 5
        $topCopilotAgentCombos | ForEach-Object {
            $parts = $_.Name -split ', '
            $agentName = if ($parts.Count -gt 0) { $parts[0] } else { "Unknown" }
            $userName = if ($parts.Count -gt 1) { $parts[1] } else { "Unknown" }
            Write-Host "  • $agentName (User: $userName): $($_.Count) interactions" -ForegroundColor White
        }
    }
    
    Write-Host "\n4. Advanced Analytics..." -ForegroundColor Yellow
    
    if ($DetailedAnalysis -and $copilotEvents -and $copilotEvents.Count -gt 0) {
        Write-Host "   Analyzing Copilot interaction patterns..." -ForegroundColor Cyan
        
        # Time-based analysis
        $hourlyUsage = $copilotEvents | Group-Object {$_.CreationDate.Hour} | Sort-Object Name
        Write-Host "   Peak Usage Hours:" -ForegroundColor White
        $hourlyUsage | Sort-Object Count -Descending | Select-Object -First 5 | ForEach-Object {
            $hour = [int]$_.Name
            $timeStr = if ($hour -eq 0) { "12 AM" } elseif ($hour -lt 12) { "$hour AM" } elseif ($hour -eq 12) { "12 PM" } else { "$($hour-12) PM" }
            Write-Host "     $timeStr`: $($_.Count) interactions" -ForegroundColor Gray
        }
        
        # Weekly pattern
        $weeklyUsage = $copilotEvents | Group-Object {$_.CreationDate.DayOfWeek} | Sort-Object Count -Descending
        Write-Host "   Usage by Day of Week:" -ForegroundColor White
        $weeklyUsage | ForEach-Object {
            Write-Host "     $($_.Name): $($_.Count) interactions" -ForegroundColor Gray
        }
        
        # Try to parse audit data for more insights
        if ($copilotEvents.Count -gt 0) {
            Write-Host "   Sample Copilot Interaction Details:" -ForegroundColor White
            $copilotEvents | Select-Object -First 3 | ForEach-Object {
                Write-Host "     $($_.CreationDate) - User: $($_.UserIds)" -ForegroundColor Gray
                if ($_.AuditData) {
                    try {
                        $auditObj = $_.AuditData | ConvertFrom-Json
                        if ($auditObj.Workload) { Write-Host "       Workload: $($auditObj.Workload)" -ForegroundColor Gray }
                        if ($auditObj.Operation) { Write-Host "       Operation: $($auditObj.Operation)" -ForegroundColor Gray }
                        if ($auditObj.ClientIP) { Write-Host "       Client IP: $($auditObj.ClientIP)" -ForegroundColor Gray }
                    } catch {
                        Write-Host "       Raw data available but not JSON parseable" -ForegroundColor Gray
                    }
                }
                Write-Host ""
            }
        }
    }
    
    Write-Host "`n=== USAGE SUMMARY ===" -ForegroundColor Green
    
    $totalInteractions = if ($copilotEvents) { $copilotEvents.Count } else { 0 }
    $uniqueUsers = if ($copilotEvents) { ($copilotEvents | Select-Object UserIds -Unique).Count } else { 0 }
    $totalBotOps = if ($allBotEvents) { $allBotEvents.Count } else { 0 }
    $spAICount = if ($spAIActivities) { $spAIActivities.Count } else { 0 }
    
    Write-Host "📊 Copilot Interactions: $totalInteractions (by $uniqueUsers unique users)" -ForegroundColor White
    Write-Host "🤖 Bot/Agent Operations: $totalBotOps" -ForegroundColor White
    Write-Host "📁 SharePoint AI Activities: $spAICount" -ForegroundColor White
    Write-Host "📅 Analysis Period: $DaysBack days" -ForegroundColor White
    
    if ($totalInteractions -gt 0) {
        $dailyAvg = [math]::Round($totalInteractions / $DaysBack, 2)
        Write-Host "📈 Daily Average: $dailyAvg interactions/day" -ForegroundColor White
        
        $adoptionRate = [math]::Round(($uniqueUsers / 100) * 100, 2)  # Assuming ~100 potential users
        Write-Host "👥 Adoption Indicator: $uniqueUsers active users" -ForegroundColor White
    }
    
    # Export comprehensive report
    if ($ExportAll) {
        $summaryReport = @()
        
        if ($copilotEvents) {
            foreach ($event in $copilotEvents) {
                $summaryReport += [PSCustomObject]@{
                    Timestamp = $event.CreationDate
                    Type = "Copilot Interaction"
                    User = $event.UserIds
                    Operation = $event.Operations
                    ObjectId = $event.ObjectId
                    ClientIP = if ($event.ClientIP) { $event.ClientIP } else { "N/A" }
                    Details = if ($event.AuditData) { $event.AuditData } else { "N/A" }
                }
            }
        }
        
        if ($allBotEvents) {
            foreach ($event in $allBotEvents) {
                $summaryReport += [PSCustomObject]@{
                    Timestamp = $event.CreationDate
                    Type = "Bot/Agent Development"
                    User = $event.UserIds
                    Operation = $event.Operations
                    ObjectId = $event.ObjectId
                    ClientIP = if ($event.ClientIP) { $event.ClientIP } else { "N/A" }
                    Details = if ($event.AuditData) { $event.AuditData } else { "N/A" }
                }
            }
        }
        
        if ($spAIActivities) {
            foreach ($event in $spAIActivities) {
                # Extract agent name from audit data
                $agentName = "N/A"
                try {
                    if ($event.AuditData) {
                        $auditObj = $event.AuditData | ConvertFrom-Json -ErrorAction SilentlyContinue
                        if ($auditObj) {
                            if ($auditObj.AgentName) { $agentName = $auditObj.AgentName }
                            elseif ($auditObj.ApplicationName -and $auditObj.ApplicationName -like "*agent*") { $agentName = $auditObj.ApplicationName }
                            elseif ($auditObj.UserAgent -and $auditObj.UserAgent -like "*agent*") { $agentName = $auditObj.UserAgent }
                            elseif ($auditObj.SourceFileName -and $auditObj.SourceFileName -like "*.agent*") { $agentName = $auditObj.SourceFileName }
                            elseif ($event.ObjectId -and $event.ObjectId -like "*.agent*") { $agentName = $event.ObjectId }
                        }
                    }
                } catch {
                    # Keep default "N/A" if parsing fails
                }
                
                $summaryReport += [PSCustomObject]@{
                    Timestamp = $event.CreationDate
                    Type = "SharePoint Agent Activity"
                    User = $event.UserIds
                    Operation = $event.Operations
                    AgentName = $agentName
                    UPN = $event.UserIds
                    ObjectId = $event.ObjectId
                    ClientIP = if ($event.ClientIP) { $event.ClientIP } else { "N/A" }
                    Details = if ($event.AuditData) { $event.AuditData } else { "N/A" }
                }
            }
        }
        
        if ($summaryReport) {
            $reportPath = ".\CopilotUsageReport-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
            $summaryReport | Sort-Object Timestamp -Descending | Export-Csv -Path $reportPath -NoTypeInformation
            Write-Host "`n✅ Comprehensive report exported to: $reportPath" -ForegroundColor Green
        }
    }
    
    Write-Host "`n💡 Recommendations:" -ForegroundColor Cyan
    if ($totalInteractions -eq 0) {
        Write-Host "   • No Copilot usage detected - consider user training/adoption campaigns" -ForegroundColor Yellow
    } elseif ($totalInteractions -lt 50) {
        Write-Host "   • Low Copilot usage - consider increasing awareness and training" -ForegroundColor Yellow
    } else {
        Write-Host "   • Good Copilot adoption - monitor trends and gather user feedback" -ForegroundColor Green
    }
    
    if ($totalBotOps -gt 0) {
        Write-Host "   • Bot/Agent development activity detected - track deployment success" -ForegroundColor Green
    }

} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
} finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}

Write-Host "`nAnalysis completed at $(Get-Date)" -ForegroundColor Gray