# Microsoft Graph and Exchange Online APIs Used in Get-CopilotUsageAnalytics.ps1  
1. Exchange Online PowerShell Module  
**Connect-ExchangeOnline** - Authenticates to Exchange Online  
**Disconnect-ExchangeOnline** - Closes the Exchange Online session  
2. Unified Audit Log API (via Search-UnifiedAuditLog cmdlet)  

### Direct Copilot Usage Tracking:  
>Search-UnifiedAuditLog -Operations "CopilotInteraction"

Endpoint: Accesses the Microsoft 365 Unified Audit Log  
Purpose: Retrieves direct Copilot interaction events  

### Bot/Agent Development Activities:  
>Search-UnifiedAuditLog -Operations $op

Where $op includes:

"BotCreate"  
"BotUpdateOperation-BotPublish"  
"BotUpdateOperation-BotNameUpdate"  
"BotUpdateOperation-BotIconUpdate"  
"BotUpdateOperation-BotAuthUpdate"    
"BotComponentCreate"  
"BotComponentUpdate"  
"BotComponentDelete"   

### SharePoint Agent Activities:  
>Search-UnifiedAuditLog -Operations $fileOp

Where $fileOp includes:

"FileAccessed"  
"FileModified"  
"FileCreated"  
"FileCopied"  
"FileDownloaded"  
"PageViewed"  
"SearchQueryPerformed"  

### Keyword-Based Search:
>Search-UnifiedAuditLog -FreeText $term  

Where $term includes:

"copilot"  
"agent"  
"AI"  
"assistant"  
"bot"  
"intelligent"  
"smart"  
"generate"  
"compose"  

### Underlying API Endpoints  
While the script uses PowerShell cmdlets, these translate to the following Microsoft Graph API endpoints:

1. Audit Log API  
>https://graph.microsoft.com/v1.0/auditLogs/directoryAudits  
https://graph.microsoft.com/beta/auditLogs/signIns

2. Security & Compliance API  
>https://graph.microsoft.com/v1.0/security/alerts  
https://manage.office.com/api/v1.0/{tenant_id}/activity/feed

3. SharePoint API (indirectly accessed via audit logs)
>https://graph.microsoft.com/v1.0/sites  
https://graph.microsoft.com/v1.0/sites/{site-id}/drives  

### API Call Pattern Summary  
Authentication: Connect-ExchangeOnline  
Main Data Retrieval: Multiple Search-UnifiedAuditLog calls with different parameters:  
Operations-based filtering (CopilotInteraction, Bot operations, File operations)  
Keyword-based search (FreeText parameter)  
Date range filtering (StartDate, EndDate)  
Result size limiting (ResultSize parameter)  
Data Processing: Local PowerShell processing of returned audit data
Cleanup: Disconnect-ExchangeOnline  

### Note on API Usage
The script primarily uses the Microsoft 365 Unified Audit Log API accessed through Exchange Online PowerShell cmdlets, rather than direct Microsoft Graph REST API calls. This approach leverages the established audit logging infrastructure to track Copilot and agent usage across the Microsoft 365 environment.