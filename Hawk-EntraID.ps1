<#
AUTHOR
Victor Staflin @TrueSec

.DESCRIPTION
This script uses the Microsoft Graph PowerShell SDK to retrieve information about custom applications
in Azure AD and their associated API permissions. It generates an HTML report and provides an option
to export the data as a CSV file.

.PREREQUISITES
1. Install the Microsoft Graph PowerShell SDK:
   Install-Module Microsoft.Graph -Scope CurrentUser

2. If you haven't already, you may need to install the Az module:
   Install-Module -Name Az -Scope CurrentUser -Repository PSGallery -Force

3. Ensure you have the necessary permissions in Azure AD to read application information.

.USAGE
1. Open PowerShell and navigate to the directory containing this script.
2. Run the script:
   .\ScriptName.ps1

3. When prompted, sign in with your Azure AD credentials.
4. The script will generate an HTML report and open it in your default browser.
5. You can export the data to CSV by clicking the "Export to CSV" button in the HTML report.

.NOTES
- This script requires an active internet connection.
- The account used to run this script must have sufficient permissions in Azure AD.
- The generated report will be saved in the same directory as the script.
#>

# Install and import the Graph module
# Install-Module Microsoft.Graph -Scope CurrentUser

# --------------------------------------------
# 1. Import Necessary Modules
# --------------------------------------------

Write-Host "📦 Loading Modules..." -ForegroundColor Cyan
#Import-Module Microsoft.Graph -ErrorAction Stop
#Import-Module ImportExcel -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Web

Write-Host "✅ Modules Loaded Successfully" -ForegroundColor Green

# --------------------------------------------
# 2. Connect to Microsoft Graph
# --------------------------------------------

Write-Host "🔐 Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "Application.Read.All", "Directory.Read.All"

Write-Host "✅ Connected to Microsoft Graph!" -ForegroundColor Green

# --------------------------------------------
# 2a. Script Parameters
# --------------------------------------------

# Ask the user if they want to include built-in Microsoft applications
# Prompt user with a default of "No" for including built-in applications
$includeBuiltIn = Read-Host "Include Microsoft built-in applications? (y/N) [default: N]"
$includeBuiltIn = if ($includeBuiltIn -eq 'y') { $true } else { $false }

# --------------------------------------------
# 3. Define Helper Functions
# --------------------------------------------

# Function to check if an app is built-in Microsoft app
function Is-BuiltInMicrosoftApp($app) {
    # List of known built-in Microsoft apps
    $builtInApps = @(
        'Microsoft Flow Service', 'Microsoft Teams AadSync', 'Azure Information Protection', 
        'SalesInsightsWebApp', 'Teams NRT DLP Ingestion Service', 'WindowsUpdate-Service',
        'Office 365 SharePoint Online', 'Office 365 Information Protection', 'AAD App Management',
        'Skype for Business Online', 'Skype Teams Firehose', 'Microsoft Defender for Cloud Apps MIP Server',
        'Microsoft password reset service', 'Power BI Premium', 'Teams CMD Services and Data',
        'Microsoft Exchange Online Protection', 'IAMTenantCrawler', 'ComplianceWorkbenchApp',
        'Skype and Teams Tenant Admin API', 'Microsoft Teams User Profile Search Service', 
        'Universal Print', 'Microsoft Azure AD Identity Protection', 'SharePoint Online Web Client Extensibility',
        'Microsoft Alchemy Service', 'Radius Aad Syncer', 'MsgDataMgmt', 'O365 Secure Score', 
        'Microsoft Intune', 'ZTNA Network Access Control Plane', 'IC3 Long Running Operations Service',
        'Microsoft Rights Management Services', 'Azure AD Identity Governance - Entitlement Management',
        'Data Classification Service', 'Office365 Shell WCSS-Server Default', 'AADReporting',
        'Azure AD Notification', 'Windows Update for Business Deployment Service', 'Microsoft To-Do',
        'Microsoft Azure Workflow', 'Microsoft Graph', 'ProjectWorkManagement', 'Microsoft Teams UIS',
        'Microsoft Teams - Teams And Channels Service', 'Bing', 'OneProfile Service', 'AAD Terms Of Use',
        'Azure Multi-Factor Auth Connector', 'Viva Engage', 'OfficeServicesManager', 
        'Office MRO Device Manager Service', 'PowerApps Service', 'Microsoft Information Protection Sync Service',
        'Microsoft Cloud App Security', 'Office365 Shell SS-Server Default', 'Cortana Runtime Service',
        'Azure Resource Graph', 'Microsoft Threat Protection', 'AAD Request Verification Service - PROD',
        'PPE-DataResidencyService', 'CPIM Service', 'SharePoint Online Web Client Extensibility Isolated',
        'PushChannel', 'Microsoft Teams Chat Aggregator', 'Lifecycle Workflows', 'Microsoft Intune Service Discovery',
        'Skype for Business', 'MS-PIM', 'Centralized Deployment', 'SharePoint Home Notifier', 
        'Office 365 Enterprise Insights', 'Azure AD Application Proxy', 'Office365DirectorySynchronizationService',
        'MIP Exchange Solutions - Teams', 'Signup', 'Microsoft Office 365 Portal', 'OneNote', 
        'Narada Notification Service', 'Office 365 Configure', 'Dynamics Lifecycle services', 
        'Azure AD Identity Protection', 'Media Analysis and Transformation Service', 'PowerApps-Advisor',
        'OMSAuthorizationServicePROD', 'OfficeClientService', 'Microsoft SharePoint Online - SharePoint Home', 
        'Cortana at Work Service', 'Microsoft.Azure.SyncFabric', 'Microsoft Teams Services', 'Microsoft Teams Graph Service',
        'Power BI Service', 'Common Data Service License Management', 'SharePoint Framework Azure AD Helper', 
        'CAP Neptune Prod CM Prod', 'Audit GraphAPI Application', 'Microsoft People Cards Service', 
        'AzureSupportCenter', 'Microsoft Insider Risk Management', 'Microsoft B2B Admin Worker',
        'Office 365 Management APIs', 'Microsoft_Azure_Support', 'Office Shredding Service', 
        'Request Approvals Read Platform', 'Microsoft Azure Signup Portal', 'My Apps', 
        'Conference Auto Attendant', 'Microsoft Graph Change Tracking', 'Office365 Zoom', 
        'MS Teams Griffin Assistant', 'Customer Experience Platform PROD', 'CompliancePolicy',
        'Messaging Bot API Application', 'Microsoft Modern Contact Master', 'DeploymentScheduler', 
        'Sway', 'ZTNA Policy Service Graph Client', 'M365 Label Analytics', 'Virtual Visits App', 
        'API Connectors 1st Party', 'Connectors', 'Skype Presence Service', 'Microsoft Intune SCCM Connector', 
        'IPSubstrate', 'Azure Portal', 'Microsoft Mobile Application Management', 'Microsoft Service Trust', 
        'M365 Compliance Drive', 'Portfolios', 'Microsoft Flow CDS Integration Service', 
        'MIP Exchange Solutions - ODB', 'Skype Core Calling Service', 'Microsoft Partner Center', 
        'OCaaS Experience Management Service', 'Microsoft App Access Panel', 'Microsoft Teams AuditService', 
        'Microsoft Device Management Checkin', 'OCaaS Worker Services', 'Azure Advanced Threat Protection', 
        'Microsoft Intune API', 'Internet resources with Global Secure Access', 'Microsoft Graph Connectors Core',
        'Windows 365', 'o365.servicecommunications.microsoft.com', 'Policy Administration Service', 
        'Microsoft Forms', 'Intune CMDeviceService', 'Azure MFA StrongAuthenticationService', 
        'SharePoint Notification Service', 'Azure Credential Configuration Endpoint Service', 
        'Intune Grouping and Targeting Client Prod', 'Microsoft Approval Management', 
        'Microsoft Substrate Management', 'Azure ESTS Service', 'Skype Teams Calling API Service', 
        'Microsoft Office Licensing Service Agents', 'Azure Multi-Factor Auth Client', 
        'Microsoft Intune Advanced Threat Protection Integration', 'Microsoft Information Protection API', 
        'Dynamics 365 Viva Sales', 'Microsoft Teams AuthSvc', 'Azure Purview', 
        'Microsoft Windows AutoPilot Service API', 'Microsoft.SMIT', 'Substrate-FileWatcher', 
        'Yggdrasil', 'Office 365 Exchange Online', 'Microsoft Teams ATP Service', 
        'People Profile Event Proxy', 'Microsoft 365 Security and Compliance Center', 
        'IDML Graph Resolver Service and CAD', 'Windows Azure Active Directory', 
        'Device Registration Service', 'Microsoft Office Licensing Service vNext', 
        'MIP Exchange Solutions', 'IpLicensingService', 'Microsoft apps with Global Secure Access', 
        'Office 365 Search Service', 'Microsoft AppPlat EMA', 'Microsoft Invitation Acceptance Portal', 
        'Dynamics Data Integration', 'Customer Service Trial PVA', 'O365 Demeter', 
        'WindowsDefenderATP', 'Customer Service Trial PVA - readonly', 'Microsoft Device Management EMM API',
        'MCAPI Authorization Prod', 'MIP Exchange Solutions - SPO', 'Windows Azure Service Management API', 
        'M365 License Manager', 'Teams NRT DLP Service', 'Microsoft O365 Scuba', 
        'Meeting Migration Service', 'IAM Supportability', 'MicrosoftEndpointDLP', 'M365 Admin Services', 
        'Graph Connector Service', 'Configuration Manager Microservice', 'Power Platform Environment Discovery Service', 
        'Exchange Rbac', 'SPAuthEvent', 'Conferencing Virtual Assistant', 'Dataverse', 
        'OCaaS Client Interaction Service', 'Teams EHR Connector', 'All private resources with Global Secure Access', 
        'Substrate Instant Revocation Pipeline', 'Windows Store for Business', 'Linkedin', 'BrowserStack'
    )


    return $builtInApps -contains $app.DisplayName
}

function Get-AppPermissions($servicePrincipal) {
    $appId = $servicePrincipal.AppId
    $permissions = @{
        AppRoles = @()
        DelegatedPermissions = @()
    }

    # Get all app role assignments
    $appRoles = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.Id -ErrorAction SilentlyContinue
    
    # Get all OAuth2 permission grants
    $oauth2Permissions = Get-MgServicePrincipalOAuth2PermissionGrant -ServicePrincipalId $servicePrincipal.Id -ErrorAction SilentlyContinue

    # Process Application Permissions (AppRoles)
    foreach ($role in $appRoles) {
        try {
            $resourceSp = Get-MgServicePrincipal -Filter "Id eq '$($role.ResourceId)'" -ErrorAction SilentlyContinue
            if ($resourceSp) {
                $roleDetails = $resourceSp.AppRoles | Where-Object { $_.Id -eq $role.AppRoleId }
                if ($roleDetails) {
                    $permissions.AppRoles += "$($resourceSp.DisplayName) - $($roleDetails.Value)"
                }
            }
        } catch {
            Write-Verbose "Error processing AppRole for $($servicePrincipal.DisplayName): $_"
        }
    }

    # Process Delegated Permissions (OAuth2PermissionGrants)
    foreach ($permission in $oauth2Permissions) {
        try {
            $resourceSp = Get-MgServicePrincipal -Filter "Id eq '$($permission.ResourceId)'" -ErrorAction SilentlyContinue
            if ($resourceSp) {
                $scopes = $permission.Scope -split " "
                foreach ($scope in $scopes) {
                    $permissions.DelegatedPermissions += "$($resourceSp.DisplayName) - $scope"
                }
            }
        } catch {
            Write-Verbose "Error processing OAuth2 permission for $($servicePrincipal.DisplayName): $_"
        }
    }

    # Explicitly set to "None" if no permissions are found
    if ($permissions.AppRoles.Count -eq 0) {
        $permissions.AppRoles = @("None")
    }
    if ($permissions.DelegatedPermissions.Count -eq 0) {
        $permissions.DelegatedPermissions = @("None")
    }

    return $permissions
}


# Function to format permissions for HTML, separating each permission with its own span tag
function Format-Permissions($permissionList) {
    # If the permission list is null, empty, or contains only whitespace
    if ($null -eq $permissionList -or 
        $permissionList.Count -eq 0 -or 
        [string]::IsNullOrWhiteSpace($permissionList) -or 
        $permissionList[0] -eq "None") {
        return "<span class='permission-tag none-permission'>None</span>"
    }

    # Split the permission list by semicolon and process each permission individually
    return ($permissionList -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { 
        $escapedPermission = [System.Web.HttpUtility]::HtmlEncode($_.Trim())
        $class = if ($escapedPermission -match 'Write|ReadWrite|Mail\.Send|FullControl') { 
            'permission-tag write' 
        } else { 
            'permission-tag read' 
        }
        "<span class='$class'>$escapedPermission</span>"
    }) -join " "
}

# Define high-risk permissions patterns with weighted scores
# Adjusted risk scores for permissions with added weight for Application permissions.
$permissionRiskScores = @{
    # Critical (Weight: Application 8, Delegated 6)
    '^Directory\.ReadWrite\.All$' = @{ App = 8; Del = 6 }
    '^RoleManagement\.ReadWrite\.Directory$' = @{ App = 8; Del = 6 }
    '^Application\.ReadWrite\.All$' = @{ App = 8; Del = 6 }
    '^AppRoleAssignment\.ReadWrite\.All$' = @{ App = 8; Del = 6 }
    '^GroupMember\.ReadWrite\.All$' = @{ App = 8; Del = 6 }
    'FullControl' = @{ App = 8; Del = 6 }

    # High (Weight: Application 6, Delegated 4)
    '^Mail\.ReadWrite\.All$' = @{ App = 6; Del = 4 }
    '^Calendars\.ReadWrite\.All$' = @{ App = 6; Del = 4 }
    '^Files\.ReadWrite\.All$' = @{ App = 6; Del = 4 }
    '^User\.ReadWrite\.All$' = @{ App = 6; Del = 4 }

    # Medium-High (Weight: Application 4, Delegated 3)
    'Mail\.Send' = @{ App = 4; Del = 3 }
    '\.ReadWrite\.' = @{ App = 4; Del = 3 }
    'AccessReview\.ReadWrite\.All' = @{ App = 4; Del = 3 }
    'AuditLog\.ReadWrite\.All' = @{ App = 4; Del = 3 }

    # Medium (Weight: Application 3, Delegated 2)
    '\.Write\.' = @{ App = 3; Del = 2 }
    'Policy\.ReadWrite' = @{ App = 3; Del = 2 }
    'Chat\.ReadWrite' = @{ App = 3; Del = 2 }

    # Low (Weight: Application 2, Delegated 1)
    '\.Read\.' = @{ App = 2; Del = 1 }
    'User\.Read' = @{ App = 2; Del = 1 }
    'profile' = @{ App = 2; Del = 1 }
}

function Calculate-RiskScore($appRoles, $delegatedPermissions) {
    # Define a base score and separate totals for read-only and write/critical permissions.
    $baseScore = 1
    $appRiskScore = 0
    $delegatedRiskScore = 0
    $readOnlyScoreCap = 3   # Cap for cumulative read-only permissions to prevent high risk from reads alone.

    # Helper function to apply weighted scoring based on the permission type (App vs. Delegated)
    function Get-WeightedScore($permission, $isAppPermission) {
        foreach ($pattern in $permissionRiskScores.Keys) {
            if ($permission -match $pattern) {
                $weight = if ($isAppPermission) { $permissionRiskScores[$pattern].App } else { $permissionRiskScores[$pattern].Del }
                return $weight
            }
        }
        return 0
    }

    # Separate permissions into read-only and write/critical categories
    $appHighPrivCount = 0
    $delegatedHighPrivCount = 0
    $readOnlyAppScore = 0
    $readOnlyDelScore = 0

    foreach ($permission in $appRoles) {
        $score = Get-WeightedScore $permission $true
        if ($score -le 2) {
            $readOnlyAppScore += $score
        } else {
            $appRiskScore += $score
            if ($score -ge 6) { $appHighPrivCount++ }
        }
    }

    foreach ($permission in $delegatedPermissions) {
        $score = Get-WeightedScore $permission $false
        if ($score -le 2) {
            $readOnlyDelScore += $score
        } else {
            $delegatedRiskScore += $score
            if ($score -ge 4) { $delegatedHighPrivCount++ }
        }
    }

    # Apply caps for read-only permissions to limit their cumulative impact.
    $appRiskScore += [Math]::Min($readOnlyScoreCap, $readOnlyAppScore)
    $delegatedRiskScore += [Math]::Min($readOnlyScoreCap, $readOnlyDelScore)

    # Apply multiplier if multiple high-privilege write permissions are found
    if ($appHighPrivCount -ge 2) { $appRiskScore *= 1.5 }
    if ($delegatedHighPrivCount -ge 2) { $delegatedRiskScore *= 1.2 }

    # Total risk score calculation
    $totalRiskScore = $baseScore + [Math]::Min(10, $appRiskScore + $delegatedRiskScore)

    # Additional factors
    if ($app.PasswordCredentialsCount -gt 2) { $totalRiskScore += 1 }
    if ($app.VerifiedPublisher -eq "No") { $totalRiskScore += 1 }

    # Ensure risk score does not exceed 10 and rounds to the nearest integer
    return [Math]::Min(10, [Math]::Round($totalRiskScore))
}



# --------------------------------------------
# 4. Retrieve Applications and Enterprise Apps
# --------------------------------------------

Write-Host "🔄 Retrieving all Applications from Microsoft Graph..." -ForegroundColor Yellow
$allApps = Get-MgApplication -All -ErrorAction SilentlyContinue

if (-not $includeBuiltIn) {  # Changed condition
    Write-Host "🔄 Filtering out built-in Microsoft applications..." -ForegroundColor Yellow
    $filteredApps = $allApps | Where-Object { -not (Is-BuiltInMicrosoftApp $_) }
    $totalAppsToProcess = $filteredApps.Count
    Write-Host "✅ Found $totalAppsToProcess non-built-in Applications to process!" -ForegroundColor Green
} else {
    $filteredApps = $allApps
    $totalAppsToProcess = $allApps.Count
    Write-Host "✅ Retrieved $totalAppsToProcess Applications!" -ForegroundColor Green
}

# Process Applications
if ($filteredApps.Count -eq 0) {
    Write-Host "⚠️ No Applications found to process." -ForegroundColor Yellow
} else {
    Write-Host "🚀 Starting to Process Applications..." -ForegroundColor Yellow
    $processedApps = 0
    
    $csvData = @()

    foreach ($app in $filteredApps) {
        $processedApps++
        $progressPercent = [math]::Round(($processedApps / $totalAppsToProcess) * 100)
        Write-Progress -Activity "Processing Applications" -Status "$processedApps of $totalAppsToProcess" -PercentComplete $progressPercent

        # Process application permissions
        $sp = Get-MgServicePrincipal -Filter "AppId eq '$($app.AppId)'" -ErrorAction SilentlyContinue
        if ($sp) {
            $permissions = Get-AppPermissions $sp
            $appRolesString = if ($permissions.AppRoles.Count -gt 0) { $permissions.AppRoles -join '; ' } else { "None" }
            $delegatedPermissionsString = if ($permissions.DelegatedPermissions.Count -gt 0) { $permissions.DelegatedPermissions -join '; ' } else { "None" }
            
            # Calculate risk score based on permissions and other factors
            $riskScore = Calculate-RiskScore $permissions.AppRoles $permissions.DelegatedPermissions

            # Calculate days until credential expiration
            $latestCredentialExpiration = ($app.PasswordCredentials | Sort-Object EndDateTime -Descending | Select-Object -First 1).EndDateTime
            $daysUntilExpiry = if ($null -ne $latestCredentialExpiration) { ($latestCredentialExpiration - (Get-Date)).Days } else { $null }

            # Add data to csvData array
            $csvData += [PSCustomObject]@{
                DisplayName = $app.DisplayName
                AppType = if (Is-BuiltInMicrosoftApp $app) { "Built-In" } else { "Custom" }
                ApplicationID = $app.AppId
                ObjectID = $app.Id
                Created = $app.CreatedDateTime
                LatestCredentialExpiration = $latestCredentialExpiration
                SignInAudience = $app.SignInAudience
                VerifiedPublisher = if ($app.VerifiedPublisher) { "Yes" } else { "No" }
                PasswordCredentialsCount = ($app.PasswordCredentials | Measure-Object).Count
                Owners = "N/A"  # You might want to add actual owners logic here
                ApplicationPermissions = $appRolesString
                DelegatedPermissions = $delegatedPermissionsString
                RiskScore = [Math]::Min(10, $riskScore)  # Cap at 10
                DaysUntilExpiry = $daysUntilExpiry
                ExpiryStatus = if ($null -eq $daysUntilExpiry) { "No Expiration" } elseif ($daysUntilExpiry -lt 0) { "Expired" } elseif ($daysUntilExpiry -le 30) { "Expiring Soon" } else { "Valid" }
            }
        }
    }
    
    Write-Progress -Activity "Processing Applications" -Completed
}

# Now handle Enterprise Apps
Write-Host "🔄 Retrieving all Enterprise Apps (Service Principals) from Microsoft Graph..." -ForegroundColor Yellow
$allServicePrincipals = Get-MgServicePrincipal -All -ErrorAction SilentlyContinue

if (-not $includeBuiltIn) {  # Changed condition
    Write-Host "🔄 Filtering out built-in Microsoft Enterprise Apps..." -ForegroundColor Yellow
    $filteredServicePrincipals = $allServicePrincipals | Where-Object { -not (Is-BuiltInMicrosoftApp $_) }
    $totalSPsToProcess = $filteredServicePrincipals.Count
    Write-Host "✅ Found $totalSPsToProcess non-built-in Enterprise Apps to process!" -ForegroundColor Green
} else {
    $filteredServicePrincipals = $allServicePrincipals
    $totalSPsToProcess = $allServicePrincipals.Count
    Write-Host "✅ Retrieved $totalSPsToProcess Enterprise Apps!" -ForegroundColor Green
}

# Process Enterprise Apps
if ($filteredServicePrincipals.Count -eq 0) {
    Write-Host "⚠️ No Enterprise Apps found to process." -ForegroundColor Yellow
} else {
    Write-Host "🚀 Starting to Process Enterprise Apps..." -ForegroundColor Yellow
    $processedSPs = 0
    
    foreach ($sp in $filteredServicePrincipals) {
        $processedSPs++
        $progressPercent = [math]::Round(($processedSPs / $totalSPsToProcess) * 100)
        Write-Progress -Activity "Processing Enterprise Apps" -Status "$processedSPs of $totalSPsToProcess" -PercentComplete $progressPercent

        # Process service principal permissions
        $permissions = Get-AppPermissions $sp
        $appRolesString = if ($permissions.AppRoles.Count -gt 0) { $permissions.AppRoles -join '; ' } else { "None" }
        $delegatedPermissionsString = if ($permissions.DelegatedPermissions.Count -gt 0) { $permissions.DelegatedPermissions -join '; ' } else { "None" }

        # Calculate risk score
        $riskScore = Calculate-RiskScore $permissions.AppRoles $permissions.DelegatedPermissions

        # Add data to csvData array
        $csvData += [PSCustomObject]@{
            DisplayName = $sp.DisplayName
            AppType = if (Is-BuiltInMicrosoftApp $sp) { "Built-In" } else { "Enterprise" }
            ApplicationID = $sp.AppId
            ObjectID = $sp.Id
            Created = $sp.CreatedDateTime
            LatestCredentialExpiration = "N/A"
            SignInAudience = "N/A"
            VerifiedPublisher = if ($sp.VerifiedPublisher) { "Yes" } else { "No" }
            PasswordCredentialsCount = 0
            Owners = "N/A"
            ApplicationPermissions = $appRolesString
            DelegatedPermissions = $delegatedPermissionsString
            RiskScore = [Math]::Min(10, $riskScore)  # Cap at 10
            DaysUntilExpiry = $null
            ExpiryStatus = "No Expiration"
        }
    }
    
    Write-Progress -Activity "Processing Enterprise Apps" -Completed
}

# --------------------------------------------
# 8. Generating HTML Report
# --------------------------------------------

# Initialize HTML content with the header and styles
$script:htmlContent = @"
<!DOCTYPE html>
<html lang='en'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>Applications and Enterprise Apps Permissions Report</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; max-width: 1200px; margin: 0 auto; padding: 20px; background-color: #f5f5f5; }
        h1 { color: #2c3e50; text-align: center; margin-bottom: 30px; }
        
        /* Info Section Styles */
        .info-section {
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            margin: 20px 0;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .info-section h2 {
            color: #2c3e50;
            margin-bottom: 20px;
            border-bottom: 2px solid #3498db;
            padding-bottom: 10px;
        }
        
        .info-table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            background-color: white;
        }
        
        .info-table th,
        .info-table td {
            padding: 12px;
            border: 1px solid #e0e0e0;
            text-align: left;
        }
        
        .info-table th {
            background-color: #3498db;
            color: white;
            font-weight: bold;
        }
        
        .info-table tr:nth-child(even) {
            background-color: #f8f9fa;
        }
        
        .info-table tr:hover {
            background-color: #f1f1f1;
        }
        
        .filter-section { margin-bottom: 20px; display: flex; align-items: center; gap: 15px; background-color: #ffffff; padding: 15px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        .filter-group label { font-weight: bold; }
        #exportBtn { background-color: #2ecc71; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-size: 1em; transition: background-color 0.3s; }
        #exportBtn:hover { background-color: #27ae60; }
        .app { background-color: #ffffff; border: 1px solid #e0e0e0; border-radius: 8px; padding: 20px; margin-bottom: 30px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        .high-risk { border: 2px solid #e74c3c; }
        .critical-risk { border: 3px solid #c0392b; background-color: #ffebee; }
        .app-name { color: #3498db; font-size: 1.4em; font-weight: bold; margin-bottom: 10px; }
        .app-type { color: #8e44ad; font-size: 1em; font-weight: bold; margin-bottom: 10px; }
        .app-id, .object-id { color: #7f8c8d; font-size: 0.9em; margin-bottom: 5px; }
        .app-dates, .app-details { color: #7f8c8d; font-size: 0.9em; margin-bottom: 10px; }
        .app-details div { margin-bottom: 5px; }
        .risk-score { 
            font-weight: bold; 
            padding: 5px 10px; 
            border-radius: 4px; 
            display: inline-block; 
        }
        .risk-score.critical {
            background-color: #ff4444;
            color: white;
            animation: pulse 2s infinite;
        }
        @keyframes pulse {
            0% { transform: scale(1); }
            50% { transform: scale(1.05); }
            100% { transform: scale(1); }
        }
        .risk-low { background-color: #dff0d8; color: #3c763d; }
        .risk-medium { background-color: #fcf8e3; color: #8a6d3b; }
        .risk-high { background-color: #f2dede; color: #a94442; }
        .risk-critical { background-color: #d9534f; color: white; }
        .permissions-title { color: #2c3e50; font-weight: bold; margin-top: 15px; font-size: 1.1em; }
        .permission-type { font-weight: bold; margin-top: 10px; color: #34495e; }
        .permission-list {
            display: flex; 
            flex-wrap: wrap; 
            gap: 8px; 
            margin: 10px 0; 
            padding: 10px;
            background-color: #f8f9fa;
            border-radius: 6px;
        }
        .permission-tag {
            display: inline-flex;
            align-items: center;
            background-color: #e9ecef;
            padding: 6px 12px;
            border-radius: 15px;
            font-size: 0.9em;
            border: 1px solid #dee2e6;
            transition: all 0.2s ease;
        }
        .permission-tag:hover {
            transform: translateY(-2px);
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .permission-tag.write {
            background-color: #ffe6e6;
            border-color: #ffcccc;
            color: #cc0000;
        }
        .permission-tag.read {
            background-color: #e6f3ff;
            border-color: #cce6ff;
            color: #0066cc;
        }
        .none-permission {
            background-color: #f8f9fa;
            color: #6c757d;
            border-color: #dee2e6;
            font-style: italic;
        }
        .statistics { 
            background-color: #3498db; 
            color: white; 
            padding: 15px; 
            border-radius: 5px; 
            margin-bottom: 20px; 
            text-align: center; 
        }
         /* Info icon and tooltip styles */
        .info-icon-container {
            display: flex;
            align-items: center;
            gap: 8px;
            color: #3498db;
        }

        .info-label {
            font-weight: bold;
            font-size: 1em;
            color: #333;
        }

        .info-icon {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 24px;
            height: 24px;
            background-color: #3498db;
            color: white;
            font-weight: bold;
            font-size: 0.9em;
            border-radius: 50%;
            cursor: pointer;
            transition: transform 0.3s;
            position: relative;
        }

        .info-icon:hover {
            transform: scale(1.1);
        }

        /* Tooltip styling */
        .tooltip-content {
            display: none;
            position: absolute;
            top: -5px;
            left: 35px;
            width: 350px;
            background-color: #ffffff;
            border: 1px solid #ddd;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
            font-size: 0.9em;
            line-height: 1.4;
            color: #333;
            z-index: 1000;
        }

        .info-icon:hover .tooltip-content {
            display: block;
        }

        /* Tooltip arrow */
        .tooltip-content::after {
            content: "";
            position: absolute;
            top: 12px;
            left: -8px;
            border-width: 8px;
            border-style: solid;
            border-color: transparent #ddd transparent transparent;
        }

        .tooltip-content::before {
            content: "";
            position: absolute;
            top: 12px;
            left: -7px;
            border-width: 8px;
            border-style: solid;
            border-color: transparent #fff transparent transparent;
        }

        /* Tooltip content text */
        .tooltip-header {
            font-weight: bold;
            color: #2c3e50;
            margin-bottom: 8px;
        }

        .tooltip-content p {
            margin: 5px 0;
        }

        .tooltip-list {
            margin: 10px 0;
            padding: 0;
            list-style-type: disc;
            padding-left: 20px;
        }

        .tooltip-content strong {
            color: #e74c3c;
        }
        
        /* Navbar Styles */
        .navbar {
            background-color: #2c3e50;
            padding: 1rem;
            position: sticky;
            top: 0;
            z-index: 1000;
        }

        .nav-content {
            max-width: 1200px;
            margin: 0 auto;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .nav-menu {
            display: flex;
            gap: 2rem;
            list-style: none;
            margin: 0;
            padding: 0;
        }

        .nav-item {
            color: white;
            text-decoration: none;
            padding: 0.5rem 1rem;
            border-radius: 4px;
            transition: background-color 0.3s;
            cursor: pointer;
        }

        .nav-item:hover {
            background-color: #34495e;
        }

        /* Application Details Table */
        .app-details-table {
            width: 100%;
            border-collapse: collapse;
            margin: 10px 0;
            background-color: white;
            border-radius: 8px;
            overflow: hidden;
        }

        .app-details-table th,
        .app-details-table td {
            padding: 12px;
            border: 1px solid #e0e0e0;
        }

        .app-details-table th {
            background-color: #f8f9fa;
            font-weight: bold;
            text-align: left;
        }

        .app-details-table tr:nth-child(even) {
            background-color: #f8f9fa;
        }

        .risk-score-cell {
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .risk-indicator {
            height: 8px;
            flex: 1;
            background: linear-gradient(to right, #2ecc71, #f1c40f, #e74c3c);
            border-radius: 4px;
            position: relative;
        }

        .risk-marker {
            position: absolute;
            width: 4px;
            height: 12px;
            background-color: #2c3e50;
            top: -2px;
            transform: translateX(-50%);
        }
    </style>
</head>
<body>
    <h1>App registrations and Enterprise Apps Permissions Report</h1>
    
    <div class="info-section">
        <h2>Enterprise Applications Overview</h2>
        <table class="info-table">
            <tr>
                <th>Topic</th>
                <th>Description</th>
                <th>Security Impact</th>
            </tr>
            <tr>
                <td>Enterprise Applications</td>
                <td>Service principal objects in Microsoft Entra (formerly Azure AD) that represent applications integrated into the organization's Entra tenant. These enable users or systems to authenticate and interact with external or internal services securely.</td>
                <td>Proper configuration is crucial as these applications have access to organizational resources.</td>
            </tr>
            <tr>
                <td>Application Permissions</td>
                <td>Grant an application direct, tenant-wide access to resources, bypassing individual user context. Often used for services running without user interaction.</td>
                <td>If misconfigured, applications with tenant-wide access can expose all user data or critical systems. A breach could compromise entire organization's data.</td>
            </tr>
            <tr>
                <td>Delegated Permissions</td>
                <td>Allow an application to act on behalf of a user, inheriting the user's permissions and requiring user sign-in for authentication.</td>
                <td>If compromised, an attacker could perform actions using the user's identity, limited to the user's permissions.</td>
            </tr>
            <tr>
                <td>Tenant-wide Consent</td>
                <td>Permissions that apply across the entire organization ("consent on behalf of organization").</td>
                <td>Due to broad access, increases potential impact of a breach. When misused, sensitive data could be exposed across the organization.</td>
            </tr>
            <tr>
                <td>Individual/Group Consent</td>
                <td>Access restricted to specific users or groups.</td>
                <td>Reduces blast radius if compromised. Limits exposure of sensitive data.</td>
            </tr>
        </table>

        <h2>Recommended Security Settings</h2>
        <table class="info-table">
            <tr>
                <th>Setting</th>
                <th>Recommendation</th>
                <th>Rationale</th>
            </tr>
            <tr>
                <td>User Settings</td>
                <td>Set 'User can register applications' to No</td>
                <td>Prevents users from creating application registrations. Grant ability back to specific individuals through application developer role.</td>
            </tr>
            <tr>
                <td>User Consent</td>
                <td>Set 'User consent for applications' to 'Do not allow user consent'</td>
                <td>Ensures all application permission requests go through proper administrative review.</td>
            </tr>
            <tr>
                <td>Admin Consent Workflow</td>
                <td>Enable admin consent workflow</td>
                <td>Provides secure process for granting access to applications requiring administrator approval. Users can submit requests for admin review.</td>
            </tr>
        </table>
    </div>

    <div class="filter-section">
        <div class="filter-group">
            <label for="riskScoreFilter">Minimum Risk Score:</label>
            <select id="riskScoreFilter">
                <option value="0">All</option>
                <option value="3">3+ (Medium Risk)</option>
                <option value="5">5+ (High Risk)</option>
                <option value="8">8+ (Critical Risk)</option>
            </select>
        </div>
        <div class="filter-group">
            <label for="expiryFilter">Credential Expiration Status:</label>
            <select id="expiryFilter">
                <option value="all">All</option>
                <option value="expired">Expired</option>
                <option value="expiring">Expiring Soon</option>
                <option value="valid">Valid</option>
                <option value="none">No Expiration</option>
            </select>
        </div>
        <div class="filter-group">
            <label for="daysFilter">Days Until Expiration:</label>
            <input type="number" id="daysFilter" placeholder="Enter days">
        </div>
        
        <!-- Info icon with tooltip content -->
        <div class="info-icon-container">
            <span class="info-label">Info:</span>
            <span class="info-icon" title="Click for risk evaluation details">i
                <div class="tooltip-content">
                    <div class="tooltip-header">Risk Evaluation Details</div>
                    <p>This report assigns a risk score from <strong>1 to 10</strong> based on permissions associated with each application, focusing on sensitivity and access level:</p>
                    
                    <ul class="tooltip-list">
                        <li><strong>Application Permissions</strong> are more sensitive as they grant direct access, contributing more to the risk score.</li>
                        <li><strong>Delegated Permissions</strong> grant access through user context and have a lower impact on the score.</li>
                        <li>High-risk permissions (e.g., write access) weigh more heavily than read-only permissions.</li>
                        <li>Multiple high-privilege permissions further increase the score.</li>
                    </ul>

                    <p><strong>Scoring Breakdown:</strong></p>
                    <ul class="tooltip-list">
                        <li><strong>Critical Permissions</strong> (e.g., `Directory.ReadWrite.All`): 8 points for application, 6 for delegated.</li>
                        <li><strong>High-Risk Permissions</strong> (e.g., `Files.ReadWrite.All`): 6 points for application, 4 for delegated.</li>
                        <li><strong>Medium-Risk Permissions</strong> (e.g., `Mail.Send`): 4 points for application, 3 for delegated.</li>
                        <li><strong>Low-Risk Permissions</strong> (e.g., `User.Read`): minimal impact on score.</li>
                    </ul>

                    <p>Additional factors:</p>
                    <ul class="tooltip-list">
                        <li>More than two password credentials add 1 point.</li>
                        <li>Unverified publisher adds 1 point.</li>
                    </ul>
                    
                    <p><strong>Note:</strong> This score provides a general risk evaluation. It is important to perform your own due diligence and align these evaluations with your organization’s security policies to ensure a thorough risk assessment.</p>
                </div>
            </span>
        </div>

        <!-- Export button -->
        <button id="exportBtn" onclick="exportToCSV()">Export to CSV</button>
    </div>
</body>
"@



foreach ($row in $csvData) {
$htmlAppContent = @"
        <div class='app $(if ([int]$row.RiskScore -gt 7) { "high-risk" })' 
             data-apptype='$([System.Web.HttpUtility]::HtmlEncode($row.AppType))' 
             data-riskscore='$([System.Web.HttpUtility]::HtmlEncode($row.RiskScore))'
             data-expirystatus='$([System.Web.HttpUtility]::HtmlEncode($row.ExpiryStatus))'
             data-daysuntilexpiry='$([System.Web.HttpUtility]::HtmlEncode($row.DaysUntilExpiry))'>
            <div class='app-name'>$([System.Web.HttpUtility]::HtmlEncode($row.DisplayName))</div>
            <table class='app-details-table'>
                <tr>
                    <th>Property</th>
                    <th>Value</th>
                </tr>
                <tr>
                    <td>Application ID</td>
                    <td>$([System.Web.HttpUtility]::HtmlEncode($row.ApplicationID))</td>
                </tr>
                <tr>
                    <td>Object ID</td>
                    <td>$([System.Web.HttpUtility]::HtmlEncode($row.ObjectID))</td>
                </tr>
                <tr>
                    <td>Created</td>
                    <td>$([System.Web.HttpUtility]::HtmlEncode($row.Created))</td>
                </tr>
                <tr>
                    <td>Latest Credential Expiration</td>
                    <td>$([System.Web.HttpUtility]::HtmlEncode($row.LatestCredentialExpiration))</td>
                </tr>
                <tr>
                    <td>Days Until Expiry</td>
                    <td>$(
                        if ($null -eq $row.DaysUntilExpiry) {
                            "<span class='expiry-status expiry-none'>No expiration</span>"
                        } elseif ($row.DaysUntilExpiry -lt 0) {
                            "<span class='expiry-status expiry-expired'>Expired ($($row.DaysUntilExpiry * -1) days ago)</span>"
                        } elseif ($row.DaysUntilExpiry -le 30) {
                            "<span class='expiry-status expiry-soon'>Expiring in $($row.DaysUntilExpiry) days</span>"
                        } else {
                            "<span class='expiry-status expiry-valid'>Valid ($($row.DaysUntilExpiry) days remaining)</span>"
                        }
                    )</td>
                </tr>
                <tr>
                    <td>Sign-in Audience</td>
                    <td>$([System.Web.HttpUtility]::HtmlEncode($row.SignInAudience))</td>
                </tr>
                <tr>
                    <td>Verified Publisher</td>
                    <td>$([System.Web.HttpUtility]::HtmlEncode($row.VerifiedPublisher))</td>
                </tr>
                <tr>
                    <td>Password Credentials Count</td>
                    <td>$([System.Web.HttpUtility]::HtmlEncode($row.PasswordCredentialsCount))</td>
                </tr>
                <tr>
                    <td>Owners</td>
                    <td>$([System.Web.HttpUtility]::HtmlEncode($row.Owners))</td>
                </tr>
                <tr>
                    <td>Risk Score</td>
                    <td class='risk-score-cell'>
                        $([System.Web.HttpUtility]::HtmlEncode($row.RiskScore)) / 10
                        <div class='risk-indicator'>
                            <div class='risk-marker' style='left: $($row.RiskScore * 10)%'></div>
                        </div>
                    </td>
                </tr>
            </table>
            <div class='permissions-title'>API Permissions:</div>
            <div class='permission-type'>Application Permissions:</div>
            <div class='permission-list'>
                $(if ([string]::IsNullOrWhiteSpace($row.ApplicationPermissions)) { 
                    "<span class='permission-tag none-permission'>None</span>" 
                } else { 
                    Format-Permissions $row.ApplicationPermissions 
                })
            </div>
            <div class='permission-type'>Delegated Permissions:</div>
            <div class='permission-list'>
                $(if ([string]::IsNullOrWhiteSpace($row.DelegatedPermissions)) { 
                    "<span class='permission-tag none-permission'>None</span>" 
                } else { 
                    Format-Permissions $row.DelegatedPermissions 
                })
            </div>
        </div>
"@



    # Append HTML content for each app
    $script:htmlContent += $htmlAppContent
}

# Add JavaScript functionality to enable filtering and export
$script:htmlContent += @"
    <script>
        const csvData = $($csvData | ConvertTo-Json);

        document.addEventListener('DOMContentLoaded', function() {
            const riskScoreFilter = document.getElementById('riskScoreFilter');
            const expiryFilter = document.getElementById('expiryFilter');
            const daysFilter = document.getElementById('daysFilter');
            const apps = document.querySelectorAll('.app');

            function filterApps() {
                const selectedRiskScore = parseInt(riskScoreFilter.value);
                const selectedExpiry = expiryFilter.value;
                const selectedDays = parseInt(daysFilter.value);

                apps.forEach(app => {
                    const riskScore = parseInt(app.dataset.riskscore);
                    const expiryStatus = app.dataset.expirystatus;
                    const daysUntilExpiry = parseInt(app.dataset.daysuntilexpiry);
                    
                    let showApp = true;

                    // Risk Score Filter
                    if (selectedRiskScore > 0 && riskScore < selectedRiskScore) {
                        showApp = false;
                    }

                    // Expiry Status Filter
                    if (selectedExpiry !== 'all') {
                        switch(selectedExpiry) {
                            case 'expired':
                                if (expiryStatus !== 'Expired') showApp = false;
                                break;
                            case 'expiring':
                                if (expiryStatus !== 'Expiring Soon') showApp = false;
                                break;
                            case 'valid':
                                if (expiryStatus !== 'Valid') showApp = false;
                                break;
                            case 'none':
                                if (expiryStatus !== 'No Expiration') showApp = false;
                                break;
                        }
                    }

                    // Days Until Expiry Filter
                    if (selectedDays && !isNaN(selectedDays)) {
                        if (isNaN(daysUntilExpiry) || daysUntilExpiry > selectedDays || daysUntilExpiry < 0) {
                            showApp = false;
                        }
                    }

                    app.style.display = showApp ? 'block' : 'none';
                });
            }

            // Add event listeners to filters
            riskScoreFilter.addEventListener('change', filterApps);
            expiryFilter.addEventListener('change', filterApps);
            daysFilter.addEventListener('change', filterApps);
        });

        function exportToCSV() {
            const apps = document.getElementsByClassName('app');
            const visibleApps = Array.from(apps).filter(app => app.style.display !== 'none');
            
            if (visibleApps.length === 0) {
                alert('No visible applications to export.');
                return;
            }

            const csvData = visibleApps.map(app => {
                return {
                    'Display Name': app.querySelector('.app-name').textContent,
                    'App Type': app.querySelector('.app-type').textContent.replace('App Type: ', ''),
                    'Application ID': app.querySelector('.app-id').textContent.replace('Application ID: ', ''),
                    'Object ID': app.querySelector('.object-id').textContent.replace('Object ID: ', ''),
                    'Created': app.querySelector('.app-dates').textContent.split('|')[0].replace('Created: ', '').trim(),
                    'Latest Credential Expiration': app.querySelector('.app-dates').textContent.split('|')[1].replace('Latest Credential Expiration: ', '').trim(),
                    'Sign-in Audience': Array.from(app.querySelectorAll('.app-details div')).find(el => el.textContent.startsWith('Sign-in Audience'))?.textContent.replace('Sign-in Audience: ', '') || '',
                    'Verified Publisher': Array.from(app.querySelectorAll('.app-details div')).find(el => el.textContent.startsWith('Verified Publisher'))?.textContent.replace('Verified Publisher: ', '') || '',
                    'Password Credentials Count': Array.from(app.querySelectorAll('.app-details div')).find(el => el.textContent.startsWith('Password Credentials Count'))?.textContent.replace('Password Credentials Count: ', '') || '',
                    'Owners': Array.from(app.querySelectorAll('.app-details div')).find(el => el.textContent.startsWith('Owners'))?.textContent.replace('Owners: ', '') || '',
                    'Risk Score': Array.from(app.querySelectorAll('.app-details div')).find(el => el.textContent.startsWith('Risk Score'))?.textContent.replace('Risk Score: ', '').split('/')[0].trim() || '',
                    'Application Permissions': Array.from(app.querySelectorAll('.permission-list')[0].querySelectorAll('.permission-tag')).map(tag => tag.textContent).join('; '),
                    'Delegated Permissions': Array.from(app.querySelectorAll('.permission-list')[1].querySelectorAll('.permission-tag')).map(tag => tag.textContent).join('; ')
                };
            });

            const headers = Object.keys(csvData[0]);
            const csvContent = [
                headers.join(','),
                ...csvData.map(row => headers.map(header => {
                    let value = row[header] || '';
                    if (typeof value === 'string') {
                        value = value.replace(/"/g, '""');
                        if (value.includes(',') || value.includes('"') || value.includes('\n')) {
                            value = `"${value}"`;
                        }
                    }
                    return value;
                }).join(','))
            ].join('\n');

            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = 'ApplicationsEnterpriseAppsReport.csv';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    </script>
</body>
</html>
"@

# --------------------------------------------
# 9. Save and Open HTML Report
# --------------------------------------------

$reportPath = Join-Path $PWD.Path "ApplicationsEnterpriseAppsReport.html"
$script:htmlContent | Out-File -FilePath $reportPath -Encoding UTF8

Write-Host "📝 HTML Report has been generated and saved as ApplicationsEnterpriseAppsReport.html" -ForegroundColor Green
Start-Process $reportPath  # Open the HTML report

# --------------------------------------------
# End of Script
# --------------------------------------------
