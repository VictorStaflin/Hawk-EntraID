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
    # First check if it's a managed identity (these are built-in)
    if ($app.ServicePrincipalType -eq "ManagedIdentity") {
        return $true
    }

    # Check if it's a built-in app type
    if ($app.Tags -contains "WindowsAzureActiveDirectoryIntegratedApp" -or 
        $app.Tags -contains "WindowsAzureActiveDirectoryGalleryApplicationPrimary") {
        return $true
    }

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


    # Check if the app name is in our known list
    if ($builtInApps -contains $app.DisplayName) {
        return $true
    }

    # Check for common Microsoft app patterns
    $microsoftPatterns = @(
        '^Microsoft',
        '^MS\s',
        '^Office\s',
        '^Azure\s',
        '^Windows\s',
        '^SharePoint\s',
        '^Teams\s',
        '^Dynamics\s',
        '^Power\s',
        '^Graph\s',
        'Microsoft$',
        '\sOnline$',
        'Azure$',
        '^AAD\s',
        'Intune$',
        'Exchange$',
        'Outlook$',
        'Skype$',
        'Yammer$',
        'Defender$',
        'PowerBI$',
        'PowerApps$',
        'Flow$',
        'Substrate',
        'Cortana',
        'CPIM Service',
        'Workflow',
        'MIP',
        'Fabric',
        'Compliance',
        'Universal Print',
        'Bing',
        'OneDrive',
        'OneNote',
        'Sway',
        'Viva',
        'M365',
        'GSA-$',
        'Linkedin',
        'Box',
        'Salesforce'
    )

    foreach ($pattern in $microsoftPatterns) {
        if ($app.DisplayName -match $pattern) {
            return $true
        }
    }

    # Check publisher domain if available
    if ($app.PublisherDomain -like "*.microsoft.com") {
        return $true
    }

    # Check service principal specific properties
    if ($app.ServicePrincipalType -eq "Application") {
        # Check if it's published by Microsoft
        if ($app.PublisherName -like "*Microsoft*") {
            return $true
        }

        # Check the app ID prefix (many Microsoft apps start with specific GUIDs)
        $microsoftAppIdPrefixes = @(
            "00000001-", "00000002-", "00000003-", "00000004-",
            "00000005-", "00000006-", "00000007-", "00000008-",
            "00000009-", "0000000a-", "0000000b-", "0000000c-"
        )
        foreach ($prefix in $microsoftAppIdPrefixes) {
            if ($app.AppId -like "$prefix*") {
                return $true
            }
        }
    }

    # Additional checks for service principals
    if ($app.PSObject.Properties.Name -contains "ServicePrincipalNames") {
        foreach ($spn in $app.ServicePrincipalNames) {
            if ($spn -like "*.microsoft.com" -or 
                $spn -like "*microsoftonline*" -or 
                $spn -like "*windows.net" -or 
                $spn -like "*azure.*") {
                return $true
            }
        }
    }

    return $false
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

if (-not $includeBuiltIn) {  
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

if (-not $includeBuiltIn) {
    Write-Host "🔄 Filtering out built-in Microsoft Enterprise Apps..." -ForegroundColor Yellow
    $VerbosePreference = "Continue"
    
    $filteredServicePrincipals = $allServicePrincipals | Where-Object {
        $isBuiltIn = Is-BuiltInMicrosoftApp $_
        if ($isBuiltIn) {
            Write-Verbose "Filtering out: $($_.DisplayName) (Type: $($_.ServicePrincipalType))"
        }
        -not $isBuiltIn
    }
    
    $VerbosePreference = "SilentlyContinue"
    $totalSPsToProcess = $filteredServicePrincipals.Count
    Write-Host "✅ Found $totalSPsToProcess non-built-in Enterprise Apps to process!" -ForegroundColor Green
    Write-Host "ℹ️ Filtered out $($allServicePrincipals.Count - $totalSPsToProcess) built-in Microsoft apps" -ForegroundColor Cyan
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

# Create HTML report
$script:htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Enterprise Applications Security Report</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css" rel="stylesheet">
    <style>
        /* Core Layout Styles */
        :root {
            --primary-color: #2563eb;
            --secondary-color: #f8fafc;
            --success-color: #22c55e;
            --warning-color: #f59e0b;
            --danger-color: #ef4444;
            --text-primary: #1e293b;
            --text-secondary: #64748b;
            --border-color: #e2e8f0;
            --card-shadow: 0 1px 3px 0 rgb(0 0 0 / 0.1), 0 1px 2px -1px rgb(0 0 0 / 0.1);
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            line-height: 1.5;
            color: var(--text-primary);
            background-color: #f1f5f9;
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
        }

        .container {
            width: 100%;
            max-width: 1400px;
            margin: 0 auto;
            padding: 1.5rem;
        }

        .header {
            background-color: white;
            padding: 2rem;
            border-radius: 0.75rem;
            margin-bottom: 2rem;
            box-shadow: var(--card-shadow);
        }

        .header-top {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 1.5rem;
        }

        .export-button {
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
            padding: 0.75rem 1.25rem;
            background-color: var(--primary-color);
            color: white;
            border: none;
            border-radius: 0.5rem;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s;
            font-size: 0.875rem;
            text-decoration: none;
        }

        .export-button:hover {
            background-color: #1d4ed8;
            transform: translateY(-1px);
        }

        .export-button:active {
            transform: translateY(0);
        }

        .export-button i {
            font-size: 1rem;
        }

        .apps-grid {
            display: grid;
            gap: 1.5rem;
            grid-template-columns: repeat(auto-fill, minmax(min(100%, 400px), 1fr));
            margin-top: 2rem;
        }

        .app-card {
            display: flex !important;
            flex-direction: column;
            background-color: white;
            border-radius: 0.75rem;
            padding: 1.5rem;
            box-shadow: var(--card-shadow);
            gap: 1rem;
        }

        .app-card[style*="display: none"] {
            display: none !important;
        }

        .app-header {
            padding: 1.5rem;
            border-bottom: 1px solid var(--border-color);
            background-color: var(--secondary-color);
        }

        .app-name {
            font-size: clamp(1rem, 1.5vw, 1.125rem);
            font-weight: 600;
            color: var(--text-primary);
            margin-bottom: 0.75rem;
            word-break: break-word;
            display: flex;
            align-items: center;
        }

        .app-meta {
            display: flex;
            flex-wrap: wrap;
            gap: 1rem;
            font-size: 0.875rem;
            color: var(--text-secondary);
        }

        .app-body {
            padding: 1.5rem;
            flex: 1;
            display: flex;
            flex-direction: column;
            gap: 1rem;
        }

        .detail-item {
            display: flex;
            flex-direction: column;
            gap: 0.5rem;
            padding: 1rem;
            background-color: var(--secondary-color);
            border-radius: 0.5rem;
        }

        .detail-label {
            font-weight: 500;
            color: var(--text-secondary);
            font-size: 0.875rem;
        }

        .detail-value {
            font-weight: 500;
            color: var(--text-primary);
            word-break: break-word;
        }

        .permissions-section {
            margin-top: auto;
            padding-top: 1.5rem;
            border-top: 1px solid var(--border-color);
        }

        .permissions-title {
            font-weight: 600;
            margin-bottom: 1rem;
            display: flex;
            align-items: center;
            gap: 0.5rem;
            font-size: 0.875rem;
        }

        .permissions-list {
            background-color: var(--secondary-color);
            border-radius: 0.5rem;
            padding: 1rem;
            font-size: 0.875rem;
        }

        .app-icon {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 2rem;
            height: 2rem;
            border-radius: 0.5rem;
            background-color: var(--primary-color);
            color: white;
            margin-right: 0.75rem;
        }

        .app-icon i {
            font-size: 1rem;
        }

        .status-badge {
            display: inline-flex;
            align-items: center;
            gap: 0.375rem;
            padding: 0.375rem 0.75rem;
            border-radius: 9999px;
            font-size: 0.875rem;
            font-weight: 500;
        }

        .status-badge.expired {
            background-color: #fee2e2;
            color: #991b1b;
        }

        .status-badge.warning {
            background-color: #fef3c7;
            color: #92400e;
        }

        .status-badge.valid {
            background-color: #dcfce7;
            color: #166534;
        }

        .status-badge.no-expiry {
            background-color: #e0f2fe;
            color: #0369a1;
        }

        /* Responsive Design */
        @media (max-width: 768px) {
            .container {
                padding: 1rem;
            }

            .header {
                padding: 1.5rem;
            }

            .apps-grid {
                grid-template-columns: 1fr;
                gap: 1rem;
            }

            .app-card {
                border-radius: 0.75rem;
            }

            .app-header, .app-body {
                padding: 1rem;
            }

            .detail-item {
                padding: 0.75rem;
            }

            .permissions-section {
                padding-top: 1rem;
            }
        }

        /* Dark mode support */
        @media (prefers-color-scheme: dark) {
            body {
                background-color: #0f172a;
                color: #f8fafc;
            }

            .header, .app-card {
                background-color: #1e293b;
            }

            .detail-item, .permissions-list {
                background-color: #334155;
            }

            select, input {
                background-color: #334155;
                color: #f8fafc;
                border-color: #475569;
            }

            .app-header {
                background-color: #1e293b;
                border-color: #475569;
            }
        }

        /* Security Enhancements - Keep these at the end to override base styles */
        [data-tooltip] {
            position: relative;
            cursor: help;
        }

        [data-tooltip]:before {
            content: attr(data-tooltip);
            position: absolute;
            bottom: 100%;
            left: 50%;
            transform: translateX(-50%);
            padding: 8px 12px;
            background-color: #1e293b;
            color: white;
            border-radius: 6px;
            font-size: 0.875rem;
            white-space: normal;
            max-width: 300px;
            opacity: 0;
            visibility: hidden;
            transition: all 0.2s;
            z-index: 1000;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
            pointer-events: none;
        }

        [data-tooltip]:hover:before {
            opacity: 1;
            visibility: visible;
            bottom: calc(100% + 5px);
        }

        .risk-score {
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
            padding: 0.5rem 1rem;
            border-radius: 9999px;
            font-weight: 600;
        }

        .risk-score.critical {
            background-color: #dc2626 !important;
            color: white !important;
        }

        .risk-score.high {
            background-color: #ef4444 !important;
            color: white !important;
        }

        .risk-score.medium {
            background-color: #f97316 !important;
            color: white !important;
        }

        .risk-score.low {
            background-color: #22c55e !important;
            color: white !important;
        }

        .permission-item {
            margin: 4px 0;
            padding: 8px 12px;
            border-radius: 4px;
            display: flex;
            align-items: center;
            gap: 8px;
            transition: all 0.2s ease;
        }

        .permission-item.critical {
            background-color: #fee2e2 !important;
            color: #991b1b !important;
            border-left: 4px solid #dc2626;
        }

        .permission-item.high {
            background-color: #fef2f2 !important;
            color: #991b1b !important;
            border-left: 4px solid #ef4444;
        }

        .permission-item.elevated {
            background-color: #faf5ff !important;
            color: #6b21a8 !important;
            border-left: 4px solid #9333ea;
        }

        .permission-item.moderate {
            background-color: #eff6ff !important;
            color: #1e40af !important;
            border-left: 4px solid #3b82f6;
        }

        .permission-item.low {
            background-color: #f0fdf4 !important;
            color: #166534 !important;
            border-left: 4px solid #22c55e;
        }

        @media (prefers-color-scheme: dark) {
            .permission-item.critical {
                background-color: rgba(220, 38, 38, 0.1) !important;
                color: #fca5a5 !important;
            }

            .permission-item.high {
                background-color: rgba(239, 68, 68, 0.1) !important;
                color: #fca5a5 !important;
            }

            .permission-item.elevated {
                background-color: rgba(147, 51, 234, 0.1) !important;
                color: #e9d5ff !important;
            }

            .permission-item.moderate {
                background-color: rgba(59, 130, 246, 0.1) !important;
                color: #bfdbfe !important;
            }

            .permission-item.low {
                background-color: rgba(34, 197, 94, 0.1) !important;
                color: #86efac !important;
            }

            [data-tooltip]:before {
                background-color: #0f172a;
                border: 1px solid #334155;
            }
        }
        
        .filter-stats {
            margin-top: 1rem;
            padding: 0.5rem;
            font-size: 0.875rem;
            color: var(--text-secondary);
            text-align: right;
            border-top: 1px solid var(--border-color);
        }
        
        .filters {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 1rem;
            background-color: white;
            padding: 1.5rem;
            border-radius: 0.5rem;
            box-shadow: var(--card-shadow);
            margin-top: 1rem;
        }

        .filter-group {
            display: flex;
            flex-direction: column;
            gap: 0.5rem;
        }

        .filter-group label {
            font-size: 0.875rem;
            font-weight: 500;
            color: var(--text-secondary);
        }

        .filter-group input,
        .filter-group select {
            width: 100%;
            padding: 0.75rem;
            border: 1px solid var(--border-color);
            border-radius: 0.5rem;
            font-size: 0.875rem;
            color: var(--text-primary);
            background-color: white;
            transition: all 0.2s;
            outline: none;
        }

        .filter-group input:focus,
        .filter-group select:focus {
            border-color: var(--primary-color);
            box-shadow: 0 0 0 2px rgba(59, 130, 246, 0.1);
        }

        .filter-group input::placeholder {
            color: var(--text-secondary);
            opacity: 0.7;
        }

        .filter-group select {
            cursor: pointer;
            background-image: url("data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' fill='none' viewBox='0 0 20 20'%3e%3cpath stroke='%236b7280' stroke-linecap='round' stroke-linejoin='round' stroke-width='1.5' d='M6 8l4 4 4-4'/%3e%3c/svg%3e");
            background-position: right 0.5rem center;
            background-repeat: no-repeat;
            background-size: 1.5em 1.5em;
            padding-right: 2.5rem;
            -webkit-appearance: none;
            -moz-appearance: none;
            appearance: none;
        }

        .filter-stats {
            grid-column: 1 / -1;
            margin-top: 0.5rem;
            padding-top: 0.5rem;
            border-top: 1px solid var(--border-color);
            font-size: 0.875rem;
            color: var(--text-secondary);
            text-align: right;
        }

        @media (max-width: 768px) {
            .filters {
                grid-template-columns: 1fr;
                padding: 1rem;
            }
        }

        @media (prefers-color-scheme: dark) {
            .filters {
                background-color: var(--dark-card-bg);
            }

            .filter-group input,
            .filter-group select {
                background-color: var(--dark-card-bg);
                border-color: var(--dark-border-color);
                color: var(--dark-text-primary);
            }

            .filter-group input::placeholder {
                color: var(--dark-text-secondary);
            }

            .filter-group select {
                background-image: url("data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' fill='none' viewBox='0 0 20 20'%3e%3cpath stroke='%239ca3af' stroke-linecap='round' stroke-linejoin='round' stroke-width='1.5' d='M6 8l4 4 4-4'/%3e%3c/svg%3e");
            }

            .filter-stats {
                border-color: var(--dark-border-color);
                color: var(--dark-text-secondary);
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="header-top">
                <h1>Enterprise Applications Security Report</h1>
                <button class="export-button" onclick="exportToCSV()">
                    <i class="fas fa-file-export"></i>
                    Export to CSV
                </button>
            </div>
            <div class="filters">
                <div class="filter-group">
                    <label for="appNameFilter">Search Applications</label>
                    <input type="text" id="appNameFilter" placeholder="Type to search...">
                </div>
                <div class="filter-group">
                    <label for="riskScoreFilter">Risk Level</label>
                    <select id="riskScoreFilter">
                        <option value="all">All Risk Levels</option>
                        <option value="critical">Critical (8-10)</option>
                        <option value="high">High (6-7)</option>
                        <option value="medium">Medium (4-5)</option>
                        <option value="low">Low (0-3)</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="appTypeFilter">Application Type</label>
                    <select id="appTypeFilter">
                        <option value="all">All Types</option>
                        <option value="internal">Internal</option>
                        <option value="external">External</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label for="expiryStatusFilter">Expiry Status</label>
                    <select id="expiryStatusFilter">
                        <option value="all">All Statuses</option>
                        <option value="valid">Valid</option>
                        <option value="expired">Expired</option>
                        <option value="no-expiry">No Expiry</option>
                    </select>
                </div>
            </div>
            <div class="filter-stats">Showing all applications</div>
        </div>
        <div class="apps-grid">
"@

foreach ($row in $csvData) {
    # Determine if the app is internal or external based on SignInAudience
    $appType = if ($row.SignInAudience -like "*AzureADMyOrg*") { 
        "internal" 
    } else { 
        "external" 
    }
    
    # Determine expiry status
    $expiryStatus = if ($null -eq $row.DaysUntilExpiry) { 
        "no-expiry" 
    } elseif ($row.DaysUntilExpiry -lt 0) { 
        "expired" 
    } else { 
        "valid" 
    }

    # Calculate risk score
    $riskScore = [int]$row.RiskScore

    $script:htmlContent += @"
            <div class='app-card' 
                data-apptype='$appType' 
                data-expirystatus='$expiryStatus'
                data-riskscore='$riskScore'>
                <div class='app-header'>
                    <div class='app-title'>
                        <div class='app-icon'>
                            <i class='fas fa-puzzle-piece'></i>
                        </div>
                        <div class='app-name'>$([System.Web.HttpUtility]::HtmlEncode($row.DisplayName))</div>
                    </div>
                    <div class='app-meta'>
                        <div class='risk-score $(
                            if ($riskScore -ge 8) { "critical" }
                            elseif ($riskScore -ge 6) { "high" }
                            elseif ($riskScore -ge 4) { "medium" }
                            else { "low" }
                        )'>
                            <i class='fas $(
                                if ($riskScore -ge 8) { "fa-radiation-alt" }
                                elseif ($riskScore -ge 6) { "fa-exclamation-triangle" }
                                elseif ($riskScore -ge 4) { "fa-shield-alt" }
                                else { "fa-check-circle" }
                            )'></i>
                            Risk Score: $riskScore
                        </div>
                        <div class='status-badge $expiryStatus'>
                            <i class='fas $(
                                if ($null -eq $row.DaysUntilExpiry) { "fa-infinity" }
                                elseif ($row.DaysUntilExpiry -lt 0) { "fa-exclamation-circle" }
                                else { "fa-check-circle" }
                            )'></i>
                            $(
                                if ($null -eq $row.DaysUntilExpiry) {
                                    "No Expiry"
                                } elseif ($row.DaysUntilExpiry -lt 0) {
                                    "Expired ($($row.DaysUntilExpiry * -1) days ago)"
                                } else {
                                    "Valid ($($row.DaysUntilExpiry) days)"
                                }
                            )
                        </div>
                    </div>
                </div>
                <div class='app-body'>
                    <div class='detail-item' data-tooltip="Unique identifier for the application in Azure AD">
                        <div class='detail-label'>Application ID:</div>
                        <div class='detail-value'>$([System.Web.HttpUtility]::HtmlEncode($row.ApplicationID))</div>
                    </div>
                    <div class='detail-item' data-tooltip="Internal Azure AD object identifier">
                        <div class='detail-label'>Object ID:</div>
                        <div class='detail-value'>$([System.Web.HttpUtility]::HtmlEncode($row.ObjectID))</div>
                    </div>
                    <div class='detail-item' data-tooltip="Date when the application was registered in Azure AD">
                        <div class='detail-label'>Created:</div>
                        <div class='detail-value'>$([System.Web.HttpUtility]::HtmlEncode($row.Created))</div>
                    </div>
                    <div class='detail-item' data-tooltip="The expiration date of the most recently expiring credential">
                        <div class='detail-label'>Latest Credential Expiration:</div>
                        <div class='detail-value'>$([System.Web.HttpUtility]::HtmlEncode($row.LatestCredentialExpiration))</div>
                    </div>
                    <div class='detail-item' data-tooltip="Defines the type of user accounts that can access the application">
                        <div class='detail-label'>Sign-in Audience:</div>
                        <div class='detail-value'>$([System.Web.HttpUtility]::HtmlEncode($row.SignInAudience))</div>
                    </div>
                    <div class='detail-item' data-tooltip="Indicates if the application publisher has been verified by Microsoft">
                        <div class='detail-label'>Verified Publisher:</div>
                        <div class='detail-value'>$([System.Web.HttpUtility]::HtmlEncode($row.VerifiedPublisher))</div>
                    </div>
                    <div class='detail-item' data-tooltip="Number of active client secrets configured for the application">
                        <div class='detail-label'>Password Credentials Count:</div>
                        <div class='detail-value'>$([System.Web.HttpUtility]::HtmlEncode($row.PasswordCredentialsCount))</div>
                    </div>
                    <div class='detail-item' data-tooltip="Users who have administrative access to manage this application">
                        <div class='detail-label'>Owners:</div>
                        <div class='detail-value'>$([System.Web.HttpUtility]::HtmlEncode($row.Owners))</div>
                    </div>
                    <div class='permissions-section'>
                        <div class='permissions-title' data-tooltip="Direct API permissions granted to the application">
                            <i class='fas fa-key'></i>
                            Application Permissions
                        </div>
                        <div class='permissions-list'>
                            $(if ([string]::IsNullOrWhiteSpace($row.ApplicationPermissions)) { 
                                "<div class='permission-item low' data-tooltip='No permissions have been granted'><i class='fas fa-ban'></i>None</div>" 
                            } else { 
                                ($row.ApplicationPermissions -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { 
                                    $permClass = if ($_ -match '(FullControl|full_access|admin|.+Admin)') { 'critical' }
                                               elseif ($_ -match '(Write|Delete|Create|Manage)') { 'high' }
                                               elseif ($_ -match 'ReadWrite') { 'elevated' }
                                               elseif ($_ -match 'Read') { 'moderate' }
                                               else { 'low' }
                                    
                                    $tooltip = switch ($permClass) {
                                        'critical' { "CRITICAL RISK: Full administrative access - Requires immediate security review" }
                                        'high' { "HIGH RISK: Write/modify permissions - Can make system changes" }
                                        'elevated' { "ELEVATED RISK: Combined read-write access - Extended privileges" }
                                        'moderate' { "MODERATE RISK: Read-only access - Limited data exposure" }
                                        'low' { "LOW RISK: Basic access - Minimal security impact" }
                                    }
                                    
                                    $icon = switch ($permClass) {
                                        'critical' { 'fa-user-shield' }
                                        'high' { 'fa-pen-fancy' }
                                        'elevated' { 'fa-edit' }
                                        'moderate' { 'fa-eye' }
                                        'low' { 'fa-info-circle' }
                                    }
                                    
                                    "<div class='permission-item $permClass' data-tooltip='$tooltip'><i class='fas $icon'></i>$([System.Web.HttpUtility]::HtmlEncode($_.Trim()))</div>"
                                }) -join ""
                            })
                        </div>
                        <div class='permissions-title' data-tooltip="Permissions that the application has been granted to access other applications on behalf of users">
                            <i class='fas fa-user-shield'></i>
                            Delegated Permissions
                        </div>
                        <div class='permissions-list'>
                            $(if ([string]::IsNullOrWhiteSpace($row.DelegatedPermissions)) { 
                                "<div class='permission-item low' data-tooltip='No permissions have been granted'><i class='fas fa-ban'></i>None</div>" 
                            } else { 
                                ($row.DelegatedPermissions -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { 
                                    $permClass = if ($_ -match '(FullControl|full_access|admin|.+Admin)') { 'critical' }
                                               elseif ($_ -match '(Write|Delete|Create|Manage)') { 'high' }
                                               elseif ($_ -match 'ReadWrite') { 'elevated' }
                                               elseif ($_ -match 'Read') { 'moderate' }
                                               else { 'low' }
                                    
                                    $tooltip = switch ($permClass) {
                                        'critical' { "CRITICAL RISK: Full administrative access - Requires immediate security review" }
                                        'high' { "HIGH RISK: Write/modify permissions - Can make system changes" }
                                        'elevated' { "ELEVATED RISK: Combined read-write access - Extended privileges" }
                                        'moderate' { "MODERATE RISK: Read-only access - Limited data exposure" }
                                        'low' { "LOW RISK: Basic access - Minimal security impact" }
                                    }
                                    
                                    $icon = switch ($permClass) {
                                        'critical' { 'fa-user-shield' }
                                        'high' { 'fa-pen-fancy' }
                                        'elevated' { 'fa-edit' }
                                        'moderate' { 'fa-eye' }
                                        'low' { 'fa-info-circle' }
                                    }
                                    
                                    "<div class='permission-item $permClass' data-tooltip='$tooltip'><i class='fas $icon'></i>$([System.Web.HttpUtility]::HtmlEncode($_.Trim()))</div>"
                                }) -join ""
                            })
                        </div>
                    </div>
                </div>
            </div>
"@
}
$script:htmlContent += @"
        </div>
    </div>
    <script>
        // Debug function to help troubleshoot filtering
        function debugElement(element, attribute) {
            if (attribute) {
                console.log('Debug ' + attribute + ':', element ? element.getAttribute(attribute) : 'Element not found');
            } else {
                console.log('Debug element:', element ? element.textContent : 'Element not found');
            }
        }

        document.addEventListener('DOMContentLoaded', function() {
            const appNameFilter = document.getElementById('appNameFilter');
            const riskScoreFilter = document.getElementById('riskScoreFilter');
            const appTypeFilter = document.getElementById('appTypeFilter');
            const expiryStatusFilter = document.getElementById('expiryStatusFilter');
            const filterStats = document.querySelector('.filter-stats');
            const appCards = document.querySelectorAll('.app-card');

            function filterCards() {
                console.log('Filtering started...');
                const searchTerm = appNameFilter.value.toLowerCase();
                const riskLevel = riskScoreFilter.value;
                const appType = appTypeFilter.value;
                const expiryStatus = expiryStatusFilter.value;

                let visibleCount = 0;

                appCards.forEach(card => {
                    // Debug logging
                    console.log('Processing card:', card);
                    debugElement(card.querySelector('.app-name'));
                    debugElement(card, 'data-riskscore');
                    debugElement(card, 'data-apptype');
                    debugElement(card, 'data-expirystatus');

                    const appName = card.querySelector('.app-name').textContent.toLowerCase();
                    const riskScore = parseInt(card.getAttribute('data-riskscore'));
                    const cardType = card.getAttribute('data-apptype');
                    const cardExpiry = card.getAttribute('data-expirystatus');

                    const matchesSearch = appName.includes(searchTerm);
                    const matchesRisk = riskLevel === 'all' || 
                        (riskLevel === 'critical' && riskScore >= 8) ||
                        (riskLevel === 'high' && riskScore >= 6 && riskScore < 8) ||
                        (riskLevel === 'medium' && riskScore >= 4 && riskScore < 6) ||
                        (riskLevel === 'low' && riskScore < 4);
                    const matchesType = appType === 'all' || cardType === appType;
                    const matchesExpiry = expiryStatus === 'all' || cardExpiry === expiryStatus;

                    // Debug logging
                    console.log({
                        appName: appName,
                        riskScore: riskScore,
                        cardType: cardType,
                        cardExpiry: cardExpiry,
                        matchesSearch: matchesSearch,
                        matchesRisk: matchesRisk,
                        matchesType: matchesType,
                        matchesExpiry: matchesExpiry
                    });

                    if (matchesSearch && matchesRisk && matchesType && matchesExpiry) {
                        card.style.display = 'flex';
                        visibleCount++;
                    } else {
                        card.style.display = 'none';
                    }
                });

                filterStats.textContent = 'Showing ' + visibleCount + ' of ' + appCards.length + ' applications';
            }

            // Add event listeners
            appNameFilter.addEventListener('input', filterCards);
            riskScoreFilter.addEventListener('change', filterCards);
            appTypeFilter.addEventListener('change', filterCards);
            expiryStatusFilter.addEventListener('change', filterCards);

            // Initial filter
            console.log('Initial filter running...');
            filterCards();
        });
    </script>
</body>
</html>
"@
{{ ... }}

# --------------------------------------------
# 9. Save and Open HTML Report
# --------------------------------------------

$reportPath = Join-Path $PWD.Path "ApplicationsEnterpriseAppsReport.html"
$script:htmlContent | Out-File -FilePath $reportPath -Encoding UTF8

Write-Host "✨ Report generation completed!" -ForegroundColor Green
Write-Host "📊 Total applications processed: $($csvData.Count)" -ForegroundColor Cyan
Write-Host "📝 Report saved as: $reportPath" -ForegroundColor Yellow

# Open the report in the default browser
try {
    Write-Host "🌐 Opening report in your default browser..." -ForegroundColor Cyan
    Start-Process $reportPath
    Write-Host "✅ Report opened successfully!" -ForegroundColor Green
} catch {
    Write-Host "⚠️ Could not automatically open the report. Please open it manually from: $reportPath" -ForegroundColor Yellow
}

Write-Host "`n💡 Tip: Use the filters at the top of the report to analyze your applications!" -ForegroundColor Magenta

# --------------------------------------------
# End of Script
# --------------------------------------------
