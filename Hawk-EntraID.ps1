<#
AUTHOR
Victor Staflin

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
     # List of built-in Microsoft apps and patterns
    $builtInApps = @(
        # Regex patterns for standard Microsoft apps
        '^Microsoft',
        '^MS\s',
        '^Office\s',
        '^Azure\s',
        '^Windows\s',
        '^SharePoint\s',
        '^Teams\s',
        '^Dynamics\s',
        '^Power\s',
        '^Skype\s',
        '^MIP\s',
        '^OCaaS\s',
        '^Intune\s',
        '^Viva\s',
        '^Substrate',
        
        # Exact matches for non-standard names
        'AAD App Management',
        'IPSubstrate',
        'Box',
        'PushChannel',
        'Bing',
        'OMSAuthorizationServicePROD',
        'OfficeServicesManager',
        'WindowsDefenderATP',
        'My Apps',
        'MCAPI Authorization Prod',
        'Exchange Rbac',
        'ComplianceAuthServer',
        'CAP Neptune Prod CM Prod',
        'ProjectWorkManagement',
        'Dataverse',
        'DeploymentScheduler',
        'IAMTenantCrawler',
        'IpLicensingService',
        'ProvisioningHealth',
        'Connectors',
        'EXO_App2025',
        'Portfolios',
        'Sway',
        'ComplianceWorkbenchApp',
        'AADReporting',
        'Signup',
        'SPAuthEvent',
        'Linkedin'
    )

    # Check if the app name matches any pattern or exact name
    if ($builtInApps | Where-Object { $app.DisplayName -match $_ }) {
        return $true
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

function Format-Permission($permission) {
    # Determine if this is a ReadWrite permission
    $isReadWrite = $permission -match 'ReadWrite'
    
    # Set the class based on permission type
    $permissionClass = if ($isReadWrite) {
        'permission-text readwrite'
    } else {
        'permission-text'
    }

    # Determine icon based on permission type
    $icon = if ($isReadWrite) {
        '<i class="fas fa-pen"></i>' # Pen icon for write permissions
    } else {
        '<i class="fas fa-eye"></i>' # Eye icon for read-only
    }

    "<div class='permission-item'>
        $icon
        <div class='$permissionClass'>
            <span class='graph-prefix'>Microsoft Graph - </span>$permission
        </div>
    </div>"
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
                ExpiryStatus = if ($null -eq $daysUntilExpiry) { "High Risk" } elseif ($daysUntilExpiry -lt 0) { "Critical" } elseif ($daysUntilExpiry -le 30) { "Warning" } else { "Valid" }
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
        /* Reset and Base Styles */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            background: linear-gradient(145deg, #0f172a, #1e293b);
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            color: #f1f5f9;
            min-height: 100vh;
            padding: 24px;
        }

        /* Container */
        .container {
            max-width: 100%;
            margin: 0 auto;
            padding: 16px;
        }

        /* Header Section */
        h1 {
            font-size: 24px;
            font-weight: 600;
            color: #f1f5f9;
            margin-bottom: 20px;
        }

        /* Filter Controls */
        .filter-controls {
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
            margin-bottom: 24px;
            align-items: center;
        }

        /* Filter Inputs - Dark Mode Optimized */
        input[type="text"],
        select {
            padding: 8px 12px;
            background: rgba(30, 41, 59, 0.8);
            border: 1px solid rgba(148, 163, 184, 0.2);
            border-radius: 6px;
            color: #e2e8f0;
            font-size: 14px;
            min-width: 200px;
            max-width: 300px;
            transition: all 0.2s ease;
        }

        input[type="text"]:focus,
        select:focus {
            border-color: #3b82f6;
            outline: none;
            box-shadow: 0 0 0 1px rgba(59, 130, 246, 0.2);
        }

        /* Dark mode select options */
        select option {
            background: #1e293b;
            color: #e2e8f0;
        }

        /* Export Button */
        .export-button {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 8px 16px;
            background: #3b82f6;
            color: #ffffff;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s ease;
            margin-bottom: 18px;
        }

        .export-button:hover {
            background: #2563eb;
        }

        .export-button i {
            font-size: 14px;
        }

        /* Filter Stats */
        .filter-stats {
            font-size: 14px;
            color: #94a3b8;
            margin-bottom: 16px;
        }

        /* Apps Grid - Responsive Layout */
        .apps-grid {
            display: grid;
            gap: 16px;
            grid-template-columns: repeat(auto-fill, minmax(min(100%, 400px), 1fr));
            width: 100%;
        }

        /* Responsive Design */
        @media (max-width: 1400px) {
            .apps-grid {
                grid-template-columns: repeat(auto-fill, minmax(min(100%, 350px), 1fr));
            }
        }

        @media (max-width: 1024px) {
            .apps-grid {
                grid-template-columns: repeat(auto-fill, minmax(min(100%, 300px), 1fr));
            }
        }

        @media (max-width: 768px) {
            .filter-controls {
                flex-direction: column;
                align-items: stretch;
            }

            input[type="text"],
            select {
                min-width: 100%;
                max-width: 100%;
            }

            .apps-grid {
                grid-template-columns: 1fr;
            }
        }

        /* Dark Mode Optimization */
        @media (prefers-color-scheme: dark) {
            input[type="text"],
            select {
                background: rgba(255, 255, 255, 0.05);
                border-color: rgba(255, 255, 255, 0.1);
            }
        }

        /* Apps Grid */
        .apps-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
            gap: 20px;
        }

        /* App Card */
        .app-card {
            background: #1e293b;
            border-radius: 8px;
            overflow: hidden;
            display: flex;
            flex-direction: column;
        }

        /* Card Header */
        .app-header {
            display: flex;
            align-items: center;
            padding: 16px;
            gap: 12px;
            position: relative; /* For tooltip positioning */
        }

        /* App Icon */
        .app-icon {
            width: 32px;
            height: 32px;
            flex-shrink: 0;
        }

        /* App Name Container */
        .app-name {
            flex: 1;
            font-size: 14px;
            color: #f1f5f9;
            position: relative;
            cursor: default; /* Shows it's hoverable */
        }

        /* Truncated Text */
        .app-name-truncate {
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            max-width: 200px; /* Adjust based on your layout */
        }

        /* Tooltip on Hover */
        .app-name:hover .app-name-tooltip {
            display: block;
        }

        .app-name-tooltip {
            display: none;
            position: absolute;
            bottom: 100%;
            left: 0;
            background: #1e293b;
            padding: 8px 12px;
            border-radius: 4px;
            font-size: 12px;
            white-space: nowrap;
            z-index: 10;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
            margin-bottom: 8px;
        }

        /* Arrow for tooltip */
        .app-name-tooltip::after {
            content: '';
            position: absolute;
            top: 100%;
            left: 20px;
            border: 6px solid transparent;
            border-top-color: #1e293b;
        }

        /* Badges */
        .risk-score,
        .validity-badge {
            display: inline-flex;
            align-items: center;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: 500;
            gap: 4px;
        }

        .risk-score.low {
            background: rgba(34, 197, 94, 0.2);
            color: #22c55e;
        }

        .risk-score.high {
            background: rgba(239, 68, 68, 0.2);
            color: #ef4444;
        }

        .validity-badge {
            background: rgba(59, 130, 246, 0.2);
            color: #60a5fa;
        }

        /* Card Details */
        .detail-item {
            padding: 12px 16px;
            border-top: 1px solid rgba(255, 255, 255, 0.1);
        }

        .detail-label {
            font-size: 12px;
            color: #94a3b8;
            margin-bottom: 4px;
        }

        .detail-value {
            font-size: 14px;
            word-break: break-all;
        }

        /* Responsive Design */
        @media (max-width: 768px) {
            body {
                padding: 16px;
            }

            .apps-grid {
                grid-template-columns: 1fr;
            }

            .badge-container {
                flex-direction: column;
            }
        }

        /* Permission Tags Layout */
        .permission-tag {
            display: inline-flex;
            align-items: center;
            padding: 4px 8px;
            margin: 2px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: 500;
            gap: 4px;
        }

        /* Read/Write Permission Colors */
        .permission-tag.read {
            background: rgba(59, 130, 246, 0.15);
            color: #60a5fa;
        }

        .permission-tag.write {
            background: rgba(234, 179, 8, 0.15);
            color: #eab308;
        }

        .permission-tag.critical {
            background: rgba(239, 68, 68, 0.15);
            color: #ef4444;
        }

        .permission-tag.none-permission {
            background: rgba(148, 163, 184, 0.15);
            color: #94a3b8;
        }

        /* Risk Score Colors */
        .risk-score {
            display: inline-flex;
            align-items: center;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: 500;
            gap: 4px;
        }

        .risk-score.critical {
            background: rgba(239, 68, 68, 0.15);
            color: #ef4444;
        }

        .risk-score.high {
            background: rgba(234, 179, 8, 0.15);
            color: #eab308;
        }

        .risk-score.medium {
            background: rgba(59, 130, 246, 0.15);
            color: #60a5fa;
        }

        .risk-score.low {
            background: rgba(34, 197, 94, 0.15);
            color: #22c55e;
        }

        /* Permissions Section */
        .permissions-section {
            padding: 16px;
            border-top: 1px solid rgba(255, 255, 255, 0.1);
        }

        .permissions-title {
            font-size: 12px;
            color: #94a3b8;
            margin-bottom: 8px;
            display: flex;
            align-items: center;
            gap: 4px;
        }

        .permissions-list {
            display: flex;
            flex-wrap: wrap;
            gap: 4px;
            margin-top: 8px;
        }

        /* Icon styles for permissions */
        .permission-tag i {
            font-size: 10px;
        }

        .permission-tag.write i,
        .permission-tag.critical i {
            color: currentColor;
        }

        /* Permission Item Styling */
        .permission-item {
            display: flex;
            align-items: flex-start;
            gap: 12px;
            padding: 8px 12px;
            background: rgba(15, 23, 42, 0.6);
            border-radius: 4px;
            margin-bottom: 4px;
        }

        .permission-item i {
            color: #94a3b8;
            font-size: 14px;
            margin-top: 2px;
        }

        .permission-text {
            color: #e2e8f0;
            font-size: 13px;
            line-height: 1.5;
        }

        .permission-text.readwrite {
            color: #ef4444; /* Bright red for ReadWrite permissions */
        }

        .graph-prefix {
            color: #94a3b8 !important; /* Always gray, even in ReadWrite permissions */
        }

        /* Permission Highlights */
        .permission-highlight {
            color: #ef4444; /* Red for high-privilege permissions */
            font-weight: 500;
        }

        /* Section Headers */
        .permissions-header {
            color: #94a3b8;
            font-size: 13px;
            font-weight: 500;
            margin: 16px 0 8px;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .permissions-header i {
            font-size: 14px;
        }

        /* Permission Categories */
        .application-permissions,
        .delegated-permissions {
            margin-bottom: 16px;
        }

        /* App Icon Variations based on type */
        .app-icon i {
            font-size: 18px;
            color: #ffffff;
        }

        /* Different icons for different app types */
        .app-icon.enterprise i {
            color: #3b82f6; /* Blue for enterprise apps */
        }

        .app-icon.service-principal i {
            color: #8b5cf6; /* Purple for service principals */
        }

        .app-icon.managed i {
            color: #10b981; /* Green for managed apps */
        }

        /* Sort Filter Styling */
        #sortFilter {
            min-width: 200px;
            padding: 8px 12px;
            background: rgba(30, 41, 59, 0.8);
            border: 1px solid rgba(148, 163, 184, 0.2);
            border-radius: 6px;
            color: #e2e8f0;
            font-size: 14px;
            cursor: pointer;
        }

        #sortFilter:hover {
            border-color: #3b82f6;
        }

        #sortFilter option {
            background: #1e293b;
            color: #e2e8f0;
            padding: 8px;
        }

        /* Info Button */
        .info-button {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 8px 16px;
            background: #3b82f6;
            color: #ffffff;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s ease;
            margin-bottom: 18px;
        }

        .info-button:hover {
            background: #2563eb;
        }

        .info-button i {
            font-size: 14px;
        }

        /* Info Panel */
        .info-panel {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.5);
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 1000;
            display: none;
        }

        .info-panel.active {
            display: flex;
        }

        .info-content {
            background: #1e293b;
            border-radius: 8px;
            padding: 24px;
            max-width: 80%;
            max-height: 80%;
            overflow: auto;
            position: relative;
        }

        .info-section {
            margin-bottom: 24px;
        }

        .info-section h3 {
            font-size: 16px;
            font-weight: 600;
            color: #f1f5f9;
            margin-bottom: 8px;
        }

        .info-section ul {
            list-style-type: disc;
            padding-left: 20px;
        }

        .info-section li {
            margin-bottom: 4px;
        }

        .permissions-info {
            display: flex;
            justify-content: space-between;
        }

        .permissions-info h4 {
            font-size: 14px;
            font-weight: 600;
            color: #f1f5f9;
            margin-bottom: 8px;
        }

        .permissions-info ul {
            list-style-type: disc;
            padding-left: 20px;
        }

        .permissions-info li {
            margin-bottom: 4px;
        }

        /* Updated Navbar Styles */
        .navbar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 16px 0;
            margin-bottom: 24px;
        }

        .navbar-buttons {
            display: flex;
            gap: 12px;
        }

        /* Info Button */
        .info-button {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 8px 16px;
            background: #475569;
            color: #ffffff;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s ease;
        }

        .info-button:hover {
            background: #334155;
        }
        .detail-item .fa-layer-group {
            color: #ffffff;
        }
        /* Info Panel */
        .info-panel {
            display: none;
            position: fixed;
            top: 0;
            right: 0;
            width: 100%;
            max-width: 600px;
            height: 100vh;
            background: #1e293b;
            box-shadow: -4px 0 12px rgba(0, 0, 0, 0.1);
            z-index: 1000;
            overflow-y: auto;
        }

        .info-panel.active {
            display: block;
        }

        .info-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 24px;
        }

        .close-button {
            background: none;
            border: none;
            color: #94a3b8;
            cursor: pointer;
            font-size: 20px;
            padding: 4px;
        }

        .close-button:hover {
            color: #f1f5f9;
        }

        /* Rest of your existing styles ... */

        /* Modal Styles */
        .modal {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0, 0, 0, 0.7);
            z-index: 1000;
            backdrop-filter: blur(4px);
        }

        .modal-content {
            position: relative;
            background: #1e293b;
            margin: 2% auto;
            width: 90%;
            max-width: 1000px;
            max-height: 90vh;
            border-radius: 12px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
            overflow: hidden;
        }

        .modal-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 20px;
            background: #0f172a;
            border-bottom: 1px solid #334155;
        }

        .modal-header h2 {
            color: #f1f5f9;
            font-size: 1.5rem;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .modal-body {
            padding: 24px;
            overflow-y: auto;
            max-height: calc(90vh - 80px);
        }

        .close-button {
            background: none;
            border: none;
            color: #94a3b8;
            cursor: pointer;
            font-size: 1.5rem;
            padding: 4px;
            transition: color 0.2s;
        }

        .close-button:hover {
            color: #f1f5f9;
        }

        /* Info Section Styles */
        .info-section {
            margin-bottom: 32px;
        }

        .info-section h3 {
            color: #f1f5f9;
            font-size: 1.25rem;
            margin-bottom: 16px;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .info-section ul {
            list-style: none;
            padding-left: 8px;
        }

        .info-section ul li {
            margin-bottom: 12px;
            color: #cbd5e1;
            line-height: 1.6;
        }

        .info-section ul li strong {
            color: #f1f5f9;
        }

        .info-section ul ul {
            margin-top: 8px;
            margin-left: 20px;
            border-left: 2px solid #334155;
            padding-left: 16px;
        }

        /* Risk Categories */
        .risk-category {
            margin-bottom: 24px;
            padding: 16px;
            border-radius: 8px;
            border: 1px solid rgba(255, 255, 255, 0.1);
        }

        .risk-category h4 {
            display: flex;
            align-items: center;
            gap: 8px;
            margin-bottom: 12px;
            font-size: 1.1rem;
        }

        .risk-category.critical {
            background: rgba(239, 68, 68, 0.1);
        }

        .risk-category.high {
            background: rgba(245, 158, 11, 0.1);
        }

        .risk-category.medium {
            background: rgba(59, 130, 246, 0.1);
        }

        .risk-category.low {
            background: rgba(16, 185, 129, 0.1);
        }

        .risk-category.critical h4 { color: #ef4444; }
        .risk-category.high h4 { color: #f59e0b; }
        .risk-category.medium h4 { color: #3b82f6; }
        .risk-category.low h4 { color: #10b981; }

        /* Methodology Styles */
        .methodology-content {
            color: #cbd5e1;
            line-height: 1.6;
        }

        .methodology-category {
            margin-bottom: 24px;
            padding: 16px;
            background: rgba(255, 255, 255, 0.05);
            border-radius: 8px;
        }

        .methodology-category h4 {
            color: #f1f5f9;
            margin-bottom: 12px;
            font-size: 1.1rem;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .scoring-table {
            width: 100%;
            border-collapse: collapse;
            margin: 16px 0;
            font-size: 0.9rem;
        }

        .scoring-table th,
        .scoring-table td {
            padding: 12px;
            border: 1px solid #334155;
            text-align: left;
        }

        .scoring-table th {
            background: #1e293b;
            color: #f1f5f9;
        }

        .scoring-table tr.critical td { background: rgba(239, 68, 68, 0.1); }
        .scoring-table tr.high td { background: rgba(245, 158, 11, 0.1); }
        .scoring-table tr.medium td { background: rgba(59, 130, 246, 0.1); }
        .scoring-table tr.low td { background: rgba(16, 185, 129, 0.1); }

        .methodology-note {
            display: flex;
            align-items: flex-start;
            gap: 12px;
            padding: 16px;
            background: rgba(59, 130, 246, 0.1);
            border-radius: 8px;
            margin-top: 24px;
        }

        .methodology-note i {
            color: #3b82f6;
            font-size: 1.2rem;
            margin-top: 3px;
        }

        .methodology-note p {
            margin: 0;
            color: #f1f5f9;
        }

        /* Validity Badges */
        .validity-badge {
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: 500;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }

        .validity-badge.high-risk {
            background: rgba(239, 68, 68, 0.15);
            color: #ef4444;
            border: 1px solid rgba(239, 68, 68, 0.3);
        }

        .validity-badge.critical {
            background: rgba(0, 0, 0, 0.2);
            color: #ef4444;
            border: 1px solid rgba(239, 68, 68, 0.5);
            animation: pulse 2s infinite;
        }

        .validity-badge.warning {
            background: rgba(245, 158, 11, 0.15);
            color: #f59e0b;
            border: 1px solid rgba(245, 158, 11, 0.3);
        }

        .validity-badge.valid {
            background: rgba(16, 185, 129, 0.15);
            color: #10b981;
            border: 1px solid rgba(16, 185, 129, 0.3);
        }

        @keyframes pulse {
            0% { opacity: 1; }
            50% { opacity: 0.5; }
            100% { opacity: 1; }
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- Updated Navbar Section -->
        <div class="navbar">
            <h1>Enterprise Applications Security Report</h1>
            <div class="navbar-buttons">
                <button class="info-button" onclick="toggleInfo()">
                    <i class="fas fa-info-circle"></i>
                    Documentation
                </button>
                <button class="export-button" onclick="exportToCSV()">
                    <i class="fas fa-file-export"></i>
                    Export to CSV
                </button>
            </div>
        </div>

        <!-- Updated Info Modal -->
        <div id="infoModal" class="modal">
            <div class="modal-content">
                <div class="modal-header">
                    <h2><i class="fas fa-info-circle"></i> Documentation</h2>
                    <button class="close-button" onclick="toggleInfo()">
                        <i class="fas fa-times"></i>
                    </button>
                </div>
                <div class="modal-body">
                    <div class="info-section">
                        <h3>🔍 Field Descriptions</h3>
                        <ul>
                            <li><strong>Application Name:</strong> The display name of the application in Azure AD.</li>
                            <li><strong>Risk Score:</strong> A calculated value (1-10) based on:
                                <ul>
                                    <li>Number and type of permissions granted</li>
                                    <li>Credential expiration status</li>
                                    <li>Publisher verification status</li>
                                    <li>Number of credentials</li>
                                </ul>
                            </li>
                            <li><strong>Application Type:</strong> Internal (single-tenant) or External (multi-tenant) application.</li>
                            <li><strong>Application/Object ID:</strong> Unique identifiers in Azure AD.</li>
                            <li><strong>Sign-in Audience:</strong> Defines who can use the application (My org only, All Microsoft accounts, etc.).</li>
                            <li><strong>Verified Publisher:</strong> Whether Microsoft has verified the application publisher.</li>
                            <li><strong>Password Credentials:</strong> Number of client secrets/certificates configured.</li>
                        </ul>
                    </div>
                    
                    <div class="info-section">
                        <h3>📊 Risk Scoring Methodology</h3>
                        <div class="methodology-content">
                            <p>Applications are assigned a risk score from 1 to 10 based on multiple security factors:</p>
                            
                            <div class="methodology-category">
                                <h4>Permission Sensitivity</h4>
                                <ul>
                                    <li>Application Permissions (direct access) have higher weight than Delegated Permissions (user context)</li>
                                    <li>Permissions are categorized and weighted as follows:</li>
                                    <table class="scoring-table">
                                        <tr>
                                            <th>Category</th>
                                            <th>Application Weight</th>
                                            <th>Delegated Weight</th>
                                            <th>Example</th>
                                        </tr>
                                        <tr class="critical">
                                            <td>Critical</td>
                                            <td>8 points</td>
                                            <td>6 points</td>
                                            <td>Directory.ReadWrite.All</td>
                                        </tr>
                                        <tr class="high">
                                            <td>High</td>
                                            <td>6 points</td>
                                            <td>4 points</td>
                                            <td>Files.ReadWrite.All</td>
                                        </tr>
                                        <tr class="medium">
                                            <td>Medium</td>
                                            <td>4 points</td>
                                            <td>3 points</td>
                                            <td>Mail.Send</td>
                                        </tr>
                                        <tr class="low">
                                            <td>Low</td>
                                            <td>2 points</td>
                                            <td>1 point</td>
                                            <td>User.Read</td>
                                        </tr>
                                    </table>
                                </ul>
                            </div>

                            <div class="methodology-category">
                                <h4>Risk Multipliers</h4>
                                <ul>
                                    <li><strong>Multiple High-Privilege Permissions:</strong>
                                        <ul>
                                            <li>Application Permissions: 1.5x multiplier for 2+ high-risk permissions</li>
                                            <li>Delegated Permissions: 1.2x multiplier for 2+ high-risk permissions</li>
                                        </ul>
                                    </li>
                                </ul>
                            </div>

                            <div class="methodology-category">
                                <h4>Additional Risk Factors</h4>
                                <ul>
                                    <li><strong>Credential Count:</strong> +1 point if more than two password credentials</li>
                                    <li><strong>Publisher Status:</strong> +1 point if publisher is not verified</li>
                                    <li><strong>Read-Only Cap:</strong> Maximum of 3 points from read-only permissions</li>
                                </ul>
                            </div>

                            <div class="methodology-note">
                                <i class="fas fa-info-circle"></i>
                                <p>The final score is capped at 10. Applications scoring 8 or higher should be prioritized for security review.</p>
                            </div>
                        </div>
                    </div>
                    
                    <div class="info-section">
                        <h3>🔑 API Permissions Guide</h3>
                        <div class="permissions-info">
                            <div class="risk-category critical">
                                <h4><i class="fas fa-radiation-alt"></i> Critical Risk Permissions</h4>
                                <ul>
                                    <li><strong>Directory.ReadWrite.All:</strong> Full access to read and write directory data
                                        <ul>
                                            <li>Create/delete users and groups</li>
                                            <li>Assign roles and manage licenses</li>
                                            <li>Update organization settings</li>
                                        </ul>
                                    </li>
                                    <li><strong>RoleManagement.ReadWrite.Directory:</strong> Manage role assignments and role definitions</li>
                                    <li><strong>Application.ReadWrite.All:</strong> Full access to manage applications and service principals</li>
                                    <li><strong>AppRoleAssignment.ReadWrite.All:</strong> Manage application role assignments</li>
                                </ul>
                            </div>

                            <div class="risk-category high">
                                <h4><i class="fas fa-exclamation-triangle"></i> High Risk Permissions</h4>
                                <ul>
                                    <li><strong>Mail.ReadWrite.All:</strong> Read and write all user mail</li>
                                    <li><strong>Files.ReadWrite.All:</strong> Full access to all SharePoint files</li>
                                    <li><strong>User.ReadWrite.All:</strong> Read and write all user profiles</li>
                                    <li><strong>Group.ReadWrite.All:</strong> Read and write all groups</li>
                                </ul>
                            </div>

                            <div class="risk-category medium">
                                <h4><i class="fas fa-shield-alt"></i> Medium Risk Permissions</h4>
                                <ul>
                                    <li><strong>Mail.Send:</strong> Send mail as any user</li>
                                    <li><strong>Sites.ReadWrite.All:</strong> Read and write all SharePoint site collections</li>
                                    <li><strong>Calendar.ReadWrite:</strong> Read and write calendar items</li>
                                    <li><strong>Device.ReadWrite.All:</strong> Read and write all device properties</li>
                                </ul>
                            </div>

                            <div class="risk-category low">
                                <h4><i class="fas fa-check-circle"></i> Low Risk Permissions</h4>
                                <ul>
                                    <li><strong>User.Read:</strong> Sign in and read user profile</li>
                                    <li><strong>Directory.Read.All:</strong> Read directory data</li>
                                    <li><strong>Group.Read.All:</strong> Read all groups</li>
                                    <li><strong>Sites.Read.All:</strong> Read items in all site collections</li>
                                </ul>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <div class="filter-controls">
            <input type="text" id="appNameFilter" placeholder="Search Applications">
            <select id="riskScoreFilter">
                <option value="all">All Risk Levels</option>
                <option value="critical">Critical (8-10)</option>
                <option value="high">High (6-7)</option>
                <option value="medium">Medium (4-5)</option>
                <option value="low">Low (0-3)</option>
            </select>
            <select id="appTypeFilter">
                <option value="all">All Types</option>
                <option value="internal">Internal</option>
                <option value="external">External</option>
            </select>
            <select id="expiryStatusFilter">
                <option value="all">All Statuses</option>
                <option value="High Risk">No Expiry (High Risk)</option>
                <option value="Critical">Expired (Critical)</option>
                <option value="Warning">Expiring Soon</option>
                <option value="Valid">Valid</option>
            </select>
            <select id="sortFilter">
                <option value="none">Sort By...</option>
                <option value="riskScore-desc">Highest Risk Score First</option>
                <option value="riskScore-asc">Lowest Risk Score First</option>
                <option value="name-asc">Name (A-Z)</option>
                <option value="name-desc">Name (Z-A)</option>
                <option value="permissions-desc">Most Permissions First</option>
                <option value="created-desc">Newest First</option>
                <option value="created-asc">Oldest First</option>
            </select>
        </div>
        <div class="filter-stats">Showing all applications</div>
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
        "High Risk"  # No expiration set
    } elseif ($row.DaysUntilExpiry -lt 0) { 
        "Critical"   # Already expired
    } elseif ($row.DaysUntilExpiry -le 30) { 
        "Warning"    # Expiring soon
    } else { 
        "Valid"      # Valid credentials
    }

    # Calculate risk score
    $riskScore = [int]$row.RiskScore

    # Update the validity badge generation
    $validityBadgeHtml = switch ($row.ExpiryStatus) {
        "High Risk" { @"
            <div class='validity-badge high-risk'>
                <i class='fas fa-exclamation-triangle'></i>
                No Expiry Set
            </div>
"@ }
        "Critical" { @"
            <div class='validity-badge critical'>
                <i class='fas fa-radiation-alt'></i>
                Expired
            </div>
"@ }
        "Warning" { @"
            <div class='validity-badge warning'>
                <i class='fas fa-exclamation-circle'></i>
                Expires in $($row.DaysUntilExpiry) days
            </div>
"@ }
        "Valid" { @"
            <div class='validity-badge valid'>
                <i class='fas fa-check-circle'></i>
                Valid ($($row.DaysUntilExpiry) days)
            </div>
"@ }
        default { @"
            <div class='validity-badge high-risk'>
                <i class='fas fa-exclamation-triangle'></i>
                No Expiry Set
            </div>
"@ }
    }

    $script:htmlContent += @"
            <div class='app-card' 
                data-apptype='$([System.Web.HttpUtility]::HtmlEncode($row.ApplicationType))'
                data-expirystatus='$expiryStatus'
                data-riskscore='$riskScore'>
                <div class='app-header'>
                    <div class='app-icon'>
                       <i class="fas fa-cube"></i> 
                    </div>
                    <div class='app-name'>$([System.Web.HttpUtility]::HtmlEncode($row.DisplayName))</div>
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
                    $validityBadgeHtml
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
                                    Format-Permission $_.Trim()
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
                                    Format-Permission $_.Trim()
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
        document.addEventListener('DOMContentLoaded', function() {
            // Cache DOM elements
            var appNameFilter = document.getElementById('appNameFilter');
            var riskScoreFilter = document.getElementById('riskScoreFilter');
            var appTypeFilter = document.getElementById('appTypeFilter');
            var expiryStatusFilter = document.getElementById('expiryStatusFilter');
            var sortFilter = document.getElementById('sortFilter');
            var filterStats = document.querySelector('.filter-stats');
            var appCards = document.querySelectorAll('.app-card');
            var infoModal = document.getElementById('infoModal');

            // Info Modal Functions
            window.toggleInfo = function() {
                if (infoModal) {
                    infoModal.style.display = infoModal.style.display === 'block' ? 'none' : 'block';
                }
            }

            // Close modal when clicking outside
            window.onclick = function(event) {
                if (event.target === infoModal) {
                    infoModal.style.display = 'none';
                }
            }

            // Close modal with Escape key
            document.addEventListener('keydown', function(event) {
                if (event.key === 'Escape' && infoModal) {
                    infoModal.style.display = 'none';
                }
            });

            // Rest of your existing filterCards and other functions...
            function filterCards() {
                var searchTerm = appNameFilter.value.toLowerCase();
                var riskLevel = riskScoreFilter.value;
                var appType = appTypeFilter.value;
                var expiryStatus = expiryStatusFilter.value;
                var sortValue = sortFilter.value;

                // Convert NodeList to Array for easier manipulation
                var cards = Array.prototype.slice.call(appCards);

                // Apply filters
                cards = cards.filter(function(card) {
                    var appName = card.querySelector('.app-name').textContent.toLowerCase();
                    var riskScore = parseInt(card.getAttribute('data-riskscore'));
                    var cardType = card.getAttribute('data-apptype');
                    var cardExpiry = card.getAttribute('data-expirystatus');

                    var matchesSearch = appName.includes(searchTerm);
                    var matchesRisk = riskLevel === 'all' || 
                        (riskLevel === 'critical' && riskScore >= 8) ||
                        (riskLevel === 'high' && riskScore >= 6 && riskScore < 8) ||
                        (riskLevel === 'medium' && riskScore >= 4 && riskScore < 6) ||
                        (riskLevel === 'low' && riskScore < 4);
                    var matchesType = appType === 'all' || cardType === appType;
                    var matchesExpiry = expiryStatus === 'all' || cardExpiry === expiryStatus;

                    return matchesSearch && matchesRisk && matchesType && matchesExpiry;
                });

                // Apply sorting
                if (sortValue !== 'none') {
                    cards.sort(function(a, b) {
                        switch (sortValue) {
                            case 'riskScore-desc':
                                return parseInt(b.getAttribute('data-riskscore')) - parseInt(a.getAttribute('data-riskscore'));
                            case 'riskScore-asc':
                                return parseInt(a.getAttribute('data-riskscore')) - parseInt(b.getAttribute('data-riskscore'));
                            case 'name-asc':
                                return a.querySelector('.app-name').textContent.localeCompare(b.querySelector('.app-name').textContent);
                            case 'name-desc':
                                return b.querySelector('.app-name').textContent.localeCompare(a.querySelector('.app-name').textContent);
                            case 'permissions-desc':
                                var getPermCount = function(card) {
                                    var lists = card.querySelectorAll('.permissions-list');
                                    var count = 0;
                                    Array.prototype.forEach.call(lists, function(list) {
                                        count += list.querySelectorAll('.permission-item').length;
                                    });
                                    return count;
                                };
                                return getPermCount(b) - getPermCount(a);
                            case 'created-desc':
                            case 'created-asc':
                                var getDate = function(card) {
                                    var dateText = card.querySelectorAll('.detail-value')[2].textContent;
                                    return new Date(dateText);
                                };
                                var dateA = getDate(a);
                                var dateB = getDate(b);
                                return sortValue === 'created-desc' ? dateB - dateA : dateA - dateB;
                            default:
                                return 0;
                        }
                    });
                }

                // Update the display
                var container = document.querySelector('.apps-grid');
                container.innerHTML = '';
                cards.forEach(function(card) {
                    container.appendChild(card);
                });

                // Update stats
                filterStats.textContent = 'Showing ' + cards.length + ' of ' + appCards.length + ' applications';
            }

            // Export function
            window.exportToCSV = function() {
                var visibleCards = Array.prototype.slice.call(document.querySelectorAll('.app-card')).filter(function(card) {
                    return window.getComputedStyle(card).display !== 'none';
                });

                var headers = [
                    'Application Name',
                    'Application Type',
                    'Application ID',
                    'Object ID',
                    'Created Date',
                    'Risk Score',
                    'Expiry Status',
                    'Sign-in Audience',
                    'Verified Publisher',
                    'Password Credentials Count',
                    'Application Permissions',
                    'Delegated Permissions'
                ];

                var csvContent = headers.join(',') + '\n';

                visibleCards.forEach(function(card) {
                    var row = [
                        card.querySelector('.app-name').textContent.replace(/,/g, ';'),
                        card.getAttribute('data-apptype'),
                        card.querySelector('.detail-value').textContent.replace(/,/g, ';'),
                        card.querySelectorAll('.detail-value')[1].textContent.replace(/,/g, ';'),
                        card.querySelectorAll('.detail-value')[2].textContent.replace(/,/g, ';'),
                        card.getAttribute('data-riskscore'),
                        card.getAttribute('data-expirystatus'),
                        card.querySelectorAll('.detail-value')[4].textContent.replace(/,/g, ';'),
                        card.querySelectorAll('.detail-value')[5].textContent.replace(/,/g, ';'),
                        card.querySelectorAll('.detail-value')[6].textContent.replace(/,/g, ';')
                    ];

                    // Add permissions
                    var appPerms = Array.prototype.slice.call(card.querySelectorAll('.permissions-list')[0].querySelectorAll('.permission-item'))
                        .map(function(p) { return p.textContent.trim(); })
                        .join(';')
                        .replace(/,/g, ';');
                    var delPerms = Array.prototype.slice.call(card.querySelectorAll('.permissions-list')[1].querySelectorAll('.permission-item'))
                        .map(function(p) { return p.textContent.trim(); })
                        .join(';')
                        .replace(/,/g, ';');

                    row.push(appPerms, delPerms);
                    csvContent += row.join(',') + '\n';
                });

                var blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                var link = document.createElement('a');
                var url = URL.createObjectURL(blob);
                
                link.setAttribute('href', url);
                link.setAttribute('download', 'EnterpriseAppsReport.csv');
                link.style.visibility = 'hidden';
                
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            };

            // Add event listeners
            var filters = ['appNameFilter', 'riskScoreFilter', 'appTypeFilter', 'expiryStatusFilter', 'sortFilter'];
            filters.forEach(function(filterId) {
                var element = document.getElementById(filterId);
                if (element) {
                    element.addEventListener('change', filterCards);
                    element.addEventListener('input', filterCards);
                }
            });

            // Initial filter application
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
