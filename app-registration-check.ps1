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

Import-Module Microsoft.Graph

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Application.Read.All", "Directory.Read.All"

# Function to get app permissions
function Get-AppPermissions($appId) {
    $appRoles = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $appId
    $oauth2Permissions = Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $appId

    $permissions = @{
        AppRoles = @()
        DelegatedPermissions = @()
    }
    
    foreach ($role in $appRoles) {
        $resourceApp = $servicePrincipals | Where-Object { $_.AppId -eq $role.ResourceAppId }
        $roleName = ($resourceApp.AppRoles | Where-Object { $_.Id -eq $role.AppRoleId }).Value
        $permissions.AppRoles += "$($resourceApp.DisplayName) - $roleName"
    }

    foreach ($permission in $oauth2Permissions) {
        $resourceApp = $servicePrincipals | Where-Object { $_.AppId -eq $permission.ResourceId }
        $permissions.DelegatedPermissions += "$($resourceApp.DisplayName) - $($permission.Scope)"
    }

    return $permissions
}

# Function to format permissions
function Format-Permissions($permissionList) {
    $filteredPermissions = $permissionList | Where-Object { $_ -ne '-' -and $_ -ne '' -and $_ -ne $null }
    if ($filteredPermissions.Count -eq 0) {
        return "<span class='permission-item none-permission'>None</span>"
    }
    return $filteredPermissions | ForEach-Object { 
        $permissions = $_ -split ' '
        $permissions | Where-Object { $_ -ne '-' -and $_ -ne '' -and $_ -ne $null } | ForEach-Object { 
            $class = if ($_ -match 'Write|ReadWrite|Mail\.Send\.Shared') { 'permission-item write-permission' } else { 'permission-item' }
            "<span class='$class'>$_</span>"
        }
    } | Join-String -Separator " "
}

# Function to format date
function Format-Date($date) {
    if ($date) {
        return $date.ToString("yyyy-MM-dd HH:mm:ss")
    }
    return "N/A"
}

# Function to get the latest credential expiration date
function Get-LatestCredentialExpiration($credentials) {
    if ($credentials -and $credentials.Count -gt 0) {
        $latestCredential = $credentials | Sort-Object EndDateTime -Descending | Select-Object -First 1
        return Format-Date $latestCredential.EndDateTime
    }
    return "N/A"
}

# Get all applications
$allApps = Get-MgApplication -All

# Get all service principals
$servicePrincipals = Get-MgServicePrincipal -All

# Filter custom applications
$customApps = $allApps | Where-Object { 
    $_.Tags -notcontains "WindowsAzureActiveDirectoryIntegratedApp" -and
    $_.PublisherDomain -ne "microsoft.com" -and
    $_.DisplayName -notmatch "^Microsoft" -and
    $_.DisplayName -notmatch "^Office" -and
    $_.DisplayName -notmatch "^Azure"
}

# Create HTML content
$htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Custom Applications Report</title>
    <style>
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            line-height: 1.6; 
            color: #333; 
            max-width: 1200px; 
            margin: 0 auto; 
            padding: 20px; 
            background-color: #f5f5f5;
        }
        h1 { 
            color: #2c3e50; 
            text-align: center;
            margin-bottom: 30px;
        }
        .app { 
            background-color: #ffffff; 
            border: 1px solid #e0e0e0; 
            border-radius: 8px; 
            padding: 20px; 
            margin-bottom: 30px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        .app-name { 
            color: #3498db; 
            font-size: 1.4em; 
            font-weight: bold;
            margin-bottom: 10px;
        }
        .app-id, .object-id { 
            color: #7f8c8d; 
            font-size: 0.9em;
            margin-bottom: 5px;
        }
        .app-dates, .app-details {
            color: #7f8c8d;
            font-size: 0.9em;
            margin-bottom: 10px;
        }
        .app-details div {
            margin-bottom: 5px;
        }
        .permissions-title { 
            color: #2c3e50; 
            font-weight: bold; 
            margin-top: 15px;
            font-size: 1.1em;
        }
        .permission-type { 
            font-weight: bold; 
            margin-top: 10px;
            color: #34495e;
        }
        .permission-list { 
            display: flex; 
            flex-wrap: wrap; 
            gap: 8px; 
            margin: 10px 0;
            min-height: 28px; /* Ensures consistent height even when empty */
        }
        .permission-item { 
            background-color: #e8e8e8; 
            border-radius: 4px; 
            padding: 4px 8px; 
            font-size: 0.9em;
        }
        .write-permission {
            background-color: #ffcccc;
            color: #cc0000;
            font-weight: bold;
        }
        .none-permission {
            color: #999;
            font-style: italic;
        }
        #exportBtn { 
            background-color: #2ecc71; 
            color: white; 
            border: none; 
            padding: 10px 20px; 
            border-radius: 5px; 
            cursor: pointer; 
            font-size: 1em;
            margin-bottom: 20px;
            transition: background-color 0.3s;
        }
        #exportBtn:hover { 
            background-color: #27ae60; 
        }
    </style>
</head>
<body>
    <h1>Custom Applications Report</h1>
    <button id="exportBtn" onclick="exportToCSV()">Export to CSV</button>
"@

$csvData = @()

# Process custom applications
foreach ($app in $customApps) {
    $sp = $servicePrincipals | Where-Object { $_.AppId -eq $app.AppId }
    if ($sp) {
        $permissions = Get-AppPermissions $sp.Id
        
        # Filter out empty permissions, dashes, and null values
        $appRoles = @($permissions.AppRoles | Where-Object { $_ -ne '-' -and $_ -ne '' -and $_ -ne $null })
        $delegatedPermissions = @($permissions.DelegatedPermissions | Where-Object { $_ -ne '-' -and $_ -ne '' -and $_ -ne $null })

        $createdDateTime = Format-Date $app.CreatedDateTime
        $latestCredentialExpiration = Get-LatestCredentialExpiration $app.PasswordCredentials

        $htmlContent += @"
    <div class="app">
        <div class="app-name">$($app.DisplayName)</div>
        <div class="app-id">Application ID: $($app.AppId)</div>
        <div class="object-id">Object ID: $($app.Id)</div>
        <div class="app-dates">Created: $createdDateTime | Latest Credential Expiration: $latestCredentialExpiration</div>
        <div class="app-details">
            <div>Sign-in Audience: $($app.SignInAudience)</div>
            <div>Verified Publisher: $($app.VerifiedPublisher.DisplayName ?? "N/A")</div>
            <div>Password Credentials Count: $($app.PasswordCredentials.Count)</div>
        </div>
        <div class="permissions-title">API Permissions:</div>
        <div class="permission-type">Application Permissions:</div>
        <div class="permission-list">
            $(Format-Permissions $appRoles)
        </div>
        <div class="permission-type">Delegated Permissions:</div>
        <div class="permission-list">
            $(Format-Permissions $delegatedPermissions)
        </div>
    </div>
"@

        $csvRow = [PSCustomObject]@{
            DisplayName = $app.DisplayName
            ApplicationID = $app.AppId
            ObjectID = $app.Id
            Created = $createdDateTime
            LatestCredentialExpiration = $latestCredentialExpiration
            SignInAudience = $app.SignInAudience
            VerifiedPublisher = $app.VerifiedPublisher.DisplayName ?? "N/A"
            PasswordCredentialsCount = $app.PasswordCredentials.Count
            ApplicationPermissions = if ($appRoles.Count -gt 0) { ($appRoles -join "; ") } else { "None" }
            DelegatedPermissions = if ($delegatedPermissions.Count -gt 0) { ($delegatedPermissions -join "; ") } else { "None" }
        }
        $csvData += $csvRow
    }
}

# Convert CSV data to a JSON string
$csvDataJson = $csvData | ConvertTo-Json -Compress

# Update the HTML content to include the CSV data and export function
$htmlContent += @"
    <script>
        const csvData = $csvDataJson;

        function exportToCSV() {
            const headers = Object.keys(csvData[0]);
            const csvContent = [
                headers.join(','),
                ...csvData.map(row => headers.map(fieldName => JSON.stringify(row[fieldName])).join(','))
            ].join('\n');

            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            if (link.download !== undefined) {
                const url = URL.createObjectURL(blob);
                link.setAttribute('href', url);
                link.setAttribute('download', 'CustomApplicationsReport.csv');
                link.style.visibility = 'hidden';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            }
        }
    </script>
</body>
</html>
"@

# Save HTML content to file
$reportPath = Join-Path $PWD.Path "CustomApplicationsReport.html"
$htmlContent | Out-File -FilePath $reportPath -Encoding UTF8

Write-Host "Report has been generated and saved as CustomApplicationsReport.html"

# Open the report in the default browser
Start-Process $reportPath

# Disconnect from Microsoft Graph
Disconnect-MgGraph