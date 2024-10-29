
# Hawk-EntraID

This PowerShell script uses the Microsoft Graph PowerShell SDK to retrieve information about custom applications in Entra ID and their associated API permissions. The script generates an HTML report, which can be viewed in a browser, and includes an option to export the data as a CSV file.

## Author

**Victor Staflin**  

## Description

The script retrieves details of custom applications in Entra ID, focusing on API permissions, such as application and delegated permissions, and generates an HTML report with the option to export the data to CSV.

### Features:
- Lists custom applications excluding those published by Microsoft or Azure.
- Displays application details (name, ID, creation date, API permissions, credential expiration, etc.).
- Includes options to export the report data as a CSV file.
- Generates a user-friendly HTML report with categorized API permissions.

## Prerequisites

Before running the script, ensure you have the following:

1. **Microsoft Graph PowerShell SDK**  
   Install it by running the following command:
   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser
   ```
2. **Powershell v7**
   
4. **Azure AD Permissions**  
   The following permissions are required in Azure AD to retrieve application information:
   - `Application.Read.All`
   - `Directory.Read.All`

5. **Internet Connection**  
   The script communicates with Microsoft Graph, so an active internet connection is required.

## Risk Scoring Methodology
The application assigns each app a risk score from 1 to 10 based on the following criteria:

Permission Sensitivity:

Application Permissions (direct access without user context) are inherently more sensitive than Delegated Permissions (access granted through a signed-in user), so application permissions contribute more heavily to the risk score.
Permission Type and Weight: Permissions are classified into Critical, High-Risk, Medium-Risk, and Low-Risk categories. Each category has a weighted score based on its potential impact:

Critical Permissions (e.g., Directory.ReadWrite.All): 8 points for application, 6 for delegated.
High-Risk Permissions (e.g., Files.ReadWrite.All): 6 points for application, 4 for delegated.
Medium-Risk Permissions (e.g., Mail.Send): 4 points for application, 3 for delegated.
Low-Risk Permissions (e.g., User.Read): minimal impact on score.
High-Privilege Permission Count: If an application has multiple high-privilege permissions, it will increase the risk score further:

Application Permissions: 1.5x multiplier if two or more high-risk permissions are present.
Delegated Permissions: 1.2x multiplier if two or more high-risk permissions are present.
Additional Risk Factors:

Credential Count: Applications with more than two password credentials increase the score by 1.
Unverified Publisher: Applications without a verified publisher status add 1 to the score.
Read-Only Permissions Cap: The score contribution from read-only permissions is capped to prevent inflation solely due to read access, with a maximum score of 3 points for multiple read-only permissions.

The final score is capped at 10 to simplify prioritization, and applications with scores near or at 10 should be prioritized for security review.



## Usage

Follow these steps to run the script and generate the report:

1. **Open PowerShell** and navigate to the directory where the script is saved.

2. Run the script:
   ```powershell
   .\ScriptName.ps1
   ```

3. **Sign in** with your Azure AD credentials when prompted.

4. The script will generate an **HTML report** and open it in your default web browser.

5. To export the data as a CSV file, **click the "Export to CSV"** button in the HTML report.

## Notes

- The script requires an active internet connection to retrieve information from Microsoft Graph.
- The account used to run the script must have sufficient permissions in Azure AD to access the necessary application details.
- The generated HTML report will be saved in the same directory as the script under the name `CustomApplicationsReport.html`.


## Troubleshooting

- **Permission Issues**: Ensure you have the necessary Azure AD permissions to read the application information. Without proper permissions, the script may fail to retrieve the data.
- **Module Installation**: Verify that the Microsoft Graph SDK and Az module are properly installed by running the following:
   ```powershell
   Get-Module -ListAvailable
   ```
   **Powershell 7**: Make sure to run the script with Powershell 7: https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.4
- **Internet Connection**: Ensure that your internet connection is active, as the script communicates with the Microsoft Graph API.

![image](https://github.com/user-attachments/assets/0066becd-d6de-4e67-88e2-2dab1ccd8424)

