
# Hawk-EntraID

This PowerShell script uses the Microsoft Graph PowerShell SDK to retrieve information about custom applications in Entra ID and their associated API permissions. The script generates an HTML report, which can be viewed in a browser, and includes an option to export the data as a CSV file.

## Author

**Victor Staflin**  
@TRUESEC

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
- **Internet Connection**: Ensure that your internet connection is active, as the script communicates with the Microsoft Graph API.

Image:![image](https://github.com/user-attachments/assets/7632f25a-7512-4041-9eb5-9ee50c1cb3b6)
