# Find-CopilotInteractions-Graph.ps1

## Overview

`Find-CopilotInteractions-Graph.ps1` is an advanced PowerShell tool that retrieves and analyzes Microsoft Copilot interactions for a set of users in your Microsoft 365 tenant, using the Microsoft Graph API. The script is designed for IT administrators and analysts who need actionable Copilot usage insights at scale across their organization.

It supports both Client Secret and Certificate-based authentication, robust error handling, throttling protection, and flexible reporting/export options (CSV or Excel). The script produces interactive and aggregated reports to help you understand Copilot adoption and usage patterns.

---

## Features

- **Bulk Copilot Interactions Fetching**: Collects prompt/response history for all users listed in a simple text file.
- **Advanced Filtering**: Optionally include/exclude AI responses for focused user activity reporting.
- **Automatic vs. User-Initiated Detection**: Flags system-generated (automatic) vs. human-initiated prompts.
- **Throttling Protection**: Implements exponential backoff and delay between requests to respect Graph API rate limits.
- **Flexible Export**: Outputs results as CSV (default) or Excel (if ImportExcel module is installed).
- **Interactive Analysis**: Optionally view results in PowerShell’s Out-GridView.
- **Aggregated Statistics**: Provides per-user and global Copilot activity breakdowns.
- **Date Range Controls**: Analyze any time window with customizable look-back/forward days.

---

## Requirements

- **PowerShell 7.x** or **Windows PowerShell 5.1**
- **Microsoft Graph PowerShell SDK**  
  Install: `Install-Module Microsoft.Graph`
- **ImportExcel PowerShell Module** *(optional, for Excel export)*  
  Install: `Install-Module ImportExcel`
- **Registered Azure AD App** with:
    - `User.Read.All` (Application permission)
    - `AiEnterpriseInteraction.Read.All` (Application permission)
- **Copilot License** assigned to all users being reported on
- **User List File**: Text file with one user principal name (UPN) per line

---

## Permissions

The Azure AD app registration used for authentication must have the following Microsoft Graph **application** permissions (not delegated):

- `User.Read.All`
- `AiEnterpriseInteraction.Read.All`

Refer to the [Microsoft docs on Graph authentication](https://learn.microsoft.com/en-us/graph/auth-v2-service) for app registration and permission assignment.

---

## Setup

### 1. Register an Azure AD Application

- Azure Portal → Azure Active Directory → App registrations → New registration
- Assign `User.Read.All` and `AiEnterpriseInteraction.Read.All` **(application permissions)**
- Generate a **Client Secret** or upload a **Certificate**
- Record your **Application (client) ID** and **Directory (tenant) ID**

### 2. Install Required PowerShell Modules

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser   # Optional, for Excel export
```

### 3. Prepare User List

Create a file such as `C:\temp\simpleuserlist.txt` with one UPN per line:

```
user1@contoso.com
user2@contoso.com
```

---

## Usage

### 1. Configure the Script

Edit the **CONFIGURATION SECTION** at the top of `Find-CopilotInteractions-Graph.ps1`:

```powershell
# Azure AD app settings
$appID = 'your-app-client-id'
$TenantId = 'your-tenant-id'
$AuthType = 'ClientSecret'   # or 'Certificate'

# For Client Secret authentication
$ClientSecret = 'your-client-secret'

# For Certificate authentication
$Thumbprint = "your-cert-thumbprint"

# Path to user list
$UserListPath = "C:\temp\simpleuserlist.txt"

# Reporting options
$IncludeAIResponses = $false      # $true to include AI responses, $false for user prompts only
$ShowGridView = $false            # $true to open Out-GridView
$ExportOption = "CSV"             # 'CSV' or 'XLSX'

# Date range
$DaysToLookBack = 60
$DaysToLookAhead = 1

# API control
$MaxResultsPerRequest = 100
$RequestDelayMs = 250    # Delay (ms) between API calls to prevent throttling
$Debug = $true           # $true to log all requests/responses
$TrackRPS = $true        # $true to track requests-per-second
```

### 2. Run the Script

```powershell
.\Find-CopilotInteractions-Graph.ps1
```

- Script will process each user in your list, showing progress and statistics.
- Results are saved to your **Downloads** folder as CSV or XLSX.
- Summary stats are displayed in the console.

---

## Parameters

| Parameter                | Description                                                                                      |
|--------------------------|--------------------------------------------------------------------------------------------------|
| `appID`                  | Application (client) ID of your Azure AD app registration                                        |
| `TenantId`               | Directory (tenant) ID                                                                            |
| `AuthType`               | Authentication type: `'ClientSecret'` or `'Certificate'`                                         |
| `ClientSecret`           | Client secret (for `'ClientSecret'` auth)                                                        |
| `Thumbprint`             | Certificate thumbprint (for `'Certificate'` auth)                                                |
| `UserListPath`           | Path to user list file (one UPN per line)                                                        |
| `IncludeAIResponses`     | `$true` to include AI responses, `$false` for user prompts only                                  |
| `ShowGridView`           | `$true` to view results in Out-GridView GUI table                                                |
| `ExportOption`           | `'CSV'` or `'XLSX'` (XLSX requires ImportExcel module)                                           |
| `DaysToLookBack`         | Days in the past to include in the report                                                        |
| `DaysToLookAhead`        | Days in the future to include in the report                                                      |
| `MaxResultsPerRequest`   | Maximum results per API call (`$top` parameter on Graph API)                                     |
| `RequestDelayMs`         | Delay (milliseconds) between API requests (to avoid throttling)                                  |
| `Debug`                  | `$true` to enable detailed API logging                                                           |
| `TrackRPS`               | `$true` to track and report API request rates                                                    |

---

## Output

- **CSV or Excel File**: Saved in your Downloads folder as `CopilotInteractions-YYYYMMDD-HHMMSS.csv` or `.xlsx`.
- **Console Summary**: Per-user and global statistics, including breakdowns by Copilot app, interaction type, and auto-generated vs. user-initiated.
- **Interactive Table**: If `ShowGridView` is enabled, opens an Out-GridView window for browsing/filtering results.
- **Debug/Error Log**: If enabled, a log file of all requests/responses or errors for troubleshooting.

**Sample Output Screenshots:**

*CSV Output:*
![CSV Output Example](https://github.com/user-attachments/assets/76b949c0-67b7-4bcb-97a6-c5e41a23279b)

*Excel Output:*
![Excel Output Example](https://github.com/user-attachments/assets/f0646d6b-22f9-44c7-958c-9bb58a42f37e)

---

## Troubleshooting

- **No Data Returned**:  
  - Ensure users have Copilot licenses and active use within the specified date range.

- **Authentication Errors**:  
  - Verify app credentials, permissions, and tenant ID.
  - Check for certificate expiry or secret validity.

- **Permission Denied**:  
  - Confirm that `User.Read.All` and `AiEnterpriseInteraction.Read.All` are assigned and admin consented.

- **Throttling / 429 Errors**:  
  - The script auto-retries with exponential backoff. If repeated, increase `$RequestDelayMs` or reduce user count.

- **Excel Export Fails**:  
  - Install ImportExcel (`Install-Module ImportExcel`), or set `ExportOption = 'CSV'`.

- **User Not Found / License Missing**:  
  - Ensure all UPNs in your user list are valid and licensed for Copilot.

- **General Script Errors**:  
  - Review the log file generated in your Downloads folder for detailed diagnostics.

---

## Example

```powershell
# Example: Analyze 30 days of Copilot usage for a list of users, exporting to Excel
$appID = 'abc64618-...'
$TenantId = '9cfc42cb-...'
$AuthType = 'ClientSecret'
$ClientSecret = 'your-secret'
$UserListPath = "C:\temp\simpleuserlist.txt"
$IncludeAIResponses = $false
$ShowGridView = $true
$ExportOption = "XLSX"
$DaysToLookBack = 30
```

---

## Credits

- **Authors:** Mike Lee, Jay Thakker, Tony Redmond
- **References:**
    - [Microsoft Graph API Throttling](https://learn.microsoft.com/en-us/graph/throttling)
    - [ImportExcel PowerShell Module](https://www.powershellgallery.com/packages/ImportExcel)
    - [Office 365 for IT Pros Scripts](https://github.com/12Knocksinna/Office365itpros)

---

## License

MIT License (see repository for details).

---

## Support

For issues or questions, please open a [GitHub issue](https://github.com/mikelee1313/Copilot/issues) or contact the script authors.

---

*(Replace the Azure AD app IDs, secrets, certificate thumbprints, and user list paths with your own values before running.)*
