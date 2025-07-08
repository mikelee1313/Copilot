Certainly! Below is a detailed README.md tailored for the `Find-CopilotInteractions-Graph.ps1` script in your Copilot repository. This README provides clear setup, usage, parameter, requirements, and troubleshooting instructions, designed for IT admins and advanced users.

---

# Find-CopilotInteractions-Graph.ps1

## Overview

`Find-CopilotInteractions-Graph.ps1` is a PowerShell script that retrieves and analyzes Microsoft Copilot interactions for a list of users using the Microsoft Graph API. It provides insights into how users interact with Copilot across different Microsoft 365 applications, helping organizations understand Copilot adoption and usage patterns.

The script supports both Client Secret and Certificate-based authentication, includes robust error handling and throttling protection, and exports results to CSV or Excel files. It is designed for IT administrators and requires appropriate Graph API permissions.

---

## Features

- **Fetch Copilot Interactions:** Retrieves Copilot user prompt and AI response data from the Graph API.
- **Multi-User Reporting:** Processes a list of users in bulk via a simple text file.
- **Advanced Filtering:** Optionally include/exclude AI responses for focused reporting.
- **Automatic vs. User-Initiated:** Attempts to identify and flag automatic (system-generated) interactions.
- **Throttling Protection:** Implements exponential backoff for Graph API rate limits.
- **Flexible Export:** Exports results to CSV by default or to Excel (if the ImportExcel module is available).
- **Interactive Analysis:** Optionally view results in an Out-GridView window.
- **Aggregated Statistics:** Provides per-user and overall summaries of Copilot activity.
- **Date Range Controls:** Easily specify the analysis time window.

---

## Requirements

- **PowerShell 7.x or Windows PowerShell 5.1**
- **Microsoft Graph PowerShell SDK** (`Install-Module Microsoft.Graph`)
- **ImportExcel PowerShell Module** (`Install-Module ImportExcel`) — *optional, for XLSX export*
- **Registered Azure AD App** with:
  - `User.Read.All` permission
  - `AiEnterpriseInteraction.Read.All` permission
- **Users must have a Copilot license assigned**
- **User List File:** A text file with user principal names (UPNs), one per line

---

## Permissions

Your Azure AD application must have the following Microsoft Graph API permissions (application permissions, not delegated):

- `User.Read.All`
- `AiEnterpriseInteraction.Read.All`

You will authenticate with either a client secret or a certificate. See [Microsoft documentation](https://learn.microsoft.com/en-us/graph/auth-v2-service) for registering apps and assigning permissions.

---

## Setup

### 1. Register an Azure AD Application

- Go to **Azure Portal** → **Azure Active Directory** → **App registrations** → **New registration**
- Grant **User.Read.All** and **AiEnterpriseInteraction.Read.All** (application permissions)
- Generate a **client secret** or upload a **certificate**
- Record the **Application (client) ID** and **Directory (tenant) ID**

### 2. Install Required PowerShell Modules

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser   # Optional, for Excel export
```

### 3. Prepare User List

Create a file, e.g., `C:\temp\simpleuserlist.txt`, with one user principal name (UPN) per line:

```
user1@contoso.com
user2@contoso.com
```

---

## Usage

### 1. Configure the Script

Edit the **CONFIGURATION SECTION** at the top of `Find-CopilotInteractions-Graph.ps1`:

```powershell
# Required Azure AD app settings
$appID = 'your-app-client-id'
$TenantId = 'your-tenant-id'
$AuthType = 'ClientSecret'   # or 'Certificate'

# If using a client secret
$ClientSecret = 'your-client-secret'

# If using a certificate
$Thumbprint = "your-cert-thumbprint"

# User list file
$UserListPath = "C:\temp\simpleuserlist.txt"

# Report options
$IncludeAIResponses = $false      # $true to include AI responses
$ShowGridView = $false            # $true to show Out-GridView
$ExportOption = "CSV"             # 'CSV' or 'XLSX'

# Date range
$DaysToLookBack = 60
$DaysToLookAhead = 1

# API query controls
$MaxResultsPerRequest = 100
```

### 2. Run the Script

```powershell
.\Find-CopilotInteractions-Graph.ps1
```

- The script will prompt and display progress for each user.
- Results are saved to your **Downloads** folder as CSV or XLSX.
- Summary statistics are shown in the console.

---

## Parameters

| Parameter              | Description                                                                                 |
|------------------------|---------------------------------------------------------------------------------------------|
| `appID`                | Application (client) ID of your registered Azure AD app                                     |
| `TenantId`             | Directory (tenant) ID                                                                       |
| `AuthType`             | Authentication type: `'ClientSecret'` or `'Certificate'`                                   |
| `ClientSecret`         | Client secret (if using `'ClientSecret'`)                                                   |
| `Thumbprint`           | Certificate thumbprint (if using `'Certificate'`)                                           |
| `UserListPath`         | Path to user list file (one UPN per line)                                                   |
| `IncludeAIResponses`   | `$true` to include AI responses, `$false` for user prompts only                             |
| `ShowGridView`         | `$true` to show Out-GridView of results                                                     |
| `ExportOption`         | `'CSV'` or `'XLSX'` (requires ImportExcel module for XLSX)                                  |
| `DaysToLookBack`       | Days in the past to include in the report                                                   |
| `DaysToLookAhead`      | Days in the future to include in the report                                                 |
| `MaxResultsPerRequest` | Max results per API call (controls `$top` parameter for Graph API)                         |

---

## Output

- **CSV or Excel File:** Saved to your Downloads folder, e.g. `CopilotInteractions-YYYYMMDD-HHMMSS.csv` or `.xlsx`.
- **Console Summary:** Per-user and global statistics, including automatic vs. user-initiated interactions.
- **Optional Interactive View:** Use `$ShowGridView = $true` to open results in a GUI table.

---

## How It Works

1. **Authentication:** Connects to Microsoft Graph using app credentials.
2. **User Iteration:** Reads UPNs from your user list and checks for a valid Copilot license.
3. **Data Retrieval:** Queries the `/beta/copilot/users/{user}/interactionHistory/getAllEnterpriseInteractions` Graph endpoint.
4. **Throttling Handling:** Retries automatically with exponential backoff if API limits are hit.
5. **Filtering:** Optionally removes AI responses for a user-centric activity view.
6. **Summarization:** Calculates per-user and overall Copilot activity statistics.
7. **Export:** Saves the final report as CSV/Excel in your Downloads folder.

---

## Troubleshooting

- **No Data Returned:** Ensure users have Copilot licenses and recent activity within the date range.
- **Authentication Errors:** Verify app credentials, permissions, and tenant ID.
- **Permission Denied:** Confirm that `User.Read.All` and `AiEnterpriseInteraction.Read.All` are granted and admin consented.
- **Throttling:** The script automatically waits and retries, but you may need to reduce the number of users or extend time between calls if rate limits persist.
- **Excel Export Fails:** Install the ImportExcel module, or set `ExportOption` to `'CSV'`.

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

- **Author(s):** Mike Lee, Jay Thakker, Tony Redmond
- **References:**
  - [Microsoft Graph API Throttling](https://learn.microsoft.com/en-us/graph/throttling)
  - [ImportExcel PowerShell Module](https://www.powershellgallery.com/packages/ImportExcel)
  - [Office 365 for IT Pros Scripts](https://github.com/12Knocksinna/Office365itpros)

---

## License

MIT License (see repository for details).

---

## Support

For issues, open a [GitHub issue](https://github.com/mikelee1313/Copilot/issues) or contact the script authors.

---

**Happy reporting!**

---
