# TenantWide_SP_Copilot_Agents_Insights.ps1

## Overview

**TenantWide_SP_Copilot_Agents_Insights.ps1** is a PowerShell script designed to discover and document all SharePoint Copilot Agents across a Microsoft 365 tenant. It generates a comprehensive CSV report with detailed information about each Copilot Agent and its host site, leveraging both the Microsoft Graph API and PnP PowerShell.

---

## Features

- Discovers all SharePoint Copilot Agents tenant-wide
- Collects:
  - Copilot Agent details (name, created date, last accessed date, owner)
  - Site information (template, owner, sensitivity label)
  - Security settings (information barriers, external sharing, access restrictions)
- Outputs:
  - CSV report with all Copilot Agents and their site details
  - Log file with detailed execution info

---

## Prerequisites

- PowerShell 7+
- [PnP.PowerShell](https://pnp.github.io/powershell/) module installed
- An Entra (Azure AD) App Registration with:
  - Application permissions: `Sites.FullControl.All`, `Sites.Read.All`
  - Certificate-based authentication configured

---

## Parameters

| Parameter     | Description                                                       |
|---------------|-------------------------------------------------------------------|
| `tenantname`  | Your M365 tenant name (without `.onmicrosoft.com`)                |
| `appID`       | The Entra Application (client) ID                                 |
| `thumbprint`  | Certificate thumbprint for authentication                         |
| `tenantid`    | The M365 tenant ID (GUID)                                         |
| `searchRegion`| Region for Microsoft Graph search (e.g., NAM, EMEA, APAC)         |

Parameters are set at the top of the script.

---

## Usage

```powershell
# Run the script from PowerShell
.\TenantWide_SP_Copilot_Agents_Insights.ps1
```

No arguments are required if you’ve set your variables at the top of the script file.

---

## Output

- **CSV file** in your TEMP directory (e.g., `SPO_Copilot_Agents_YYYY-MM-DD_HH-mm-ss.csv`)
- **Log file** in your TEMP directory

---

## Example

```powershell
.\TenantWide_SP_Copilot_Agents_Insights.ps1
```

---

## References

- [Microsoft Graph API Search Documentation](https://learn.microsoft.com/en-us/graph/api/search-query?view=graph-rest-1.0&tabs=http)
- [Microsoft 365 Copilot](https://learn.microsoft.com/microsoft-365-copilot/)
- [Insights report on SharePoint agents](https://learn.microsoft.com/en-us/sharepoint/insights-on-sharepoint-agents)

---

## Notes

- Author: Mike Lee
- Created: 7/7/2025
- License/Disclaimer: Provided AS IS, without warranty or support.
