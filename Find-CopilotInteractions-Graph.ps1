<#
.SYNOPSIS
    Retrieves and analyzes Microsoft Copilot interactions for specified users using the # API query settings
$MaxResultsPerRequest = 100  # Maximum number of results per Graph API request (used in $top parameter)
$RequestDelayMs = 250  # Delay in milliseconds between requests to help prevent RPS throttling (1000ms = 1 second)
                      # The Copilot API has a limit of 30 RPS per application. Set to 0 to disable delays.
                      # Recommended: 250ms (allows ~4 RPS) or 100ms (allows ~10 RPS) for safety margin

# Debug and monitoring options
$Debug = $true  # Set to $true to log ALL requests and responses (successful and failed) to the log file
$TrackRPS = $true  # Set to $false to disable RPS tracking, calculations, and progress displayGraph API.

.DESCRIPTION
    This script uses the Microsoft Graph API to fetch and analyze Copilot interactions for a list of users over a
    specified time period. It provides insights into user interaction patterns with Copilot across different 
    Microsoft 365 applications.

    The script requires an app registration with User.Read.All and AiEnterpriseInteraction.Read.All permissions,
    and can authenticate using either a client secret or certificate.

.PARAMETER appID
    The application (client) ID of the registered app in Azure AD.

.PARAMETER TenantId
    The tenant ID where the app is registered.

.PARAMETER AuthType
    The authentication type to use. Valid values: 'ClientSecret' or 'Certificate'.

.PARAMETER ClientSecret
    The client secret for the registered app (used when AuthType is 'ClientSecret').

.PARAMETER Thumbprint
    The certificate thumbprint (used when AuthType is 'Certificate').

.PARAMETER UserListPath
    Path to a text file containing user principal names (one per line).

.PARAMETER IncludeAIResponses
    When true, includes AI responses in the report. When false, only user prompts are shown.

.PARAMETER ShowGridView
    When true, displays an interactive GridView of the results at the end.

.PARAMETER ExportOption
    The format to use when exporting data. Valid values: 'XLSX' or 'CSV'. XLSX requires the ImportExcel module.

.PARAMETER DaysToLookBack
    Number of days in the past to check for Copilot interactions.

.PARAMETER DaysToLookAhead
    Number of days in the future to include in the search range.

.PARAMETER MaxResultsPerRequest
    Maximum number of results per Microsoft Graph API request (controls the $top parameter).

.EXAMPLE
    PS> .\Find-CopilotInteractions-Graph.ps1

.NOTES
Author: Mike Lee | Jay Thakker | Tony Redmond (Office 365 for IT Pros)
Date: 8/7/25
Version: 1.8

    - Requires Microsoft Graph PowerShell SDK module
    - Requires an app with User.Read.All and AiEnterpriseInteraction.Read.All permissions
    - Optional: ImportExcel module for Excel output (falls back to CSV if not available)
    - Users must have a Copilot license assigned
    - The script attempts to identify automatic interactions vs. user-initiated interactions
    - Output files are saved to the Downloads folder
    - Includes throttling protection for Microsoft Graph API calls with exponential backoff
    - Logs all non-200 HTTP responses to help diagnose throttling and other API issues
    - Includes configurable delay between requests to help prevent RPS throttling

.LINK
    https://github.com/12Knocksinna/Office365itpros/blob/master/Find-CopilotInteractions-Graph.PS1
    https://www.powershellgallery.com/packages/ImportExcel/7.8.10
#>

#############################################################
#                  CONFIGURATION SECTION                    #
#############################################################

# Authentication settings - Required
# The app must have User.Read.All and AiEnterpriseInteraction.Read.All permissions
$appID = 'abc64618-283f-47ba-a185-50d935d51d57'
$TenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3'

# Authentication type: Choose 'ClientSecret' or 'Certificate'
$AuthType = 'Certificate'  # Valid values: 'ClientSecret' or 'Certificate'

# Client Secret authentication (used when $AuthType = 'ClientSecret')
$ClientSecret = ''

# Certificate authentication (used when $AuthType = 'Certificate')
$Thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"

# User list - Path to file containing user principal names (one per line)
$UserListPath = "C:\temp\simpleuserlist.txt"   # File should contain one UPN per line, no headers

# Report filtering options
$IncludeAIResponses = $true  # Set to $false to exclude AI responses and only show user prompts
$ShowGridView = $false  # Set to $false to disable the interactive GridView popup at the end

# Export options
$ExportOption = "XLSX"  # Valid values: 'XLSX' or 'CSV'. XLSX requires the ImportExcel module.

# Date range for report
$DaysToLookBack = 60  # Number of days in the past to check
$DaysToLookAhead = 1  # Number of days in the future to include

# API query settings
$MaxResultsPerRequest = 100  # Maximum number of results per Graph API request (used in $top parameter)
$RequestDelayMs = 250  # Delay in milliseconds between requests to help prevent RPS throttling (1000ms = 1 second)
# The Copilot API has a limit of 30 RPS per application. Set to 0 to disable delays.
# Recommended: 250ms (allows ~4 RPS) or 100ms (allows ~10 RPS) for safety margin

# Debug and monitoring options
$Debug = $false  # Set to $true to log ALL requests and responses (successful and failed) to the log file
$TrackRPS = $true  # Set to $false to disable RPS tracking, calculations, and progress display
#############################################################
#                   END CONFIGURATION                       #
#############################################################

# Initialize logging variables
$LogFilePrefix = if ($Debug) { "CopilotScript-Debug" } else { "CopilotScript-Errors" }
$LogFilePath = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\$LogFilePrefix-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
$LogFileHeaderWritten = $false

# Initialize RPS tracking variables (only if tracking is enabled)
if ($TrackRPS) {
    $RequestTimes = [System.Collections.Generic.List[DateTime]]::new()
    $TotalRequests = 0
    $ScriptStartTime = Get-Date
    Write-Host "RPS tracking enabled - monitoring request rates" -ForegroundColor Cyan
}
else {
    Write-Host "RPS tracking disabled" -ForegroundColor Gray
}

# Debug mode notification
if ($Debug) {
    Write-Host "Debug mode enabled - logging ALL requests and responses to log file" -ForegroundColor Yellow
}
else {
    Write-Host "Debug mode disabled - logging only non-200 responses to log file" -ForegroundColor Gray
}

# Function to write log entry directly to disk
function Write-LogEntryToDisk {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$LogEntry
    )
    
    # Write CSV header if this is the first entry
    if (-not $script:LogFileHeaderWritten) {
        $LogEntry | Export-Csv -Path $script:LogFilePath -NoTypeInformation -Encoding Utf8
        $script:LogFileHeaderWritten = $true
    }
    else {
        # Append to existing file without header
        $LogEntry | Export-Csv -Path $script:LogFilePath -NoTypeInformation -Encoding Utf8 -Append
    }
}

# Function to calculate current RPS
function Get-CurrentRPS {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [int]$WindowSeconds = 10
    )
    
    # Return empty values if RPS tracking is disabled
    if (-not $script:TrackRPS) {
        return [PSCustomObject]@{
            RequestsInWindow = "N/A"
            WindowSeconds    = $WindowSeconds
            CurrentRPS       = "N/A"
            TotalRequests    = "N/A"
            ElapsedMinutes   = "N/A"
            AverageRPS       = "N/A"
        }
    }
    
    $now = Get-Date
    $cutoffTime = $now.AddSeconds(-$WindowSeconds)
    
    # Remove requests older than the window
    $RequestTimes.RemoveAll([Predicate[DateTime]] { param($time) $time -lt $cutoffTime })
    
    # Calculate RPS over the window
    $requestsInWindow = $RequestTimes.Count
    $actualRPS = if ($WindowSeconds -gt 0) { [math]::Round($requestsInWindow / $WindowSeconds, 2) } else { 0 }
    
    return [PSCustomObject]@{
        RequestsInWindow = $requestsInWindow
        WindowSeconds    = $WindowSeconds
        CurrentRPS       = $actualRPS
        TotalRequests    = $script:TotalRequests
        ElapsedMinutes   = [math]::Round(($now - $script:ScriptStartTime).TotalMinutes, 2)
        AverageRPS       = if (($now - $script:ScriptStartTime).TotalSeconds -gt 0) { 
            [math]::Round($script:TotalRequests / ($now - $script:ScriptStartTime).TotalSeconds, 2) 
        }
        else { 0 }
    }
}

# Function to log ALL requests and responses when Debug mode is enabled
function Write-DebugLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        
        [Parameter(Mandatory = $true)]
        [string]$Method,
        
        [Parameter(Mandatory = $true)]
        [int]$StatusCode,
        
        [Parameter(Mandatory = $false)]
        [string]$StatusDescription = "",
        
        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage = "",
        
        [Parameter(Mandatory = $false)]
        [hashtable]$ResponseHeaders = @{},
        
        [Parameter(Mandatory = $false)]
        [string]$UserContext = "",
        
        [Parameter(Mandatory = $false)]
        [string]$RequestType = "API Call"
    )
    
    # Get current RPS information if tracking is enabled
    $rpsInfo = if ($script:TrackRPS) { Get-CurrentRPS } else { 
        [PSCustomObject]@{
            CurrentRPS       = "N/A"
            RequestsInWindow = "N/A"
            TotalRequests    = "N/A"
            ElapsedMinutes   = "N/A"
            AverageRPS       = "N/A"
        }
    }
    
    $LogEntry = [PSCustomObject]@{
        Timestamp               = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        RequestType             = $RequestType
        UserContext             = $UserContext
        Method                  = $Method
        Uri                     = $Uri
        StatusCode              = $StatusCode
        StatusDescription       = $StatusDescription
        ErrorMessage            = $ErrorMessage
        RetryAfterHeader        = if ($ResponseHeaders.ContainsKey("Retry-After")) { $ResponseHeaders["Retry-After"] } else { "N/A" }
        ThrottleScope           = if ($ResponseHeaders.ContainsKey("x-ms-throttle-scope")) { $ResponseHeaders["x-ms-throttle-scope"] } else { "N/A" }
        ThrottleInformation     = if ($ResponseHeaders.ContainsKey("x-ms-throttle-information")) { $ResponseHeaders["x-ms-throttle-information"] } else { "N/A" }
        ResourceUnit            = if ($ResponseHeaders.ContainsKey("x-ms-resource-unit")) { $ResponseHeaders["x-ms-resource-unit"] } else { "N/A" }
        ThrottleLimitPercentage = if ($ResponseHeaders.ContainsKey("x-ms-throttle-limit-percentage")) { $ResponseHeaders["x-ms-throttle-limit-percentage"] } else { "N/A" }
        CurrentRPS_10sec        = $rpsInfo.CurrentRPS
        RequestsInWindow_10sec  = $rpsInfo.RequestsInWindow
        TotalRequests           = $rpsInfo.TotalRequests
        ElapsedMinutes          = $rpsInfo.ElapsedMinutes
        AverageRPS_Overall      = $rpsInfo.AverageRPS
    }
    
    # Write log entry directly to disk
    Write-LogEntryToDisk -LogEntry $LogEntry
    
    # Also write to console for immediate feedback in debug mode
    if ($StatusCode -eq 200) {
        Write-Host "[$($LogEntry.Timestamp)] DEBUG - Status: $StatusCode, User: $UserContext, RPS: $($rpsInfo.CurrentRPS)" -ForegroundColor Green
    }
    else {
        Write-Host "[$($LogEntry.Timestamp)] DEBUG - Status: $StatusCode, User: $UserContext, RPS: $($rpsInfo.CurrentRPS)" -ForegroundColor Red
    }
}

# Function to log non-200 responses
function Write-ErrorLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        
        [Parameter(Mandatory = $true)]
        [string]$Method,
        
        [Parameter(Mandatory = $true)]
        [int]$StatusCode,
        
        [Parameter(Mandatory = $false)]
        [string]$StatusDescription = "",
        
        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage = "",
        
        [Parameter(Mandatory = $false)]
        [hashtable]$ResponseHeaders = @{},
        
        [Parameter(Mandatory = $false)]
        [string]$UserContext = ""
    )
    
    # Get current RPS information if tracking is enabled
    $rpsInfo = if ($script:TrackRPS) { Get-CurrentRPS } else { 
        [PSCustomObject]@{
            CurrentRPS       = "N/A"
            RequestsInWindow = "N/A"
            TotalRequests    = "N/A"
            ElapsedMinutes   = "N/A"
            AverageRPS       = "N/A"
        }
    }
    
    $LogEntry = [PSCustomObject]@{
        Timestamp               = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        UserContext             = $UserContext
        Method                  = $Method
        Uri                     = $Uri
        StatusCode              = $StatusCode
        StatusDescription       = $StatusDescription
        ErrorMessage            = $ErrorMessage
        RetryAfterHeader        = if ($ResponseHeaders.ContainsKey("Retry-After")) { $ResponseHeaders["Retry-After"] } else { "N/A" }
        ThrottleScope           = if ($ResponseHeaders.ContainsKey("x-ms-throttle-scope")) { $ResponseHeaders["x-ms-throttle-scope"] } else { "N/A" }
        ThrottleInformation     = if ($ResponseHeaders.ContainsKey("x-ms-throttle-information")) { $ResponseHeaders["x-ms-throttle-information"] } else { "N/A" }
        ResourceUnit            = if ($ResponseHeaders.ContainsKey("x-ms-resource-unit")) { $ResponseHeaders["x-ms-resource-unit"] } else { "N/A" }
        ThrottleLimitPercentage = if ($ResponseHeaders.ContainsKey("x-ms-throttle-limit-percentage")) { $ResponseHeaders["x-ms-throttle-limit-percentage"] } else { "N/A" }
        CurrentRPS_10sec        = $rpsInfo.CurrentRPS
        RequestsInWindow_10sec  = $rpsInfo.RequestsInWindow
        TotalRequests           = $rpsInfo.TotalRequests
        ElapsedMinutes          = $rpsInfo.ElapsedMinutes
        AverageRPS_Overall      = $rpsInfo.AverageRPS
    }
    
    # Write log entry directly to disk
    Write-LogEntryToDisk -LogEntry $LogEntry
    
    # Also write to console for immediate feedback
    Write-Host "[$($LogEntry.Timestamp)] ERROR - Status: $StatusCode, User: $UserContext, RPS: $($rpsInfo.CurrentRPS)" -ForegroundColor Red
    if ($LogEntry.ThrottleScope -ne "N/A") {
        Write-Host "  Throttle Scope: $($LogEntry.ThrottleScope)" -ForegroundColor Yellow
    }
    if ($LogEntry.ThrottleInformation -ne "N/A") {
        Write-Host "  Throttle Info: $($LogEntry.ThrottleInformation)" -ForegroundColor Yellow
    }
    if ($LogEntry.RetryAfterHeader -ne "N/A") {
        Write-Host "  Retry-After: $($LogEntry.RetryAfterHeader) seconds" -ForegroundColor Yellow
    }
}

# Function to handle throttling for Microsoft Graph requests
# This implements best practices from https://learn.microsoft.com/en-us/graph/throttling
# It automatically handles 429 "Too Many Requests" responses with either:
# 1. The Retry-After header value if provided by the server
# 2. Exponential backoff if no Retry-After header is present
# Enhanced with comprehensive logging for all non-200 responses
function Invoke-MgGraphRequestWithThrottleHandling {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        
        [Parameter(Mandatory = $true)]
        [string]$Method,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 10,
        
        [Parameter(Mandatory = $false)]
        [int]$InitialBackoffSeconds = 2,
        
        [Parameter(Mandatory = $false)]
        [string]$UserContext = ""
    )
    
    $retryCount = 0
    $backoffSeconds = $InitialBackoffSeconds
    $success = $false
    $result = $null
    
    Write-Verbose "Making Graph request to $Uri"
    
    while (-not $success -and $retryCount -lt $MaxRetries) {
        try {
            $requestStartTime = Get-Date
            $result = Invoke-MgGraphRequest -Uri $Uri -Method $Method -ErrorAction Stop
            $success = $true
            
            # Track successful request for RPS calculation (only if tracking enabled)
            if ($script:TrackRPS) {
                $script:TotalRequests++
                $script:RequestTimes.Add($requestStartTime)
                
                # Get current RPS for progress reporting
                $rpsInfo = Get-CurrentRPS
                
                # Display RPS info for high-volume operations (every 10th request or when RPS > 5)
                if (($script:TotalRequests % 10 -eq 0) -or ($rpsInfo.CurrentRPS -gt 5)) {
                    Write-Host ("Request #$($script:TotalRequests) - Current RPS: $($rpsInfo.CurrentRPS), Avg: $($rpsInfo.AverageRPS)") -ForegroundColor Gray
                }
            }
            
            # Log successful request if Debug mode is enabled
            if ($script:Debug) {
                Write-DebugLog -Uri $Uri -Method $Method -StatusCode 200 -StatusDescription "OK" -UserContext $UserContext -RequestType "Successful Request"
            }
            
            # Add a small delay between successful requests to help prevent RPS throttling
            if ($script:RequestDelayMs -gt 0) {
                Start-Sleep -Milliseconds $script:RequestDelayMs
            }
        }
        catch {
            # Track failed request for RPS calculation too (only if tracking enabled)
            if ($script:TrackRPS) {
                $script:TotalRequests++
                $script:RequestTimes.Add((Get-Date))
            }
            
            $statusCode = $_.Exception.Response.StatusCode.value__
            $statusDescription = $_.Exception.Response.ReasonPhrase
            $errorMessage = $_.Exception.Message
            
            # Extract response headers for logging
            $responseHeaders = @{}
            if ($_.Exception.Response.Headers) {
                foreach ($header in $_.Exception.Response.Headers) {
                    $responseHeaders[$header.Key] = ($header.Value -join ", ")
                }
            }
            
            # Log based on Debug mode setting
            if ($script:Debug) {
                # In debug mode, log all responses (including errors) using debug log
                Write-DebugLog -Uri $Uri -Method $Method -StatusCode $statusCode -StatusDescription $statusDescription -ErrorMessage $errorMessage -ResponseHeaders $responseHeaders -UserContext $UserContext -RequestType "Failed Request"
            }
            else {
                # In normal mode, only log non-200 responses using error log
                Write-ErrorLog -Uri $Uri -Method $Method -StatusCode $statusCode -StatusDescription $statusDescription -ErrorMessage $errorMessage -ResponseHeaders $responseHeaders -UserContext $UserContext
            }
            
            # Check if this is a throttling error (429)
            if ($statusCode -eq 429) {
                # Get the Retry-After header if it exists
                $retryAfter = $null
                if ($_.Exception.Response.Headers.Contains("Retry-After")) {
                    $retryAfter = [int]($_.Exception.Response.Headers.GetValues("Retry-After") | Select-Object -First 1)
                    Write-Host "Request throttled. Retry-After header suggests waiting for $retryAfter seconds." -ForegroundColor Yellow
                }
                else {
                    # If no Retry-After header, use exponential backoff
                    $retryAfter = $backoffSeconds
                    Write-Host "Request throttled. Using exponential backoff: waiting for $retryAfter seconds." -ForegroundColor Yellow
                    # Increase backoff for next potential retry (exponential)
                    $backoffSeconds = $backoffSeconds * 2
                }
                
                $retryCount++
                if ($retryCount -lt $MaxRetries) {
                    Write-Host "Throttling detected. Waiting before retry. Attempt $retryCount of $MaxRetries..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $retryAfter
                }
                else {
                    Write-Host "Maximum retry attempts reached ($MaxRetries). Giving up." -ForegroundColor Red
                    throw $_
                }
            }
            else {
                # Not a throttling error, rethrow immediately
                throw $_
            }
        }
    }
    
    return $result
}

# Disconnect any previous Graph session before connecting
# Disconnect-MgGraph | Out-Null

# Connect to Microsoft Graph based on authentication type
Write-Host "Connecting to Microsoft Graph using $AuthType authentication..." -ForegroundColor Cyan

if ($AuthType -eq 'ClientSecret') {
    # Connect using client secret
    $SecureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
    $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $appID, $SecureSecret
    
    try {
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $Credential -ErrorAction Stop
        Write-Host "Successfully connected using Client Secret authentication" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to connect using Client Secret authentication" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Exit
    }
}
elseif ($AuthType -eq 'Certificate') {
    # Connect using certificate
    try {
        Connect-MgGraph -AppId $appID -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome -ErrorAction Stop
        Write-Host "Successfully connected using Certificate authentication" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to connect using Certificate authentication" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Exit
    }
}
else {
    Write-Host "Invalid authentication type: $AuthType. Valid values are 'ClientSecret' or 'Certificate'." -ForegroundColor Red
    Exit
}

# Calculate date range from configuration
$StartDate = (Get-Date).AddDays(-$DaysToLookBack).toString('yyyy-MM-ddT00:00:00Z')
$EndDate = (Get-Date).AddDays($DaysToLookAhead).toString('yyyy-MM-ddT00:00:00Z')
$StartDateForReport = Get-Date $StartDate -format 'dd-MMM-yyyy'
$EndDateForReport = Get-Date $EndDate -format 'dd-MMM-yyyy'

# Create a master report list to hold all users' data
$MasterReport = [System.Collections.Generic.List[Object]]::new()

# Load user list from file
if (Test-Path $UserListPath) {
    $userList = Get-Content $UserListPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.StartsWith("#") }
    if ($userList.Count -eq 0) {
        Write-Host "No valid users found in $UserListPath. The file may be empty or contain only comments." -ForegroundColor Red
        Exit
    }
}
else {
    Write-Host "User list file not found at $UserListPath. Please create this file with one user principal name per line." -ForegroundColor Red
    # Create a sample file with instructions
    New-Item -Path $UserListPath -ItemType File -Force -ErrorAction SilentlyContinue | Out-Null
    "# Enter one user principal name per line" | Out-File -FilePath $UserListPath -Encoding utf8
    "# Example: user1@contoso.com" | Out-File -FilePath $UserListPath -Encoding utf8 -Append
    "# Example: user2@contoso.com" | Out-File -FilePath $UserListPath -Encoding utf8 -Append
    Write-Host "A sample file has been created at $UserListPath. Please edit this file and run the script again." -ForegroundColor Yellow
    Exit
}

Write-Host "Starting Copilot interaction analysis for $($userList.Count) users"
Write-Host "Date range: $StartDateForReport to $EndDateForReport"
if (-not $IncludeAIResponses) {
    Write-Host "AI Responses will be filtered out (only showing user prompts)" -ForegroundColor Cyan
}
else {
    Write-Host "Both user prompts and AI responses will be included" -ForegroundColor Cyan
}
Write-Host ""

# Process each user
foreach ($UserPrincipalName in $userList) {
    Write-Host "`n=========================================================="
    Write-Host "Processing user: $UserPrincipalName"
    Write-Host "==========================================================`n"
    
    $User = Get-MgUser -UserId $UserPrincipalName.trim() -ErrorAction SilentlyContinue
    If (!$User) {
        Write-Host ("User {0} not found. Skipping to next user." -f $UserPrincipalName) -ForegroundColor Yellow
        Continue
    }
    
    # Has the account got a Copilot license?
    [array]$UserLicenses = Get-MgUserLicenseDetail -UserId $User.Id | Select-Object -ExpandProperty SkuId
    If ("639dec6b-bb19-468b-871c-c5c441c4b0cb" -notin $UserLicenses) {
        Write-Host ("User {0} does not have a Copilot license, so we can't check their Copilot interactions. Skipping to next user." -f $User.DisplayName) -ForegroundColor Yellow
        Continue
    }

    $Uri = ("https://graph.microsoft.com/v1.0/copilot/users/{0}/interactionHistory/getAllEnterpriseInteractions?`$filter=createdDateTime gt {1} and createdDateTime lt {2}&`$top={3}" `
            -f $User.Id, $StartDate, $EndDate, $MaxResultsPerRequest)

    Write-Host ("Searching for Copilot interactions for {0} between {1} and {2}" -f $User.DisplayName, $StartDateForReport, $EndDateForReport)
    [array]$CopilotData = $null
    # Get the first set of records
    Try {
        [array]$Data = Invoke-MgGraphRequestWithThrottleHandling -Uri $Uri -Method GET -UserContext $User.UserPrincipalName
        $CopilotData = $Data.Value
        If (!($CopilotData)) {
            Write-Host ("No Copilot interactions found for {0} between {1} and {2}. Skipping to next user." -f $User.DisplayName, $StartDateForReport, $EndDateForReport) -ForegroundColor Yellow
            Continue
        }
    }
    Catch {
        Write-Host "Error retrieving Copilot interaction data:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        if ($_.Exception.Response) {
            Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__ -ForegroundColor Red
            Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription -ForegroundColor Red
        }
        Write-Host "Skipping to next user." -ForegroundColor Yellow
        Continue
    }

    $Nextlink = $Data.'@odata.nextLink'
    While ($null -ne $Nextlink) {
        Write-Host ("Fetching more records - currently at {0}" -f $CopilotData.count)
        Try {
            [array]$Data = Invoke-MgGraphRequestWithThrottleHandling -Uri $Nextlink -Method GET -UserContext $User.UserPrincipalName
            $CopilotData += $Data.Value
            $Nextlink = $Data.'@odata.nextLink'
        }
        Catch {
            Write-Host "Error retrieving additional Copilot interaction data:" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
            # Break the loop if we encounter an error
            $Nextlink = $null
        }
    }
    # Remove any null records
    $CopilotData = $CopilotData | Where-Object { $_ -ne $null }
    $CopilotData = $CopilotData | Sort-Object { $_.createdDateTime -as [datetime] }
    
    # Apply AI response filtering based on configuration
    if (-not $IncludeAIResponses) {
        $OriginalCount = $CopilotData.Count
        $CopilotData = $CopilotData | Where-Object { $_.interactionType -ne "aiResponse" }
        Write-Host ("Filtering: Removed {0} AI responses based on configuration" -f ($OriginalCount - $CopilotData.Count)) -ForegroundColor Cyan
    }

    Write-Host ("{0} Copilot interactions for {1} between {2} and {3} have been retrieved" -f $CopilotData.count, $User.DisplayName, $StartDateForReport, $EndDateForReport)
    $Report = [System.Collections.Generic.List[Object]]::new()

    ForEach ($Record in $CopilotData) {

        If ($Record.createdDateTime) {
            $Timestamp = Get-Date $Record.createdDateTime -format 'dd-MMM-yyyy HH:mm:ss'
        }
        else {
            $Timestamp = $null
        }
        Switch ($Record.interactionType) {
            "userPrompt" {
                $AppName = $User.displayname
                $AppId = $record.from.user.id
            }
            "aiResponse" {
                $AppName = $Record.from.application.displayname
                $AppId = $Record.from.application.id
            }
            Default {
                $AppName = $Record.interfrom.application.displayname
                $AppId = $Record.from.application.id
            }
        }

        # Capture the full body content without truncation
        $Body = $Record.body.content.ToString()
        
        $AutoGeneratedFlag = $False
        # This section checks for some of the fingerprints that indicate that the interaction is automatic rather than user-generated
        Switch ($AppName) {
            "Copilot in Outlook" {
                $AppName = "Outlook"
                If ($Body -like "*VisualTheming*" -or $Body -like "*data:image;base64*") {
                    $AutoGeneratedFlag = $True
                }
            }
            "Copilot in Word" {
                If ($Body -like "*[AutoGenerated]*") {
                    $AutoGeneratedFlag = $True  
                }
            }
        }

        $ReportLine = [pscustomobject]@{
            User            = $User.UserPrincipalName
            DisplayName     = $User.DisplayName
            Timestamp       = $Timestamp
            'Copilot App'   = $AppName
            AppId           = $AppId
            Contexts        = ($Record.contexts.displayName -join ", ")
            InteractionType = $Record.interactionType
            ThreadId        = $Record.sessionid
            Body            = $Body
            Attachments     = ($Record.attachments.name -join ", ")
            Mentions        = ($Record.mentions.name -join ", ")
            Links           = ($Record.Links.LinkUrl -join ", ")
            AutoGenerated   = $AutoGeneratedFlag
        }
        $Report.Add($ReportLine)
        # Also add to the master report
        $MasterReport.Add($ReportLine)
    }

    # Display per-user statistics
    Write-Host "`nStatistics for $($User.DisplayName):"
    # Some basic computations
    $NumberOfAutomaticInteractions = $Report | Where-Object { $_.AutoGenerated -eq $True } | Measure-Object | Select-Object -ExpandProperty Count
    $UserInteractions = $Report | Where-Object { $_.InteractionType -eq "userPrompt" } | Measure-Object | Select-Object -ExpandProperty Count
    $CopilotResponses = $Report.Count - ($UserInteractions + $NumberOfAutomaticInteractions)
    
    if ($Report.Count -gt 0) {
        $PercentCopilotResponses = ($CopilotResponses / $Report.Count).ToString("P")
        $PercentAutomaticInteractions = ($NumberOfAutomaticInteractions / $Report.Count).ToString("P")
        $PercentUserInteractions = ($UserInteractions / $Report.Count).ToString("P")
    
        Write-Host ("{0} of the {1} interactions are automatic ({2})" -f $NumberOfAutomaticInteractions, $Report.Count, $PercentAutomaticInteractions)
        Write-Host ("{0} of the interactions are user prompts ({1})" -f $UserInteractions, $PercentUserInteractions)
        Write-Host ("{0} of the interactions are Copilot responses ({1})" -f $CopilotResponses, $PercentCopilotResponses)
        
        $Report | Group-Object 'Copilot App' | Select-Object Name, Count | Sort-Object Count -Descending | Format-Table
    }
}

# End of user foreach loop

# Final report with combined data from all users
Write-Host "`n=========================================================="
Write-Host "Processing complete. Generating final report for all users."
Write-Host "==========================================================`n"

if ($MasterReport.Count -eq 0) {
    Write-Host "No Copilot interactions found for any users in the specified date range." -ForegroundColor Yellow
    Exit
}

# Allow viewing the master report
if ($ShowGridView) {
    Write-Host "Opening interactive GridView..." -ForegroundColor Cyan
    $MasterReport | Out-GridView -Title "Copilot Interactions for All Users"
}
else {
    Write-Host "Interactive GridView disabled (set ShowGridView to true in configuration to enable)" -ForegroundColor Cyan
}

# Some basic computations across all users
$NumberOfAutomaticInteractions = $MasterReport | Where-Object { $_.AutoGenerated -eq $True } | Measure-Object | Select-Object -ExpandProperty Count
$UserInteractions = $MasterReport | Where-Object { $_.InteractionType -eq "userPrompt" } | Measure-Object | Select-Object -ExpandProperty Count
$CopilotResponses = $MasterReport.Count - ($UserInteractions + $NumberOfAutomaticInteractions)
$PercentCopilotResponses = ($CopilotResponses / $MasterReport.Count).ToString("P")
$PercentAutomaticInteractions = ($NumberOfAutomaticInteractions / $MasterReport.Count).ToString("P")
$PercentUserInteractions = ($UserInteractions / $MasterReport.Count).ToString("P")

Write-Host ""
Write-Host ("Aggregated Copilot interactions for all users between {0} and {1}" -f $StartDateForReport, $EndDateForReport)
Write-Host ("Total records: {0}" -f $MasterReport.Count)
Write-Host ("Unique users with data: {0}" -f ($MasterReport | Select-Object User -Unique | Measure-Object).Count)
Write-Host ""

Write-Host "Summary by user:"
$MasterReport | Group-Object DisplayName | Select-Object Name, Count | Sort-Object Count -Descending | Format-Table

Write-Host "Summary by application:"
$MasterReport | Group-Object 'Copilot App' | Select-Object Name, Count | Sort-Object Count -Descending | Format-Table

Write-Host ("{0} of the {1} interactions are automatic ({2})" -f $NumberOfAutomaticInteractions, $MasterReport.Count, $PercentAutomaticInteractions)
Write-Host ("{0} of the interactions are user prompts ({1})" -f $UserInteractions, $PercentUserInteractions)
Write-Host ("{0} of the interactions are Copilot responses ({1})" -f $CopilotResponses, $PercentCopilotResponses)

# Generate reports
$TimeStamp = Get-Date -Format "yyyyMMdd-HHmmss"

# Check if we can generate an Excel file if that's the preferred option
$ExcelGenerated = $false
if ($ExportOption -eq "XLSX") {
    if (Get-Module ImportExcel -ListAvailable) {
        $ExcelGenerated = $true
        $ExcelTitle = ("Copilot interactions between {0} and {1}" -f $StartDateForReport, $EndDateForReport)
        Import-Module ImportExcel -ErrorAction SilentlyContinue
        $OutputXLSXFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\CopilotInteractions-$TimeStamp.xlsx"
        If (Test-Path $OutputXLSXFile) {
            Remove-Item $OutputXLSXFile -ErrorAction SilentlyContinue
        }
        $MasterReport | Export-Excel -Path $OutputXLSXFile -WorksheetName "Copilot Interactions" -Title $ExcelTitle -TitleBold -TableName "CopilotInteractions" 
        Write-Host ("An Excel worksheet containing the report data is available in {0}" -f $OutputXLSXFile)
    }
    else {
        Write-Host "The ImportExcel module is not available. Falling back to CSV export." -ForegroundColor Yellow
        $ExportOption = "CSV"  # Fall back to CSV
    }
}

# Generate CSV if that's the preferred option or if Excel export failed
if ($ExportOption -eq "CSV" -or -not $ExcelGenerated) {
    $OutputCSVFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\CopilotInteractions-$TimeStamp.csv"
    $MasterReport | Export-Csv -Path $OutputCSVFile -NoTypeInformation -Encoding Utf8
    Write-Host ("A CSV file containing the report data is available in {0}" -f $OutputCSVFile)
}

# Display RPS Performance Summary (only if RPS tracking is enabled)
if ($TrackRPS) {
    Write-Host ""
    Write-Host "=========================================================="
    Write-Host "API PERFORMANCE SUMMARY"
    Write-Host "=========================================================="
    $finalRpsInfo = Get-CurrentRPS
    Write-Host ("Total API requests made: {0}" -f $finalRpsInfo.TotalRequests) -ForegroundColor Cyan
    Write-Host ("Script execution time: {0} minutes" -f $finalRpsInfo.ElapsedMinutes) -ForegroundColor Cyan
    Write-Host ("Average RPS throughout execution: {0}" -f $finalRpsInfo.AverageRPS) -ForegroundColor Cyan
    Write-Host ("Current RPS (last 10 seconds): {0}" -f $finalRpsInfo.CurrentRPS) -ForegroundColor Cyan
    Write-Host ("Configured delay between requests: {0}ms" -f $RequestDelayMs) -ForegroundColor Cyan

    # Display RPS efficiency analysis
    $theoreticalMaxRPS = if ($RequestDelayMs -gt 0) { [math]::Round(1000 / $RequestDelayMs, 2) } else { "Unlimited" }
    Write-Host ("Theoretical max RPS with current delay: {0}" -f $theoreticalMaxRPS) -ForegroundColor Cyan

    if ($finalRpsInfo.AverageRPS -gt 25) {
        Write-Host "WARNING: Average RPS is approaching the 30 RPS limit for Copilot API!" -ForegroundColor Yellow
    }
    elseif ($finalRpsInfo.AverageRPS -gt 15) {
        Write-Host "INFO: RPS is in a reasonable range but monitor for throttling." -ForegroundColor Yellow
    }
    else {
        Write-Host "INFO: RPS is conservative and should avoid throttling." -ForegroundColor Green
    }
}
else {
    Write-Host ""
    Write-Host "RPS tracking was disabled - no performance metrics available" -ForegroundColor Gray
}

# Save error log if there were any logged entries
if ($LogFileHeaderWritten -and (Test-Path $LogFilePath)) {
    # Read the log file to get statistics
    $LogEntries = Import-Csv -Path $LogFilePath
    
    Write-Host ""
    Write-Host "=========================================================="
    Write-Host "ERROR LOG SUMMARY"
    Write-Host "=========================================================="
    Write-Host ("Total responses logged: {0}" -f $LogEntries.Count) -ForegroundColor Yellow
    Write-Host ("Log file saved to: {0}" -f $LogFilePath) -ForegroundColor Yellow
    
    # Display summary of errors by type
    $ErrorSummary = $LogEntries | Group-Object StatusCode | Select-Object Name, Count | Sort-Object Count -Descending
    Write-Host ""
    Write-Host "Response breakdown by status code:"
    $ErrorSummary | Format-Table -AutoSize
    
    # Display throttling information if present
    $ThrottleEntries = $LogEntries | Where-Object { $_.StatusCode -eq 429 }
    if ($ThrottleEntries.Count -gt 0) {
        Write-Host ""
        Write-Host "Throttling Analysis (Status 429):" -ForegroundColor Yellow
        Write-Host ("Total throttled requests: {0}" -f $ThrottleEntries.Count)
        
        # Show RPS information for throttled requests (only if RPS tracking was enabled)
        if ($TrackRPS) {
            $avgRpsAtThrottle = ($ThrottleEntries | Where-Object { $_.CurrentRPS_10sec -ne "N/A" } | Measure-Object -Property CurrentRPS_10sec -Average).Average
            $maxRpsAtThrottle = ($ThrottleEntries | Where-Object { $_.CurrentRPS_10sec -ne "N/A" } | Measure-Object -Property CurrentRPS_10sec -Maximum).Maximum
            if ($avgRpsAtThrottle -gt 0) {
                Write-Host ("Average RPS when throttling occurred: {0:F2}" -f $avgRpsAtThrottle) -ForegroundColor Yellow
                Write-Host ("Peak RPS when throttling occurred: {0:F2}" -f $maxRpsAtThrottle) -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "RPS data not available (tracking was disabled)" -ForegroundColor Gray
        }
        
        $ThrottleScopes = $ThrottleEntries | Where-Object { $_.ThrottleScope -ne "N/A" } | Group-Object ThrottleScope
        if ($ThrottleScopes.Count -gt 0) {
            Write-Host ""
            Write-Host "Throttle scopes encountered:"
            $ThrottleScopes | ForEach-Object {
                Write-Host ("  {0}: {1} occurrences" -f $_.Name, $_.Count)
            }
        }
        
        $ThrottleReasons = $ThrottleEntries | Where-Object { $_.ThrottleInformation -ne "N/A" } | Group-Object ThrottleInformation
        if ($ThrottleReasons.Count -gt 0) {
            Write-Host ""
            Write-Host "Throttle reasons:"
            $ThrottleReasons | ForEach-Object {
                Write-Host ("  {0}: {1} occurrences" -f $_.Name, $_.Count)
            }
        }
        
        Write-Host ""
        Write-Host "RECOMMENDATION:" -ForegroundColor Cyan
        Write-Host "Based on the Microsoft Graph throttling limits documentation:" -ForegroundColor Cyan
        Write-Host "- The /copilot/users/{user-id}/interactionHistory endpoint has a limit of 30 RPS per application" -ForegroundColor Cyan
        Write-Host "- Consider adding delays between requests or reducing concurrent processing" -ForegroundColor Cyan
        Write-Host "- Review the detailed log file for specific throttle scopes and reasons" -ForegroundColor Cyan
    }
    
    Write-Host "=========================================================="
}
else {
    Write-Host ""
    if ($Debug) {
        Write-Host "No requests were logged (possibly no users processed or script error)." -ForegroundColor Yellow
    }
    else {
        Write-Host "No HTTP errors encountered during script execution." -ForegroundColor Green
    }
}
