#Requires -Version 5.1

<#
.SYNOPSIS
    Reports Duo users in bypass mode across all subclients.

.DESCRIPTION
    Connects to the Duo Accounts API to enumerate all subclients, then
    queries each subclient's Admin API for users with bypass status.
    Results are grouped by subclient in console output and exported to CSV.

.PARAMETER OutputPath
    Path for the CSV output file. Defaults to the script directory with a
    timestamped filename (DuoBypassUsers-YYYYMMDD-HHmmss.csv).

.PARAMETER ResetCredentials
    Force re-entry of stored Duo credentials. ConnectWise credentials in
    the shared credentials file are preserved.

.EXAMPLE
    .\Get-DuoBypassUsers.ps1
    Queries all subclients and saves results to the script directory.

.EXAMPLE
    .\Get-DuoBypassUsers.ps1 -OutputPath "C:\Reports\bypass.csv"
    Saves results to the specified path.

.EXAMPLE
    .\Get-DuoBypassUsers.ps1 -ResetCredentials
    Clears stored Duo credentials and prompts for new ones.

.NOTES
    Author: CodeBlue Technology LLC
    Requires: DuoSecurity PowerShell module
    Shares credentials.xml with ProvisionDuoAccount.ps1
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$OutputPath,

    [Parameter()]
    [switch]$ResetCredentials
)

# Credential storage path - shared with ProvisionDuoAccount.ps1
$script:ConfigPath     = Join-Path $PSScriptRoot "Config"
$script:CredentialFile = Join-Path $script:ConfigPath "credentials.xml"

#region Helper Functions

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [Parameter()]
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $colors = @{
        'Info'    = 'Cyan'
        'Success' = 'Green'
        'Warning' = 'Yellow'
        'Error'   = 'Red'
    }

    Write-Host "[$timestamp] " -NoNewline -ForegroundColor Gray
    Write-Host "[$Level] " -NoNewline -ForegroundColor $colors[$Level]
    Write-Host $Message
}

#endregion

#region Credential Management

function Get-StoredCredentials {
    <#
    .SYNOPSIS
        Loads or prompts for Duo Accounts API credentials.
        Preserves any existing ConnectWise credentials in the shared file.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$Reset
    )

    if (-not (Test-Path $script:ConfigPath)) {
        New-Item -ItemType Directory -Path $script:ConfigPath -Force | Out-Null
    }

    # Try to load existing credentials (may contain both Duo and ConnectWise sections)
    $existingCreds = $null
    if (Test-Path $script:CredentialFile) {
        try {
            $existingCreds = Import-Clixml -Path $script:CredentialFile
        }
        catch {
            Write-Log "Could not read credential file, will prompt for new credentials" -Level Warning
        }
    }

    # Return existing if Duo section is present and no reset requested
    if ($existingCreds -and $existingCreds.Duo -and (-not $Reset)) {
        Write-Log "Loaded stored credentials" -Level Info
        return $existingCreds
    }

    # Prompt for Duo Accounts API credentials only
    Write-Host "`n=== Duo MSP Credentials ===" -ForegroundColor Cyan
    Write-Host "These credentials are for your Duo Accounts API (MSP level)`n"

    $duoApiHost        = Read-Host "Duo API Host (e.g., api-xxxxxxxx.duosecurity.com)"
    $duoIntegrationKey = Read-Host "Duo Integration Key"
    $duoSecretKey      = Read-Host "Duo Secret Key"

    # Merge into existing creds so we don't lose ConnectWise keys
    if (-not $existingCreds) {
        $existingCreds = @{}
    }

    $existingCreds.Duo = @{
        Type           = 'Accounts'
        ApiHost        = $duoApiHost
        IntegrationKey = $duoIntegrationKey
        SecretKey      = $duoSecretKey
    }

    try {
        $existingCreds | Export-Clixml -Path $script:CredentialFile -Force
        Write-Log "Credentials saved to $script:CredentialFile" -Level Success
    }
    catch {
        Write-Log "Failed to save credentials: $_" -Level Warning
    }

    return $existingCreds
}

#endregion

#region Prerequisites

function Test-Prerequisites {
    [CmdletBinding()]
    param()

    if (-not (Get-Module -ListAvailable -Name 'DuoSecurity')) {
        Write-Log "DuoSecurity module not found. Installing..." -Level Warning
        try {
            Install-Module -Name 'DuoSecurity' -Force -Scope CurrentUser -AllowClobber
            Write-Log "Successfully installed DuoSecurity" -Level Success
        }
        catch {
            Write-Log "Failed to install DuoSecurity: $_" -Level Error
            return $false
        }
    }

    try {
        Import-Module 'DuoSecurity' -Force -ErrorAction Stop
        Write-Log "DuoSecurity module loaded" -Level Success
        return $true
    }
    catch {
        Write-Log "Failed to import DuoSecurity: $_" -Level Error
        return $false
    }
}

#endregion

#region Duo Functions

function Get-BypassUsersForSubclient {
    <#
    .SYNOPSIS
        Returns all bypass users for the currently selected Duo subaccount.
        Pagination is handled internally by the DuoSecurity module.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SubclientName
    )

    $bypassUsers = [System.Collections.Generic.List[PSCustomObject]]::new()

    try {
        # Get-DuoUsers handles all pagination internally via Invoke-DuoPaginatedRequest.
        # It does not accept -Limit or -Offset parameters.
        $allUsers = @(Get-DuoUsers -ErrorAction Stop)
    }
    catch {
        Write-Log "    Error retrieving users for '$SubclientName': $_" -Level Error
        Write-Output -NoEnumerate $bypassUsers
        return
    }

    foreach ($user in $allUsers) {
        if ($user.status -eq 'bypass') {
            $bypassUsers.Add([PSCustomObject]@{
                Subclient  = $SubclientName
                Username   = $user.username
                RealName   = $user.realname
                Email      = $user.email
                Status     = $user.status
                IsEnrolled = $user.is_enrolled
                UserId     = $user.user_id
                LastLogin  = $user.last_login
            })
        }
    }

    Write-Output -NoEnumerate $bypassUsers
}

function Get-AllBypassUsers {
    <#
    .SYNOPSIS
        Iterates every Duo subclient and collects users in bypass status.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials
    )

    # Authenticate to Accounts API
    try {
        $duoParams = $Credentials.Duo
        Set-DuoApiAuth @duoParams
        Write-Log "Authenticated to Duo Accounts API" -Level Success
    }
    catch {
        Write-Log "Failed to authenticate to Duo Accounts API: $_" -Level Error
        return $null
    }

    # Retrieve all subclients
    Write-Log "Retrieving subclient list..." -Level Info
    try {
        $subaccounts = @(Get-DuoAccounts)
    }
    catch {
        Write-Log "Failed to retrieve subaccounts: $_" -Level Error
        return $null
    }

    if ($subaccounts.Count -eq 0) {
        Write-Log "No subaccounts found" -Level Warning
        return @()
    }

    # If Get-DuoAccounts returned an API error response instead of account objects,
    # the result is a single object with a 'stat' property (e.g. stat='FAIL').
    if ($subaccounts[0] | Get-Member -Name 'stat' -MemberType NoteProperty -ErrorAction SilentlyContinue) {
        Write-Log "Accounts API error: $($subaccounts[0].stat) - $($subaccounts[0].message)" -Level Error
        Write-Log "Verify that credentials.xml contains valid Accounts API (MSP-level) keys, not Admin API keys." -Level Error
        return $null
    }

    Write-Log "Found $($subaccounts.Count) subclient(s)" -Level Info

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($account in $subaccounts) {
        # Skip the parent/MSP account — it has no name and cannot be queried via Admin API
        if ([string]::IsNullOrWhiteSpace($account.name)) {
            Write-Log "Skipping account with no name (account_id: $($account.account_id))" -Level Warning
            continue
        }

        Write-Log "Querying: $($account.name)" -Level Info

        try {
            # Switch Admin API context to this subclient
            Select-DuoAccount -Name $account.name

            $bypassUsers = Get-BypassUsersForSubclient -SubclientName $account.name

            if ($bypassUsers.Count -gt 0) {
                Write-Log "  $($bypassUsers.Count) bypass user(s) found" -Level Warning
                $results.AddRange($bypassUsers)
            }
            else {
                Write-Log "  No bypass users" -Level Success
            }
        }
        catch {
            Write-Log "  Error querying '$($account.name)': $_" -Level Error
        }
    }

    # Write-Output -NoEnumerate prevents PowerShell from enumerating the List through
    # the pipeline, which would make an empty collection appear as $null to the caller.
    Write-Output -NoEnumerate $results
}

#endregion

#region ConnectWise Functions

function Search-ConnectWiseCompany {
    <#
    .SYNOPSIS
        Searches for a company in ConnectWise Manage by name.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SearchTerm
    )

    Write-Log "Searching ConnectWise for '$SearchTerm'..." -Level Info

    try {
        $companies = Get-CWMCompany -condition "name contains '$SearchTerm' and status/id = 1 and deletedFlag = false" -all |
            Sort-Object -Property id -Unique

        if (-not $companies -or $companies.Count -eq 0) {
            Write-Log "No companies found matching '$SearchTerm'" -Level Warning
            return $null
        }

        if ($companies.Count -eq 1) {
            Write-Log "Found: $($companies[0].name)" -Level Success
            return $companies[0]
        }

        # Multiple matches - present selection menu
        Write-Host "`nMultiple companies found:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $companies.Count; $i++) {
            Write-Host "  [$($i + 1)] $($companies[$i].name)"
        }
        Write-Host "  [0] Cancel"

        do {
            $selection = Read-Host "`nSelect company (1-$($companies.Count))"
            $selectionInt = 0
            [int]::TryParse($selection, [ref]$selectionInt) | Out-Null
        } while ($selectionInt -lt 0 -or $selectionInt -gt $companies.Count)

        if ($selectionInt -eq 0) {
            Write-Log "Selection cancelled" -Level Warning
            return $null
        }

        $selected = $companies[$selectionInt - 1]
        Write-Log "Selected: $($selected.name)" -Level Success
        return $selected
    }
    catch {
        Write-Log "Error searching ConnectWise: $_" -Level Error
        return $null
    }
}

function New-BypassTicketForCompany {
    <#
    .SYNOPSIS
        Creates a single ConnectWise ticket under the specified company.
        Optionally attaches a CSV file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$CompanyId,

        [Parameter(Mandatory)]
        [string]$CompanyName,

        [Parameter(Mandatory)]
        $BypassUsers,   # List/array of PSCustomObject

        [Parameter()]
        [string]$CsvPath,

        [Parameter()]
        [string]$Summarysuffix
    )

    $auditDate = Get-Date -Format "yyyy-MM-dd HH:mm"
    $summary   = "Duo Bypass Users Audit"
    if ($Summarysuffix) { $summary += " — $Summarysuffix" }

    $desc  = "Duo Bypass Users Audit — $auditDate`n`n"
    $desc += "Total bypass users: $($BypassUsers.Count)`n`n"

    $grouped = @($BypassUsers) | Group-Object Subclient | Sort-Object Name
    foreach ($group in $grouped) {
        $desc += "$($group.Name): $($group.Count) user(s)`n"
        foreach ($user in $group.Group | Sort-Object Username) {
            $enrolledStr = if ($user.IsEnrolled) { 'enrolled' } else { 'NOT enrolled' }
            $emailStr    = if ($user.Email) { " <$($user.Email)>" } else { '' }
            $desc += "  - $($user.Username)$emailStr [$enrolledStr]`n"
        }
        $desc += "`n"
    }

    $ticket = $null
    try {
        $ticket = New-CWMTicket `
            -summary  $summary `
            -company  @{ id = $CompanyId } `
            -initialDescription $desc `
            -ErrorAction Stop
        Write-Log "CW Ticket #$($ticket.id) created for '$CompanyName': '$($ticket.summary)'" -Level Success
    }
    catch {
        Write-Log "Failed to create ticket for '$CompanyName': $_" -Level Error
        return
    }

    # Attach CSV if provided
    if ($CsvPath -and (Test-Path $CsvPath)) {
        try {
            New-CWMDocument `
                -recordType 'Ticket' `
                -recordId   $ticket.id `
                -title      'Duo Bypass Users Report' `
                -FilePath   $CsvPath `
                -Private    $false `
                -ErrorAction Stop
            Write-Log "CSV attached to ticket #$($ticket.id)" -Level Success
        }
        catch {
            Write-Log "Ticket created but CSV attachment failed: $_" -Level Warning
        }
    }
}

function Invoke-BypassTicketCreation {
    <#
    .SYNOPSIS
        Orchestrates ConnectWise ticket creation for bypass user audit results.
        Offers per-client or single-company ticketing modes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials,

        [Parameter(Mandatory)]
        [string]$CsvPath,

        [Parameter(Mandatory)]
        $BypassUsers
    )

    # ConnectWise credentials are optional — skip gracefully if absent
    if (-not $Credentials.ConnectWise) {
        Write-Log "No ConnectWise credentials found in credentials.xml — skipping ticket creation." -Level Warning
        return
    }

    # Install CWM module if needed
    if (-not (Get-Module -ListAvailable -Name 'ConnectWiseManageAPI')) {
        Write-Log "ConnectWiseManageAPI module not found. Installing..." -Level Warning
        try {
            Install-Module -Name 'ConnectWiseManageAPI' -Scope CurrentUser -Force -AllowClobber
            Write-Log "ConnectWiseManageAPI installed" -Level Success
        }
        catch {
            Write-Log "Failed to install ConnectWiseManageAPI: $_  Skipping ticket creation." -Level Warning
            return
        }
    }

    try {
        Import-Module 'ConnectWiseManageAPI' -Force -ErrorAction Stop
    }
    catch {
        Write-Log "Failed to import ConnectWiseManageAPI: $_  Skipping ticket creation." -Level Warning
        return
    }

    # Connect
    try {
        $cwParams = $Credentials.ConnectWise
        Connect-CWM @cwParams -ErrorAction Stop
        Write-Log "Connected to ConnectWise" -Level Info
    }
    catch {
        Write-Log "Failed to connect to ConnectWise: $_  Skipping ticket creation." -Level Warning
        return
    }

    # Prompt for ticket mode
    Write-Host ""
    Write-Host "How should ConnectWise tickets be created?" -ForegroundColor Cyan
    Write-Host "  [1] One ticket per client (match Duo subclient to CW company)"
    Write-Host "  [2] All under one company"
    Write-Host "  [0] Skip"

    do {
        $mode = Read-Host "`nSelect option (0-2)"
    } while ($mode -notin '0', '1', '2')

    if ($mode -eq '0') {
        Write-Log "Skipped ConnectWise ticket creation." -Level Info
        return
    }

    if ($mode -eq '2') {
        # Single-company mode
        $searchTerm = Read-Host "`nEnter company name to search"
        $company = Search-ConnectWiseCompany -SearchTerm $searchTerm
        if (-not $company) {
            Write-Log "No company selected — skipping ticket creation." -Level Warning
            return
        }

        New-BypassTicketForCompany `
            -CompanyId   $company.id `
            -CompanyName $company.name `
            -BypassUsers $BypassUsers `
            -CsvPath     $CsvPath
        return
    }

    # Per-client mode
    $grouped = @($BypassUsers) | Group-Object Subclient | Sort-Object Name
    $ticketsCreated = 0

    foreach ($group in $grouped) {
        $subclientName = $group.Name
        Write-Host "`n--- $subclientName ($($group.Count) bypass user(s)) ---" -ForegroundColor Yellow

        # Try to auto-match
        $company = Search-ConnectWiseCompany -SearchTerm $subclientName

        while (-not $company) {
            Write-Host "`nNo match for '$subclientName'." -ForegroundColor Yellow
            Write-Host "  [1] Search with a different term"
            Write-Host "  [0] Skip this client"

            $choice = Read-Host "Select option"
            if ($choice -eq '0') {
                Write-Log "Skipped ticket for '$subclientName'" -Level Info
                break
            }

            $altTerm = Read-Host "Enter search term"
            $company = Search-ConnectWiseCompany -SearchTerm $altTerm
        }

        if ($company) {
            New-BypassTicketForCompany `
                -CompanyId     $company.id `
                -CompanyName   $company.name `
                -BypassUsers   $group.Group `
                -Summarysuffix $subclientName
            $ticketsCreated++
        }
    }

    Write-Log "$ticketsCreated ticket(s) created across $($grouped.Count) subclient(s)." -Level Info
}

#endregion

#region Main Execution

function Invoke-BypassUserReport {
    [CmdletBinding()]
    param(
        [string]$CsvOutputPath,
        [switch]$ResetCreds
    )

    Write-Host "`n" -NoNewline
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "  Duo Bypass User Report" -ForegroundColor Cyan
    Write-Host "  Code Blue Technology" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""

    if (-not (Test-Prerequisites)) {
        Write-Log "Prerequisites check failed. Exiting." -Level Error
        return
    }

    $credentials = Get-StoredCredentials -Reset:$ResetCreds
    if (-not $credentials) {
        Write-Log "Failed to obtain credentials. Exiting." -Level Error
        return
    }

    $bypassUsers = Get-AllBypassUsers -Credentials $credentials

    if ($null -eq $bypassUsers) {
        Write-Log "Report aborted due to errors." -Level Error
        return
    }

    Write-Host ""
    Write-Host "==========================================" -ForegroundColor $(if ($bypassUsers.Count -gt 0) { 'Yellow' } else { 'Green' })

    if ($bypassUsers.Count -eq 0) {
        Write-Host "  No bypass users found across any subclient." -ForegroundColor Green
        Write-Host "==========================================" -ForegroundColor Green
        return
    }

    Write-Host "  $($bypassUsers.Count) bypass user(s) found" -ForegroundColor Yellow
    Write-Host "==========================================" -ForegroundColor Yellow
    Write-Host ""

    # Display grouped summary in console
    $grouped = $bypassUsers | Group-Object -Property Subclient | Sort-Object Name
    foreach ($group in $grouped) {
        Write-Host "  $($group.Name)  ($($group.Count) user(s))" -ForegroundColor Yellow
        foreach ($user in $group.Group | Sort-Object Username) {
            $enrolledTag = if ($user.IsEnrolled) { 'enrolled' } else { 'NOT enrolled' }
            $emailPart   = if ($user.Email) { " <$($user.Email)>" } else { '' }
            Write-Host "    - $($user.Username)$emailPart  [$enrolledTag]" -ForegroundColor White
        }
        Write-Host ""
    }

    # Determine CSV path
    if (-not $CsvOutputPath) {
        $timestamp     = Get-Date -Format "yyyyMMdd-HHmmss"
        $CsvOutputPath = Join-Path $PSScriptRoot "DuoBypassUsers-$timestamp.csv"
    }

    try {
        $bypassUsers |
            Sort-Object Subclient, Username |
            Export-Csv -Path $CsvOutputPath -NoTypeInformation -Force
        Write-Log "Report saved to: $CsvOutputPath" -Level Success
    }
    catch {
        Write-Log "Failed to write CSV: $_" -Level Error
        return
    }

    # ConnectWise ticket creation
    Invoke-BypassTicketCreation -Credentials $credentials -CsvPath $CsvOutputPath -BypassUsers $bypassUsers
}

# Entry point
Invoke-BypassUserReport -CsvOutputPath $OutputPath -ResetCreds:$ResetCredentials

#endregion
