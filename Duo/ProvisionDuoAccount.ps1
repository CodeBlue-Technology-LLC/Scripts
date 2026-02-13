#Requires -Version 5.1

<#
.SYNOPSIS
    Provisions Duo Security subaccounts for ConnectWise Manage companies.

.DESCRIPTION
    This script integrates Duo MSP with ConnectWise Manage to automate
    subaccount creation and configuration for MSP customers.

    Features:
    - Searches ConnectWise Manage for company information
    - Creates Duo subaccount with company name
    - Applies Code Blue Technology standard settings

.PARAMETER CompanyName
    The name of the company to search for in ConnectWise. If not provided,
    the script will prompt for input.

.PARAMETER SkipSettings
    Skip the automatic configuration of Duo settings.

.PARAMETER ResetCredentials
    Force re-entry of stored credentials. Use this if your API keys have changed.

.EXAMPLE
    .\Duo.ps1
    Runs interactively, prompting for company name.

.EXAMPLE
    .\Duo.ps1 -CompanyName "Contoso Ltd"
    Searches for "Contoso Ltd" in ConnectWise and provisions Duo account.

.EXAMPLE
    .\Duo.ps1 -ResetCredentials
    Clears stored credentials and prompts for new API keys.

.NOTES
    Author: Code Blue Technology
    Requires: DuoSecurity, ConnectWiseManageAPI PowerShell modules
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$CompanyName,

    [Parameter()]
    [switch]$SkipSettings,

    [Parameter()]
    [switch]$ResetCredentials
)

#region Configuration

# Credential storage path (DPAPI encrypted - only you on this machine can decrypt)
$script:ConfigPath = Join-Path $PSScriptRoot "Config"
$script:CredentialFile = Join-Path $script:ConfigPath "credentials.xml"

# Standard Duo settings
$StandardSettings = @{
    Timezone                        = 'America/New_York'
    EnrollmentUniversalPromptEnabled = $true
    HelpdeskCanSendEnrollEmail      = $true
}

#endregion

#region Credential Management

function Get-StoredCredentials {
    <#
    .SYNOPSIS
        Retrieves or prompts for API credentials, storing them securely with DPAPI.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$Reset
    )

    # Ensure config directory exists
    if (-not (Test-Path $script:ConfigPath)) {
        New-Item -ItemType Directory -Path $script:ConfigPath -Force | Out-Null
    }

    # Check for existing credentials
    if ((Test-Path $script:CredentialFile) -and (-not $Reset)) {
        try {
            $creds = Import-Clixml -Path $script:CredentialFile
            Write-Log "Loaded stored credentials" -Level Info
            return $creds
        }
        catch {
            Write-Log "Failed to load stored credentials, will prompt for new ones" -Level Warning
        }
    }

    # Prompt for new credentials
    Write-Host "`n=== Duo MSP Credentials ===" -ForegroundColor Cyan
    Write-Host "These credentials are for your Duo Accounts API (MSP level)`n"

    $duoApiHost = Read-Host "Duo API Host (e.g., api-xxxxxxxx.duosecurity.com)"
    $duoIntegrationKey = Read-Host "Duo Integration Key"
    $duoSecretKey = Read-Host "Duo Secret Key"

    Write-Host "`n=== ConnectWise Manage Credentials ===" -ForegroundColor Cyan
    Write-Host "These credentials are for your ConnectWise Manage API`n"

    $cwmServer = Read-Host "CWM Server (e.g., na.myconnectwise.net)"
    $cwmCompany = Read-Host "CWM Company ID"
    $cwmClientId = Read-Host "CWM Client ID"
    $cwmPublicKey = Read-Host "CWM Public Key"
    $cwmPrivateKey = Read-Host "CWM Private Key"

    $credentials = @{
        Duo = @{
            Type           = 'Accounts'
            ApiHost        = $duoApiHost
            IntegrationKey = $duoIntegrationKey
            SecretKey      = $duoSecretKey
        }
        ConnectWise = @{
            Server     = $cwmServer
            Company    = $cwmCompany
            clientId   = $cwmClientId
            pubKey     = $cwmPublicKey
            privateKey = $cwmPrivateKey
        }
    }

    # Save credentials (encrypted with DPAPI)
    try {
        $credentials | Export-Clixml -Path $script:CredentialFile -Force
        Write-Log "Credentials saved securely to $script:CredentialFile" -Level Success
    }
    catch {
        Write-Log "Failed to save credentials: $_" -Level Warning
    }

    return $credentials
}

#endregion

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

#region Prerequisites

function Test-Prerequisites {
    <#
    .SYNOPSIS
        Checks and installs required PowerShell modules.
    #>
    [CmdletBinding()]
    param()

    $requiredModules = @('DuoSecurity', 'ConnectWiseManageAPI')
    $allPresent = $true

    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Log "Module '$module' not found. Installing..." -Level Warning
            try {
                Install-Module -Name $module -Force -Scope CurrentUser -AllowClobber
                Write-Log "Successfully installed $module" -Level Success
            }
            catch {
                Write-Log "Failed to install $module : $_" -Level Error
                $allPresent = $false
            }
        }
    }

    # Import modules
    foreach ($module in $requiredModules) {
        try {
            Import-Module $module -Force -ErrorAction Stop
        }
        catch {
            Write-Log "Failed to import $module : $_" -Level Error
            $allPresent = $false
        }
    }

    if ($allPresent) {
        Write-Log "All prerequisites satisfied" -Level Success
    }

    return $allPresent
}

#endregion

#region Connection Functions

function Connect-APIs {
    <#
    .SYNOPSIS
        Connects to both Duo and ConnectWise APIs using stored credentials.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials
    )

    # Connect to ConnectWise
    try {
        $cwmParams = $Credentials.ConnectWise
        Connect-CWM @cwmParams
        Write-Log "Connected to ConnectWise Manage" -Level Success
    }
    catch {
        Write-Log "Failed to connect to ConnectWise: $_" -Level Error
        return $false
    }

    # Connect to Duo Accounts API
    try {
        $duoParams = $Credentials.Duo
        Set-DuoApiAuth @duoParams
        Write-Log "Connected to Duo MSP (Accounts API)" -Level Success
    }
    catch {
        Write-Log "Failed to connect to Duo: $_" -Level Error
        return $false
    }

    return $true
}

#endregion

#region ConnectWise Functions

function Search-ConnectWiseCompany {
    <#
    .SYNOPSIS
        Searches for a company in ConnectWise Manage.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SearchTerm
    )

    Write-Log "Searching ConnectWise for '$SearchTerm'..." -Level Info

    try {
        $companies = Get-CWMCompany -condition "name contains '$SearchTerm'" -all

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

#endregion

#region Duo Functions

function New-DuoCustomerSubaccount {
    <#
    .SYNOPSIS
        Creates a new Duo subaccount for a customer.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccountName
    )

    Write-Log "Checking if Duo account '$AccountName' already exists..." -Level Info

    try {
        $existingAccounts = Get-DuoAccounts
        $existing = $existingAccounts | Where-Object { $_.name -eq $AccountName }

        if ($existing) {
            Write-Host "`nAccount '$AccountName' already exists in Duo!" -ForegroundColor Yellow
            Write-Host "  [1] Use existing account"
            Write-Host "  [2] Create new account with different name"
            Write-Host "  [0] Cancel"

            $choice = Read-Host "`nSelect option"

            switch ($choice) {
                '1' {
                    Write-Log "Using existing account: $AccountName" -Level Info
                    return $existing
                }
                '2' {
                    $newName = Read-Host "Enter new account name"
                    return New-DuoCustomerSubaccount -AccountName $newName
                }
                default {
                    Write-Log "Operation cancelled" -Level Warning
                    return $null
                }
            }
        }

        # Create new account
        Write-Log "Creating Duo subaccount: $AccountName" -Level Info
        $newAccount = New-DuoAccount -Name $AccountName

        if ($newAccount) {
            Write-Log "Successfully created Duo subaccount: $AccountName" -Level Success
            Write-Host "  Account ID: $($newAccount.account_id)" -ForegroundColor Gray
            Write-Host "  API Host: $($newAccount.api_hostname)" -ForegroundColor Gray
            return $newAccount
        }
        else {
            Write-Log "Failed to create Duo subaccount" -Level Error
            return $null
        }
    }
    catch {
        Write-Log "Error creating Duo subaccount: $_" -Level Error
        return $null
    }
}

function Set-DuoStandardSettings {
    <#
    .SYNOPSIS
        Applies Code Blue Technology standard settings to the Duo account.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccountName
    )

    Write-Log "Switching to subaccount context: $AccountName" -Level Info

    try {
        # Select the subaccount to work with Admin API
        Select-DuoAccount -Name $AccountName

        Write-Log "Applying standard settings..." -Level Info

        $settingsApplied = @()

        # Apply settings one at a time to handle any individual failures
        try {
            Update-DuoSettings -Timezone $StandardSettings.Timezone
            $settingsApplied += "Timezone: $($StandardSettings.Timezone)"
        }
        catch { Write-Log "Failed to set Timezone: $_" -Level Warning }

        try {
            Update-DuoSettings -EnrollmentUniversalPromptEnabled $StandardSettings.EnrollmentUniversalPromptEnabled
            $settingsApplied += "Universal Prompt: Enabled"
        }
        catch { Write-Log "Failed to set Universal Prompt: $_" -Level Warning }

        try {
            Update-DuoSettings -HelpdeskCanSendEnrollEmail $StandardSettings.HelpdeskCanSendEnrollEmail
            $settingsApplied += "Helpdesk Enrollment Email: Enabled"
        }
        catch { Write-Log "Failed to set Helpdesk Enrollment: $_" -Level Warning }

        try {
            Update-DuoSettings -FraudEmail $StandardSettings.FraudEmail -FraudEmailEnabled $StandardSettings.FraudEmailEnabled
            $settingsApplied += "Fraud Email: $($StandardSettings.FraudEmail)"
        }
        catch { Write-Log "Failed to set Fraud Email: $_" -Level Warning }

        if ($settingsApplied.Count -gt 0) {
            Write-Log "Applied settings:" -Level Success
            foreach ($setting in $settingsApplied) {
                Write-Host "  - $setting" -ForegroundColor Green
            }
        }

        return $true
    }
    catch {
        Write-Log "Error applying settings: $_" -Level Error
        return $false
    }
}

#endregion

#region Main Execution

function Invoke-DuoProvisioning {
    <#
    .SYNOPSIS
        Main orchestration function for Duo provisioning.
    #>
    [CmdletBinding()]
    param(
        [string]$Company,
        [switch]$SkipSettingsConfig,
        [switch]$ResetCreds
    )

    Write-Host "`n" -NoNewline
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "  Duo MSP Customer Provisioning Tool" -ForegroundColor Cyan
    Write-Host "  Code Blue Technology" -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host ""

    # Step 1: Check prerequisites
    if (-not (Test-Prerequisites)) {
        Write-Log "Prerequisites check failed. Exiting." -Level Error
        return
    }

    # Step 2: Get credentials
    $credentials = Get-StoredCredentials -Reset:$ResetCreds
    if (-not $credentials) {
        Write-Log "Failed to obtain credentials. Exiting." -Level Error
        return
    }

    # Step 3: Connect to APIs
    if (-not (Connect-APIs -Credentials $credentials)) {
        Write-Log "Failed to connect to APIs. Exiting." -Level Error
        return
    }

    # Step 4: Get company name if not provided
    if (-not $Company) {
        $Company = Read-Host "`nEnter Connectwise Company name"
    }

    # Step 5: Search ConnectWise
    $cwmCompany = Search-ConnectWiseCompany -SearchTerm $Company
    if (-not $cwmCompany) {
        $useManual = Read-Host "Would you like to enter a company name manually? (Y/N)"
        if ($useManual -eq 'Y') {
            $accountName = Read-Host "Enter company name for Duo account - should match CWM"
        }
        else {
            Write-Log "No company selected. Exiting." -Level Warning
            return
        }
    }
    else {
        $accountName = $cwmCompany.name
    }

    # Step 6: Confirm before creating
    Write-Host "`n--- Summary ---" -ForegroundColor Yellow
    Write-Host "Company Name: $accountName"
    if ($cwmCompany) {
        if ($cwmCompany.addressLine1) { Write-Host "Address: $($cwmCompany.addressLine1)" }
        if ($cwmCompany.city) { Write-Host "City: $($cwmCompany.city), $($cwmCompany.state) $($cwmCompany.zip)" }
    }
    Write-Host ""

    $confirm = Read-Host "Create Duo account for '$accountName'? (Y/N)"
    if ($confirm -ne 'Y') {
        Write-Log "Operation cancelled by user" -Level Warning
        return
    }

    # Step 7: Create Duo subaccount
    $duoAccount = New-DuoCustomerSubaccount -AccountName $accountName
    if (-not $duoAccount) {
        return
    }

    # Step 8: Configure settings
    if (-not $SkipSettingsConfig) {
        Set-DuoStandardSettings -AccountName $accountName
    }

    # Step 9: Final summary
    Write-Host "`n" -NoNewline
    Write-Host "=======================================" -ForegroundColor Green
    Write-Host "  Provisioning Complete!" -ForegroundColor Green
    Write-Host "=======================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Account: $accountName" -ForegroundColor White
    Write-Host "Account ID: $($duoAccount.account_id)" -ForegroundColor Gray
    Write-Host "API Host: $($duoAccount.api_hostname)" -ForegroundColor Gray
    Write-Host ""
    # Build admin portal URL (replace api- with admin-)
    $adminHost = $duoAccount.api_hostname -replace '^api-', 'admin-'
    $directorySyncUrl = "https://$adminHost/users/directorysync"

    Write-Host "Next Steps:" -ForegroundColor Yellow
    Write-Host "  1. Configure directory sync: $directorySyncUrl" -ForegroundColor White
    Write-Host "  2. Add applications/integrations as needed" -ForegroundColor White
    Write-Host "  3. Send enrollment links to users" -ForegroundColor White
    Write-Host ""
}

# Run the main function
Invoke-DuoProvisioning -Company $CompanyName -SkipSettingsConfig:$SkipSettings -ResetCreds:$ResetCredentials

#endregion
