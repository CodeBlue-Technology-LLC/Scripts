<#
.SYNOPSIS
    Migrate DNS records from Bluehost (WHM/cPanel) to Cloudflare.

.DESCRIPTION
    Interactive migration tool that transfers DNS records from a Bluehost
    WHM reseller account to Cloudflare. Looks up the company in ConnectWise
    Manage by matching the domain against contact email domains, creates a
    Cloudflare subaccount, imports DNS records, and documents in ITGlue.

    Nameserver changes at Bluehost must be done manually (no registrar API).

.PARAMETER Domain
    The domain to migrate (e.g., "example.com").

.PARAMETER CustomerName
    Name for the Cloudflare account/subtenant. If omitted, uses the
    ConnectWise company name found via email domain lookup.

.PARAMETER AccountId
    Use an existing Cloudflare account instead of creating a new one.

.PARAMETER ListDomains
    List domains on the WHM server whose nameservers point to Bluehost (ready to migrate).

.PARAMETER ListDomainsAll
    List all domains on the WHM server with nameserver status (Bluehost, Cloudflare, GoDaddy, Network Solutions, Other).

.PARAMETER ListCloudflareAccounts
    List all accounts in Cloudflare (read-only).

.PARAMETER ListCloudflareZones
    List all zones in Cloudflare (read-only).

.PARAMETER PreviewRecords
    Preview DNS records for a domain without making changes.

.EXAMPLE
    .\Bluehost-Cloudflare-Migrate-DNS.ps1 -ListDomains

.EXAMPLE
    .\Bluehost-Cloudflare-Migrate-DNS.ps1 -ListDomainsAll

.EXAMPLE
    .\Bluehost-Cloudflare-Migrate-DNS.ps1 -PreviewRecords "example.com"

.EXAMPLE
    .\Bluehost-Cloudflare-Migrate-DNS.ps1 -Domain "example.com"

.EXAMPLE
    .\Bluehost-Cloudflare-Migrate-DNS.ps1 -Domain "example.com" -CustomerName "Acme Corp"
#>
[CmdletBinding(DefaultParameterSetName = 'Migrate')]
param(
    [Parameter(ParameterSetName = 'Migrate')]
    [Parameter(ParameterSetName = 'Preview', Mandatory)]
    [string]$Domain,

    [Parameter(ParameterSetName = 'Migrate')]
    [string]$CustomerName,

    [Parameter(ParameterSetName = 'Migrate')]
    [string]$AccountId,

    [Parameter(ParameterSetName = 'ListDomains')]
    [switch]$ListDomains,

    [Parameter(ParameterSetName = 'ListDomainsAll')]
    [switch]$ListDomainsAll,

    [Parameter(ParameterSetName = 'ListCFAccounts')]
    [switch]$ListCloudflareAccounts,

    [Parameter(ParameterSetName = 'ListCFZones')]
    [switch]$ListCloudflareZones,

    [Parameter(ParameterSetName = 'Preview')]
    [switch]$PreviewRecords,

    [Parameter(ParameterSetName = 'ResetHost')]
    [switch]$ResetHostCredentials
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Import modules
Import-Module "$ScriptDir\Modules\WHM.psm1" -Force
Import-Module "$ScriptDir\Modules\Cloudflare.psm1" -Force

# Check for ConnectWiseManageAPI module
if (-not (Get-Module -ListAvailable -Name ConnectWiseManageAPI)) {
    Write-Host "ConnectWiseManageAPI module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name ConnectWiseManageAPI -Scope CurrentUser -Force -AllowClobber
        Write-Host "ConnectWiseManageAPI installed successfully." -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to install ConnectWiseManageAPI: $($_.Exception.Message)"
        Write-Warning "Company lookup and ticket creation will be skipped."
    }
}

# Check for ITGlueAPI module
if (-not (Get-Module -ListAvailable -Name ITGlueAPI)) {
    Write-Host "ITGlueAPI module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name ITGlueAPI -Scope CurrentUser -Force -AllowClobber
        Write-Host "ITGlueAPI installed successfully." -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to install ITGlueAPI: $($_.Exception.Message)"
        Write-Warning "ITGlue asset creation will be skipped."
    }
}

# Credential management
$CredentialsPath = "$ScriptDir\Config\credentials.xml"

function Initialize-Credentials {
    <#
    .SYNOPSIS
        Loads or prompts for API credentials, storing them securely with DPAPI.
    #>
    param([switch]$Reset)

    $configDir = Split-Path $CredentialsPath -Parent
    if (-not (Test-Path $configDir)) {
        New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    }

    $creds = $null
    if ((Test-Path $CredentialsPath) -and (-not $Reset)) {
        try {
            $creds = Import-Clixml -Path $CredentialsPath
        }
        catch {
            Write-Warning "Failed to load stored credentials, will prompt for new ones."
        }
    }

    $needsSave = $false

    # --- Cloudflare ---
    if (-not $creds -or -not $creds.Cloudflare) {
        Write-Host "`n=== Cloudflare Credentials ===" -ForegroundColor Cyan
        $cfEmail = Read-Host "Cloudflare Email"
        $cfApiKey = Read-Host "Cloudflare API Key"
        if (-not $creds) { $creds = @{} }
        $creds.Cloudflare = @{ Email = $cfEmail; ApiKey = $cfApiKey }
        $needsSave = $true
    }

    # --- WHM/cPanel/Bluehost ---
    if (-not $creds.WHM) {
        Write-Host "`n=== Bluehost/cPanel Connection ===" -ForegroundColor Cyan
        Write-Host "  [1] WHM (reseller account - port 2087)" -ForegroundColor White
        Write-Host "  [2] cPanel (single account - port 2083)" -ForegroundColor White
        $connType = Read-Host "Connection type (1 or 2)"
        $type = if ($connType -eq '2') { 'cPanel' } else { 'WHM' }

        $whmServer = Read-Host "Server hostname (e.g., server123.bluehost.com)"
        $whmUser = Read-Host "Username"
        $whmToken = Read-Host "API Token"
        $creds.WHM = @{
            Server      = $whmServer
            Username    = $whmUser
            AccessToken = $whmToken
            Type        = $type
        }
        $needsSave = $true
    }

    # --- ConnectWise ---
    if (-not $creds.ConnectWise) {
        Write-Host "`n=== ConnectWise Manage Credentials ===" -ForegroundColor Cyan
        $cwServer = Read-Host "CWM Server (e.g., na.myconnectwise.net)"
        $cwCompany = Read-Host "CWM Company ID"
        $cwClientId = Read-Host "CWM Client ID"
        $cwPubKey = Read-Host "CWM Public Key"
        $cwPrivKey = Read-Host "CWM Private Key"
        $creds.ConnectWise = @{
            Server     = $cwServer
            Company    = $cwCompany
            clientId   = $cwClientId
            pubKey     = $cwPubKey
            privateKey = $cwPrivKey
        }
        $needsSave = $true
    }

    # --- ITGlue ---
    if (-not $creds.ITGlue) {
        Write-Host "`n=== ITGlue Credentials ===" -ForegroundColor Cyan
        $creds.ITGlue = @{}
        $needsSave = $true
    }

    # ITGlue Base URI
    if (-not $creds.ITGlue.BaseUri) {
        $itgBaseUri = Read-Host "ITGlue API Base URL [https://api.itglue.com]"
        if ([string]::IsNullOrWhiteSpace($itgBaseUri)) { $itgBaseUri = 'https://api.itglue.com' }
        $creds.ITGlue.BaseUri = $itgBaseUri
        $needsSave = $true
    }

    # ITGlue API Key
    if (-not $creds.ITGlue.ApiKey) {
        $creds.ITGlue.ApiKey = Read-Host "ITGlue API Key"
        $needsSave = $true
    }

    # ITGlue Subdomain
    if (-not $creds.ITGlue.Subdomain) {
        $creds.ITGlue.Subdomain = Read-Host "ITGlue Subdomain (e.g., yourcompany)"
        $needsSave = $true
    }

    # ITGlue DNS Credential Password Asset ID
    if (-not $creds.ITGlue.DnsCredentialAssetId) {
        $itgAssetId = Read-Host "ITGlue Password Asset ID for DNS credentials (or Enter to skip)"
        if (-not [string]::IsNullOrWhiteSpace($itgAssetId)) {
            $creds.ITGlue.DnsCredentialAssetId = [int]$itgAssetId
        }
        $needsSave = $true
    }

    # ITGlue Flexible Asset Type ID
    if (-not $creds.ITGlue.FlexibleAssetTypeId) {
        $existingId = $creds.ITGlue.FlexibleAssetTypeId
        $promptSuffix = if ($existingId) { " [$existingId]" } else { "" }

        Write-Host "`n=== ITGlue Flexible Asset Type (DNS/Registrar) ===" -ForegroundColor Cyan
        $configChoice = Read-Host "Configure DNS/Registrar flexible asset type? (Y/n)"
        if ($configChoice -ne 'n' -and $configChoice -ne 'N') {
            try {
                Import-Module ITGlueAPI -Force -ErrorAction Stop
                Add-ITGlueBaseURI -base_uri $creds.ITGlue.BaseUri
                Add-ITGlueAPIKey -Api_Key $creds.ITGlue.ApiKey

                $fatypes = Get-ITGlueFlexibleAssetTypes
                if ($fatypes.data) {
                    Write-Host "`nExisting flexible asset types:" -ForegroundColor Yellow
                    $i = 1
                    foreach ($fat in $fatypes.data) {
                        Write-Host "  [$i] $($fat.attributes.name) (ID: $($fat.id))" -ForegroundColor White
                        $i++
                    }
                    Write-Host "  [C] Create new DNS/Registrar type" -ForegroundColor Cyan
                    Write-Host ""

                    $selection = Read-Host "Select a type by number, or C to create new$promptSuffix"

                    if ([string]::IsNullOrWhiteSpace($selection) -and $existingId) {
                        # Keep existing
                    }
                    elseif ($selection -eq 'C' -or $selection -eq 'c') {
                        $creds.ITGlue.FlexibleAssetTypeId = New-DnsRegistrarAssetType
                        $needsSave = $true
                    }
                    elseif ($selection -match '^\d+$') {
                        $idx = [int]$selection - 1
                        if ($idx -ge 0 -and $idx -lt $fatypes.data.Count) {
                            $selected = $fatypes.data[$idx]
                            $creds.ITGlue.FlexibleAssetTypeId = [int]$selected.id
                            $needsSave = $true

                            $fields = Get-ITGlueFlexibleAssetFields -flexible_asset_type_id $selected.id
                            $requiredTraits = @('provider', 'dns', 'registrar', 'dns-credentials', 'management-url', 'notes', 'domain-s')
                            $existingTraits = $fields.data | ForEach-Object { $_.attributes.'name-key' }
                            $missing = $requiredTraits | Where-Object { $_ -notin $existingTraits }
                            if ($missing) {
                                Write-Warning "Selected type is missing these trait fields: $($missing -join ', ')"
                                Write-Warning "ITGlue asset creation may not work correctly without them."
                            }
                            else {
                                Write-Host "All required fields present." -ForegroundColor Green
                            }
                        }
                    }
                }
            }
            catch {
                Write-Warning "Could not connect to ITGlue: $($_.Exception.Message)"
                $manualId = Read-Host "Enter Flexible Asset Type ID manually (or Enter to skip)"
                if (-not [string]::IsNullOrWhiteSpace($manualId)) {
                    $creds.ITGlue.FlexibleAssetTypeId = [int]$manualId
                    $needsSave = $true
                }
            }
        }
    }

    if ($needsSave) {
        try {
            $creds | Export-Clixml -Path $CredentialsPath -Force
            Write-Host "`nCredentials saved to $CredentialsPath" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to save credentials: $($_.Exception.Message)"
        }
    }

    return $creds
}

function New-DnsRegistrarAssetType {
    <#
    .SYNOPSIS
        Creates a new DNS/Registrar flexible asset type in ITGlue with required fields.
    #>
    Write-Host "Creating DNS/Registrar flexible asset type..." -ForegroundColor Yellow

    $typeData = @{
        type = 'flexible-asset-types'
        attributes = @{
            name = 'DNS / Registrar'
            icon = 'globe'
            description = 'DNS and domain registrar information'
            'show-in-menu' = $true
        }
    }

    $newType = New-ITGlueFlexibleAssetTypes -data $typeData
    $typeId = $newType.data.id
    Write-Host "Created flexible asset type: $($newType.data.attributes.name) (ID: $typeId)" -ForegroundColor Green

    $fields = @(
        @{ name = 'Provider'; kind = 'Text'; required = $true; 'show-in-list' = $true }
        @{ name = 'DNS'; kind = 'Checkbox'; required = $false; 'show-in-list' = $true }
        @{ name = 'Registrar'; kind = 'Checkbox'; required = $false; 'show-in-list' = $true }
        @{ name = 'DNS Credentials'; kind = 'Password'; required = $false; 'show-in-list' = $false }
        @{ name = 'Management URL'; kind = 'Text'; required = $false; 'show-in-list' = $true }
        @{ name = 'Notes'; kind = 'Textbox'; required = $false; 'show-in-list' = $false }
        @{ name = 'Domain(s)'; kind = 'Tag'; 'tag-type' = 'Domains'; required = $false; 'show-in-list' = $true }
    )

    $order = 1
    foreach ($field in $fields) {
        $fieldData = @{
            type = 'flexible-asset-fields'
            attributes = @{
                order                    = $order
                name                     = $field.name
                kind                     = $field.kind
                required                 = $field.required
                'show-in-list'           = $field.'show-in-list'
                'use-for-title'          = ($order -eq 1)
                'flexible-asset-type-id' = [int]$typeId
            }
        }
        if ($field.'tag-type') {
            $fieldData.attributes.'tag-type' = $field.'tag-type'
        }
        New-ITGlueFlexibleAssetFields -data $fieldData | Out-Null
        Write-Host "  Created field: $($field.name)" -ForegroundColor Cyan
        $order++
    }

    Write-Host "Flexible asset type setup complete." -ForegroundColor Green
    return [int]$typeId
}

# Load credentials
$Credentials = Initialize-Credentials
$WHMCreds = $Credentials.WHM
$CloudflareCreds = $Credentials.Cloudflare
$ConnectWiseCreds = $Credentials.ConnectWise
$ITGlueCreds = $Credentials.ITGlue

function Get-AllCloudflareAccounts {
    <#
    .SYNOPSIS
        Retrieves all Cloudflare accounts with pagination (default API page size is 20).
    #>
    $headers = @{
        "X-Auth-Email" = $CloudflareCreds.Email
        "X-Auth-Key"   = $CloudflareCreds.ApiKey
        "Content-Type" = "application/json"
    }
    $allAccounts = @()
    $page = 1
    do {
        $response = Invoke-RestMethod -Uri "https://api.cloudflare.com/client/v4/accounts?page=$page&per_page=50" `
            -Headers $headers -Method Get
        if ($response.success -and $response.result) {
            $allAccounts += $response.result
        }
        $page++
    } while ($response.result -and $response.result.Count -eq 50)
    return $allAccounts
}

function Write-Banner {
    param([string]$Text)
    Write-Host ""
    Write-Host ("=" * 60) -ForegroundColor Cyan
    Write-Host " $Text" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan
    Write-Host ""
}

function Confirm-Action {
    param(
        [string]$Message,
        [string]$Caption = "Confirm"
    )

    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Proceed with the action."
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Cancel the action."
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
    $result = $Host.UI.PromptForChoice($Caption, $Message, $options, 0)
    return $result -eq 0
}

function Find-CWCompanyByDomain {
    <#
    .SYNOPSIS
        Looks up a ConnectWise Manage company by matching a domain against
        contact email addresses, then falls back to company website field.
    .PARAMETER Domain
        The domain to search for (e.g., "example.com").
    .RETURNS
        The matching CW company object, or $null if none found.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Domain
    )

    $matchingCompany = $null

    # Strategy 1: Search contacts by email domain
    Write-Host "Searching ConnectWise contacts with email domain '@$Domain'..." -ForegroundColor Yellow
    try {
        $contacts = @(Get-CWMContact -childConditions "communicationItems/value like '*@$Domain*'" -all)
        if ($contacts.Count -gt 0) {
            # Extract unique company IDs from matching contacts
            $companyIds = $contacts |
                Where-Object { $_.company -and $_.company.id } |
                ForEach-Object { $_.company.id } |
                Select-Object -Unique

            if ($companyIds) {
                $matchingCompanies = @()
                foreach ($compId in $companyIds) {
                    $matchingCompanies += Get-CWMCompany -id $compId
                }

                if ($matchingCompanies.Count -eq 1) {
                    $matchingCompany = $matchingCompanies[0]
                    Write-Host "  Found company via contact email: $($matchingCompany.name) (ID: $($matchingCompany.id))" -ForegroundColor Cyan
                }
                elseif ($matchingCompanies.Count -gt 1) {
                    Write-Host "  Multiple companies found with contacts using '@$Domain':" -ForegroundColor Yellow
                    $i = 1
                    foreach ($co in $matchingCompanies) {
                        Write-Host "  [$i] $($co.name) (ID: $($co.id))" -ForegroundColor White
                        $i++
                    }
                    $selection = Read-Host "Select company (1-$($matchingCompanies.Count))"
                    if ($selection -match '^\d+$') {
                        $idx = [int]$selection - 1
                        if ($idx -ge 0 -and $idx -lt $matchingCompanies.Count) {
                            $matchingCompany = $matchingCompanies[$idx]
                        }
                    }
                }
            }
        }
    }
    catch {
        Write-Host "  Contact email search failed: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "  Trying fallback methods..." -ForegroundColor Yellow
    }

    # Strategy 2: Search by company website field
    if (-not $matchingCompany) {
        Write-Host "  Searching ConnectWise companies by website field..." -ForegroundColor Yellow
        try {
            $companies = @(Get-CWMCompany -condition "website like '*$Domain*'" -all)
            if ($companies.Count -gt 0) {
                if ($companies.Count -eq 1) {
                    $matchingCompany = $companies[0]
                    Write-Host "  Found company via website: $($matchingCompany.name) (ID: $($matchingCompany.id))" -ForegroundColor Cyan
                }
                else {
                    Write-Host "  Multiple companies found with website matching '$Domain':" -ForegroundColor Yellow
                    $i = 1
                    foreach ($co in $companies) {
                        Write-Host "  [$i] $($co.name) (ID: $($co.id))" -ForegroundColor White
                        $i++
                    }
                    $selection = Read-Host "Select company (1-$($companies.Count))"
                    if ($selection -match '^\d+$') {
                        $idx = [int]$selection - 1
                        if ($idx -ge 0 -and $idx -lt $companies.Count) {
                            $matchingCompany = $companies[$idx]
                        }
                    }
                }
            }
        }
        catch {
            Write-Host "  Website search failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    # Strategy 3: Manual search
    if (-not $matchingCompany) {
        Write-Host "  No automatic match found for '$Domain'." -ForegroundColor Yellow
        $searchTerm = Read-Host "Enter company name to search ConnectWise (or Enter to skip)"
        if (-not [string]::IsNullOrWhiteSpace($searchTerm)) {
            try {
                # Strip punctuation that breaks CW API conditions
                $sanitizedSearch = $searchTerm -replace "[',`"\\]", ''
                $companies = @(Get-CWMCompany -condition "name like '*$sanitizedSearch*'" -all)
                if ($companies.Count -gt 0) {
                    $i = 1
                    foreach ($co in $companies) {
                        Write-Host "  [$i] $($co.name) (ID: $($co.id))" -ForegroundColor White
                        $i++
                    }
                    $selection = Read-Host "Select company (1-$($companies.Count), or Enter to skip)"
                    if ($selection -match '^\d+$') {
                        $idx = [int]$selection - 1
                        if ($idx -ge 0 -and $idx -lt $companies.Count) {
                            $matchingCompany = $companies[$idx]
                        }
                    }
                }
                else {
                    Write-Host "  No companies found matching '$searchTerm'." -ForegroundColor Yellow
                }
            }
            catch {
                Write-Host "  Manual search failed: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }

    return $matchingCompany
}

# Handle reset host credentials
if ($ResetHostCredentials) {
    $creds = $null
    if (Test-Path $CredentialsPath) {
        $creds = Import-Clixml -Path $CredentialsPath
    }
    if ($creds -and $creds.WHM) {
        $current = $creds.WHM
        Write-Host "Current connection:" -ForegroundColor Yellow
        Write-Host "  Type:   $($current.Type)" -ForegroundColor Cyan
        Write-Host "  Server: $($current.Server)" -ForegroundColor Cyan
        Write-Host "  User:   $($current.Username)" -ForegroundColor Cyan
        Write-Host ""
    }
    if ($creds) {
        $creds.Remove('WHM')
        $creds | Export-Clixml -Path $CredentialsPath -Force
    }
    Write-Host "Host credentials cleared. Run the script again to enter new credentials." -ForegroundColor Green
    exit 0
}

# Handle read-only operations
switch ($PSCmdlet.ParameterSetName) {
    { $_ -in 'ListDomains', 'ListDomainsAll' } {
        $showAll = ($PSCmdlet.ParameterSetName -eq 'ListDomainsAll')
        $title = if ($showAll) { "All WHM Domains (with NS status)" } else { "WHM Domains Ready to Migrate (Bluehost NS)" }
        Write-Banner $title
        $domains = Get-WHMDomains -Credentials $WHMCreds
        Write-Host "Checking nameservers for $($domains.Count) domains..." -ForegroundColor Yellow
        Write-Host ""

        $domainResults = foreach ($d in $domains) {
            $domainName = $d.domain
            $nsStatus = 'Unknown'
            $nsDetail = ''
            try {
                $nsRecords = Resolve-DnsName -Name $domainName -Type NS -ErrorAction Stop -DnsOnly | Where-Object { $_.QueryType -eq 'NS' }
                $nameservers = $nsRecords | ForEach-Object { $_.NameHost.ToLower().TrimEnd('.') }

                if ($nameservers) {
                    $pointsToBluehost = $true
                    foreach ($ns in $nameservers) {
                        if ($ns -notmatch '\.(bluehost\.com|hostmonster\.com|justhost\.com|fastdomain\.com|rhostbh\.com)$') {
                            $pointsToBluehost = $false
                            break
                        }
                    }

                    if ($pointsToBluehost) {
                        $nsStatus = 'Bluehost'
                    }
                    elseif ($nameservers | Where-Object { $_ -match '\.cloudflare\.com$' }) {
                        $nsStatus = 'Cloudflare'
                    }
                    elseif ($nameservers | Where-Object { $_ -match '\.domaincontrol\.com$' }) {
                        $nsStatus = 'GoDaddy'
                    }
                    elseif ($nameservers | Where-Object { $_ -match '\.worldnic\.com$' }) {
                        $nsStatus = 'Network Solutions'
                    }
                    else {
                        $nsStatus = 'Other'
                    }
                    $nsDetail = ($nameservers | Select-Object -First 2) -join ', '
                }
                else {
                    $nsStatus = 'No NS'
                }
            }
            catch {
                $nsStatus = 'Error'
            }

            [PSCustomObject]@{
                Domain      = $domainName
                Nameservers = $nsStatus
                Detail      = $nsDetail
            }
        }

        if ($showAll) {
            $domainResults | Format-Table Domain, Nameservers, Detail -AutoSize
        }
        else {
            $bluehostOnly = $domainResults | Where-Object { $_.Nameservers -eq 'Bluehost' }
            if ($bluehostOnly) {
                $bluehostOnly | Format-Table Domain -AutoSize
            }
            else {
                Write-Host "No domains with Bluehost nameservers found." -ForegroundColor Yellow
            }
        }

        $bluehostCount = ($domainResults | Where-Object { $_.Nameservers -eq 'Bluehost' }).Count
        Write-Host "Bluehost: $bluehostCount | Total: $($domains.Count)" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "To migrate a domain, run:" -ForegroundColor Yellow
        Write-Host "  .\Bluehost-Cloudflare-Migrate-DNS.ps1 -Domain ""example.com""" -ForegroundColor Cyan
        exit 0
    }

    'ListCFAccounts' {
        Write-Banner "Cloudflare Accounts"
        $accounts = Get-AllCloudflareAccounts
        $accounts | Format-Table id, name, type -AutoSize
        exit 0
    }

    'ListCFZones' {
        Write-Banner "Cloudflare Zones"
        $zones = Get-CloudflareZones -Credentials $CloudflareCreds
        $zones | Format-Table id, name, status, @{L='Account';E={$_.account.name}} -AutoSize
        exit 0
    }

    'Preview' {
        Write-Banner "DNS Records Preview: $Domain (WHM)"
        $records = Get-WHMDNSRecords -Credentials $WHMCreds -Domain $Domain

        Write-Host "Found $($records.Count) records:`n" -ForegroundColor Yellow

        $records | ForEach-Object {
            $priorityStr = if ($_.priority) { " (Priority: $($_.priority))" } else { "" }
            $ttlStr = if ($_.ttl) { " TTL:$($_.ttl)" } else { "" }
            Write-Host "  $($_.type.PadRight(6)) $($_.name.PadRight(30)) -> $($_.data)$priorityStr$ttlStr"
        }

        Write-Host ""
        exit 0
    }
}

# If no domain specified, show usage
if (-not $Domain) {
    Write-Host "No domain specified. Use one of:" -ForegroundColor Yellow
    Write-Host "  .\Bluehost-Cloudflare-Migrate-DNS.ps1 -ListDomains          # Domains ready to migrate" -ForegroundColor Cyan
    Write-Host "  .\Bluehost-Cloudflare-Migrate-DNS.ps1 -ListDomainsAll       # All domains with NS status" -ForegroundColor Cyan
    Write-Host "  .\Bluehost-Cloudflare-Migrate-DNS.ps1 -Domain ""example.com"" # Migrate a specific domain" -ForegroundColor Cyan
    exit 0
}

# Full migration workflow
Write-Banner "Cloudflare DNS Migration (from Bluehost/WHM)"
Write-Host "Domain: $Domain" -ForegroundColor Yellow

# Step 1: Retrieve DNS records from WHM
Write-Banner "Step 1: Retrieving DNS Records from WHM"
try {
    $whmRecords = Get-WHMDNSRecords -Credentials $WHMCreds -Domain $Domain
    Write-Host "Found $($whmRecords.Count) DNS records." -ForegroundColor Green

    $whmRecords | ForEach-Object {
        $priorityStr = if ($_.priority) { " (Priority: $($_.priority))" } else { "" }
        Write-Host "  $($_.type.PadRight(6)) $($_.name.PadRight(30)) -> $($_.data)$priorityStr"
    }
}
catch {
    Write-Error "Failed to retrieve DNS records from WHM: $($_.Exception.Message)"
    exit 1
}

# Pre-flight: Verify nameservers point to Bluehost
Write-Host ""
Write-Host "Checking current nameserver configuration for $Domain..." -ForegroundColor Yellow
try {
    $nsRecords = Resolve-DnsName -Name $Domain -Type NS -ErrorAction Stop | Where-Object { $_.QueryType -eq 'NS' }
    $currentNS = $nsRecords | ForEach-Object { $_.NameHost.ToLower().TrimEnd('.') }

    if ($currentNS) {
        Write-Host "Current nameservers:" -ForegroundColor Cyan
        $currentNS | ForEach-Object { Write-Host "  $_" -ForegroundColor White }

        # Check if nameservers point to Bluehost/cPanel hosting
        $isBluehost = $true
        foreach ($ns in $currentNS) {
            if ($ns -notmatch '\.(bluehost\.com|hostmonster\.com|justhost\.com|fastdomain\.com|rhostbh\.com)$') {
                $isBluehost = $false
                break
            }
        }

        if (-not $isBluehost) {
            Write-Host ""
            Write-Host ("!" * 60) -ForegroundColor Red
            Write-Host " WARNING: Nameservers are NOT pointing to Bluehost!" -ForegroundColor Red
            Write-Host ("!" * 60) -ForegroundColor Red
            Write-Host ""
            Write-Host "DNS for this domain is not managed by Bluehost. The records" -ForegroundColor Yellow
            Write-Host "in WHM may NOT reflect the live DNS records." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Migration aborted. Resolve the nameserver situation before" -ForegroundColor Yellow
            Write-Host "migrating this domain to Cloudflare." -ForegroundColor Yellow
            exit 1
        }
        else {
            Write-Host "Nameservers confirmed pointing to Bluehost." -ForegroundColor Green
        }
    }
    else {
        Write-Warning "Could not resolve NS records for $Domain. Proceeding with caution."
    }
}
catch {
    Write-Warning "Could not check nameservers for $Domain : $($_.Exception.Message)"
    Write-Warning "Proceeding without nameserver validation."
}

# Step 2: Look up company in ConnectWise by email domain
Write-Banner "Step 2: ConnectWise Company Lookup"
$cwCompany = $null
if ($ConnectWiseCreds -and (Get-Module -ListAvailable -Name ConnectWiseManageAPI)) {
    try {
        Import-Module ConnectWiseManageAPI -Force
        Connect-CWM @ConnectWiseCreds

        $cwCompany = Find-CWCompanyByDomain -Domain $Domain

        if ($cwCompany) {
            Write-Host ""
            Write-Host "Company: $($cwCompany.name) (ID: $($cwCompany.id))" -ForegroundColor Green

            # Check if company is active
            $companyStatus = $cwCompany.status.name
            if ($companyStatus -and $companyStatus -ne 'Active') {
                Write-Host ""
                Write-Host ("!" * 60) -ForegroundColor Red
                Write-Host " WARNING: Company '$($cwCompany.name)' is NOT active!" -ForegroundColor Red
                Write-Host " Status: $companyStatus" -ForegroundColor Red
                Write-Host ("!" * 60) -ForegroundColor Red
                Write-Host ""
                Write-Host "This company is not an active client. DNS should not be" -ForegroundColor Yellow
                Write-Host "migrated for inactive companies." -ForegroundColor Yellow
                Write-Host ""
                Write-Host "Skipping domain '$Domain'." -ForegroundColor Yellow
                exit 0
            }

            Write-Host "Status: Active" -ForegroundColor Green

            # Use CW company name as CustomerName if not explicitly provided
            if (-not $CustomerName) {
                $CustomerName = $cwCompany.name
                Write-Host "Using company name as Cloudflare account name: $CustomerName" -ForegroundColor Cyan
            }
        }
        else {
            Write-Host "No ConnectWise company found for domain '$Domain'." -ForegroundColor Yellow
            Write-Host "WARNING: Active status could not be verified." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Warning "ConnectWise lookup failed: $($_.Exception.Message)"
    }
}
else {
    Write-Host "ConnectWise integration not available. Skipping company lookup." -ForegroundColor Gray
}

# Ensure we have a CustomerName or AccountId
if (-not $CustomerName -and -not $AccountId) {
    $CustomerName = Read-Host "Enter a name for the Cloudflare account (or provide -AccountId to use existing)"
    if ([string]::IsNullOrWhiteSpace($CustomerName)) {
        Write-Error "You must specify either -CustomerName or -AccountId."
        exit 1
    }
}

# Sanitize CustomerName for Cloudflare (invalid characters)
if ($CustomerName) {
    $originalName = $CustomerName
    $CustomerName = $CustomerName -replace '&', 'and'
    $CustomerName = $CustomerName -replace '[<>]', ''
    $CustomerName = $CustomerName -replace '\s{2,}', ' '
    $CustomerName = $CustomerName.Trim()
    if ($CustomerName -ne $originalName) {
        Write-Host "Cloudflare account name sanitized: '$originalName' -> '$CustomerName'" -ForegroundColor Yellow
    }
}

# Step 3: Create or select Cloudflare account
Write-Banner "Step 3: Cloudflare Account Setup"

if ($AccountId) {
    Write-Host "Using existing account: $AccountId" -ForegroundColor Yellow
    # Fetch the specific account directly (avoids pagination issues)
    $cfAccount = @{ id = $AccountId; name = $AccountId }
    try {
        $headers = @{
            "X-Auth-Email" = $CloudflareCreds.Email
            "X-Auth-Key"   = $CloudflareCreds.ApiKey
            "Content-Type" = "application/json"
        }
        $acctResponse = Invoke-RestMethod -Uri "https://api.cloudflare.com/client/v4/accounts/$AccountId" `
            -Headers $headers -Method Get
        if ($acctResponse.success) {
            $cfAccount = $acctResponse.result
            Write-Host "Account name: $($cfAccount.name)" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Warning "Could not fetch account details, but will proceed with ID: $AccountId"
    }
}
else {
    # Check if an account with this name already exists
    Write-Host "Checking for existing Cloudflare account '$CustomerName'..." -ForegroundColor Yellow
    $existingAccounts = Get-AllCloudflareAccounts
    $normalizedCustomer = $CustomerName -replace '[^\w\s]', '' -replace '\s+', ' '
    $matchingAccounts = $existingAccounts | Where-Object {
        $normalizedName = $_.name -replace '[^\w\s]', '' -replace '\s+', ' '
        $normalizedName -eq $normalizedCustomer
    }

    if ($matchingAccounts) {
        $matchingAccount = $matchingAccounts | Select-Object -First 1
        Write-Host ""
        Write-Host "An existing Cloudflare account was found with this name:" -ForegroundColor Yellow
        Write-Host "  Name: $($matchingAccount.name)" -ForegroundColor Cyan
        Write-Host "  ID:   $($matchingAccount.id)" -ForegroundColor Cyan
        Write-Host ""

        $useExisting = New-Object System.Management.Automation.Host.ChoiceDescription "&Use existing", "Add the zone to the existing account."
        $createNew = New-Object System.Management.Automation.Host.ChoiceDescription "Create &new", "Create a new account with the same name."
        $cancel = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel", "Cancel the operation."
        $accountChoice = $Host.UI.PromptForChoice("Account already exists", "Account '$($matchingAccount.name)' already exists. What would you like to do?", @($useExisting, $createNew, $cancel), 0)

        switch ($accountChoice) {
            0 {
                $cfAccount = $matchingAccount
                Write-Host "Using existing account: $($cfAccount.name) ($($cfAccount.id))" -ForegroundColor Green
            }
            1 {
                try {
                    $cfAccount = New-CloudflareAccount -Credentials $CloudflareCreds -Name $CustomerName -Confirm:$false
                    Write-Host "New account created: $($cfAccount.id)" -ForegroundColor Green
                }
                catch {
                    Write-Error "Failed to create Cloudflare account: $($_.Exception.Message)"
                    exit 1
                }
            }
            2 {
                Write-Host "Cancelled by user." -ForegroundColor Yellow
                exit 0
            }
        }
    }
    else {
        Write-Host "No existing account found." -ForegroundColor Gray
        Write-Host "Will create new account: $CustomerName" -ForegroundColor Yellow
        Write-Host ""

        if (-not (Confirm-Action "Create new Cloudflare account '$CustomerName'?")) {
            Write-Host "Cancelled by user." -ForegroundColor Yellow
            exit 0
        }

        try {
            $cfAccount = New-CloudflareAccount -Credentials $CloudflareCreds -Name $CustomerName -Confirm:$false
            Write-Host "Account created: $($cfAccount.id)" -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to create Cloudflare account: $($_.Exception.Message)"
            exit 1
        }
    }
}

# Step 4: Add zone to Cloudflare
Write-Banner "Step 4: Add Zone to Cloudflare"
Write-Host "Domain: $Domain" -ForegroundColor Yellow
Write-Host "Account: $($cfAccount.id)" -ForegroundColor Yellow
Write-Host ""

if (-not (Confirm-Action "Add zone '$Domain' to Cloudflare?")) {
    Write-Host "Cancelled by user." -ForegroundColor Yellow
    exit 0
}

try {
    $cfZone = New-CloudflareZone -Credentials $CloudflareCreds -AccountId $cfAccount.id -Domain $Domain -Confirm:$false
    Write-Host ""
    Write-Host "Zone created successfully!" -ForegroundColor Green
    Write-Host "Zone ID: $($cfZone.id)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "IMPORTANT: Update nameservers at Bluehost to:" -ForegroundColor Yellow
    $cfZone.name_servers | ForEach-Object { Write-Host "  $_" -ForegroundColor White }
}
catch {
    Write-Error "Failed to create zone: $($_.Exception.Message)"
    exit 1
}

# Step 5: Import DNS records
Write-Banner "Step 5: Import DNS Records"
Write-Host "Ready to import $($whmRecords.Count) records to Cloudflare." -ForegroundColor Yellow
Write-Host ""

# Preview first
Import-CloudflareDNS -Credentials $CloudflareCreds -ZoneId $cfZone.id -Records $whmRecords -PreviewOnly

if (-not (Confirm-Action "Proceed with DNS record import?")) {
    Write-Host "Cancelled by user." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Zone was created but no DNS records were imported." -ForegroundColor Yellow
    Write-Host "You can import records later manually." -ForegroundColor Cyan
    exit 0
}

try {
    $importResult = Import-CloudflareDNS -Credentials $CloudflareCreds -ZoneId $cfZone.id -Records $whmRecords -Confirm:$false
}
catch {
    Write-Error "Failed to import DNS records: $($_.Exception.Message)"
    exit 1
}

# Step 6: Nameserver instructions (manual for Bluehost)
Write-Banner "Step 6: Update Nameservers at Bluehost (Manual)"
Write-Host "Cloudflare assigned nameservers:" -ForegroundColor Yellow
$cfZone.name_servers | ForEach-Object { Write-Host "  $_" -ForegroundColor Cyan }
Write-Host ""
Write-Host "Bluehost does not have a public registrar API." -ForegroundColor Yellow
Write-Host "You must update the nameservers manually:" -ForegroundColor Yellow
Write-Host ""
Write-Host "  1. Log in to your Bluehost account" -ForegroundColor White
Write-Host "  2. Go to Domains > My Domains" -ForegroundColor White
Write-Host "  3. Click 'Manage' next to '$Domain'" -ForegroundColor White
Write-Host "  4. Under 'Name Servers', select 'Custom' and enter:" -ForegroundColor White
$cfZone.name_servers | ForEach-Object { Write-Host "     $_" -ForegroundColor Cyan }
Write-Host "  5. Save changes" -ForegroundColor White
Write-Host ""

$null = Read-Host "Press Enter when nameservers have been updated (or to continue without updating)"

# Step 7: Create ITGlue DNS/Registrar Asset
Write-Banner "Step 7: Document in ITGlue"
$itgAssetCreated = $false
if ($ITGlueCreds -and $ITGlueCreds.FlexibleAssetTypeId -and (Get-Module -ListAvailable -Name ITGlueAPI)) {
    Write-Host "Creating ITGlue DNS/Registrar asset..." -ForegroundColor Yellow
    try {
        Import-Module ITGlueAPI -Force
        Add-ITGlueBaseURI -base_uri $ITGlueCreds.BaseUri
        Add-ITGlueAPIKey -Api_Key $ITGlueCreds.ApiKey

        # Use ConnectWise company name to find ITGlue organization (falls back to Cloudflare account name)
        $orgName = if ($cwCompany) { $cwCompany.name } else { $cfAccount.name }
        Write-Host "  Searching ITGlue for org: $orgName" -ForegroundColor Cyan

        if ($orgName) {
            $allOrgs = @()
            $pageNum = 1
            do {
                $page = Get-ITGlueOrganizations -page_size 1000 -page_number $pageNum
                if ($page.data) {
                    $allOrgs += $page.data
                }
                $pageNum++
            } while ($page.data -and $page.data.Count -eq 1000)

            # Normalize name for flexible matching (& -> and, strip punctuation, collapse spaces)
            $normalizedSearch = $orgName -replace '&', 'and' -replace '[^\w\s]', '' -replace '\s+', ' ' -replace '^\s+|\s+$', ''
            $matchingOrg = $allOrgs | Where-Object {
                $normalizedName = $_.attributes.name -replace '&', 'and' -replace '[^\w\s]', '' -replace '\s+', ' ' -replace '^\s+|\s+$', ''
                $normalizedName -eq $normalizedSearch
            } | Select-Object -First 1

            if ($matchingOrg) {
                $itgOrgId = $matchingOrg.id
                Write-Host "  Found ITGlue org: $($matchingOrg.attributes.name) (ID: $itgOrgId)" -ForegroundColor Cyan

                # Build management URL
                $managementUrl = "https://dash.cloudflare.com/$($cfAccount.id)/$Domain/dns/records"

                # Look up domain asset for tagging
                $domainTag = $null
                try {
                    $domains = Get-ITGlueDomains -filter_organization_id $itgOrgId
                    if ($domains.data) {
                        $matchingDomain = $domains.data | Where-Object { $_.attributes.name -eq $Domain }
                        if ($matchingDomain) {
                            $domainTag = @($matchingDomain.id)
                            Write-Host "  Found matching domain in ITGlue: $Domain (ID: $($matchingDomain.id))" -ForegroundColor Cyan
                        }
                    }
                }
                catch {
                    Write-Host "  Could not query ITGlue domains: $($_.Exception.Message)" -ForegroundColor Yellow
                }

                # Build traits
                $traits = @{
                    'provider' = 'CloudFlare'
                    'dns' = $true
                    'registrar' = $false
                    'dns-credentials' = $ITGlueCreds.DnsCredentialAssetId
                    'management-url' = $managementUrl
                    'notes' = "Cloudflare Zone ID: $($cfZone.id)<br>Cloudflare Account ID: $($cfAccount.id)<br>Nameservers: $($cfZone.name_servers -join ', ')<br>Migrated: $(Get-Date -Format 'yyyy-MM-dd')<br>Source: Bluehost WHM"
                }

                if ($domainTag) {
                    $traits['domain-s'] = $domainTag
                }

                # Create the flexible asset
                $assetData = @{
                    type = 'flexible-assets'
                    attributes = @{
                        'organization-id' = [int]$itgOrgId
                        'flexible-asset-type-id' = $ITGlueCreds.FlexibleAssetTypeId
                        traits = $traits
                    }
                }

                $itgAsset = New-ITGlueFlexibleAssets -data $assetData
                if ($itgAsset.data) {
                    $itgAssetCreated = $true
                    Write-Host "  ITGlue asset created: $($itgAsset.data.attributes.name)" -ForegroundColor Green
                }
            }
            else {
                Write-Host "  Organization '$orgName' not found in ITGlue." -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "  Could not determine organization name for ITGlue." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Failed to create ITGlue asset: $($_.Exception.Message)" -ForegroundColor Red
    }
}
else {
    Write-Host "ITGlue integration not configured. Skipping." -ForegroundColor Gray
}

# Summary
Write-Banner "Migration Complete"
Write-Host "Domain:     $Domain" -ForegroundColor Green
Write-Host "Account ID: $($cfAccount.id)" -ForegroundColor Green
Write-Host "Zone ID:    $($cfZone.id)" -ForegroundColor Green
Write-Host ""
Write-Host "Records Imported: $($importResult.Success.Count)" -ForegroundColor Green
if ($importResult.Failed.Count -gt 0) {
    Write-Host "Records Failed:   $($importResult.Failed.Count)" -ForegroundColor Red
}
if ($cwCompany) {
    Write-Host "CW Company:     $($cwCompany.name)" -ForegroundColor Green
}
if ($itgAssetCreated) {
    Write-Host "ITGlue Asset:   Created" -ForegroundColor Green
}
Write-Host ""
Write-Host "NEXT STEPS:" -ForegroundColor Yellow
$stepNum = 1

Write-Host "$stepNum. Update nameservers at Bluehost to:" -ForegroundColor White
$cfZone.name_servers | ForEach-Object { Write-Host "   $_" -ForegroundColor Cyan }
$stepNum++

Write-Host "$stepNum. Wait for DNS propagation (up to 48 hours)" -ForegroundColor White
$stepNum++
Write-Host "$stepNum. Verify zone is active in Cloudflare dashboard" -ForegroundColor White
$stepNum++

if ($itgAssetCreated) {
    Write-Host "$stepNum. Update ITGlue asset 'Registrar' field if/when registrar transfer is completed" -ForegroundColor White
}
