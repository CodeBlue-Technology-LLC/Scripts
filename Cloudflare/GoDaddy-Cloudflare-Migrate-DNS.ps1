<#
.SYNOPSIS
    Migrate DNS records from GoDaddy to Cloudflare.

.DESCRIPTION
    Interactive migration tool that transfers DNS records from GoDaddy to Cloudflare.
    All write operations require explicit user confirmation.

.PARAMETER Domain
    The domain to migrate (e.g., "example.com").

.PARAMETER CustomerName
    Name for the Cloudflare account/subtenant.

.PARAMETER AccountId
    Use an existing Cloudflare account instead of creating a new one.

.PARAMETER ListGoDaddyDomains
    List all domains in the GoDaddy account (read-only).

.PARAMETER ListCloudflareAccounts
    List all accounts in Cloudflare (read-only).

.PARAMETER ListCloudflareZones
    List all zones in Cloudflare (read-only).

.PARAMETER PreviewRecords
    Preview DNS records for a domain without making changes.

.EXAMPLE
    .\Migrate-DNS.ps1 -ListGoDaddyDomains

.EXAMPLE
    .\Migrate-DNS.ps1 -PreviewRecords "example.com"

.EXAMPLE
    .\Migrate-DNS.ps1 -Domain "example.com" -CustomerName "Acme Corp"

.EXAMPLE
    .\Migrate-DNS.ps1 -Domain "example.com" -AccountId "existing-account-id"
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

    [Parameter(ParameterSetName = 'ListGoDaddy')]
    [switch]$ListGoDaddyDomains,

    [Parameter(ParameterSetName = 'ListCFAccounts')]
    [switch]$ListCloudflareAccounts,

    [Parameter(ParameterSetName = 'ListCFZones')]
    [switch]$ListCloudflareZones,

    [Parameter(ParameterSetName = 'Preview')]
    [switch]$PreviewRecords
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Import modules
Import-Module "$ScriptDir\Modules\GoDaddy.psm1" -Force
Import-Module "$ScriptDir\Modules\Cloudflare.psm1" -Force

# Check for ConnectWiseManageAPI module (for ticket creation)
if (-not (Get-Module -ListAvailable -Name ConnectWiseManageAPI)) {
    Write-Host "ConnectWiseManageAPI module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name ConnectWiseManageAPI -Scope CurrentUser -Force -AllowClobber
        Write-Host "ConnectWiseManageAPI installed successfully." -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to install ConnectWiseManageAPI: $($_.Exception.Message)"
        Write-Warning "Ticket creation will be skipped."
    }
}

# Check for ITGlueAPI module (for asset documentation)
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

    # --- GoDaddy ---
    if (-not $creds.GoDaddy) {
        Write-Host "`n=== GoDaddy Credentials ===" -ForegroundColor Cyan
        $gdKey = Read-Host "GoDaddy API Key"
        $gdSecret = Read-Host "GoDaddy API Secret"
        $creds.GoDaddy = @{ ApiKey = $gdKey; ApiSecret = $gdSecret }
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

                # List existing flexible asset types
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
                        # Create new flexible asset type
                        $creds.ITGlue.FlexibleAssetTypeId = New-DnsRegistrarAssetType
                        $needsSave = $true
                    }
                    elseif ($selection -match '^\d+$') {
                        $idx = [int]$selection - 1
                        if ($idx -ge 0 -and $idx -lt $fatypes.data.Count) {
                            $selected = $fatypes.data[$idx]
                            $creds.ITGlue.FlexibleAssetTypeId = [int]$selected.id
                            $needsSave = $true

                            # Validate fields
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

    # Create required fields
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
$GoDaddyCreds = $Credentials.GoDaddy
$CloudflareCreds = $Credentials.Cloudflare
$ConnectWiseCreds = $Credentials.ConnectWise
$ITGlueCreds = $Credentials.ITGlue

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

# Handle read-only operations
switch ($PSCmdlet.ParameterSetName) {
    'ListGoDaddy' {
        Write-Banner "GoDaddy Domains"
        $domains = Get-GoDaddyDomains -Credentials $GoDaddyCreds | Where-Object { $_.status -notin 'CANCELLED', 'CANCELLED_TRANSFER', 'PENDING_DNS_INACTIVE', 'TRANSFERRED_OUT', 'UPDATED_OWNERSHIP' }
        $domains | Format-Table domain, status, expires -AutoSize
        exit 0
    }

    'ListCFAccounts' {
        Write-Banner "Cloudflare Accounts"
        $accounts = Get-CloudflareAccounts -Credentials $CloudflareCreds
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
        Write-Banner "DNS Records Preview: $Domain"
        $records = Get-GoDaddyDNSRecords -Credentials $GoDaddyCreds -Domain $Domain

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

# If no domain specified, list available GoDaddy domains
if (-not $Domain) {
    Write-Banner "Available GoDaddy Domains"
    $domains = Get-GoDaddyDomains -Credentials $GoDaddyCreds | Where-Object { $_.status -notin 'CANCELLED', 'CANCELLED_TRANSFER', 'PENDING_DNS_INACTIVE', 'TRANSFERRED_OUT', 'UPDATED_OWNERSHIP' }
    $domains | Format-Table domain, status, expires -AutoSize
    Write-Host "To migrate a domain, run:" -ForegroundColor Yellow
    Write-Host "  .\Migrate-DNS.ps1 -Domain ""example.com"" -CustomerName ""Customer Name""" -ForegroundColor Cyan
    exit 0
}

# Full migration workflow
Write-Banner "Cloudflare DNS Migration"
Write-Host "Domain: $Domain" -ForegroundColor Yellow

if (-not $CustomerName -and -not $AccountId) {
    Write-Error "You must specify either -CustomerName (to create new account) or -AccountId (to use existing)."
    exit 1
}

# Step 1: Query GoDaddy for DNS records
Write-Banner "Step 1: Retrieving DNS Records from GoDaddy"
try {
    $gdRecords = Get-GoDaddyDNSRecords -Credentials $GoDaddyCreds -Domain $Domain
    Write-Host "Found $($gdRecords.Count) DNS records." -ForegroundColor Green

    $gdRecords | ForEach-Object {
        $priorityStr = if ($_.priority) { " (Priority: $($_.priority))" } else { "" }
        Write-Host "  $($_.type.PadRight(6)) $($_.name.PadRight(30)) -> $($_.data)$priorityStr"
    }
}
catch {
    if ($_.Exception.Message -match '404') {
        Write-Warning "GoDaddy returned 404 for '$Domain'. This typically means DNS is not hosted at GoDaddy for this domain."
        Write-Warning "Verify that the domain exists in your GoDaddy account and that DNS records are managed there."
    }
    else {
        Write-Error "Failed to retrieve DNS records from GoDaddy: $($_.Exception.Message)"
    }
    exit 1
}

# Pre-flight: Check nameservers
Write-Host ""
Write-Host "Checking current nameserver configuration..." -ForegroundColor Yellow
try {
    $domainDetails = Get-GoDaddyDomainDetails -Credentials $GoDaddyCreds -Domain $Domain

    $currentNS = $domainDetails.nameServers
    $isGoDaddyNS = $true
    foreach ($ns in $currentNS) {
        if ($ns -notmatch '\.(domaincontrol\.com|secureserver\.net)$') {
            $isGoDaddyNS = $false
            break
        }
    }

    if (-not $isGoDaddyNS) {
        Write-Host ""
        Write-Host ("!" * 60) -ForegroundColor Red
        Write-Host " WARNING: Nameservers are NOT hosted on GoDaddy defaults!" -ForegroundColor Red
        Write-Host ("!" * 60) -ForegroundColor Red
        Write-Host ""
        Write-Host "Current nameservers:" -ForegroundColor Yellow
        $currentNS | ForEach-Object { Write-Host "  $_" -ForegroundColor White }
        Write-Host ""
        Write-Host "DNS may be managed by a third party. The records retrieved from" -ForegroundColor Yellow
        Write-Host "GoDaddy may NOT reflect the live DNS records for this domain." -ForegroundColor Yellow
        Write-Host ""

        if (-not (Confirm-Action "Nameservers are not GoDaddy defaults. Continue anyway?")) {
            Write-Host "Cancelled by user." -ForegroundColor Yellow
            exit 0
        }
    }
    else {
        Write-Host "Nameservers are on GoDaddy defaults." -ForegroundColor Green
    }
}
catch {
    Write-Warning "Could not check nameserver configuration: $($_.Exception.Message)"
    Write-Warning "Proceeding without nameserver validation."
}

# Pre-flight: Check for GoDaddy website hosting
Write-Host ""
Write-Host "Checking for GoDaddy website hosting..." -ForegroundColor Yellow

$gdHostingIndicators = @()
foreach ($rec in $gdRecords) {
    $data = $rec.data
    # Check for GoDaddy hosting CNAMEs
    if ($rec.type -eq 'CNAME' -and $data -match '\.(godaddysites\.com|secureserver\.net|secureserversites\.net)$') {
        $gdHostingIndicators += $rec
    }
    # Check for GoDaddy hosting A record IPs (common ranges)
    if ($rec.type -eq 'A' -and ($data -match '^184\.168\.' -or $data -match '^50\.63\.' -or $data -match '^160\.153\.' -or $data -match '^173\.201\.')) {
        $gdHostingIndicators += $rec
    }
}

if ($gdHostingIndicators.Count -gt 0) {
    Write-Host ""
    Write-Host ("!" * 60) -ForegroundColor Red
    Write-Host " WARNING: GoDaddy website hosting detected!" -ForegroundColor Red
    Write-Host ("!" * 60) -ForegroundColor Red
    Write-Host ""
    Write-Host "The following records point to GoDaddy-hosted services:" -ForegroundColor Yellow
    $gdHostingIndicators | ForEach-Object {
        Write-Host "  $($_.type.PadRight(6)) $($_.name.PadRight(30)) -> $($_.data)" -ForegroundColor White
    }
    Write-Host ""
    Write-Host "Transferring DNS away from GoDaddy may break the website if" -ForegroundColor Yellow
    Write-Host "it is actively hosted on GoDaddy's platform (Website Builder," -ForegroundColor Yellow
    Write-Host "Managed WordPress, or cPanel hosting)." -ForegroundColor Yellow
    Write-Host ""

    if (-not (Confirm-Action "GoDaddy hosting detected. Continue with migration?")) {
        Write-Host "Cancelled by user." -ForegroundColor Yellow
        exit 0
    }
}
else {
    Write-Host "No GoDaddy website hosting detected." -ForegroundColor Green
}

# Step 2: Create or select Cloudflare account
Write-Banner "Step 2: Cloudflare Account Setup"

if ($AccountId) {
    Write-Host "Using existing account: $AccountId" -ForegroundColor Yellow
    # Fetch account details to get the name
    $cfAccounts = Get-CloudflareAccounts -Credentials $CloudflareCreds
    $cfAccount = $cfAccounts | Where-Object { $_.id -eq $AccountId }
    if (-not $cfAccount) {
        Write-Error "Account not found: $AccountId"
        exit 1
    }
    Write-Host "Account name: $($cfAccount.name)" -ForegroundColor Cyan
}
else {
    # Check if an account with this name already exists
    Write-Host "Checking for existing Cloudflare account '$CustomerName'..." -ForegroundColor Yellow
    $existingAccounts = Get-CloudflareAccounts -Credentials $CloudflareCreds
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

# Step 3: Add zone to Cloudflare
Write-Banner "Step 3: Add Zone to Cloudflare"
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
    Write-Host "IMPORTANT: Update nameservers at your registrar to:" -ForegroundColor Yellow
    $cfZone.name_servers | ForEach-Object { Write-Host "  $_" -ForegroundColor White }
}
catch {
    Write-Error "Failed to create zone: $($_.Exception.Message)"
    exit 1
}

# Step 4: Import DNS records
Write-Banner "Step 4: Import DNS Records"
Write-Host "Ready to import $($gdRecords.Count) records to Cloudflare." -ForegroundColor Yellow
Write-Host ""

# Preview first
Import-CloudflareDNS -Credentials $CloudflareCreds -ZoneId $cfZone.id -Records $gdRecords -PreviewOnly

if (-not (Confirm-Action "Proceed with DNS record import?")) {
    Write-Host "Cancelled by user." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Zone was created but no DNS records were imported." -ForegroundColor Yellow
    Write-Host "You can import records later with:" -ForegroundColor Cyan
    Write-Host "  Import-CloudflareDNS -ZoneId '$($cfZone.id)' ..." -ForegroundColor White
    exit 0
}

try {
    $importResult = Import-CloudflareDNS -Credentials $CloudflareCreds -ZoneId $cfZone.id -Records $gdRecords -Confirm:$false
}
catch {
    Write-Error "Failed to import DNS records: $($_.Exception.Message)"
    exit 1
}

# Step 5: Update nameservers at GoDaddy
Write-Banner "Step 5: Update Nameservers at GoDaddy"
Write-Host "Cloudflare assigned nameservers:" -ForegroundColor Yellow
$cfZone.name_servers | ForEach-Object { Write-Host "  $_" -ForegroundColor Cyan }
Write-Host ""

$nameserversUpdated = $false
if (Confirm-Action "Update GoDaddy to use these Cloudflare nameservers?") {
    try {
        Set-GoDaddyNameservers -Credentials $GoDaddyCreds -Domain $Domain -NameServers $cfZone.name_servers -Confirm:$false
        $nameserversUpdated = $true
    }
    catch {
        Write-Host "Failed to update nameservers: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "You will need to update nameservers manually at GoDaddy." -ForegroundColor Yellow
    }
}
else {
    Write-Host "Skipped nameserver update." -ForegroundColor Yellow
}

# Step 6: Create ITGlue DNS/Registrar Asset
Write-Banner "Step 6: Document in ITGlue"
$itgAssetCreated = $false
if ($ITGlueCreds -and $ITGlueCreds.FlexibleAssetTypeId -and (Get-Module -ListAvailable -Name ITGlueAPI)) {
    Write-Host "Creating ITGlue DNS/Registrar asset..." -ForegroundColor Yellow
    try {
        Import-Module ITGlueAPI -Force
        Add-ITGlueBaseURI -base_uri $ITGlueCreds.BaseUri
        Add-ITGlueAPIKey -Api_Key $ITGlueCreds.ApiKey

        # Use Cloudflare subaccount name to find ITGlue organization
        $orgName = $cfAccount.name
        Write-Host "  Searching ITGlue for org: $orgName" -ForegroundColor Cyan

        if ($orgName) {
            # Paginate through all ITGlue organizations
            $allOrgs = @()
            $pageNum = 1
            do {
                $page = Get-ITGlueOrganizations -page_size 1000 -page_number $pageNum
                if ($page.data) {
                    $allOrgs += $page.data
                }
                $pageNum++
            } while ($page.data -and $page.data.Count -eq 1000)

            # Normalize name for flexible matching (remove punctuation)
            $normalizedSearch = $orgName -replace '[^\w\s]', '' -replace '\s+', ' '
            $matchingOrg = $allOrgs | Where-Object {
                $normalizedName = $_.attributes.name -replace '[^\w\s]', '' -replace '\s+', ' '
                $normalizedName -eq $normalizedSearch
            } | Select-Object -First 1

            if ($matchingOrg) {
                $itgOrgId = $matchingOrg.id
                Write-Host "  Found ITGlue org: $($matchingOrg.attributes.name) (ID: $itgOrgId)" -ForegroundColor Cyan

                # Build management URL with account ID and zone name
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
                    'notes' = "Cloudflare Zone ID: $($cfZone.id)<br>Cloudflare Account ID: $($cfAccount.id)<br>Nameservers: $($cfZone.name_servers -join ', ')<br>Migrated: $(Get-Date -Format 'yyyy-MM-dd')"
                }

                # Add domain tag if found
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

# Step 7: Unlock domain and retrieve auth code for transfer (optional)
Write-Banner "Step 7: Unlock Domain & Retrieve Auth Code (Optional)"
Write-Host "If you plan to transfer this domain's registration to Cloudflare," -ForegroundColor Yellow
Write-Host "the domain must be unlocked and the auth code retrieved." -ForegroundColor Yellow
Write-Host ""

$domainUnlocked = $false
$authCode = $null
$ticketCreated = $false
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Proceed with the action."
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Cancel the action."
$unlockChoice = $Host.UI.PromptForChoice("Confirm", "Unlock '$Domain' and retrieve auth code for registrar transfer?", @($yes, $no), 1)
if ($unlockChoice -eq 0) {
    try {
        # Remove domain privacy first (required for transfer)
        Write-Host "Removing domain privacy for $Domain..." -ForegroundColor Yellow
        try {
            Remove-GoDaddyPrivacy -Credentials $GoDaddyCreds -Domain $Domain -Confirm:$false
        }
        catch {
            Write-Warning "Could not remove privacy: $($_.Exception.Message)"
            Write-Warning "You may need to remove privacy manually in GoDaddy."
        }

        Unlock-GoDaddyDomain -Credentials $GoDaddyCreds -Domain $Domain -Confirm:$false
        $domainUnlocked = $true

        # Force Cloudflare to refresh transfer eligibility (clears cached WHOIS status)
        Write-Host ""
        Write-Host "Refreshing Cloudflare transfer eligibility for $Domain..." -ForegroundColor Yellow
        try {
            $registrarStatus = Get-CloudflareRegistrarDomain -Credentials $CloudflareCreds -AccountId $cfAccount.id -Domain $Domain
            if ($registrarStatus) {
                Write-Host "  Registry status: $($registrarStatus.registry_statuses)" -ForegroundColor Cyan
                Write-Host "  Current registrar: $($registrarStatus.current_registrar)" -ForegroundColor Cyan
                Write-Host "Cloudflare transfer eligibility refreshed." -ForegroundColor Green
            }
        }
        catch {
            Write-Warning "Could not refresh Cloudflare status: $($_.Exception.Message)"
            Write-Warning "You may need to wait for Cloudflare to detect the unlock, or check the dashboard manually."
        }

        # Warn about .us domains requiring FOA email approval
        if ($Domain -match '\.us$') {
            Write-Host ""
            Write-Host ("!" * 60) -ForegroundColor Yellow
            Write-Host " NOTE: .us domains require Form of Authorization (FOA)" -ForegroundColor Yellow
            Write-Host ("!" * 60) -ForegroundColor Yellow
            Write-Host ""
            Write-Host "After initiating the transfer in Cloudflare, check the" -ForegroundColor Yellow
            Write-Host "registrant contact email for an FOA approval email." -ForegroundColor Yellow
            Write-Host "The transfer will not proceed until the FOA is completed." -ForegroundColor Yellow
        }

        # Retrieve auth code
        Write-Host ""
        Write-Host "Retrieving auth code for $Domain..." -ForegroundColor Yellow
        $authCode = Get-GoDaddyAuthCode -Credentials $GoDaddyCreds -Domain $Domain
        if ($authCode) {
            Write-Host "Auth code retrieved successfully." -ForegroundColor Green
        }
        else {
            Write-Host "Could not retrieve auth code. You may need to get it manually from GoDaddy." -ForegroundColor Yellow
        }

        # Create ConnectWise ticket for domain transfer (with auth code)
        if ($ConnectWiseCreds -and (Get-Module -ListAvailable -Name ConnectWiseManageAPI)) {
            Write-Host ""
            $cwYes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Proceed with the action."
            $cwNo = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Cancel the action."
            $cwChoice = $Host.UI.PromptForChoice("Confirm", "Create a ConnectWise ticket for the domain transfer?", @($cwYes, $cwNo), 0)
            if ($cwChoice -eq 0) {
                # Prompt for contact email
                Write-Host ""
                Write-Host "Use the default company contact, or enter an alternate email?" -ForegroundColor Yellow
                Write-Host "  [1] Default company contact" -ForegroundColor Cyan
                Write-Host "  [2] Enter alternate email address" -ForegroundColor Cyan
                $emailChoice = Read-Host "Select (1 or 2)"
                $contactEmail = $null
                if ($emailChoice -eq '2') {
                    $contactEmail = Read-Host "Enter email address for the ticket contact"
                    if ([string]::IsNullOrWhiteSpace($contactEmail)) {
                        Write-Host "No email entered. Using default company contact." -ForegroundColor Yellow
                        $contactEmail = $null
                    }
                }

                Write-Host ""
                Write-Host "Creating ConnectWise ticket for domain transfer..." -ForegroundColor Yellow
                try {
                    Import-Module ConnectWiseManageAPI -Force
                    Connect-CWM @ConnectWiseCreds

                    # Search for company using Cloudflare subaccount name
                    $companyName = $cfAccount.name
                    Write-Host "  Searching ConnectWise for company: $companyName" -ForegroundColor Cyan
                    $companies = Get-CWMCompany -condition "name like '%$companyName%'"

                    # Find matching company using normalized name comparison
                    $normalizedSearch = $companyName -replace '[^\w\s]', '' -replace '\s+', ' '
                    $matchingCompany = $companies | Where-Object {
                        $normalizedName = $_.name -replace '[^\w\s]', '' -replace '\s+', ' '
                        $normalizedName -eq $normalizedSearch
                    } | Select-Object -First 1

                    if ($matchingCompany) {
                        Write-Host "  Found company: $($matchingCompany.name) (ID: $($matchingCompany.identifier))" -ForegroundColor Cyan

                        # Build ticket description with auth code, links, and reminders
                        $cfDashboardUrl = "https://dash.cloudflare.com/$($cfAccount.id)/$Domain"
                        $ticketDescription = "Transfer domain $Domain from GoDaddy to Cloudflare.`n`n"
                        $ticketDescription += "Domain: $Domain`n"
                        $ticketDescription += "Cloudflare Account ID: $($cfAccount.id)`n"
                        $ticketDescription += "Cloudflare Zone ID: $($cfZone.id)`n"
                        $ticketDescription += "`nCloudflare Dashboard: $cfDashboardUrl`n"
                        if ($itgAssetCreated -and $itgAsset.data) {
                            $itgSubdomain = if ($ITGlueCreds.Subdomain) { $ITGlueCreds.Subdomain } else { 'app' }
                            $itgAssetUrl = "https://$itgSubdomain.itglue.com/$itgOrgId/assets/records/$($itgAsset.data.id)"
                            $ticketDescription += "`nITGlue Asset: $itgAssetUrl`n"
                        }
                        if ($authCode) {
                            $ticketDescription += "`nAuth Code: $authCode`n"
                        }
                        else {
                            $ticketDescription += "`nAuth Code: COULD NOT RETRIEVE - obtain manually from GoDaddy`n"
                        }
                        $ticketDescription += "`nDomain has been unlocked at GoDaddy and is ready for transfer."
                        if ($itgAssetCreated) {
                            $ticketDescription += "`n`nIMPORTANT: After the registrar transfer is complete, update the ITGlue asset to set 'Registrar' to true."
                        }

                        $ticketParams = @{
                            summary = "Transfer $Domain to Cloudflare"
                            company = @{identifier = $matchingCompany.identifier}
                            initialDescription = $ticketDescription
                        }
                        if ($contactEmail) {
                            $ticketParams.contactEmailAddress = $contactEmail
                            Write-Host "  Ticket email set to: $contactEmail" -ForegroundColor Cyan
                        }

                        $ticket = New-CWMTicket @ticketParams
                        $ticketCreated = $true
                        Write-Host "Ticket #$($ticket.id) created: $($ticket.summary)" -ForegroundColor Green
                    } else {
                        Write-Host "  Company '$companyName' not found in ConnectWise." -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host "Failed to create ConnectWise ticket: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            else {
                Write-Host "Skipped ConnectWise ticket creation." -ForegroundColor Yellow
            }
        }
    }
    catch {
        Write-Host "Failed to unlock domain: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "You will need to unlock the domain manually at GoDaddy." -ForegroundColor Yellow
    }
}
else {
    Write-Host "Skipped domain unlock. Domain remains locked at GoDaddy." -ForegroundColor Yellow
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
if ($authCode) {
    Write-Host "Auth Code:    Retrieved" -ForegroundColor Green
}
if ($ticketCreated) {
    Write-Host "CW Ticket:      #$($ticket.id)" -ForegroundColor Green
}
if ($itgAssetCreated) {
    Write-Host "ITGlue Asset:   Created" -ForegroundColor Green
}
Write-Host ""
Write-Host "NEXT STEPS:" -ForegroundColor Yellow
$stepNum = 1

if (-not $nameserversUpdated) {
    Write-Host "$stepNum. Update nameservers at GoDaddy to:" -ForegroundColor White
    $cfZone.name_servers | ForEach-Object { Write-Host "   $_" -ForegroundColor Cyan }
    $stepNum++
}

Write-Host "$stepNum. Wait for DNS propagation (up to 48 hours)" -ForegroundColor White
$stepNum++
Write-Host "$stepNum. Verify zone is active in Cloudflare dashboard" -ForegroundColor White
$stepNum++

if ($domainUnlocked) {
    Write-Host "$stepNum. Transfer domain registration via Cloudflare dashboard" -ForegroundColor White
    Write-Host "   (Domain is unlocked and ready for transfer)" -ForegroundColor Cyan
}
elseif ($nameserversUpdated) {
    Write-Host "$stepNum. (Optional) Transfer registration: unlock domain at GoDaddy first" -ForegroundColor Gray
}
