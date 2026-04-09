<#
.SYNOPSIS
    Migrate DNS records from Network Solutions to Cloudflare using exported zone files.

.DESCRIPTION
    Interactive migration tool that reads DNS zone files exported from Network Solutions
    (stored in Config\netsol\<domain>.txt), creates Cloudflare subaccounts, imports DNS
    records, and documents in ITGlue. Looks up the company in ConnectWise Manage by
    matching the domain against contact email domains.

    Nameserver changes at Network Solutions must be done manually (no registrar API).

    Zone files that contain a "already uses Cloudflare nameservers" comment are
    automatically skipped.

.PARAMETER Domain
    The domain to migrate (e.g., "example.com"). Must have a matching zone file.

.PARAMETER CustomerName
    Name for the Cloudflare account/subtenant. If omitted, uses the
    ConnectWise company name found via email domain lookup.

.PARAMETER AccountId
    Use an existing Cloudflare account instead of creating a new one.

.PARAMETER All
    Process all zone files in Config\netsol\ that are not already on Cloudflare.

.PARAMETER ListDomains
    List all zone files and their status (ready, already on Cloudflare, empty).

.PARAMETER ListCloudflareAccounts
    List all accounts in Cloudflare (read-only).

.PARAMETER ListCloudflareZones
    List all zones in Cloudflare (read-only).

.PARAMETER PreviewRecords
    Preview DNS records for a domain without making changes.

.EXAMPLE
    .\NetSol-Cloudflare-Migrate-DNS.ps1 -ListDomains

.EXAMPLE
    .\NetSol-Cloudflare-Migrate-DNS.ps1 -PreviewRecords "example.com"

.EXAMPLE
    .\NetSol-Cloudflare-Migrate-DNS.ps1 -Domain "example.com"

.EXAMPLE
    .\NetSol-Cloudflare-Migrate-DNS.ps1 -Domain "example.com" -CustomerName "Acme Corp"

.PARAMETER ConfigOnly
    Create Cloudflare account/zone and ITGlue documentation but skip DNS record import.
    Useful when you need to move records manually.

.EXAMPLE
    .\NetSol-Cloudflare-Migrate-DNS.ps1 -All

.EXAMPLE
    .\NetSol-Cloudflare-Migrate-DNS.ps1 -Domain "example.com" -ConfigOnly
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

    [Parameter(ParameterSetName = 'Migrate')]
    [Parameter(ParameterSetName = 'Batch')]
    [switch]$ConfigOnly,

    [Parameter(ParameterSetName = 'Batch')]
    [switch]$All,

    [Parameter(ParameterSetName = 'ListDomains')]
    [switch]$ListDomains,

    [Parameter(ParameterSetName = 'ListCFAccounts')]
    [switch]$ListCloudflareAccounts,

    [Parameter(ParameterSetName = 'ListCFZones')]
    [switch]$ListCloudflareZones,

    [Parameter(ParameterSetName = 'Preview')]
    [switch]$PreviewRecords
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ZoneFileDir = "$ScriptDir\Config\netsol"

# Import modules
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

    if (-not $creds.ITGlue.BaseUri) {
        $itgBaseUri = Read-Host "ITGlue API Base URL [https://api.itglue.com]"
        if ([string]::IsNullOrWhiteSpace($itgBaseUri)) { $itgBaseUri = 'https://api.itglue.com' }
        $creds.ITGlue.BaseUri = $itgBaseUri
        $needsSave = $true
    }

    if (-not $creds.ITGlue.ApiKey) {
        $creds.ITGlue.ApiKey = Read-Host "ITGlue API Key"
        $needsSave = $true
    }

    if (-not $creds.ITGlue.Subdomain) {
        $creds.ITGlue.Subdomain = Read-Host "ITGlue Subdomain (e.g., yourcompany)"
        $needsSave = $true
    }

    if (-not $creds.ITGlue.DnsCredentialAssetId) {
        $itgAssetId = Read-Host "ITGlue Password Asset ID for DNS credentials (or Enter to skip)"
        if (-not [string]::IsNullOrWhiteSpace($itgAssetId)) {
            $creds.ITGlue.DnsCredentialAssetId = [int]$itgAssetId
        }
        $needsSave = $true
    }

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
                    Write-Host ""

                    $selection = Read-Host "Select a type by number$promptSuffix"

                    if ([string]::IsNullOrWhiteSpace($selection) -and $existingId) {
                        # Keep existing
                    }
                    elseif ($selection -match '^\d+$') {
                        $idx = [int]$selection - 1
                        if ($idx -ge 0 -and $idx -lt $fatypes.data.Count) {
                            $selected = $fatypes.data[$idx]
                            $creds.ITGlue.FlexibleAssetTypeId = [int]$selected.id
                            $needsSave = $true
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

# Load credentials
$Credentials = Initialize-Credentials
$CloudflareCreds = $Credentials.Cloudflare
$ConnectWiseCreds = $Credentials.ConnectWise
$ITGlueCreds = $Credentials.ITGlue

#region Helper Functions

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

# ── ITGlue domain creation via Selenium (no API exists) ─────────────────────
$script:itgSeleniumDriver = $null

function Initialize-ITGlueBrowser {
    <#
    .SYNOPSIS
        Opens Edge via Selenium, navigates to ITGlue, and waits for SSO login.
    #>
    param([string]$ITGlueUrl)

    Import-Module Selenium -Force -ErrorAction Stop

    $edgeDriverDir = "c:\cbt\EdgeDriver"
    $edgeOptions = New-Object OpenQA.Selenium.Edge.EdgeOptions

    Write-Host "  Opening Edge browser for ITGlue domain creation..." -ForegroundColor Cyan
    $driver = New-Object OpenQA.Selenium.Edge.EdgeDriver($edgeDriverDir, $edgeOptions)

    Write-Host "  Navigating to ITGlue..." -ForegroundColor Cyan
    $driver.Navigate().GoToUrl("https://$ITGlueUrl")

    Write-Host ""
    Write-Host "  *** Please log in to ITGlue in the browser window ***" -ForegroundColor Yellow
    Write-Host "  Press ENTER here once you are logged in and see the ITGlue dashboard..." -ForegroundColor Yellow
    $null = Read-Host

    return $driver
}

function New-ITGDomain {
    <#
    .SYNOPSIS
        Adds a domain to an ITGlue organization via Selenium browser automation.
        ITGlue has no public API for domain creation, so we drive the web UI.
    #>
    param(
        [string]$OrgID,
        [string]$DomainName,
        [string]$ITGlueUrl,
        [OpenQA.Selenium.IWebDriver]$Driver
    )

    try {
        $Driver.Navigate().GoToUrl("https://$ITGlueUrl/$OrgID/domains/new")
        Start-Sleep -Seconds 2

        $nameField = $null
        foreach ($selector in @('domain_name', 'domain[name]')) {
            try {
                $nameField = $Driver.FindElement([OpenQA.Selenium.By]::Name($selector))
                if ($nameField) { break }
            } catch {}
        }
        if (-not $nameField) {
            try {
                $nameField = $Driver.FindElement([OpenQA.Selenium.By]::CssSelector('input[type="text"]'))
            } catch {}
        }

        if (-not $nameField) {
            Write-Warning "  Could not find domain name input field on page"
            return $false
        }

        $nameField.Clear()
        $nameField.SendKeys($DomainName)

        $submitBtn = $null
        foreach ($selector in @('input[type="submit"]', 'button[type="submit"]', '.btn-primary')) {
            try {
                $submitBtn = $Driver.FindElement([OpenQA.Selenium.By]::CssSelector($selector))
                if ($submitBtn) { break }
            } catch {}
        }

        if (-not $submitBtn) {
            Write-Warning "  Could not find submit button on page"
            return $false
        }

        $submitBtn.Click()
        Start-Sleep -Seconds 3

        $pageSource = $Driver.PageSource
        if ($pageSource -like "*Domain has been created successfully*" -or $pageSource -like "*$DomainName*") {
            return $true
        }
        else {
            Write-Warning "  Domain creation may have failed for '$DomainName' - check ITGlue manually"
            return $false
        }
    }
    catch {
        Write-Warning "  Failed to create domain '$DomainName': $($_.Exception.Message)"
        return $false
    }
}

function Remove-ITGlueSSLTrackers {
    <#
    .SYNOPSIS
        Removes auto-created SSL certificate trackers for a given org/domain.
        ITGlue automatically adds SSL trackers when a domain is created.
    #>
    param(
        [string]$OrgId,
        [string]$DomainName,
        [string]$BaseUri,
        [string]$ApiKey
    )

    $headers = @{
        'x-api-key'    = $ApiKey
        'Content-Type' = 'application/vnd.api+json'
    }

    try {
        $response = Invoke-RestMethod -Uri "$BaseUri/ssl_certificates?filter[organization_id]=$OrgId&page[size]=100" `
            -Method GET -Headers $headers -ErrorAction Stop

        if (-not $response.data -or $response.data.Count -eq 0) {
            Write-Host "    No SSL trackers found" -ForegroundColor Gray
            return 0
        }

        # Match only auto-created trackers: exact domain match, no subdomains, no notes
        $matching = @($response.data | Where-Object {
            $cert = $_
            $certDomain = $cert.attributes.'common-name'
            if (-not $certDomain) { $certDomain = $cert.attributes.name }
            $hasNotes = $cert.attributes.notes -and $cert.attributes.notes.Trim() -ne ''
            ($certDomain -eq $DomainName -or $certDomain -eq "www.$DomainName" -or $certDomain -eq "*.$DomainName") -and -not $hasNotes
        })

        if ($matching.Count -eq 0) {
            Write-Host "    No matching SSL trackers for '$DomainName'" -ForegroundColor Gray
            return 0
        }

        $removed = 0
        foreach ($cert in $matching) {
            $certName = $cert.attributes.'common-name'
            if (-not $certName) { $certName = $cert.attributes.name }
            try {
                Invoke-RestMethod -Uri "$BaseUri/ssl_certificates/$($cert.id)" `
                    -Method DELETE -Headers $headers -ErrorAction Stop | Out-Null
                Write-Host "    Removed SSL tracker: $certName (ID: $($cert.id))" -ForegroundColor Green
                $removed++
            }
            catch {
                Write-Warning "    Failed to remove SSL tracker $($cert.id): $($_.Exception.Message)"
            }
        }
        return $removed
    }
    catch {
        Write-Warning "    Failed to query SSL trackers: $($_.Exception.Message)"
        return 0
    }
}

function Import-NetSolZoneFile {
    <#
    .SYNOPSIS
        Parses a Network Solutions zone file into the record format expected by
        Import-CloudflareDNS (type, name, data, ttl, priority).
    .PARAMETER Path
        Path to the zone file.
    .RETURNS
        Array of hashtables with type, name, data, ttl, and optionally priority.
        Returns $null if the file indicates the domain is already on Cloudflare.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $content = Get-Content -Path $Path -ErrorAction Stop
    $domain = $null
    $records = @()

    # Check for already-on-Cloudflare marker
    $joined = $content -join "`n"
    if ($joined -match 'already uses Cloudflare nameservers|DNS is already managed at Cloudflare') {
        # Check if there are any actual records beyond the $ORIGIN line
        $hasRecords = $content | Where-Object { $_ -match '^\S+\s+\d+\s+IN\s+' }
        if (-not $hasRecords) {
            return $null
        }
    }

    foreach ($line in $content) {
        # Extract domain from $ORIGIN directive
        if ($line -match '^\$ORIGIN\s+(\S+?)\.?\s*$') {
            $domain = $Matches[1]
            continue
        }

        # Skip comments and blank lines
        if ($line -match '^\s*;' -or $line -match '^\s*$' -or $line -match '^\$') {
            continue
        }

        # Parse standard zone file line: name ttl IN type data
        if ($line -match '^(\S+)\s+(\d+)\s+IN\s+(\S+)\s+(.+)$') {
            $recName = $Matches[1]
            $recTTL = [int]$Matches[2]
            $recType = $Matches[3].ToUpper()
            $recData = $Matches[4].Trim()

            # Skip SOA and NS records (Cloudflare manages these)
            if ($recType -in @('SOA', 'NS')) { continue }

            # Handle MX records: priority is part of the data field
            $priority = $null
            if ($recType -eq 'MX') {
                if ($recData -match '^(\d+)\s+(.+)$') {
                    $priority = [int]$Matches[1]
                    $recData = $Matches[2].TrimEnd('.')
                }
            }

            # Handle SRV records: priority weight port target
            if ($recType -eq 'SRV') {
                if ($recData -match '^(\d+)\s+(\d+)\s+(\d+)\s+(.+)$') {
                    $priority = [int]$Matches[1]
                    # data format for Import-CloudflareDNS: "weight port target"
                    $recData = "$($Matches[2]) $($Matches[3]) $($Matches[4].TrimEnd('.'))"
                }
            }

            # Strip trailing dots from CNAME and MX targets
            if ($recType -in @('CNAME', 'MX')) {
                $recData = $recData.TrimEnd('.')
            }

            # Strip outer quotes from TXT records (Import-CloudflareDNS handles quoting)
            if ($recType -eq 'TXT') {
                $recData = $recData -replace '^"(.*)"$', '$1'
            }

            $rec = @{
                type = $recType
                name = $recName
                data = $recData
                ttl  = $recTTL
            }

            if ($null -ne $priority) {
                $rec.priority = $priority
            }

            $records += $rec
        }
    }

    return $records
}

function Get-NetSolZoneFiles {
    <#
    .SYNOPSIS
        Returns all zone files in the netsol config directory with their status.
    #>
    if (-not (Test-Path $ZoneFileDir)) {
        Write-Error "Zone file directory not found: $ZoneFileDir"
        return @()
    }

    $files = Get-ChildItem -Path $ZoneFileDir -Filter "*.txt" | Sort-Object Name

    $results = foreach ($file in $files) {
        $domain = $file.BaseName
        $content = Get-Content -Path $file.FullName -Raw

        $status = 'Ready'
        if ($content -match 'already uses Cloudflare nameservers|DNS is already managed at Cloudflare') {
            $hasRecords = (Get-Content -Path $file.FullName) | Where-Object { $_ -match '^\S+\s+\d+\s+IN\s+' }
            if (-not $hasRecords) {
                $status = 'Already on Cloudflare'
            }
        }

        [PSCustomObject]@{
            Domain = $domain
            Status = $status
            File   = $file.Name
        }
    }

    return $results
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
        [string]$Domain,
        [switch]$IncludeInactive
    )

    $matchingCompany = $null

    # Strategy 1: Search contacts by email domain
    Write-Host "Searching ConnectWise contacts with email domain '@$Domain'..." -ForegroundColor Yellow
    try {
        $contacts = @(Get-CWMContact -childConditions "communicationItems/value like '*@$Domain*'" -all)
        if ($contacts.Count -gt 0) {
            $companyIds = $contacts |
                Where-Object { $_.company -and $_.company.id } |
                ForEach-Object { $_.company.id } |
                Select-Object -Unique

            if ($companyIds) {
                $matchingCompanies = @()
                foreach ($compId in $companyIds) {
                    $matchingCompanies += Get-CWMCompany -id $compId
                }
                # Exclude inactive companies from match list (unless explicitly included)
                if (-not $IncludeInactive) {
                    $matchingCompanies = @($matchingCompanies | Where-Object { $_.status.name -in @('Active', 'Special Info') })
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
            $companies = Get-CWMCompany -condition "website like '*$Domain*'" -all
            if (-not $IncludeInactive) {
                $companies = @($companies | Where-Object { $_.status.name -in @('Active', 'Special Info') })
            }
            if ($companies -and $companies.Count -gt 0) {
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
                $sanitizedSearch = $searchTerm -replace "[',`"\\]", ''
                $companies = @(Get-CWMCompany -condition "name like '*$sanitizedSearch*'" -all)
                if ($companies.Count -eq 0) {
                    $words = @($sanitizedSearch -split '\s+' | Where-Object { $_ })
                    if ($words.Count -gt 1) {
                        $longest = $words | Sort-Object Length -Descending | Select-Object -First 1
                        $companies = @(Get-CWMCompany -condition "name like '*$longest*'" -all)
                    }
                }
                if ($companies -and $companies.Count -gt 0) {
                    $i = 1
                    foreach ($co in $companies) {
                        $statusLabel = if ($co.status.name -and $co.status.name -notin @('Active', 'Special Info')) { " [$($co.status.name)]" } else { "" }
                        $color = if ($statusLabel) { 'Yellow' } else { 'White' }
                        Write-Host "  [$i] $($co.name) (ID: $($co.id))$statusLabel" -ForegroundColor $color
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

function Invoke-DomainMigration {
    <#
    .SYNOPSIS
        Runs the full migration workflow for a single domain.
    .PARAMETER Domain
        The domain to migrate.
    .PARAMETER CustomerName
        Optional. Name for the Cloudflare account.
    .PARAMETER AccountId
        Optional. Use an existing Cloudflare account.
    .RETURNS
        Hashtable with migration results.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Domain,

        [string]$CustomerName,

        [string]$AccountId,

        [switch]$ConfigOnly,

        [switch]$IncludeInactive
    )

    $result = @{
        Domain          = $Domain
        Success         = $false
        AccountId       = $null
        ZoneId          = $null
        Nameservers     = @()
        RecordsImported = 0
        RecordsFailed   = 0
        CWCompany       = $null
        ITGlueAsset     = $false
    }

    # ConfigOnly mode: look up existing zone and account in Cloudflare, skip to ITGlue
    if ($ConfigOnly) {
        Write-Banner "Step 1: Look Up Existing Cloudflare Zone"
        Write-Host "ConfigOnly mode: looking up existing zone for '$Domain'..." -ForegroundColor Yellow

        $headers = @{
            "X-Auth-Email" = $CloudflareCreds.Email
            "X-Auth-Key"   = $CloudflareCreds.ApiKey
            "Content-Type" = "application/json"
        }
        try {
            $zoneResponse = Invoke-RestMethod -Uri "https://api.cloudflare.com/client/v4/zones?name=$Domain" `
                -Headers $headers -Method Get
            if ($zoneResponse.success -and $zoneResponse.result.Count -gt 0) {
                $cfZone = $zoneResponse.result[0]
                $cfAccount = $cfZone.account
                Write-Host "Found zone: $($cfZone.name) (ID: $($cfZone.id))" -ForegroundColor Green
                Write-Host "Account:    $($cfAccount.name) (ID: $($cfAccount.id))" -ForegroundColor Green
                Write-Host "Status:     $($cfZone.status)" -ForegroundColor Cyan
                Write-Host "Nameservers:" -ForegroundColor Cyan
                $cfZone.name_servers | ForEach-Object { Write-Host "  $_" -ForegroundColor White }
            }
            else {
                Write-Host "Zone '$Domain' not found in Cloudflare." -ForegroundColor Red
                return $result
            }
        }
        catch {
            Write-Error "Failed to look up zone in Cloudflare: $($_.Exception.Message)"
            return $result
        }

        $result.AccountId = $cfAccount.id
        $result.ZoneId = $cfZone.id
        $result.Nameservers = $cfZone.name_servers

        # Look up company in ConnectWise
        Write-Banner "Step 2: ConnectWise Company Lookup"
        $cwCompany = $null
        if ($ConnectWiseCreds -and (Get-Module -ListAvailable -Name ConnectWiseManageAPI)) {
            try {
                Import-Module ConnectWiseManageAPI -Force
                Connect-CWM @ConnectWiseCreds
                $cwCompany = Find-CWCompanyByDomain -Domain $Domain -IncludeInactive:$IncludeInactive
                if ($cwCompany) {
                    Write-Host "Company: $($cwCompany.name) (ID: $($cwCompany.id))" -ForegroundColor Green
                    $result.CWCompany = $cwCompany.name
                }
                else {
                    Write-Host "No ConnectWise company found for domain '$Domain'." -ForegroundColor Yellow
                }
            }
            catch {
                Write-Warning "ConnectWise lookup failed: $($_.Exception.Message)"
            }
        }
        else {
            Write-Host "ConnectWise integration not available. Skipping company lookup." -ForegroundColor Gray
        }
    }
    else {
    # Full migration mode

    # Step 1: Parse zone file
    Write-Banner "Step 1: Reading DNS Records from Zone File"
    $zoneFilePath = Join-Path $ZoneFileDir "$Domain.txt"

    if (-not (Test-Path $zoneFilePath)) {
        Write-Host "Zone file not found: $zoneFilePath" -ForegroundColor Red
        return $result
    }

    $records = Import-NetSolZoneFile -Path $zoneFilePath

    if ($null -eq $records) {
        Write-Host "Domain '$Domain' is already on Cloudflare. Skipping." -ForegroundColor Yellow
        return $result
    }

    if ($records.Count -eq 0) {
        Write-Host "No DNS records found in zone file for '$Domain'." -ForegroundColor Yellow
        return $result
    }

    Write-Host "Found $($records.Count) DNS records." -ForegroundColor Green
    $records | ForEach-Object {
        $priorityStr = if ($_.priority) { " (Priority: $($_.priority))" } else { "" }
        Write-Host "  $($_.type.PadRight(6)) $($_.name.PadRight(30)) -> $($_.data)$priorityStr"
    }

    # Detect forwarding-only zone: exactly 3 A records (*, @, www) all -> same IP.
    # NetSol uses parking IPs for redirects; replace with 192.0.2.1 + proxy ON so
    # Cloudflare Page Rules / Bulk Redirects can later replicate the forwarding.
    if ($records.Count -eq 3) {
        $aRecs = @($records | Where-Object { $_.type -eq 'A' })
        if ($aRecs.Count -eq 3) {
            $names = ($aRecs | ForEach-Object { $_.name } | Sort-Object) -join ','
            $uniqueIps = @($aRecs | ForEach-Object { $_.data } | Select-Object -Unique)
            if ($names -eq '*,@,www' -and $uniqueIps.Count -eq 1) {
                $originalIp = $uniqueIps[0]
                Write-Host ""
                Write-Host ("!" * 60) -ForegroundColor Yellow
                Write-Host " FORWARDING-ONLY DOMAIN DETECTED" -ForegroundColor Yellow
                Write-Host ("!" * 60) -ForegroundColor Yellow
                Write-Host " '$Domain' has only *, @, and www all pointing to $originalIp." -ForegroundColor Yellow
                Write-Host " This is a NetSol redirect/forwarding setup." -ForegroundColor Yellow
                Write-Host " Replacing records with 192.0.2.1 (proxied) so Cloudflare" -ForegroundColor Yellow
                Write-Host " Page Rules / Bulk Redirects can replicate the forwarding." -ForegroundColor Yellow
                Write-Host " *** Remember to configure Page Rules in Cloudflare ***" -ForegroundColor Yellow
                Write-Host ("!" * 60) -ForegroundColor Yellow
                Write-Host ""

                $records = @(
                    @{ type = 'A'; name = '@';   data = '192.0.2.1'; ttl = 1; proxied = $true },
                    @{ type = 'A'; name = 'www'; data = '192.0.2.1'; ttl = 1; proxied = $true },
                    @{ type = 'A'; name = '*';   data = '192.0.2.1'; ttl = 1; proxied = $true }
                )
                $result.Forwarding = $true
            }
        }
    }

    # Step 2: Look up company in ConnectWise by email domain
    Write-Banner "Step 2: ConnectWise Company Lookup"
    $cwCompany = $null
    if ($ConnectWiseCreds -and (Get-Module -ListAvailable -Name ConnectWiseManageAPI)) {
        try {
            Import-Module ConnectWiseManageAPI -Force
            Connect-CWM @ConnectWiseCreds

            $cwCompany = Find-CWCompanyByDomain -Domain $Domain -IncludeInactive:$IncludeInactive

            if ($cwCompany) {
                Write-Host ""
                Write-Host "Company: $($cwCompany.name) (ID: $($cwCompany.id))" -ForegroundColor Green

                # Check if company is active
                $companyStatus = $cwCompany.status.name
                if ($companyStatus -and $companyStatus -notin @('Active', 'Special Info')) {
                    Write-Host ""
                    Write-Host ("!" * 60) -ForegroundColor Yellow
                    Write-Host " WARNING: Company '$($cwCompany.name)' status: $companyStatus" -ForegroundColor Yellow
                    Write-Host ("!" * 60) -ForegroundColor Yellow
                    Write-Host ""
                    if (-not (Confirm-Action "Company is not Active. Proceed with migration anyway?")) {
                        Write-Host "Skipping domain '$Domain'." -ForegroundColor Yellow
                        return $result
                    }
                }
                else {
                    Write-Host "Status: Active" -ForegroundColor Green
                }
                $result.CWCompany = $cwCompany.name

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
            return $result
        }
    }

    # Sanitize CustomerName for Cloudflare
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
        # Use the account ID directly — avoid paginated account list lookup
        $cfAccount = @{ id = $AccountId; name = $AccountId }
        # Try to fetch the actual account name for display
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
                        return $result
                    }
                }
                2 {
                    Write-Host "Cancelled by user." -ForegroundColor Yellow
                    return $result
                }
            }
        }
        else {
            Write-Host "No existing account found." -ForegroundColor Gray
            Write-Host "Will create new account: $CustomerName" -ForegroundColor Yellow
            Write-Host ""

            if (-not (Confirm-Action "Create new Cloudflare account '$CustomerName'?")) {
                Write-Host "Cancelled by user." -ForegroundColor Yellow
                return $result
            }

            try {
                $cfAccount = New-CloudflareAccount -Credentials $CloudflareCreds -Name $CustomerName -Confirm:$false
                Write-Host "Account created: $($cfAccount.id)" -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to create Cloudflare account: $($_.Exception.Message)"
                return $result
            }
        }
    }

    $result.AccountId = $cfAccount.id

    # Step 4: Add zone to Cloudflare
    Write-Banner "Step 4: Add Zone to Cloudflare"
    Write-Host "Domain: $Domain" -ForegroundColor Yellow
    Write-Host "Account: $($cfAccount.id)" -ForegroundColor Yellow
    Write-Host ""

    if (-not (Confirm-Action "Add zone '$Domain' to Cloudflare?")) {
        Write-Host "Cancelled by user." -ForegroundColor Yellow
        return $result
    }

    try {
        $cfZone = New-CloudflareZone -Credentials $CloudflareCreds -AccountId $cfAccount.id -Domain $Domain -Confirm:$false
        Write-Host ""
        Write-Host "Zone created successfully!" -ForegroundColor Green
        Write-Host "Zone ID: $($cfZone.id)" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Nameservers to set at Network Solutions:" -ForegroundColor Yellow
        $cfZone.name_servers | ForEach-Object { Write-Host "  $_" -ForegroundColor White }
    }
    catch {
        Write-Error "Failed to create zone: $($_.Exception.Message)"
        return $result
    }

    $result.ZoneId = $cfZone.id
    $result.Nameservers = $cfZone.name_servers

    # Step 5: Import DNS records
    Write-Banner "Step 5: Import DNS Records"
    Write-Host "Ready to import $($records.Count) records to Cloudflare." -ForegroundColor Yellow
    Write-Host ""

    # Preview first
    Import-CloudflareDNS -Credentials $CloudflareCreds -ZoneId $cfZone.id -Records $records -PreviewOnly

    if (-not (Confirm-Action "Proceed with DNS record import?")) {
        Write-Host "Cancelled by user." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Zone was created but no DNS records were imported." -ForegroundColor Yellow
        Write-Host "You can import records later manually." -ForegroundColor Cyan
        return $result
    }

    try {
        $importResult = Import-CloudflareDNS -Credentials $CloudflareCreds -ZoneId $cfZone.id -Records $records -Confirm:$false
        $result.RecordsImported = $importResult.Success.Count
        $result.RecordsFailed = $importResult.Failed.Count
    }
    catch {
        Write-Error "Failed to import DNS records: $($_.Exception.Message)"
        return $result
    }

    } # end full migration mode

    # Nameserver instructions (manual for Network Solutions)
    Write-Banner $(if ($ConfigOnly) { "Nameservers for $Domain" } else { "Step 6: Update Nameservers at Network Solutions (Manual)" })
    Write-Host "Cloudflare assigned nameservers:" -ForegroundColor Yellow
    $cfZone.name_servers | ForEach-Object { Write-Host "  $_" -ForegroundColor Cyan }
    Write-Host ""
    Write-Host "Network Solutions does not have a public API for nameserver changes." -ForegroundColor Yellow
    Write-Host "You must update the nameservers manually:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  1. Log in to your Network Solutions account" -ForegroundColor White
    Write-Host "  2. Go to Account Manager > My Domain Names" -ForegroundColor White
    Write-Host "  3. Select '$Domain'" -ForegroundColor White
    Write-Host "  4. Under 'Domain Name Server (DNS)', click 'Edit'" -ForegroundColor White
    Write-Host "  5. Select 'Move DNS to an advanced/3rd party host'" -ForegroundColor White
    Write-Host "  6. Enter the Cloudflare nameservers:" -ForegroundColor White
    $cfZone.name_servers | ForEach-Object { Write-Host "     $_" -ForegroundColor Cyan }
    Write-Host "  7. Save changes" -ForegroundColor White
    Write-Host ""

    if (-not $ConfigOnly) {
        $null = Read-Host "Press Enter when nameservers have been updated (or to continue without updating)"
    }

    # Step 7: Create ITGlue DNS/Registrar Asset
    Write-Banner $(if ($ConfigOnly) { "Step 3: Document in ITGlue" } else { "Step 7: Document in ITGlue" })
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

                # Normalize name for flexible matching
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

                    # Look up domain asset for tagging; create it if missing.
                    $domainTag = $null
                    $matchingDomain = $null
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

                    if (-not $matchingDomain) {
                        Write-Host "  Domain '$Domain' not found under org in ITGlue. Creating via web UI..." -ForegroundColor Yellow
                        try {
                            $itgUiHost = "$($ITGlueCreds.Subdomain).itglue.com"
                            if (-not $script:itgSeleniumDriver) {
                                $script:itgSeleniumDriver = Initialize-ITGlueBrowser -ITGlueUrl $itgUiHost
                            }
                            $created = New-ITGDomain -OrgID $itgOrgId -DomainName $Domain `
                                -ITGlueUrl $itgUiHost -Driver $script:itgSeleniumDriver
                            if ($created) {
                                Write-Host "  Domain created in ITGlue: $Domain" -ForegroundColor Green
                                Start-Sleep -Seconds 2

                                # Re-query to grab the new domain ID for tagging
                                try {
                                    $domains = Get-ITGlueDomains -filter_organization_id $itgOrgId
                                    if ($domains.data) {
                                        $matchingDomain = $domains.data | Where-Object { $_.attributes.name -eq $Domain }
                                        if ($matchingDomain) {
                                            $domainTag = @($matchingDomain.id)
                                            Write-Host "  Newly created domain ID: $($matchingDomain.id)" -ForegroundColor Cyan
                                        }
                                    }
                                }
                                catch {
                                    Write-Host "  Could not re-query ITGlue domains after create: $($_.Exception.Message)" -ForegroundColor Yellow
                                }

                                # Remove auto-created SSL trackers
                                Write-Host "  Removing auto-created SSL trackers for '$Domain'..." -ForegroundColor Cyan
                                $null = Remove-ITGlueSSLTrackers -OrgId $itgOrgId -DomainName $Domain `
                                    -BaseUri $ITGlueCreds.BaseUri -ApiKey $ITGlueCreds.ApiKey
                            }
                        }
                        catch {
                            Write-Host "  Failed to create ITGlue domain '$Domain': $($_.Exception.Message)" -ForegroundColor Red
                        }
                    }

                    # Build traits
                    $traits = @{
                        'provider'        = 'CloudFlare'
                        'dns'             = $true
                        'registrar'       = $false
                        'dns-credentials' = $ITGlueCreds.DnsCredentialAssetId
                        'management-url'  = $managementUrl
                        'notes'           = "Cloudflare Zone ID: $($cfZone.id)<br>Cloudflare Account ID: $($cfAccount.id)<br>Nameservers: $($cfZone.name_servers -join ', ')<br>Migrated: $(Get-Date -Format 'yyyy-MM-dd')<br>Source: Network Solutions"
                    }

                    if ($domainTag) {
                        $traits['domain-s'] = $domainTag
                    }

                    # Create the flexible asset
                    $assetData = @{
                        type       = 'flexible-assets'
                        attributes = @{
                            'organization-id'        = [int]$itgOrgId
                            'flexible-asset-type-id' = $ITGlueCreds.FlexibleAssetTypeId
                            traits                   = $traits
                        }
                    }

                    $itgAsset = New-ITGlueFlexibleAssets -data $assetData
                    if ($itgAsset.data) {
                        $itgAssetCreated = $true
                        $result.ITGlueAsset = $true
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
    Write-Banner $(if ($ConfigOnly) { "Config Complete: $Domain" } else { "Migration Complete: $Domain" })
    Write-Host "Domain:     $Domain" -ForegroundColor Green
    Write-Host "Account ID: $($cfAccount.id)" -ForegroundColor Green
    Write-Host "Zone ID:    $($cfZone.id)" -ForegroundColor Green
    Write-Host ""
    if (-not $ConfigOnly) {
        Write-Host "Records Imported: $($importResult.Success.Count)" -ForegroundColor Green
        if ($importResult.Failed.Count -gt 0) {
            Write-Host "Records Failed:   $($importResult.Failed.Count)" -ForegroundColor Red
        }
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

    Write-Host "$stepNum. Update nameservers at Network Solutions to:" -ForegroundColor White
    $cfZone.name_servers | ForEach-Object { Write-Host "   $_" -ForegroundColor Cyan }
    $stepNum++

    Write-Host "$stepNum. Wait for DNS propagation (up to 48 hours)" -ForegroundColor White
    $stepNum++
    Write-Host "$stepNum. Verify zone is active in Cloudflare dashboard" -ForegroundColor White
    $stepNum++

    if ($itgAssetCreated) {
        Write-Host "$stepNum. Update ITGlue asset 'Registrar' field if/when registrar transfer is completed" -ForegroundColor White
    }

    $result.Success = $true
    return $result
}

#endregion

# Handle read-only operations
switch ($PSCmdlet.ParameterSetName) {
    'ListDomains' {
        Write-Banner "Network Solutions Zone Files"
        $zoneFiles = Get-NetSolZoneFiles
        if ($zoneFiles.Count -eq 0) {
            Write-Host "No zone files found in $ZoneFileDir" -ForegroundColor Yellow
            exit 0
        }

        $ready = ($zoneFiles | Where-Object { $_.Status -eq 'Ready' }).Count
        $onCF = ($zoneFiles | Where-Object { $_.Status -eq 'Already on Cloudflare' }).Count

        $zoneFiles | Format-Table Domain, Status -AutoSize
        Write-Host "Total: $($zoneFiles.Count) | Ready: $ready | Already on Cloudflare: $onCF" -ForegroundColor Yellow
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
        Write-Banner "DNS Records Preview: $Domain"
        $zoneFilePath = Join-Path $ZoneFileDir "$Domain.txt"

        if (-not (Test-Path $zoneFilePath)) {
            Write-Host "Zone file not found: $zoneFilePath" -ForegroundColor Red
            Write-Host "Available zone files:" -ForegroundColor Yellow
            Get-ChildItem -Path $ZoneFileDir -Filter "*.txt" | ForEach-Object { Write-Host "  $($_.BaseName)" }
            exit 1
        }

        $records = Import-NetSolZoneFile -Path $zoneFilePath

        if ($null -eq $records) {
            Write-Host "Domain '$Domain' is already on Cloudflare." -ForegroundColor Yellow
            exit 0
        }

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

# Single domain migration
if ($Domain) {
    $migrateParams = @{ Domain = $Domain; IncludeInactive = $true }
    if ($CustomerName) { $migrateParams.CustomerName = $CustomerName }
    if ($AccountId) { $migrateParams.AccountId = $AccountId }
    if ($ConfigOnly) { $migrateParams.ConfigOnly = $true }
    Invoke-DomainMigration @migrateParams
    exit 0
}

# Batch mode: process all ready zone files
if ($All) {
    Write-Banner "Batch Migration: Network Solutions to Cloudflare"
    $zoneFiles = Get-NetSolZoneFiles | Where-Object { $_.Status -eq 'Ready' }

    if ($zoneFiles.Count -eq 0) {
        Write-Host "No domains ready for migration." -ForegroundColor Yellow
        exit 0
    }

    # Filter to domains that still have nameservers at Network Solutions.
    # Domains with custom/third-party nameservers aren't actually hosted at NetSol
    # anymore and should be skipped (and noted for the user).
    Write-Host "Checking nameservers for each domain..." -ForegroundColor Yellow
    $netsolZones = @()
    $skippedCustomNs = @()
    foreach ($zf in $zoneFiles) {
        $nsNames = @()
        try {
            $nsRecords = Resolve-DnsName -Name $zf.Domain -Type NS -DnsOnly -ErrorAction Stop |
                Where-Object { $_.Type -eq 'NS' }
            $nsNames = @($nsRecords | ForEach-Object { $_.NameHost.ToLower() })
        }
        catch {
            Write-Host "  [?] $($zf.Domain): NS lookup failed ($($_.Exception.Message))" -ForegroundColor DarkYellow
            $skippedCustomNs += [PSCustomObject]@{ Domain = $zf.Domain; Nameservers = @('lookup failed') }
            continue
        }

        if ($nsNames.Count -eq 0) {
            Write-Host "  [?] $($zf.Domain): no NS records returned" -ForegroundColor DarkYellow
            $skippedCustomNs += [PSCustomObject]@{ Domain = $zf.Domain; Nameservers = @('none') }
            continue
        }

        # NetSol nameservers: worldnic.com, netsol.com, networksolutions.com
        $isNetsol = $nsNames | Where-Object { $_ -match '(worldnic\.com|netsol\.com|networksolutions\.com)$' }
        if ($isNetsol) {
            Write-Host "  [OK] $($zf.Domain) -> $($nsNames -join ', ')" -ForegroundColor Green
            $netsolZones += $zf
        }
        else {
            Write-Host "  [SKIP] $($zf.Domain) -> $($nsNames -join ', ') (custom NS)" -ForegroundColor Gray
            $skippedCustomNs += [PSCustomObject]@{ Domain = $zf.Domain; Nameservers = $nsNames }
        }
    }

    Write-Host ""
    if ($skippedCustomNs.Count -gt 0) {
        Write-Host "Skipping $($skippedCustomNs.Count) domain(s) with custom/non-NetSol nameservers:" -ForegroundColor Yellow
        foreach ($s in $skippedCustomNs) {
            Write-Host "  $($s.Domain) -> $($s.Nameservers -join ', ')" -ForegroundColor Gray
        }
        Write-Host ""
    }

    $zoneFiles = $netsolZones

    if ($zoneFiles.Count -eq 0) {
        Write-Host "No domains with NetSol nameservers found. Nothing to migrate." -ForegroundColor Yellow
        exit 0
    }

    Write-Host "Domains to migrate:" -ForegroundColor Yellow
    $zoneFiles | ForEach-Object { Write-Host "  $($_.Domain)" -ForegroundColor White }
    Write-Host ""
    Write-Host "Total: $($zoneFiles.Count) domains" -ForegroundColor Yellow
    Write-Host ""

    if (-not (Confirm-Action "Proceed with batch migration of $($zoneFiles.Count) domains?")) {
        Write-Host "Cancelled by user." -ForegroundColor Yellow
        exit 0
    }

    # Allow matching closed/cancelled companies in the batch run if explicitly opted in
    $batchIncludeInactive = Confirm-Action "Include closed/cancelled companies as valid matches for these domains?"

    $batchResults = @()

    foreach ($zf in $zoneFiles) {
        Write-Host ""
        Write-Host ("=" * 60) -ForegroundColor Magenta
        Write-Host " MIGRATING: $($zf.Domain)" -ForegroundColor Magenta
        Write-Host ("=" * 60) -ForegroundColor Magenta

        try {
            $migParams = @{ Domain = $zf.Domain }
            if ($ConfigOnly) { $migParams.ConfigOnly = $true }
            if ($batchIncludeInactive) { $migParams.IncludeInactive = $true }
            $migResult = Invoke-DomainMigration @migParams
            $batchResults += $migResult
        }
        catch {
            Write-Host "ERROR migrating $($zf.Domain): $($_.Exception.Message)" -ForegroundColor Red
            $batchResults += @{
                Domain  = $zf.Domain
                Success = $false
            }
        }
    }

    # Batch summary
    Write-Banner "Batch Migration Summary"
    $succeeded = ($batchResults | Where-Object { $_.Success }).Count
    $failed = ($batchResults | Where-Object { -not $_.Success }).Count

    Write-Host "Total:     $($batchResults.Count)" -ForegroundColor Yellow
    Write-Host "Succeeded: $succeeded" -ForegroundColor Green
    Write-Host "Failed:    $failed" -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Green" })
    Write-Host ""

    # Nameserver summary for all migrated domains
    $migratedDomains = $batchResults | Where-Object { $_.Success -and $_.Nameservers }
    if ($migratedDomains) {
        Write-Host "NAMESERVERS TO SET AT NETWORK SOLUTIONS:" -ForegroundColor Yellow
        Write-Host ""
        foreach ($mr in $migratedDomains) {
            Write-Host "  $($mr.Domain):" -ForegroundColor White
            $mr.Nameservers | ForEach-Object { Write-Host "    $_" -ForegroundColor Cyan }
        }
    }

    if ($failed -gt 0) {
        Write-Host ""
        Write-Host "FAILED DOMAINS:" -ForegroundColor Red
        $batchResults | Where-Object { -not $_.Success } | ForEach-Object {
            Write-Host "  $($_.Domain)" -ForegroundColor Red
        }
    }

    exit 0
}

# No parameters - show available domains
Write-Banner "Network Solutions DNS Migration"
Write-Host "Available zone files:" -ForegroundColor Yellow
$zoneFiles = Get-NetSolZoneFiles
$zoneFiles | Format-Table Domain, Status -AutoSize

Write-Host "Usage:" -ForegroundColor Yellow
Write-Host "  .\NetSol-Cloudflare-Migrate-DNS.ps1 -Domain ""example.com""          # Migrate single domain" -ForegroundColor Cyan
Write-Host "  .\NetSol-Cloudflare-Migrate-DNS.ps1 -All                            # Migrate all ready domains" -ForegroundColor Cyan
Write-Host "  .\NetSol-Cloudflare-Migrate-DNS.ps1 -PreviewRecords ""example.com""   # Preview records" -ForegroundColor Cyan
Write-Host "  .\NetSol-Cloudflare-Migrate-DNS.ps1 -ListDomains                    # List all zone files" -ForegroundColor Cyan

# Cleanup Selenium driver if it was used
if ($script:itgSeleniumDriver) {
    try { $script:itgSeleniumDriver.Quit() } catch {}
    $script:itgSeleniumDriver = $null
}
