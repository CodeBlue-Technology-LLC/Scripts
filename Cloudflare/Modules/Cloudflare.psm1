# Cloudflare API Module
# Write operations require -Confirm for safety

$script:CloudflareApiBase = "https://api.cloudflare.com/client/v4"

function Get-CloudflareAuthHeader {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials
    )

    @{
        "X-Auth-Email" = $Credentials.Email
        "X-Auth-Key"   = $Credentials.ApiKey
        "Content-Type" = "application/json"
    }
}

#region Read-Only Operations

function Get-CloudflareAccounts {
    <#
    .SYNOPSIS
        Lists all accounts accessible to the authenticated user.
    .DESCRIPTION
        Read-only operation to retrieve Cloudflare accounts.
    .PARAMETER Credentials
        Hashtable containing Email and ApiKey for Cloudflare.
    .EXAMPLE
        $creds = (Import-Clixml .\Config\credentials.xml).Cloudflare
        Get-CloudflareAccounts -Credentials $creds
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials
    )

    $headers = Get-CloudflareAuthHeader -Credentials $Credentials

    try {
        $response = Invoke-RestMethod -Uri "$script:CloudflareApiBase/accounts" `
            -Headers $headers `
            -Method Get

        if ($response.success) {
            return $response.result
        }
        else {
            throw "API Error: $($response.errors | ConvertTo-Json -Compress)"
        }
    }
    catch {
        Write-Error "Failed to retrieve Cloudflare accounts: $($_.Exception.Message)"
        throw
    }
}

function Get-CloudflareZones {
    <#
    .SYNOPSIS
        Lists all zones in the account.
    .DESCRIPTION
        Read-only operation to retrieve Cloudflare zones.
    .PARAMETER Credentials
        Hashtable containing Email and ApiKey for Cloudflare.
    .PARAMETER AccountId
        Optional. Filter zones by account ID.
    .EXAMPLE
        $creds = (Import-Clixml .\Config\credentials.xml).Cloudflare
        Get-CloudflareZones -Credentials $creds
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials,

        [Parameter()]
        [string]$AccountId
    )

    $headers = Get-CloudflareAuthHeader -Credentials $Credentials

    $uri = "$script:CloudflareApiBase/zones"
    if ($AccountId) {
        $uri += "?account.id=$AccountId"
    }

    try {
        $response = Invoke-RestMethod -Uri $uri `
            -Headers $headers `
            -Method Get

        if ($response.success) {
            return $response.result
        }
        else {
            throw "API Error: $($response.errors | ConvertTo-Json -Compress)"
        }
    }
    catch {
        Write-Error "Failed to retrieve Cloudflare zones: $($_.Exception.Message)"
        throw
    }
}

function Get-CloudflareDNSRecords {
    <#
    .SYNOPSIS
        Lists DNS records for a zone.
    .DESCRIPTION
        Read-only operation to retrieve DNS records from a Cloudflare zone.
    .PARAMETER Credentials
        Hashtable containing Email and ApiKey for Cloudflare.
    .PARAMETER ZoneId
        The zone ID to query.
    .EXAMPLE
        Get-CloudflareDNSRecords -Credentials $creds -ZoneId "abc123"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials,

        [Parameter(Mandatory)]
        [string]$ZoneId
    )

    $headers = Get-CloudflareAuthHeader -Credentials $Credentials

    try {
        $response = Invoke-RestMethod -Uri "$script:CloudflareApiBase/zones/$ZoneId/dns_records" `
            -Headers $headers `
            -Method Get

        if ($response.success) {
            return $response.result
        }
        else {
            throw "API Error: $($response.errors | ConvertTo-Json -Compress)"
        }
    }
    catch {
        Write-Error "Failed to retrieve DNS records: $($_.Exception.Message)"
        throw
    }
}

#endregion

#region Write Operations (Require Confirmation)

function New-CloudflareAccount {
    <#
    .SYNOPSIS
        Creates a new Cloudflare account (subtenant).
    .DESCRIPTION
        Creates a new account under the tenant. Requires confirmation.
    .PARAMETER Credentials
        Hashtable containing Email and ApiKey for Cloudflare.
    .PARAMETER Name
        Name for the new account.
    .PARAMETER Type
        Account type. Default is "standard".
    .EXAMPLE
        New-CloudflareAccount -Credentials $creds -Name "Acme Corp" -Confirm
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter()]
        [ValidateSet("standard", "enterprise")]
        [string]$Type = "standard"
    )

    $headers = Get-CloudflareAuthHeader -Credentials $Credentials

    $body = @{
        name = $Name
        type = $Type
    } | ConvertTo-Json

    if ($PSCmdlet.ShouldProcess("Cloudflare", "Create new account '$Name'")) {
        try {
            $response = Invoke-RestMethod -Uri "$script:CloudflareApiBase/accounts" `
                -Headers $headers `
                -Method Post `
                -Body $body

            if ($response.success) {
                Write-Host "Account '$Name' created successfully." -ForegroundColor Green
                return $response.result
            }
            else {
                throw "API Error: $($response.errors | ConvertTo-Json -Compress)"
            }
        }
        catch {
            Write-Error "Failed to create Cloudflare account: $($_.Exception.Message)"
            throw
        }
    }
}

function New-CloudflareZone {
    <#
    .SYNOPSIS
        Adds a zone (domain) to a Cloudflare account.
    .DESCRIPTION
        Creates a new zone in the specified account. Requires confirmation.
    .PARAMETER Credentials
        Hashtable containing Email and ApiKey for Cloudflare.
    .PARAMETER AccountId
        The account ID to add the zone to.
    .PARAMETER Domain
        The domain name (e.g., "example.com").
    .PARAMETER Type
        Zone type. Default is "full".
    .EXAMPLE
        New-CloudflareZone -Credentials $creds -AccountId "abc123" -Domain "example.com" -Confirm
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials,

        [Parameter(Mandatory)]
        [string]$AccountId,

        [Parameter(Mandatory)]
        [string]$Domain,

        [Parameter()]
        [ValidateSet("full", "partial", "secondary")]
        [string]$Type = "full"
    )

    $headers = Get-CloudflareAuthHeader -Credentials $Credentials

    $body = @{
        name = $Domain
        account = @{
            id = $AccountId
        }
        type = $Type
    } | ConvertTo-Json -Depth 3

    if ($PSCmdlet.ShouldProcess("Cloudflare", "Add zone '$Domain' to account $AccountId")) {
        try {
            $response = Invoke-RestMethod -Uri "$script:CloudflareApiBase/zones" `
                -Headers $headers `
                -Method Post `
                -Body $body

            if ($response.success) {
                Write-Host "Zone '$Domain' created successfully." -ForegroundColor Green
                Write-Host "Nameservers: $($response.result.name_servers -join ', ')" -ForegroundColor Cyan
                return $response.result
            }
            else {
                throw "API Error: $($response.errors | ConvertTo-Json -Compress)"
            }
        }
        catch {
            Write-Error "Failed to create Cloudflare zone: $($_.Exception.Message)"
            throw
        }
    }
}

function New-CloudflareDNSRecord {
    <#
    .SYNOPSIS
        Creates a single DNS record in a Cloudflare zone.
    .DESCRIPTION
        Creates a DNS record. Requires confirmation.
    .PARAMETER Credentials
        Hashtable containing Email and ApiKey for Cloudflare.
    .PARAMETER ZoneId
        The zone ID to add the record to.
    .PARAMETER Type
        Record type (A, AAAA, CNAME, MX, TXT, etc.).
    .PARAMETER Name
        Record name (e.g., "www" or "@" for root).
    .PARAMETER Content
        Record content/value.
    .PARAMETER TTL
        TTL in seconds. 1 = automatic.
    .PARAMETER Priority
        Priority for MX/SRV records.
    .PARAMETER Proxied
        Whether to proxy through Cloudflare (only for A/AAAA/CNAME).
    .EXAMPLE
        New-CloudflareDNSRecord -Credentials $creds -ZoneId "abc" -Type "A" -Name "www" -Content "1.2.3.4"
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials,

        [Parameter(Mandatory)]
        [string]$ZoneId,

        [Parameter(Mandatory)]
        [ValidateSet("A", "AAAA", "CNAME", "MX", "TXT", "NS", "SRV", "CAA", "SPF")]
        [string]$Type,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [string]$Content,

        [Parameter()]
        [int]$TTL = 1,

        [Parameter()]
        [int]$Priority,

        [Parameter()]
        [bool]$Proxied = $false
    )

    $headers = Get-CloudflareAuthHeader -Credentials $Credentials

    $body = @{
        type    = $Type
        name    = $Name
        content = $Content
        ttl     = $TTL
    }

    if ($PSBoundParameters.ContainsKey('Priority') -and $Type -in @("MX", "SRV")) {
        $body.priority = $Priority
    }

    if ($Type -in @("A", "AAAA", "CNAME")) {
        $body.proxied = $Proxied
    }

    $bodyJson = $body | ConvertTo-Json

    if ($PSCmdlet.ShouldProcess("Zone $ZoneId", "Create $Type record '$Name' -> '$Content'")) {
        try {
            $response = Invoke-RestMethod -Uri "$script:CloudflareApiBase/zones/$ZoneId/dns_records" `
                -Headers $headers `
                -Method Post `
                -Body $bodyJson

            if ($response.success) {
                return $response.result
            }
            else {
                throw "API Error: $($response.errors | ConvertTo-Json -Compress)"
            }
        }
        catch {
            Write-Error "Failed to create DNS record: $($_.Exception.Message)"
            throw
        }
    }
}

function Import-CloudflareDNS {
    <#
    .SYNOPSIS
        Imports multiple DNS records to a Cloudflare zone.
    .DESCRIPTION
        Bulk imports DNS records with preview. Requires confirmation.
    .PARAMETER Credentials
        Hashtable containing Email and ApiKey for Cloudflare.
    .PARAMETER ZoneId
        The zone ID to import records to.
    .PARAMETER Records
        Array of DNS records (from GoDaddy format).
    .PARAMETER PreviewOnly
        If set, only shows what would be imported without making changes.
    .EXAMPLE
        $godaddyRecords = Get-GoDaddyDNSRecords -Credentials $gd -Domain "example.com"
        Import-CloudflareDNS -Credentials $cf -ZoneId "abc" -Records $godaddyRecords
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials,

        [Parameter(Mandatory)]
        [string]$ZoneId,

        [Parameter(Mandatory)]
        [array]$Records,

        [Parameter()]
        [switch]$PreviewOnly
    )

    # Convert GoDaddy records to Cloudflare format
    $cfRecords = foreach ($record in $Records) {
        # Skip SOA and NS records (Cloudflare manages these)
        if ($record.type -in @("SOA", "NS")) {
            Write-Verbose "Skipping $($record.type) record (Cloudflare managed)"
            continue
        }

        # Skip parked domain A records (GoDaddy parking/placeholder IPs)
        if ($record.type -eq 'A' -and $record.data -match '^(0\.0\.0\.0|Parked|34\.102\.136\.\d+|50\.63\.202\.\d+|184\.168\.131\.\d+)$') {
            Write-Host "  Skipping parked A record: $($record.name) -> $($record.data)" -ForegroundColor Gray
            continue
        }

        # SRV records require structured data for Cloudflare API
        if ($record.type -eq 'SRV') {
            # GoDaddy SRV: name = "_service._proto", data = "weight port target"
            $srvName = $record.name
            $service = ''
            $proto = ''
            if ($srvName -match '^(_[^.]+)\.(_[^.]+)') {
                $service = $Matches[1]
                $proto = $Matches[2]
            }
            $srvParts = $record.data -split '\s+'
            if ($srvParts.Count -ge 3) {
                @{
                    type     = 'SRV'
                    name     = $srvName
                    ttl      = if ($record.ttl -lt 60) { 1 } else { $record.ttl }
                    data     = @{
                        service  = $service
                        proto    = $proto
                        name     = '@'
                        priority = [int]$record.priority
                        weight   = [int]$srvParts[0]
                        port     = [int]$srvParts[1]
                        target   = $srvParts[2]
                    }
                }
            }
            else {
                Write-Host "  Skipping malformed SRV record: $($record.name) -> $($record.data)" -ForegroundColor Yellow
            }
            continue
        }

        @{
            type     = $record.type
            name     = if ($record.name -eq "@") { "@" } else { $record.name }
            content  = if ($record.type -eq 'TXT' -and $record.data -notmatch '^\s*".*"\s*$') { "`"$($record.data)`"" } else { $record.data }
            ttl      = if ($record.ttl -lt 60) { 1 } else { $record.ttl }
            priority = $record.priority
        }
    }

    # Preview
    Write-Host "`n=== DNS Records to Import ===" -ForegroundColor Cyan
    Write-Host "Total records: $($cfRecords.Count)" -ForegroundColor Yellow
    Write-Host ""

    $cfRecords | ForEach-Object {
        $priorityStr = if ($_.priority) { " (Priority: $($_.priority))" } else { "" }
        Write-Host "  $($_.type.PadRight(6)) $($_.name.PadRight(30)) -> $($_.content)$priorityStr"
    }
    Write-Host ""

    if ($PreviewOnly) {
        Write-Host "Preview only - no changes made." -ForegroundColor Yellow
        return $cfRecords
    }

    if ($PSCmdlet.ShouldProcess("Zone $ZoneId", "Import $($cfRecords.Count) DNS records")) {
        $results = @{
            Success = @()
            Failed  = @()
        }

        foreach ($record in $cfRecords) {
            try {
                # SRV records use a structured data object instead of content
                if ($record.type -eq 'SRV' -and $record.data) {
                    $headers = Get-CloudflareAuthHeader -Credentials $Credentials
                    $body = @{
                        type = 'SRV'
                        data = $record.data
                        ttl  = $record.ttl
                    } | ConvertTo-Json -Depth 3
                    $response = Invoke-RestMethod -Uri "$script:CloudflareApiBase/zones/$ZoneId/dns_records" `
                        -Headers $headers `
                        -Method Post `
                        -Body $body `
                        -ContentType "application/json"
                    if ($response.success) {
                        $results.Success += $record
                        Write-Host "  [OK] $($record.type) $($record.name)" -ForegroundColor Green
                    }
                    else {
                        throw "API Error: $($response.errors | ConvertTo-Json -Compress)"
                    }
                }
                else {
                    $params = @{
                        Credentials = $Credentials
                        ZoneId      = $ZoneId
                        Type        = $record.type
                        Name        = $record.name
                        Content     = $record.content
                        TTL         = $record.ttl
                        Confirm     = $false
                    }

                    if ($null -ne $record.priority -and $record.type -in @("MX")) {
                        $params.Priority = $record.priority
                    }

                    $result = New-CloudflareDNSRecord @params
                    $results.Success += $record
                    Write-Host "  [OK] $($record.type) $($record.name)" -ForegroundColor Green
                }
            }
            catch {
                $results.Failed += @{
                    Record = $record
                    Error  = $_.Exception.Message
                }
                Write-Host "  [FAIL] $($record.type) $($record.name): $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        Write-Host "`n=== Import Summary ===" -ForegroundColor Cyan
        Write-Host "Success: $($results.Success.Count)" -ForegroundColor Green
        Write-Host "Failed:  $($results.Failed.Count)" -ForegroundColor $(if ($results.Failed.Count -gt 0) { "Red" } else { "Green" })

        return $results
    }
}

#endregion

function Get-CloudflareRegistrarDomain {
    <#
    .SYNOPSIS
        Queries Cloudflare's registrar API for a domain's transfer eligibility.
    .DESCRIPTION
        Calls GET /accounts/{account_id}/registrar/domains/{domain} which forces
        Cloudflare to refresh its cached WHOIS/RDAP data for the domain. This is
        needed after unlocking a domain at GoDaddy, as Cloudflare may cache the
        old locked status and gray out the transfer option until refreshed.
    .PARAMETER Credentials
        Hashtable containing Email and ApiKey for Cloudflare.
    .PARAMETER AccountId
        The Cloudflare account ID that owns the zone.
    .PARAMETER Domain
        The domain name to check (e.g., "example.com").
    .EXAMPLE
        $creds = (Import-Clixml .\Config\credentials.xml).Cloudflare
        Get-CloudflareRegistrarDomain -Credentials $creds -AccountId "abc123" -Domain "example.com"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials,

        [Parameter(Mandatory)]
        [string]$AccountId,

        [Parameter(Mandatory)]
        [string]$Domain
    )

    $headers = Get-CloudflareAuthHeader -Credentials $Credentials

    try {
        $response = Invoke-RestMethod -Uri "$script:CloudflareApiBase/accounts/$AccountId/registrar/domains/$Domain" `
            -Headers $headers `
            -Method Get

        if ($response.success) {
            return $response.result
        }
        else {
            $errors = ($response.errors | ForEach-Object { $_.message }) -join "; "
            Write-Error "Cloudflare API error: $errors"
            return $null
        }
    }
    catch {
        Write-Error "Failed to query registrar status for $Domain : $($_.Exception.Message)"
        throw
    }
}

Export-ModuleMember -Function Get-CloudflareAccounts, Get-CloudflareZones, Get-CloudflareDNSRecords,
                              New-CloudflareAccount, New-CloudflareZone, New-CloudflareDNSRecord,
                              Import-CloudflareDNS, Get-CloudflareRegistrarDomain
