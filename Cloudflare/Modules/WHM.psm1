# WHM / cPanel API Module (for Bluehost reseller and cPanel accounts)
# Supports both WHM API v1 (port 2087, reseller) and cPanel UAPI (port 2083, single account)
# Read-only operations

function Get-AuthHeader {
    <#
    .SYNOPSIS
        Returns the authorization header based on connection type (WHM or cPanel).
    #>
    param(
        [Parameter(Mandatory)]
        $Credentials
    )

    if ($Credentials.Type -eq 'cPanel') {
        @{ Authorization = "cpanel $($Credentials.Username):$($Credentials.AccessToken)" }
    }
    else {
        @{ Authorization = "whm $($Credentials.Username):$($Credentials.AccessToken)" }
    }
}

function Get-BaseUrl {
    param(
        [Parameter(Mandatory)]
        $Credentials
    )

    if ($Credentials.Type -eq 'cPanel') {
        "https://$($Credentials.Server):2083"
    }
    else {
        "https://$($Credentials.Server):2087"
    }
}

function Get-WHMDomains {
    <#
    .SYNOPSIS
        Lists all DNS zones. Uses WHM listzones for reseller accounts,
        or cPanel DomainInfo/list_domains for single cPanel accounts.
    .PARAMETER Credentials
        Hashtable/PSCustomObject containing Server, Username, AccessToken, and Type (WHM or cPanel).
    .EXAMPLE
        $creds = (Import-Clixml .\Config\credentials.xml).WHM
        Get-WHMDomains -Credentials $creds
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Credentials
    )

    $headers = Get-AuthHeader -Credentials $Credentials
    $baseUrl = Get-BaseUrl -Credentials $Credentials

    if ($Credentials.Type -eq 'cPanel') {
        # cPanel UAPI - list_domains returns main + addon + parked + sub domains
        try {
            $response = Invoke-RestMethod -Uri "$baseUrl/execute/DomainInfo/list_domains" `
                -Headers $headers `
                -Method Get

            if ($response.status -eq 1) {
                $allDomains = @()

                # Main domain
                if ($response.data.main_domain) {
                    $allDomains += [PSCustomObject]@{ domain = $response.data.main_domain }
                }

                # Addon domains
                if ($response.data.addon_domains) {
                    foreach ($d in $response.data.addon_domains) {
                        $allDomains += [PSCustomObject]@{ domain = $d }
                    }
                }

                # Parked domains (aliases)
                if ($response.data.parked_domains) {
                    foreach ($d in $response.data.parked_domains) {
                        $allDomains += [PSCustomObject]@{ domain = $d }
                    }
                }

                return $allDomains
            }
            else {
                $errors = if ($response.errors) { $response.errors -join '; ' } else { "Unknown error" }
                throw "cPanel API Error: $errors"
            }
        }
        catch {
            Write-Error "Failed to retrieve cPanel domains: $($_.Exception.Message)"
            throw
        }
    }
    else {
        # WHM API - listzones
        try {
            $response = Invoke-RestMethod -Uri "$baseUrl/json-api/listzones?api.version=1" `
                -Headers $headers `
                -Method Get

            if ($response.metadata.result -eq 1) {
                return $response.data.zone
            }
            else {
                $reason = if ($response.metadata.reason) { $response.metadata.reason } else { "Unknown error" }
                throw "WHM API Error: $reason"
            }
        }
        catch {
            Write-Error "Failed to retrieve WHM domains: $($_.Exception.Message)"
            throw
        }
    }
}

function Get-WHMDNSRecords {
    <#
    .SYNOPSIS
        Retrieves DNS records for a specific domain.
        Uses WHM dumpzone for reseller accounts, or cPanel DNS/parse_zone for single accounts.
    .PARAMETER Credentials
        Hashtable/PSCustomObject containing Server, Username, AccessToken, and Type.
    .PARAMETER Domain
        The domain name to query (e.g., "example.com").
    .EXAMPLE
        $creds = (Import-Clixml .\Config\credentials.xml).WHM
        Get-WHMDNSRecords -Credentials $creds -Domain "example.com"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Credentials,

        [Parameter(Mandatory)]
        [string]$Domain
    )

    $headers = Get-AuthHeader -Credentials $Credentials
    $baseUrl = Get-BaseUrl -Credentials $Credentials

    # cPanel-specific subdomains to skip (A records only)
    $cpanelPrefixes = @(
        'cpanel.', 'webmail.', 'whm.', 'webdisk.',
        'cpcalendars.', 'cpcontacts.',
        'autodiscover.', 'autoconfig.',
        'mail.', 'www.'
    )

    # cPanel-specific SRV/TXT service records and other hosting artifacts to skip
    $cpanelServiceNames = @(
        '_autodiscover._tcp', '_carddav._tcp', '_carddavs._tcp',
        '_caldav._tcp', '_caldavs._tcp',
        '_cpanel-dcv-test-record', 'ftp'
    )

    # Skip ACME challenge records (Let's Encrypt validation - ephemeral)
    $acmePattern = '^_acme-challenge'

    if ($Credentials.Type -eq 'cPanel') {
        # cPanel UAPI - parse_zone returns structured records
        try {
            $response = Invoke-RestMethod -Uri "$baseUrl/execute/DNS/parse_zone?zone=$Domain" `
                -Headers $headers `
                -Method Get

            if ($response.status -eq 1 -and $response.data) {
                $records = foreach ($raw in $response.data) {
                    $converted = ConvertFrom-CPanelZoneRecord -Record $raw -Domain $Domain
                    if (-not $converted) { continue }

                    # Skip localhost/loopback A records
                    if ($converted.type -eq 'A' -and $converted.data -eq '127.0.0.1') {
                        Write-Verbose "Skipping localhost record: $($converted.name) -> 127.0.0.1"
                        continue
                    }

                    # Skip cPanel-specific subdomain A records
                    $isCpanelRecord = $false
                    foreach ($prefix in $cpanelPrefixes) {
                        if ($converted.name -eq ($prefix.TrimEnd('.')) -and $converted.type -eq 'A') {
                            $isCpanelRecord = $true
                            break
                        }
                    }
                    if ($isCpanelRecord) {
                        Write-Verbose "Skipping cPanel service record: $($converted.type) $($converted.name)"
                        continue
                    }

                    # Skip cPanel-specific SRV/TXT service records
                    if ($converted.name -in $cpanelServiceNames) {
                        Write-Verbose "Skipping cPanel service record: $($converted.type) $($converted.name)"
                        continue
                    }

                    # Skip ACME challenge records (Let's Encrypt validation - ephemeral)
                    if ($converted.name -match $acmePattern) {
                        Write-Verbose "Skipping ACME challenge record: $($converted.name)"
                        continue
                    }

                    $converted
                }

                return $records
            }
            else {
                $errors = if ($response.errors) { $response.errors -join '; ' } else { "Unknown error" }
                throw "cPanel API Error: $errors"
            }
        }
        catch {
            Write-Error "Failed to retrieve DNS records for $Domain : $($_.Exception.Message)"
            throw
        }
    }
    else {
        # WHM API - dumpzone
        try {
            $response = Invoke-RestMethod -Uri "$baseUrl/json-api/dumpzone?api.version=1&domain=$Domain" `
                -Headers $headers `
                -Method Get

            if ($response.metadata.result -eq 1) {
                $rawRecords = $response.data.zone | Select-Object -First 1 | Select-Object -ExpandProperty record

                $records = foreach ($raw in $rawRecords) {
                    $converted = ConvertFrom-WHMZoneRecord -Record $raw -Domain $Domain
                    if (-not $converted) { continue }

                    # Skip localhost/loopback A records
                    if ($converted.type -eq 'A' -and $converted.data -eq '127.0.0.1') {
                        Write-Verbose "Skipping localhost record: $($converted.name) -> 127.0.0.1"
                        continue
                    }

                    # Skip cPanel-specific subdomain A records
                    $isCpanelRecord = $false
                    foreach ($prefix in $cpanelPrefixes) {
                        if ($converted.name -eq ($prefix.TrimEnd('.')) -and $converted.type -eq 'A') {
                            $isCpanelRecord = $true
                            break
                        }
                    }
                    if ($isCpanelRecord) {
                        Write-Verbose "Skipping cPanel service record: $($converted.type) $($converted.name)"
                        continue
                    }

                    # Skip cPanel-specific SRV/TXT service records
                    if ($converted.name -in $cpanelServiceNames) {
                        Write-Verbose "Skipping cPanel service record: $($converted.type) $($converted.name)"
                        continue
                    }

                    # Skip ACME challenge records (Let's Encrypt validation - ephemeral)
                    if ($converted.name -match $acmePattern) {
                        Write-Verbose "Skipping ACME challenge record: $($converted.name)"
                        continue
                    }

                    $converted
                }

                return $records
            }
            else {
                $reason = if ($response.metadata.reason) { $response.metadata.reason } else { "Unknown error" }
                throw "WHM API Error: $reason"
            }
        }
        catch {
            Write-Error "Failed to retrieve DNS records for $Domain : $($_.Exception.Message)"
            throw
        }
    }
}

function ConvertFrom-CPanelZoneRecord {
    <#
    .SYNOPSIS
        Converts a cPanel UAPI parse_zone record (base64-encoded fields)
        into the normalized format expected by Import-CloudflareDNS.
    #>
    param(
        [Parameter(Mandatory)]
        $Record,

        [Parameter(Mandatory)]
        [string]$Domain
    )

    # Skip non-record entries (comments, control lines)
    if ($Record.type -ne 'record') { return $null }

    $type = $Record.record_type
    if (-not $type) { return $null }

    # Skip SOA and NS records
    if ($type -in @('SOA', 'NS')) { return $null }

    $ttl = if ($Record.ttl) { [int]$Record.ttl } else { 14400 }

    # Decode base64 helper
    function DecodeB64 { param([string]$s) [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($s)) }

    # Decode name
    $rawName = if ($Record.dname_b64) { (DecodeB64 $Record.dname_b64) -replace '\.$', '' } else { return $null }

    # Decode data array
    $dataValues = @()
    if ($Record.data_b64) {
        foreach ($d in $Record.data_b64) {
            $dataValues += (DecodeB64 $d) -replace '\.$', ''
        }
    }

    # Convert FQDN name to relative name
    $domainLower = $Domain.ToLower()
    $nameLower = $rawName.ToLower()

    if ($nameLower -eq $domainLower) {
        $name = '@'
    }
    elseif ($nameLower.EndsWith(".$domainLower")) {
        $name = $rawName.Substring(0, $rawName.Length - $Domain.Length - 1)
    }
    else {
        $name = $rawName
    }

    switch ($type) {
        'A' {
            @{
                type = 'A'
                name = $name
                data = $dataValues[0]
                ttl  = $ttl
            }
        }
        'AAAA' {
            @{
                type = 'AAAA'
                name = $name
                data = $dataValues[0]
                ttl  = $ttl
            }
        }
        'CNAME' {
            @{
                type = 'CNAME'
                name = $name
                data = $dataValues[0]
                ttl  = $ttl
            }
        }
        'MX' {
            # data_b64: [priority, exchange]
            $priority = if ($dataValues.Count -ge 2) { [int]$dataValues[0] } else { 10 }
            $exchange = if ($dataValues.Count -ge 2) { $dataValues[1] } else { $dataValues[0] }
            @{
                type     = 'MX'
                name     = $name
                data     = $exchange
                ttl      = $ttl
                priority = $priority
            }
        }
        'TXT' {
            $txtData = $dataValues -join ''
            if ($txtData -match '^\s*"(.*?)"\s*$') {
                $txtData = $Matches[1]
            }
            @{
                type = 'TXT'
                name = $name
                data = $txtData
                ttl  = $ttl
            }
        }
        'SRV' {
            # data_b64: [priority, weight, port, target]
            $srvPriority = if ($dataValues.Count -ge 4) { [int]$dataValues[0] } else { 0 }
            $srvWeight = if ($dataValues.Count -ge 4) { $dataValues[1] } else { '0' }
            $srvPort = if ($dataValues.Count -ge 4) { $dataValues[2] } else { '0' }
            $srvTarget = if ($dataValues.Count -ge 4) { $dataValues[3] } else { $dataValues[0] }
            $srvData = "$srvWeight $srvPort $srvTarget"
            @{
                type     = 'SRV'
                name     = $name
                data     = $srvData
                ttl      = $ttl
                priority = $srvPriority
            }
        }
        'CAA' {
            # data_b64: [flag, tag, value]
            $caaData = if ($dataValues.Count -ge 3) { "$($dataValues[0]) $($dataValues[1]) `"$($dataValues[2])`"" } else { $dataValues -join ' ' }
            @{
                type = 'CAA'
                name = $name
                data = $caaData
                ttl  = $ttl
            }
        }
        default {
            Write-Verbose "Skipping unsupported record type: $type"
            $null
        }
    }
}

function ConvertFrom-WHMZoneRecord {
    <#
    .SYNOPSIS
        Converts a raw WHM dumpzone record into the normalized format
        expected by Import-CloudflareDNS (type, name, data, ttl, priority).
    #>
    param(
        [Parameter(Mandatory)]
        $Record,

        [Parameter(Mandatory)]
        [string]$Domain
    )

    $type = $Record.type
    if (-not $type) { return $null }

    # Skip SOA and NS records (Cloudflare manages these)
    if ($type -in @('SOA', 'NS')) { return $null }

    $ttl = if ($Record.ttl) { [int]$Record.ttl } else { 14400 }

    # Convert FQDN name to relative name
    $rawName = ($Record.name -replace '\.$', '')  # strip trailing dot
    $domainLower = $Domain.ToLower()
    $nameLower = $rawName.ToLower()

    if ($nameLower -eq $domainLower) {
        $name = '@'
    }
    elseif ($nameLower.EndsWith(".$domainLower")) {
        $name = $rawName.Substring(0, $rawName.Length - $Domain.Length - 1)
    }
    else {
        $name = $rawName
    }

    switch ($type) {
        'A' {
            @{
                type = 'A'
                name = $name
                data = $Record.address
                ttl  = $ttl
            }
        }
        'AAAA' {
            @{
                type = 'AAAA'
                name = $name
                data = $Record.address
                ttl  = $ttl
            }
        }
        'CNAME' {
            $target = ($Record.cname -replace '\.$', '')
            @{
                type = 'CNAME'
                name = $name
                data = $target
                ttl  = $ttl
            }
        }
        'MX' {
            $exchange = ($Record.exchange -replace '\.$', '')
            @{
                type     = 'MX'
                name     = $name
                data     = $exchange
                ttl      = $ttl
                priority = if ($Record.preference) { [int]$Record.preference } else { 10 }
            }
        }
        'TXT' {
            $txtData = $Record.txtdata
            # WHM may return TXT data with or without quotes
            if ($txtData -match '^\s*"(.*?)"\s*$') {
                $txtData = $Matches[1]
            }
            @{
                type = 'TXT'
                name = $name
                data = $txtData
                ttl  = $ttl
            }
        }
        'SRV' {
            # WHM SRV: target, port, weight, priority in separate fields
            $target = ($Record.target -replace '\.$', '')
            $srvData = "$($Record.weight) $($Record.port) $target"
            @{
                type     = 'SRV'
                name     = $name
                data     = $srvData
                ttl      = $ttl
                priority = if ($Record.priority) { [int]$Record.priority } else { 0 }
            }
        }
        'CAA' {
            $caaData = "$($Record.flag) $($Record.tag) `"$($Record.value)`""
            @{
                type = 'CAA'
                name = $name
                data = $caaData
                ttl  = $ttl
            }
        }
        default {
            Write-Verbose "Skipping unsupported record type: $type"
            $null
        }
    }
}

Export-ModuleMember -Function Get-WHMDomains, Get-WHMDNSRecords
