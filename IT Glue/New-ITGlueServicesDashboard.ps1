<#
.SYNOPSIS
    Builds a per-client "All Services" dashboard (gavsto Bootstrap-panel style) for an IT Glue
    organization and renders it to a local HTML preview file.

.DESCRIPTION
    Pulls live, READ-ONLY data from the managed-service vendors this org uses and renders a grid
    of colored Bootstrap-3 panels - the same HTML that can later be PATCHed into an IT Glue
    Flexible Asset "Textbox" trait to "go live" (see the plan / README).

    Tiles (first build):
      - Antivirus     : auto-detect SentinelOne OR Bitdefender; protected device/license count.
      - Duo           : # users protected (or "No Duo").
      - Backup        : Veeam (VSPC REST API) if present, else Cove; TB + device/VM counts.
      - M365 Backup   : Cove M365 backup (only if the org has M365), else "None".
      - Domains       : list of IT Glue domains, each linking to its IT Glue domain page.
      - M365 Licenses : read from IT Glue's native "Microsoft Licenses" flexible asset.

    CLIENT MATCHING IS AUTOMATIC. Each vendor is resolved against the org by normalized name.
    A confident single match is used silently; an ambiguous one shows a pick-list; if a vendor
    returns no candidate, you are asked once. Every answer (including "not managed") is cached in
    Config\service-map.json so you are asked at most once per client per vendor. -Remap re-resolves.

    READ-ONLY: this script only performs GET/list/lookup calls against IT Glue and the vendor APIs
    (aside from each vendor's own auth-token POST). It never writes to IT Glue.

.PARAMETER Organization
    IT Glue organization name (exact, then partial) or numeric ID. Prompts if omitted.

.PARAMETER OutFile
    Output HTML path. Defaults to .\Output\<org>-services.html

.PARAMETER Reset
    Re-prompt for all stored credentials.

.PARAMETER Remap
    Clear the cached service-map resolution (for the selected org, or all with no -Organization)
    and re-resolve from scratch.

.PARAMETER NoOpen
    Do not auto-open the generated HTML in the default browser.

.EXAMPLE
    .\New-ITGlueServicesDashboard.ps1 -Organization "Contoso Ltd"

.EXAMPLE
    .\New-ITGlueServicesDashboard.ps1 -Organization "Contoso Ltd" -Remap
#>
[CmdletBinding()]
param(
    [string]$Organization,
    [string]$OutFile,
    [switch]$Reset,
    [switch]$Remap,
    [switch]$NoOpen
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# -- Paths ---------------------------------------------------------------------
$ConfigDir        = Join-Path $ScriptDir 'Config'
$CredentialsPath  = Join-Path $ConfigDir 'credentials.xml'
$ServiceMapPath   = Join-Path $ConfigDir 'service-map.json'
$OutputDir        = Join-Path $ScriptDir 'Output'

# ==============================================================================
# Small helpers
# ==============================================================================
function Write-Status {
    param([string]$Message, [ValidateSet('Info','Success','Warning','Error','Detail')][string]$Level = 'Info')
    $color = @{ Info='Cyan'; Success='Green'; Warning='Yellow'; Error='Red'; Detail='Gray' }[$Level]
    Write-Host $Message -ForegroundColor $color
}

function ConvertTo-NormalizedName {
    # Lowercase, strip punctuation + common legal suffixes, collapse whitespace.
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    $n = $Name.ToLowerInvariant()
    $n = $n -replace '[^a-z0-9 ]', ' '            # drop punctuation
    $n = $n -replace '\b(llc|inc|ltd|limited|corp|corporation|co|llp|pllc|plc|company|the)\b', ' '
    $n = ($n -replace '\s+', ' ').Trim()
    return $n
}

function Get-HtmlEncoded {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

# ==============================================================================
# Credentials
# ==============================================================================
function Initialize-Credentials {
    param([switch]$Reset)

    if (-not (Test-Path $ConfigDir)) { New-Item -ItemType Directory -Path $ConfigDir -Force | Out-Null }

    $creds = @{}
    if ((Test-Path $CredentialsPath) -and (-not $Reset)) {
        try { $creds = Import-Clixml -Path $CredentialsPath }
        catch { Write-Status "Could not read stored credentials; will prompt." Warning; $creds = @{} }
    }
    if ($creds -isnot [hashtable]) { $creds = @{} }

    $changed = $false

    # -- IT Glue (required) --
    if (-not $creds.ITGlue) {
        Write-Host "`n=== IT Glue API ===" -ForegroundColor Cyan
        $base = Read-Host "IT Glue API Base URL [https://api.itglue.com]"
        if ([string]::IsNullOrWhiteSpace($base)) { $base = 'https://api.itglue.com' }
        $creds.ITGlue = @{
            BaseUri   = $base
            ApiKey    = Read-Host "IT Glue API Key"
            Subdomain = Read-Host "IT Glue Subdomain (e.g. 'yourcompany' from yourcompany.itglue.com)"
        }
        $changed = $true
    }

    # -- Optional vendor sections: prompt once, allow skipping --
    if (-not $creds.SentinelOne) {
        if ((Read-Host "`nConfigure SentinelOne? (y/N)") -match '^[Yy]') {
            $instances = @()
            foreach ($inst in @('EDR','MDR')) {
                if ((Read-Host "  Add SentinelOne $inst instance? (y/N)") -match '^[Yy]') {
                    $instances += @{
                        Name   = $inst
                        URL    = (Read-Host "    $inst URL (e.g. https://usea1-xxxx$($inst.ToLower()).sentinelone.net)").TrimEnd('/')
                        APIKey = Read-Host "    $inst API Token"
                    }
                }
            }
            $creds.SentinelOne = @{ Configured = ($instances.Count -gt 0); Instances = $instances }
        } else { $creds.SentinelOne = @{ Configured = $false } }
        $changed = $true
    }

    if (-not $creds.Bitdefender) {
        if ((Read-Host "Configure Bitdefender GravityZone? (y/N)") -match '^[Yy]') {
            $creds.Bitdefender = @{
                Configured = $true
                URL        = (Read-Host "  GravityZone API URL (e.g. https://cloud.gravityzone.bitdefender.com)").TrimEnd('/')
                ApiKey     = Read-Host "  GravityZone API Key"
            }
        } else { $creds.Bitdefender = @{ Configured = $false } }
        $changed = $true
    }

    if (-not $creds.Duo) {
        if ((Read-Host "Configure Duo (Accounts/MSP API)? (y/N)") -match '^[Yy]') {
            $creds.Duo = @{
                Configured     = $true
                Type           = 'Accounts'
                ApiHost        = Read-Host "  Duo API Host (e.g. api-xxxxxxxx.duosecurity.com)"
                IntegrationKey = Read-Host "  Duo Integration Key"
                SecretKey      = Read-Host "  Duo Secret Key"
            }
        } else { $creds.Duo = @{ Configured = $false } }
        $changed = $true
    }

    if (-not $creds.Cove) {
        if ((Read-Host "Configure Cove Data Protection (N-able)? (y/N)") -match '^[Yy]') {
            $cred = Get-Credential -Message "Cove Data Protection (backup.management) login"
            $partner = Read-Host "  Your Cove (MSP) Partner Name exactly as shown in backup.management (e.g. 'Contoso MSP LLC')"
            $creds.Cove = @{ Configured = $true; Credential = $cred; PartnerName = $partner }
        } else { $creds.Cove = @{ Configured = $false } }
        $changed = $true
    }

    if (-not $creds.VeeamVspc) {
        if ((Read-Host "Configure Veeam Service Provider Console? (y/N)") -match '^[Yy]') {
            $url    = (Read-Host "  VSPC URL (e.g. https://vspc.yourdomain.com:1280)").TrimEnd('/')
            $apiKey = Read-Host "  VSPC API Key (Simple Key 'private key' string, read-only recommended)"
            $creds.VeeamVspc = @{ Configured = $true; URL = $url; ApiKey = $apiKey }
        } else { $creds.VeeamVspc = @{ Configured = $false } }
        $changed = $true
    }

    if ($changed) {
        $creds | Export-Clixml -Path $CredentialsPath
        Write-Status "Credentials saved to $CredentialsPath" Success
    }
    return $creds
}

# ==============================================================================
# Service-map cache + generic resolver
# ==============================================================================
function Get-ServiceMap {
    if (-not (Test-Path $ServiceMapPath)) { return @{} }
    try {
        $raw = Get-Content $ServiceMapPath -Raw | ConvertFrom-Json
        $map = @{}
        foreach ($p in $raw.PSObject.Properties) { $map[$p.Name] = $p.Value }
        return $map
    } catch {
        Write-Status "Could not parse service-map.json; starting fresh." Warning
        return @{}
    }
}

function Save-ServiceMap {
    param([hashtable]$Map)
    $Map | ConvertTo-Json -Depth 8 | Set-Content -Path $ServiceMapPath -Encoding UTF8
}

function Get-OrgMapEntry {
    # Returns (creating if needed) the cache entry object for this org.
    param([hashtable]$Map, [string]$OrgId, [string]$OrgName)
    if (-not $Map.ContainsKey($OrgId)) {
        $Map[$OrgId] = [pscustomobject]@{
            orgName    = $OrgName
            resolvedOn = (Get-Date).ToString('yyyy-MM-dd')
            _unmanaged = @()
        }
    }
    return $Map[$OrgId]
}

function Resolve-OrgServiceId {
    <#
        Generic per-vendor resolver. Returns the chosen candidate object, or $null if the org is
        "not managed" by this vendor. Caches the decision under $entry.$VendorKey.

        -Candidates: array of objects each having .Id and .Name (extra props are preserved).
    #>
    param(
        [Parameter(Mandatory)][pscustomobject]$Entry,
        [Parameter(Mandatory)][string]$VendorKey,      # e.g. 'sentinelOne','bitdefender','duo','cove','veeam'
        [Parameter(Mandatory)][string]$VendorLabel,    # display name
        [Parameter(Mandatory)][string]$OrgName,
        [object[]]$Candidates = @()
    )

    # 1) cache hit (resolved id, or confirmed unmanaged)
    if ($Entry._unmanaged -contains $VendorKey) { return $null }
    $cached = $Entry.PSObject.Properties[$VendorKey]
    if ($cached -and $cached.Value) {
        $cv = $cached.Value
        # re-hydrate against current candidate list when possible (keeps Name/extra fresh)
        $match = $Candidates | Where-Object { "$($_.Id)" -eq "$($cv.Id)" } | Select-Object -First 1
        if ($match) { return $match }
        return $cv   # candidate list unavailable this run; trust the cached id
    }

    # 2) auto-match by normalized name
    $target = ConvertTo-NormalizedName $OrgName
    $matches = @($Candidates | Where-Object { (ConvertTo-NormalizedName $_.Name) -eq $target })
    if ($matches.Count -eq 1) {
        Write-Status "  [$VendorLabel] auto-matched '$($matches[0].Name)'" Detail
        Add-Resolution -Entry $Entry -VendorKey $VendorKey -Candidate $matches[0]
        return $matches[0]
    }

    # 3) interactive: pick-list of candidates (if any), plus none/manual
    Write-Host "`n  [$VendorLabel] No confident match for '$OrgName'." -ForegroundColor Yellow
    if ($Candidates.Count -gt 0) {
        $sorted = $Candidates | Sort-Object Name
        for ($i = 0; $i -lt $sorted.Count; $i++) {
            Write-Host ("    [{0}] {1}" -f ($i + 1), $sorted[$i].Name)
        }
        Write-Host "    [0] None / not managed by $VendorLabel"
        do {
            $sel = Read-Host "  Select $VendorLabel customer for '$OrgName' (0-$($sorted.Count))"
            $n = -1; [int]::TryParse($sel, [ref]$n) | Out-Null
        } while ($n -lt 0 -or $n -gt $sorted.Count)

        if ($n -eq 0) { Add-Unmanaged -Entry $Entry -VendorKey $VendorKey; return $null }
        $chosen = $sorted[$n - 1]
        Add-Resolution -Entry $Entry -VendorKey $VendorKey -Candidate $chosen
        return $chosen
    }
    else {
        # 4) no candidates at all - manual id, or confirm not-managed
        $manual = Read-Host "  Enter $VendorLabel customer/ID for '$OrgName' (blank = not managed)"
        if ([string]::IsNullOrWhiteSpace($manual)) { Add-Unmanaged -Entry $Entry -VendorKey $VendorKey; return $null }
        $obj = [pscustomobject]@{ Id = $manual; Name = $manual }
        Add-Resolution -Entry $Entry -VendorKey $VendorKey -Candidate $obj
        return $obj
    }
}

function Add-Resolution {
    param([pscustomobject]$Entry, [string]$VendorKey, [object]$Candidate)
    $val = [pscustomobject]@{ Id = "$($Candidate.Id)"; Name = $Candidate.Name }
    foreach ($extra in @('Instance','Provider')) {
        if ($Candidate.PSObject.Properties[$extra]) {
            $val | Add-Member -NotePropertyName $extra -NotePropertyValue $Candidate.$extra -Force
        }
    }
    $Entry | Add-Member -NotePropertyName $VendorKey -NotePropertyValue $val -Force
    $script:ServiceMapDirty = $true
}

function Add-Unmanaged {
    param([pscustomobject]$Entry, [string]$VendorKey)
    if ($Entry._unmanaged -notcontains $VendorKey) {
        $Entry._unmanaged = @($Entry._unmanaged) + $VendorKey
    }
    $script:ServiceMapDirty = $true
}

# ==============================================================================
# Vendor clients (READ-ONLY)
# ==============================================================================

# ---- SentinelOne ----
function Get-S1Sites {
    param([string]$BaseURL, [string]$APIKey, [string]$InstanceName)
    $all = @(); $cursor = $null
    $headers = @{ 'Authorization' = "ApiToken $APIKey" }
    do {
        $uri = "$BaseURL/web/api/v2.1/sites?states=active&limit=1000"
        if ($cursor) { $uri += "&cursor=$cursor" }
        $resp = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
        foreach ($s in $resp.data.sites) {
            $all += [pscustomobject]@{
                Id             = $s.id
                Name           = $s.name
                Instance       = $InstanceName
                Provider       = 'SentinelOne'
                ActiveLicenses = [int]$s.activeLicenses
                TotalLicenses  = [int]$s.totalLicenses
            }
        }
        $cursor = $resp.pagination.nextCursor
    } while ($cursor)
    return $all
}

# ---- Bitdefender GravityZone (JSON-RPC) ----
# NOTE: exact JSON-RPC method names/endpoints per GravityZone version should be confirmed against
#       https://www.bitdefender.com/business/support/en/77209-api.html - calls are wrapped so any
#       mismatch degrades gracefully to "not available" rather than throwing.
function Invoke-BdJsonRpc {
    param([string]$BaseURL, [string]$ApiKey, [string]$Service, [string]$Method, [hashtable]$Params = @{})
    $auth = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${ApiKey}:"))
    $body = @{ id = 1; jsonrpc = '2.0'; method = $Method; params = $Params } | ConvertTo-Json -Depth 8
    $resp = Invoke-RestMethod -Uri "$BaseURL/api/v1.0/jsonrpc/$Service" -Method Post `
        -Headers @{ Authorization = $auth } -ContentType 'application/json' -Body $body
    if ($resp.error) { throw "Bitdefender API error: $($resp.error.message)" }
    return $resp.result
}

function Get-BdCompanies {
    param([string]$BaseURL, [string]$ApiKey)
    $companies = @(); $page = 1
    do {
        $res = Invoke-BdJsonRpc -BaseURL $BaseURL -ApiKey $ApiKey -Service 'companies' `
            -Method 'getCompaniesList' -Params @{ page = $page; perPage = 100 }
        foreach ($c in $res.items) {
            $companies += [pscustomobject]@{ Id = $c.id; Name = $c.name; Provider = 'Bitdefender' }
        }
        $page++
    } while ($res.items -and $companies.Count -lt [int]$res.total)
    return $companies
}

function Get-BdEndpointCount {
    param([string]$BaseURL, [string]$ApiKey, [string]$CompanyId)
    # Licensing usage gives used seats per company on most MSP setups.
    try {
        $res = Invoke-BdJsonRpc -BaseURL $BaseURL -ApiKey $ApiKey -Service 'licensing' `
            -Method 'getMonthlyUsagePerProductType' -Params @{ targetId = $CompanyId }
        $used = ($res.usage | Measure-Object -Property usedSlots -Sum).Sum
        if ($used) { return [int]$used }
    } catch { }
    return $null
}

# ---- Cove Data Protection (backup.management JSON-RPC, visa auth) ----
function Connect-Cove {
    param([pscredential]$Credential)
    $url = 'https://api.backup.management/jsonapi'
    $data = @{ jsonrpc='2.0'; id='login'; method='Login'; params=@{
        username = $Credential.UserName
        password = $Credential.GetNetworkCredential().Password
    } }
    $resp = Invoke-RestMethod -Uri $url -Method Post -ContentType 'application/json' -Body (ConvertTo-Json $data)
    if (-not $resp.visa) { throw "Cove login failed: $($resp.error.message)" }
    return $resp.visa
}

function Invoke-CoveJsonRpc {
    param([string]$Visa, [string]$Method, [hashtable]$Params)
    $url = 'https://api.backup.management/jsonapi'
    $data = @{ jsonrpc = '2.0'; id = '1'; visa = $Visa; method = $Method; params = $Params }
    $resp = Invoke-RestMethod -Uri $url -Method Post -ContentType 'application/json; charset=utf-8' `
        -Headers @{ Authorization = "Bearer $Visa" } -Body ([Text.Encoding]::UTF8.GetBytes((ConvertTo-Json $data -Depth 6)))
    if ($resp.error) { throw "Cove API error ($Method): $($resp.error.message)" }
    return $resp
}

function Get-CovePartners {
    # Resolve the MSP's own partner id by name, then enumerate child partners (customers).
    param([string]$Visa, [string]$MyPartnerName)
    $info = Invoke-CoveJsonRpc -Visa $Visa -Method 'GetPartnerInfo' -Params @{ name = $MyPartnerName }
    $parentId = [int]$info.result.result.Id
    if (-not $parentId) { throw "Could not resolve Cove partner id for '$MyPartnerName'." }

    $enum = Invoke-CoveJsonRpc -Visa $Visa -Method 'EnumeratePartners' -Params @{
        parentPartnerId = $parentId; fetchRecursively = $true; fields = @(0, 1, 3, 4, 8, 10, 14)
    }
    $partners = @()
    foreach ($p in $enum.result.result) {
        if ([string]::IsNullOrWhiteSpace($p.Name)) { continue }
        $partners += [pscustomobject]@{ Id = [string]$p.Id; Name = $p.Name; Provider = 'Cove' }
    }
    return $partners
}

function Get-CoveDevices {
    param([string]$Visa, [string]$PartnerId, [int]$MaxDevices = 15000)
    $url = 'https://api.backup.management/jsonapi'
    $data = @{ jsonrpc='2.0'; id='2'; visa=$Visa; method='EnumerateAccountStatistics'; params=@{ query=@{
        PartnerId         = [int]$PartnerId
        Columns           = @('AU','AR','AN','MN','OS','OT','PD','AP','PN','T3','US','I81')
        OrderBy           = 'CD DESC'
        StartRecordNumber = 0
        RecordsCount      = $MaxDevices
    } } }
    $resp = Invoke-RestMethod -Uri $url -Method Post -ContentType 'application/json; charset=utf-8' `
        -Headers @{ Authorization = "Bearer $Visa" } -Body ([Text.Encoding]::UTF8.GetBytes((ConvertTo-Json $data -Depth 6)))
    $devices = @()
    foreach ($d in $resp.result.result) {
        $devices += [pscustomobject]@{
            DeviceName  = ($d.Settings.AN -join '')
            OSType      = ($d.Settings.OT -join '')          # e.g. Workstation / Server
            Product     = ($d.Settings.PN -join '')
            DataSources = ($d.Settings.AP -join '')          # e.g. "FileSystem, Exchange, ..."
            Physicality = ($d.Settings.I81 -join '')
            SelectedGB  = [Math]::Round([Decimal](($d.Settings.T3 -join '') / 1GB), 2)
            UsedGB      = [Math]::Round([Decimal](($d.Settings.US -join '') / 1GB), 2)
        }
    }
    return $devices
}

# ---- Veeam Service Provider Console (REST) ----
# Auth: a VSPC API key (Simple Key 'private key') is a PERMANENT bearer credential used directly
#       in the Authorization header -- no /api/v3/token exchange. The read-only toggle on the key
#       applies to VSPC and all proxied Veeam-server requests.
#       Ref: https://helpcenter.veeam.com/docs/vac/rest/api_keys.html
# NOTE: endpoint shapes follow the documented v3 API; calls are wrapped so a mismatch degrades.
function Test-VspcAuth {
    param([string]$BaseURL, [string]$ApiKey)
    # Lightweight read-only call to validate the key up front and surface a clear error.
    $null = Invoke-RestMethod -Uri "$BaseURL/api/v3/about" -Method Get `
        -Headers @{ Authorization = "Bearer $ApiKey" }
    return $ApiKey
}

function Get-VspcCompanies {
    param([string]$BaseURL, [string]$Token)
    $resp = Invoke-RestMethod -Uri "$BaseURL/api/v3/organizations/companies?limit=1000" -Method Get `
        -Headers @{ Authorization = "Bearer $Token" }
    $companies = @()
    foreach ($c in $resp.data) {
        $companies += [pscustomobject]@{ Id = $c.instanceUid; Name = $c.name; Provider = 'Veeam' }
    }
    return $companies
}

function Get-VspcCompanyBackup {
    param([string]$BaseURL, [string]$Token, [string]$CompanyUid)
    # Best-effort: protected workloads + storage usage for the company.
    $h = @{ Authorization = "Bearer $Token" }
    $result = [pscustomobject]@{ Servers = 0; Workstations = 0; VMs = 0; UsedTB = 0 }
    try {
        $pw = Invoke-RestMethod -Uri "$BaseURL/api/v3/protectedWorkloads/managedVirtualMachines?filter=[{\""property\"":\""organizationUid\"",\""operation\"":\""equals\"",\""value\"":\""$CompanyUid\""}]&limit=1000" -Method Get -Headers $h
        $result.VMs = @($pw.data).Count
    } catch { }
    try {
        $pc = Invoke-RestMethod -Uri "$BaseURL/api/v3/protectedWorkloads/computersManagedByConsole?filter=[{\""property\"":\""organizationUid\"",\""operation\"":\""equals\"",\""value\"":\""$CompanyUid\""}]&limit=1000" -Method Get -Headers $h
        foreach ($c in $pc.data) {
            if ("$($c.platform)$($c.operatingSystem)" -match 'server') { $result.Servers++ } else { $result.Workstations++ }
        }
    } catch { }
    try {
        $bytes = (Invoke-RestMethod -Uri "$BaseURL/api/v3/organizations/companies/$CompanyUid/backupResources" -Method Get -Headers $h).data |
            Measure-Object -Property usedStorageQuota -Sum
        if ($bytes.Sum) { $result.UsedTB = [Math]::Round($bytes.Sum / 1TB, 2) }
    } catch { }
    return $result
}

# ---- IT Glue (READ-ONLY) ----
function Get-ItgAllOrganizations {
    $all = @(); $page = 1
    do {
        $r = Get-ITGlueOrganizations -page_size 1000 -page_number $page
        if ($r.data) { $all += $r.data }
        $page++
    } while ($r.data -and $r.data.Count -eq 1000)
    return $all
}

function Resolve-ItgOrganization {
    <#
        Robust org lookup. NOTE: IT Glue's filter[name] treats commas as a value list, so names
        like "Christine S. Rausch, MD, PC" break it. We therefore match client-side: numeric id,
        then case-insensitive exact, then normalized, then unique 'contains'.
    #>
    param([string]$Query)

    if ($Query -match '^\d+$') {
        try { $o = (Get-ITGlueOrganizations -id $Query).data; if ($o) { return $o } } catch { }
    }

    $all = Get-ItgAllOrganizations
    if (-not $all -or $all.Count -eq 0) { return $null }

    $q = $Query.Trim()
    # 1) case-insensitive exact (handles case/whitespace differences)
    $hit = $all | Where-Object { $_.attributes.name.Trim() -ieq $q } | Select-Object -First 1
    if ($hit) { return $hit }

    # 2) normalized (punctuation/legal-suffix insensitive)
    $target = ConvertTo-NormalizedName $q
    $norm = @($all | Where-Object { (ConvertTo-NormalizedName $_.attributes.name) -eq $target })
    if ($norm.Count -ge 1) { return $norm[0] }

    # 3) unique substring match
    $contains = @($all | Where-Object { $_.attributes.name -like "*$q*" })
    if ($contains.Count -eq 1) { return $contains[0] }
    if ($contains.Count -gt 1) {
        Write-Host "`nMultiple organizations match '$Query':" -ForegroundColor Yellow
        $sorted = $contains | Sort-Object { $_.attributes.name }
        for ($i = 0; $i -lt $sorted.Count; $i++) { Write-Host ("  [{0}] {1}" -f ($i + 1), $sorted[$i].attributes.name) }
        do { $sel = Read-Host "Select organization (1-$($sorted.Count))"; $n = -1; [int]::TryParse($sel, [ref]$n) | Out-Null }
        while ($n -lt 1 -or $n -gt $sorted.Count)
        return $sorted[$n - 1]
    }
    return $null
}

function Get-ItgFlexAssetTypeId {
    param([string[]]$NameLike)
    $page = 1
    do {
        $res = Get-ITGlueFlexibleAssetTypes -page_size 1000 -page_number $page
        foreach ($t in $res.data) {
            foreach ($pattern in $NameLike) {
                if ($t.attributes.name -like $pattern) { return $t.id }
            }
        }
        $page++
    } while ($res.data -and $res.data.Count -eq 1000)
    return $null
}

# ==============================================================================
# Tile providers - each returns a normalized tile object
#   @{ Title; Shading; Content; Detail; Link; IsInfo; InfoItems }
# ==============================================================================
function New-Tile {
    param(
        [string]$Title, [string]$Shading, [string]$Content = '', [string]$Detail = '',
        [string]$Link = '', [switch]$IsInfo, [object[]]$InfoItems = @()
    )
    [pscustomobject]@{
        Title = $Title; Shading = $Shading; Content = $Content; Detail = $Detail
        Link = $Link; IsInfo = [bool]$IsInfo; InfoItems = $InfoItems
    }
}

function Get-DashAntivirus {
    param([pscustomobject]$Entry, [string]$OrgName, [hashtable]$Creds)

    # Try SentinelOne first
    if ($Creds.SentinelOne.Configured) {
        $sites = @()
        foreach ($inst in $Creds.SentinelOne.Instances) {
            try { $sites += Get-S1Sites -BaseURL $inst.URL -APIKey $inst.APIKey -InstanceName $inst.Name }
            catch { Write-Status "  SentinelOne $($inst.Name) query failed: $($_.Exception.Message)" Warning }
        }
        $site = Resolve-OrgServiceId -Entry $Entry -VendorKey 'sentinelOne' -VendorLabel 'SentinelOne' -OrgName $OrgName -Candidates $sites
        if ($site) {
            $count = $site.ActiveLicenses
            return New-Tile -Title 'SentinelOne' -Shading 'success' -Content "$count Protected" -Detail "Active endpoints"
        }
    }

    # Then Bitdefender
    if ($Creds.Bitdefender.Configured) {
        $companies = @()
        try { $companies = Get-BdCompanies -BaseURL $Creds.Bitdefender.URL -ApiKey $Creds.Bitdefender.ApiKey }
        catch { Write-Status "  Bitdefender query failed: $($_.Exception.Message)" Warning }
        $company = Resolve-OrgServiceId -Entry $Entry -VendorKey 'bitdefender' -VendorLabel 'Bitdefender' -OrgName $OrgName -Candidates $companies
        if ($company) {
            $count = Get-BdEndpointCount -BaseURL $Creds.Bitdefender.URL -ApiKey $Creds.Bitdefender.ApiKey -CompanyId $company.Id
            $content = if ($null -ne $count) { "$count Protected" } else { "Protected" }
            return New-Tile -Title 'Bitdefender' -Shading 'success' -Content $content -Detail "Endpoints"
        }
    }

    return New-Tile -Title 'Antivirus' -Shading 'danger' -Content 'No AV'
}

function Get-DashDuo {
    param([pscustomobject]$Entry, [string]$OrgName, [hashtable]$Creds)
    if (-not $Creds.Duo.Configured) { return New-Tile -Title 'Duo' -Shading 'danger' -Content 'No Duo' }

    $accounts = @()
    try {
        $duoAuth = @{ Type = 'Accounts'; ApiHost = $Creds.Duo.ApiHost; IntegrationKey = $Creds.Duo.IntegrationKey; SecretKey = $Creds.Duo.SecretKey }
        Set-DuoApiAuth @duoAuth
        foreach ($a in @(Get-DuoAccounts)) {
            if (-not [string]::IsNullOrWhiteSpace($a.name)) {
                $accounts += [pscustomobject]@{ Id = $a.account_id; Name = $a.name; Provider = 'Duo' }
            }
        }
    } catch { Write-Status "  Duo query failed: $($_.Exception.Message)" Warning }

    $acct = Resolve-OrgServiceId -Entry $Entry -VendorKey 'duo' -VendorLabel 'Duo' -OrgName $OrgName -Candidates $accounts
    if (-not $acct) { return New-Tile -Title 'Duo' -Shading 'danger' -Content 'No Duo' }

    try {
        Select-DuoAccount -Name $acct.Name
        $users = @(Get-DuoUsers)
        $protected = @($users | Where-Object { $_.status -eq 'active' }).Count
        return New-Tile -Title 'Duo' -Shading 'success' -Content "$protected Users" -Detail "Protected by Duo"
    } catch {
        Write-Status "  Duo user count failed for '$($acct.Name)': $($_.Exception.Message)" Warning
        return New-Tile -Title 'Duo' -Shading 'warning' -Content 'Duo (count n/a)'
    }
}

function Get-DashBackup {
    param([pscustomobject]$Entry, [string]$OrgName, [hashtable]$Creds, [ref]$CoveDevicesOut)

    # Veeam first
    if ($Creds.VeeamVspc.Configured) {
        try {
            $token = Test-VspcAuth -BaseURL $Creds.VeeamVspc.URL -ApiKey $Creds.VeeamVspc.ApiKey
            $companies = Get-VspcCompanies -BaseURL $Creds.VeeamVspc.URL -Token $token
            $company = Resolve-OrgServiceId -Entry $Entry -VendorKey 'veeam' -VendorLabel 'Veeam VSPC' -OrgName $OrgName -Candidates $companies
            if ($company) {
                $b = Get-VspcCompanyBackup -BaseURL $Creds.VeeamVspc.URL -Token $token -CompanyUid $company.Id
                $detail = "$($b.Servers) servers - $($b.Workstations) workstations - $($b.VMs) VMs"
                return New-Tile -Title 'Veeam Backup' -Shading 'success' -Content "$($b.UsedTB) TB" -Detail $detail
            }
        } catch { Write-Status "  Veeam VSPC query failed: $($_.Exception.Message)" Warning }
    }

    # Then Cove
    if ($Creds.Cove.Configured) {
        try {
            $visa = Connect-Cove -Credential $Creds.Cove.Credential
            $partners = Get-CovePartners -Visa $visa -MyPartnerName $Creds.Cove.PartnerName
            $partner = Resolve-OrgServiceId -Entry $Entry -VendorKey 'cove' -VendorLabel 'Cove' -OrgName $OrgName -Candidates $partners
            if ($partner) {
                $devices = Get-CoveDevices -Visa $visa -PartnerId $partner.Id
                $CoveDevicesOut.Value = $devices   # cached for the M365-backup tile
                # Endpoint-style devices (exclude M365 cloud-to-cloud accounts)
                $endpoint = @($devices | Where-Object { $_.DataSources -notmatch '365|Exchange Online|OneDrive|SharePoint|Microsoft 365' })
                $servers = @($endpoint | Where-Object { $_.OSType -match 'server' }).Count
                $workstations = @($endpoint | Where-Object { $_.OSType -notmatch 'server' }).Count
                $tb = [Math]::Round((($endpoint | Measure-Object -Property UsedGB -Sum).Sum) / 1024, 2)
                $detail = "$servers servers - $workstations workstations"
                return New-Tile -Title 'Cove Backup' -Shading 'success' -Content "$tb TB" -Detail $detail
            }
        } catch { Write-Status "  Cove query failed: $($_.Exception.Message)" Warning }
    }

    return New-Tile -Title 'Backup' -Shading 'danger' -Content 'No Backup'
}

function Get-DashM365Backup {
    param([bool]$HasM365, [object[]]$CoveDevices)
    if (-not $HasM365) { return New-Tile -Title 'M365 Backup' -Shading 'blank' -Content 'N/A (no M365)' }
    if ($null -eq $CoveDevices) { return New-Tile -Title 'M365 Backup' -Shading 'danger' -Content 'None' }

    $m365 = @($CoveDevices | Where-Object { $_.DataSources -match '365|Exchange Online|OneDrive|SharePoint|Microsoft 365' -or $_.Product -match '365' })
    if ($m365.Count -eq 0) { return New-Tile -Title 'M365 Backup' -Shading 'danger' -Content 'None' }
    $tb = [Math]::Round((($m365 | Measure-Object -Property UsedGB -Sum).Sum) / 1024, 2)
    return New-Tile -Title 'M365 Backup' -Shading 'success' -Content "$($m365.Count) Protected" -Detail "$tb TB - Cove"
}

function Get-DashDomains {
    param([string]$OrgId, [string]$LinkBase)
    $domains = @()
    try {
        $res = Get-ITGlueDomains -filter_organization_id $OrgId
        if ($res.data) { $domains = $res.data }
    } catch { Write-Status "  IT Glue domains query failed: $($_.Exception.Message)" Warning }

    if (-not $domains -or $domains.Count -eq 0) {
        return New-Tile -Title 'Domains' -Shading 'danger' -Content 'No Domains'
    }

    $items = foreach ($d in ($domains | Sort-Object { $_.attributes.name })) {
        $name = Get-HtmlEncoded $d.attributes.name
        $href = "$LinkBase/domains/$($d.id)"
        [pscustomobject]@{ Shading = 'info'; AlertText = "<a href=`"$href`">$name</a>" }
    }
    return New-Tile -Title 'Domains' -Shading 'info' -IsInfo -InfoItems $items
}

function Get-DashM365 {
    param([string]$OrgId, [string]$LinkBase, [ref]$HasM365Out)
    $typeId = Get-ItgFlexAssetTypeId -NameLike @('*Microsoft Licenses*','*Office 365 Licenses*','*Microsoft 365 Licenses*')
    if (-not $typeId) {
        $HasM365Out.Value = $false
        return New-Tile -Title 'M365 Licenses' -Shading 'danger' -Content 'No M365'
    }

    $assets = @()
    try {
        $res = Get-ITGlueFlexibleAssets -filter_flexible_asset_type_id $typeId -filter_organization_id $OrgId
        if ($res.data) { $assets = $res.data }
    } catch { Write-Status "  M365 license asset query failed: $($_.Exception.Message)" Warning }

    if (-not $assets -or $assets.Count -eq 0) {
        $HasM365Out.Value = $false
        return New-Tile -Title 'M365 Licenses' -Shading 'danger' -Content 'No M365'
    }
    $HasM365Out.Value = $true

    # Build a license list. The synced trait names vary, so probe common keys per asset.
    $items = @()
    foreach ($a in $assets) {
        $traits = $a.attributes.traits
        $skuName = $null; $active = $null; $consumed = $null
        if ($traits) {
            foreach ($p in $traits.PSObject.Properties) {
                switch -Regex ($p.Name) {
                    'name|sku|product|subscription' { if (-not $skuName) { $skuName = "$($p.Value)" } }
                    'active|purchased|total|prepaid' { if ($null -eq $active -and $p.Value -match '^\d+$') { $active = $p.Value } }
                    'consumed|assigned|used'         { if ($null -eq $consumed -and $p.Value -match '^\d+$') { $consumed = $p.Value } }
                }
            }
        }
        $label = if ($skuName) { Get-HtmlEncoded $skuName } else { Get-HtmlEncoded $a.attributes.name }
        if ($null -ne $consumed -and $null -ne $active) { $label += " - $consumed/$active" }
        $href = "$LinkBase/assets/$($a.id)"
        $items += [pscustomobject]@{ Shading = 'success'; AlertText = "<a href=`"$href`">$label</a>" }
    }
    return New-Tile -Title 'M365 Licenses' -Shading 'info' -IsInfo -InfoItems $items
}

# ==============================================================================
# Rendering
# ==============================================================================
function ConvertTo-PanelHtml {
    param([pscustomobject]$Tile, [int]$Size = 3)
    $title = $Tile.Title
    if ($Tile.Link) { $title = "<a href=`"$($Tile.Link)`">$title</a>" }

    if ($Tile.IsInfo) {
        $content = if ($Tile.InfoItems.Count -gt 0) { ,$Tile.InfoItems } else { @([pscustomobject]@{ Shading='warning'; AlertText='None' }) }
        return New-BootstrapInfoPanel -PanelShading $Tile.Shading -PanelTitle $title -PanelContent $content -PanelSize $Size
    }
    return New-BootstrapSinglePanel -PanelShading $Tile.Shading -PanelTitle $title `
        -PanelContent $Tile.Content -ContentAsBadge -PanelAdditionalDetail $Tile.Detail -PanelSize $Size
}

function New-DashboardHtml {
    # Inner panels markup only - this is the exact string a future -Publish would store in IT Glue.
    param([pscustomobject[]]$Row1, [pscustomobject[]]$Row2)
    $r1 = ($Row1 | ForEach-Object { ConvertTo-PanelHtml -Tile $_ -Size 3 }) -join ''
    $r2 = ($Row2 | ForEach-Object { ConvertTo-PanelHtml -Tile $_ -Size 6 }) -join ''
    return "<div class=`"row`">$r1</div><div class=`"row`">$r2</div>"
}

function New-PreviewDocument {
    # Wrap the dashboard markup in a Bootstrap-3 CDN page for local browser preview only.
    param([string]$DashboardHtml, [string]$OrgName)
    @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>$([System.Net.WebUtility]::HtmlEncode($OrgName)) - All Services</title>
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap-theme.min.css">
</head>
<body>
<div class="container-fluid" style="margin-top:20px">
<h1>All $([System.Net.WebUtility]::HtmlEncode($OrgName)) Services</h1>
<p class="text-muted">Local preview - generated $(Get-Date -Format 'yyyy-MM-dd HH:mm')</p>
$DashboardHtml
</div>
</body>
</html>
"@
}

# ==============================================================================
# MAIN
# ==============================================================================
Write-Host "`n=== IT Glue Per-Client Services Dashboard ===" -ForegroundColor Cyan

# Modules
if (-not (Get-Module -ListAvailable -Name ITGlueAPI)) {
    Write-Status "Installing ITGlueAPI module..." Warning
    Install-Module -Name ITGlueAPI -Scope CurrentUser -Force -AllowClobber
}
Import-Module ITGlueAPI -Force

# Bootstrap helpers
. (Join-Path $ScriptDir 'lib\ITGlue-BootStrapHelpers.ps1')

# Credentials
$creds = Initialize-Credentials -Reset:$Reset
Add-ITGlueBaseURI -base_uri $creds.ITGlue.BaseUri
Add-ITGlueAPIKey  -Api_Key  $creds.ITGlue.ApiKey

# Duo module (only if configured)
if ($creds.Duo.Configured) {
    if (-not (Get-Module -ListAvailable -Name DuoSecurity)) {
        Write-Status "Installing DuoSecurity module..." Warning
        Install-Module -Name DuoSecurity -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module DuoSecurity -Force
}

# Resolve organization
if ([string]::IsNullOrWhiteSpace($Organization)) { $Organization = Read-Host "IT Glue organization name or ID" }

$org = Resolve-ItgOrganization -Query $Organization
if (-not $org) { throw "No IT Glue organization found matching '$Organization'." }

$orgId   = "$($org.id)"
$orgName = $org.attributes.name
$linkBase = "https://$($creds.ITGlue.Subdomain).itglue.com/$orgId"
Write-Status "Organization: $orgName (ID $orgId)" Success

# Service map
$script:ServiceMap = Get-ServiceMap
$script:ServiceMapDirty = $false
if ($Remap) {
    if ($script:ServiceMap.ContainsKey($orgId)) { $script:ServiceMap.Remove($orgId); $script:ServiceMapDirty = $true }
    Write-Status "Cleared cached resolution for this org (-Remap)." Detail
}
$entry = Get-OrgMapEntry -Map $script:ServiceMap -OrgId $orgId -OrgName $orgName

# -- Gather tiles --
Write-Status "`nGathering service data..." Info
$coveDevices = $null
$hasM365 = $false

$tileAv     = Get-DashAntivirus -Entry $entry -OrgName $orgName -Creds $creds
$tileDuo    = Get-DashDuo       -Entry $entry -OrgName $orgName -Creds $creds
$tileBackup = Get-DashBackup    -Entry $entry -OrgName $orgName -Creds $creds -CoveDevicesOut ([ref]$coveDevices)
$tileM365   = Get-DashM365      -OrgId $orgId -LinkBase $linkBase -HasM365Out ([ref]$hasM365)
$tileM365bk = Get-DashM365Backup -HasM365 $hasM365 -CoveDevices $coveDevices
$tileDomains = Get-DashDomains  -OrgId $orgId -LinkBase $linkBase

# Persist any new resolutions
if ($script:ServiceMapDirty) {
    $entry.resolvedOn = (Get-Date).ToString('yyyy-MM-dd')
    Save-ServiceMap -Map $script:ServiceMap
    Write-Status "Service-map cache updated: $ServiceMapPath" Detail
}

# -- Render --
$row1 = @($tileAv, $tileDuo, $tileBackup, $tileM365bk)
$row2 = @($tileM365, $tileDomains)
$dashboardHtml = New-DashboardHtml -Row1 $row1 -Row2 $row2

# Output
if ([string]::IsNullOrWhiteSpace($OutFile)) {
    if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
    $safe = ($orgName -replace '[^\w\-]+', '_').Trim('_')
    $OutFile = Join-Path $OutputDir "$safe-services.html"
}
New-PreviewDocument -DashboardHtml $dashboardHtml -OrgName $orgName | Set-Content -Path $OutFile -Encoding UTF8
Write-Status "`nDashboard written to: $OutFile" Success

if (-not $NoOpen) { Invoke-Item $OutFile }
Write-Host ""
