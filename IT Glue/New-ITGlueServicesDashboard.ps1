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

# This box's PowerShell 7 (pwsh) has a broken Invoke-RestMethod that returns $null, which makes
# ITGlueAPI / vendor REST calls fail with "Object reference not set...". Run under Windows
# PowerShell 5.1 (powershell.exe) only.
if ($PSVersionTable.PSVersion.Major -ge 6) {
    throw "Run this script under Windows PowerShell 5.1 (powershell.exe), not PowerShell $($PSVersionTable.PSVersion). pwsh 7's Invoke-RestMethod is broken on this machine."
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# -- Paths ---------------------------------------------------------------------
# This script keeps its OWN credentials file so it never shares config with other IT Glue
# scripts (e.g. Update-ITGlueCompanyInfo.ps1 owns Config\credentials.xml). On first run we may
# SEED the IT Glue section from that shared file, but we only ever WRITE our own file.
$ConfigDir         = Join-Path $ScriptDir 'Config'
$CredentialsPath   = Join-Path $ConfigDir 'dashboard-credentials.xml'
$SharedCredsPath   = Join-Path $ConfigDir 'credentials.xml'
$ServiceMapPath    = Join-Path $ConfigDir 'service-map.json'
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

    # -- Seed IT Glue creds from the shared file (read-only) so we don't re-prompt. We still
    #    only ever WRITE our own dashboard-credentials.xml. --
    if ((-not $creds.ITGlue) -and (Test-Path $SharedCredsPath)) {
        try {
            $shared = Import-Clixml -Path $SharedCredsPath
            if ($shared.ITGlue) {
                $creds.ITGlue = $shared.ITGlue
                $changed = $true
                Write-Status "Seeded IT Glue credentials from shared $($SharedCredsPath | Split-Path -Leaf) (into our own file)." Detail
            }
        } catch { }
    }

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

    # -- ConnectWise Manage: seed CWMConnectionInfo from any sibling repo Credentials.xml (read-only
    #    pull; we still only write our own file). Used for the ConnectWise contacts count. --
    if (-not $creds.CWM) {
        $repoRoot = Split-Path -Parent $ScriptDir
        $ci = $null; $ciFile = $null
        $ciFile = Get-ChildItem -Path $repoRoot -Recurse -Filter 'Credentials.xml' -ErrorAction SilentlyContinue |
            Where-Object { try { [bool]((Import-Clixml $_.FullName).CWMConnectionInfo) } catch { $false } } |
            Select-Object -First 1 -ExpandProperty FullName
        if ($ciFile) { try { $ci = (Import-Clixml $ciFile).CWMConnectionInfo } catch {} }
        if ($ci) {
            $creds.CWM = @{ Configured = $true; ConnectionInfo = $ci }
            Write-Status "Seeded ConnectWise Manage credentials from a sibling Credentials.xml (into our own file)." Detail
        } else {
            $creds.CWM = @{ Configured = $false }
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

    # Re-prompt if missing, or if an old username/password-shaped entry has no ApiToken yet.
    if ((-not $creds.Cove) -or ($creds.Cove.Configured -and -not $creds.Cove.ApiToken)) {
        if ((Read-Host "Configure Cove Data Protection (N-able)? (y/N)") -match '^[Yy]') {
            # API-key auth: create an API user in the Cove console (Management > Users, API access
            # enabled) and use its generated token as the password. API users skip 2FA.
            # Ref: https://developer.n-able.com/n-able-cove/docs/authorization
            Write-Host "  Cove uses an API USER (token), not your console login. Create one under" -ForegroundColor Gray
            Write-Host "  Management > Users (API access enabled); the token is shown only once." -ForegroundColor Gray
            $partner  = Read-Host "  Cove (MSP) Partner Name exactly as shown in the console (e.g. 'Contoso MSP LLC')"
            $apiUser  = Read-Host "  Cove API user name"
            $apiToken = Read-Host "  Cove API user token (used as the password)"
            $creds.Cove = @{ Configured = $true; PartnerName = $partner; ApiUser = $apiUser; ApiToken = $apiToken }
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

    if (-not $creds.Automate) {
        if ((Read-Host "Configure ConnectWise Automate (device links)? (y/N)") -match '^[Yy]') {
            $autoUrl = (Read-Host "  Automate URL (e.g. https://automate.yourdomain.com)").TrimEnd('/')
            $autoCid = Read-Host "  Automate API Client ID (GUID)"
            $autoUsr = Read-Host "  Automate API username (local-login capable)"
            $autoPwd = Read-Host "  Automate API password"
            $creds.Automate = @{ Configured = $true; Url = $autoUrl; ClientId = $autoCid; Username = $autoUsr; Password = $autoPwd }
        } else { $creds.Automate = @{ Configured = $false } }
        $changed = $true
    }

    if (-not $creds.CIPP) {
        if ((Read-Host "Configure CIPP (M365 - MFA / AD sync)? (y/N)") -match '^[Yy]') {
            $cippUrl = (Read-Host "  CIPP API URL (e.g. https://cippXXXX.azurewebsites.net)").TrimEnd('/')
            $cippTen = Read-Host "  CIPP/partner Tenant ID (GUID)"
            $cippCid = Read-Host "  API client Application (client) ID"
            $cippSec = Read-Host "  API client Secret"
            $creds.CIPP = @{
                Configured = $true; ApiUrl = $cippUrl; TenantId = $cippTen
                ClientId = $cippCid; ClientSecret = $cippSec; Scope = "api://$cippCid/.default"
            }
        } else { $creds.CIPP = @{ Configured = $false } }
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
    $nameMatches = @($Candidates | Where-Object { (ConvertTo-NormalizedName $_.Name) -eq $target })
    if ($nameMatches.Count -eq 1) {
        Write-Status "  [$VendorLabel] auto-matched '$($nameMatches[0].Name)'" Detail
        Add-Resolution -Entry $Entry -VendorKey $VendorKey -Candidate $nameMatches[0]
        return $nameMatches[0]
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

# ---- Duo (Accounts + Admin API via native REST/HMAC) ----
# We deliberately do NOT use the DuoSecurity PSGallery module: its manifest requires PowerShell
# 7.0, but this machine must run under Windows PowerShell 5.1 (pwsh 7's Invoke-RestMethod is
# broken here). The Duo API is a simple HMAC-SHA1 signed REST API, so we sign requests ourselves.
# One Accounts API (MSP-level) key enumerates all subaccounts; child Admin API calls are made by
# passing the child's account_id. Ref: https://duo.com/docs/accountsapi , https://duo.com/docs/adminapi
function Invoke-DuoApi {
    param(
        [Parameter(Mandatory)][string]$ApiHost,
        [Parameter(Mandatory)][string]$IKey,
        [Parameter(Mandatory)][string]$SKey,
        [ValidateSet('GET','POST')][string]$Method = 'GET',
        [Parameter(Mandatory)][string]$Path,
        [hashtable]$Params = @{}
    )
    $hostL = $ApiHost.ToLowerInvariant()
    $date  = [DateTime]::UtcNow.ToString("ddd, dd MMM yyyy HH:mm:ss '-0000'", [Globalization.CultureInfo]::InvariantCulture)

    # Canonical params: sorted, RFC-3986 percent-encoded, joined with '&'.
    $canonParams = ''
    if ($Params.Count -gt 0) {
        $pairs = foreach ($k in ($Params.Keys | Sort-Object)) {
            '{0}={1}' -f [Uri]::EscapeDataString($k), [Uri]::EscapeDataString([string]$Params[$k])
        }
        $canonParams = ($pairs -join '&')
    }

    # Sign: date\nMETHOD\nhost\npath\nparams  (HMAC-SHA1, hex), Basic auth = base64(ikey:sig).
    $canon = ($date, $Method.ToUpper(), $hostL, $Path, $canonParams) -join "`n"
    $hmac  = New-Object System.Security.Cryptography.HMACSHA1
    $hmac.Key = [Text.Encoding]::ASCII.GetBytes($SKey)
    $sig = (($hmac.ComputeHash([Text.Encoding]::ASCII.GetBytes($canon)) | ForEach-Object { $_.ToString('x2') }) -join '')
    # Use X-Duo-Date (not Date): in Windows PowerShell 5.1, 'Date' is a restricted header that
    # Invoke-RestMethod silently refuses to set from -Headers, so the server never sees the value
    # we signed with -> 401. Duo supports X-Duo-Date for exactly this case.
    $headers = @{
        Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${IKey}:${sig}"))
        'X-Duo-Date'  = $date
    }

    $uri = "https://$hostL$Path"
    if ($Method.ToUpper() -eq 'GET') {
        if ($canonParams) { $uri += "?$canonParams" }
        return Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
    }
    return Invoke-RestMethod -Uri $uri -Method Post -Headers $headers `
        -ContentType 'application/x-www-form-urlencoded' -Body $canonParams
}

function Get-DuoAccountList {
    param([hashtable]$Creds)
    $resp = Invoke-DuoApi -ApiHost $Creds.Duo.ApiHost -IKey $Creds.Duo.IntegrationKey -SKey $Creds.Duo.SecretKey `
        -Method POST -Path '/accounts/v1/account/list'
    return @($resp.response)
}

function Get-DuoUserCounts {
    # Page the child account's Admin API users (300/page) and tally by status. Duo statuses include
    # 'active', 'bypass', 'disabled', 'locked out'. We report active and bypass.
    param([hashtable]$Creds, [string]$AccountId)
    $active = 0; $bypass = 0; $offset = 0
    do {
        $resp = Invoke-DuoApi -ApiHost $Creds.Duo.ApiHost -IKey $Creds.Duo.IntegrationKey -SKey $Creds.Duo.SecretKey `
            -Method GET -Path '/admin/v1/users' -Params @{ account_id = $AccountId; limit = '300'; offset = "$offset" }
        $active += @($resp.response | Where-Object { $_.status -eq 'active' }).Count
        $bypass += @($resp.response | Where-Object { $_.status -eq 'bypass' }).Count
        $next = $resp.metadata.next_offset
        if ($null -ne $next -and "$next" -ne '') { $offset = [int]$next } else { $offset = $null }
    } while ($null -ne $offset)
    return [pscustomobject]@{ Active = $active; Bypass = $bypass }
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
    # getCompaniesList lives on the NETWORK service (Partners product only) and returns a FLAT array
    # of {id,name} - it is not paginated and needs no params.
    # Ref: https://www.bitdefender.com/business/support/en/77211-128478-getcompanieslist.html
    param([string]$BaseURL, [string]$ApiKey)
    $res = Invoke-BdJsonRpc -BaseURL $BaseURL -ApiKey $ApiKey -Service 'network' -Method 'getCompaniesList'
    $companies = @()
    foreach ($c in $res) {
        $companies += [pscustomobject]@{ Id = $c.id; Name = $c.name; Provider = 'Bitdefender' }
    }
    return $companies
}

function Get-BdEndpointCount {
    # Count endpoints actually PROTECTED by Bitdefender (BEST agent installed) for the company,
    # recursing into subgroups. getEndpointsList 'total' includes unmanaged/stale inventory, so we
    # count items with managedWithBest=true instead. Max perPage is 100.
    param([string]$BaseURL, [string]$ApiKey, [string]$CompanyId)
    try {
        $count = 0; $page = 1
        do {
            $res = Invoke-BdJsonRpc -BaseURL $BaseURL -ApiKey $ApiKey -Service 'network' `
                -Method 'getEndpointsList' -Params @{
                    parentId = $CompanyId; page = $page; perPage = 100
                    filters  = @{ depth = @{ allItemsRecursively = $true } }
                }
            foreach ($it in $res.items) { if ($it.managedWithBest) { $count++ } }
            $page++
        } while ($res.items -and $page -le [int]$res.pagesCount)
        return $count
    } catch { }
    return $null
}

# ---- Cove Data Protection (backup.management JSON-RPC, visa auth) ----
function Connect-Cove {
    # API-user auth: partner + API-user name + token-as-password -> session visa (15-min validity).
    # The 'partner' param is REQUIRED; omitting it returns "Unknown partner/username or bad password".
    param([string]$PartnerName, [string]$ApiUser, [string]$ApiToken)
    $url = 'https://api.backup.management/jsonapi'
    $data = @{ jsonrpc='2.0'; id='login'; method='Login'; params=@{
        partner  = $PartnerName
        username = $ApiUser
        password = $ApiToken
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

function Get-VspcPaged {
    # GET a VSPC v3 collection, paging at the API max of 500 (limit>500 returns 400). Returns all items.
    param([string]$BaseURL, [string]$Token, [string]$Path)
    $items = @(); $offset = 0; $limit = 500
    $h = @{ Authorization = "Bearer $Token" }
    do {
        $sep  = if ($Path -match '\?') { '&' } else { '?' }
        $resp = Invoke-RestMethod -Uri "$BaseURL$Path${sep}limit=$limit&offset=$offset" -Method Get -Headers $h
        if ($resp.data) { $items += $resp.data }
        $total = [int]$resp.meta.pagingInfo.total
        $offset += $limit
    } while (($offset -lt $total) -and $resp.data)
    return $items
}

function Get-VspcCompanies {
    param([string]$BaseURL, [string]$Token)
    $companies = @()
    foreach ($c in (Get-VspcPaged -BaseURL $BaseURL -Token $Token -Path '/api/v3/organizations/companies')) {
        $companies += [pscustomobject]@{ Id = $c.instanceUid; Name = $c.name; Provider = 'Veeam' }
    }
    return $companies
}

function Get-VspcCompanyBackup {
    # Protected-workload counts for the company. The ?filter= param breaks the computers endpoints
    # (they fall through to the SPA), so we page each list and filter client-side by organizationUid.
    # NOTE: no reliable used-storage (TB) endpoint in this VSPC build - backupResources returns the
    # SPA - so TB is omitted for now (see follow-up notes). operationMode gives Server/Workstation.
    param([string]$BaseURL, [string]$Token, [string]$CompanyUid)
    # Returns the NAMES of protected workloads per type (not just counts).
    $result = [pscustomobject]@{ Servers = @(); Workstations = @(); VMs = @() }
    try {
        $vms = Get-VspcPaged -BaseURL $BaseURL -Token $Token -Path '/api/v3/protectedWorkloads/virtualMachines' |
            Where-Object { $_.organizationUid -eq $CompanyUid }
        $result.VMs = @($vms | ForEach-Object { $_.name } | Where-Object { $_ } | Sort-Object)
    } catch { }
    try {
        $agents = Get-VspcPaged -BaseURL $BaseURL -Token $Token -Path '/api/v3/protectedWorkloads/computersManagedByConsole' |
            Where-Object { $_.organizationUid -eq $CompanyUid }
        $result.Servers      = @($agents | Where-Object { $_.operationMode -eq 'Server' }      | ForEach-Object { $_.name } | Sort-Object)
        $result.Workstations = @($agents | Where-Object { $_.operationMode -eq 'Workstation' } | ForEach-Object { $_.name } | Sort-Object)
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
# Brand styling (colors + logos) via Brandfetch
# Known brands are baked below (accent color + light/colored icon URL) so rendering makes no API
# call; any unmapped domain is resolved live and cached. The API key is hard-coded per request.
# SECURITY: this key is visible to anyone with the repo - restrict it in Brandfetch (domain
# allowlist) or rotate if the repo is shared.
# ==============================================================================
$script:BrandfetchApiKey = 'ob7cfkVJLOtwFWiIVN60OjkZywKeFtvnS6l95rfZ7oC0cGUHDtjhcpy_Kt9JuX6l96rWcyWYyP_zCn5D3MPZHA'
$script:BrandMap = @{
    'itglue.com'      = @{ Color = '#3860be'; Logo = 'https://cdn.brandfetch.io/idwyCKX16_/w/350/h/350/theme/dark/icon.jpeg?c=1bxljtlqy4vv80d5kade1ohx1blb2u0lzBN' }
    'connectwise.com' = @{ Color = '#5ea4de'; Logo = 'https://cdn.brandfetch.io/idejubPisD/w/400/h/400/theme/dark/icon.jpeg?c=1bxljtlqy4vv80d5kade1ohx1blb2u0lzBN' }
    'sentinelone.com' = @{ Color = '#6b0aea'; Logo = 'https://cdn.brandfetch.io/idqbZJLrXa/w/400/h/400/theme/dark/icon.png?c=1bxljtlqy4vv80d5kade1ohx1blb2u0lzBN' }
    'bitdefender.com' = @{ Color = '#EB0000'; Logo = 'https://cdn.brandfetch.io/idtcaX4QNF/w/400/h/400/theme/dark/icon.png?c=1bxljtlqy4vv80d5kade1ohx1blb2u0lzBN' }
    'duo.com'         = @{ Color = '#74bf4b'; Logo = 'https://cdn.brandfetch.io/id7s-uuKhk/w/400/h/400/theme/dark/icon.png?c=1bxljtlqy4vv80d5kade1ohx1blb2u0lzBN' }
    'veeam.com'       = @{ Color = '#03D15F'; Logo = 'https://cdn.brandfetch.io/idVHk_jeH3/w/400/h/400/theme/dark/icon.jpeg?c=1bxljtlqy4vv80d5kade1ohx1blb2u0lzBN' }
    'n-able.com'      = @{ Color = '#c046ff'; Logo = 'https://cdn.brandfetch.io/idpmq52ZKx/w/400/h/400/theme/dark/icon.jpeg?c=1bxljtlqy4vv80d5kade1ohx1blb2u0lzBN' }
    'microsoft.com'   = @{ Color = '#00A4EF'; Logo = 'https://cdn.brandfetch.io/idchmboHEZ/w/800/h/800/theme/dark/symbol.png?c=1bxljtlqy4vv80d5kade1ohx1blb2u0lzBN' }
    # ConnectWise Automate: no Brandfetch entry / favicon - logo is an inlined data URI (downscaled
    # from the supplied automate.png) so it's self-contained.
    'automate'        = @{ Color = '#57b947'; Logo = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAArqSURBVHhe3Vt/jFxVFV4RxGpAEFCMBEGaYNCAutmd997M7tTSdue9N7PbUrZIQgygVlskUAkQf9TVkviH2mibQNNGA02lSSvQ7s6bNzM7224pVdtSBNtCgNRqxaYpNCimhqKU8fvunGm2u/fNe/NjG7pfcrKZmXvOPefcc8859763bWcaxuCc6YntfYutnP1d03ceNHP2UivnPmCVMnd3Dvb0ybCpCyNnr5u5d0G5e8fc0yj5x3nlxLbesrW551oZOvVwfXH2Rw3PfiNeSpfNnDOBun8/r2x4zoAMn3qw8m5v1/Y+rfGkxNZeOMB+uW1g4BxhmVpA+K/vfmau1nhFvjghZ8eFZeqg009dCMPejA/rw79KzAem56wStqkD00/d3FVr9YWYH7ANjrYPpT8irO8fJAeS53Y93deTHE2eK19FBpLbE1EcQGKeMHx3nrBGRtxPGZhngXxsLTrXpS60iu5wcudNZWvY3Z8oZRYaG26eJj/XxHQ/db6Zc9+yQsK/SspRnvOUsIfCLLpfgOy1jJ7uP8wrm1l3qfzUGnQ8deMlVjG9h2WKCjJRca9iPx+wSukl7aVZH5OhE5AouZ+NF92fQMGTTHLjjdWRVXDLZsH9tzmcvqMr3/MpETUBCT91HeT+Brr9j/qgyihSenrOchnWHLAXL4Uiz0tyOk3R+JYMJsPExfQ/zGL6h2Zx9ieErQ1R0gvlsoiaE9g2p/GFEuahExgJkP0WaH08784Q0SinPdditddA9n+V4XTYWH46AY0Vom6lsDSGDs+5HMK0xo+leImOwIR55w3swV+Znv0M9zFpgnJ1EoxXjmDUQYch0CpExwmt4WMJ0cbuEuOX92/s/6CYVB/MrP3kjN3zaxo/lljiTimr+b0pgkFd2+BURkVUpzISkBMMP5UUk+oDStedKiHphJ8FFB/J8O9rtXJUTTB0jKzzyqSs6BkgOVvcL+Y0hpjn3K5ygGaCRojhm0DyZGJkdFE2//Izv282Z1RJdZy+88+E514spjSG9j0Lz8OZ/SCV000UhVialME8CPnO24bv7MP3g1idNcgzK/HdGtAgvtuPfXui6hDyjZVTD8nq/0LMaA7o4b9d7QHqIiYuJi1kcrOQzhkF92vmFuczIlaL2GjfVVYxcztKXJ6NkzhNLz+AGEWgd1AqrxSxzQGt74ch+DVJKuGEqsGxzNow4ndGPv1lEVUXzFK6A/3EJkaDmjtiNeK2MrL2OhHTGhjZ1L2qzdRMeBpBSa4avH/UKqTnC3tTgJxbsKeP0RGhTsC2wdj3LM/5orC3Bqqf95wXa0YBjUfIQ4HnOzfMvlpYWwLeJULuPsqv5QQ6KZa1h4StdbB29F5g+Pbf2PXpJq6uPPbuczc8mrxI2FoKdSYpuHtVXghwQqVjtHe1lVt8q4RsvTTJLVBr4oJ7uGOjc7mwTAo6n0xdgbxwlOcQnR4k5oBYzr5FWJoHk6Dh2YeDwp/limXSGGqw5awTsUF7dmK0N7BMMvkiCe6U4c0jlnV71d7TTMaIkKy7WoaHgq1pfCRtm8PufagSA/Fi+q7E1ky8rT/6oSXm2Y+pBk0XkT4SYSldjg9mPi/DmwPqcOBtjqrxefd4bCTzSRkeCGO5MQ17+CEYfYTJigZUiVsI37+IfuEOGV4TnX7qCmy5t4M6R7UoeffnMrx+mKPpzyGhPYAJdqpJApoRNVGE1W/fOOtKZPE/8UaJqzNBFuTTCSy1mG99lGs3JLtKFIyXJfKwRd5T+sMO2iNswbC2zL0mXnDvRZJ52iw47/Kyg3stsBOj0uzf82lTRGjBG2GE5cuUF5RETxFkJnfdxHGPC3sgsDVnhOqH39W8sId2Kftgp4hAghtKX4ofFiGUS/DWOyoc+YgqQh/Okojk+Pf21e3niTgtGCFc+VDjqwTFVdPlObeKCC1Ucs7aR7QRNY5UooZdtI92KnuHM4vazJx7z8w/91fa1jpPYqzHOHBsEn206BycfbXyfp2y1cpm7ZfCbnMwNqc6RI2MIKIutHfmCwvKCKPUg4H7KIQUn+f8VHTRAifJexo9SLHsxrw57SJKC8y/vHH955VVg9OwAJ74POdu0UULOOCxmo/DapDiy6buFFFaGEPO/U05wPBSP2pUgDrq5uxviC5a4Oy/SbWuGv4wol7YYktElBaIgMUNO1g5IGt/v5kIAP93RBctJjsCEMH3NR4BcxkBzhJVJjQDwkitUM5+SHTRYtJzQM7+WVMOYKmZ8ez8SjhHfHRVpUpo2xtFFy0mvwrYg3VXAbm+55W/qqVW3ulDl/ZbKHmMP6hjbQSFuULYAgfDlJysPoD9B/uQwCP6GFKlD3aphYadyl7YLaIqSI6iKSpmbkNjNGQVnP8wRNg8UCGd0MqhA04IufKarE6QT4PVsbiGfgnUe9pBe5RdsI92iohg8ELRLKUXo2MaZU8dNEnFKPuXwhaIus4CfrSzAFb/kcD9D+NxSDtJ/WkH7RG2+gGBQ0GlDEbxmeCbyU19obdArTwN8qEtZPHBqVYvteq+u0KGNwfuxUBPgxgFxqAd+ejZkvuArL0ycEtJNJlDqQ4Z3hyszb0XQPCxoEMHE4xVSr9rbk59SVgmFbHNTgwJ+CTn1elTqSDOXgz9QIWjBVD1NuhOEN+pSQvuK3SWsEwKEo+7FyPs/6LmC0ioKiJDOtS6wZcfkAyPBvYKUEbtu7y75bqN/R8StpbC2GBMM4vOdtUhBhjPqoDkuJ+P9IStNWCtR7Y/GHQxqohOgPdRbrZ0rL3xEmFtCRIbUpchV2xXuSjAeBL3vpFNPStsrQMmXVwrEZ4iOoHNRin9asMvJowDtt9MOPUgm5haxleJ87f0WpxPhbCnDtW6iz+NoCT3KJskazjzcCzfd5WIqguWn7oGW241o05l9QjGkzg3tsG+tnKLkiAOTN9q5EDDayiVFwrucZS8R42Ck5pV6q/5tkZytO8irLaDRLfWKrqqG6UcnfxapFrenN38s0kmNHj+r029H1B90QkrY/nOUXRpW42csxqOXRbznB8YvrMMtZvvCIyie3ud46Rn18qLQhIFL4gZjQNKfl01HJpJ6iY0KOwceb6ggVzdKikH4XvVWQb19nWScqI37sBTD9iPwwEH0KVpJ3i/kzrIefZzYk79QPa9i6ujE362kDrD+O5XxaT6gHP8+kbeE6zZKzRBrAQqrAMOQDriZY+ZbfCNUZ7wkJR2BB46hGiw5InXETW/Br3Kz0xEze5nVoDqC5JInLuxJZks/8WqFPZARLXuOSfX1Kv3lbfE07tUGRznBFYGVaaK7qF40f1ewk9dRh5OiFK2CDV8D3sHKj+WLxLhPF+Rnab8bYnhzK3Vum6MuJ+G7GX47QgdrYs4ZXzRHZ6+InU+eZrCDZuSrM07xKNqZSUUX0rw0qTGASg+nDHQvz8S5drqFPEyo+CeNEvujxMjmetF1AQYG+Z8nEdqbL0DdJZqlhBxfEcYkZNvifFVqFUtuh6dEB92d1sjmdvCngmOBcL3cFQnqBcccvZ2YQ0F/28BC/RNLMheZXzBXR92T9kYBtrO6d7W1yWf6gK2z6qoFUWN89yFwhoZNLp7R2P6TTo6PecrUZIi9zz+Hq/mkykDucI+FFYiVbb3nCeEbWoBh5MVYdtAPZPIu73CMrWAxBavZmqd8ZUkaR+J+s9YZx2YpIyscyDoXoE1Hd3nwzJ8agLVYPnMfRP/e5yNliqxfsqQoWcAbW3/B5+yVWJNKxhhAAAAAElFTkSuQmCC' }
}
function Resolve-Brand {
    param([string]$Domain)
    if ([string]::IsNullOrWhiteSpace($Domain)) { return $null }
    if ($script:BrandMap.ContainsKey($Domain)) { return $script:BrandMap[$Domain] }
    try {
        $b = Invoke-RestMethod -Uri "https://api.brandfetch.io/v2/brands/$Domain" -TimeoutSec 20 `
            -Headers @{ Authorization = "Bearer $($script:BrandfetchApiKey)" }
        $color = ($b.colors | Where-Object { $_.type -eq 'accent' } | Select-Object -First 1).hex
        if (-not $color) { $color = ($b.colors | Where-Object { $_.type -eq 'brand' } | Select-Object -First 1).hex }
        $icon = $b.logos | Where-Object { $_.type -in 'icon','symbol' } | Select-Object -First 1
        $fmt  = $icon.formats | Where-Object { $_.format -eq 'png' } | Select-Object -First 1
        if (-not $fmt) { $fmt = $icon.formats | Select-Object -First 1 }
        $entry = @{ Color = $color; Logo = $fmt.src }
        $script:BrandMap[$Domain] = $entry
        return $entry
    } catch { return $null }
}

# ==============================================================================
# Tile providers - each returns a card object: @{ Title; Items[] }
# An item is @{ Label; Sub; Link; Brand; Bg; Muted; Kind }:
#   Brand -> domain key into BrandMap; sets the logo + accent color.
#   Bg    -> explicit background colour that overrides everything (e.g. red for an alert).
#   Muted -> force the neutral translucent background even with a brand (secondary 365/license lines).
#   Kind 'pill' = logo + Label button; Kind 'number' = big Label over small Sub (always muted, no logo).
# ==============================================================================
function New-CardItem {
    param(
        [string]$Label, [string]$Sub = '', [string]$Link = '', [string]$Brand = '',
        [string]$Bg = '', [switch]$Muted, [ValidateSet('pill','number')][string]$Kind = 'pill'
    )
    [pscustomobject]@{ Label = $Label; Sub = $Sub; Link = $Link; Brand = $Brand; Bg = $Bg; Muted = [bool]$Muted; Kind = $Kind }
}

function New-Card {
    param([string]$Title, [object[]]$Items)
    [pscustomobject]@{ Title = $Title; Items = @($Items) }
}

function Get-DashAntivirus {
    # Category tile: lists every active AV product the client has (SentinelOne and/or Bitdefender).
    param([pscustomobject]$Entry, [string]$OrgName, [hashtable]$Creds)
    $lines = @()

    if ($Creds.SentinelOne.Configured) {
        $sites = @()
        foreach ($inst in $Creds.SentinelOne.Instances) {
            try { $sites += Get-S1Sites -BaseURL $inst.URL -APIKey $inst.APIKey -InstanceName $inst.Name }
            catch { Write-Status "  SentinelOne $($inst.Name) query failed: $($_.Exception.Message)" Warning }
        }
        $site = Resolve-OrgServiceId -Entry $Entry -VendorKey 'sentinelOne' -VendorLabel 'SentinelOne' -OrgName $OrgName -Candidates $sites
        # Hide a 0-count product (matched by name but protecting nothing).
        if ($site -and [int]$site.ActiveLicenses -gt 0) { $lines += New-CardItem -Label "$([int]$site.ActiveLicenses) devices" -Brand 'sentinelone.com' }
    }

    if ($Creds.Bitdefender.Configured) {
        $companies = @()
        try { $companies = Get-BdCompanies -BaseURL $Creds.Bitdefender.URL -ApiKey $Creds.Bitdefender.ApiKey }
        catch { Write-Status "  Bitdefender query failed: $($_.Exception.Message)" Warning }
        $company = Resolve-OrgServiceId -Entry $Entry -VendorKey 'bitdefender' -VendorLabel 'Bitdefender' -OrgName $OrgName -Candidates $companies
        if ($company) {
            $count = Get-BdEndpointCount -BaseURL $Creds.Bitdefender.URL -ApiKey $Creds.Bitdefender.ApiKey -CompanyId $company.Id
            if ([int]$count -gt 0) { $lines += New-CardItem -Label "$([int]$count) devices" -Brand 'bitdefender.com' }
        }
    }

    return New-Card -Title 'Endpoint Security' -Items $lines
}

function Get-DashMfa {
    # Category tile: MFA providers the client has (currently Duo).
    param([pscustomobject]$Entry, [string]$OrgName, [hashtable]$Creds)
    $lines = @()
    if ($Creds.Duo.Configured) {
        $accounts = @()
        try {
            foreach ($a in (Get-DuoAccountList -Creds $Creds)) {
                if (-not [string]::IsNullOrWhiteSpace($a.name)) {
                    $accounts += [pscustomobject]@{ Id = $a.account_id; Name = $a.name; Provider = 'Duo'; ApiHostname = $a.api_hostname }
                }
            }
        } catch { Write-Status "  Duo query failed: $($_.Exception.Message)" Warning }

        $acct = Resolve-OrgServiceId -Entry $Entry -VendorKey 'duo' -VendorLabel 'Duo' -OrgName $OrgName -Candidates $accounts
        if ($acct) {
            # Duo's per-customer admin panel is the api_hostname with 'api-' swapped for 'admin-'.
            $adminUrl = if ($acct.ApiHostname) { 'https://' + ($acct.ApiHostname -replace '^api-', 'admin-') } else { '' }
            try {
                $c = Get-DuoUserCounts -Creds $Creds -AccountId $acct.Id
                $lines += New-CardItem -Label "Duo: $($c.Active) active, $($c.Bypass) bypass" -Link $adminUrl -Brand 'duo.com'
            }
            catch { Write-Status "  Duo user count failed for '$($acct.Name)': $($_.Exception.Message)" Warning; $lines += New-CardItem -Label 'Duo: n/a' -Brand 'duo.com' }
        }
    }
    return New-Card -Title 'MFA' -Items $lines
}

function Get-DashBackup {
    # Category tile: lists every active backup product. Endpoint/server backup (Veeam, Cove) and
    # Cove M365 cloud backup all live here; the M365 line is explicitly labelled "365" so it is not
    # confused with server/workstation backup.
    param([pscustomobject]$Entry, [string]$OrgName, [hashtable]$Creds)
    $lines = @()

    if ($Creds.VeeamVspc.Configured) {
        try {
            $token = Test-VspcAuth -BaseURL $Creds.VeeamVspc.URL -ApiKey $Creds.VeeamVspc.ApiKey
            $companies = Get-VspcCompanies -BaseURL $Creds.VeeamVspc.URL -Token $token
            $company = Resolve-OrgServiceId -Entry $Entry -VendorKey 'veeam' -VendorLabel 'Veeam VSPC' -OrgName $OrgName -Candidates $companies
            if ($company) {
                $b = Get-VspcCompanyBackup -BaseURL $Creds.VeeamVspc.URL -Token $token -CompanyUid $company.Id
                # List protected device names, ordered servers -> workstations -> VMs.
                $names = @($b.Servers) + @($b.Workstations) + @($b.VMs)
                if ($names.Count -gt 0) {
                    $lines += New-CardItem -Label "Veeam: $($names -join ', ')" -Brand 'veeam.com'
                }
            }
        } catch { Write-Status "  Veeam VSPC query failed: $($_.Exception.Message)" Warning }
    }

    if ($Creds.Cove.Configured) {
        try {
            $visa = Connect-Cove -PartnerName $Creds.Cove.PartnerName -ApiUser $Creds.Cove.ApiUser -ApiToken $Creds.Cove.ApiToken
            $partners = Get-CovePartners -Visa $visa -MyPartnerName $Creds.Cove.PartnerName
            $partner = Resolve-OrgServiceId -Entry $Entry -VendorKey 'cove' -VendorLabel 'Cove' -OrgName $OrgName -Candidates $partners
            if ($partner) {
                $devices = Get-CoveDevices -Visa $visa -PartnerId $partner.Id
                # Cove M365 cloud accounts have Physicality 'Undefined'; endpoint/server backups do
                # not. Split them so the M365 line is clearly labelled. Ref: CWM-SyncNableCoveData.ps1.
                $endpoint = @($devices | Where-Object { "$($_.Physicality)" -ne 'Undefined' })
                if ($endpoint.Count -gt 0) {
                    $servers = @($endpoint | Where-Object { $_.OSType -match 'server' }).Count
                    $workstations = $endpoint.Count - $servers
                    $tb = [Math]::Round((($endpoint | Measure-Object -Property UsedGB -Sum).Sum) / 1024, 2)
                    $lines += New-CardItem -Label "Cove: $tb TB ($servers srv, $workstations wks)" -Brand 'n-able.com'
                }
                $m365 = @($devices | Where-Object { "$($_.Physicality)" -eq 'Undefined' })
                if ($m365.Count -gt 0) {
                    $tb365 = [Math]::Round((($m365 | Measure-Object -Property UsedGB -Sum).Sum) / 1024, 2)
                    $tenants = if ($m365.Count -eq 1) { '1 tenant' } else { "$($m365.Count) tenants" }
                    $lines += New-CardItem -Label "Cove 365: $tb365 TB ($tenants)" -Brand 'n-able.com' -Muted
                }
            }
        } catch { Write-Status "  Cove query failed: $($_.Exception.Message)" Warning }
    }

    return New-Card -Title 'Backup' -Items $lines
}

function Get-DashDomains {
    param([string]$OrgId, [string]$LinkBase)
    $domains = @()
    try {
        $res = Get-ITGlueDomains -filter_organization_id $OrgId
        if ($res.data) { $domains = $res.data }
    } catch { Write-Status "  IT Glue domains query failed: $($_.Exception.Message)" Warning }

    $items = foreach ($d in ($domains | Sort-Object { $_.attributes.name })) {
        # Muted pill, no icon (matches the design's domain list).
        New-CardItem -Label $d.attributes.name -Link "$LinkBase/domains/$($d.id)"
    }
    return New-Card -Title 'Domains' -Items $items
}

# ---- CIPP (M365 data via the CIPP API: MFA + Entra/AD sync) ----
function Invoke-CippGraph {
    # Proxy a Graph request through CIPP (handles the multi-tenant GDAP auth for us).
    param([string]$TenantFilter, [string]$Endpoint)
    $uri = "$($script:CippApiUrl)/api/ListGraphRequest?TenantFilter=$([uri]::EscapeDataString($TenantFilter))&Endpoint=$([uri]::EscapeDataString($Endpoint))"
    $r = Invoke-RestMethod -Uri $uri -Headers @{ Authorization = "Bearer $($script:CippToken)" }
    if ($r -isnot [array] -and ($r.PSObject.Properties.Name -contains 'Results')) { return $r.Results }
    return $r
}
function Resolve-CippTenant {
    # Map the IT Glue org to its CIPP/M365 tenant. We query ListTenants?TenantFilter=<domain> for each
    # of the org's IT Glue domains; the tenant's DEFAULT domain resolves to a single clean tenant
    # (secondary domains 400 in Graph, and CIPP's occasional aggregated all-tenants row - a single
    # object whose defaultDomainName is many domains joined by spaces - is rejected). Cached on success;
    # a miss is NOT cached (re-resolve next run) so transient CIPP responses can't poison the cache.
    param([pscustomobject]$Entry, [string]$OrgId)
    if (-not $script:CippConnected) { return $null }
    $cached = $Entry.PSObject.Properties['cipp']
    if ($cached -and $cached.Value -and "$($cached.Value.Domain)" -notmatch '\s') { return $cached.Value }

    $orgDomains = @()
    try { $orgDomains = @((Get-ITGlueDomains -filter_organization_id $OrgId).data.attributes.name) } catch {}
    $orgDomains = @($orgDomains | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() })

    $match = $null
    foreach ($d in $orgDomains) {
        $t = $null
        try { $t = @(Invoke-RestMethod -Uri "$($script:CippApiUrl)/api/ListTenants?TenantFilter=$([uri]::EscapeDataString($d))" -Headers @{ Authorization = "Bearer $($script:CippToken)" })[0] } catch {}
        if ($t -and $t.defaultDomainName -and "$($t.defaultDomainName)" -notmatch '\s') { $match = $t; break }
    }
    if (-not $match) { return $null }
    $cid = if ($match.customerId) { $match.customerId } else { $match.RowKey }
    $val = [pscustomobject]@{ Id = "$cid"; Name = $match.displayName; Domain = $match.defaultDomainName }
    $Entry | Add-Member -NotePropertyName 'cipp' -NotePropertyValue $val -Force
    $script:ServiceMapDirty = $true
    return $val
}
function Get-CippMfaItem {
    param([string]$TenantFilter)
    try {
        $rows = @(Invoke-CippGraph -TenantFilter $TenantFilter -Endpoint 'reports/authenticationMethods/userRegistrationDetails')
        $reg  = @($rows | Where-Object { $_.isMfaRegistered -eq $true }).Count
        return New-CardItem -Label "MFA: $reg/$($rows.Count)" -Brand 'microsoft.com' -Muted
    } catch { Write-Status "  CIPP MFA query failed: $($_.Exception.Message)" Warning; return $null }
}
function Get-CippSyncItem {
    param([string]$TenantFilter)
    try {
        $o = @(Invoke-CippGraph -TenantFilter $TenantFilter -Endpoint 'organization')[0]
        if ($o.onPremisesSyncEnabled) {
            $rel = ''
            if ($o.onPremisesLastSyncDateTime) {
                $span = (Get-Date).ToUniversalTime() - ([datetime]$o.onPremisesLastSyncDateTime).ToUniversalTime()
                $rel  = if ($span.TotalMinutes -lt 60) { " ($([int]$span.TotalMinutes)m ago)" }
                        elseif ($span.TotalHours -lt 24) { " ($([int]$span.TotalHours)h ago)" }
                        else { " ($([int]$span.TotalDays)d ago)" }
            }
            return New-CardItem -Label "AD Sync: Synced$rel" -Brand 'microsoft.com' -Muted
        }
        return New-CardItem -Label 'AD Sync: Cloud-only' -Brand 'microsoft.com' -Muted
    } catch { Write-Status "  CIPP org/sync query failed: $($_.Exception.Message)" Warning; return $null }
}

function Get-DashM365 {
    # M365 card: MFA + AD-sync (via CIPP, tenant matched by domain) on top, then the license list.
    param([pscustomobject]$Entry, [string]$OrgId, [string]$LinkBase)
    $items = @()

    if ($script:CippConnected) {
        $tenant = Resolve-CippTenant -Entry $Entry -OrgId $OrgId
        if ($tenant) {
            $mfa  = Get-CippMfaItem  -TenantFilter $tenant.Domain; if ($mfa)  { $items += $mfa }
            $sync = Get-CippSyncItem -TenantFilter $tenant.Domain; if ($sync) { $items += $sync }
        }
    }

    # Licenses from the IT Glue Microsoft Licenses flexible asset. The synced trait names vary, so
    # probe common keys per asset.
    $typeId = Get-ItgFlexAssetTypeId -NameLike @('*Microsoft Licenses*','*Office 365 Licenses*','*Microsoft 365 Licenses*')
    $assets = @()
    if ($typeId) {
        try {
            $res = Get-ITGlueFlexibleAssets -filter_flexible_asset_type_id $typeId -filter_organization_id $OrgId
            if ($res.data) { $assets = $res.data }
        } catch { Write-Status "  M365 license asset query failed: $($_.Exception.Message)" Warning }
    }
    foreach ($a in $assets) {
        $traits = $a.attributes.traits
        $skuName = $null; $active = $null; $consumed = $null
        if ($traits) {
            foreach ($p in $traits.PSObject.Properties) {
                # IT Glue's synced Microsoft Licenses asset stores counts as decimals (e.g. "70.0"),
                # so accept an optional fractional part and cast to int. Skip 'unused' so it is not
                # mistaken for 'consumed' via the 'used' token.
                switch -Regex ($p.Name) {
                    '^unused$' { continue }
                    'name|sku|product|subscription' { if (-not $skuName) { $skuName = "$($p.Value)" } }
                    'active|purchased|total|prepaid' { if ($null -eq $active   -and "$($p.Value)" -match '^\d+(\.\d+)?$') { $active   = [int][double]$p.Value } }
                    'consumed|assigned|used'         { if ($null -eq $consumed -and "$($p.Value)" -match '^\d+(\.\d+)?$') { $consumed = [int][double]$p.Value } }
                }
            }
        }
        # Hide free/unlimited/empty SKUs: Microsoft reports those with a total (denominator) of 0
        # (e.g. Teams Exploratory, Exchange Online Essentials) or a large round sentinel like
        # 1000 / 10000 / 100000 / 1000000 (e.g. Power Automate Free, Windows Store For Business).
        # Real paid licenses have a specific seat count under 1000. Skip when we parsed such a total.
        if ($null -ne $active -and ([int]$active -le 0 -or [int]$active -ge 1000)) { continue }

        $label = if ($skuName) { $skuName } else { $a.attributes.name }
        if ($null -ne $consumed -and $null -ne $active) { $label += " - $consumed/$active" }
        # Over-provisioned (consumed > purchased, e.g. 4/3) -> red alert pill; otherwise muted.
        $over = ($null -ne $consumed -and $null -ne $active -and [int]$consumed -gt [int]$active)
        $bg   = if ($over) { '#DC3545' } else { '' }
        $items += New-CardItem -Label $label -Link "$LinkBase/assets/$($a.id)" -Brand 'microsoft.com' -Muted -Bg $bg
    }
    return New-Card -Title 'M365' -Items $items
}

# ---- Users (IT Glue contacts + ConnectWise contacts) ----
function Get-ItgContactCount {
    param([string]$OrgId)
    try { return [int](Get-ITGlueContacts -organization_id $OrgId -page_size 1).meta.'total-count' } catch { return $null }
}

function Resolve-CwmCompany {
    # Match the IT Glue org to a non-deleted CWM company by normalized name; cache in service-map.
    param([pscustomobject]$Entry, [string]$OrgName)
    if (-not $script:CwmConnected) { return $null }
    if ($Entry._unmanaged -contains 'connectwise') { return $null }
    $cached = $Entry.PSObject.Properties['connectwise']
    if ($cached -and $cached.Value) { return $cached.Value }   # avoid re-fetching all companies

    $cands = @()
    try {
        foreach ($c in (Get-CWMCompany -condition "deletedFlag=false" -all)) {
            $cands += [pscustomobject]@{ Id = "$($c.id)"; Name = $c.name }
        }
    } catch { Write-Status "  ConnectWise company query failed: $($_.Exception.Message)" Warning; return $null }
    return Resolve-OrgServiceId -Entry $Entry -VendorKey 'connectwise' -VendorLabel 'ConnectWise' -OrgName $OrgName -Candidates $cands
}

function Get-CwmActiveContactCount {
    param([string]$CompanyId)
    try { return @(Get-CWMContact -condition "company/id=$CompanyId and inactiveFlag=false" -all).Count } catch { return $null }
}

function Get-CwmMemberQty {
    # Quantity of part CBT-PF-MEMBER on the company's "IT Services Agreement - PeopleFirst Support"
    # (agreement type "IT Services Agreement"). Returns $null if the company has no such agreement
    # (so the Users card skips the pill); 0 if the agreement exists but carries no member seats.
    param([string]$CompanyId)
    try {
        $agr = Get-CWMAgreement -condition "company/id=$CompanyId and name='IT Services Agreement - PeopleFirst Support' and type/name='IT Services Agreement'" |
            Select-Object -First 1
        if (-not $agr) { return $null }
        $adds = Get-CWMAgreementAddition -AgreementID $agr.id -condition "product/identifier='CBT-PF-MEMBER'" -all
        return [int](($adds | Measure-Object -Property quantity -Sum).Sum)
    } catch { return $null }
}

function Get-DashUsers {
    param([pscustomobject]$Entry, [string]$OrgId, [string]$OrgName, [string]$LinkBase, [hashtable]$Creds)
    $items = @()
    # ConnectWise first (on top of IT Glue). CW Manage can't deep-link a company's Contacts tab (that
    # state is an opaque LZMA-compressed URL fragment), so the contacts pill links to the company's
    # BILLING CONTACT record (falls back to the company record).
    if ($script:CwmConnected) {
        $co = Resolve-CwmCompany -Entry $Entry -OrgName $OrgName
        if ($co) {
            # Agreement on top: PeopleFirst member seats (CBT-PF-MEMBER qty on the PeopleFirst Support
            # agreement). Skipped entirely if the company has no such agreement.
            $memQty = Get-CwmMemberQty -CompanyId $co.Id
            if ($null -ne $memQty) { $items += New-CardItem -Label "Agreement: $memQty" -Brand 'connectwise.com' }

            # Then CWM contacts (links to the company's billing contact / company record).
            $cw = Get-CwmActiveContactCount -CompanyId $co.Id
            $base = "https://$($Creds.CWM.ConnectionInfo.Server)/v4_6_release/services/system_io/router/openrecord.rails"
            $billId = $null
            try { $billId = (Get-CWMCompany -id $co.Id).billingContact.id } catch {}
            $cwHref = if ($billId) { "$base`?recordType=ContactFV&recid=$billId" } else { "$base`?recordType=CompanyFV&recid=$($co.Id)" }
            $cwLabel = if ($null -ne $cw) { "Contacts: $cw" } else { "Contacts: n/a" }
            $items += New-CardItem -Label $cwLabel -Link $cwHref -Brand 'connectwise.com'
        }
    }
    # IT Glue contacts (no active/inactive concept in IT Glue -> total contacts for the org).
    $itg = Get-ItgContactCount -OrgId $OrgId
    $itgLabel = if ($null -ne $itg) { "Contacts: $itg" } else { "Contacts: n/a" }
    $items += New-CardItem -Label $itgLabel -Link "$LinkBase/contacts" -Brand 'itglue.com'

    return New-Card -Title 'Users' -Items $items
}

# ---- Workstations / Servers (IT Glue configurations, by type + Active status) ----
function Get-ItgConfigTypeId {
    param([string]$Name)
    foreach ($t in (Get-ITGlueConfigurationTypes).data) { if ($t.attributes.name -ieq $Name) { return $t.id } }
    return $null
}
function Get-ItgConfigStatusId {
    param([string]$Name)
    foreach ($s in (Get-ITGlueConfigurationStatuses).data) { if ($s.attributes.name -ieq $Name) { return $s.id } }
    return $null
}
function Resolve-AutomateClient {
    # Match the IT Glue org to a ConnectWise Automate client and cache its client Id (the Automate
    # companyId used in the computers URL) under 'automate'. We query clients?condition=Name contains
    # '<token>' (the bulk clients list intermittently returns one aggregated row), then normalized-match.
    # A miss is NOT cached (re-resolve next run) so a flaky response can't poison the cache.
    param([pscustomobject]$Entry, [string]$OrgName)
    if (-not $script:AutomateConnected) { return $null }
    $cached = $Entry.PSObject.Properties['automate']
    if ($cached -and $cached.Value) { return $cached.Value.Id }

    $target = ConvertTo-NormalizedName $OrgName
    $token  = ($OrgName -split '[ ,]+' | Where-Object { $_.Length -ge 4 } | Select-Object -First 1)
    if (-not $token) { $token = $OrgName }
    $cands = @()
    try {
        $cond  = "Name contains '$($token -replace "'","''")'"
        $cands = @(Invoke-RestMethod -Uri "$($script:AutomateUrl)/cwa/api/v1/clients?condition=$([uri]::EscapeDataString($cond))&pagesize=200" -Headers $script:AutomateHeaders)
    } catch { Write-Status "  Automate client query failed: $($_.Exception.Message)" Warning }
    $m = $cands | Where-Object { (ConvertTo-NormalizedName $_.Name) -eq $target } | Select-Object -First 1
    if (-not $m) { return $null }
    $val = [pscustomobject]@{ Id = "$($m.Id)"; Name = $m.Name }
    $Entry | Add-Member -NotePropertyName 'automate' -NotePropertyValue $val -Force
    $script:ServiceMapDirty = $true
    return $val.Id
}

function Get-ConfigCountPill {
    # A pill card-item "<Label>: <count>". Links to the client's ConnectWise Automate computers page
    # when -Link is supplied (Automate icon via -Brand); otherwise to the IT Glue configs list
    # filtered by type. (Future: append a stale count, e.g. devices not seen in Automate for 90+ days.)
    param([string]$Label, [string]$TypeName, [string]$OrgId, [string]$LinkBase, [string]$TypeId,
          [string]$StatusId, [string]$Link = '', [string]$Brand = '')
    if (-not $TypeId -or -not $StatusId) { return New-CardItem -Label "${Label}: n/a" }
    $count = $null
    try {
        $count = [int](Get-ITGlueConfigurations -filter_organization_id $OrgId -filter_configuration_type_id $TypeId `
            -filter_configuration_status_id $StatusId -page_size 1).meta.'total-count'
    } catch { Write-Status "  IT Glue $Label query failed: $($_.Exception.Message)" Warning }
    if (-not $Link) {
        # Fallback: IT Glue UI filter deep-link (#partial=...&filters=[Type:<name>]).
        $link = "$LinkBase/configurations#partial=&sortBy=name:asc&filters=%5BType:$([Uri]::EscapeDataString($TypeName))%5D"
    } else { $link = $Link }
    return New-CardItem -Label "${Label}: $([int]$count)" -Link $link -Brand $Brand -Muted
}

function Get-DashDevices {
    # 'Devices' card: Workstations + Servers pills (active IT Glue configs). Pills link to the client's
    # ConnectWise Automate computers page when the org matches an Automate client; else to IT Glue.
    param([pscustomobject]$Entry, [string]$OrgId, [string]$OrgName, [string]$LinkBase, [string]$WsTypeId, [string]$SvTypeId, [string]$StatusId)
    $autoLink = ''; $brand = ''
    $autoId = Resolve-AutomateClient -Entry $Entry -OrgName $OrgName
    if ($autoId) { $autoLink = "$($script:AutomateUrl)/automate/browse/companies/computers?companyId=$autoId"; $brand = 'automate' }
    $items = @(
        (Get-ConfigCountPill -Label 'Workstations' -TypeName 'Managed Workstation' -OrgId $OrgId -LinkBase $LinkBase -TypeId $WsTypeId -StatusId $StatusId -Link $autoLink -Brand $brand),
        (Get-ConfigCountPill -Label 'Servers'      -TypeName 'Managed Server'      -OrgId $OrgId -LinkBase $LinkBase -TypeId $SvTypeId -StatusId $StatusId -Link $autoLink -Brand $brand)
    )
    return New-Card -Title 'Devices' -Items $items
}

# ==============================================================================
# Rendering
# ==============================================================================
$script:CardColor = '#051554'   # navy card background

function ConvertTo-CardItemHtml {
    param([pscustomobject]$Item)
    $brand = Resolve-Brand $Item.Brand
    if ($Item.Bg) {
        # Explicit colour override (e.g. red alert) wins over brand/muted.
        $colored = $true; $bg = $Item.Bg; $border = 'none'
    } else {
        $colored = ($brand -and $brand.Color -and -not $Item.Muted)
        $bg      = if ($colored) { $brand.Color } else { 'rgba(255,255,255,0.12)' }
        $border  = if ($colored) { 'none' } else { '1px solid rgba(255,255,255,0.2)' }
    }

    if ($Item.Kind -eq 'number') {
        $inner = "<span style=`"font-size:30px;font-weight:800;line-height:1;`">$(Get-HtmlEncoded $Item.Label)</span>" +
                 "<span style=`"font-size:13px;font-weight:600;opacity:0.85;margin-top:4px;`">$(Get-HtmlEncoded $Item.Sub)</span>"
        $style = "display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;" +
                 "padding:16px 14px;border-radius:10px;text-decoration:none;word-break:break-word;background:$bg;color:#fff;border:$border;"
    }
    else {
        $icon = ''
        if ($brand -and $brand.Logo) {
            $icon = "<span style=`"display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;border-radius:7px;background:#fff;flex:none;`">" +
                    "<img src=`"$($brand.Logo)`" alt=`"`" width=`"18`" height=`"18`" style=`"display:block;width:18px;height:18px;object-fit:contain;`"></span>"
        }
        $weight = if ($colored) { '700' } else { '600' }
        $inner  = "$icon<span>$(Get-HtmlEncoded $Item.Label)</span>"
        $style  = "display:flex;align-items:center;justify-content:center;gap:10px;text-align:center;" +
                  "padding:12px 14px;border-radius:10px;font-size:14px;font-weight:$weight;line-height:1.3;" +
                  "text-decoration:none;word-break:break-word;background:$bg;color:#fff;border:$border;"
    }

    if ($Item.Link) { return "<a href=`"$($Item.Link)`" target=`"_blank`" rel=`"noopener`" style=`"$style`">$inner</a>" }
    return "<div style=`"$style`">$inner</div>"
}

function ConvertTo-CardHtml {
    param([pscustomobject]$Card)
    $items = @($Card.Items)
    if ($items.Count -eq 0) { $items = @(New-CardItem -Label 'None') }   # empty category -> muted 'None'
    $body = ($items | ForEach-Object { ConvertTo-CardItemHtml $_ }) -join ''
    $h3 = "<h3 style=`"margin:0;color:#fff;font-size:13px;font-weight:700;letter-spacing:0.08em;text-transform:uppercase;text-align:center;opacity:0.82;`">$(Get-HtmlEncoded $Card.Title)</h3>"
    $inner = "<div style=`"flex:1;display:flex;flex-direction:column;justify-content:center;gap:10px;`">$body</div>"
    return "<div style=`"background:$($script:CardColor);border-radius:16px;padding:22px 20px;display:flex;flex-direction:column;gap:14px;aspect-ratio:1 / 1;box-shadow:0 2px 6px rgba(5,21,84,0.12);`">$h3$inner</div>"
}

function New-DashboardHtml {
    # Responsive card grid - this is the exact markup a future -Publish would store in IT Glue
    # (all card styling is inline, which IT Glue preserves; only the page chrome uses <style>).
    param([pscustomobject[]]$Cards)
    $cards = ($Cards | ForEach-Object { ConvertTo-CardHtml -Card $_ }) -join ''
    return "<div style=`"display:grid;grid-template-columns:repeat(auto-fit, minmax(220px, 1fr));gap:18px;`">$cards</div>"
}

function New-PreviewDocument {
    # Local browser preview only: light page chrome + header around the card grid.
    param([string]$DashboardHtml, [string]$OrgName)
    $enc = [System.Net.WebUtility]::HtmlEncode($OrgName)
    $gen = Get-Date -Format 'yyyy-MM-dd HH:mm'
    @"
<!DOCTYPE html>
<html lang="en"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$enc - PeopleFirst Services</title>
<style>
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  body { background: #eef0f5; font-family: -apple-system, BlinkMacSystemFont, "Helvetica Neue", Arial, sans-serif; color: #1c2333; -webkit-font-smoothing: antialiased; }
</style>
</head>
<body>
<div style="max-width: 1180px; margin: 0 auto; padding: 36px 28px 56px;">
  <header style="margin-bottom: 28px;">
    <h1 style="margin: 0; font-size: 26px; font-weight: 800; letter-spacing: -0.01em; color: #051554;">$enc</h1>
    <p style="margin: 6px 0 0; font-size: 15px; font-weight: 600; color: #5b6478;">PeopleFirst Services Overview</p>
    <p style="margin: 4px 0 0; font-size: 12px; color: #97a0b5;">Local preview &middot; generated $gen</p>
  </header>
$DashboardHtml
</div>
</body></html>
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

# Credentials
$creds = Initialize-Credentials -Reset:$Reset
Add-ITGlueBaseURI -base_uri $creds.ITGlue.BaseUri
Add-ITGlueAPIKey  -Api_Key  $creds.ITGlue.ApiKey

# Duo needs no module: we call the Accounts/Admin API directly (see Invoke-DuoApi). This avoids the
# DuoSecurity module, whose manifest requires PowerShell 7.0 (incompatible with our 5.1 requirement).

# ConnectWise Manage (only if seeded) - used for the ConnectWise contacts count in the Users tile.
$script:CwmConnected = $false
if ($creds.CWM.Configured) {
    try {
        if (-not (Get-Module -ListAvailable -Name ConnectWiseManageAPI)) {
            Write-Status "Installing ConnectWiseManageAPI module..." Warning
            Install-Module -Name ConnectWiseManageAPI -Scope CurrentUser -Force -AllowClobber
        }
        Import-Module ConnectWiseManageAPI -Force
        $cwmCi = $creds.CWM.ConnectionInfo
        Connect-CWM @cwmCi -ErrorAction Stop
        $script:CwmConnected = $true
    } catch { Write-Status "ConnectWise Manage connect failed: $($_.Exception.Message)" Warning }
}

# CIPP (M365 data) - one client-credentials token, reused for every client. Per-org tenant lookup
# is done in Resolve-CippTenant via ListTenants?TenantFilter=<domain>.
$script:CippConnected = $false
if ($creds.CIPP.Configured) {
    try {
        $tok = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$($creds.CIPP.TenantId)/oauth2/v2.0/token" -Method Post `
            -ContentType 'application/x-www-form-urlencoded' -Body @{
                client_id = $creds.CIPP.ClientId; client_secret = $creds.CIPP.ClientSecret
                grant_type = 'client_credentials'; scope = $creds.CIPP.Scope
            }
        $script:CippToken   = $tok.access_token
        $script:CippApiUrl  = $creds.CIPP.ApiUrl.TrimEnd('/')
        $script:CippConnected = $true
    } catch { Write-Status "CIPP connect failed: $($_.Exception.Message)" Warning }
}

# ConnectWise Automate - token + client list (used to link the Devices pills to the client's
# computers). The ClientId header is required on every Automate API call.
$script:AutomateConnected = $false
if ($creds.Automate.Configured) {
    try {
        $script:AutomateUrl = $creds.Automate.Url.TrimEnd('/')
        $aTok = (Invoke-RestMethod -Uri "$script:AutomateUrl/cwa/api/v1/apitoken" -Method Post `
            -Headers @{ ClientId = $creds.Automate.ClientId } -ContentType 'application/json' `
            -Body (@{ UserName = $creds.Automate.Username; Password = $creds.Automate.Password } | ConvertTo-Json)).AccessToken
        $script:AutomateHeaders = @{ Authorization = "Bearer $aTok"; ClientId = $creds.Automate.ClientId }
        $script:AutomateConnected = $true
    } catch { Write-Status "Automate connect failed: $($_.Exception.Message)" Warning }
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

# Resolve IT Glue configuration type/status ids once (by name) for the Workstation/Server tiles.
$wsTypeId       = Get-ItgConfigTypeId   -Name 'Managed Workstation'
$svTypeId       = Get-ItgConfigTypeId   -Name 'Managed Server'
$activeStatusId = Get-ItgConfigStatusId -Name 'Active'

$tileUsers   = Get-DashUsers   -Entry $entry -OrgId $orgId -OrgName $orgName -LinkBase $linkBase -Creds $creds
$tileDevices = Get-DashDevices -Entry $entry -OrgId $orgId -OrgName $orgName -LinkBase $linkBase -WsTypeId $wsTypeId -SvTypeId $svTypeId -StatusId $activeStatusId
$tileAv      = Get-DashAntivirus -Entry $entry -OrgName $orgName -Creds $creds
$tileMfa    = Get-DashMfa       -Entry $entry -OrgName $orgName -Creds $creds
$tileBackup = Get-DashBackup    -Entry $entry -OrgName $orgName -Creds $creds
$tileM365   = Get-DashM365      -Entry $entry -OrgId $orgId -LinkBase $linkBase
$tileDomains = Get-DashDomains  -OrgId $orgId -LinkBase $linkBase

# Persist any new resolutions
if ($script:ServiceMapDirty) {
    $entry.resolvedOn = (Get-Date).ToString('yyyy-MM-dd')
    Save-ServiceMap -Map $script:ServiceMap
    Write-Status "Service-map cache updated: $ServiceMapPath" Detail
}

# -- Render --
$cards = @($tileUsers, $tileDevices, $tileAv, $tileMfa, $tileBackup, $tileM365, $tileDomains)
$dashboardHtml = New-DashboardHtml -Cards $cards

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
