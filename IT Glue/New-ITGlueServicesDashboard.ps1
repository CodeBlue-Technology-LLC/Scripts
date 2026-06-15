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
    $result = [pscustomobject]@{ Servers = 0; Workstations = 0; VMs = 0 }
    try {
        $vms = Get-VspcPaged -BaseURL $BaseURL -Token $Token -Path '/api/v3/protectedWorkloads/virtualMachines' |
            Where-Object { $_.organizationUid -eq $CompanyUid }
        $result.VMs = @($vms).Count
    } catch { }
    try {
        $agents = Get-VspcPaged -BaseURL $BaseURL -Token $Token -Path '/api/v3/protectedWorkloads/computersManagedByConsole' |
            Where-Object { $_.organizationUid -eq $CompanyUid }
        $result.Servers      = @($agents | Where-Object { $_.operationMode -eq 'Server' }).Count
        $result.Workstations = @($agents | Where-Object { $_.operationMode -eq 'Workstation' }).Count
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

function New-CategoryTile {
    # One tile per category (Antivirus, Backup, MFA...) listing EVERY active product as a
    # "Product: value" line. Green when any product is present; red 'None' when the client has none.
    # Each line is either a plain string, or an object @{ Text; Link } to render the line as a link.
    param([string]$Title, [object[]]$Lines)
    if (-not $Lines -or $Lines.Count -eq 0) { return New-Tile -Title $Title -Shading 'danger' -Content 'None' }
    $items = foreach ($l in $Lines) {
        if ($l -is [string]) {
            [pscustomobject]@{ Shading = 'success'; AlertText = (Get-HtmlEncoded $l) }
        } else {
            $txt  = Get-HtmlEncoded $l.Text
            $html = if ($l.Link) { "<a href=`"$($l.Link)`">$txt</a>" } else { $txt }
            [pscustomobject]@{ Shading = 'success'; AlertText = $html }
        }
    }
    return New-Tile -Title $Title -Shading 'success' -IsInfo -InfoItems $items
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
        if ($site) { $lines += "SentinelOne: $([int]$site.ActiveLicenses)" }
    }

    if ($Creds.Bitdefender.Configured) {
        $companies = @()
        try { $companies = Get-BdCompanies -BaseURL $Creds.Bitdefender.URL -ApiKey $Creds.Bitdefender.ApiKey }
        catch { Write-Status "  Bitdefender query failed: $($_.Exception.Message)" Warning }
        $company = Resolve-OrgServiceId -Entry $Entry -VendorKey 'bitdefender' -VendorLabel 'Bitdefender' -OrgName $OrgName -Candidates $companies
        if ($company) {
            $count = Get-BdEndpointCount -BaseURL $Creds.Bitdefender.URL -ApiKey $Creds.Bitdefender.ApiKey -CompanyId $company.Id
            $lines += "Bitdefender: $([int]$count)"
        }
    }

    return New-CategoryTile -Title 'Antivirus' -Lines $lines
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
                $lines += [pscustomobject]@{ Text = "Duo: $($c.Active) active, $($c.Bypass) bypass"; Link = $adminUrl }
            }
            catch { Write-Status "  Duo user count failed for '$($acct.Name)': $($_.Exception.Message)" Warning; $lines += "Duo: n/a" }
        }
    }
    return New-CategoryTile -Title 'MFA' -Lines $lines
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
                $total = [int]$b.Servers + [int]$b.Workstations + [int]$b.VMs
                if ($total -gt 0) {
                    $lines += "Veeam: $($b.Servers) srv, $($b.Workstations) wks, $($b.VMs) VM"
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
                    $lines += "Cove: $tb TB ($servers srv, $workstations wks)"
                }
                $m365 = @($devices | Where-Object { "$($_.Physicality)" -eq 'Undefined' })
                if ($m365.Count -gt 0) {
                    $tb365 = [Math]::Round((($m365 | Measure-Object -Property UsedGB -Sum).Sum) / 1024, 2)
                    $tenants = if ($m365.Count -eq 1) { '1 tenant' } else { "$($m365.Count) tenants" }
                    $lines += "Cove 365: $tb365 TB ($tenants)"
                }
            }
        } catch { Write-Status "  Cove query failed: $($_.Exception.Message)" Warning }
    }

    return New-CategoryTile -Title 'Backup' -Lines $lines
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

        $label = if ($skuName) { Get-HtmlEncoded $skuName } else { Get-HtmlEncoded $a.attributes.name }
        if ($null -ne $consumed -and $null -ne $active) { $label += " - $consumed/$active" }
        $href = "$LinkBase/assets/$($a.id)"
        $items += [pscustomobject]@{ Shading = 'success'; AlertText = "<a href=`"$href`">$label</a>" }
    }
    return New-Tile -Title 'M365 Licenses' -Shading 'info' -IsInfo -InfoItems $items
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

function Get-DashUsers {
    param([pscustomobject]$Entry, [string]$OrgId, [string]$OrgName, [string]$LinkBase, [hashtable]$Creds)
    $items = @()
    # IT Glue: contacts have no active/inactive concept, so this is the org's total contact count.
    $itg = Get-ItgContactCount -OrgId $OrgId
    $itgLabel = if ($null -ne $itg) { "IT Glue: $itg" } else { "IT Glue: n/a" }
    $items += [pscustomobject]@{ Shading = 'info'; AlertText = "<a href=`"$LinkBase/contacts`">$itgLabel</a>" }
    # ConnectWise: active contacts for the matched company. CW Manage can't deep-link a company's
    # Contacts tab (that state is an opaque LZMA-compressed URL fragment), so we link to the
    # company's BILLING CONTACT record (lands directly in a contact); fall back to the company record.
    if ($script:CwmConnected) {
        $co = Resolve-CwmCompany -Entry $Entry -OrgName $OrgName
        if ($co) {
            $cw = Get-CwmActiveContactCount -CompanyId $co.Id
            $base = "https://$($Creds.CWM.ConnectionInfo.Server)/v4_6_release/services/system_io/router/openrecord.rails"
            $billId = $null
            try { $billId = (Get-CWMCompany -id $co.Id).billingContact.id } catch {}
            $cwHref = if ($billId) { "$base`?recordType=ContactFV&recid=$billId" } else { "$base`?recordType=CompanyFV&recid=$($co.Id)" }
            $cwLabel = if ($null -ne $cw) { "ConnectWise: $cw" } else { "ConnectWise: n/a" }
            $items += [pscustomobject]@{ Shading = 'info'; AlertText = "<a href=`"$cwHref`">$cwLabel</a>" }
        }
    }
    return New-Tile -Title 'Users' -Shading 'info' -IsInfo -InfoItems $items
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
function Get-DashConfigTile {
    param([string]$Title, [string]$OrgId, [string]$LinkBase, [string]$TypeId, [string]$TypeName, [string]$StatusId)
    if (-not $TypeId -or -not $StatusId) { return New-Tile -Title $Title -Shading 'warning' -Content 'n/a' }
    $count = $null
    try {
        $count = [int](Get-ITGlueConfigurations -filter_organization_id $OrgId -filter_configuration_type_id $TypeId `
            -filter_configuration_status_id $StatusId -page_size 1).meta.'total-count'
    } catch { Write-Status "  IT Glue $Title query failed: $($_.Exception.Message)" Warning }
    $shading = if ($count -gt 0) { 'success' } else { 'danger' }
    # IT Glue UI filter deep-link (matches the app's #partial=...&filters=[Type:<name>] hash format).
    $typeEnc = [Uri]::EscapeDataString($TypeName)
    $link = "$LinkBase/configurations#partial=&sortBy=name:asc&filters=%5BType:$typeEnc%5D"
    return New-Tile -Title $Title -Shading $shading -Content "$([int]$count)" -Detail 'Active' -Link $link
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
    # Singles render at col-sm-3 (wrap to multiple rows); Infos (lists) render at col-sm-6.
    param([pscustomobject[]]$Singles, [pscustomobject[]]$Infos)
    $s = ($Singles | ForEach-Object { ConvertTo-PanelHtml -Tile $_ -Size 3 }) -join ''
    $i = ($Infos   | ForEach-Object { ConvertTo-PanelHtml -Tile $_ -Size 6 }) -join ''
    return "<div class=`"row`">$s</div><div class=`"row`">$i</div>"
}

function New-PreviewDocument {
    # Wrap the dashboard markup in a Bootstrap-3 CDN page for local browser preview only.
    param([string]$DashboardHtml, [string]$OrgName)
    @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>$([System.Net.WebUtility]::HtmlEncode($OrgName)) - PeopleFirst Services</title>
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap-theme.min.css">
</head>
<body>
<div class="container-fluid" style="margin-top:20px">
<h1>$([System.Net.WebUtility]::HtmlEncode($OrgName)) - PeopleFirst Services</h1>
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
$hasM365 = $false

# Resolve IT Glue configuration type/status ids once (by name) for the Workstation/Server tiles.
$wsTypeId       = Get-ItgConfigTypeId   -Name 'Managed Workstation'
$svTypeId       = Get-ItgConfigTypeId   -Name 'Managed Server'
$activeStatusId = Get-ItgConfigStatusId -Name 'Active'

$tileUsers  = Get-DashUsers     -Entry $entry -OrgId $orgId -OrgName $orgName -LinkBase $linkBase -Creds $creds
$tileWs     = Get-DashConfigTile -Title 'Workstations' -OrgId $orgId -LinkBase $linkBase -TypeId $wsTypeId -TypeName 'Managed Workstation' -StatusId $activeStatusId
$tileSv     = Get-DashConfigTile -Title 'Servers'      -OrgId $orgId -LinkBase $linkBase -TypeId $svTypeId -TypeName 'Managed Server'      -StatusId $activeStatusId
$tileAv     = Get-DashAntivirus -Entry $entry -OrgName $orgName -Creds $creds
$tileMfa    = Get-DashMfa       -Entry $entry -OrgName $orgName -Creds $creds
$tileBackup = Get-DashBackup    -Entry $entry -OrgName $orgName -Creds $creds
$tileM365   = Get-DashM365      -OrgId $orgId -LinkBase $linkBase -HasM365Out ([ref]$hasM365)
$tileDomains = Get-DashDomains  -OrgId $orgId -LinkBase $linkBase

# Persist any new resolutions
if ($script:ServiceMapDirty) {
    $entry.resolvedOn = (Get-Date).ToString('yyyy-MM-dd')
    Save-ServiceMap -Map $script:ServiceMap
    Write-Status "Service-map cache updated: $ServiceMapPath" Detail
}

# -- Render --
$singles = @($tileUsers, $tileWs, $tileSv, $tileAv, $tileMfa, $tileBackup)
$infos   = @($tileM365, $tileDomains)
$dashboardHtml = New-DashboardHtml -Singles $singles -Infos $infos

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
