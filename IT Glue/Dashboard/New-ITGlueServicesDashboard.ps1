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
      - Duo           : # active users protected (or "No Duo"); expands to a bypass/disabled/
                        locked-out/not-enrolled/stale breakdown with the at-risk usernames.
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
$BrandCachePath    = Join-Path $ConfigDir 'brand-cache.json'
$OutputDir        = Join-Path $ScriptDir 'Output'

# A device is "stale" if it hasn't checked in for this many days. Used dashboard-wide wherever staleness
# is measured (currently the Automate Workstations pill); a vendor's own built-in staleness measure, if
# one exists, takes precedence over this fallback.
$script:StaleDays  = 90

# ==============================================================================
# Small helpers
# ==============================================================================
function Write-Status {
    param([string]$Message, [ValidateSet('Info','Success','Warning','Error','Detail')][string]$Level = 'Info')
    $color = @{ Info='Cyan'; Success='Green'; Warning='Yellow'; Error='Red'; Detail='Gray' }[$Level]
    Write-Host $Message -ForegroundColor $color
}

function ConvertTo-NormalizedName {
    # Lowercase, strip punctuation + common legal suffixes, collapse whitespace. Used only to MATCH a
    # vendor's candidate names against the canonical (ConnectWise Manage) org name - it never affects a
    # displayed name. Kept deliberately simple: vendors are expected to carry the same name as CWM, so
    # this just absorbs case/punctuation/suffix formatting. Where a specific vendor mangles a name
    # (e.g. disallows a character), add a targeted rule here as that case comes up.
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

    if (-not $creds.Cloudflare) {
        if ((Read-Host "Configure Cloudflare (color domains in your self-service portal)? (y/N)") -match '^[Yy]') {
            # An API token with Zone:Read over the portal account(s). We list its zones and colour any
            # matching client domain. Create at https://dash.cloudflare.com/profile/api-tokens.
            $cfToken = Read-Host "  Cloudflare API Token (Zone:Read)"
            $creds.Cloudflare = @{ Configured = $true; ApiToken = $cfToken }
        } else { $creds.Cloudflare = @{ Configured = $false } }
        $changed = $true
    }

    if (-not $creds.UniFi) {
        if ((Read-Host "Configure UniFi (Site Manager cloud API)? (y/N)") -match '^[Yy]') {
            # Site Manager API key (X-API-KEY), created at unifi.ui.com (Settings > API). Read-only
            # is sufficient - we only list hosts/sites to match the org and link to its cloud console.
            $creds.UniFi = @{ Configured = $true; ApiKey = Read-Host "  UniFi Site Manager API Key (X-API-KEY)" }
        } else { $creds.UniFi = @{ Configured = $false } }
        $changed = $true
    }

    if (-not $creds.MySonicWall) {
        if ((Read-Host "Configure SonicWall (MySonicWall - Capture Client / CAS / firewalls)? (y/N)") -match '^[Yy]') {
            # MySonicWall API key (X-api-key header), from My Workspace > User Groups > User List >
            # Generate API Key. We use get-cloud-tenants (Capture Client/CAS counts) + get-firewalls
            # (the registered firewall fleet) + serviceInfo (per-firewall license expiry).
            $swKey = Read-Host "  MySonicWall API Key"
            # The product group that holds your registered firewalls; blank = auto-detect (the group
            # with the most products, i.e. your own MySonicWall account).
            $swGrp = Read-Host "  Firewall product group ID (blank = auto-detect)"
            $creds.MySonicWall = @{ Configured = $true; ApiKey = $swKey; FirewallGroupId = $swGrp.Trim() }
        } else { $creds.MySonicWall = @{ Configured = $false } }
        $changed = $true
    }

    if (-not $creds.WatchGuard) {
        if ((Read-Host "Configure WatchGuard (Cloud - firewalls)? (y/N)") -match '^[Yy]') {
            # WatchGuard Cloud API (USA region): read-only API Access client (AccessID + password) for the
            # OAuth token, plus the separate API Key sent in the WatchGuard-API-Key header on each call.
            # AccountId is the service-provider account (e.g. ACC-1326807) from WatchGuard Cloud My Account.
            $creds.WatchGuard = @{
                Configured = $true
                AccessId   = Read-Host "  WatchGuard API Access ID"
                Password   = Read-Host "  WatchGuard API Access password"
                ApiKey     = Read-Host "  WatchGuard API Key"
                AccountId  = Read-Host "  WatchGuard Account ID (e.g. ACC-XXXXXXX)"
            }
        } else { $creds.WatchGuard = @{ Configured = $false } }
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
        [object[]]$Candidates = @(),
        # When set, a no-confident-match is silently recorded as unmanaged instead of prompting.
        # Used by the AV tile so a client confidently matched on one platform isn't prompted for another.
        [switch]$NoPromptUnmanaged
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

    # 3) no confident match. Caller may opt to record unmanaged silently (e.g. already matched on a
    #    sibling AV platform) instead of interactively prompting.
    if ($NoPromptUnmanaged) { Add-Unmanaged -Entry $Entry -VendorKey $VendorKey; return $null }

    # otherwise: interactive pick-list of candidates (if any), plus none/manual
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

function Test-VendorMatched {
    # True if the org is confidently resolvable for this vendor WITHOUT prompting: either a cached
    # resolution (not unmanaged) or exactly one normalized-name match in the candidate list. Mirrors
    # the confident-match paths of Resolve-OrgServiceId but never mutates state or prompts.
    param([pscustomobject]$Entry, [string]$VendorKey, [string]$OrgName, [object[]]$Candidates = @())
    if ($Entry._unmanaged -contains $VendorKey) { return $false }
    $cached = $Entry.PSObject.Properties[$VendorKey]
    if ($cached -and $cached.Value) { return $true }
    $target = ConvertTo-NormalizedName $OrgName
    $n = @($Candidates | Where-Object { (ConvertTo-NormalizedName $_.Name) -eq $target }).Count
    return ($n -eq 1)
}

function Add-Resolution {
    param([pscustomobject]$Entry, [string]$VendorKey, [object]$Candidate)
    $val = [pscustomobject]@{ Id = "$($Candidate.Id)"; Name = $Candidate.Name }
    # Preserve every extra candidate property (e.g. Duo's ApiHostname, CIPP's Domain, SentinelOne's
    # Instance/ConsoleURL) so the cached resolution can drive deep links and per-account API calls.
    foreach ($p in $Candidate.PSObject.Properties) {
        if ($p.Name -in @('Id','Name')) { continue }
        $val | Add-Member -NotePropertyName $p.Name -NotePropertyValue $p.Value -Force
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
                ConsoleURL     = $BaseURL          # instance console host (EDR vs MDR), for the deep link
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
    try {
        if ($Method.ToUpper() -eq 'GET') {
            if ($canonParams) { $uri += "?$canonParams" }
            return Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
        }
        return Invoke-RestMethod -Uri $uri -Method Post -Headers $headers `
            -ContentType 'application/x-www-form-urlencoded' -Body $canonParams
    }
    catch {
        # Surface Duo's JSON error body: in PS 5.1 a 4xx throws a WebException whose .Message is only
        # the generic "(400) Bad Request", but the response stream carries Duo's {code,message}. Read it
        # and rethrow with the detail so callers' Write-Status warnings show the real reason.
        $detail = ''
        $resp = $_.Exception.Response
        if ($resp) {
            try {
                $sr = New-Object IO.StreamReader($resp.GetResponseStream())
                $body = $sr.ReadToEnd(); $sr.Close()
                $detail = $body
                try { $j = $body | ConvertFrom-Json; if ($j.message) { $detail = "Duo $($j.code): $($j.message)" + $(if ($j.message_detail) { " ($($j.message_detail))" }) } } catch {}
            } catch {}
        }
        if ($detail) { throw "$($_.Exception.Message) - $detail" }
        throw
    }
}

function Get-DuoAccountList {
    param([hashtable]$Creds)
    $resp = Invoke-DuoApi -ApiHost $Creds.Duo.ApiHost -IKey $Creds.Duo.IntegrationKey -SKey $Creds.Duo.SecretKey `
        -Method POST -Path '/accounts/v1/account/list'
    return @($resp.response)
}

function Get-DuoUserCounts {
    # Page the child account's Admin API users (300/page) and tally by status, plus two posture gaps,
    # all from the same response (no extra calls). Duo statuses: 'active', 'bypass', 'disabled',
    # 'locked out' (note the SPACE - not an underscore). We report active/bypass/disabled/locked-out
    # counts, plus 'not enrolled' (an active user with no 2FA device) and 'stale' (active, no login in
    # 90+ days), and collect bypass + not-enrolled names for the expandable pill's panel.
    # -ApiHost overrides the parent MSP host with the child account's own api_hostname when known;
    # a child on a different Duo deployment 400s if queried against the parent host.
    param([hashtable]$Creds, [string]$AccountId, [string]$ApiHost)
    $duoHost = if ($ApiHost) { $ApiHost } else { $Creds.Duo.ApiHost }
    $active = 0; $bypass = 0; $disabled = 0; $lockedOut = 0; $notEnrolled = 0; $stale = 0; $offset = 0
    $bypassNames = @(); $notEnrolledNames = @()
    # Stale threshold: last_login is epoch SECONDS (may be null). Compute the cutoff once.
    $staleBefore = [DateTimeOffset]::UtcNow.AddDays(-90).ToUnixTimeSeconds()
    # username -> realname -> email fallback so a listed name is never blank.
    $nameOf = { param($u) $n = "$($u.username)"; if (-not $n) { $n = "$($u.realname)" }; if (-not $n) { $n = "$($u.email)" }; $n }
    do {
        $resp = Invoke-DuoApi -ApiHost $duoHost -IKey $Creds.Duo.IntegrationKey -SKey $Creds.Duo.SecretKey `
            -Method GET -Path '/admin/v1/users' -Params @{ account_id = $AccountId; limit = '300'; offset = "$offset" }
        foreach ($u in @($resp.response)) {
            switch ("$($u.status)") {
                'active'     { $active++ }
                'bypass'     { $bypass++;    $n = & $nameOf $u; if ($n) { $bypassNames += $n } }
                'disabled'   { $disabled++ }
                'locked out' { $lockedOut++ }
            }
            # Gaps are scoped to active accounts only: a disabled/locked-out user can't sign in, so a
            # missing 2FA device or stale login isn't a live risk for them.
            if ("$($u.status)" -eq 'active') {
                if ($u.is_enrolled -eq $false) { $notEnrolled++; $n = & $nameOf $u; if ($n) { $notEnrolledNames += $n } }
                $ll = $u.last_login
                if ($null -eq $ll -or [int64]$ll -lt $staleBefore) { $stale++ }
            }
        }
        $next = $resp.metadata.next_offset
        if ($null -ne $next -and "$next" -ne '') { $offset = [int]$next } else { $offset = $null }
    } while ($null -ne $offset)
    return [pscustomobject]@{
        Active = $active; Bypass = $bypass; Disabled = $disabled; LockedOut = $lockedOut
        NotEnrolled = $notEnrolled; Stale = $stale
        BypassNames = @($bypassNames | Sort-Object -Unique); NotEnrolledNames = @($notEnrolledNames | Sort-Object -Unique)
    }
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

function ConvertTo-UtcOrNull {
    # Parse an ISO timestamp to a UTC [datetime], or $null if blank/unparseable.
    param($Value)
    $s = "$Value"
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    try { return ([datetimeoffset]$s).UtcDateTime } catch { return $null }
}
function Format-RelativeAge {
    # "<n>m ago" / "<n>h ago" / "<n>d ago" from a UTC time.
    param([datetime]$Utc)
    $span = (Get-Date).ToUniversalTime() - $Utc
    if ($span.TotalMinutes -lt 60) { return "$([int]$span.TotalMinutes)m ago" }
    elseif ($span.TotalHours -lt 24) { return "$([int]$span.TotalHours)h ago" }
    else { return "$([int]$span.TotalDays)d ago" }
}
function Format-BackupSize {
    # Bytes -> "X.X TB" (>=1 TB) else "X GB". '' for non-positive/unknown.
    param([Nullable[int64]]$Bytes)
    if (-not $Bytes -or $Bytes -le 0) { return '' }
    if ($Bytes -ge 1TB) { return ('{0:0.0} TB' -f ($Bytes / 1TB)) }
    return ('{0:0} GB' -f ($Bytes / 1GB))
}

function Get-VspcCompanyBackup {
    # Per-protected-entity backup detail for the company. The ?filter= param breaks the computers
    # endpoints (they fall through to the SPA), so we page each list and filter client-side by
    # organizationUid. Returns one object per entity: Name, Type (Server/Workstation/VM), LatestBackupUtc
    # (last successful backup), TotalSizeBytes, and BackupType. We cover all three workload sources:
    #   virtualMachines              -> VM backups (size via totalRestorePointSize)
    #   computersManagedByConsole    -> agents managed by our cloud console (Cloud Connect)
    #   computersManagedByBackupServer -> agents managed by a B&R server (role: Client/Hosted/CloudConnect)
    # BackupType combines the workload with where it's managed, via the managing server's role:
    #   Client -> 'Local B&R', Hosted -> 'Hosted', CloudConnect -> 'Cloud Connect'.
    # Only VMs expose a size; agents' TotalSizeBytes is $null. De-duped by name (the B&R-agent endpoint
    # returns a row per protection group). Ordered Server -> Workstation -> VM, then by name.
    param([string]$BaseURL, [string]$Token, [string]$CompanyUid)
    $dot = [char]0x00B7
    $destMap = @{ 'Client' = 'Local B&R'; 'Hosted' = 'Hosted'; 'CloudConnect' = 'Cloud Connect' }
    # server uid -> role
    $role = @{}
    try { foreach ($s in (Get-VspcPaged -BaseURL $BaseURL -Token $Token -Path '/api/v3/infrastructure/backupServers')) { $role["$($s.instanceUid)"] = "$($s.backupServerRoleType)" } } catch { }

    $entities = @()
    # VMs (B&R-managed; role gives Local B&R / Hosted / Cloud Connect)
    try {
        foreach ($vm in @(Get-VspcPaged -BaseURL $BaseURL -Token $Token -Path '/api/v3/protectedWorkloads/virtualMachines' | Where-Object { $_.organizationUid -eq $CompanyUid })) {
            $size = $null; try { $size = [int64]$vm.totalRestorePointSize } catch { }
            $dest = $destMap["$($role["$($vm.backupServerUid)"])"]
            $bt = if ($dest) { "VM $dot $dest" } else { 'VM' }
            $entities += [pscustomobject]@{
                Name = "$($vm.name)"; Type = 'VM'; LatestBackupUtc = (ConvertTo-UtcOrNull $vm.latestRestorePointDate)
                TotalSizeBytes = $size; BackupType = $bt
            }
        }
    } catch { }
    # Agents managed by our cloud console -> Cloud Connect
    try {
        foreach ($a in @(Get-VspcPaged -BaseURL $BaseURL -Token $Token -Path '/api/v3/protectedWorkloads/computersManagedByConsole' | Where-Object { $_.organizationUid -eq $CompanyUid })) {
            $entities += [pscustomobject]@{
                Name = "$($a.name)"; Type = "$($a.operationMode)"; LatestBackupUtc = (ConvertTo-UtcOrNull $a.latestRestorePointDate)
                TotalSizeBytes = $null; BackupType = "Agent $dot Cloud Connect"
            }
        }
    } catch { }
    # Agents managed by a B&R server -> role gives Local B&R / Hosted / Cloud Connect
    try {
        foreach ($a in @(Get-VspcPaged -BaseURL $BaseURL -Token $Token -Path '/api/v3/protectedWorkloads/computersManagedByBackupServer' | Where-Object { $_.organizationUid -eq $CompanyUid })) {
            $dest = $destMap["$($role["$($a.backupServerUid)"])"]
            $bt = if ($dest) { "Agent $dot $dest" } else { 'Agent' }
            $entities += [pscustomobject]@{
                Name = "$($a.name)"; Type = "$($a.operationMode)"; LatestBackupUtc = (ConvertTo-UtcOrNull $a.latestRestorePointDate)
                TotalSizeBytes = $null; BackupType = $bt
            }
        }
    } catch { }

    # De-dup by name: keep the freshest record (and prefer one carrying a size).
    $deduped = foreach ($g in ($entities | Where-Object { $_.Name } | Group-Object { $_.Name.ToLowerInvariant() })) {
        @($g.Group | Sort-Object @{ Expression = { $_.LatestBackupUtc }; Descending = $true }, @{ Expression = { [bool]$_.TotalSizeBytes }; Descending = $true })[0]
    }
    $order = @{ 'Server' = 0; 'Workstation' = 1; 'VM' = 2 }
    return @($deduped | Sort-Object @{ Expression = { $o = $order["$($_.Type)"]; if ($null -eq $o) { 9 } else { $o } } }, @{ Expression = { $_.Name } })
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
# Icon styling (colors + logos): vendor brands AND integration icons
# Known entries are baked into IconMap below (accent color + icon URL or inlined data URI) so
# rendering makes no API call; any unmapped domain is resolved live via Brandfetch and cached. The
# API key is hard-coded per request.
# SECURITY: this key is visible to anyone with the repo - restrict it in Brandfetch (domain
# allowlist) or rotate if the repo is shared.
# ==============================================================================
$script:BrandfetchApiKey = 'ob7cfkVJLOtwFWiIVN60OjkZywKeFtvnS6l95rfZ7oC0cGUHDtjhcpy_Kt9JuX6l96rWcyWYyP_zCn5D3MPZHA'

# Persistent Brandfetch cache (domain -> @{ Color; Logo }). Brandfetch's free tier is rate-limited, so we
# resolve each unmapped brand at most ONCE EVER: misses and hits are written to brand-cache.json and
# reused on every future run (a miss is negatively cached so we don't retry it). Static vendor logos live
# in $script:IconMap and never hit the API at all.
$script:BrandCache = @{}
$script:BrandCacheDirty = $false
if (Test-Path $BrandCachePath) {
    try {
        $raw = Get-Content $BrandCachePath -Raw | ConvertFrom-Json
        foreach ($p in $raw.PSObject.Properties) { $script:BrandCache[$p.Name] = @{ Color = $p.Value.Color; Logo = $p.Value.Logo } }
    } catch { }
}
$script:IconMap = @{
    'itglue.com'      = @{ Color = '#3860be'; Logo = 'data:image/png;base64,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' }
    'connectwise.com' = @{ Color = '#5ea4de'; Logo = 'https://cdn.brandfetch.io/idejubPisD/w/400/h/400/theme/dark/icon.jpeg?c=1bxljtlqy4vv80d5kade1ohx1blb2u0lzBN' }
    'sentinelone.com' = @{ Color = '#6b0aea'; Logo = 'https://cdn.brandfetch.io/idqbZJLrXa/w/400/h/400/theme/dark/icon.png?c=1bxljtlqy4vv80d5kade1ohx1blb2u0lzBN' }
    'bitdefender.com' = @{ Color = '#EB0000'; Logo = 'https://cdn.brandfetch.io/idtcaX4QNF/w/400/h/400/theme/dark/icon.png?c=1bxljtlqy4vv80d5kade1ohx1blb2u0lzBN' }
    'duo.com'         = @{ Color = '#74bf4b'; Logo = 'https://cdn.brandfetch.io/id7s-uuKhk/w/400/h/400/theme/dark/icon.png?c=1bxljtlqy4vv80d5kade1ohx1blb2u0lzBN' }
    'cloudflare.com'  = @{ Color = '#F38020'; Logo = 'data:image/png;base64,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' }
    'veeam.com'       = @{ Color = '#03D15F'; Logo = 'https://cdn.brandfetch.io/idVHk_jeH3/w/400/h/400/theme/dark/icon.jpeg?c=1bxljtlqy4vv80d5kade1ohx1blb2u0lzBN' }
    'n-able.com'      = @{ Color = '#c046ff'; Logo = 'https://cdn.brandfetch.io/idpmq52ZKx/w/400/h/400/theme/dark/icon.jpeg?c=1bxljtlqy4vv80d5kade1ohx1blb2u0lzBN' }
    'sonicwall.com'   = @{ Color = '#F36F21'; Logo = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAAAnCAYAAAC/mE48AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAASLSURBVGhD5VhviFRVFJ+KzHn3vl12y6y0KKRECrQIEfwQCAkRgRKCuiBROO7M+zOjo6IFJaJFEP0RBD9Y/vkkC30I/9AH16gvQn0oVApaKMlVFCx2YVvX3X3n+jt3zizDOLrqPufP2x/8uG/u3HfvOeeee865LxUHqDhb0Sb3BQrV68bTGQrUbgr0AfL195GvfiZP/Ynni2j/BQfJ08P4PQqO4zlCS5aebblvDByJPDVkPDVAvrqK/n60fej/DXP+iPaY8fVB9H+J5w/QbkD7NtZ9DXK8RBvSc6iYUiJirHiACupFCtx3KdRfYeEfIiucHjcbXWM2g1ukLYKbwIIuMS8MKxhUsfI/Zvmd8hy8Bs/Jc/Ma1evxGF8byDME9sMoZ8Fe9PXA+PvAT8A8+laSl35VdLo9KEg9AqWXscJRoM/ByjSxcFlBFhYLN5wsBxuJZdsGbpdn9MMgV+BFp8GvwRBcJirWhglnPmPyagcU75uYlFveqVqL15ssB3sH7zzLxi367M77+iQ26nPju++Q37YYz7NErcnBgymvP4PVBuzE4lYNJyvMspS9D8rj+F2GvCfh1p8i9nSR7yw0uVlaVLl72EBS0BfM1tICNQWpFysVZqIPSvahvwdtAXFoicl0tIvoUwesuNO6Ebt5tTD1Ip9hjiviedjd8+AReOR6ju4mk3pYxI0XVnne9UYEs/I5xi7j7A4j5pxC+z48cSltTKVFxPsHCLHCulg9lee1ykoH6j/s8lHeZZNvf1bEqg/Ywtj9v63b1RI0bko2wS6PQOkTlHfXwbVnizj1B+VUlz1vtYSNixzM2MNgZCh+BtxKofu8iNBYQMCecoSNnaI4FCbwGBXcN81HqQdl6eYA0t6F2PO8KB6VFD9MBecVWa75YAWOM/hxLAGjUH1LgfOyLNO8gAeMxGIA3vVScDs3FrhvyfTND+TbX6dc+Egeh/J7UKg4MnVrAClw/5SyAFdrBU5pqkumbC3gSrjmnrMAjg6KlzHqdt6Q6VoP5LmPwgsG7+niwykup7MyVesCgfDAXR8DLmo8dVqmaG2YnLMIrkw2ktdSthY56HlqrUzR+kBJfOiOvYDPvq+vUZCeK6+3PlCtPYk0NnhHVSFHfk/31+WqWk+Q76yzGWGyo8AG8PWl+/XZuaHAUfjGfhippXiZpSMwSn77c/JacoDb2gwUNT9NGg9KQXC1vJYsmIx+jGv62xZIbICc+k5eSR7ISz9Fofr9lp7AxyDU1ylomyevJA+oD57Anf6XW8YEWwk6R2R4MkFBZxuKpKPWCNXZgX9zRZh1lsvw5ALu/oX9iltdJ+AqjaB53mTbO2RocsFXXlyaBm4KjogTuFUmNyBWgrrd+TgSvTY4Vt4g8TvKqY9lWPKBLLANhhia8AaOB0VtxrM6I0OSD8rqBTDCcWsE/qwGj8Bv83/QuUSGTA+gXlhNefWHPRbIFuOe/uuf9+Z2yt/TA/wxlAK9HR5w2XzomlHP7ZW/pheoWz0Oj9hFRX191G87KN3TD5Sd+XTkOXuHcx1rzKpVD0l3kyGVugFEn8NZDBF64QAAAABJRU5ErkJggg==' }
    # Capture Client band (Devices card) - same SonicWall logo, distinct label.
    'captureclient'   = @{ Color = '#F36F21'; Logo = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAD8AAABGCAYAAAB7ckzWAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAABB6SURBVHhe7VtpbFTXFab73h9tk7aquiutqi5qVbVqparLr6pVq6pSWwiLsdnM1hCcOOwBSkiJISyOsQ3esMFmMYvN7kCAtkA2whYwYQ0Yxxjv4xnPPvP1fOf5jsfjZ3u8TNqqHOl45p158975zj33u+fceR7xQB7IA3kgD+T/QXw+HyihUAhtbW1obW2F3+9XWzAY7GGjeDwedHR0IBAI6PuGhkZcuXIFZ8+ew5kzb+DUqZfx0kvHsG/ffuzeXYFdu3Zjx46d2Lp1O7ZsKRMt7XwtQ1nZNhQWFiMtbQ5mzUrH7NlPic7R4/T0+Zg/fzGWLn0WK1aswurVmdiwIR/l5Xtw9OgxvPbaGVy4cBE3b95Sf4zQV6/Xi3A4HDkmBmIhTgpxR8DzRJ7Q0tISAcoTjc3I3bu1ePnlV1BV9aKA2oPNm0uRl1eI3NyNyM7egPXrc1Wzs3PVlpubJw5bunFjfqcWqOblFaCgoAhZWTmYOHEakpOnICUlNaI8Hj9+MpKSJmHcuIkYO3YCRo9OwaOPjlcdM2aCnpOaOhNPPDEPzz67Evn5Raio2KcDwcGhEA8xEIsJiILnyLlcLrjdbj2JxlgbM6C29h3s3LkbOTkbFBhfCYgA8vMLFcRgtKioWK+VmvpXTJo0HZMnzxiQ8jsmcAzQ6NHJEphkfT9v3mLNDgOWWIiJ2IhxBA+amprgcDj0JAo/pK29vV2Pm5ubFSRB2wEYig4VfF86duxEzZBXX31dcVCIk9iIW0eeBgLmCFP5AW2cRzw+dKhKUjnH1vmhaiLBU5kBc+Ys1JHmNCZOYtORN3PejvAoDQ0NKC4uiczP4VaCJ1dMmTIzIeB5TXLG+fMXIgGg6Jw34O0Ij3Lr1q0hz+u+NNHgqeSBw4dfVFw9CM/pdGqK80OqsQWDAVy7dk2ds3N8OPTdAM95v3VruQLnlCY2Tfu+CI+B4DqaqPlOtcDnJnjkU+QeGxUbSbwb4dFgihYqP6CN70+ePKXrtp3jw6EEz+snEjxJjzWAIXNi05GPJjy7OX/kyFFlYzvHh0MNeDqZKPAslBYsWKKEbiQuwjt48PD/PPjk5FStALlyGYkQnl3akxR4wrsDPkedTBR4lsrsG8hrxBRJ+94Ij1UdPxsKeC6PVC6V1K7aPj9S79O+du0LkRqer6bGnzBhqirL14kTp2twBhMgXmPmzDQdUIOXr92WOruR37//gIDf2Cs4Ok8QbGIYJI6iaWrYvPC8TZtKtAFiB8fObtu2HdIn7NLGaM+eSvmsDAsXLtF5yXo8PX0BHn88XR2eNm2WkqFVv6dqcFi2WjX8eH0lm48Zk6LExsCZgJlA8XjGjNlobOQK5uta6vqb85WVexWIAUugBEmAfF9cvFlazF0SpIPaxrKlvXr1Gm7fvoObUiDVSBfY1u6E1x+ExxdAU1s7WhwuWP2WtM2irS6Pql/qDx67fUHU3W/C3Xfq0dDUgvqGJrx9u0aqtIt4/fU3pGM7j9OnX5GsrNLA5ebm44UXcoTRMzRwBErQSUmTNTB/+ctYDdydO3f0nhSd8wZ8b2xfWbkPmZlZCpYBYF9eVXVEnDijnZ7b41WHneJ5h1U8RcQpCNs9UlXpUQghr9TVDbVw3LuNUFu9fHgfoea7aK95C+211yybW9rOjlY4m+vham3S7xlxy1tP53sj9LTLW6DDH0LNO/dw9fpNVFe/pU3N3r0HZJBKce+eXL9TFLwhPNPYmLSnjSdwqauoqNT+uKbmLlweK1gUjl6HNwBPSz2CNRcQOL8f/mM5COxeiGDxFASy/ohgxi8RfuZHwOJvAwseQXjuVxBO/wLw5OeBJz4LzH4I4dmfEX1Ijj8H8LM5X0J43tcQkvOx+DvA8h8jvPIXCKz9LQK5oxDcPBWBnXMROJSBwD8LEDxbAW/1cbhvnYW/6a5Ew63+MegcGCMMo0sMDo/47PP339L65CQKL+IQtK4WWS5ungZeWodA0UQElv8UofQvAjM+CkwaAaSIJne+ThSdTH0PMOV9QOr7gakfAKZ9sFM/1FOnip3nUHk+v8fv89rUCaLR9+Ax7XJeeMbHEE57GFj4DeDvP0Fg9a/hzRmJQOlMYN9S4EQ2Ok6WovXCUbgba7sIj0FgulM9JDwhQLeEjqMbbKmF719F8GX/WUbuq3Kj93bdnODo8PSPADM/Jvrx/4DKfXn/VAkUfRonOrrzlTbJsMDSH8C/4VEEDqyA9+xeeBtu90J4ahHbtX/ClzMKoVmf7ooyR0QibO/Eu6Vyf2aJyTT6lSp+PfVFhDN+Dl/RJLgPr0HorRNAWy3CAVnbfcJpAswwiC8Yjia8oLCyC60y1P5rJ4EXfiejKiPMi/NGtk68S8pgT5UpwGnEjGO2Pfk5BeotmY6OozkI3XxF2NAqX0mKzFgjbc4OWS3u6NTu1s8z7R1S87qEJdnCBnfNk7SWGxE057GdMwlXgpWpRLD0Q+Z+eN7Xgaw/SL29HO5zB9Bed0uXRiN86+jwolaY/uLFS9K/H5HVqViXP9YJJSVlmt2sZxgEXefdctDU7oajtRFY9xsgSW42/cPiwCdiHEqwMtAkNo4s05msn/V7ePf9HQ4SlIPLniUK1O3Ffd0yv6rbbFzn5817GlOnPqaVIjcxzUbmn/40BkePHtfvEnhXhScE55K1OpDxKwu4nWOJUpKUmbcc6WU/RGh7GvwXDsDfUqfOGvFIkVT7Tp0UN69KP1CCpUuXawVIoKNGcRs7Rau/6MrOKANw8eKbkZaWq5lV4fHKFQssZrRzcLiV89esFtM/qvPWs2cJ2qv/1VkMWdLu9uH6jVs4dfJlreKefnqZgrLK2WQFzbK1v1p/woRpUiI/hrffvq2gmfoUnfO+2mrLoUSTGq/PEZbUDi7+HkJ7Zd2tvajsy5KEK0yHZCDTeMeO3ViyZLnOVSt9xytYjqodwL6UJe7ChUs13XsQXogEN56jnqA5ziJFRjk861MIFKQAV2X5kZWFbOyRQXBI3V9dfUX6hAKkpc2NpPFgwcYqg1dYWCJ3Cytwjj43NXTO47mfWfPOzvGhKFObI73om/BW/g2tt6/oCFNkiZXm56owchVKS7di1arVOrrjx08ZVMvalzKQx4//o/POMYQHVmxcQ+0ADFgle0hcHOn5j8BXtRY+Z6uWxkzrxsZGbYjY0rIzpBYWbhKmXq+ghxs4M4f7AOwwA4Gun+K43CnhhZ+UJYUO24IZiApvcKQlvXFoBYKedjgFMBuhuro6if4JBcp+n62w2RNgr5+ZacDbgxiscuqQO9issXqldiO8EIuHoYz8X2W0SWYy2v7MPyBw/7penHLjxg3t8wmW+3Tc+DCgo8GvW5elI2QHYCg6alQSdu7co74QuPmJmgFQ8Fgl6/tg5zyBk9CmyXp9fH2khayru6e/kJiND26CxII2mijwzCSu+/zdgYCNsomLEF64PL2T7W3A9akCXMrO0IxPwn2xSqNLOX/uvILl5kdfoI12gR86s0crgc+du0gbNc5vI90Iz3/piIy8lJUDbUdZnU39EPxvVimZsffnaHNO26V3b5oo8GR5pjxHm+D5pAaV7wlcCc/nk6gs+77VwdmB7E1ZoVUs0mjyQtu2bUdWVrYtwL40EeAtlp+G+npr24p7FNyNZhZ0r/DkTfjo2oGl/rQPIjT7YQTbG3Wes7EYDHBqIsCzzl+5co2CJFiC5zxnyrO+jxCe2xeAo+k+QrpDI+RlBzZWuXmQN1qrtMuXqzXV7YDFowTPffvhBM/anzu84bC1Kct1nYCpJDzadM5rS+uVuvpEgbVOxzP3mfKHngWzhlvbZHQ7YPEowa9Zkzls4NnBLV78jAI1e5F8NdKN8JQEvD64A0EEuVnAAHAJswNtVMAHDz0Hh/QI1tNYg39qY7jBk+iOHTsRITpDdrHv+dq1jSXaVl+D0FNfttb9vgIgAeKOaLN8lc4PFfxwpT2XNz6719LSrKsPA0BhhWdPePKHJ5EI2mQSuy4dlxZXAHLbuLdOT9rSsDREbR1elEiv/d8y8uzgOOrEwvQmaGKLJTzaLMKT3CcBOJ3tFilIVPxndsnoy9KnNb9NADq3qd03zmB75eFutfpAdbjAjxs3SX/rYzpTosmtd8KTPyQARsUI6cF5ohBhFj/MALspIKkfLn8Sx9+4guwhPLlB8KtXr9NdGTtQ8SqXt5MnT1sAOoWp3z/hRREBjV4hQN3+ZQawbufGYmwA2MykfQb1l1/FhuJtcZWydsrH3IYKnumemZmtac50JoZYcrOzdfvRgoRAYgjIHKGwSXGer0Io7fPW2h4NntOBK0PhaBw5fRbrB7ncDRV8Soq1j1dff199jq7mBkR4TBOmhbHxS1zOXDXVCD3zY6sKjN7L183IEWg/nInC8oPI25hvC7AvHSp4tq2VlfsVkCEyYiAWQ3jRth6Ex4jwAyMkCNpIDhIqJcEOTwe8m6ZYGcAtKjMNSIppn8Klw1uRtWkHCgaY/oMFz81N/va+bNkKAWP9GEGf+WqE/sfaiJM2nfP8E0t4PJm26IC4vH6tBN2nSvUnZK3ymAUMAjlh0SM4tGMz1heWCaj4A0Dwzz+/dsDg+dgKX7l3QKGv9DkWaKytB+FRmQY08ERjo3azyTlcSDoaa+HPHWX9nGQKIkl/1/xvobiwEBuKttgCtdPBgme6HznykgKij9E+09/ebMRpbPaEJ3OFwrnR0xZEi8uLFvla4GwFsOS7Fhcw/VNH4M6inyG7eAfyCzbZgo1VPtayatWauMEz3UeOTNKtbgrXbPpIXynR5GZn60F4JAGewDShxmNz0CYXCXqcCEmdj8cfsoIweQRey0hGVskuW7CxOpCRJ3D+d0X0I+TR5Eb/6KfxuT9b/4TXKb3aHE7dpUVzDbD1MWsFkCAcWvMUsop3CsC+5/9AwHNfnwG4fv1GJJWN2JFbXITHlIglPNpiAxKPzX33Mnwbk+B77NMoXzkX2YWlPQBHa7xpz8/Zrp47d17vY0du9KU/G3HSpuAZQaYAlSdRh2STG2h1XXMOzeVLUJK3Abkbex99gl+5crUWK3agqSxiOM/37TvIK0fuaUafOhjbIAjP3sZjKucThX/5WNit69eE/Ir0qcvBgCdw9uhlZdvVaRKcHZHFa+tGeIyAITLOEaYKP6CNpWK8Nh5T+b7L5oRfrssfL8xjpwMFzxHPyytShzly0ffg/XlMf+K1ESdtOvJDJrwYGzuoWBtv9uablzofSe0egL7Ajxw5DuvWrY+s2UZ4bd6D9zcSr404aUsI4RG8nY3nXrp0WZ/HjQ4AwWdk9ARP4PzvCC5pvFZ0MHnMe8SCisfWg/CYHlQ6SWWa8Jivw2mjcLeX4Plvod3BW+WqNceTdMTNaPP78d6jPxuPqZr2/IAy3IRnZ+MrI37lyluRh5gJ/rnnnlfwBjj/PZXnceQMafG75nrGxvvzmP7Ea+tBePyQ0WBkmF78YKA2Hvdn4yuP+Z4PLXMXh/t/GRnPdz6RkYRt28rVOZPuvE70d6mxNjOiA7HpyPc15/kFI/HaeIPebHw1Qgdqa2tRXr4TixYt1d0YPh1N5+iTEV4n9rvGFj2X47V1m/P/KfD8nMpd1S1btuq/nVOYmtHf5Xu779LWH1A7Wxd4N/4NFBCeX8COYG4AAAAASUVORK5CYII=' }
    'cloudappsecurity' = @{ Color = '#F36F21'; Logo = 'data:image/png;base64,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' }
    'watchguard.com'  = @{ Color = '#ED1C24'; Logo = 'data:image/png;base64,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' }
    'ui.com'          = @{ Color = '#2282FF'; Logo = 'https://cdn.brandfetch.io/idLAa3-67H/w/200/h/200/theme/dark/icon.jpeg?c=1bxljtlqy4vv80d5kade1ohx1blb2u0lzBN' }
    'microsoft.com'   = @{ Color = '#3A599A'; Logo = 'https://cdn.brandfetch.io/idchmboHEZ/w/800/h/800/theme/dark/symbol.png?c=1bxljtlqy4vv80d5kade1ohx1blb2u0lzBN' }
    # Entra/Azure AD-sync pill icon - inlined data URI (downscaled from the supplied
    # azure-active-directory.png) so it's self-contained like the Automate logo.
    'azuread'         = @{ Color = '#00ADEF'; Logo = 'data:image/png;base64,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' }
    # Domain controller pill icon - inlined data URI (downscaled from the supplied windowsserver.png).
    'winserver'       = @{ Color = '#0C499C'; Logo = 'data:image/png;base64,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' }
    # ConnectWise Automate: no Brandfetch entry / favicon - logo is an inlined data URI (downscaled
    # from the supplied automate.png) so it's self-contained.
    'automate'        = @{ Color = '#57b947'; Logo = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAArqSURBVHhe3Vt/jFxVFV4RxGpAEFCMBEGaYNCAutmd997M7tTSdue9N7PbUrZIQgygVlskUAkQf9TVkviH2mibQNNGA02lSSvQ7s6bNzM7224pVdtSBNtCgNRqxaYpNCimhqKU8fvunGm2u/fNe/NjG7pfcrKZmXvOPefcc8859763bWcaxuCc6YntfYutnP1d03ceNHP2UivnPmCVMnd3Dvb0ybCpCyNnr5u5d0G5e8fc0yj5x3nlxLbesrW551oZOvVwfXH2Rw3PfiNeSpfNnDOBun8/r2x4zoAMn3qw8m5v1/Y+rfGkxNZeOMB+uW1g4BxhmVpA+K/vfmau1nhFvjghZ8eFZeqg009dCMPejA/rw79KzAem56wStqkD00/d3FVr9YWYH7ANjrYPpT8irO8fJAeS53Y93deTHE2eK19FBpLbE1EcQGKeMHx3nrBGRtxPGZhngXxsLTrXpS60iu5wcudNZWvY3Z8oZRYaG26eJj/XxHQ/db6Zc9+yQsK/SspRnvOUsIfCLLpfgOy1jJ7uP8wrm1l3qfzUGnQ8deMlVjG9h2WKCjJRca9iPx+wSukl7aVZH5OhE5AouZ+NF92fQMGTTHLjjdWRVXDLZsH9tzmcvqMr3/MpETUBCT91HeT+Brr9j/qgyihSenrOchnWHLAXL4Uiz0tyOk3R+JYMJsPExfQ/zGL6h2Zx9ieErQ1R0gvlsoiaE9g2p/GFEuahExgJkP0WaH08784Q0SinPdditddA9n+V4XTYWH46AY0Vom6lsDSGDs+5HMK0xo+leImOwIR55w3swV+Znv0M9zFpgnJ1EoxXjmDUQYch0CpExwmt4WMJ0cbuEuOX92/s/6CYVB/MrP3kjN3zaxo/lljiTimr+b0pgkFd2+BURkVUpzISkBMMP5UUk+oDStedKiHphJ8FFB/J8O9rtXJUTTB0jKzzyqSs6BkgOVvcL+Y0hpjn3K5ygGaCRojhm0DyZGJkdFE2//Izv282Z1RJdZy+88+E514spjSG9j0Lz8OZ/SCV000UhVialME8CPnO24bv7MP3g1idNcgzK/HdGtAgvtuPfXui6hDyjZVTD8nq/0LMaA7o4b9d7QHqIiYuJi1kcrOQzhkF92vmFuczIlaL2GjfVVYxcztKXJ6NkzhNLz+AGEWgd1AqrxSxzQGt74ch+DVJKuGEqsGxzNow4ndGPv1lEVUXzFK6A/3EJkaDmjtiNeK2MrL2OhHTGhjZ1L2qzdRMeBpBSa4avH/UKqTnC3tTgJxbsKeP0RGhTsC2wdj3LM/5orC3Bqqf95wXa0YBjUfIQ4HnOzfMvlpYWwLeJULuPsqv5QQ6KZa1h4StdbB29F5g+Pbf2PXpJq6uPPbuczc8mrxI2FoKdSYpuHtVXghwQqVjtHe1lVt8q4RsvTTJLVBr4oJ7uGOjc7mwTAo6n0xdgbxwlOcQnR4k5oBYzr5FWJoHk6Dh2YeDwp/limXSGGqw5awTsUF7dmK0N7BMMvkiCe6U4c0jlnV71d7TTMaIkKy7WoaHgq1pfCRtm8PufagSA/Fi+q7E1ky8rT/6oSXm2Y+pBk0XkT4SYSldjg9mPi/DmwPqcOBtjqrxefd4bCTzSRkeCGO5MQ17+CEYfYTJigZUiVsI37+IfuEOGV4TnX7qCmy5t4M6R7UoeffnMrx+mKPpzyGhPYAJdqpJApoRNVGE1W/fOOtKZPE/8UaJqzNBFuTTCSy1mG99lGs3JLtKFIyXJfKwRd5T+sMO2iNswbC2zL0mXnDvRZJ52iw47/Kyg3stsBOj0uzf82lTRGjBG2GE5cuUF5RETxFkJnfdxHGPC3sgsDVnhOqH39W8sId2Kftgp4hAghtKX4ofFiGUS/DWOyoc+YgqQh/Okojk+Pf21e3niTgtGCFc+VDjqwTFVdPlObeKCC1Ucs7aR7QRNY5UooZdtI92KnuHM4vazJx7z8w/91fa1jpPYqzHOHBsEn206BycfbXyfp2y1cpm7ZfCbnMwNqc6RI2MIKIutHfmCwvKCKPUg4H7KIQUn+f8VHTRAifJexo9SLHsxrw57SJKC8y/vHH955VVg9OwAJ74POdu0UULOOCxmo/DapDiy6buFFFaGEPO/U05wPBSP2pUgDrq5uxviC5a4Oy/SbWuGv4wol7YYktElBaIgMUNO1g5IGt/v5kIAP93RBctJjsCEMH3NR4BcxkBzhJVJjQDwkitUM5+SHTRYtJzQM7+WVMOYKmZ8ez8SjhHfHRVpUpo2xtFFy0mvwrYg3VXAbm+55W/qqVW3ulDl/ZbKHmMP6hjbQSFuULYAgfDlJysPoD9B/uQwCP6GFKlD3aphYadyl7YLaIqSI6iKSpmbkNjNGQVnP8wRNg8UCGd0MqhA04IufKarE6QT4PVsbiGfgnUe9pBe5RdsI92iohg8ELRLKUXo2MaZU8dNEnFKPuXwhaIus4CfrSzAFb/kcD9D+NxSDtJ/WkH7RG2+gGBQ0GlDEbxmeCbyU19obdArTwN8qEtZPHBqVYvteq+u0KGNwfuxUBPgxgFxqAd+ejZkvuArL0ycEtJNJlDqQ4Z3hyszb0XQPCxoEMHE4xVSr9rbk59SVgmFbHNTgwJ+CTn1elTqSDOXgz9QIWjBVD1NuhOEN+pSQvuK3SWsEwKEo+7FyPs/6LmC0ioKiJDOtS6wZcfkAyPBvYKUEbtu7y75bqN/R8StpbC2GBMM4vOdtUhBhjPqoDkuJ+P9IStNWCtR7Y/GHQxqohOgPdRbrZ0rL3xEmFtCRIbUpchV2xXuSjAeBL3vpFNPStsrQMmXVwrEZ4iOoHNRin9asMvJowDtt9MOPUgm5haxleJ87f0WpxPhbCnDtW6iz+NoCT3KJskazjzcCzfd5WIqguWn7oGW241o05l9QjGkzg3tsG+tnKLkiAOTN9q5EDDayiVFwrucZS8R42Ck5pV6q/5tkZytO8irLaDRLfWKrqqG6UcnfxapFrenN38s0kmNHj+r029H1B90QkrY/nOUXRpW42csxqOXRbznB8YvrMMtZvvCIyie3ud46Rn18qLQhIFL4gZjQNKfl01HJpJ6iY0KOwceb6ggVzdKikH4XvVWQb19nWScqI37sBTD9iPwwEH0KVpJ3i/kzrIefZzYk79QPa9i6ujE362kDrD+O5XxaT6gHP8+kbeE6zZKzRBrAQqrAMOQDriZY+ZbfCNUZ7wkJR2BB46hGiw5InXETW/Br3Kz0xEze5nVoDqC5JInLuxJZks/8WqFPZARLXuOSfX1Kv3lbfE07tUGRznBFYGVaaK7qF40f1ewk9dRh5OiFK2CDV8D3sHKj+WLxLhPF+Rnab8bYnhzK3Vum6MuJ+G7GX47QgdrYs4ZXzRHZ6+InU+eZrCDZuSrM07xKNqZSUUX0rw0qTGASg+nDHQvz8S5drqFPEyo+CeNEvujxMjmetF1AQYG+Z8nEdqbL0DdJZqlhBxfEcYkZNvifFVqFUtuh6dEB92d1sjmdvCngmOBcL3cFQnqBcccvZ2YQ0F/28BC/RNLMheZXzBXR92T9kYBtrO6d7W1yWf6gK2z6qoFUWN89yFwhoZNLp7R2P6TTo6PecrUZIi9zz+Hq/mkykDucI+FFYiVbb3nCeEbWoBh5MVYdtAPZPIu73CMrWAxBavZmqd8ZUkaR+J+s9YZx2YpIyscyDoXoE1Hd3nwzJ8agLVYPnMfRP/e5yNliqxfsqQoWcAbW3/B5+yVWJNKxhhAAAAAElFTkSuQmCC' }
    'godaddy.com' = @{ Color = '#00a4a6'; Logo = 'data:image/png;base64,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' }
    'amazonaws.com' = @{ Color = '#FF9900'; Logo = 'data:image/png;base64,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' }
    'azure.microsoft.com' = @{ Color = '#00A4EF'; Logo = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAIISURBVHhe7dsxT9tAGAbgG+PwC3ASOjUxxOpvKSUF0YEh2U7Kkj9AoLSildpO/RUlHbpUlTOkCwInA2FCmAlWX0wiVZCA/PUuvWZp1nxL35Ne3fp+j89eThbDdTe6f5FTw+cuW9KtgkoqyzVhV/OkWHt3VVbNXoktB33P7JEGyI3oZYHuNvJsoe0Vuq0s1+38onn6tP7x2qfX/VW2HF6s0W7PG4nb9Vx8r0tpCLbQVkEDuNLOL/bConx/VaZmz2PLwfmqBijFAAAAAAAAAAAAAABziy4qAAAAAAAAAAAAAAB2fgAAAAAAAAAAAACAuUUXFQD8BZgWerUyLcWWnSc0rLgNO7/Y6xYbn9UzMghc+XTjawCPRLLhVu8289I8Ea6k23mpNvO+nV+8CT3/w2VZmpPAlbd9T+6HpaqtgPUfry9JVXz7JUVrwJcg1buavQLjH0t+GjpyEvAl7ThyHGT1K9AakOgQiYAxoc7XZPYRnASZBkVLRN0sX86yNGk7JMTRINYngEQr4YtBMCfBLvNETKkHXYgr9NPRAJkYAAAAAAAAAAAAAODfkosMAACgdwAAwM4PAAAAAAAAAAAAADC36KICAAAAAAAA+AOQjMT3R5oicMXcRbYGs5+nNUDd3NWZUmw5ngKMzOVopEspvfMlSJU4ime/z4/b2VradZQuxJa0k1EP7Uz0G5gXpC3uM4p6AAAAAElFTkSuQmCC' }
    'google.com' = @{ Color = '#4285F4'; Logo = 'data:image/png;base64,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' }
    'namecheap.com' = @{ Color = '#d4202c'; Logo = 'data:image/png;base64,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' }
    'networksolutions.com' = @{ Color = '#007932'; Logo = 'data:image/png;base64,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' }
}
function Resolve-Brand {
    param([string]$Domain)
    if ([string]::IsNullOrWhiteSpace($Domain)) { return $null }
    if ($script:IconMap.ContainsKey($Domain)) { return $script:IconMap[$Domain] }
    # Already resolved (this run or a previous one)? Use the persisted result - no API call. A negatively
    # cached miss has empty Color+Logo, which we surface as $null without re-fetching.
    if ($script:BrandCache.ContainsKey($Domain)) {
        $c = $script:BrandCache[$Domain]
        if ($c -and ($c.Color -or $c.Logo)) { return $c } else { return $null }
    }
    try {
        $b = Invoke-RestMethod -Uri "https://api.brandfetch.io/v2/brands/$Domain" -TimeoutSec 20 `
            -Headers @{ Authorization = "Bearer $($script:BrandfetchApiKey)" }
        $color = ($b.colors | Where-Object { $_.type -eq 'accent' } | Select-Object -First 1).hex
        if (-not $color) { $color = ($b.colors | Where-Object { $_.type -eq 'brand' } | Select-Object -First 1).hex }
        $icon = $b.logos | Where-Object { $_.type -in 'icon','symbol' } | Select-Object -First 1
        $fmt  = $icon.formats | Where-Object { $_.format -eq 'png' } | Select-Object -First 1
        if (-not $fmt) { $fmt = $icon.formats | Select-Object -First 1 }
        $entry = @{ Color = "$color"; Logo = "$($fmt.src)" }
        $script:BrandCache[$Domain] = $entry; $script:BrandCacheDirty = $true
        return $entry
    } catch {
        $script:BrandCache[$Domain] = @{ Color = ''; Logo = '' }  # negative cache: don't retry this miss
        $script:BrandCacheDirty = $true
        return $null
    }
}

# ==============================================================================
# Tile providers - each returns a card object: @{ Title; Items[] }
# An item is @{ Label; Sub; Link; Brand; Bg; Muted; Alert; Kind }:
#   Brand -> domain/integration key into IconMap; sets the logo + accent color.
#   Bg    -> explicit background colour that overrides everything.
#   Muted -> force the neutral translucent background even with a brand (secondary 365/license lines).
#   Alert -> outline the pill in red to flag a problem WITHOUT recolouring it. Used for MFA gaps,
#            over-provisioned licenses, Duo bypass users, etc.
#   Kind 'pill' = logo + Label button; Kind 'number' = big Label over small Sub (always muted, no logo).
# ==============================================================================
function New-CardItem {
    # -Detail makes the pill EXPANDABLE: it renders as a <details>/<summary> disclosure (native, no JS),
    # and clicking it opens a panel of detail rows below the pill, inside the card. -Detail is an array
    # of New-DetailRow objects (K/V rows, or list items when only V is set); -DetailHead is an optional
    # panel heading. When a pill is expandable AND has a -Link, the link moves INSIDE the panel (rendered
    # as "<LinkText> ->") instead of making the whole pill an anchor. An empty -Detail leaves the pill
    # non-expandable (a plain pill, or a direct link if -Link is set) - so detail-less data degrades cleanly.
    param(
        [string]$Label, [string]$Sub = '', [string]$Link = '', [string]$LinkText = '', [string]$Brand = '',
        [string]$Bg = '', [switch]$Muted, [switch]$Alert, [string]$Ring = '', [ValidateSet('pill','number')][string]$Kind = 'pill',
        [object[]]$Detail = @(), [string]$DetailHead = ''
    )
    [pscustomobject]@{ Label = $Label; Sub = $Sub; Link = $Link; LinkText = $LinkText; Brand = $Brand; Bg = $Bg; Muted = [bool]$Muted; Alert = [bool]$Alert; Ring = $Ring; Kind = $Kind; Detail = @($Detail); DetailHead = $DetailHead }
}

function New-DetailRow {
    # One row in an expandable pill's panel. With -K it renders as a "key .... value" row (value coloured
    # by -State: 'bad' red, 'ok' green); with only -V it renders as a plain list item (e.g. a username or
    # device name). -Link makes the value a clickable link (e.g. a firewall WAN IP). See New-CardItem -Detail.
    param([string]$K = '', [string]$V = '', [ValidateSet('','ok','bad')][string]$State = '', [string]$Link = '')
    [pscustomobject]@{ K = $K; V = $V; State = $State; Link = $Link }
}

function New-Card {
    # -NoBand opts the card out of the auto vendor-banding (Group-CardItems) so its pills render flat
    # (used by Domains, whose pills are many different DNS hosts - no single vendor to band by).
    param([string]$Title, [object[]]$Items, [switch]$NoBand)
    [pscustomobject]@{ Title = $Title; Items = @($Items); NoBand = [bool]$NoBand }
}

# Brand key -> vendor display name, used to label the auto-generated vendor bands (Group-CardItems).
# A pill is only banded if its Brand appears here (so unknown/iconless pills render flat).
$script:BrandName = @{
    'connectwise.com' = 'ConnectWise'; 'itglue.com' = 'IT Glue'; 'duo.com' = 'Duo'
    'microsoft.com' = 'Microsoft 365'; 'azuread' = 'Microsoft 365'; 'winserver' = 'Domains'
    'automate' = 'ConnectWise Automate'; 'sentinelone.com' = 'SentinelOne'; 'bitdefender.com' = 'Bitdefender'
    'sonicwall.com' = 'SonicWall'; 'captureclient' = 'Capture Client'; 'cloudappsecurity' = 'Cloud App Security'; 'watchguard.com' = 'WatchGuard'; 'ui.com' = 'UniFi'
    'veeam.com' = 'Veeam'; 'n-able.com' = 'Cove'; 'cloudflare.com' = 'Cloudflare'
}

function New-CardGroup {
    # A labelled sub-box inside a card: the box is the brand colour with a small brand label/logo, and
    # its inner pills render in the default (neutral) pill style. Used to cluster a vendor's pills
    # (e.g. all Cove backup pills under a 'Cove' box). Treated as a card item with Kind 'group'.
    param([string]$Label, [string]$Brand = '', [object[]]$Items)
    [pscustomobject]@{ Kind = 'group'; Label = $Label; Brand = $Brand; Items = @($Items) }
}

function Get-DashAntivirus {
    # Category tile: lists every active AV product the client has (SentinelOne, SonicWall Capture Client,
    # and/or Bitdefender).
    param([pscustomobject]$Entry, [string]$OrgName, [hashtable]$Creds, [pscustomobject]$MswTenant)
    $lines = @()

    # Gather candidate lists for every AV platform up front, then decide prompting. If the client is
    # confidently matched on ANY AV platform, we don't interactively prompt for the others (a client is
    # usually on a single AV product); the unmatched platforms are recorded as unmanaged silently.
    $sites = @()
    if ($Creds.SentinelOne.Configured) {
        foreach ($inst in $Creds.SentinelOne.Instances) {
            try { $sites += Get-S1Sites -BaseURL $inst.URL -APIKey $inst.APIKey -InstanceName $inst.Name }
            catch { Write-Status "  SentinelOne $($inst.Name) query failed: $($_.Exception.Message)" Warning }
        }
    }
    $companies = @()
    if ($Creds.Bitdefender.Configured) {
        try { $companies = Get-BdCompanies -BaseURL $Creds.Bitdefender.URL -ApiKey $Creds.Bitdefender.ApiKey }
        catch { Write-Status "  Bitdefender query failed: $($_.Exception.Message)" Warning }
    }
    $s1Matched = Test-VendorMatched -Entry $Entry -VendorKey 'sentinelOne' -OrgName $OrgName -Candidates $sites
    $bdMatched = Test-VendorMatched -Entry $Entry -VendorKey 'bitdefender' -OrgName $OrgName -Candidates $companies
    $anyAvMatched = $s1Matched -or $bdMatched

    if ($Creds.SentinelOne.Configured) {
        $site = Resolve-OrgServiceId -Entry $Entry -VendorKey 'sentinelOne' -VendorLabel 'SentinelOne' -OrgName $OrgName -Candidates $sites -NoPromptUnmanaged:($anyAvMatched -and -not $s1Matched)
        # Hide a 0-count product (matched by name but protecting nothing). Prefix EDR:/MDR: per the
        # instance the client's site was found in (SentinelOne has separate EDR and MDR consoles), and
        # deep-link into that console's site-scoped Unified Assets / endpoint inventory.
        if ($site -and [int]$site.ActiveLicenses -gt 0) {
            $s1Link = if ($site.ConsoleURL) {
                "$($site.ConsoleURL)/unified-assets?_scopeId=$($site.Id)&_scopeLevel=site&_categoryId=inventory&activeTabFromPreferences=true&activeTab=endpoint"
            } else { '' }
            $lines += New-CardItem -Label "$($site.Instance): $([int]$site.ActiveLicenses) devices" -Link $s1Link -Brand 'sentinelone.com'
        }
    }

    # SonicWall Capture Client: the active-endpoint count from the client's MySonicWall CSC tile
    # (single number; the total-licenses denominator isn't exposed by the billing API key).
    if ($MswTenant) {
        $ccCount = Get-MswServiceCount -Tenant $MswTenant -ServiceName 'Capture Client'
        if ($null -ne $ccCount) { $lines += New-CardItem -Label "$ccCount devices" -Brand 'captureclient' }
    }

    if ($Creds.Bitdefender.Configured) {
        $company = Resolve-OrgServiceId -Entry $Entry -VendorKey 'bitdefender' -VendorLabel 'Bitdefender' -OrgName $OrgName -Candidates $companies -NoPromptUnmanaged:($anyAvMatched -and -not $bdMatched)
        if ($company) {
            $count = Get-BdEndpointCount -BaseURL $Creds.Bitdefender.URL -ApiKey $Creds.Bitdefender.ApiKey -CompanyId $company.Id
            # GravityZone has no per-company deep link, so link to the console Network page (EDR-only,
            # so no EDR/MDR label on this pill).
            if ([int]$count -gt 0) {
                $bdLink = "$($Creds.Bitdefender.URL.TrimEnd('/'))/#!/network"
                $lines += New-CardItem -Label "$([int]$count) devices" -Link $bdLink -Brand 'bitdefender.com'
            }
        }
    }

    # Duo Authentication for Windows Logon coverage (Automate software inventory). Only shown when the
    # client actually has Duo - reuse the per-client Duo match the Identity card resolves and caches on
    # the service-map entry (the build order resolves Identity before this card, so the cache is warm).
    if ($Entry.PSObject.Properties['duo'] -and $Entry.duo) {
        $duoDev = Get-AutomateDuoLogonItem -Entry $Entry -OrgName $OrgName
        if ($duoDev) { $lines += $duoDev }
    }

    return New-Card -Title 'Endpoint Security' -Items $lines
}

function Get-DashIdentity {
    # 'Identity' card. Pills, in order: Duo MFA, Microsoft 365 total users, MFA enforcement, Entra AD-sync
    # (all three via CIPP), and the on-prem AD domain + domain controllers (via Automate).
    param([pscustomobject]$Entry, [string]$OrgId, [string]$OrgName, [hashtable]$Creds, [pscustomobject]$CippTenant)
    $lines = @()

    # 1) Duo. The duo.com logo identifies the pill, so the label drops the 'Duo:' prefix.
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
                $c = Get-DuoUserCounts -Creds $Creds -AccountId $acct.Id -ApiHost $acct.ApiHostname
                # Expand the pill into a posture summary (counts) followed by the actionable name lists,
                # mirroring the M365 users / MFA-enforced pills. The Duo admin link tucks into the panel.
                # Red values: bypass (skips MFA) and locked-out (the two the pill's red outline keys off).
                # Each count row that has at-risk names lists them directly beneath it (the renderer nests
                # the consecutive name rows into an indented list), so "Not enrolled 1 -> the name" reads
                # as one unit - matching the redesign (no separate caption, no duplicated label).
                $detail = @()
                $detail += New-DetailRow -K 'Bypass' -V "$($c.Bypass)" -State $(if ($c.Bypass -gt 0) {'bad'} else {''})
                if ($c.BypassNames.Count)      { $detail += @($c.BypassNames      | ForEach-Object { New-DetailRow -V $_ }) }
                $detail += New-DetailRow -K 'Disabled'     -V "$($c.Disabled)"
                $detail += New-DetailRow -K 'Locked out'   -V "$($c.LockedOut)" -State $(if ($c.LockedOut -gt 0) {'bad'} else {''})
                $detail += New-DetailRow -K 'Not enrolled' -V "$($c.NotEnrolled)"
                if ($c.NotEnrolledNames.Count) { $detail += @($c.NotEnrolledNames | ForEach-Object { New-DetailRow -V $_ }) }
                $detail += New-DetailRow -K 'Stale (90+ days)' -V "$($c.Stale)"
                $lines += New-CardItem -Label "$($c.Active) active users" -Link $adminUrl -LinkText 'Open Duo admin' `
                    -Detail $detail -Brand 'duo.com' -Alert:($c.Bypass -gt 0 -or $c.LockedOut -gt 0)
            }
            # The client IS in Duo (account resolved), so still render a linked Duo pill - just without
            # the count - rather than a dead 'n/a'. The console warning carries the real failure reason.
            catch { Write-Status "  Duo user count failed for '$($acct.Name)': $($_.Exception.Message)" Warning; $lines += New-CardItem -Label 'Managed' -Link $adminUrl -Brand 'duo.com' }
        }
    }

    # 2-4) Microsoft 365 total users + MFA enforcement + Entra AD-sync (CIPP). Tenant is resolved once up
    # front and passed in. Grouped in a 'Microsoft 365' sub-box.
    if ($CippTenant -and $CippTenant.Domain) {
        $msItems = @()
        $usr  = Get-CippUsersItem -TenantFilter $CippTenant.Domain; if ($usr)  { $msItems += $usr }
        $mfa  = Get-CippMfaItem  -TenantFilter $CippTenant.Domain; if ($mfa)  { $msItems += $mfa }
        $sync = Get-CippSyncItem -TenantFilter $CippTenant.Domain; if ($sync) { $msItems += $sync }
        if ($msItems.Count -gt 0) { $lines += New-CardGroup -Label 'Microsoft 365' -Brand 'microsoft.com' -Items $msItems }
    }

    # 4) On-prem AD domain + domain controllers (Automate).
    $dc = Get-AutomateDcItem -Entry $Entry -OrgName $OrgName
    if ($dc) { $lines += $dc }

    return New-Card -Title 'Identity' -Items $lines
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
                # One multi-line pill per protected entity inside a 'Veeam' sub-box: device name on top,
                # "<age> ago [, size] - <backup type>" below. Each pill gets a red ring if its own last
                # backup is missing or >25h old (25, not 24, to avoid false positives from daily drift).
                $cutoff = (Get-Date).ToUniversalTime().AddHours(-25)
                $dot = [char]0x00B7
                $veeamItems = @()
                foreach ($e in (Get-VspcCompanyBackup -BaseURL $Creds.VeeamVspc.URL -Token $token -CompanyUid $company.Id)) {
                    $parts = @(if ($e.LatestBackupUtc) { Format-RelativeAge $e.LatestBackupUtc } else { 'no backup' })
                    $size = Format-BackupSize $e.TotalSizeBytes; if ($size) { $parts += $size }
                    if ($e.BackupType) { $parts += $e.BackupType }
                    $ring = if ((-not $e.LatestBackupUtc) -or ($e.LatestBackupUtc -lt $cutoff)) { '#DC3545' } else { '' }
                    $veeamItems += New-CardItem -Label $e.Name -Sub ($parts -join " $dot ") -Ring $ring
                }
                if ($veeamItems.Count -gt 0) { $lines += New-CardGroup -Label 'Veeam' -Brand 'veeam.com' -Items $veeamItems }
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
                # Both go in a 'Cove' sub-box, so the pills drop the 'Cove' prefix.
                $coveItems = @()
                $endpoint = @($devices | Where-Object { "$($_.Physicality)" -ne 'Undefined' })
                if ($endpoint.Count -gt 0) {
                    $servers = @($endpoint | Where-Object { $_.OSType -match 'server' }).Count
                    $workstations = $endpoint.Count - $servers
                    $tb = [Math]::Round((($endpoint | Measure-Object -Property UsedGB -Sum).Sum) / 1024, 2)
                    $coveItems += New-CardItem -Label "$tb TB ($servers srv, $workstations wks)" -Brand 'n-able.com'
                }
                $m365 = @($devices | Where-Object { "$($_.Physicality)" -eq 'Undefined' })
                if ($m365.Count -gt 0) {
                    # M365 backup: just list the protected tenant(s) by name - no TB/tenant count.
                    $m365Names = @($m365 | ForEach-Object { "$($_.DeviceName)" } | Where-Object { $_ } | Sort-Object -Unique)
                    $coveItems += New-CardItem -Label ($m365Names -join ', ') -Brand 'n-able.com'
                }
                if ($coveItems.Count -gt 0) { $lines += New-CardGroup -Label 'Cove' -Brand 'n-able.com' -Items $coveItems }
            }
        } catch { Write-Status "  Cove query failed: $($_.Exception.Message)" Warning }
    }

    return New-Card -Title 'Backup' -Items $lines
}

function Get-CloudflareZones {
    # All zone (domain) names the API token can see, paged. Used to tell domains that are in OUR
    # self-service Cloudflare portal from domains merely using Cloudflare NS via a web host.
    param([string]$ApiToken)
    $zones = @(); $page = 1
    $h = @{ Authorization = "Bearer $ApiToken" }
    do {
        $r = Invoke-RestMethod -Uri "https://api.cloudflare.com/client/v4/zones?per_page=50&page=$page" -Headers $h -Method Get
        if ($r.result) { $zones += @($r.result | ForEach-Object { "$($_.name)".ToLowerInvariant() }) }
        $totalPages = [int]$r.result_info.total_pages
        $page++
    } while ($page -le $totalPages -and $totalPages -gt 0)
    return $zones
}

function Get-CloudflarePortalSet {
    # The set of zones in OUR Cloudflare portal (lowercased zone name -> $true), used only to colour
    # domain pills with the Cloudflare logo. Loaded LAZILY on first use and cached for the run, so a
    # single-org run with no Domains card never makes the (whole-account) zone call. Returns an empty
    # set when Cloudflare isn't configured or the load fails (domains just fall back to their NS brand).
    if ($null -ne $script:CloudflarePortal) { return $script:CloudflarePortal }
    $script:CloudflarePortal = @{}
    if ($script:CloudflareConfigured) {
        try {
            foreach ($z in (Get-CloudflareZones -ApiToken $script:CloudflareToken)) { $script:CloudflarePortal[$z] = $true }
            Write-Status "Cloudflare portal: $($script:CloudflarePortal.Count) zones loaded." Detail
        } catch { Write-Status "Cloudflare zone load failed: $($_.Exception.Message)" Warning }
    }
    return $script:CloudflarePortal
}

function Get-NameserverBrand {
    # Map a domain's NS hostnames to a DNS-provider brand key (for the icon). IT Glue only stores the
    # registrar, not the DNS host, so the live NS records are the real signal. Returns '' if unknown.
    param([string[]]$NameHosts)
    $h = (($NameHosts | ForEach-Object { "$_" }) -join ' ').ToLowerInvariant()
    switch -Regex ($h) {
        'cloudflare\.com'                          { return 'cloudflare.com' }
        'domaincontrol\.com'                       { return 'godaddy.com' }          # GoDaddy
        'awsdns'                                   { return 'amazonaws.com' }        # Route 53
        'azure-dns\.'                              { return 'azure.microsoft.com' }
        'googledomains\.com|ns-cloud-.*google'     { return 'google.com' }
        'registrar-servers\.com'                   { return 'namecheap.com' }
        'dnsmadeeasy\.com'                         { return 'dnsmadeeasy.com' }
        'name-services\.com|worldnic\.com'         { return 'networksolutions.com' }
        'nsone\.net'                               { return 'ns1.com' }
        'akam\.net|akamai'                         { return 'akamai.com' }
        'digitalocean\.com'                        { return 'digitalocean.com' }
        'dnsimple'                                 { return 'dnsimple.com' }
        default {
            # Fallback: the registrable domain of the first NS host (e.g. ns1.example.com -> example.com),
            # which Resolve-Brand will try to fetch an icon for via Brandfetch.
            if ($NameHosts.Count -gt 0 -and "$($NameHosts[0])" -match '([^.]+\.[^.]+)$') { return $matches[1] }
            return ''
        }
    }
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
function Get-CippTenants {
    # All CIPP/GDAP tenants, fetched once and cached for the run. Each becomes a resolver candidate
    # {Id; Name=displayName; Domain=defaultDomainName}. CIPP's occasional aggregated all-tenants row
    # (a single object whose defaultDomainName is many domains joined by spaces) is dropped.
    if ($null -ne $script:CippTenants) { return $script:CippTenants }
    $script:CippTenants = @()
    try {
        $resp = Invoke-RestMethod -Uri "$($script:CippApiUrl)/api/ListTenants" -Headers @{ Authorization = "Bearer $($script:CippToken)" }
        foreach ($t in @($resp)) {
            if (-not $t.defaultDomainName -or "$($t.defaultDomainName)" -match '\s') { continue }
            $cid = if ($t.customerId) { $t.customerId } else { $t.RowKey }
            $script:CippTenants += [pscustomobject]@{ Id = "$cid"; Name = $t.displayName; Domain = $t.defaultDomainName }
        }
    } catch { Write-Status "  CIPP tenant list query failed: $($_.Exception.Message)" Warning }
    return $script:CippTenants
}
function Resolve-CippTenant {
    # Map the IT Glue org to its CIPP/M365 tenant, returning {Id; Name; Domain}. Resolution order:
    #   1) cached resolution;
    #   2) auto-confirm by IT Glue domain (ListTenants?TenantFilter=<domain> -> a single clean tenant);
    #   3) the generic resolver against the full tenant list (normalized display-name auto-match, with
    #      an interactive pick-list fallback) - so orgs whose IT Glue domains are only vanity/secondary
    #      domains (which 400 in step 2) can still be mapped, consistent with the other vendors.
    param([pscustomobject]$Entry, [string]$OrgId, [string]$OrgName)
    if (-not $script:CippConnected) { return $null }

    $tenants = Get-CippTenants
    if (-not $tenants -or $tenants.Count -eq 0) {
        # Tenant list unavailable this run: trust a valid cached resolution, else give up (don't poison).
        $cached = $Entry.PSObject.Properties['cipp']
        if ($cached -and $cached.Value -and "$($cached.Value.Domain)" -notmatch '\s') { return $cached.Value }
        return $null
    }

    # 2) auto-confirm by domain (only when not already resolved/unmanaged)
    if (-not ($Entry._unmanaged -contains 'cipp') -and -not ($Entry.PSObject.Properties['cipp'] -and $Entry.cipp)) {
        $orgDomains = @()
        try { $orgDomains = @((Get-ITGlueDomains -filter_organization_id $OrgId).data.attributes.name) } catch {}
        $orgDomains = @($orgDomains | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() })
        foreach ($d in $orgDomains) {
            $t = $null
            try { $t = @(Invoke-RestMethod -Uri "$($script:CippApiUrl)/api/ListTenants?TenantFilter=$([uri]::EscapeDataString($d))" -Headers @{ Authorization = "Bearer $($script:CippToken)" })[0] } catch {}
            if ($t -and $t.defaultDomainName -and "$($t.defaultDomainName)" -notmatch '\s') {
                $cid  = if ($t.customerId) { $t.customerId } else { $t.RowKey }
                $cand = $tenants | Where-Object { $_.Id -eq "$cid" -or $_.Domain -eq $t.defaultDomainName } | Select-Object -First 1
                if (-not $cand) { $cand = [pscustomobject]@{ Id = "$cid"; Name = $t.displayName; Domain = $t.defaultDomainName } }
                Write-Status "  [Microsoft 365 (CIPP)] auto-matched '$($cand.Name)' by domain '$d'" Detail
                Add-Resolution -Entry $Entry -VendorKey 'cipp' -Candidate $cand
                return $cand
            }
        }
    }

    # 3) name auto-match + interactive pick-list (caching/re-hydration handled by the generic resolver)
    return Resolve-OrgServiceId -Entry $Entry -VendorKey 'cipp' -VendorLabel 'Microsoft 365 (CIPP)' -OrgName $OrgName -Candidates $tenants
}
function Get-CippUsersItem {
    # Total M365 user count pill for the Identity card's Microsoft 365 sub-box. The headline count is
    # member accounts only (guests excluded - their account/licensing is owned by their home tenant).
    # Expands to a breakdown: licensed / unlicensed member counts + the guest count + the global admin
    # count, with the global admins then listed out by UPN. The member count splits as Total = Licensed +
    # Unlicensed.
    param([string]$TenantFilter)
    try {
        $users = @(Invoke-CippGraph -TenantFilter $TenantFilter -Endpoint 'users?$select=displayName,userPrincipalName,userType,accountEnabled,assignedLicenses&$top=999')
        if ($users.Count -eq 0) { return $null }
        $guests     = @($users | Where-Object { "$($_.userType)" -eq 'Guest' })
        $members    = @($users | Where-Object { "$($_.userType)" -ne 'Guest' })
        $licensed   = @($members | Where-Object { @($_.assignedLicenses).Count -gt 0 })
        $unlicensed = @($members | Where-Object { @($_.assignedLicenses).Count -eq 0 })
        $total      = $members.Count

        # Global admins: members of the Global Administrator directory role (always activated, so it shows
        # in directoryRoles). Two steps - find the role's object id, then fetch its members - because CIPP's
        # ListGraphRequest does not forward an $expand=members on the roles call. Members can be
        # users/groups/service principals; we list only the user accounts (those with a UPN) - the partner /
        # app service principals that also hold GA aren't "admins" in the sense being surfaced here.
        $admins = @()
        try {
            $roles = @(Invoke-CippGraph -TenantFilter $TenantFilter -Endpoint 'directoryRoles')
            $ga = $roles | Where-Object { "$($_.roleTemplateId)" -eq '62e90394-69f5-4237-9190-012177145e10' -or "$($_.displayName)" -eq 'Global Administrator' } | Select-Object -First 1
            if ($ga -and $ga.id) {
                $gaMembers = @(Invoke-CippGraph -TenantFilter $TenantFilter -Endpoint "directoryRoles/$($ga.id)/members")
                $admins = @($gaMembers | Where-Object { $_.userPrincipalName } | ForEach-Object { "$($_.userPrincipalName)" } | Sort-Object -Unique)
            }
        } catch { Write-Status "  CIPP global-admins query failed: $($_.Exception.Message)" Warning }

        $detail = @()
        $detail += New-DetailRow -K 'Licensed users'   -V "$($licensed.Count)"
        $detail += New-DetailRow -K 'Unlicensed users' -V "$($unlicensed.Count)"
        $detail += New-DetailRow -K 'Guest users'      -V "$($guests.Count)"
        # Global admins: a count row, then the admin UPNs nested in an indented list beneath it (the
        # renderer buffers the consecutive name rows into the member list) - "Global admins 1 -> the UPN".
        if ($admins.Count -gt 0) {
            $detail += New-DetailRow -K 'Global admins' -V "$($admins.Count)"
            $detail += @($admins | ForEach-Object { New-DetailRow -V $_ })
        }

        return New-CardItem -Label "$total users" -Brand 'microsoft.com' -Detail $detail
    } catch { Write-Status "  CIPP users query failed: $($_.Exception.Message)" Warning; return $null }
}
function Get-CippMfaItem {
    # MFA posture via CIPP's ListMFAUsers report (Get-CIPPMFAState). Unlike the raw Graph
    # userRegistrationDetails report, this returns per-user accountEnabled / userType / CA-coverage /
    # security-defaults / per-user MFA state, so we can both scope the denominator and measure real
    # ENFORCEMENT (not just "has registered a method"). Real-time call (no UseReportDB), one per tenant.
    param([string]$TenantFilter)
    try {
        $resp = Invoke-RestMethod -Uri "$($script:CippApiUrl)/api/ListMFAUsers?TenantFilter=$([uri]::EscapeDataString($TenantFilter))" -Headers @{ Authorization = "Bearer $($script:CippToken)" }
        # CIPP returns the user list as a bare array on a fresh compute but wraps it in an OData-style
        # { value: [...] } (or { Results: [...] }) envelope when serving from cache. Unwrap both, else
        # the envelope collapses to a single bogus row (the "MFA: 1/1" bug).
        $users =
            if ($resp -is [array]) { $resp }
            elseif ($resp -and ($resp.PSObject.Properties.Name -contains 'value'))   { @($resp.value) }
            elseif ($resp -and ($resp.PSObject.Properties.Name -contains 'Results')) { @($resp.Results) }
            else { @($resp) }
        $users = @($users)
        # Denominator: only accounts that can actually sign in and that we own the MFA for.
        # Drop guests (their home tenant owns MFA) and disabled accounts (can't sign in -> no risk).
        # Enabled-but-unlicensed accounts (e.g. shared mailboxes whose sign-in isn't blocked) stay in:
        # they are a live risk until sign-in is disabled.
        $relevant = @($users | Where-Object { "$($_.UserType)" -ne 'Guest' -and $_.AccountEnabled -eq $true })
        if ($relevant.Count -eq 0) { return $null }
        # Enforced = MFA is actually required for the user, via any of: a Conditional Access policy,
        # Security Defaults (tenant-wide), or legacy per-user MFA set to enabled/enforced. Split the
        # relevant set so we can both count the enforced and list the ones that aren't.
        $notEnforced = @($relevant | Where-Object {
            -not ($_.CoveredBySD -eq $true -or
                  "$($_.CoveredByCA)" -like 'Enforced*' -or
                  "$($_.PerUser)" -in @('enforced','enabled'))
        })
        $enforced = $relevant.Count - $notEnforced.Count
        # Make the pill expandable to list the un-enforced users (by UPN/email), like the Duo bypass
        # pill above. A plain (periwinkle) 'Not enforced' label captions the list - matching the M365
        # users / WatchGuard pill look. Empty when everyone's covered -> the pill degrades to a plain
        # non-expandable pill.
        $detail = @()
        if ($notEnforced.Count -gt 0) {
            $detail += New-DetailRow -K 'Not enforced' -V "$($notEnforced.Count)"
            $detail += @($notEnforced |
                ForEach-Object { (("$($_.UPN)".Trim()) -split '@')[0] } |
                Where-Object { $_ } | Sort-Object |
                ForEach-Object { New-DetailRow -V $_ })
        }
        # Flag (red outline) when not everyone who can sign in is covered.
        return New-CardItem -Label "$enforced/$($relevant.Count) MFA enforced" -Brand 'microsoft.com' `
            -Detail $detail -Alert:($enforced -lt $relevant.Count)
    } catch { Write-Status "  CIPP MFA query failed: $($_.Exception.Message)" Warning; return $null }
}
function Get-CippSyncItem {
    # Entra Connect (AD) sync pill for the Identity card's Microsoft 365 sub-box. When sync is on, the
    # pill shows "Synced (<age> ago)" and EXPANDS to the M365 admin "Synced from on-premises vs In cloud"
    # breakdown: the on-prem domain(s) members sync from, the synced-member count, and the cloud-only
    # member count. Guests are excluded from both counts (their account/sync is owned by their home tenant).
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
            # Per-user breakdown (members only - drop guests). A user is "synced from on-premises" when
            # onPremisesSyncEnabled is true; the on-prem domain it syncs from is onPremisesDomainName.
            $detail = @()
            try {
                $users   = @(Invoke-CippGraph -TenantFilter $TenantFilter -Endpoint 'users?$select=userType,onPremisesSyncEnabled,onPremisesDomainName&$top=999')
                $members = @($users | Where-Object { "$($_.userType)" -ne 'Guest' })
                $synced  = @($members | Where-Object { $_.onPremisesSyncEnabled -eq $true })
                $cloud   = @($members | Where-Object { $_.onPremisesSyncEnabled -ne $true })
                $domains = @($synced | ForEach-Object { "$($_.onPremisesDomainName)".Trim() } | Where-Object { $_ } | Sort-Object -Unique)
                if ($domains.Count -eq 1) {
                    $detail += New-DetailRow -K 'Syncing from' -V $domains[0]
                } elseif ($domains.Count -gt 1) {
                    $detail += New-DetailRow -K 'Syncing from'
                    $detail += @($domains | ForEach-Object { New-DetailRow -V $_ })
                }
                $detail += New-DetailRow -K 'Synced users'     -V "$($synced.Count)"
                $detail += New-DetailRow -K 'Cloud-only users' -V "$($cloud.Count)"
            } catch { Write-Status "  CIPP sync-user breakdown failed: $($_.Exception.Message)" Warning }
            return New-CardItem -Label "Synced$rel" -Brand 'azuread' -Muted -Detail $detail
        }
        return New-CardItem -Label 'Cloud-only' -Brand 'azuread' -Muted
    } catch { Write-Status "  CIPP org/sync query failed: $($_.Exception.Message)" Warning; return $null }
}

function Get-CippDeviceItems {
    # Microsoft 365 device pills for the Devices card (via CIPP/Graph). Two lenses, both filtered to
    # "active in the last 90 days" so retired records drop off:
    #   1) INTUNE managed devices (deviceManagement/managedDevices), deduped by azureADDeviceId, split
    #      Computers (Windows/macOS) vs Mobile (iOS/iPadOS/Android). Intune's own joinType is unreliable
    #      (often 'unknown'), so we don't use it here.
    #   2) ENTRA join type (devices.trustType - each device has exactly one, so no double counting):
    #      AzureAd -> Azure AD Joined, ServerAd -> Hybrid Joined, Workplace -> Registered. Empty
    #      trustType (mostly MDM-only mobile, already counted under Intune Mobile) is skipped.
    # NOTE: the Intune and Entra counts are deliberately different lenses (management vs identity) and
    # CAN overlap for a device that is both Intune-managed and AAD/Hybrid joined - that's expected.
    param([string]$TenantFilter)
    $items = @()
    $cutoff = (Get-Date).ToUniversalTime().AddDays(-90)

    try {
        $md = @(Invoke-CippGraph -TenantFilter $TenantFilter -Endpoint 'deviceManagement/managedDevices')
        $md = @($md | Where-Object { $_.lastSyncDateTime -and ([datetime]$_.lastSyncDateTime).ToUniversalTime() -ge $cutoff })
        $seen = @{}; $uniq = @()
        foreach ($d in $md) {
            $k = if ($d.azureADDeviceId) { "$($d.azureADDeviceId)" } else { "id:$($d.id)" }
            if (-not $seen.ContainsKey($k)) { $seen[$k] = $true; $uniq += $d }
        }
        $computers = @($uniq | Where-Object { "$($_.operatingSystem)" -match '(?i)windows|mac' }).Count
        $mobile    = @($uniq | Where-Object { "$($_.operatingSystem)" -match '(?i)ios|ipad|android' }).Count
        if ($computers -gt 0) { $items += New-CardItem -Label "$computers Intune computers" -Brand 'microsoft.com' }
        if ($mobile    -gt 0) { $items += New-CardItem -Label "$mobile Intune mobile" -Brand 'microsoft.com' }
    } catch { Write-Status "  CIPP Intune devices query failed: $($_.Exception.Message)" Warning }

    try {
        $dev = @(Invoke-CippGraph -TenantFilter $TenantFilter -Endpoint 'devices')
        $dev = @($dev | Where-Object { $_.approximateLastSignInDateTime -and ([datetime]$_.approximateLastSignInDateTime).ToUniversalTime() -ge $cutoff })
        foreach ($m in @(
            @{ Tt = 'AzureAd';   Label = 'Azure AD Joined' },
            @{ Tt = 'ServerAd';  Label = 'Hybrid Joined' },
            @{ Tt = 'Workplace'; Label = 'Registered' })) {
            $n = @($dev | Where-Object { "$($_.trustType)" -eq $m.Tt }).Count
            if ($n -gt 0) { $items += New-CardItem -Label "$n $($m.Label)" -Brand 'microsoft.com' }
        }
    } catch { Write-Status "  CIPP Entra devices query failed: $($_.Exception.Message)" Warning }

    return $items
}

function Get-CippLicenseItems {
    # Live M365 license records via CIPP's ListLicenses (GDAP). Returns $null on failure (caller falls
    # back to IT Glue), or an array (possibly empty) of {Name; Used; Total; Link; Over} records on
    # success - the caller rolls these up into one "assigned/total licenses" pill. CIPP field names vary
    # by version, so probe a few candidates defensively. Free/unlimited/empty SKUs are hidden with the
    # same heuristic as the IT Glue path (total of 0, or a large round sentinel >= 1000).
    param([string]$TenantFilter)
    try {
        $resp = Invoke-RestMethod -Uri "$($script:CippApiUrl)/api/ListLicenses?TenantFilter=$([uri]::EscapeDataString($TenantFilter))" -Headers @{ Authorization = "Bearer $($script:CippToken)" }
        $lic =
            if ($resp -is [array]) { $resp }
            elseif ($resp -and ($resp.PSObject.Properties.Name -contains 'Results')) { @($resp.Results) }
            elseif ($resp -and ($resp.PSObject.Properties.Name -contains 'value'))   { @($resp.value) }
            else { @($resp) }
        $out = @()
        foreach ($l in @($lic)) {
            if (-not $l) { continue }
            $name = $null
            foreach ($k in @('License','SkuName','skuPartNumber','Sku','skuId')) { if ($l.PSObject.Properties[$k] -and "$($l.$k)") { $name = "$($l.$k)"; break } }
            if (-not $name) { continue }
            $used = $null; $total = $null
            foreach ($k in @('CountUsed','consumedUnits','Used','TotalUsed')) { if ($l.PSObject.Properties[$k] -and "$($l.$k)" -match '^\d+$') { $used = [int]$l.$k; break } }
            foreach ($k in @('TotalLicenses','prepaidUnits','Total','CountTotal')) { if ($l.PSObject.Properties[$k] -and "$($l.$k)" -match '^\d+$') { $total = [int]$l.$k; break } }
            if ($null -eq $total -and $null -ne $used -and $l.PSObject.Properties['CountAvailable'] -and "$($l.CountAvailable)" -match '^\d+$') { $total = $used + [int]$l.CountAvailable }
            if ($null -ne $total -and ($total -le 0 -or $total -ge 1000)) { continue }
            $over = ($null -ne $used -and $null -ne $total -and $used -gt $total)
            $out += [pscustomobject]@{ Name = $name; Used = $used; Total = $total; Link = ''; Over = $over }
        }
        return ,$out
    } catch { Write-Status "  CIPP license query failed: $($_.Exception.Message)" Warning; return $null }
}

function Get-M365DomainItems {
    # The org's verified custom M365 domains, sourced from the M365 portal (CIPP Graph 'domains'),
    # returned as a single 'Domains' band (New-CardGroup) so each domain renders as its own pill under
    # a Domains header. Excludes any *.onmicrosoft.com domain (the initial tenant domain AND the
    # *.mail.onmicrosoft.com MOERA, which reports isInitial=false) and any unverified / incomplete-setup
    # domain, so only the domains that actually matter (mail/web) are listed. Each pill links to its IT
    # Glue domain record when one exists, else to the M365 admin Domains page; the default domain gets a
    # green ring, and pills carry a DNS-host icon from their live NS records (Cloudflare-portal domains
    # show the Cloudflare logo). Returns @() (no band) when M365 isn't resolved or there are no domains.
    param([pscustomobject]$CippTenant, [string]$OrgId, [string]$LinkBase)
    if (-not ($CippTenant -and $CippTenant.Domain)) { return @() }

    $domains = @()
    try {
        $domains = @(Invoke-CippGraph -TenantFilter $CippTenant.Domain -Endpoint 'domains') |
                   Where-Object { $_.isVerified -and -not $_.isInitial -and "$($_.id)" -notmatch '(?i)\.onmicrosoft\.com$' }
    } catch { Write-Status "  CIPP domains query failed: $($_.Exception.Message)" Warning; return @() }
    if (-not $domains) { return @() }

    $defaultDomain = ''
    $cd = @($domains | Where-Object { $_.isDefault } | Select-Object -First 1)
    if ($cd.Count -gt 0) { $defaultDomain = "$($cd[0].id)".ToLowerInvariant() }

    # IT Glue domain name -> record id, to keep the per-pill IT Glue link where the domain exists there.
    $itgById = @{}
    try {
        foreach ($d in @((Get-ITGlueDomains -filter_organization_id $OrgId).data)) {
            $n = "$($d.attributes.name)".ToLowerInvariant()
            if ($n -and -not $itgById.ContainsKey($n)) { $itgById[$n] = "$($d.id)" }
        }
    } catch { Write-Status "  IT Glue domains query failed: $($_.Exception.Message)" Warning }

    # Default domain first, then the rest alphabetically.
    $sorted = $domains | Sort-Object `
        @{ Expression = { if ("$($_.id)".ToLowerInvariant() -eq $defaultDomain) { 0 } else { 1 } } }, `
        @{ Expression = { "$($_.id)" } }

    $adminBase = 'https://admin.microsoft.com/#/Domains/Details'
    $cfPortal = Get-CloudflarePortalSet   # lazily loads (once per run) the Cloudflare zone set on first use
    $pills = @()
    foreach ($d in $sorted) {
        $name = "$($d.id)"
        $low  = $name.ToLowerInvariant()
        $link = if ($itgById.ContainsKey($low)) { "$LinkBase/domains/$($itgById[$low])" } else { "$adminBase/$name" }
        # M365 default domain -> green outline ring, keeping the pill's normal styling/icon.
        $ring = if ($defaultDomain -and $low -eq $defaultDomain) { '#28A745' } else { '' }
        # Icon: Cloudflare logo for domains on OUR portal cluster, else the DNS-host icon mapped from the
        # live NS records (IT Glue logo as a last resort so every pill still has an icon).
        $brand = ''
        if ($cfPortal[$low]) {
            $brand = 'cloudflare.com'
        } else {
            try {
                $ns = @(Resolve-DnsName -Name $name -Type NS -QuickTimeout -ErrorAction Stop |
                        Where-Object { $_.QueryType -eq 'NS' } | ForEach-Object { $_.NameHost })
                if ($ns.Count -gt 0) { $brand = Get-NameserverBrand -NameHosts $ns }
            } catch { }
            if (-not $brand) { $brand = 'itglue.com' }
        }
        $pills += New-CardItem -Label $name -Link $link -Brand $brand -Ring $ring
    }
    if ($pills.Count -eq 0) { return @() }
    return New-CardGroup -Label 'Domains' -Items $pills
}

function Get-CippMailboxItem {
    # Headline pill for the Email card: total Exchange Online mailboxes, expanding to the same breakdown
    # the M365 admin centre shows - user (licensed) mailboxes, shared mailboxes, distribution lists and
    # Microsoft 365 groups. Mailboxes come from CIPP's ListMailboxes (Exchange Online), which carries a
    # recipientTypeDetails per row; DLs/groups come from Graph. System mailboxes (Discovery/Arbitration/
    # etc.) are excluded - only UserMailbox + SharedMailbox count toward the headline.
    param([string]$TenantFilter)
    try {
        $resp = Invoke-RestMethod -Uri "$($script:CippApiUrl)/api/ListMailboxes?TenantFilter=$([uri]::EscapeDataString($TenantFilter))" -Headers @{ Authorization = "Bearer $($script:CippToken)" }
        $mb =
            if ($resp -is [array]) { $resp }
            elseif ($resp -and ($resp.PSObject.Properties.Name -contains 'Results')) { @($resp.Results) }
            elseif ($resp -and ($resp.PSObject.Properties.Name -contains 'value'))   { @($resp.value) }
            else { @($resp) }
        $mb = @($mb)
        $user   = @($mb | Where-Object { "$($_.recipientTypeDetails)" -eq 'UserMailbox' }).Count
        $shared = @($mb | Where-Object { "$($_.recipientTypeDetails)" -eq 'SharedMailbox' }).Count
        $total  = $user + $shared

        # Distribution lists vs Microsoft 365 (Unified) groups, from Graph. A DL is a mail-enabled,
        # non-security group that is NOT a Unified (M365) group; mail-enabled security groups are excluded.
        $unified = 0; $distro = 0
        try {
            $groups = @(Invoke-CippGraph -TenantFilter $TenantFilter -Endpoint 'groups?$select=groupTypes,mailEnabled,securityEnabled&$top=999')
            $unified = @($groups | Where-Object { $_.groupTypes -contains 'Unified' }).Count
            $distro  = @($groups | Where-Object { $_.mailEnabled -eq $true -and $_.securityEnabled -eq $false -and ($_.groupTypes -notcontains 'Unified') }).Count
        } catch { Write-Status "  CIPP groups query failed: $($_.Exception.Message)" Warning }

        $detail = @(
            New-DetailRow -K 'Licensed mailboxes'   -V "$user"
            New-DetailRow -K 'Shared mailboxes'     -V "$shared"
            New-DetailRow -K 'Distribution lists'   -V "$distro"
            New-DetailRow -K 'Microsoft 365 groups' -V "$unified"
        )
        return New-CardItem -Label "$total mailboxes" -Brand 'microsoft.com' -Detail $detail
    } catch { Write-Status "  CIPP mailboxes query failed: $($_.Exception.Message)" Warning; return $null }
}

function Get-DashM365 {
    # 'Email' card: M365 license pills (live CIPP/GDAP, IT Glue fallback) plus the org's verified
    # custom M365 domains (Get-M365DomainItems). The licenses go in an explicit 'Microsoft 365' band
    # and the card is built -NoBand so the flat domain pills below aren't auto-grouped by their
    # DNS-host brand. (MFA enforcement + AD-sync live on the Identity card.)
    param([pscustomobject]$Entry, [string]$OrgId, [string]$LinkBase, [pscustomobject]$CippTenant, [pscustomobject]$MswTenant)
    # Per-SKU license records {Name; Used; Total; Link; Over}, rolled up below into ONE summary pill.
    $licenseRecs = @()

    if ($CippTenant -and $CippTenant.Domain) {
        $cipRecs = Get-CippLicenseItems -TenantFilter $CippTenant.Domain
        if ($null -ne $cipRecs) { $licenseRecs = @($cipRecs) }
    }
    if ($licenseRecs.Count -eq 0) {
    # Fallback: licenses from the IT Glue Microsoft Licenses flexible asset. The synced trait names
    # vary, so probe common keys per asset.
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

        $name = if ($skuName) { $skuName } else { $a.attributes.name }
        # Over-provisioned (consumed > purchased, e.g. 4/3) -> red outline flag.
        $over = ($null -ne $consumed -and $null -ne $active -and [int]$consumed -gt [int]$active)
        $licenseRecs += [pscustomobject]@{ Name = $name; Used = $consumed; Total = $active; Link = "$LinkBase/assets/$($a.id)"; Over = $over }
    }
    }

    # Roll the per-SKU records into ONE pill: "<assigned>/<total> licenses assigned", expanding to the
    # per-license breakdown (each SKU's assigned/total, the over-provisioned ones flagged red). Licenses
    # missing a count contribute 0 to the totals but still list. The pill flags red if any SKU is over.
    $licenseItem = $null
    if ($licenseRecs.Count -gt 0) {
        $assigned = ($licenseRecs | ForEach-Object { [int]$_.Used }  | Measure-Object -Sum).Sum
        $totalLic = ($licenseRecs | ForEach-Object { [int]$_.Total } | Measure-Object -Sum).Sum
        $anyOver  = [bool](@($licenseRecs | Where-Object { $_.Over }).Count) -or ($assigned -gt $totalLic)
        $detail = @($licenseRecs | Sort-Object Name | ForEach-Object {
            $v = if ($null -ne $_.Used -and $null -ne $_.Total) { "$($_.Used)/$($_.Total)" }
                 elseif ($null -ne $_.Used) { "$($_.Used)" } else { '' }
            New-DetailRow -K $_.Name -V $v -State $(if ($_.Over) { 'bad' } else { '' }) -Link $_.Link
        })
        $licenseItem = New-CardItem -Label "$assigned/$totalLic licenses assigned" -Brand 'microsoft.com' -Detail $detail -DetailHead 'Licenses' -Alert:$anyOver
    }

    $items = @()
    # Microsoft 365 band: the headline mailbox-count pill on top, then the single license-summary pill.
    $m365Items = @()
    if ($CippTenant -and $CippTenant.Domain) {
        $mbItem = Get-CippMailboxItem -TenantFilter $CippTenant.Domain
        if ($mbItem) { $m365Items += $mbItem }
    }
    if ($licenseItem) { $m365Items += $licenseItem }
    if ($m365Items.Count -gt 0) {
        $items += New-CardGroup -Label 'Microsoft 365' -Brand 'microsoft.com' -Items $m365Items
    }
    # SonicWall Cloud App Security (CAS): protected-user count from the client's MySonicWall CSC tile.
    if ($MswTenant) {
        $casCount = Get-MswServiceCount -Tenant $MswTenant -ServiceName 'CAS2.0'
        if ($null -ne $casCount) {
            $items += New-CardGroup -Label 'Cloud App Security' -Brand 'cloudappsecurity' -Items @(New-CardItem -Label "$casCount users")
        }
    }
    $items += @(Get-M365DomainItems -CippTenant $CippTenant -OrgId $OrgId -LinkBase $LinkBase)
    return New-Card -Title 'Email' -Items $items -NoBand
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
            $cwmItems = @()
            # Agreement on top: PeopleFirst member seats (CBT-PF-MEMBER qty on the PeopleFirst Support
            # agreement). Skipped entirely if the company has no such agreement.
            $memQty = Get-CwmMemberQty -CompanyId $co.Id
            if ($null -ne $memQty) { $cwmItems += New-CardItem -Label "$memQty members" -Brand 'connectwise.com' }

            # Then CWM contacts (links to the company's billing contact / company record).
            $cw = Get-CwmActiveContactCount -CompanyId $co.Id
            $base = "https://$($Creds.CWM.ConnectionInfo.Server)/v4_6_release/services/system_io/router/openrecord.rails"
            $billId = $null
            try { $billId = (Get-CWMCompany -id $co.Id).billingContact.id } catch {}
            $cwHref = if ($billId) { "$base`?recordType=ContactFV&recid=$billId" } else { "$base`?recordType=CompanyFV&recid=$($co.Id)" }
            $cwLabel = if ($null -ne $cw) { "$cw contacts" } else { "n/a contacts" }
            $cwmItems += New-CardItem -Label $cwLabel -Link $cwHref -Brand 'connectwise.com'

            # Multiple CWM pills -> cluster them in a 'ConnectWise Manage' sub-box (same pattern as the
            # Veeam / Microsoft 365 groups); a single pill stays inline.
            if ($cwmItems.Count -ge 2) { $items += New-CardGroup -Label 'ConnectWise Manage' -Brand 'connectwise.com' -Items $cwmItems }
            else { $items += $cwmItems }
        }
    }
    # IT Glue contacts (no active/inactive concept in IT Glue -> total contacts for the org).
    $itg = Get-ItgContactCount -OrgId $OrgId
    $itgLabel = if ($null -ne $itg) { "$itg contacts" } else { "n/a contacts" }
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

function Get-AutomateDcItem {
    # On-prem AD pill(s): one pill per local AD domain (e.g. "sscv.local"), each EXPANDING to list that
    # domain's domain controllers. Each DC in the panel is a clickable link straight to its own page in
    # ConnectWise Automate. ConnectWise Automate tags a domain controller by prefixing its agent
    # DomainName with 'DC:' (members show the bare domain, e.g. 'DC:sscv.local' vs 'sscv.local'), so DCs
    # are the computers whose DomainName starts with 'DC:' and the domain is that value with the prefix
    # stripped. Returns an array of pills (one per domain), or $null when no Automate match / no DC.
    param([pscustomobject]$Entry, [string]$OrgName)
    $autoId = Resolve-AutomateClient -Entry $Entry -OrgName $OrgName
    if (-not $autoId) { return $null }
    try {
        $cond = "Client.Id = $autoId"
        $raw  = Invoke-RestMethod -Uri "$($script:AutomateUrl)/cwa/api/v1/computers?condition=$([uri]::EscapeDataString($cond))&pagesize=1000" -Headers $script:AutomateHeaders
    } catch { Write-Status "  Automate computers query failed: $($_.Exception.Message)" Warning; return $null }
    # The computers endpoint sometimes returns the array nested one level; flatten defensively.
    $comps = New-Object System.Collections.Generic.List[object]
    foreach ($x in $raw) { if ($x -is [array]) { $comps.AddRange($x) } else { $comps.Add($x) } }
    $dcs = @($comps | Where-Object { "$($_.DomainName)" -match '^DC:' })
    if ($dcs.Count -eq 0) { return $null }

    # Group the DCs by their AD domain (most clients have one; handle multiple defensively). Each DC links
    # to its own Automate computer page (browse SPA) so a tech lands on that device, not the company list.
    $byDomain = @{}
    foreach ($dc in $dcs) {
        $d = ($dc.DomainName -replace '^DC:', '').Trim()
        if (-not $d) { continue }
        if (-not $byDomain.ContainsKey($d)) { $byDomain[$d] = New-Object System.Collections.Generic.List[object] }
        $byDomain[$d].Add($dc)
    }
    $compLink = "$($script:AutomateUrl)/automate/browse/companies/computers?companyId=$autoId"
    $pills = @()
    foreach ($d in ($byDomain.Keys | Sort-Object)) {
        $rows = @(New-DetailRow -K 'Domain controllers' -V "$($byDomain[$d].Count)") + @($byDomain[$d] |
            Sort-Object { "$($_.ComputerName)" } |
            ForEach-Object { New-DetailRow -V "$($_.ComputerName)" -Link "$($script:AutomateUrl)/automate/browse/computers/$($_.Id)" })
        $pills += New-CardItem -Label $d -Brand 'winserver' -Muted `
            -Detail $rows -Link $compLink -LinkText 'Open in Automate'
    }
    return $pills
}

function Get-AutomateDuoLogonItem {
    # Endpoint-Security pill: how many of the client's devices have the "Duo Authentication for
    # Windows Logon" agent installed, e.g. "Duo Logon: 57 of 63 devices". A coverage gap (< 100%) is
    # flagged with a red outline. Sourced from ConnectWise Automate software inventory because Duo's
    # own API does not report which machines have the logon agent installed. Returns $null when the
    # org has no Automate match or no computers.
    #
    # Automate query shapes were validated live against this instance:
    #   - The computers endpoint condition does NOT support the Applications child collection (400),
    #     so the GUI's "[Computer.Applications.Name] Contains ..." search can't be issued directly.
    #   - The software endpoint DOES accept a flat condition on its own columns; each row carries a
    #     ClientId + ComputerId, so we filter by client and name in one call and count distinct
    #     computers. (Note: the software endpoint uses the flat 'ClientId' field, whereas the
    #     computers endpoint uses the 'Client.Id' navigation field.)
    param([pscustomobject]$Entry, [string]$OrgName)
    $autoId = Resolve-AutomateClient -Entry $Entry -OrgName $OrgName
    if (-not $autoId) { return $null }

    # Denominator: the client's total computers (same call shape as Get-AutomateDcItem).
    try {
        $cond = "Client.Id = $autoId"
        $rawComps = Invoke-RestMethod -Uri "$($script:AutomateUrl)/cwa/api/v1/computers?condition=$([uri]::EscapeDataString($cond))&pagesize=1000" -Headers $script:AutomateHeaders
    } catch { Write-Status "  Automate Duo-logon computers query failed: $($_.Exception.Message)" Warning; return $null }
    $comps = New-Object System.Collections.Generic.List[object]
    foreach ($x in $rawComps) { if ($x -is [array]) { $comps.AddRange($x) } else { $comps.Add($x) } }
    $total = $comps.Count
    if ($total -eq 0) { return $null }

    # Numerator: distinct computers whose installed software name contains 'Duo Authentication'.
    try {
        $swCond = "ClientId = $autoId and Name contains 'Duo Authentication'"
        $rawSw  = Invoke-RestMethod -Uri "$($script:AutomateUrl)/cwa/api/v1/Computers/Software?condition=$([uri]::EscapeDataString($swCond))&pagesize=1000" -Headers $script:AutomateHeaders
    } catch { Write-Status "  Automate Duo-logon software query failed: $($_.Exception.Message)" Warning; return $null }
    $sw = New-Object System.Collections.Generic.List[object]
    foreach ($x in $rawSw) { if ($x -is [array]) { $sw.AddRange($x) } else { $sw.Add($x) } }
    $coveredIds = @($sw | ForEach-Object { "$($_.ComputerId)" } | Where-Object { $_ } | Sort-Object -Unique)
    $covered = $coveredIds.Count

    # The uncovered devices (in the client's fleet but without the Duo Logon agent) become the expandable
    # pill's panel, so the gap is one click away. Computers carry .Id; the software rows carry .ComputerId.
    $missing = @($comps | Where-Object { "$($_.Id)" -and ($coveredIds -notcontains "$($_.Id)") } |
        ForEach-Object { "$($_.ComputerName)" } | Where-Object { $_ } | Sort-Object -Unique)
    $detail = @(New-DetailRow -K 'Missing Duo Logon agent' -V "$($missing.Count)") + @($missing | ForEach-Object { New-DetailRow -V $_ })

    $link = "$($script:AutomateUrl)/automate/browse/companies/computers?companyId=$autoId"
    return New-CardItem -Label "Duo Logon: $covered of $total devices" -Link $link -LinkText 'Open in Automate' `
        -Detail $detail -Brand 'duo.com' -Alert:($covered -lt $total)
}

function Get-ConfigCountPill {
    # A pill card-item "<Label>: <count>". Links to the client's ConnectWise Automate computers page
    # when -Link is supplied (Automate icon via -Brand); otherwise to the IT Glue configs list
    # filtered by type. (Future: append a stale count, e.g. devices not seen in Automate for 90+ days.)
    param([string]$Label, [string]$TypeName, [string]$OrgId, [string]$LinkBase, [string]$TypeId,
          [string]$StatusId, [string]$Link = '', [string]$Brand = '')
    $noun = $Label.ToLowerInvariant()
    if (-not $TypeId -or -not $StatusId) { return New-CardItem -Label "n/a $noun" }
    $count = $null
    try {
        $count = [int](Get-ITGlueConfigurations -filter_organization_id $OrgId -filter_configuration_type_id $TypeId `
            -filter_configuration_status_id $StatusId -page_size 1).meta.'total-count'
    } catch { Write-Status "  IT Glue $Label query failed: $($_.Exception.Message)" Warning }
    if (-not $Link) {
        # Fallback: IT Glue UI filter deep-link (#partial=...&filters=[Type:<name>]).
        $link = "$LinkBase/configurations#partial=&sortBy=name:asc&filters=%5BType:$([Uri]::EscapeDataString($TypeName))%5D"
    } else { $link = $Link }
    return New-CardItem -Label "$([int]$count) $noun" -Link $link -Brand $Brand -Muted
}

function Get-AutomateComputerSet {
    # All ConnectWise Automate computers for a client (flattened, cached per-run by client Id so the
    # Workstations and Servers pills share a single fetch). Returns @() on error / no match.
    param([string]$AutoId)
    if (-not $script:AutomateConnected -or -not $AutoId) { return @() }
    if (-not $script:AutomateCompCache) { $script:AutomateCompCache = @{} }
    if ($script:AutomateCompCache.ContainsKey($AutoId)) { return $script:AutomateCompCache[$AutoId] }
    try {
        $cond = "Client.Id = $AutoId"
        $raw  = Invoke-RestMethod -Uri "$($script:AutomateUrl)/cwa/api/v1/computers?condition=$([uri]::EscapeDataString($cond))&pagesize=1000" -Headers $script:AutomateHeaders
    } catch { Write-Status "  Automate computers query failed: $($_.Exception.Message)" Warning; return @() }
    # The computers endpoint sometimes returns the array nested one level; flatten defensively.
    $comps = New-Object System.Collections.Generic.List[object]
    foreach ($x in $raw) { if ($x -is [array]) { $comps.AddRange($x) } else { $comps.Add($x) } }
    $arr = $comps.ToArray()
    $script:AutomateCompCache[$AutoId] = $arr
    return $arr
}

function Get-CompLastContact {
    # The most recent check-in time for an Automate computer, as [datetime] (local), or $null. Automate
    # exposes several near-identical timestamps; we take the latest of the agent-contact fields so a
    # machine isn't called stale just because one field lags.
    param([pscustomobject]$Comp)
    $best = $null
    foreach ($f in 'LastContact','RemoteAgentLastContact','LastHeartbeat') {
        $p = $Comp.PSObject.Properties[$f]
        if (-not $p -or -not $p.Value) { continue }
        $dt = [datetime]::MinValue
        if ([datetime]::TryParse("$($p.Value)", [ref]$dt)) { if (-not $best -or $dt -gt $best) { $best = $dt } }
    }
    return $best
}

function Format-Age {
    # A compact "time since" label for a [timespan]: e.g. 90d, 14h, 5m. Days once past 24h, else hours,
    # else minutes (floored, never below 1m so a just-offline device doesn't read "0").
    param([timespan]$Span)
    if ($Span.TotalDays    -ge 1) { return "$([int][math]::Floor($Span.TotalDays))d" }
    if ($Span.TotalHours   -ge 1) { return "$([int][math]::Floor($Span.TotalHours))h" }
    return "$([math]::Max(1,[int][math]::Floor($Span.TotalMinutes)))m"
}

function Get-AutomateDriveUsage {
    # Used/total GB across a computer's non-removable (internal, present) drives, via the drives endpoint.
    # Sizes come back in MB. Returns @{ UsedGB; TotalGB } or $null when the call fails / no internal drive.
    param([string]$CompId)
    try {
        $raw = Invoke-RestMethod -Uri "$($script:AutomateUrl)/cwa/api/v1/computers/$CompId/drives" -Headers $script:AutomateHeaders
    } catch { return $null }
    $drv = New-Object System.Collections.Generic.List[object]
    foreach ($x in $raw) { if ($x -is [array]) { $drv.AddRange($x) } else { $drv.Add($x) } }
    $internal = @($drv | Where-Object { ("$($_.IsInternal)" -eq 'True') -and ("$($_.IsMissing)" -ne 'True') -and ([double]("0" + "$($_.Size)") -gt 0) })
    if ($internal.Count -eq 0) { return $null }
    $sizeMb = ($internal | Measure-Object -Property Size -Sum).Sum
    $freeMb = ($internal | Measure-Object -Property FreeSpace -Sum).Sum
    return @{ UsedGB = [int][math]::Round(($sizeMb - $freeMb) / 1024); TotalGB = [int][math]::Round($sizeMb / 1024) }
}

function Get-AutomateWorkstationPill {
    # Workstations pill, Automate-sourced: "<n> workstations" and, when any haven't checked in for
    # $script:StaleDays, a "· <s> stale" tail with a red ring. Expands to the stale device names, each a
    # direct link to that machine's Automate page. Returns $null when the client has no Automate workstations.
    param([object[]]$Comps, [string]$AutoLink)
    $ws = @($Comps | Where-Object { "$($_.Type)" -eq 'Workstation' })
    if ($ws.Count -eq 0) { return $null }
    $cutoff = (Get-Date).AddDays(-$script:StaleDays)
    $stale  = @($ws | Where-Object { $lc = Get-CompLastContact $_; (-not $lc) -or ($lc -lt $cutoff) })
    $label  = if ($stale.Count -gt 0) { "$($ws.Count) workstations · $($stale.Count) stale" } else { "$($ws.Count) workstations" }
    if ($stale.Count -eq 0) {
        return New-CardItem -Label $label -Link $AutoLink -LinkText 'Open in Automate' -Brand 'automate' -Muted
    }
    $rows = @(New-DetailRow -K "Stale (>$($script:StaleDays)d)" -V "$($stale.Count)") + @($stale |
        Sort-Object { Get-CompLastContact $_ } |
        ForEach-Object {
            $lc = Get-CompLastContact $_
            $age = if ($lc) { Format-Age ((Get-Date) - $lc) } else { 'never' }
            New-DetailRow -V "$($_.ComputerName) ($age)" -Link "$($script:AutomateUrl)/automate/browse/computers/$($_.Id)" })
    # Amber (warning) ring, not red: a stale workstation is softer than an offline server (which gets the
    # red error ring), so the two read as distinct severities at a glance.
    return New-CardItem -Label $label -Brand 'automate' -Ring '#E0A800' `
        -Detail $rows -Link $AutoLink -LinkText 'Open in Automate'
}

function Get-AutomateServerPill {
    # Servers pill, Automate-sourced: "<n> servers" with a red ring (and "· <o> offline" tail) when any
    # server is offline. Expands to every server with its used/total GB across non-removable drives; an
    # offline server shows how long it's been offline (red) instead. Each row links to the machine's
    # Automate page. Returns $null when the client has no Automate servers.
    param([object[]]$Comps, [string]$AutoLink)
    $sv = @($Comps | Where-Object { "$($_.Type)" -eq 'Server' })
    if ($sv.Count -eq 0) { return $null }
    $offline = @($sv | Where-Object { "$($_.Status)" -ne 'Online' })
    $label = if ($offline.Count -gt 0) { "$($sv.Count) servers · $($offline.Count) offline" } else { "$($sv.Count) servers" }
    $rows = @($sv |
        Sort-Object { "$($_.ComputerName)" } |
        ForEach-Object {
            $isOff = "$($_.Status)" -ne 'Online'
            $link  = "$($script:AutomateUrl)/automate/browse/computers/$($_.Id)"
            if ($isOff) {
                $lc = Get-CompLastContact $_
                $v  = if ($lc) { "offline $(Format-Age ((Get-Date) - $lc))" } else { 'offline' }
                New-DetailRow -K "$($_.ComputerName)" -V $v -State 'bad' -Link $link
            } else {
                $u = Get-AutomateDriveUsage -CompId "$($_.Id)"
                $v = if ($u) { "$($u.UsedGB) / $($u.TotalGB) GB" } else { 'n/a' }
                New-DetailRow -K "$($_.ComputerName)" -V $v -Link $link
            }
        })
    return New-CardItem -Label $label -Brand 'automate' -Alert:($offline.Count -gt 0) `
        -Detail $rows -DetailHead 'Servers' -Link $AutoLink -LinkText 'Open in Automate'
}

function Get-DashDevices {
    # 'Devices' card: a ConnectWise Automate band plus, when CIPP is connected, a Microsoft 365 band with
    # Intune (Computers/Mobile) and Entra join-type device counts. When the org matches an Automate client,
    # the Workstations/Servers pills are Automate-sourced and live (workstation staleness, server
    # offline/drive usage, each expanding to per-device rows that deep-link to the machine's Automate page);
    # otherwise they fall back to active-IT-Glue-config counts that link to the org's Automate computer list.
    param([pscustomobject]$Entry, [string]$OrgId, [string]$OrgName, [string]$LinkBase, [string]$WsTypeId, [string]$SvTypeId, [string]$StatusId, [pscustomobject]$CippTenant)
    $autoLink = ''; $brand = ''
    $autoId = Resolve-AutomateClient -Entry $Entry -OrgName $OrgName
    if ($autoId) { $autoLink = "$($script:AutomateUrl)/automate/browse/companies/computers?companyId=$autoId"; $brand = 'automate' }

    $wsPill = $null; $svPill = $null
    if ($autoId) {
        $comps  = Get-AutomateComputerSet -AutoId $autoId
        $wsPill = Get-AutomateWorkstationPill -Comps $comps -AutoLink $autoLink
        $svPill = Get-AutomateServerPill      -Comps $comps -AutoLink $autoLink
    }
    # Fall back to the IT Glue active-config count when Automate has no match / no machine of that type.
    if (-not $wsPill) { $wsPill = Get-ConfigCountPill -Label 'Workstations' -TypeName 'Managed Workstation' -OrgId $OrgId -LinkBase $LinkBase -TypeId $WsTypeId -StatusId $StatusId -Link $autoLink -Brand $brand }
    if (-not $svPill) { $svPill = Get-ConfigCountPill -Label 'Servers'      -TypeName 'Managed Server'      -OrgId $OrgId -LinkBase $LinkBase -TypeId $SvTypeId -StatusId $StatusId -Link $autoLink -Brand $brand }
    $items = @($wsPill, $svPill)
    # Microsoft 365 / Intune device counts (auto-banded under a 'Microsoft 365' strip via their brand).
    if ($CippTenant -and $CippTenant.Domain) {
        $items += @(Get-CippDeviceItems -TenantFilter $CippTenant.Domain)
    }
    return New-Card -Title 'Devices' -Items $items
}

# ==============================================================================
# SonicWall (MySonicWall API) + WatchGuard Cloud (READ-ONLY)
# ------------------------------------------------------------------------------
# MySonicWall: one X-api-key call (get-cloud-tenants) yields the MSP-wide tenant list + each tenant's
# Capture Client / CAS active counts; get-firewalls (under our own product group) yields the whole
# registered hardware fleet, joined to a client by friendlyName; serviceInfo gives per-firewall expiry.
# WatchGuard: OAuth + the allocation 'firebox' summary yields each Firebox's model/serial/expiry,
# also joined by friendlyName. Online/offline isn't exposed by either (would need NSM / WG monitoring).
# ==============================================================================
function ConvertTo-IntSafe {
    # Parse a possibly-empty / negative integer string to [int]; default when not a clean integer.
    param([string]$Value, [int]$Default = 0)
    $n = 0
    if ([int]::TryParse(("$Value").Trim(), [ref]$n)) { return $n }
    return $Default
}

function Get-MswTenants {
    # get-cloud-tenants -> the MSP-wide tenant list (cached for the run). Also records the account
    # userName (required by get-firewalls) and auto-detects the product group holding the registered
    # firewall fleet (our own MySonicWall account = by far the most products), unless one is configured.
    if ($null -ne $script:MswTenants) { return $script:MswTenants }
    $r = Invoke-RestMethod -Uri "$($script:MswBase)/api/hgms/get-cloud-tenants" -Headers $script:MswHeaders
    $c = $r.content
    $script:MswUser = "$($c.userName)"
    $script:MswTenants = @($c.arrTenants)
    if (-not $script:MswFirewallGroupId) {
        $top = $script:MswTenants | Sort-Object { [int]("0" + "$($_.productCount)") } -Descending | Select-Object -First 1
        if ($top) { $script:MswFirewallGroupId = "$($top.productGroupID)" }
    }
    return $script:MswTenants
}

function Get-MswServiceCount {
    # The single integer a tenant's CSC tile shows for a cloud service (Capture Client active endpoints,
    # CAS protected users), from cloudServices.avaiableRatio. $null when absent or unlicensed ('-1').
    param([pscustomobject]$Tenant, [string]$ServiceName)
    if (-not $Tenant) { return $null }
    $svc = @($Tenant.cloudServices) | Where-Object { $_.serviceName -eq $ServiceName } | Select-Object -First 1
    if (-not $svc) { return $null }
    $v = "$($svc.avaiableRatio)".Trim()
    if ($v -notmatch '^\d+$') { return $null }
    return [int]$v
}

function Resolve-MswTenant {
    # Match the org to a MySonicWall tenant by normalized name (cache under entry.sonicwall). Returns the
    # LIVE tenant object (with cloudServices) so callers can read the Capture Client / CAS counts.
    param([pscustomobject]$Entry, [string]$OrgName)
    if (-not $script:MswConnected) { return $null }
    if ($Entry._unmanaged -contains 'sonicwall') { return $null }
    $tenants = @(Get-MswTenants)
    $cands = @($tenants | ForEach-Object { [pscustomobject]@{ Id = "$($_.productGroupID)"; Name = "$($_.name)" } })
    $chosen = Resolve-OrgServiceId -Entry $Entry -VendorKey 'sonicwall' -VendorLabel 'SonicWall' -OrgName $OrgName -Candidates $cands
    if (-not $chosen) { return $null }
    return $tenants | Where-Object { "$($_.productGroupID)" -eq "$($chosen.Id)" } | Select-Object -First 1
}

function Get-MswFirewalls {
    # The full registered firewall fleet (hardware only), cached for the run. friendlyName carries the
    # client identity. Cloud-only product rows (Capture Client / CAS instances) are excluded.
    if ($null -ne $script:MswFleet) { return $script:MswFleet }
    Get-MswTenants | Out-Null   # ensure $script:MswUser + $script:MswFirewallGroupId are set
    $script:MswFleet = @()
    if (-not $script:MswFirewallGroupId -or -not $script:MswUser) { return $script:MswFleet }
    try {
        $uri = "$($script:MswBase)/api/hgms/get-firewalls?productGroupId=$($script:MswFirewallGroupId)&userName=$([uri]::EscapeDataString($script:MswUser))"
        $r = Invoke-RestMethod -Uri $uri -Headers $script:MswHeaders
        $rows = @($r.content.rows) | Where-Object { "$($_.model)" -and ("$($_.model)" -notmatch '(?i)capture client|^CAS$') }
        $script:MswFleet = @($rows | ForEach-Object {
            [pscustomobject]@{
                Id = "$($_.serialNumber)"; Name = "$($_.friendlyName)"
                Model = (("$($_.model)" -replace '<[^>]+>', '').Trim())
                Firmware = "$($_.firmwareVersion)"; LicenseExpired = [bool]$_.licenseExpired
            }
        })
    } catch { Write-Status "  SonicWall get-firewalls failed: $($_.Exception.Message)" Warning }
    return $script:MswFleet
}

function Get-MswFirewallExpiry {
    # Soonest expiry among a firewall's ACTIVATED services (serviceInfo). @{ Date; Days; Soon } or $null.
    param([string]$Serial)
    try {
        $si = Invoke-RestMethod -Uri "$($script:MswBase)/api/product/serviceInfo?serial=$([uri]::EscapeDataString($Serial))" -Headers $script:MswHeaders
        $active = foreach ($t in $si.serviceTypeList) {
            foreach ($s in $t.serviceList) {
                if ("$($s.serviceStatus)" -eq 'ACTIVATED' -and "$($s.expirydate)".Trim()) { $s }
            }
        }
        $soonest = $active | Sort-Object { ConvertTo-IntSafe "$($_.nDaysToExpiry)" 999999 } | Select-Object -First 1
        if ($soonest) {
            return @{ Date = "$($soonest.expirydate)".Trim(); Days = (ConvertTo-IntSafe "$($soonest.nDaysToExpiry)" 999999); Soon = ("$($soonest.isSoonExpiring)" -eq 'YES' -or "$($soonest.expired)" -eq 'YES') }
        }
    } catch { }
    return $null
}

function Get-WgFireboxes {
    # All allocated Fireboxes (cached for the run), normalized to Id=serial, Name=friendlyName, Model,
    # Mac, ExpiryDate, Days. friendlyName is the only company identifier the WG API exposes (see notes).
    if ($null -ne $script:WgFleet) { return $script:WgFleet }
    $script:WgFleet = @()
    $all = @(); $offset = 0; $page = 500
    try {
        do {
            $url = "$($script:WgApiBase)/platform/allocation/v2/$($script:WgAccountId)/assets/summary/firebox?allocationStatus=allocated&limit=$page&offset=$offset"
            $resp = Invoke-RestMethod -Uri $url -Headers $script:WgHeaders
            $batch = @($resp.data); $all += $batch
            $total = [int]$resp.pageControls.totalItems; $offset += $page
        } while ($offset -lt $total -and $batch.Count -gt 0)
    } catch { Write-Status "  WatchGuard firebox query failed: $($_.Exception.Message)" Warning; return $script:WgFleet }
    $script:WgFleet = @($all | ForEach-Object {
        [pscustomobject]@{
            Id = "$($_.serialOrLicense)"; Name = "$($_.friendlyName)"; Model = "$($_.modelName)"
            Mac = "$($_.macAddress)"; ExpiryDate = "$($_.licenseExpiryDate)"; Days = (ConvertTo-IntSafe "$($_.daysUntilExpiry)")
        }
    })
    return $script:WgFleet
}

function Resolve-FirewallSet {
    # Returns the firewall object(s) for this org (one pill each), or @() if confirmed none. Generic over
    # SonicWall/WatchGuard: their APIs only identify a firewall's owner by friendlyName, so we confirm
    # once via a fuzzy-ranked multi-select pick-list and cache the chosen serials under $Entry.$VendorKey;
    # each run we rehydrate against the live fleet so labels/expiry stay fresh (mirrors Resolve-UnifiSites).
    param([pscustomobject]$Entry, [string]$OrgName, [object[]]$Devices, [string]$VendorKey, [string]$Label)
    if ($Entry._unmanaged -contains $VendorKey) { return @() }
    $cached = $Entry.PSObject.Properties[$VendorKey]
    if ($cached -and $cached.Value) {
        $ids = @(@($cached.Value) | ForEach-Object { "$($_.Id)" })
        $hit = @($Devices | Where-Object { $ids -contains "$($_.Id)" })
        if ($hit.Count -gt 0) { return $hit }
        return @($cached.Value)   # live fleet unavailable this run; trust the cache
    }
    if (-not $Devices -or $Devices.Count -eq 0) { return @() }

    $target = ConvertTo-NormalizedName $OrgName
    $scored = @($Devices | ForEach-Object {
        $n = ConvertTo-NormalizedName $_.Name
        $score = 0
        if ($n) {
            if ($n -eq $target) { $score += 1000 }
            elseif ($target -like "$n*" -or $n -like "$target*") { $score += 100 }
        }
        $score += @(($target -split ' ') | Where-Object { $_ -and (($n -split ' ') -contains $_) }).Count
        [pscustomobject]@{ Dev = $_; Score = $score }
    } | Sort-Object -Property Score -Descending)

    $showAll = $false
    while ($true) {
        $list = if ($showAll) { @($scored) } else { @($scored | Select-Object -First 12) }
        Write-Host "`n  [$Label] Select firewall(s) for '$OrgName' (e.g. 1,3 - multiple allowed):" -ForegroundColor Yellow
        for ($i = 0; $i -lt $list.Count; $i++) { Write-Host ("    [{0}] {1}  ({2})" -f ($i + 1), $list[$i].Dev.Name, $list[$i].Dev.Model) }
        if (-not $showAll -and $scored.Count -gt $list.Count) { Write-Host "    [a] show ALL $($scored.Count)" }
        Write-Host "    [0] none / not in $Label"
        $sel = Read-Host "  Choice"
        if ($sel -match '^[Aa]$') { $showAll = $true; continue }
        $tokens = @($sel -split '[,\s]+' | Where-Object { $_ -ne '' })
        if ($tokens.Count -eq 0) { continue }
        if ($tokens -contains '0') { Add-Unmanaged -Entry $Entry -VendorKey $VendorKey; return @() }
        $picked = @(); $bad = $false
        foreach ($t in $tokens) { $k = -1; [int]::TryParse($t, [ref]$k) | Out-Null; if ($k -ge 1 -and $k -le $list.Count) { $picked += $list[$k - 1].Dev } else { $bad = $true } }
        if ($bad -or $picked.Count -eq 0) { Write-Host "  Invalid selection." -ForegroundColor Red; continue }
        $picked = @($picked | Sort-Object -Property Id -Unique)
        $val = @($picked | ForEach-Object { [pscustomobject]@{ Id = "$($_.Id)"; Name = $_.Name; Model = $_.Model } })
        $Entry | Add-Member -NotePropertyName $VendorKey -NotePropertyValue $val -Force
        $script:ServiceMapDirty = $true
        return $picked
    }
}

# ---- Network (SonicWall + WatchGuard firewalls + UniFi Site Manager) ----
function Get-ItgConfigs {
    # All configurations for an org of a given type + status as FULL objects (not just a count, which
    # is what Get-ConfigCountPill returns). Paged like Get-ItgAllOrganizations. ARCHIVED configs are
    # excluded client-side: IT Glue's 'archived' flag is independent of configuration status, so an
    # archived item can still be status 'Active' and would otherwise slip through the status filter.
    param([string]$OrgId, [string]$TypeId, [string]$StatusId)
    $all = @(); $page = 1
    do {
        $r = Get-ITGlueConfigurations -filter_organization_id $OrgId `
            -filter_configuration_type_id $TypeId -filter_configuration_status_id $StatusId `
            -page_size 1000 -page_number $page
        if ($r.data) { $all += $r.data }
        $page++
    } while ($r.data -and $r.data.Count -eq 1000)
    return @($all | Where-Object { -not $_.attributes.archived })
}

function ConvertTo-SerialKey {
    # Normalise a serial for matching: lowercased with all non-alphanumerics stripped. WatchGuard Cloud
    # renders a serial with a hyphen before the last 4 chars while the sticker on the box has none, so the
    # IT Glue config and the API serial differ only by punctuation - stripping it lets both forms match.
    param([string]$S)
    return ("$S" -replace '[^A-Za-z0-9]', '').ToLowerInvariant()
}

function Test-WanIp {
    # $true if $Ip is a valid PUBLIC IPv4 (a WAN address): not RFC1918 private, loopback, link-local,
    # 0.0.0.0/8, or multicast/reserved (>=224). CGNAT (100.64/10) is intentionally allowed as WAN.
    param([string]$Ip)
    if ($Ip -notmatch '^\d{1,3}(\.\d{1,3}){3}$') { return $false }
    $o = @($Ip -split '\.' | ForEach-Object { [int]$_ })
    if (@($o | Where-Object { $_ -gt 255 }).Count) { return $false }
    if ($o[0] -eq 0 -or $o[0] -eq 127 -or $o[0] -ge 224) { return $false }
    if ($o[0] -eq 10) { return $false }
    if ($o[0] -eq 172 -and $o[1] -ge 16 -and $o[1] -le 31) { return $false }
    if ($o[0] -eq 192 -and $o[1] -eq 168) { return $false }
    if ($o[0] -eq 169 -and $o[1] -eq 254) { return $false }
    return $true
}

function Get-WanCandidate {
    # Parse a value that may be an IP or "IP:port" (an interface IP, a primary-ip, or text in Notes) and
    # return @{ Ip; Port } when the IP is public/WAN, else $null. A colon yields a port ONLY when it sits
    # between an IPv4 and digits (interface NAMES like X1 / 0 never reach here as ports).
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $m = [regex]::Match($Value, '(\d{1,3}(?:\.\d{1,3}){3})(?::(\d{1,5}))?')
    if (-not $m.Success) { return $null }
    if (-not (Test-WanIp $m.Groups[1].Value)) { return $null }
    $port = if ($m.Groups[2].Success) { [int]$m.Groups[2].Value } else { $null }
    return @{ Ip = $m.Groups[1].Value; Port = $port }
}

function Get-ItgConfigIndex {
    # Active configs for the org, for firewall->IT Glue matching: @{ BySerial; List } (serial-number,
    # lowercased, -> config; plus the raw list for a name fallback). One paged call, cached per run/org.
    param([string]$OrgId, [string]$StatusId)
    if ($script:CfgIndex -and $script:CfgIndexOrg -eq $OrgId) { return $script:CfgIndex }
    $all = @(); $page = 1
    do {
        $p = @{ filter_organization_id = $OrgId; page_size = 1000; page_number = $page }
        if ($StatusId) { $p.filter_configuration_status_id = $StatusId }
        try { $r = Get-ITGlueConfigurations @p } catch { Write-Status "  IT Glue configs query failed: $($_.Exception.Message)" Warning; break }
        if ($r.data) { $all += $r.data }
        $page++
    } while ($r.data -and $r.data.Count -eq 1000)
    $all = @($all | Where-Object { -not $_.attributes.archived })
    $bySerial = @{}
    foreach ($c in $all) {
        $s = ConvertTo-SerialKey $c.attributes.'serial-number'
        if ($s -and -not $bySerial.ContainsKey($s)) { $bySerial[$s] = $c }
    }
    $script:CfgIndexOrg = $OrgId
    $script:CfgIndex = [pscustomobject]@{ BySerial = $bySerial; List = $all }
    return $script:CfgIndex
}

function Find-ItgConfigForFirewall {
    # The IT Glue config for a firewall: matched by serial first, then a normalized-name fallback.
    param([pscustomobject]$Index, [string]$Serial, [string]$Name)
    $s = ConvertTo-SerialKey $Serial
    if ($s -and $Index.BySerial.ContainsKey($s)) { return $Index.BySerial[$s] }
    $tn = ConvertTo-NormalizedName $Name
    if ($tn) {
        $hit = $Index.List | Where-Object { (ConvertTo-NormalizedName "$($_.attributes.name)") -eq $tn } | Select-Object -First 1
        if ($hit) { return $hit }
    }
    return $null
}

function Get-IfaceWanRank {
    # Order a config interface as a WAN candidate: 0 = looks like the primary WAN (name/notes/port say
    # WAN or 'primary', SonicWall X1, WatchGuard port 0), 1 = a secondary WAN, 2 = anything else. Used to
    # surface real WAN interfaces ahead of incidental ones when collecting a firewall's WAN IPs.
    param([pscustomobject]$Iface, [string]$Vendor)
    $nm = "$($Iface.attributes.name)"; $nt = "$($Iface.attributes.notes)"; $pt = "$($Iface.attributes.port)"
    if (($nm -match '(?i)wan') -or ($nt -match '(?i)\bwan\b|primary') -or ($pt -match '(?i)wan') -or
        ($Vendor -eq 'sonicwall' -and $nm -match '(?i)^x1$') -or
        ($Vendor -eq 'watchguard' -and (($nm -match '(?i)^(0|eth0|if0)$') -or ($pt -match '^0(\b|/)')))) { return 0 }
    if ($nt -match '(?i)secondary') { return 1 }
    return 2
}

function Resolve-FirewallWanIps {
    # ALL of a firewall's WAN IPs (+admin port each) from its IT Glue config, in priority order and deduped.
    # Sources, in order: (1) the config Primary IP; (2) config interfaces, WAN/primary-ranked first then
    # secondary then the rest; (3) IPs parsed out of the Notes field (CW-synced) - this is how a 2nd WAN
    # often surfaces when it isn't a separate interface. Only PUBLIC IPv4s are kept. Returns an array of
    # @{ Ip; Port } (empty if none). Port: from an explicit "IP:port", else WatchGuard 8080 / SonicWall 443.
    param([pscustomobject]$Config, [ValidateSet('watchguard','sonicwall')][string]$Vendor)
    if (-not $Config) { return @() }
    $a = $Config.attributes
    $raw = @("$($a.'primary-ip')")
    $ifs = @()
    try { $ifs = @((Get-ITGlueConfigurationInterfaces -conf_id $Config.id).data) } catch {}
    foreach ($rank in 0, 1, 2) {
        foreach ($i in $ifs) { if ((Get-IfaceWanRank -Iface $i -Vendor $Vendor) -eq $rank) { $raw += "$($i.attributes.'ip-address')" } }
    }
    foreach ($mm in [regex]::Matches("$($a.notes)", '\d{1,3}(?:\.\d{1,3}){3}(?::\d{1,5})?')) { $raw += $mm.Value }

    $out = @(); $seen = @{}
    foreach ($v in $raw) {
        $c = Get-WanCandidate $v
        if (-not $c -or $seen.ContainsKey($c.Ip)) { continue }
        if (-not $c.Port) { $c.Port = if ($Vendor -eq 'watchguard') { 8080 } else { 443 } }
        $seen[$c.Ip] = $true
        $out += $c
    }
    return $out
}

function New-WanIpRows {
    # The WAN IP detail row(s) for a firewall pill, from its already-matched IT Glue config: one clickable
    # https row per WAN IP, labelled 'WAN IP', 'WAN IP 2', ... (port shown unless 443). @() if none.
    param([pscustomobject]$Config, [ValidateSet('watchguard','sonicwall')][string]$Vendor)
    $ips = @(Resolve-FirewallWanIps -Config $Config -Vendor $Vendor)
    $rows = @()
    for ($k = 0; $k -lt $ips.Count; $k++) {
        $w = $ips[$k]
        $showPort = ($w.Port -and $w.Port -ne 443)
        $disp = if ($showPort) { "$($w.Ip):$($w.Port)" } else { $w.Ip }
        $url  = if ($showPort) { "https://$($w.Ip):$($w.Port)" } else { "https://$($w.Ip)" }
        $label = if ($k -eq 0) { 'WAN IP' } else { "WAN IP $($k + 1)" }
        $rows += New-DetailRow $label $disp -Link $url
    }
    return $rows
}

function Get-ItgPasswordIndex {
    # Org passwords that are EMBEDDED on a configuration, indexed by that config id (string -> @(passwords)).
    # One paged call, cached per run/org. Used to deep-link a firewall pill to its credentials.
    param([string]$OrgId)
    if ($script:PwIndex -and $script:PwIndexOrg -eq $OrgId) { return $script:PwIndex }
    $map = @{}; $page = 1
    do {
        try { $r = Get-ITGluePasswords -filter_organization_id $OrgId -page_size 1000 -page_number $page }
        catch { Write-Status "  IT Glue passwords query failed: $($_.Exception.Message)" Warning; break }
        foreach ($p in @($r.data)) {
            if ("$($p.attributes.'resource-type')" -eq 'configurations') {
                $cid = "$($p.attributes.'resource-id')"
                if ($cid) { if (-not $map.ContainsKey($cid)) { $map[$cid] = @() }; $map[$cid] += $p }
            }
        }
        $page++
    } while ($r.data -and $r.data.Count -eq 1000)
    $script:PwIndexOrg = $OrgId
    $script:PwIndex = $map
    return $map
}

function Get-FirewallCredentialUrl {
    # The 'Credentials' link for a firewall pill: the embedded password on its IT Glue config (preferring
    # the one whose username is 'admin' when several exist), else the IT Glue configuration page itself.
    param([pscustomobject]$Config, [hashtable]$PwIndex, [string]$LinkBase)
    if (-not $Config) { return '' }
    # Filter nulls: indexing a missing hashtable key yields $null, which @() would wrap into a 1-element
    # array and falsely trip the "has passwords" branch.
    $pws = @($PwIndex["$($Config.id)"] | Where-Object { $_ })
    if ($pws.Count -gt 0) {
        $admin = @($pws | Where-Object { "$($_.attributes.username)" -match '(?i)admin' }) | Select-Object -First 1
        $pick = if ($admin) { $admin } else { $pws[0] }
        return "$LinkBase/passwords/$($pick.id)"
    }
    return "$LinkBase/configurations/$($Config.id)"
}

function Get-UnifiSites {
    # Site Manager cloud API: one candidate per site across every host the API key can see.
    #   Id      = siteId          (stable cache key)
    #   Name    = meta.desc, or the host's reportedState.name when desc is blank/'Default'
    #             (this MSP runs multiple clients as separate sites on shared hosts; the client name
    #              lives in meta.desc, while meta.name is an internal slug like 'default'/'haeumwvi')
    #   Slug    = meta.name       (the site segment of the cloud URL)
    #   HostId  = hostId          (the host segment of the cloud URL)
    #   PathSeg = 'consoles' for UniFi-OS consoles, 'network-servers' for self-hosted Network servers.
    #             Driven by the host's 'type' field ('console' vs 'network-server') - the id FORMAT is
    #             not a reliable discriminator (both types appear as GUIDs and as 'MAC...:ts' strings).
    # Both collections page via nextToken.
    $hostInfo = @{}   # hostId -> @{ Name; PathSeg }
    try {
        $next = $null
        do {
            $uri = "$($script:UnifiBase)/v1/hosts?pageSize=200"
            if ($next) { $uri += "&nextToken=$([uri]::EscapeDataString($next))" }
            $r = Invoke-RestMethod -Uri $uri -Headers $script:UnifiHeaders -Method Get
            foreach ($h in $r.data) {
                $seg = if ("$($h.type)" -eq 'network-server') { 'network-servers' } else { 'consoles' }
                $hostInfo["$($h.id)"] = @{ Name = "$($h.reportedState.name)"; PathSeg = $seg }
            }
            $next = $r.nextToken
        } while ($next)
    } catch { Write-Status "  UniFi hosts query failed: $($_.Exception.Message)" Warning }

    $sites = @(); $next = $null
    do {
        $uri = "$($script:UnifiBase)/v1/sites?pageSize=200"
        if ($next) { $uri += "&nextToken=$([uri]::EscapeDataString($next))" }
        $r = Invoke-RestMethod -Uri $uri -Headers $script:UnifiHeaders -Method Get
        foreach ($s in $r.data) {
            $hi   = $hostInfo["$($s.hostId)"]
            $desc = "$($s.meta.desc)".Trim()
            $name = if ($desc -and $desc -ne 'Default') { $desc } elseif ($hi) { $hi.Name } else { '' }
            if (-not $name) { $name = if ($desc) { $desc } else { "$($s.meta.name)" } }
            $seg  = if ($hi -and $hi.PathSeg) { $hi.PathSeg } else { 'consoles' }
            $sites += [pscustomobject]@{
                Id = "$($s.siteId)"; Name = $name; Slug = "$($s.meta.name)"
                HostId = "$($s.hostId)"; PathSeg = $seg; Provider = 'UniFi'
            }
        }
        $next = $r.nextToken
    } while ($next)
    return $sites
}

function Resolve-UnifiSites {
    # Returns the UniFi site object(s) for this org (one pill each), or @() if confirmed "not in
    # UniFi". UniFi names are unreliable join keys (length-capped, '- <site>' suffixes, can't always
    # rename), so we confirm once via a fuzzy-ranked pick-list and cache the chosen siteIds under
    # $Entry.unifi; each run we rehydrate against the live site list so labels/links stay fresh.
    param([pscustomobject]$Entry, [string]$OrgName, [object[]]$Sites)

    if ($Entry._unmanaged -contains 'unifi') { return @() }
    $cached = $Entry.PSObject.Properties['unifi']
    if ($cached -and $cached.Value) {
        $ids = @(@($cached.Value) | ForEach-Object { "$($_.Id)" })
        $hit = @($Sites | Where-Object { $ids -contains "$($_.Id)" })
        if ($hit.Count -gt 0) { return $hit }
        return @($cached.Value)   # live list unavailable this run; trust the cache
    }
    if (-not $Sites -or $Sites.Count -eq 0) { return @() }

    # Fuzzy rank: strip a trailing ' - <suffix>', normalize, score by exact/prefix match + shared tokens.
    $target = ConvertTo-NormalizedName $OrgName
    $scored = @($Sites | ForEach-Object {
        $base = ($_.Name -replace '\s*-\s*[^-]+$', '')
        $n    = ConvertTo-NormalizedName $base
        $score = 0
        if ($n) {
            if ($n -eq $target) { $score += 1000 }
            elseif ($target -like "$n*" -or $n -like "$target*") { $score += 100 }
        }
        $score += @(($target -split ' ') | Where-Object { $_ -and (($n -split ' ') -contains $_) }).Count
        [pscustomobject]@{ Site = $_; Score = $score }
    } | Sort-Object -Property Score -Descending)

    $showAll = $false
    while ($true) {
        $list = if ($showAll) { @($scored) } else { @($scored | Select-Object -First 12) }
        Write-Host "`n  [UniFi] Select site(s) for '$OrgName' (e.g. 1,3 - multiple allowed):" -ForegroundColor Yellow
        for ($i = 0; $i -lt $list.Count; $i++) { Write-Host ("    [{0}] {1}" -f ($i + 1), $list[$i].Site.Name) }
        if (-not $showAll -and $scored.Count -gt $list.Count) { Write-Host "    [a] show ALL $($scored.Count) sites" }
        Write-Host "    [0] none / not in UniFi"
        $sel = Read-Host "  Choice"
        if ($sel -match '^[Aa]$') { $showAll = $true; continue }
        $tokens = @($sel -split '[,\s]+' | Where-Object { $_ -ne '' })
        if ($tokens.Count -eq 0) { continue }
        if ($tokens -contains '0') { Add-Unmanaged -Entry $Entry -VendorKey 'unifi'; return @() }
        $picked = @(); $bad = $false
        foreach ($t in $tokens) {
            $n = -1; [int]::TryParse($t, [ref]$n) | Out-Null
            if ($n -ge 1 -and $n -le $list.Count) { $picked += $list[$n - 1].Site } else { $bad = $true }
        }
        if ($bad -or $picked.Count -eq 0) { Write-Host "  Invalid selection." -ForegroundColor Red; continue }
        $picked = @($picked | Sort-Object -Property Id -Unique)
        $val = @($picked | ForEach-Object { [pscustomobject]@{ Id = "$($_.Id)"; Name = $_.Name; Slug = $_.Slug; HostId = $_.HostId; PathSeg = $_.PathSeg } })
        $Entry | Add-Member -NotePropertyName 'unifi' -NotePropertyValue $val -Force
        $script:ServiceMapDirty = $true
        return $picked
    }
}

function Get-DashNetwork {
    # 'Network' card: SonicWall + WatchGuard firewalls (live from the MySonicWall / WatchGuard Cloud
    # APIs, matched to this client by firewall friendlyName; one pill each with model + firmware +
    # license expiry, plus a WAN IP pulled from the matched IT Glue config) and UniFi (one pill per
    # matched site, linking to its unifi.ui.com console). -StatusId is the IT Glue 'Active' status id,
    # used to scope the config lookup for the WAN IP match.
    param([pscustomobject]$Entry, [string]$OrgId, [string]$OrgName, [string]$LinkBase, [string]$StatusId = '')
    $items = @()

    # SonicWall (MySonicWall): hardware fleet joined by friendlyName; per-firewall expiry via serviceInfo.
    if ($script:MswConnected) {
        $fleet = @()
        try { $fleet = Get-MswFirewalls } catch { Write-Status "  SonicWall fleet query failed: $($_.Exception.Message)" Warning }
        $matched = @(Resolve-FirewallSet -Entry $Entry -OrgName $OrgName -Devices $fleet -VendorKey 'sonicwallFw' -Label 'SonicWall')
        if ($matched.Count -gt 0) {
            $cfgIdx = Get-ItgConfigIndex -OrgId $OrgId -StatusId $StatusId
            $pwIdx  = Get-ItgPasswordIndex -OrgId $OrgId
            $pills = foreach ($f in ($matched | Sort-Object { "$($_.Name)" })) {
                # Pill label = the IT Glue config name (more descriptive), falling back to the MySonicWall
                # friendlyName when unmatched. Panel rows: Cloud (the MySonicWall console name), Model,
                # Firmware, Serial, Expiry, then WAN IP(s); plus a 'Credentials' link (to the firewall's
                # embedded IT Glue password, preferring the 'admin' one, else the config page).
                $cfg = Find-ItgConfigForFirewall -Index $cfgIdx -Serial $f.Id -Name $f.Name
                $fwLabel = if ($cfg -and $cfg.attributes.name) { "$($cfg.attributes.name)" } else { "$($f.Name)" }
                $detail = @()
                if ($cfg)        { $detail += New-DetailRow 'Cloud'    "$($f.Name)" }
                if ($f.Model)    { $detail += New-DetailRow 'Model'    $f.Model }
                if ($f.Firmware) { $detail += New-DetailRow 'Firmware' $f.Firmware }
                if ($f.Id)       { $detail += New-DetailRow 'Serial'   $f.Id }
                $ring = ''
                $exp = Get-MswFirewallExpiry -Serial $f.Id
                if ($exp -and $exp.Date) {
                    $soon = ($exp.Soon -or $exp.Days -le 30)
                    $detail += New-DetailRow 'Expires' "$($exp.Date) ($($exp.Days)d)" ($(if ($soon) { 'bad' } else { '' }))
                    if ($soon) { $ring = '#DC3545' }
                } elseif ($f.LicenseExpired) {
                    $detail += New-DetailRow 'License' 'expired' 'bad'; $ring = '#DC3545'
                }
                $detail += @(New-WanIpRows -Config $cfg -Vendor 'sonicwall')
                $credUrl = Get-FirewallCredentialUrl -Config $cfg -PwIndex $pwIdx -LinkBase $LinkBase
                New-CardItem -Label $fwLabel -Detail $detail -Ring $ring -Link $credUrl -LinkText 'Credentials'
            }
            $items += New-CardGroup -Label 'SonicWall' -Brand 'sonicwall.com' -Items $pills
        }
    }

    # WatchGuard Cloud: allocated Fireboxes joined by friendlyName; model + license expiry.
    if ($script:WgConnected) {
        $fleet = @()
        try { $fleet = Get-WgFireboxes } catch { Write-Status "  WatchGuard query failed: $($_.Exception.Message)" Warning }
        $matched = @(Resolve-FirewallSet -Entry $Entry -OrgName $OrgName -Devices $fleet -VendorKey 'watchguard' -Label 'WatchGuard')
        if ($matched.Count -gt 0) {
            $cfgIdx = Get-ItgConfigIndex -OrgId $OrgId -StatusId $StatusId
            $pwIdx  = Get-ItgPasswordIndex -OrgId $OrgId
            $pills = foreach ($f in ($matched | Sort-Object { "$($_.Name)" })) {
                # Pill label = the IT Glue config name (more descriptive), falling back to the WatchGuard
                # Cloud friendlyName when unmatched. Panel rows: Cloud (the WatchGuard Cloud name), Model,
                # Serial, MAC, Expiry, then WAN IP(s); plus a 'Credentials' link (to the firewall's embedded
                # IT Glue password, preferring the 'admin' one, else the config page). The WG API exposes no
                # firmware on this summary, so that row is omitted.
                $cfg = Find-ItgConfigForFirewall -Index $cfgIdx -Serial $f.Id -Name $f.Name
                $fwLabel = if ($cfg -and $cfg.attributes.name) { "$($cfg.attributes.name)" } else { "$($f.Name)" }
                $detail = @()
                if ($cfg)     { $detail += New-DetailRow 'Cloud'  "$($f.Name)" }
                if ($f.Model) { $detail += New-DetailRow 'Model'  $f.Model }
                if ($f.Id)    { $detail += New-DetailRow 'Serial' $f.Id }
                if ($f.Mac)   { $detail += New-DetailRow 'MAC'    $f.Mac }
                $ring = ''
                # These Fireboxes are rented, so they always read as near-expiry. Only flag (red ring) when
                # the license has ACTUALLY expired (negative days); show the real expiry date either way in
                # the panel, coloured red once expired.
                if ($f.ExpiryDate) {
                    $expired = ($f.Days -lt 0)
                    $detail += New-DetailRow 'Expires' "$($f.ExpiryDate) ($($f.Days)d)" ($(if ($expired) { 'bad' } else { '' }))
                    if ($expired) { $ring = '#DC3545' }
                }
                $detail += @(New-WanIpRows -Config $cfg -Vendor 'watchguard')
                $credUrl = Get-FirewallCredentialUrl -Config $cfg -PwIndex $pwIdx -LinkBase $LinkBase
                New-CardItem -Label $fwLabel -Detail $detail -Ring $ring -Link $credUrl -LinkText 'Credentials'
            }
            $items += New-CardGroup -Label 'WatchGuard' -Brand 'watchguard.com' -Items $pills
        }
    }

    if ($script:UnifiConnected) {
        $sites = @()
        try { $sites = Get-UnifiSites } catch { Write-Status "  UniFi query failed: $($_.Exception.Message)" Warning }
        $matched = @(Resolve-UnifiSites -Entry $Entry -OrgName $OrgName -Sites $sites)
        if ($matched.Count -gt 0) {
            $pills = foreach ($s in ($matched | Sort-Object { "$($_.Name)" })) {
                $slug = if ($s.Slug) { $s.Slug } else { 'default' }
                $seg  = if ($s.PathSeg) { $s.PathSeg } else { 'consoles' }
                $link = if ($s.HostId) { "https://unifi.ui.com/$seg/$($s.HostId)/network/$slug/dashboard" } else { 'https://unifi.ui.com' }
                New-CardItem -Label "$($s.Name)" -Link $link
            }
            $items += New-CardGroup -Label 'UniFi' -Brand 'ui.com' -Items $pills
        }
    }

    return New-Card -Title 'Network' -Items $items
}

# ==============================================================================
# Rendering
# ==============================================================================
$script:CardColor = '#051554'   # navy card background

function ConvertTo-CardGroupHtml {
    # A labelled sub-box: a brand-coloured strip across the top (brand name only) over a neutral body
    # of pills. Each inner pill carries the brand logo on its LEFT; brand colour lives only in the
    # strip (flooding every pill looked too loud for bright brands like WatchGuard red / Veeam green).
    # -TwoCol lays the inner pills out in two columns (used inside a wide/double-width card).
    param([pscustomobject]$Group, [switch]$TwoCol)
    $brand = Resolve-Brand $Group.Brand
    $brandColor = if ($brand -and $brand.Color) { $brand.Color } else { 'rgba(255,255,255,0.55)' }
    $strip = "<div style=`"background:$brandColor;padding:8px 12px;`">" +
             "<span style=`"font-size:11px;font-weight:700;letter-spacing:0.06em;text-transform:uppercase;color:#fff;`">$(Get-HtmlEncoded $Group.Label)</span></div>"
    # Inner pills inherit the group's brand so each shows the brand logo on its left; they stay neutral.
    $inner = (@($Group.Items) | ForEach-Object { if (-not $_.Brand) { $_.Brand = $Group.Brand }; ConvertTo-CardItemHtml $_ }) -join ''
    $body = if ($TwoCol) { 'display:grid;grid-template-columns:1fr 1fr;gap:8px;padding:10px;' } else { 'display:flex;flex-direction:column;gap:8px;padding:10px;' }
    return "<div style=`"background:rgba(255,255,255,0.06);border-radius:12px;overflow:hidden;`">$strip" +
           "<div style=`"$body`">$inner</div></div>"
}

function Format-DetailMembers {
    # Wrap a run of V-only member rows (already rendered HTML) in the redesign's indented, left-bordered
    # list. A long list (> 7 rows, e.g. many stale devices) becomes an internally scrolling box so the
    # pill panel stays compact; the scrollbar is styled via the .scrolllist rule in New-DashboardHtml.
    param([string[]]$Members)
    if (-not $Members -or $Members.Count -eq 0) { return '' }
    if ($Members.Count -gt 7) {
        return "<div class=`"scrolllist`" style=`"max-height:116px;overflow:auto;margin:2px 0 0 4px;padding-left:11px;padding-right:6px;border-left:2px solid rgba(139,150,176,0.30);`">$($Members -join '')</div>"
    }
    return "<div style=`"margin:2px 0 5px 4px;padding-left:11px;border-left:2px solid rgba(139,150,176,0.30);`">$($Members -join '')</div>"
}

function ConvertTo-CardItemHtml {
    param([pscustomobject]$Item, [switch]$TwoCol)
    if ($Item.Kind -eq 'group') { return ConvertTo-CardGroupHtml $Item -TwoCol:$TwoCol }
    $brand = Resolve-Brand $Item.Brand
    # Pills are always NEUTRAL now - brand identity is the logo (left) plus, for sub-boxes, the coloured
    # header strip. An explicit Bg still wins (kept for any deliberate override). Interactive pills
    # (link / expand) sit a touch brighter than static info pills so they read as actionable.
    $isInteractive = [bool]($Item.Link -or (@($Item.Detail).Count -gt 0))
    if     ($Item.Bg)       { $bg = $Item.Bg; $border = 'none' }
    elseif ($isInteractive) { $bg = 'rgba(255,255,255,0.10)'; $border = '1px solid rgba(255,255,255,0.18)' }
    else                    { $bg = 'rgba(255,255,255,0.06)'; $border = '1px solid rgba(255,255,255,0.12)' }

    # Outline ring (box-shadow: no layout size, sits flush to the pill). Explicit Ring colour wins;
    # otherwise Alert draws the red error ring. Width/colours match the redesign (2px, softer red).
    $ringColor = if ($Item.Ring) { $Item.Ring } elseif ($Item.Alert) { '#e0556a' } else { '' }
    $ring = if ($ringColor) { "box-shadow:0 0 0 2px $ringColor;" } else { '' }

    if ($Item.Kind -eq 'number') {
        $inner = "<span style=`"font-size:30px;font-weight:800;line-height:1;`">$(Get-HtmlEncoded $Item.Label)</span>" +
                 "<span style=`"font-size:13px;font-weight:600;opacity:0.85;margin-top:4px;`">$(Get-HtmlEncoded $Item.Sub)</span>"
        $style = "display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;" +
                 "padding:16px 14px;border-radius:10px;text-decoration:none;word-break:break-word;background:$bg;color:#fff;border:$border;$ring"
    }
    else {
        $icon = ''
        if ($brand -and $brand.Logo) {
            $icon = "<span style=`"display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;border-radius:7px;background:#fff;flex:none;`">" +
                    "<img src=`"$($brand.Logo)`" alt=`"`" width=`"18`" height=`"18`" style=`"display:block;width:18px;height:18px;object-fit:contain;`"></span>"
        }
        $weight = '600'
        # Icon pinned left; label takes the remaining width and centres within it (flex:1). With no icon
        # the label span fills the whole pill, so iconless pills stay fully centred. A Sub value adds a
        # small second line under the label (multi-line pill, e.g. a device's last-backup/size stats).
        if ($Item.Sub) {
            $labelHtml = "<span style=`"flex:1;min-width:0;display:flex;flex-direction:column;align-items:center;gap:2px;`">" +
                         "<span>$(Get-HtmlEncoded $Item.Label)</span>" +
                         "<span style=`"font-size:11px;font-weight:500;color:#8b96b0;`">$(Get-HtmlEncoded $Item.Sub)</span></span>"
        } else {
            $labelHtml = "<span style=`"flex:1;min-width:0;text-align:center;`">$(Get-HtmlEncoded $Item.Label)</span>"
        }
        $inner  = "$icon$labelHtml"
        $style  = "display:flex;align-items:center;justify-content:flex-start;gap:10px;text-align:center;" +
                  "padding:12px 14px;border-radius:10px;font-size:14px;font-weight:$weight;line-height:1.3;" +
                  "text-decoration:none;word-break:break-word;background:$bg;color:#fff;border:$border;$ring"

        # Expandable pill: render as a native <details>/<summary> disclosure (no JS, no required CSS - the
        # toggle is built into the browser and the panel is hidden natively when closed). The default
        # disclosure triangle is removed inline (summary is display:flex + list-style:none); our own caret
        # sits on the right and flips on open via the one cosmetic rule in New-DashboardHtml. Everything is
        # inline so it survives publishing to IT Glue; if that rule (or <details> itself) were ever stripped,
        # the worst case is the caret not flipping / the panel rendering always-open - never broken.
        if (@($Item.Detail).Count -gt 0) {
            $caret = "<span class=`"xcaret`" style=`"flex:none;font-size:13px;line-height:1;opacity:0.55;color:#8b96b0;width:16px;text-align:center;transition:transform .15s;`">&#9662;</span>"
            $out = New-Object System.Collections.Generic.List[string]
            if ($Item.DetailHead) {
                $out.Add("<div style=`"font-size:12.5px;font-weight:600;color:#8b96b0;margin:0 0 5px;`">$(Get-HtmlEncoded $Item.DetailHead)</div>")
            }
            # V-only entries buffer into an indented, left-bordered member list (scrollable when long) - the
            # redesign's treatment for name lists (stale devices, DCs, ...). K/V entries render as stat rows
            # with a muted label and a bright (or cyan-linked) value; a K-only entry is a sub-caption.
            $members = New-Object System.Collections.Generic.List[string]
            foreach ($r in @($Item.Detail)) {
                $isMember = ($r.V -and -not $r.K)
                if (-not $isMember -and $members.Count -gt 0) { $out.Add((Format-DetailMembers $members)); $members.Clear() }
                if ($r.K -and -not $r.V) {
                    $out.Add("<div style=`"font-size:12.5px;font-weight:600;color:#8b96b0;padding:3px 0;`">$(Get-HtmlEncoded $r.K)</div>")
                } elseif ($r.K) {
                    # State colours the value: 'bad' red, 'ok' green (redesign palette). A LINKED value is a
                    # cyan underlined link that navigates without toggling the disclosure (it's in the panel).
                    $vc = switch ($r.State) { 'bad' { 'color:#e0556a;' } 'ok' { 'color:#3fbf73;' } default { '' } }
                    $valHtml = if ($r.Link) {
                        "<a href=`"$($r.Link)`" target=`"_blank`" rel=`"noopener`" style=`"flex:none;color:#34c5f0;font-weight:600;text-decoration:underline;text-underline-offset:2px;text-align:right;$vc`">$(Get-HtmlEncoded $r.V)</a>"
                    } else {
                        "<span style=`"flex:none;color:#eef1f8;font-weight:600;$vc`">$(Get-HtmlEncoded $r.V)</span>"
                    }
                    $out.Add("<div style=`"display:flex;justify-content:space-between;gap:12px;font-size:12.5px;padding:3px 0;`">" +
                             "<span style=`"flex:1;min-width:0;color:#8b96b0;line-height:1.3;`">$(Get-HtmlEncoded $r.K)</span>$valHtml</div>")
                } else {
                    # V-only entry. With a -Link it's a cyan underlined member link (e.g. a stale device or
                    # domain controller -> its Automate page); otherwise a plain bright line.
                    if ($r.Link) {
                        $members.Add("<a href=`"$($r.Link)`" target=`"_blank`" rel=`"noopener`" style=`"display:block;font-size:12.5px;font-weight:600;color:#34c5f0;text-decoration:underline;text-underline-offset:2px;padding:1px 0;word-break:break-word;`">$(Get-HtmlEncoded $r.V)</a>")
                    } else {
                        $members.Add("<div style=`"font-size:12.5px;color:#eef1f8;padding:1px 0;word-break:break-word;`">$(Get-HtmlEncoded $r.V)</div>")
                    }
                }
            }
            if ($members.Count -gt 0) { $out.Add((Format-DetailMembers $members)); $members.Clear() }
            $rows = ($out -join '')
            if ($Item.Link) {
                $lt = if ($Item.LinkText) { $Item.LinkText } else { 'Open' }
                $rows += "<a href=`"$($Item.Link)`" target=`"_blank`" rel=`"noopener`" style=`"display:inline-flex;align-items:center;gap:5px;margin-top:11px;font-size:12px;font-weight:700;color:#34c5f0;text-decoration:underline;text-underline-offset:2px;`">$(Get-HtmlEncoded $lt) &#8599;</a>"
            }
            # color:#fff is set explicitly: the card container has no text colour, so panel rows would
            # otherwise inherit the browser default (black) and be illegible on the dark panel.
            $panel = "<div style=`"margin-top:8px;background:rgba(0,0,0,0.28);border:1px solid rgba(255,255,255,0.10);border-radius:10px;padding:11px 13px;color:#fff;`">$rows</div>"
            return "<details style=`"width:100%;`"><summary class=`"pill`" style=`"$style cursor:pointer;list-style:none;`">$icon$labelHtml$caret</summary>$panel</details>"
        }
    }

    if ($Item.Link) {
        # Non-expandable link pill: trailing cyan up-right arrow (redesign) + hover class. Number pills,
        # which have a centred column layout, get no arrow.
        $arrow = if ($Item.Kind -ne 'number') { "<span style=`"flex:none;color:#34c5f0;font-size:13px;`">&#8599;</span>" } else { '' }
        return "<a class=`"pill`" href=`"$($Item.Link)`" target=`"_blank`" rel=`"noopener`" style=`"$style`">$inner$arrow</a>"
    }
    return "<div style=`"$style`">$inner</div>"
}

# Height score above which a card spans two grid columns and lays its bands out two-up (so a tall,
# band-heavy card doesn't become very tall & narrow). Score ~ rows of content: each band = 1 (header)
# + its pills; each plain pill = 1.
$script:TallCardThreshold = 6

function Get-CardHeightScore {
    # Approximate content height in "rows": each band adds a header row plus its pills; plain/number
    # pills add one. Computed on the already-banded item list.
    param([object[]]$Items)
    $n = 0
    foreach ($it in @($Items)) {
        if ($it.Kind -eq 'group') { $n += 1 + @($it.Items).Count } else { $n += 1 }
    }
    return $n
}

function Group-CardItems {
    # Wrap runs of ADJACENT same-brand top-level pills into vendor bands (sub-boxes) so every card gets
    # brand header strips. Existing 'group' items and 'number' pills pass through unchanged; pills whose
    # brand has no display name (or none) render bare. Order is preserved - only adjacent pills merge.
    param([object[]]$Items)
    $out = New-Object System.Collections.Generic.List[object]
    $run = New-Object System.Collections.Generic.List[object]
    $runBrand = $null
    foreach ($it in @($Items)) {
        $bandable = ($it.Kind -ne 'group') -and ($it.Kind -ne 'number') -and $it.Brand -and $script:BrandName.ContainsKey("$($it.Brand)")
        if ($bandable -and ($run.Count -eq 0 -or "$($it.Brand)" -eq "$runBrand")) {
            $runBrand = "$($it.Brand)"; $run.Add($it); continue
        }
        if ($run.Count -gt 0) {
            $out.Add((New-CardGroup -Label $script:BrandName["$runBrand"] -Brand $runBrand -Items @($run.ToArray())))
            $run = New-Object System.Collections.Generic.List[object]; $runBrand = $null
        }
        if ($bandable) { $runBrand = "$($it.Brand)"; $run.Add($it) } else { $out.Add($it) }
    }
    if ($run.Count -gt 0) {
        $out.Add((New-CardGroup -Label $script:BrandName["$runBrand"] -Brand $runBrand -Items @($run.ToArray())))
    }
    return $out.ToArray()
}

function ConvertTo-CardHtml {
    param([pscustomobject]$Card)
    $items = @($Card.Items)
    if ($items.Count -eq 0) { $items = @(New-CardItem -Label 'None') }   # empty category -> muted 'None'
    # Auto-band the card's pills by vendor (unless opted out, e.g. Domains).
    if (-not $Card.NoBand) { $items = @(Group-CardItems -Items $items) }
    $h3 = "<h3 style=`"margin:0;color:#fff;font-size:13px;font-weight:700;letter-spacing:0.08em;text-transform:uppercase;text-align:center;opacity:0.82;`">$(Get-HtmlEncoded $Card.Title)</h3>"

    # Cards size to their content (no fixed square - band-heavy cards would otherwise overflow it).
    # A tall, band-heavy card spans TWO grid columns and lays its bands out two-up, roughly halving the
    # height so it doesn't become tall & narrow. A modest min-height keeps light cards from looking thin.
    $base = "background:$($script:CardColor);border-radius:16px;padding:22px 20px;display:flex;flex-direction:column;gap:14px;min-height:200px;box-shadow:0 2px 6px rgba(5,21,84,0.12);"
    if ((Get-CardHeightScore -Items $items) -gt $script:TallCardThreshold) {
        # Two-up via CSS multi-column: bands pack by their natural height (no stretching, so a single-pill
        # band like Duo has no dead space); break-inside keeps each band intact, margin gives the gap.
        $body = (@($items) | ForEach-Object { "<div style=`"break-inside:avoid;margin-bottom:10px;`">$(ConvertTo-CardItemHtml $_)</div>" }) -join ''
        $inner = "<div style=`"flex:1;column-count:2;column-gap:10px;`">$body</div>"
        return "<div style=`"grid-column:span 2;$base`">$h3$inner</div>"
    }

    $body = ($items | ForEach-Object { ConvertTo-CardItemHtml $_ }) -join ''
    $inner = "<div style=`"flex:1;display:flex;flex-direction:column;justify-content:center;gap:10px;`">$body</div>"
    return "<div style=`"$base`">$h3$inner</div>"
}

function New-DashboardHtml {
    # Responsive card grid - this is the exact markup a future -Publish would store in IT Glue
    # (all card styling is inline, which IT Glue preserves). The single <style> rule below is the one
    # exception: it only flips the expand caret on open (purely cosmetic). The expandable pills themselves
    # are native <details>/<summary> and work with no CSS at all, so if IT Glue strips this rule the only
    # effect is the caret not rotating - the pills still open/close.
    param([pscustomobject[]]$Cards)
    $cards = ($Cards | ForEach-Object { ConvertTo-CardHtml -Card $_ }) -join ''
    # The one <style> block: all cosmetic, degrades gracefully if IT Glue strips it (caret won't flip,
    # pills won't brighten on hover, the scroll list keeps the OS scrollbar) - pills still open/close.
    $caretCss = "<style>" +
        "details[open] > summary .xcaret{transform:rotate(180deg);}" +
        ".pill{transition:background .15s;}" +
        ".pill:hover{background:rgba(255,255,255,0.16)!important;}" +
        ".scrolllist{scrollbar-width:thin;scrollbar-color:rgba(120,180,230,0.45) transparent;}" +
        ".scrolllist::-webkit-scrollbar{width:6px;}" +
        ".scrolllist::-webkit-scrollbar-thumb{background:rgba(120,180,230,0.45);border-radius:3px;}" +
        ".scrolllist::-webkit-scrollbar-thumb:hover{background:rgba(120,180,230,0.7);}" +
        ".scrolllist::-webkit-scrollbar-track{background:transparent;}" +
        "</style>"
    return "$caretCss<div style=`"display:grid;grid-template-columns:repeat(auto-fit, minmax(260px, 1fr));gap:18px;align-items:start;`">$cards</div>"
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
<div style="max-width: 1240px; margin: 0 auto; padding: 36px 28px 56px;">
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

# Cloudflare - the zone list (used to colour domains in our self-service portal) loads LAZILY on the
# first Domains card via Get-CloudflarePortalSet, so a single-org run with no domains never makes the
# whole-account call. Here we only stash the token/flag; $null portal = "not yet loaded".
$script:CloudflarePortal = $null
$script:CloudflareConfigured = [bool]$creds.Cloudflare.Configured
$script:CloudflareToken = if ($creds.Cloudflare.Configured) { $creds.Cloudflare.ApiToken } else { '' }

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

# UniFi Site Manager (cloud) - one API key spans every client console/site under the account. Used by
# the Network card to list a client's UniFi sites and deep-link to their unifi.ui.com cloud console.
$script:UnifiConnected = $false
if ($creds.UniFi.Configured) {
    try {
        $script:UnifiBase    = 'https://api.ui.com'
        $script:UnifiHeaders = @{ 'X-API-KEY' = $creds.UniFi.ApiKey; 'Accept' = 'application/json' }
        $null = Invoke-RestMethod -Uri "$script:UnifiBase/v1/hosts?pageSize=1" -Headers $script:UnifiHeaders -Method Get
        $script:UnifiConnected = $true
    } catch { Write-Status "UniFi connect failed: $($_.Exception.Message)" Warning }
}

# SonicWall (MySonicWall) - one X-api-key spans every client tenant. Drives the Capture Client (Devices)
# and CAS (Email) counts and the SonicWall firewall pills (Network). Tenant list + firewall fleet are
# fetched lazily and cached for the run (Get-MswTenants / Get-MswFirewalls).
$script:MswConnected = $false
$script:MswTenants = $null; $script:MswFleet = $null; $script:MswUser = ''
if ($creds.MySonicWall.Configured) {
    try {
        $script:MswBase = 'https://api.mysonicwall.com'
        $script:MswHeaders = @{ 'X-api-key' = $creds.MySonicWall.ApiKey; 'Accept' = 'application/json' }
        $script:MswFirewallGroupId = "$($creds.MySonicWall.FirewallGroupId)"
        $null = Get-MswTenants   # validates the key and warms the tenant cache / userName / firewall group
        $script:MswConnected = $true
        Write-Status "MySonicWall: $($script:MswTenants.Count) tenants (firewall group $script:MswFirewallGroupId)." Detail
    } catch { Write-Status "MySonicWall connect failed: $($_.Exception.Message)" Warning }
}

# WatchGuard Cloud - OAuth token (cached for the run) + API-Key header. Drives the WatchGuard firewall
# pills on the Network card. Fireboxes are fetched lazily and cached (Get-WgFireboxes).
$script:WgConnected = $false; $script:WgFleet = $null
if ($creds.WatchGuard.Configured) {
    try {
        $script:WgApiBase   = 'https://api.usa.cloud.watchguard.com/rest'
        $script:WgAccountId = $creds.WatchGuard.AccountId
        $basic = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($creds.WatchGuard.AccessId):$($creds.WatchGuard.Password)"))
        $wgTok = (Invoke-RestMethod -Uri 'https://api.usa.cloud.watchguard.com/oauth/token' -Method Post `
            -Headers @{ Authorization = "Basic $basic"; Accept = 'application/json' } `
            -ContentType 'application/x-www-form-urlencoded' -Body 'grant_type=client_credentials&scope=api-access').access_token
        if (-not $wgTok) { throw "No access_token returned." }
        $script:WgHeaders = @{ Authorization = "Bearer $wgTok"; 'WatchGuard-API-Key' = $creds.WatchGuard.ApiKey; Accept = 'application/json' }
        $script:WgConnected = $true
    } catch { Write-Status "WatchGuard connect failed: $($_.Exception.Message)" Warning }
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

# Canonical match name: ConnectWise Manage is the source of truth (IT Glue syncs from it), so every
# other vendor is matched against the CWM company name. The CWM company is resolved from the IT Glue
# name (the one bootstrap match); if CWM isn't connected or the org isn't in CWM, fall back to the
# IT Glue name. $orgName stays the IT Glue name for display (header/title/filename).
$cwmCompany = if ($script:CwmConnected) { Resolve-CwmCompany -Entry $entry -OrgName $orgName } else { $null }
$matchName  = if ($cwmCompany -and $cwmCompany.Name) { $cwmCompany.Name } else { $orgName }
if ($matchName -ne $orgName) { Write-Status "Matching vendors against ConnectWise name: '$matchName'" Detail }

# Resolve the CIPP/M365 tenant once (may prompt) and share it between the Identity and M365 tiles.
$cippTenant = if ($script:CippConnected) { Resolve-CippTenant -Entry $entry -OrgId $orgId -OrgName $matchName } else { $null }
# Resolve the MySonicWall tenant once (may prompt) and share it between the Devices (Capture Client) and
# Email (CAS) tiles. Firewalls are matched separately by friendlyName inside Get-DashNetwork.
$mswTenant  = if ($script:MswConnected) { Resolve-MswTenant -Entry $entry -OrgName $matchName } else { $null }

$usersCard   = Get-DashUsers   -Entry $entry -OrgId $orgId -OrgName $matchName -LinkBase $linkBase -Creds $creds
$tileDevices = Get-DashDevices -Entry $entry -OrgId $orgId -OrgName $matchName -LinkBase $linkBase -WsTypeId $wsTypeId -SvTypeId $svTypeId -StatusId $activeStatusId -CippTenant $cippTenant
$tileNetwork = Get-DashNetwork -Entry $entry -OrgId $orgId -OrgName $matchName -LinkBase $linkBase -StatusId $activeStatusId
# Build Identity before Endpoint Security: Identity resolves+caches the per-client Duo match that the
# Endpoint Security card's "Duo Logon" pill gates on. Display order is fixed by $cards below, so this
# has no visual effect.
# The Users tile (ConnectWise + IT Glue people) is merged INTO the Identity card: ConnectWise pills
# first (PeopleFirst agreement, then contacts), then IT Glue, then the identity pills (Duo, Microsoft
# 365, Domain Controllers). Auto-banding turns each vendor run into its own labelled band.
$identityCard = Get-DashIdentity -Entry $entry -OrgId $orgId -OrgName $matchName -Creds $creds -CippTenant $cippTenant
$tileIdentity = New-Card -Title 'Identity' -Items (@($usersCard.Items) + @($identityCard.Items))
$tileAv      = Get-DashAntivirus -Entry $entry -OrgName $matchName -Creds $creds -MswTenant $mswTenant
# Merge Endpoint Security INTO the Devices card: Devices pills first (Automate + M365/Intune bands),
# then the Endpoint Security pills (AV vendors + Duo Logon). Auto-banding keeps each vendor run labelled.
$tileDevices = New-Card -Title 'Devices' -Items (@($tileDevices.Items) + @($tileAv.Items))
$tileBackup = Get-DashBackup    -Entry $entry -OrgName $matchName -Creds $creds
$tileM365   = Get-DashM365      -Entry $entry -OrgId $orgId -LinkBase $linkBase -CippTenant $cippTenant -MswTenant $mswTenant

# Persist any new resolutions
if ($script:ServiceMapDirty) {
    $entry.resolvedOn = (Get-Date).ToString('yyyy-MM-dd')
    Save-ServiceMap -Map $script:ServiceMap
    Write-Status "Service-map cache updated: $ServiceMapPath" Detail
}
# Persist any newly-resolved Brandfetch lookups so they're never fetched again.
if ($script:BrandCacheDirty) {
    $script:BrandCache | ConvertTo-Json -Depth 4 | Set-Content -Path $BrandCachePath -Encoding UTF8
    Write-Status "Brand cache updated: $BrandCachePath" Detail
}

# -- Render --
# Fixed display order: row 1 = Identity (now incl. Users), Email (now incl. verified M365 domains);
# row 2 = Devices (now incl. Endpoint Security), Backup, Network.
$cards = @($tileIdentity, $tileM365, $tileDevices, $tileBackup, $tileNetwork)
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
