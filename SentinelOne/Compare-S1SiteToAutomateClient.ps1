<#
.SYNOPSIS
    READ-ONLY reconciliation of ConnectWise Automate clients against SentinelOne sites.

.DESCRIPTION
    Reconciles the set of companies that exist in BOTH ConnectWise Automate and
    SentinelOne, and verifies that each one resolves to exactly one S1 site by
    normalized name. This is a pre-flight validation for a downstream agent migration.

    Scope: the intersection, by design. Most Automate clients have no S1 subtenant and
    most are never going to -- those are reported as a count and carried in the CSV,
    but they are NOT failures and do NOT affect the exit code. The migration only ever
    runs against companies present on both sides.

    The one risk that scoping creates is a company that SHOULD have matched but did not,
    because the two names drifted apart. That case is caught separately: any unmatched
    client whose closest S1 site clears -RenameHintSimilarity is surfaced as a
    "possible rename" warning, so a real client is never silently dropped from scope.

    Both sides are pulled once, in full, and matched in memory. Automate is reached
    through the AutomateAPI PowerShell module; SentinelOne through its v2.1 REST API.

    Outputs:
      1. Console summary.
      2. Console table of every AMBIGUOUS row (the operator's fix list), plus any
         possible-rename warnings.
      3. Timestamped CSV of the complete result set, all statuses.
      4. Timestamped CSV of orphan S1 sites (sites matching no Automate client).
      5. Timestamped install site map CSV -- the deliverable. One row per matched
         company, reduced to a single variable the migration script keys off:
         AutomateClientId -> S1SiteId. Unmatched and ambiguous companies are excluded
         by construction, so the migration cannot install into a guessed site.

    Exit codes:
      0 = every company present on both sides maps 1:1 (safe to run the migration)
      1 = at least one AMBIGUOUS company; the mapping is not deterministic
      2 = fatal error (auth, network, empty result, missing prerequisite)

    Unmatched clients do not set a non-zero exit code. Possible-rename hints do not
    either -- they are advisory, because a false hint must not block the pipeline.

.NOTES
    ############################################################################
    #  THIS SCRIPT IS STRICTLY READ-ONLY AND IDEMPOTENT.                       #
    #                                                                          #
    #  It performs GET operations only. It does NOT install, uninstall, or     #
    #  reboot anything; it does NOT write to SentinelOne (no externalId        #
    #  stamping); it does NOT modify Automate records.                         #
    #                                                                          #
    #  Do NOT bolt install/uninstall or any write logic onto this script.      #
    #  That is a separate, later phase and belongs in a separately named       #
    #  script. Callers rely on being able to run this repeatedly and safely.   #
    ############################################################################

    Prerequisites:
      - AutomateAPI module (Install-Module AutomateAPI)
      - A registered Automate integrator ClientId GUID (https://developer.connectwise.com/ClientID).
        Mandatory on Automate 20 Patch 11 and above; auth fails without it regardless
        of credential validity.
      - A SentinelOne API token from an ACCOUNT-scoped service user
        (Settings > Users > Service Users). A site-scoped token sees only its own site
        and every match will fail. The script warns if the returned site count looks
        implausibly low.

    Run under Windows PowerShell 5.1 (powershell.exe), not PowerShell 7.

    On first run the script prompts for connection details and stores them in
    Config\S1-AutomateReconcile-Credentials.xml (DPAPI-encrypted, user/machine
    specific). It will offer to seed those values from other credential stores in
    this repo; it reads them but never writes to them.

.PARAMETER AutomateServer
    Automate server FQDN, e.g. automate.contoso.com. Scheme is stripped if supplied.

.PARAMETER AutomateClientId
    Automate integrator ClientId GUID. Mandatory for the API on 20 Patch 11+.

.PARAMETER AutomateCredential
    Automate API credential. Prompted for if not supplied and not in the store.

.PARAMETER AutomateTwoFactorToken
    Current 2FA passcode, if the Automate account enforces two-factor.

.PARAMETER S1BaseUri
    SentinelOne console base URI. Defaults to the destination EDR console.

.PARAMETER S1ApiToken
    SentinelOne API token as a SecureString.

.PARAMETER LegalSuffix
    Standalone tokens stripped during normalization. Word-boundary matched, so 'co'
    does not turn 'Coastal' into 'astal'. Extend for tenant-specific conventions.

.PARAMETER StripLeadingArticle
    Leading articles removed during normalization, making a leading 'The' order
    independent. Default: 'the'.

.PARAMETER CollapseInitials
    Merge runs of two or more adjacent single-letter tokens, so 'P.C.' -> 'pc' and
    'R. W. Gordon' -> 'rw gordon'. Applied identically to both sides. Default: $true.
    Pass -CollapseInitials:$false to disable.

.PARAMETER ReportUnmatchedClients
    Also print the full list of Automate clients that have no S1 site. Off by default:
    that list is large, expected, and not actionable -- most Automate clients simply
    have no S1 subtenant. They are always present in the full CSV regardless.

.PARAMETER MinimumSimilarity
    Similarity floor (0..1) for the nearest-name suggestion columns in the CSVs.
    Suggestions below this are left blank rather than shown as noise. Default 0.6.

.PARAMETER RenameHintSimilarity
    Similarity floor (0..1) above which an unmatched client is flagged on the console as
    a possible rename rather than a company that is legitimately absent from S1.
    Deliberately stricter than -MinimumSimilarity. Default 0.8.

.PARAMETER ShowMatched
    Also print the OK rows. Default is summary plus exceptions only.

.PARAMETER IncludeRegistrationToken
    Add each site's SentinelOne registrationToken to the install site map, so the
    migration phase has everything it needs to enroll into the right site. Off by
    default: the token is a per-site enrollment secret, and including it turns a
    diagnostic report into a credential-bearing file. Nothing is written to
    SentinelOne either way -- the token is read from the existing site record.

.PARAMETER OutputPath
    Directory for the CSV reports. Defaults to Reports\ under the script directory.

.EXAMPLE
    .\Compare-S1SiteToAutomateClient.ps1

    Uses the stored configuration. Reconciles the companies present on both sides and
    writes the install site map. Exits 0 when that mapping is deterministic.

.EXAMPLE
    .\Compare-S1SiteToAutomateClient.ps1 -ReportUnmatchedClients -ShowMatched

    Adds the full unmatched-client list and the clean rows to the console output.

.EXAMPLE
    .\Compare-S1SiteToAutomateClient.ps1 -LegalSuffix inc,llc,ltd,corp,co,pc,pllc,group

    Extends the legal-suffix strip list for this run.
#>

[CmdletBinding()]
param(
    [string]$AutomateServer,

    [string]$AutomateClientId,

    [System.Management.Automation.PSCredential]$AutomateCredential,

    [string]$AutomateTwoFactorToken,

    [string]$S1BaseUri,

    [System.Security.SecureString]$S1ApiToken,

    [string[]]$LegalSuffix = @('inc', 'llc', 'ltd', 'corp', 'co', 'pc', 'pllc', 'plc', 'llp', 'lp'),

    [string[]]$StripLeadingArticle = @('the'),

    [bool]$CollapseInitials = $true,

    [switch]$ReportUnmatchedClients,

    [ValidateRange(0.0, 1.0)]
    [double]$MinimumSimilarity = 0.6,

    [ValidateRange(0.0, 1.0)]
    [double]$RenameHintSimilarity = 0.8,

    [switch]$ShowMatched,

    [switch]$IncludeRegistrationToken,

    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

# Windows PowerShell 5.1 defaults to SSL3/TLS1.0 for some calls; both endpoints require TLS 1.2.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$script:ExitCode = 0

#region Helpers

function Write-Step {
    param([string]$Message, [string]$Color = 'Cyan')
    Write-Host $Message -ForegroundColor $Color
}

function Get-NormalizedName {
    <#
        Normalization spec, applied identically to both sides:
          - lowercase
          - punctuation to spaces (word chars and spaces survive)
          - optionally merge runs of adjacent single-letter tokens (P.C. -> pc)
          - drop standalone legal-suffix tokens (word-boundary matched)
          - drop a leading article
          - collapse whitespace, trim

        Deliberately conservative. Over-aggressive stripping collapses genuinely
        different companies onto the same string, which shows up as a false AMBIGUOUS.
    #>
    param(
        [string]$Name,
        [string[]]$Suffix,
        [string[]]$Article,
        [bool]$MergeInitials
    )

    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }

    $text = $Name.ToLowerInvariant()

    # Keep word characters and whitespace; everything else becomes a separator.
    $text = $text -replace '[^\w\s]', ' '
    $text = ($text -replace '\s+', ' ').Trim()

    if ($MergeInitials) {
        # 'p c' -> 'pc', 'r w gordon' -> 'rw gordon'. Only runs of 2+ single letters
        # are merged; a lone initial ('charles e walton') is left alone.
        $tokens = @($text -split ' ' | Where-Object { $_ })
        $merged = New-Object System.Collections.Generic.List[string]
        $run = New-Object System.Collections.Generic.List[string]

        $flushRun = {
            if ($run.Count -ge 2) { $merged.Add(($run -join '')) }
            elseif ($run.Count -eq 1) { $merged.Add($run[0]) }
            $run.Clear()
        }

        foreach ($token in $tokens) {
            if ($token.Length -eq 1 -and $token -match '^[a-z]$') {
                $run.Add($token)
            }
            else {
                & $flushRun
                $merged.Add($token)
            }
        }
        & $flushRun

        $text = $merged -join ' '
    }

    if ($Suffix -and $Suffix.Count -gt 0) {
        $pattern = ($Suffix | Where-Object { $_ } | ForEach-Object { [regex]::Escape($_.ToLowerInvariant()) }) -join '|'
        if ($pattern) { $text = $text -replace "\b(?:$pattern)\b", ' ' }
    }

    $text = ($text -replace '\s+', ' ').Trim()

    if ($Article -and $Article.Count -gt 0) {
        $pattern = ($Article | Where-Object { $_ } | ForEach-Object { [regex]::Escape($_.ToLowerInvariant()) }) -join '|'
        if ($pattern) { $text = $text -replace "^(?:$pattern)\b", ' ' }
    }

    return ($text -replace '\s+', ' ').Trim()
}

function Get-LevenshteinDistance {
    # Two-row implementation; the full matrix form is both slower and awkward to index
    # in PowerShell. Used only for NO_MATCH hinting, never for matching itself.
    param([string]$Source, [string]$Target)

    $n = $Source.Length
    $m = $Target.Length
    if ($n -eq 0) { return $m }
    if ($m -eq 0) { return $n }

    $previous = New-Object 'int[]' ($m + 1)
    $current = New-Object 'int[]' ($m + 1)
    for ($j = 0; $j -le $m; $j++) { $previous[$j] = $j }

    for ($i = 1; $i -le $n; $i++) {
        $current[0] = $i
        $sourceChar = $Source[$i - 1]
        for ($j = 1; $j -le $m; $j++) {
            $cost = if ($sourceChar -eq $Target[$j - 1]) { 0 } else { 1 }
            $deletion = $previous[$j] + 1
            $insertion = $current[$j - 1] + 1
            $substitution = $previous[$j - 1] + $cost
            $current[$j] = [Math]::Min([Math]::Min($deletion, $insertion), $substitution)
        }
        $swap = $previous; $previous = $current; $current = $swap
    }

    return $previous[$m]
}

function Get-ClosestName {
    <#
        Returns the closest candidate to $Name with a similarity ratio, or $null if
        nothing clears $Floor. Candidates are pre-filtered on length delta, which is a
        lower bound on edit distance and cheaply skips most of the comparison set.
    #>
    param(
        [string]$Name,
        [string[]]$Candidates,
        [double]$Floor
    )

    if ([string]::IsNullOrWhiteSpace($Name) -or -not $Candidates) { return $null }

    $best = $null
    $bestDistance = [int]::MaxValue

    foreach ($candidate in $Candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        if ([Math]::Abs($Name.Length - $candidate.Length) -ge $bestDistance) { continue }

        $distance = Get-LevenshteinDistance -Source $Name -Target $candidate
        if ($distance -lt $bestDistance) {
            $bestDistance = $distance
            $best = $candidate
            if ($distance -eq 0) { break }
        }
    }

    if ($null -eq $best) { return $null }

    $longest = [Math]::Max($Name.Length, $best.Length)
    $similarity = if ($longest -eq 0) { 0 } else { 1 - ($bestDistance / $longest) }
    if ($similarity -lt $Floor) { return $null }

    return [PSCustomObject]@{
        Name       = $best
        Distance   = $bestDistance
        Similarity = [Math]::Round($similarity, 3)
    }
}

#endregion

#region Configuration

$ConfigPath = Join-Path $PSScriptRoot 'Config\S1-AutomateReconcile-Credentials.xml'

function Import-StoredConfig {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    try { return Import-Clixml -Path $Path }
    catch {
        Write-Warning "Could not read $Path ($($_.Exception.Message)). Re-prompting."
        return $null
    }
}

function Get-SeedConfig {
    <#
        Other scripts in this repo keep their own credential stores. Reading one to
        pre-fill this script's first run saves retyping a 36-character GUID; this
        script owns its own file and never writes back to theirs.
    #>
    param([string]$S1Host)

    $seedPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'IT Glue\Dashboard\Config\dashboard-credentials.xml'
    if (-not (Test-Path $seedPath)) { return $null }

    try { $seed = Import-Clixml -Path $seedPath } catch { return $null }

    $result = @{ Source = $seedPath; Automate = $null; S1 = $null }

    if ($seed.Automate -and $seed.Automate.Url) {
        $result.Automate = @{
            Server   = ($seed.Automate.Url -replace '^https?://', '').TrimEnd('/')
            Username = $seed.Automate.Username
            Password = $seed.Automate.Password
            ClientId = $seed.Automate.ClientId
        }
    }

    if ($seed.SentinelOne -and $seed.SentinelOne.Instances) {
        $instance = $seed.SentinelOne.Instances | Where-Object {
            $_.URL -and ([Uri]$_.URL).Host -eq $S1Host
        } | Select-Object -First 1
        if ($instance) {
            $result.S1 = @{ BaseUri = $instance.URL.TrimEnd('/'); ApiToken = $instance.APIKey }
        }
    }

    if (-not $result.Automate -and -not $result.S1) { return $null }
    return $result
}

$DefaultS1BaseUri = 'https://usea1-cw04edr.sentinelone.net'
$stored = Import-StoredConfig -Path $ConfigPath

# Resolve each setting: explicit parameter wins, then the store, then a prompt.
if (-not $S1BaseUri) {
    $S1BaseUri = if ($stored -and $stored.S1 -and $stored.S1.BaseUri) { $stored.S1.BaseUri } else { $DefaultS1BaseUri }
}
$S1BaseUri = $S1BaseUri.TrimEnd('/')

$seed = $null
if (-not $stored) {
    Write-Step "No stored configuration found at $ConfigPath" 'Yellow'
    $seed = Get-SeedConfig -S1Host ([Uri]$S1BaseUri).Host
    if ($seed) {
        Write-Host "  Found reusable settings in $($seed.Source)" -ForegroundColor Gray
        $answer = Read-Host "  Seed this script's configuration from that file? [Y/n]"
        if ($answer -and $answer -notmatch '^(y|yes)$') { $seed = $null }
    }
}

if (-not $AutomateServer) {
    $AutomateServer =
        if ($stored -and $stored.Automate.Server) { $stored.Automate.Server }
        elseif ($seed -and $seed.Automate) { $seed.Automate.Server }
        else { Read-Host 'Automate server FQDN (e.g. automate.contoso.com)' }
}
$AutomateServer = ($AutomateServer -replace '^https?://', '').TrimEnd('/')

if (-not $AutomateClientId) {
    $AutomateClientId =
        if ($stored -and $stored.Automate.ClientId) { $stored.Automate.ClientId }
        elseif ($seed -and $seed.Automate) { $seed.Automate.ClientId }
        else { Read-Host 'Automate integrator ClientId GUID (https://developer.connectwise.com/ClientID)' }
}

# Fail before the first call rather than surfacing a misleading auth error. A missing
# ClientId header is the single most common first-run failure on Automate 20 Patch 11+.
if ([string]::IsNullOrWhiteSpace($AutomateClientId)) {
    Write-Host 'ERROR: AutomateClientId is empty.' -ForegroundColor Red
    Write-Host '       Automate 20 Patch 11 and above require a ClientId header on every API request,' -ForegroundColor Red
    Write-Host '       including the token request. Register an integrator GUID at' -ForegroundColor Red
    Write-Host '       https://developer.connectwise.com/ClientID and re-run with -AutomateClientId.' -ForegroundColor Red
    exit 2
}

if (-not $AutomateCredential) {
    if ($stored -and $stored.Automate.Username -and $stored.Automate.Password) {
        $AutomateCredential = New-Object System.Management.Automation.PSCredential(
            $stored.Automate.Username, $stored.Automate.Password)
    }
    elseif ($seed -and $seed.Automate -and $seed.Automate.Username -and $seed.Automate.Password) {
        $seededPassword = ConvertTo-SecureString $seed.Automate.Password -AsPlainText -Force
        $AutomateCredential = New-Object System.Management.Automation.PSCredential(
            $seed.Automate.Username, $seededPassword)
    }
    else {
        $AutomateCredential = Get-Credential -Message "Automate API credential for $AutomateServer"
    }
}

if (-not $S1ApiToken) {
    if ($stored -and $stored.S1 -and $stored.S1.ApiToken) {
        $S1ApiToken = $stored.S1.ApiToken
    }
    elseif ($seed -and $seed.S1) {
        $S1ApiToken = ConvertTo-SecureString $seed.S1.ApiToken -AsPlainText -Force
    }
    else {
        $S1ApiToken = Read-Host -AsSecureString "SentinelOne API token for $S1BaseUri (account-scoped service user)"
    }
}

if (-not $AutomateCredential -or -not $S1ApiToken) {
    Write-Host 'ERROR: Automate credential and SentinelOne API token are both required.' -ForegroundColor Red
    exit 2
}

# Persist to this script's own store (DPAPI-encrypted by Export-Clixml, user+machine bound).
if (-not $stored) {
    $configDir = Split-Path -Parent $ConfigPath
    if (-not (Test-Path $configDir)) { New-Item -ItemType Directory -Path $configDir -Force | Out-Null }

    @{
        Automate = @{
            Server   = $AutomateServer
            Username = $AutomateCredential.UserName
            Password = $AutomateCredential.Password
            ClientId = $AutomateClientId
        }
        S1       = @{
            BaseUri  = $S1BaseUri
            ApiToken = $S1ApiToken
        }
    } | Export-Clixml -Path $ConfigPath

    Write-Step "Configuration saved to $ConfigPath" 'Green'
}

if (-not $OutputPath) { $OutputPath = Join-Path $PSScriptRoot 'Reports' }
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }

#endregion

Write-Host ''
Write-Step '=== SentinelOne <-> Automate reconciliation (READ-ONLY) ==='
Write-Host "  Automate      : $AutomateServer" -ForegroundColor Gray
Write-Host "  SentinelOne   : $S1BaseUri" -ForegroundColor Gray
Write-Host "  Scope         : companies present in both systems" -ForegroundColor Gray
Write-Host ''

#region Pull Automate clients

try {
    if (-not (Get-Module -ListAvailable -Name AutomateAPI)) {
        throw "The AutomateAPI module is not installed. Run: Install-Module AutomateAPI -Scope CurrentUser"
    }
    Import-Module AutomateAPI -ErrorAction Stop
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 2
}

try {
    Write-Step 'Connecting to Automate...'
    $connectParams = @{
        Server     = $AutomateServer
        Credential = $AutomateCredential
        ClientID   = $AutomateClientId
        Force      = $true
        Quiet      = $true
    }
    if ($AutomateTwoFactorToken) { $connectParams.TwoFactorToken = $AutomateTwoFactorToken }

    Connect-AutomateAPI @connectParams | Out-Null
}
catch {
    $message = $_.Exception.Message
    Write-Host "ERROR: Automate authentication failed against '$AutomateServer'." -ForegroundColor Red
    Write-Host "       $message" -ForegroundColor Red
    if ($message -match '401|403|[Uu]nauthorized|[Ff]orbidden') {
        Write-Host '       Check the username/password, the ClientId GUID, and whether the account' -ForegroundColor Yellow
        Write-Host '       requires a 2FA passcode (-AutomateTwoFactorToken).' -ForegroundColor Yellow
    }
    elseif ($message -match 'resolve|timed out|No such host|connection|SSL|TLS') {
        Write-Host "       Looks like a network/TLS problem reaching $AutomateServer, not a credential problem." -ForegroundColor Yellow
    }
    exit 2
}

try {
    Write-Step 'Pulling Automate clients...'
    $automateClients = @(Get-AutomateClient -AllClients)
}
catch {
    Write-Host "ERROR: Failed to retrieve Automate clients: $($_.Exception.Message)" -ForegroundColor Red
    exit 2
}

if ($automateClients.Count -eq 0) {
    Write-Host 'ERROR: Automate returned zero clients. The account may lack visibility to any client.' -ForegroundColor Red
    exit 2
}
Write-Host "  $($automateClients.Count) clients" -ForegroundColor Gray

# Agent counts define the in-scope set. The clients endpoint exposes no status field,
# so "has agents" is the only server-side signal for whether a client is live.
$agentCountByClientId = @{}
try {
    Write-Step 'Pulling Automate agent counts...'
    $automateComputers = @(Get-AutomateComputer -IncludeFields 'Id,ComputerName,Client')
    foreach ($computer in $automateComputers) {
        $clientId = $computer.Client.Id
        if ($null -eq $clientId) { continue }
        if ($agentCountByClientId.ContainsKey($clientId)) { $agentCountByClientId[$clientId]++ }
        else { $agentCountByClientId[$clientId] = 1 }
    }
    Write-Host "  $($automateComputers.Count) agents across $($agentCountByClientId.Keys.Count) clients" -ForegroundColor Gray
}
catch {
    # Agent counts are informational only; the match itself does not depend on them.
    Write-Warning "Could not retrieve Automate agent counts ($($_.Exception.Message)). Continuing without them."
}

#endregion

#region Pull SentinelOne sites

$s1Sites = @()
try {
    Write-Step 'Pulling SentinelOne sites...'

    $tokenBstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($S1ApiToken)
    try { $tokenPlain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($tokenBstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($tokenBstr) }

    $headers = @{ 'Authorization' = "ApiToken $tokenPlain" }
    $cursor = $null

    do {
        $uri = "$S1BaseUri/web/api/v2.1/sites?limit=1000"
        if ($cursor) { $uri += "&cursor=$([Uri]::EscapeDataString($cursor))" }

        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
        if ($response.data.sites) { $s1Sites += $response.data.sites }
        $cursor = $response.pagination.nextCursor
    } while ($cursor)
}
catch {
    $message = $_.Exception.Message
    Write-Host "ERROR: Failed to retrieve SentinelOne sites from $S1BaseUri." -ForegroundColor Red
    Write-Host "       $message" -ForegroundColor Red
    if ($message -match '401|403|[Uu]nauthorized|[Ff]orbidden') {
        Write-Host '       The API token was rejected. Confirm it is current (S1 tokens expire) and' -ForegroundColor Yellow
        Write-Host '       that it belongs to an account-scoped service user.' -ForegroundColor Yellow
    }
    exit 2
}
finally {
    if (Get-Variable -Name tokenPlain -Scope 0 -ErrorAction SilentlyContinue) {
        Remove-Variable -Name tokenPlain -Scope 0 -ErrorAction SilentlyContinue
    }
}

if ($s1Sites.Count -eq 0) {
    Write-Host 'ERROR: SentinelOne returned zero sites. Nothing to reconcile against.' -ForegroundColor Red
    exit 2
}

$stateSummary = ($s1Sites | Group-Object state | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', '
Write-Host "  $($s1Sites.Count) sites ($stateSummary)" -ForegroundColor Gray

# A site-scoped token returns only its own site, which would fail every match while
# looking like a data problem rather than a permissions problem.
if ($s1Sites.Count -le 1) {
    Write-Warning "Only $($s1Sites.Count) site returned. That is almost certainly a SITE-SCOPED API token."
    Write-Warning "Generate a token from an ACCOUNT-scoped service user (Settings > Users > Service Users) and re-run."
}

#endregion

#region Normalize and index

$normalizeArgs = @{
    Suffix        = $LegalSuffix
    Article       = $StripLeadingArticle
    MergeInitials = $CollapseInitials
}

# All sites participate in matching, including non-active ones. A decommissioned site
# sharing a name is exactly what produces a real AMBIGUOUS, and hiding it would let the
# migration silently enroll endpoints into a stale site later. State is annotated on
# every matched site so the operator can see which duplicate to retire.
# Buckets are plain arrays, not Generic.List: Windows PowerShell 5.1 throws
# "Argument types do not match" when @() wraps a List[object], and these buckets are
# read back through @(...) below. Groups are 1-2 items, so += costs nothing here.
$sitesByNormalizedName = @{}
foreach ($site in $s1Sites) {
    $normalized = Get-NormalizedName -Name $site.name @normalizeArgs
    if (-not $sitesByNormalizedName.ContainsKey($normalized)) {
        $sitesByNormalizedName[$normalized] = @()
    }
    $sitesByNormalizedName[$normalized] += $site
}

# Every Automate client is evaluated; the name match itself decides what is in scope.
# Filtering by agent count here would hide a company that exists on both sides but
# currently has no agents checked in, which is exactly a case worth mapping.
$inScopeClients = @($automateClients)

# Two different raw Automate names collapsing to one normalized string is a
# normalization-rule problem, not a data problem, and the operator needs to see it.
$clientsByNormalizedName = @{}
foreach ($client in $inScopeClients) {
    $normalized = Get-NormalizedName -Name $client.Name @normalizeArgs
    if (-not $clientsByNormalizedName.ContainsKey($normalized)) {
        $clientsByNormalizedName[$normalized] = @()
    }
    $clientsByNormalizedName[$normalized] += $client
}

$collidingClientNames = @{}
foreach ($key in $clientsByNormalizedName.Keys) {
    $group = $clientsByNormalizedName[$key]
    if ($group.Count -gt 1) {
        $distinctRaw = @($group | ForEach-Object { $_.Name } | Sort-Object -Unique)
        if ($distinctRaw.Count -gt 1) { $collidingClientNames[$key] = $distinctRaw }
    }
}

$s1NormalizedNames = @($sitesByNormalizedName.Keys | Where-Object { $_ })
$clientNormalizedNames = @($clientsByNormalizedName.Keys | Where-Object { $_ })

#endregion

#region Match

Write-Step 'Matching...'
$results = New-Object System.Collections.Generic.List[object]
$matchedNormalizedSiteNames = @{}

foreach ($client in $inScopeClients) {
    $normalized = Get-NormalizedName -Name $client.Name @normalizeArgs
    $agentCount = if ($agentCountByClientId.ContainsKey($client.Id)) { $agentCountByClientId[$client.Id] } else { 0 }

    # Three PowerShell traps converge on these four lines, so leave them as written:
    #   1. Test with ContainsKey, never @($hash[$key]).Count -- @($null).Count is 1,
    #      which would silently report every unmatched client as a clean match.
    #   2. Assign in a plain statement, never "$x = if (...) { @(...) }" -- an if block
    #      unrolls a one-element array to a scalar, and MatchCount then comes back empty.
    #   3. Name it siteMatches, not matches -- $Matches is a type-constrained automatic.
    $siteMatches = @()
    if ($normalized -and $sitesByNormalizedName.ContainsKey($normalized)) {
        $siteMatches = @($sitesByNormalizedName[$normalized])
    }

    $status = switch ($siteMatches.Count) {
        0 { 'NO_MATCH' }
        1 { 'OK' }
        default { 'AMBIGUOUS' }
    }

    $notes = New-Object System.Collections.Generic.List[string]
    if (-not $normalized) { $notes.Add('Client name normalizes to an empty string') }
    if ($collidingClientNames.ContainsKey($normalized)) {
        $notes.Add("Normalization collision with: $(($collidingClientNames[$normalized] | Where-Object { $_ -ne $client.Name }) -join ' | ')")
    }
    if ($status -eq 'AMBIGUOUS') {
        $states = @($siteMatches | ForEach-Object { $_.state } | Sort-Object -Unique)
        if ($states -contains 'active' -and @($states | Where-Object { $_ -ne 'active' }).Count -gt 0) {
            $notes.Add('Duplicate sites differ in state; retire the non-active one')
        }
    }

    $nearest = $null
    if ($status -eq 'NO_MATCH' -and $normalized) {
        $nearest = Get-ClosestName -Name $normalized -Candidates $s1NormalizedNames -Floor $MinimumSimilarity
    }

    $nearestSiteName = $null
    if ($nearest) {
        $nearestSiteName = ($sitesByNormalizedName[$nearest.Name] | Select-Object -First 1).name
    }

    if ($siteMatches.Count -gt 0) { $matchedNormalizedSiteNames[$normalized] = $true }

    $results.Add([PSCustomObject]@{
            AutomateClientId    = $client.Id
            AutomateClientName  = $client.Name
            NormalizedName      = $normalized
            AutomateAgentCount  = $agentCount
            MatchCount          = $siteMatches.Count
            Status              = $status
            # Singular id/name, populated only on an unambiguous match. This is the
            # deterministic install target for the downstream migration; the plural
            # columns below exist for the AMBIGUOUS case and must not be parsed as one id.
            S1SiteId            = $(if ($siteMatches.Count -eq 1) { $siteMatches[0].id } else { $null })
            S1SiteName          = $(if ($siteMatches.Count -eq 1) { $siteMatches[0].name } else { $null })
            S1SiteNames         = (@($siteMatches | ForEach-Object { $_.name }) -join '; ')
            S1SiteIds           = (@($siteMatches | ForEach-Object { $_.id }) -join '; ')
            S1SiteStates        = (@($siteMatches | ForEach-Object { $_.state }) -join '; ')
            S1ActiveLicenses    = (@($siteMatches | ForEach-Object { $_.activeLicenses }) -join '; ')
            NearestS1Site       = $nearestSiteName
            NearestS1Distance   = if ($nearest) { $nearest.Distance } else { $null }
            NearestS1Similarity = if ($nearest) { $nearest.Similarity } else { $null }
            Notes               = ($notes -join ' / ')
        })
}

# Reverse check. A stale or duplicate S1 site is a common cause of a later AMBIGUOUS,
# so orphans are worth surfacing before the migration rather than after.
$orphanSites = New-Object System.Collections.Generic.List[object]
foreach ($site in $s1Sites) {
    $normalized = Get-NormalizedName -Name $site.name @normalizeArgs
    if ($matchedNormalizedSiteNames.ContainsKey($normalized)) { continue }

    $nearest = Get-ClosestName -Name $normalized -Candidates $clientNormalizedNames -Floor $MinimumSimilarity
    $nearestClientName = $null
    if ($nearest) {
        $nearestClientName = ($clientsByNormalizedName[$nearest.Name] | Select-Object -First 1).Name
    }

    $notes = New-Object System.Collections.Generic.List[string]
    if ($site.isDefault) { $notes.Add('SentinelOne default site') }
    if ($site.state -ne 'active') { $notes.Add("Site state is '$($site.state)'") }
    if ($site.name -match '^[0-9A-F]{12}-S1SW$') { $notes.Add('Auto-generated site name; likely a stray enrollment') }

    $orphanSites.Add([PSCustomObject]@{
            S1SiteId                  = $site.id
            S1SiteName                = $site.name
            NormalizedName            = $normalized
            State                     = $site.state
            SiteType                  = $site.siteType
            ActiveLicenses            = $site.activeLicenses
            TotalLicenses             = $site.totalLicenses
            CreatedAt                 = $site.createdAt
            NearestAutomateClient     = $nearestClientName
            NearestAutomateDistance   = if ($nearest) { $nearest.Distance } else { $null }
            NearestAutomateSimilarity = if ($nearest) { $nearest.Similarity } else { $null }
            Notes                     = ($notes -join ' / ')
        })
}

#endregion

#region Report

$okRows = @($results | Where-Object { $_.Status -eq 'OK' })
$noMatchRows = @($results | Where-Object { $_.Status -eq 'NO_MATCH' })
$ambiguousRows = @($results | Where-Object { $_.Status -eq 'AMBIGUOUS' })

# An unmatched client whose closest site is very similar is probably a rename, not a
# company that is genuinely absent from S1. Scoping to the intersection would drop it
# silently, so it gets called out separately -- as a warning, never as a gate.
$renameHintRows = @(
    $noMatchRows | Where-Object {
        $_.NearestS1Similarity -and [double]$_.NearestS1Similarity -ge $RenameHintSimilarity
    }
)

Write-Host ''
Write-Step '--- Summary ---'
Write-Host ("  {0,-34} {1,5}" -f 'Mapped 1:1 (in both systems)', $okRows.Count) -ForegroundColor Green
Write-Host ("  {0,-34} {1,5}" -f 'AMBIGUOUS (blocks migration)', $ambiguousRows.Count) -ForegroundColor $(if ($ambiguousRows.Count) { 'Red' } else { 'Green' })
Write-Host ''
Write-Host ("  {0,-34} {1,5}" -f 'Automate-only (no S1 site)', $noMatchRows.Count) -ForegroundColor Gray
Write-Host ("  {0,-34} {1,5}" -f 'S1-only (orphan sites)', $orphanSites.Count) -ForegroundColor Gray
Write-Host ("  {0,-34} {1,5}" -f 'Automate clients examined', $results.Count) -ForegroundColor Gray
Write-Host ("  {0,-34} {1,5}" -f 'S1 sites examined', $s1Sites.Count) -ForegroundColor Gray

if ($renameHintRows.Count -gt 0) {
    Write-Host ''
    Write-Step "--- POSSIBLE RENAMES: unmatched, but a very similar S1 site exists (>= $RenameHintSimilarity) ---" 'Yellow'
    Write-Host '    Advisory only. If one of these is genuinely the same company, align the names' -ForegroundColor Gray
    Write-Host '    so it joins the mapped set; otherwise ignore it.' -ForegroundColor Gray
    $renameHintRows |
        Sort-Object -Property @{Expression = { [double]$_.NearestS1Similarity }; Descending = $true } |
        Format-Table AutomateClientName, AutomateAgentCount, NearestS1Site, NearestS1Similarity -AutoSize -Wrap |
        Out-String -Width 4096 | Write-Host
}

if ($ambiguousRows.Count -gt 0) {
    Write-Host ''
    Write-Step '--- AMBIGUOUS: more than one SentinelOne site matches (resolve by hand) ---' 'Red'
    $ambiguousRows |
        Sort-Object AutomateClientName |
        Format-Table AutomateClientName, AutomateAgentCount, MatchCount, S1SiteNames, S1SiteIds, S1SiteStates, Notes -AutoSize -Wrap |
        Out-String -Width 4096 | Write-Host
}

if ($ReportUnmatchedClients -and $noMatchRows.Count -gt 0) {
    Write-Host ''
    Write-Step '--- Automate clients with no SentinelOne site (out of scope, informational) ---' 'Gray'
    $noMatchRows |
        Sort-Object -Property @{Expression = 'AutomateAgentCount'; Descending = $true }, AutomateClientName |
        Format-Table AutomateClientName, AutomateAgentCount, NormalizedName, NearestS1Site, NearestS1Similarity -AutoSize -Wrap |
        Out-String -Width 4096 | Write-Host
}

if ($orphanSites.Count -gt 0) {
    Write-Host ''
    Write-Step '--- ORPHAN SentinelOne sites: no Automate client matched ---' 'Yellow'
    Write-Host '    Informational: these do not affect the exit code. Sites still holding active' -ForegroundColor Gray
    Write-Host '    licenses are the ones worth chasing.' -ForegroundColor Gray
    $orphanSites |
        Sort-Object -Property @{Expression = 'ActiveLicenses'; Descending = $true }, S1SiteName |
        Format-Table S1SiteName, State, ActiveLicenses, NearestAutomateClient, NearestAutomateSimilarity, Notes -AutoSize -Wrap |
        Out-String -Width 4096 | Write-Host
}

if ($collidingClientNames.Keys.Count -gt 0) {
    Write-Host ''
    Write-Step '--- Normalization collisions: different Automate names reduce to one string ---' 'Yellow'
    Write-Host '    These are a normalization-rule problem. Tighten -LegalSuffix if they are false collisions.' -ForegroundColor Gray
    foreach ($key in ($collidingClientNames.Keys | Sort-Object)) {
        Write-Host ("    '{0}' <- {1}" -f $key, ($collidingClientNames[$key] -join ' | ')) -ForegroundColor Yellow
    }
}

if ($ShowMatched -and $okRows.Count -gt 0) {
    Write-Host ''
    Write-Step '--- OK ---' 'Green'
    $okRows |
        Sort-Object AutomateClientName |
        Format-Table AutomateClientName, AutomateAgentCount, S1SiteNames, S1SiteIds, S1SiteStates, S1ActiveLicenses -AutoSize -Wrap |
        Out-String -Width 4096 | Write-Host
}

$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$resultCsv = Join-Path $OutputPath "${timestamp}_S1_Automate_Reconciliation.csv"
$orphanCsv = Join-Path $OutputPath "${timestamp}_S1_Automate_OrphanSites.csv"
$siteMapCsv = Join-Path $OutputPath "${timestamp}_S1_Automate_SiteMap.csv"

$results | Sort-Object Status, AutomateClientName | Export-Csv -Path $resultCsv -NoTypeInformation -Encoding UTF8
Write-Host ''
Write-Step "Full result set : $resultCsv" 'Green'

if ($orphanSites.Count -gt 0) {
    $orphanSites | Sort-Object S1SiteName | Export-Csv -Path $orphanCsv -NoTypeInformation -Encoding UTF8
    Write-Step "Orphan S1 sites : $orphanCsv" 'Green'
}

# Install map: the OK rows only, reduced to the pair the migration phase needs --
# Automate client id in, SentinelOne site id out. NO_MATCH and AMBIGUOUS rows are
# deliberately absent so a downstream run cannot enroll endpoints into a guessed site.
if ($okRows.Count -gt 0) {
    $siteById = @{}
    foreach ($site in $s1Sites) { $siteById[[string]$site.id] = $site }

    $siteMap = foreach ($row in ($okRows | Sort-Object AutomateClientName)) {
        $site = $siteById[[string]$row.S1SiteId]

        $entry = [ordered]@{
            AutomateClientId   = $row.AutomateClientId
            AutomateClientName = $row.AutomateClientName
            AutomateAgentCount = $row.AutomateAgentCount
            S1SiteId           = $row.S1SiteId
            S1SiteName         = $row.S1SiteName
            S1SiteState        = $site.state
            S1ActiveLicenses   = $site.activeLicenses
        }
        # The registration token is the per-site enrollment secret the installer consumes.
        # Off by default: it turns a diagnostic report into a credential-bearing file.
        if ($IncludeRegistrationToken) { $entry.S1RegistrationToken = $site.registrationToken }

        [PSCustomObject]$entry
    }

    $siteMap | Export-Csv -Path $siteMapCsv -NoTypeInformation -Encoding UTF8
    Write-Step "Install site map : $siteMapCsv ($($okRows.Count) clients)" 'Green'

    if ($IncludeRegistrationToken) {
        Write-Warning "$siteMapCsv contains SentinelOne site registration tokens. Treat it as a secret and delete it when the migration is done."
    }
    else {
        Write-Host '                   (add -IncludeRegistrationToken to embed each site enrollment token)' -ForegroundColor Gray
    }
}

#endregion

# The gate is AMBIGUOUS only. A company absent from one side is out of scope, not a
# fault, and must not block the migration for the companies that do map cleanly.
Write-Host ''
if ($ambiguousRows.Count -gt 0) {
    Write-Host ("NOT SAFE TO MIGRATE: {0} company/companies match more than one SentinelOne site." -f $ambiguousRows.Count) -ForegroundColor Red
    Write-Host 'Resolve the duplicate sites listed above, then re-run. Ambiguity is never auto-resolved.' -ForegroundColor Red
    $script:ExitCode = 1
}
else {
    Write-Host ("SAFE TO MIGRATE: all {0} companies present in both systems map to exactly one SentinelOne site." -f $okRows.Count) -ForegroundColor Green
    if ($renameHintRows.Count -gt 0) {
        Write-Host ("Note: {0} unmatched client(s) closely resemble an S1 site name -- review the possible-rename list above." -f $renameHintRows.Count) -ForegroundColor Yellow
    }
    $script:ExitCode = 0
}

exit $script:ExitCode
