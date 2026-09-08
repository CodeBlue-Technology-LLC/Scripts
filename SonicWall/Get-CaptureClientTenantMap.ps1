<#
.SYNOPSIS
    READ-ONLY: builds the SonicWall Capture Client tenant -> Automate client map and
    collects each tenant's Cleaner Utility authentication code.

.DESCRIPTION
    Pulls every Capture Client console tenant, fetches the per-tenant "Authentication
    Code for Tenant" (the value shown under Management > Client Installers >
    Utilities, passed to SonicWallCaptureClientCleanerUtility.exe -authCode), and
    resolves each tenant to exactly one ConnectWise Automate client by normalized
    company name.

    ############################################################################
    #  THIS SCRIPT IS STRICTLY READ-ONLY.                                      #
    #  It performs GET operations against SonicWall and Automate only. Nothing  #
    #  is written to either system. Writing the codes into Automate EDFs is a   #
    #  separate script: Set-AutomateCaptureClientAuthCode.ps1.                  #
    ############################################################################

    The map is the deliverable, and it deliberately contains only 1:1 matches, so the
    writer can never stamp a guessed client. Tenants that resolve to zero or to more
    than one Automate client are reported and excluded.

    Names are joined after normalization (lowercase, punctuation folded, legal
    suffixes and a leading article dropped). Company names that will never agree
    across the two systems are pinned by hand in the override CSV rather than by
    loosening normalization for everyone.

    Outputs, all timestamped, into Reports\:
      1. *_SWCC_TenantMap.csv     - the deliverable. One row per 1:1 match:
                                    TenantId, TenantName, AutomateClientId,
                                    AutomateClientName, CleanerAuthCode.
      2. *_SWCC_TenantMap_All.csv - every tenant with its status, for triage.

    Exit codes:
      0 = every tenant resolved 1:1 (safe to run the writer)
      1 = at least one tenant is UNMATCHED or AMBIGUOUS; the map is partial
      2 = fatal error (auth, network, empty result, missing prerequisite)

    An exit code of 1 is normal here: not every Capture Client tenant has a
    corresponding Automate client. It means "review the exclusions", not "stop".

.NOTES
    Run under Windows PowerShell 5.1 (powershell.exe), not PowerShell 7.

    Prerequisites:
      - AutomateAPI module (Install-Module AutomateAPI)
      - A registered Automate integrator ClientId GUID
        (https://developer.connectwise.com/ClientID). Mandatory on Automate 20
        Patch 11 and above; auth fails without it regardless of credential validity.
      - A MySonicWall API key whose user can see the Capture Client tenants
        (MySonicWall > My Account > API Key). Read-only rights are sufficient.

    The Automate account only needs READ access here. The writer script needs more.

.PARAMETER MySonicWallApiKey
    Override the stored MySonicWall API key for this run. Accepts a plain string or a
    SecureString; it is never written back to the credential store.

.PARAMETER ConsoleBaseUri
    Capture Client console base URI. Default 'https://captureclient-36.sonicwall.com'
    (US). EU tenants use 'https://captureclient-36eu.sonicwall.com'.

.PARAMETER AutomateServer
    Automate server FQDN. Taken from the credential store when omitted.

.PARAMETER AutomateCredential
    Override the stored Automate credential.

.PARAMETER AutomateClientId
    Automate integrator ClientId GUID. Taken from the credential store when omitted.

.PARAMETER OverrideCsv
    CSV of manual tenant -> client pins, applied before name matching and always
    winning over it. Columns: TenantName, AutomateClientId. TenantName may be given
    with or without the -CCxxxxxxxx console suffix. Defaults to
    Config\CaptureClient-ClientOverrides.csv when that file exists.

.PARAMETER LegalSuffix
    Standalone tokens stripped during normalization. Word-boundary matched, so 'co'
    does not turn 'Coastal' into 'astal'.

.PARAMETER OutputDir
    Where the CSVs are written. Defaults to Reports\ beside this script.

.PARAMETER ReportUnmatchedClients
    Also list Automate clients that have no Capture Client tenant. Off by default:
    that list is large, expected, and not actionable.

.EXAMPLE
    .\Get-CaptureClientTenantMap.ps1

    Full read-only run. Prompts once for anything not already stored.

.EXAMPLE
    .\Get-CaptureClientTenantMap.ps1 -ConsoleBaseUri https://captureclient-36eu.sonicwall.com

    Run against the EU console.
#>

[CmdletBinding()]
param(
    [object]$MySonicWallApiKey,

    [string]$ConsoleBaseUri,

    [string]$AutomateServer,

    [System.Management.Automation.PSCredential]$AutomateCredential,

    [string]$AutomateClientId,

    [string]$OverrideCsv,

    [string[]]$LegalSuffix = @('inc', 'llc', 'ltd', 'corp', 'co', 'pc', 'pllc', 'plc', 'llp', 'lp'),

    [string[]]$StripLeadingArticle = @('the'),

    [string]$OutputDir,

    [switch]$ReportUnmatchedClients
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

. (Join-Path $PSScriptRoot 'CaptureClientApi.ps1')

$ConfigPath = Join-Path $PSScriptRoot 'Config\CaptureClient-Credentials.xml'
$DefaultOverrideCsv = Join-Path $PSScriptRoot 'Config\CaptureClient-ClientOverrides.csv'
if (-not $OutputDir) { $OutputDir = Join-Path $PSScriptRoot 'Reports' }

$script:ExitCode = 0

function Write-Step { param([string]$Message, [string]$Color = 'Cyan') Write-Host $Message -ForegroundColor $Color }

#region Configuration

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
    $seedPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'SentinelOne\Config\S1-AutomateReconcile-Credentials.xml'
    if (-not (Test-Path $seedPath)) { return $null }

    try { $seed = Import-Clixml -Path $seedPath } catch { return $null }
    if (-not $seed.Automate -or -not $seed.Automate.Server) { return $null }

    return @{
        Source   = $seedPath
        Automate = @{
            Server   = $seed.Automate.Server
            Username = $seed.Automate.Username
            Password = $seed.Automate.Password
            ClientId = $seed.Automate.ClientId
        }
    }
}

$stored = Import-StoredConfig -Path $ConfigPath
$seed = $null
if (-not $stored) {
    Write-Step "No stored configuration found at $ConfigPath" 'Yellow'
    $seed = Get-SeedConfig
    if ($seed) {
        Write-Host "  Found reusable Automate settings in $($seed.Source)" -ForegroundColor Gray
        $answer = Read-Host "  Seed this script's Automate configuration from that file? [Y/n]"
        if ($answer -and $answer -notmatch '^(y|yes)$') { $seed = $null }
    }
}

if (-not $ConsoleBaseUri) {
    $ConsoleBaseUri =
        if ($stored -and $stored.CaptureClient -and $stored.CaptureClient.BaseUri) { $stored.CaptureClient.BaseUri }
        else { 'https://captureclient-36.sonicwall.com' }
}
$ConsoleBaseUri = $ConsoleBaseUri.TrimEnd('/')

if (-not $AutomateServer) {
    $AutomateServer =
        if ($stored -and $stored.Automate.Server) { $stored.Automate.Server }
        elseif ($seed) { $seed.Automate.Server }
        else { Read-Host 'Automate server FQDN (e.g. automate.contoso.com)' }
}
$AutomateServer = ($AutomateServer -replace '^https?://', '').TrimEnd('/')

if (-not $AutomateClientId) {
    $AutomateClientId =
        if ($stored -and $stored.Automate.ClientId) { $stored.Automate.ClientId }
        elseif ($seed) { $seed.Automate.ClientId }
        else { Read-Host 'Automate integrator ClientId GUID' }
}

if (-not $AutomateCredential) {
    if ($stored -and $stored.Automate.Username) {
        $AutomateCredential = New-Object System.Management.Automation.PSCredential($stored.Automate.Username, $stored.Automate.Password)
    }
    elseif ($seed) {
        $AutomateCredential = New-Object System.Management.Automation.PSCredential($seed.Automate.Username, $seed.Automate.Password)
    }
    else {
        $AutomateCredential = Get-Credential -Message 'Automate credential (read access is enough for this script)'
    }
}

# The MySonicWall key is the one secret this script owns outright.
$apiKeySecure = $null
if ($MySonicWallApiKey) {
    $apiKeySecure = if ($MySonicWallApiKey -is [System.Security.SecureString]) { $MySonicWallApiKey }
                    else { ConvertTo-SecureString ([string]$MySonicWallApiKey) -AsPlainText -Force }
}
elseif ($stored -and $stored.MySonicWall -and $stored.MySonicWall.ApiKey) {
    $apiKeySecure = $stored.MySonicWall.ApiKey
}
else {
    $apiKeySecure = Read-Host 'MySonicWall API key' -AsSecureString
}

$storedProductGroupId = if ($stored -and $stored.CaptureClient -and $stored.CaptureClient.ProductGroupId) {
    [string]$stored.CaptureClient.ProductGroupId
} else { $null }

#endregion

Write-Host ''
Write-Step '=== Capture Client tenant map -- READ ONLY ==='
Write-Host ("  Console   : {0}" -f $ConsoleBaseUri) -ForegroundColor Gray
Write-Host ("  Automate  : {0} as {1}" -f $AutomateServer, $AutomateCredential.UserName) -ForegroundColor Gray
Write-Host ''

#region Capture Client

try {
    if (-not $storedProductGroupId) {
        Write-Step 'Locating a Capture Client ProductGroupID in MySonicWall...'
        $mswTenants = Get-MswCloudTenant -ApiKey $apiKeySecure
        Write-Host ("  MySonicWall tenants visible: {0}" -f $mswTenants.Count) -ForegroundColor Gray

        $storedProductGroupId = Get-CaptureClientProductGroupId -Tenant $mswTenants
        if (-not $storedProductGroupId) {
            throw 'No MySonicWall tenant visible to this API key has Capture Client enabled.'
        }
        Write-Host ("  Using ProductGroupID {0}" -f $storedProductGroupId) -ForegroundColor Gray
    }

    Write-Step 'Authenticating to Capture Client...'
    $session = Connect-CaptureClient -ApiKey $apiKeySecure -ProductGroupId $storedProductGroupId -BaseUri $ConsoleBaseUri
    Write-Host ("  Signed in as {0}; token valid until {1:yyyy-MM-dd HH:mm}" -f $session.Email, $session.Expires) -ForegroundColor Gray
}
catch {
    Write-Host "ERROR: Capture Client authentication failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 2
}

try {
    Write-Step 'Retrieving tenants...'
    $ccTenants = @(Get-CaptureClientTenant -Session $session)
}
catch {
    Write-Host "ERROR: Could not list Capture Client tenants: $($_.Exception.Message)" -ForegroundColor Red
    exit 2
}

if ($ccTenants.Count -eq 0) {
    Write-Host 'ERROR: Capture Client returned no tenants.' -ForegroundColor Red
    exit 2
}
Write-Host ("  Tenants: {0}" -f $ccTenants.Count) -ForegroundColor Gray

Write-Step 'Retrieving per-tenant authentication codes...'
$authCode = @{}
$codeFailures = @()
$done = 0
foreach ($tenant in $ccTenants) {
    $done++
    Write-Progress -Activity 'Capture Client authentication codes' -Status $tenant.tenantName `
        -PercentComplete ([int](100 * $done / $ccTenants.Count))
    try {
        $code = Get-CaptureClientCleanerAuthCode -Session $session -TenantId $tenant.tenantId
        if ($code -match '^[0-9A-Fa-f]{32,}$') { $authCode[$tenant.tenantId] = $code.ToUpperInvariant() }
        else { $codeFailures += [pscustomobject]@{ TenantName = $tenant.tenantName; Reason = "Unexpected value '$code'" } }
    }
    catch {
        $codeFailures += [pscustomobject]@{ TenantName = $tenant.tenantName; Reason = $_.Exception.Message }
    }
}
Write-Progress -Activity 'Capture Client authentication codes' -Completed
Write-Host ("  Codes retrieved: {0} of {1}" -f $authCode.Count, $ccTenants.Count) -ForegroundColor Gray

if ($codeFailures.Count -gt 0) {
    Write-Host ''
    Write-Warning "Could not read an authentication code for $($codeFailures.Count) tenant(s):"
    $codeFailures | ForEach-Object { Write-Host ("    {0}: {1}" -f $_.TenantName, $_.Reason) -ForegroundColor Yellow }
    $script:ExitCode = 1
}

#endregion

#region Automate

try {
    if (-not (Get-Module -ListAvailable -Name AutomateAPI)) {
        throw 'The AutomateAPI module is not installed. Run: Install-Module AutomateAPI -Scope CurrentUser'
    }
    Import-Module AutomateAPI -ErrorAction Stop

    Write-Step 'Connecting to Automate...'
    Connect-AutomateAPI -Server $AutomateServer -Credential $AutomateCredential -ClientID $AutomateClientId -Force -Quiet | Out-Null

    $automateClients = @(Get-AutomateClient -AllClients)
}
catch {
    Write-Host "ERROR: Automate connection failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 2
}

if ($automateClients.Count -eq 0) {
    Write-Host 'ERROR: Automate returned no clients.' -ForegroundColor Red
    exit 2
}
Write-Host ("  Clients: {0}" -f $automateClients.Count) -ForegroundColor Gray

#endregion

#region Overrides

if (-not $OverrideCsv -and (Test-Path $DefaultOverrideCsv)) { $OverrideCsv = $DefaultOverrideCsv }

$overrideByName = @{}
if ($OverrideCsv) {
    if (-not (Test-Path $OverrideCsv)) {
        Write-Host "ERROR: Override CSV not found: $OverrideCsv" -ForegroundColor Red
        exit 2
    }
    $clientById = @{}
    foreach ($client in $automateClients) { $clientById["$($client.Id)"] = $client }

    foreach ($row in @(Import-Csv -Path $OverrideCsv)) {
        if (-not $row.TenantName -or -not $row.AutomateClientId) { continue }
        $key = (Get-BareTenantName -TenantName $row.TenantName).ToLowerInvariant()
        if (-not $clientById.ContainsKey("$($row.AutomateClientId)")) {
            Write-Warning "Override for '$($row.TenantName)' points at Automate client $($row.AutomateClientId), which does not exist. Ignored."
            continue
        }
        $overrideByName[$key] = $clientById["$($row.AutomateClientId)"]
    }
    Write-Host ("  Overrides loaded: {0} from {1}" -f $overrideByName.Count, (Split-Path -Leaf $OverrideCsv)) -ForegroundColor Gray
}

#endregion

#region Match

Write-Host ''
Write-Step 'Matching tenants to clients...'

$normalizeArgs = @{ Suffix = $LegalSuffix; Article = $StripLeadingArticle; MergeInitials = $true }

$clientsByNormalizedName = @{}
foreach ($client in $automateClients) {
    $normalized = Get-NormalizedCompanyName -Name $client.Name @normalizeArgs
    if (-not $normalized) { continue }
    if (-not $clientsByNormalizedName.ContainsKey($normalized)) { $clientsByNormalizedName[$normalized] = @() }
    $clientsByNormalizedName[$normalized] += $client
}

$results = @()
$matchedClientIds = @{}

foreach ($tenant in $ccTenants) {
    $bareName = Get-BareTenantName -TenantName $tenant.tenantName
    $normalized = Get-NormalizedCompanyName -Name $bareName @normalizeArgs
    $code = if ($authCode.ContainsKey($tenant.tenantId)) { $authCode[$tenant.tenantId] } else { '' }

    $candidates = @()
    $source = 'name'
    if ($overrideByName.ContainsKey($bareName.ToLowerInvariant())) {
        $candidates = @($overrideByName[$bareName.ToLowerInvariant()])
        $source = 'override'
    }
    elseif ($clientsByNormalizedName.ContainsKey($normalized)) {
        $candidates = @($clientsByNormalizedName[$normalized])
    }

    $status =
        if ($candidates.Count -eq 0) { 'UNMATCHED' }
        elseif ($candidates.Count -gt 1) { 'AMBIGUOUS' }
        elseif (-not $code) { 'NO_AUTH_CODE' }
        else { 'MATCHED' }

    if ($status -eq 'MATCHED') { $matchedClientIds["$($candidates[0].Id)"] = $true }

    $results += [pscustomobject]@{
        Status             = $status
        MatchSource        = if ($candidates.Count -eq 1) { $source } else { '' }
        TenantId           = $tenant.tenantId
        TenantName         = $tenant.tenantName
        TenantNameBare     = $bareName
        NormalizedName     = $normalized
        AutomateClientId   = if ($candidates.Count -eq 1) { $candidates[0].Id } else { '' }
        AutomateClientName = if ($candidates.Count -eq 1) { $candidates[0].Name } else { ($candidates | ForEach-Object { $_.Name }) -join ' | ' }
        CleanerAuthCode    = $code
        DevicesTotal       = $tenant.devicesTotal
        LicensesTotal      = $tenant.licensesTotal
        LicenseStatus      = $tenant.licenseStatus
    }
}

#endregion

#region Report

$byStatus = @{}
foreach ($group in ($results | Group-Object Status)) { $byStatus[$group.Name] = @($group.Group) }
function Get-Bucket { param([string]$Name) if ($byStatus.ContainsKey($Name)) { $byStatus[$Name] } else { @() } }

$matched = @(Get-Bucket 'MATCHED')
$ambiguous = @(Get-Bucket 'AMBIGUOUS')
$unmatched = @(Get-Bucket 'UNMATCHED')
$noCode = @(Get-Bucket 'NO_AUTH_CODE')

Write-Host ''
Write-Step '--- Summary ---'
Write-Host ("  Capture Client tenants : {0}" -f $results.Count)
Write-Host ("  Matched 1:1            : {0}" -f $matched.Count) -ForegroundColor Green
if ($ambiguous.Count) { Write-Host ("  Ambiguous              : {0}" -f $ambiguous.Count) -ForegroundColor Red }
if ($unmatched.Count) { Write-Host ("  Unmatched              : {0}" -f $unmatched.Count) -ForegroundColor Yellow }
if ($noCode.Count) { Write-Host ("  Matched, no auth code  : {0}" -f $noCode.Count) -ForegroundColor Yellow }

if ($ambiguous.Count -gt 0) {
    Write-Host ''
    Write-Step '--- AMBIGUOUS (fix these; they are excluded from the map) ---' 'Red'
    $ambiguous | Sort-Object TenantNameBare |
        Format-Table -AutoSize TenantNameBare, AutomateClientName, DevicesTotal | Out-String -Width 200 | Write-Host
    $script:ExitCode = 1
}

if ($unmatched.Count -gt 0) {
    Write-Host ''
    Write-Step '--- UNMATCHED (no Automate client resolved) ---' 'Yellow'
    Write-Host '    Pin any of these by hand in the override CSV:' -ForegroundColor Gray
    Write-Host ("    {0}" -f $DefaultOverrideCsv) -ForegroundColor Gray
    Write-Host ''
    $unmatched | Sort-Object -Descending { [int]$_.DevicesTotal } |
        Format-Table -AutoSize TenantNameBare, DevicesTotal, LicensesTotal | Out-String -Width 200 | Write-Host
    $script:ExitCode = 1
}

if ($ReportUnmatchedClients) {
    $orphans = @($automateClients | Where-Object { -not $matchedClientIds.ContainsKey("$($_.Id)") })
    Write-Host ''
    Write-Step ("--- Automate clients with no Capture Client tenant ({0}) ---" -f $orphans.Count) 'Gray'
    $orphans | Sort-Object Name | Format-Table -AutoSize Id, Name | Out-String -Width 200 | Write-Host
}

#endregion

#region Output

if (-not (Test-Path $OutputDir)) { $null = New-Item -ItemType Directory -Path $OutputDir -Force }
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'

$allPath = Join-Path $OutputDir "${stamp}_SWCC_TenantMap_All.csv"
$results | Sort-Object Status, TenantNameBare | Export-Csv -Path $allPath -NoTypeInformation -Encoding UTF8

$mapPath = Join-Path $OutputDir "${stamp}_SWCC_TenantMap.csv"
$matched | Sort-Object AutomateClientName |
    Select-Object TenantId, TenantName, AutomateClientId, AutomateClientName, CleanerAuthCode |
    Export-Csv -Path $mapPath -NoTypeInformation -Encoding UTF8

Write-Host ''
Write-Step '--- Files ---'
Write-Host ("  Map (1:1 only) : {0}" -f $mapPath) -ForegroundColor Gray
Write-Host ("  All tenants    : {0}" -f $allPath) -ForegroundColor Gray

#endregion

#region Persist configuration

$toStore = @{
    MySonicWall   = @{ ApiKey = $apiKeySecure }
    CaptureClient = @{ BaseUri = $ConsoleBaseUri; ProductGroupId = $storedProductGroupId }
    Automate      = @{
        Server   = $AutomateServer
        Username = $AutomateCredential.UserName
        Password = $AutomateCredential.Password
        ClientId = $AutomateClientId
    }
}
$configDir = Split-Path -Parent $ConfigPath
if (-not (Test-Path $configDir)) { $null = New-Item -ItemType Directory -Path $configDir -Force }
$toStore | Export-Clixml -Path $ConfigPath -Force

#endregion

Write-Host ''
if ($script:ExitCode -eq 0) {
    Write-Step 'Every Capture Client tenant resolved to exactly one Automate client.' 'Green'
}
else {
    Write-Step 'Map written, with exclusions. Review the lists above before writing.' 'Yellow'
}
Write-Host ''
Write-Host 'Next: .\Set-AutomateCaptureClientAuthCode.ps1        (dry run)' -ForegroundColor Gray
Write-Host '      .\Set-AutomateCaptureClientAuthCode.ps1 -Apply (writes)' -ForegroundColor Gray
Write-Host ''

exit $script:ExitCode
