<#
.SYNOPSIS
    READ-ONLY export of every agent's uninstall passphrase from a SentinelOne console.

.DESCRIPTION
    Produces a CSV of device name -> uninstall passphrase, for use when manually removing
    agents from a console you are migrating away from.

    Console-agnostic: point it at any SentinelOne-compatible console with -BaseUri. That
    includes SentinelOne OEM products such as SonicWall Capture Client, provided the
    tenant exposes the v2.1 API and the token carries agent-passphrase permission.

    #########################################################################
    #  READ-ONLY. Performs GET requests only. Does not uninstall, move,     #
    #  decommission, or modify any agent.                                   #
    #########################################################################

    The output contains uninstall passphrases in plaintext. Anyone holding that file can
    remove protection from the listed endpoints, so treat it as a credential: keep it off
    shared drives, and delete it once the migration is signed off.

.PARAMETER BaseUri
    Console base URI, e.g. https://usea1-cw04edr.sentinelone.net or the Capture Client
    management URL.

.PARAMETER ApiToken
    API token as a SecureString. Prompted for if omitted.

.PARAMETER OutputPath
    Output CSV. Defaults to a timestamped file in Reports\.

.PARAMETER IncludeDecommissioned
    Include agents already decommissioned in the console. Off by default -- they have no
    live endpoint to uninstall.

.PARAMETER SiteName
    Only export agents in these site names. Useful for migrating one customer at a time.

.EXAMPLE
    .\Export-S1AgentPassphrases.ps1 -BaseUri https://captureclient.sonicwall.com

    Prompts for the token, exports every active agent's passphrase.

.EXAMPLE
    .\Export-S1AgentPassphrases.ps1 -BaseUri https://captureclient.sonicwall.com -SiteName 'Acme Health'

    Exports one site's worth, for a phased cutover.

.NOTES
    Verified against a live SentinelOne v2.1 console: GET /web/api/v2.1/agents/passphrases
    returns computerName, passphrase, uuid, domain, lastLoggedInUserName and paginates via
    pagination.nextCursor. Agent metadata (site, OS, last seen) is joined from /agents so
    the CSV doubles as a migration worksheet.

    A token lacking passphrase rights typically returns 401/403 on the passphrases route
    while /agents still works -- the script calls that out specifically, because it is
    easy to mistake for a bad token.

    Run under Windows PowerShell 5.1.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$BaseUri,

    [System.Security.SecureString]$ApiToken,

    [string]$OutputPath,

    [switch]$IncludeDecommissioned,

    [string[]]$SiteName
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$BaseUri = $BaseUri.TrimEnd('/')
if ($BaseUri -notmatch '^https?://') { $BaseUri = "https://$BaseUri" }

if (-not $ApiToken) {
    $ApiToken = Read-Host -AsSecureString "API token for $BaseUri"
}

$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ApiToken)
try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
$headers = @{ Authorization = "ApiToken $plain" }

function Get-AllPages {
    <#
        Every v2.1 collection paginates the same way: data[] plus pagination.nextCursor.
        Returns the accumulated data rows.
    #>
    param([string]$Path, [int]$Limit = 1000)

    $rows = @()
    $cursor = $null
    do {
        $uri = "$BaseUri/web/api/v2.1/$Path"
        $uri += $(if ($Path -match '\?') { '&' } else { '?' })
        $uri += "limit=$Limit"
        if ($cursor) { $uri += "&cursor=$([Uri]::EscapeDataString($cursor))" }

        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
        if ($response.data) { $rows += $response.data }
        $cursor = $response.pagination.nextCursor
        Write-Host ("`r  {0} retrieved..." -f $rows.Count) -NoNewline -ForegroundColor DarkGray
    } while ($cursor)
    Write-Host ''

    return $rows
}

Write-Host ''
Write-Host "=== SentinelOne passphrase export (READ-ONLY) ===" -ForegroundColor Cyan
Write-Host "  Console : $BaseUri" -ForegroundColor Gray
Write-Host ''

# --- agents (metadata) ---
Write-Host 'Retrieving agents...' -ForegroundColor Cyan
try {
    $agentFilter = 'agents?'
    if (-not $IncludeDecommissioned) { $agentFilter += 'isDecommissioned=false' }
    $agents = @(Get-AllPages -Path $agentFilter)
}
catch {
    $code = $null; try { $code = $_.Exception.Response.StatusCode.value__ } catch {}
    Write-Host "ERROR: Could not list agents from $BaseUri (HTTP $code)." -ForegroundColor Red
    Write-Host "       $($_.Exception.Message)" -ForegroundColor Red
    if ($code -in 401, 403) {
        Write-Host '       The token was rejected. Confirm it is current and account-scoped.' -ForegroundColor Yellow
    }
    elseif (-not $code) {
        Write-Host "       No HTTP response -- check the console URL is correct and reachable." -ForegroundColor Yellow
    }
    exit 2
}

if ($agents.Count -eq 0) {
    Write-Host 'ERROR: No agents returned. Nothing to export.' -ForegroundColor Red
    exit 2
}
Write-Host "  $($agents.Count) agents" -ForegroundColor Gray

# --- passphrases ---
Write-Host 'Retrieving passphrases...' -ForegroundColor Cyan
try {
    $passphrases = @(Get-AllPages -Path 'agents/passphrases?')
}
catch {
    $code = $null; try { $code = $_.Exception.Response.StatusCode.value__ } catch {}
    Write-Host "ERROR: Could not retrieve passphrases (HTTP $code)." -ForegroundColor Red
    if ($code -in 401, 403) {
        Write-Host '       Listing agents worked but passphrases did not, so the token is valid and' -ForegroundColor Yellow
        Write-Host '       simply lacks passphrase rights. In the console, the service user role needs' -ForegroundColor Yellow
        Write-Host '       the agent uninstall-passphrase permission.' -ForegroundColor Yellow
    }
    elseif ($code -eq 404) {
        Write-Host '       This console does not expose /agents/passphrases. Some OEM builds omit it;' -ForegroundColor Yellow
        Write-Host '       the passphrase may only be available per-device in the console UI, or the' -ForegroundColor Yellow
        Write-Host '       product may use a single policy-level uninstall password instead.' -ForegroundColor Yellow
    }
    exit 2
}
Write-Host "  $($passphrases.Count) passphrases" -ForegroundColor Gray

$passphraseById = @{}
foreach ($p in $passphrases) { $passphraseById[[string]$p.id] = $p }

# --- join ---
$rows = foreach ($a in $agents) {
    if ($SiteName -and $a.siteName -notin $SiteName) { continue }

    $p = $passphraseById[[string]$a.id]

    [PSCustomObject]@{
        ComputerName     = $a.computerName
        Passphrase       = $(if ($p) { $p.passphrase } else { '' })
        SiteName         = $a.siteName
        GroupName        = $a.groupName
        Domain           = $a.domain
        OS               = $a.osName
        AgentVersion     = $a.agentVersion
        LastActive       = $a.lastActiveDate
        NetworkStatus    = $a.networkStatus
        IsActive         = $a.isActive
        LastLoggedInUser = $a.lastLoggedInUserName
        AgentId          = $a.id
        UUID             = $a.uuid
    }
}
$rows = @($rows)

if ($rows.Count -eq 0) {
    Write-Host 'ERROR: No agents matched the filter.' -ForegroundColor Red
    if ($SiteName) { Write-Host "       Sites present: $((@($agents | ForEach-Object { $_.siteName } | Sort-Object -Unique)) -join ', ')" -ForegroundColor Yellow }
    exit 2
}

$missing = @($rows | Where-Object { -not $_.Passphrase })

if (-not $OutputPath) {
    $reports = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path $reports)) { New-Item -ItemType Directory -Path $reports -Force | Out-Null }
    $host_ = ([Uri]$BaseUri).Host -replace '\.', '-'
    $OutputPath = Join-Path $reports ("{0}_S1_Passphrases_{1}.csv" -f (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'), $host_)
}

$rows | Sort-Object SiteName, ComputerName | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host ''
Write-Host '--- Summary ---' -ForegroundColor Cyan
Write-Host ("  Devices exported     : {0}" -f $rows.Count) -ForegroundColor Green
Write-Host ("  With a passphrase    : {0}" -f ($rows.Count - $missing.Count)) -ForegroundColor Green
Write-Host ("  MISSING a passphrase : {0}" -f $missing.Count) -ForegroundColor $(if ($missing.Count) { 'Yellow' } else { 'Green' })
Write-Host ("  Sites covered        : {0}" -f (@($rows | ForEach-Object { $_.SiteName } | Sort-Object -Unique)).Count) -ForegroundColor Gray

if ($missing.Count -gt 0) {
    Write-Host ''
    Write-Host '  Devices with no passphrase returned (retrieve these from the console UI):' -ForegroundColor Yellow
    $missing | Select-Object -First 15 | Format-Table ComputerName, SiteName, LastActive -AutoSize | Out-String -Width 4096 | Write-Host
}

Write-Host ''
Write-Host "CSV: $OutputPath" -ForegroundColor Green
Write-Warning 'This file contains uninstall passphrases in plaintext. Anyone with it can remove protection from these endpoints. Delete it once the migration is done.'

Remove-Variable -Name plain -ErrorAction SilentlyContinue
exit 0
