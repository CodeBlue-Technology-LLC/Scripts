<#
.SYNOPSIS
    READ-ONLY: compares billed SentinelOne quantities in ConnectWise Manage against
    the actual endpoint counts in the SentinelOne EDR portal.

.DESCRIPTION
    Answers "are we billing what we are actually running?" for SentinelOne.

    Billed side  : every currently-billing agreement addition whose product identifier
                   is -ProductIdentifier (default SENT-ONE-CTRL), on any ACTIVE
                   agreement. "Currently billing" means the addition's effectiveDate
                   is on or before today AND its cancelledDate is empty or in the
                   future -- cancelled lines stay on the agreement record with a
                   cancelledDate, so they must be filtered out explicitly rather than
                   assumed absent.

    Actual side  : activeLicenses per active site in the SentinelOne EDR console --
                   agents actually consuming a licence.

                   The site's totalLicenses is deliberately NOT reported. It is a soft
                   per-site allotment an admin types in, not a cost basis: sites are
                   routinely running MORE agents than their allotment (several sit at
                   totalLicenses 0 while running dozens), so it is not even a ceiling.
                   The account itself reports totalLicenses -1 (uncapped) with
                   activeAgents equal to the sum of site activeLicenses, i.e. nothing
                   is pre-purchased and consumption is the only figure that exists.

    ############################################################################
    #  THIS SCRIPT IS STRICTLY READ-ONLY.                                      #
    #  It performs GET operations against ConnectWise Manage and SentinelOne.   #
    #  It does NOT change any billing quantity. Use CWM-SyncSentinelOneCounts   #
    #  .ps1 for that.                                                           #
    ############################################################################

    A company may carry more than one matching addition (for example on both an IT
    Services and a Data Center agreement). Quantities are SUMMED per company, so the
    billed figure is what the customer actually pays for in total, not one line of it.

    Both sides are matched on a normalized company name. Companies present on only one
    side are still reported, with zero in the missing column -- those are usually the
    most interesting rows: a billed line with no deployment, or a live site nobody is
    invoicing.

    Output: one CSV with four columns -- Company Name, Billed Quantity, Active
    Licenses, Status -- plus a console summary of the discrepancies.

    Status explains a zero, which is the point of it:
      Billing               currently invoiced; the two quantities are comparable
      Starts <date> (qty N) addition exists but its effective date is in the future,
                            so a live site legitimately shows 0 billed today
      <status> agreement    addition is current but its agreement is Expired,
        (qty N)             Inactive or Cancelled -- nothing is being invoiced
      Cancelled             the only addition carries a past cancelledDate
      Not billed            agents deployed with no matching addition at all
      No S1 site            billed with no site in the console

.NOTES
    Run under Windows PowerShell 5.1 (powershell.exe), not PowerShell 7.

    Prerequisites:
      - ConnectWiseManageAPI module (Install-Module ConnectWiseManageAPI)
      - A Credentials.xml holding a CWMConnectionInfo block (the Veeam / Cove / RDS
        sync scripts all create one). Auto-discovered when -ConfigPath is omitted.
      - SentinelOne EDR base URI and API token, taken from
        Config\S1-AutomateReconcile-Credentials.xml.

.PARAMETER ProductIdentifier
    CWM product identifier to treat as the billed SentinelOne line.
    Default 'SENT-ONE-CTRL'. The MDR product is 'SENT-ONE-MDR'.

.PARAMETER ConfigPath
    Credentials.xml holding a CWMConnectionInfo block. Auto-discovered when omitted.

.PARAMETER S1ConfigPath
    Credential store holding the SentinelOne base URI and API token.
    Defaults to Config\S1-AutomateReconcile-Credentials.xml beside this script.

.PARAMETER S1BaseUri
    Override the SentinelOne console URI, e.g. the MDR instance.

.PARAMETER IncludeZeroRows
    Keep rows where both the billed and the active quantity are zero. Off by default:
    they carry no information and only pad the file.

.PARAMETER OutputPath
    Directory for the CSV. Defaults to Reports\ beside this script.

.EXAMPLE
    .\Compare-S1BilledVsActual.ps1

    Compares SENT-ONE-CTRL billing against the EDR console and writes the CSV.

.EXAMPLE
    .\Compare-S1BilledVsActual.ps1 -ProductIdentifier SENT-ONE-MDR -S1BaseUri https://usea1-cw04mdr.sentinelone.net

    Same comparison for the MDR product and console.
#>

[CmdletBinding()]
param(
    [string]$ProductIdentifier = 'SENT-ONE-CTRL',

    [string]$ConfigPath,

    [string]$S1ConfigPath,

    [string]$S1BaseUri,

    [switch]$IncludeZeroRows,

    [ValidateSet('All', 'ActiveOnly')]
    [string]$AgreementScope = 'All',

    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if (-not $OutputPath) { $OutputPath = Join-Path $PSScriptRoot 'Reports' }
if (-not $S1ConfigPath) { $S1ConfigPath = Join-Path $PSScriptRoot 'Config\S1-AutomateReconcile-Credentials.xml' }

function Write-Step { param([string]$Message, [string]$Color = 'Cyan') Write-Host $Message -ForegroundColor $Color }

function Get-NormalizedCompanyName {
    <#
        Company-name normalization, applied identically to both sides of the join.
        Conservative on purpose: aggressive stripping collapses genuinely different
        companies onto one string, which silently merges two customers' numbers.

          - lowercase
          - drop apostrophes outright, so O'Neal matches ONeal
          - remaining punctuation becomes a separator
          - merge runs of two or more single-letter tokens (P.C. -> pc)
          - drop standalone legal-suffix tokens, word-boundary matched
          - drop a leading article
    #>
    param(
        [string]$Name,
        [string[]]$Suffix = @('inc', 'llc', 'ltd', 'corp', 'co', 'pc', 'pllc', 'plc', 'llp', 'lp'),
        [string[]]$Article = @('the')
    )

    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }

    $text = $Name.ToLowerInvariant()
    $text = $text -replace '[''\u2018\u2019]', ''
    $text = $text -replace '[^\w\s]', ' '
    $text = ($text -replace '\s+', ' ').Trim()

    $tokens = @($text -split ' ' | Where-Object { $_ })
    $merged = New-Object System.Collections.Generic.List[string]
    $run = New-Object System.Collections.Generic.List[string]
    $flushRun = {
        if ($run.Count -ge 2) { $merged.Add(($run -join '')) }
        elseif ($run.Count -eq 1) { $merged.Add($run[0]) }
        $run.Clear()
    }
    foreach ($token in $tokens) {
        if ($token.Length -eq 1 -and $token -match '^[a-z]$') { $run.Add($token) }
        else { & $flushRun; $merged.Add($token) }
    }
    & $flushRun
    $text = $merged -join ' '

    $pattern = ($Suffix | ForEach-Object { [regex]::Escape($_) }) -join '|'
    if ($pattern) { $text = $text -replace "\b(?:$pattern)\b", ' ' }
    $text = ($text -replace '\s+', ' ').Trim()

    $pattern = ($Article | ForEach-Object { [regex]::Escape($_) }) -join '|'
    if ($pattern) { $text = $text -replace "^(?:$pattern)\b", ' ' }

    return ($text -replace '\s+', ' ').Trim()
}

#region Configuration

if (-not $ConfigPath) {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $found = Get-ChildItem -Path $repoRoot -Recurse -Filter '*redential*.xml' -ErrorAction SilentlyContinue |
        Where-Object { try { (Import-Clixml -Path $_.FullName).CWMConnectionInfo } catch { $false } } |
        Select-Object -First 1
    if (-not $found) {
        Write-Host 'ERROR: No Credentials.xml with a CWMConnectionInfo block found under the repo.' -ForegroundColor Red
        Write-Host '       Pass -ConfigPath explicitly.' -ForegroundColor Red
        exit 2
    }
    $ConfigPath = $found.FullName
}

if (-not (Test-Path $S1ConfigPath)) {
    Write-Host "ERROR: No SentinelOne credential store at $S1ConfigPath" -ForegroundColor Red
    exit 2
}

$cwmConfig = Import-Clixml -Path $ConfigPath
$s1Config = Import-Clixml -Path $S1ConfigPath

if (-not $S1BaseUri) { $S1BaseUri = [string]$s1Config.S1.BaseUri }
$S1BaseUri = $S1BaseUri.TrimEnd('/')
if (-not $S1BaseUri) {
    Write-Host 'ERROR: No SentinelOne base URI available. Pass -S1BaseUri.' -ForegroundColor Red
    exit 2
}

$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($s1Config.S1.ApiToken)
try { $s1Token = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }

#endregion

Write-Host ''
Write-Step '=== SentinelOne billed vs actual -- READ ONLY ==='
Write-Host ("  Product   : {0}" -f $ProductIdentifier) -ForegroundColor Gray
Write-Host ("  Console   : {0}" -f $S1BaseUri) -ForegroundColor Gray
Write-Host ("  CWM config: {0}" -f (Split-Path -Leaf $ConfigPath)) -ForegroundColor Gray
Write-Host ''

#region SentinelOne

try {
    Write-Step 'Reading SentinelOne sites...'
    $headers = @{ Authorization = "ApiToken $s1Token" }
    $sites = @()
    $cursor = $null
    do {
        $uri = "$S1BaseUri/web/api/v2.1/sites?states=active&limit=1000"
        if ($cursor) { $uri += "&cursor=$cursor" }
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
        $sites += @($response.data.sites)
        $cursor = $response.pagination.nextCursor
    } while ($cursor)
}
catch {
    Write-Host "ERROR: SentinelOne query failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 2
}

# "Default site" is the console's catch-all, not a customer.
$sites = @($sites | Where-Object { $_.name -ne 'Default site' })
Write-Host ("  Active sites: {0}" -f $sites.Count) -ForegroundColor Gray

#endregion

#region ConnectWise Manage

try {
    if (-not (Get-Module -ListAvailable -Name ConnectWiseManageAPI)) {
        throw 'The ConnectWiseManageAPI module is not installed. Run: Install-Module ConnectWiseManageAPI -Scope CurrentUser'
    }
    Import-Module ConnectWiseManageAPI -ErrorAction Stop

    Write-Step 'Connecting to ConnectWise Manage...'
    $connectionInfo = $cwmConfig.CWMConnectionInfo
    Connect-CWM @connectionInfo | Out-Null

    # Scanning every status, not just Active, is what lets a line be reported as
    # "starts next month" or "stranded on an expired agreement" instead of silently
    # vanishing. It costs one additions call per agreement, so -AgreementScope
    # ActiveOnly is there for a fast re-run once the exceptions are known.
    $agreements = if ($AgreementScope -eq 'ActiveOnly') {
        @(Get-CWMAgreement -condition 'agreementStatus="Active"' -all)
    } else {
        @(Get-CWMAgreement -all)
    }
}
catch {
    Write-Host "ERROR: ConnectWise Manage connection failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 2
}
Write-Host ("  Agreements to scan: {0} ({1})" -f $agreements.Count, $AgreementScope) -ForegroundColor Gray

Write-Step "Reading agreement additions for $ProductIdentifier..."
# CWM returns UTC timestamps with a Z suffix. Casting one straight to [datetime]
# converts it to local time, which rolls a midnight-UTC date back to the previous day
# in any western timezone -- an effective date of 2026-09-01 would read as 08-31 and
# a line could be judged effective a day early. Compare and display in UTC throughout.
$now = (Get-Date).ToUniversalTime()
$billedLines = @()
$cancelledCount = 0
$index = 0

foreach ($agreement in $agreements) {
    $index++
    Write-Progress -Activity 'Reading agreement additions' `
        -Status "$index / $($agreements.Count) - $($agreement.company.name)" `
        -PercentComplete ($index / [Math]::Max($agreements.Count, 1) * 100)

    try { $additions = @(Get-CWMAgreementAddition -parentId $agreement.id -all) }
    catch {
        Write-Warning "Could not read additions on agreement $($agreement.id) ($($agreement.company.name)): $($_.Exception.Message)"
        continue
    }

    foreach ($addition in $additions) {
        $identifier = if ($addition.product) { [string]$addition.product.identifier } else { '' }
        if ($identifier -ne $ProductIdentifier) { continue }

        # Cancelled lines remain on the record with a cancelledDate; a line not yet
        # effective is equally not being billed today. Both are kept and labelled
        # rather than dropped, because "starts next month" is the single most common
        # reason a live site looks unbilled.
        $effective = if ($addition.effectiveDate) { ([datetimeoffset]$addition.effectiveDate).UtcDateTime } else { $null }
        $cancelled = if ($addition.cancelledDate) { ([datetimeoffset]$addition.cancelledDate).UtcDateTime } else { $null }
        $isEffective = (-not $effective) -or ($effective -le $now)
        $isLive = (-not $cancelled) -or ($cancelled -gt $now)
        $agreementLive = ([string]$agreement.agreementStatus -eq 'Active')

        $state =
            if (-not $isLive) { 'Cancelled' }
            elseif (-not $isEffective) { 'Future' }
            elseif (-not $agreementLive) { 'StrandedAgreement' }
            else { 'Billing' }

        if ($state -ne 'Billing') { $cancelledCount++ }

        $billedLines += [pscustomobject]@{
            Company         = [string]$agreement.company.name
            Quantity        = [double]$addition.quantity
            State           = $state
            EffectiveDate   = $effective
            AgreementStatus = [string]$agreement.agreementStatus
        }
    }
}
Write-Progress -Activity 'Reading agreement additions' -Completed

$byState = $billedLines | Group-Object State
Write-Host ("  {0} additions found : {1}" -f $ProductIdentifier, $billedLines.Count) -ForegroundColor Gray
foreach ($group in ($byState | Sort-Object Name)) {
    Write-Host ("    {0,-18} {1}" -f $group.Name, $group.Count) -ForegroundColor Gray
}

#endregion

#region Join

# Sum per company: a customer can hold the same addition on more than one agreement.
# Only Billing lines count toward the billed quantity -- a future or cancelled line is
# not money changing hands today -- but the other states are kept alongside so the row
# can explain itself.
$billedByName = @{}
foreach ($line in $billedLines) {
    $key = Get-NormalizedCompanyName -Name $line.Company
    if (-not $key) { continue }
    if (-not $billedByName.ContainsKey($key)) {
        $billedByName[$key] = [pscustomobject]@{
            Name       = $line.Company
            Quantity   = 0.0
            Future     = @()
            Stranded   = @()
            Cancelled  = @()
        }
    }
    $entry = $billedByName[$key]
    switch ($line.State) {
        'Billing'           { $entry.Quantity += $line.Quantity }
        'Future'            { $entry.Future += $line }
        'StrandedAgreement' { $entry.Stranded += $line }
        'Cancelled'         { $entry.Cancelled += $line }
    }
}

$actualByName = @{}
foreach ($site in $sites) {
    $key = Get-NormalizedCompanyName -Name $site.name
    if (-not $key) { continue }
    if (-not $actualByName.ContainsKey($key)) {
        $actualByName[$key] = [pscustomobject]@{ Name = [string]$site.name; Active = 0.0 }
    }
    $actualByName[$key].Active += [double]$site.activeLicenses
}

$keys = @($billedByName.Keys) + @($actualByName.Keys) | Sort-Object -Unique

$rows = @()
foreach ($key in $keys) {
    $billed = $billedByName[$key]
    $actual = $actualByName[$key]

    $billedQty = if ($billed) { [int]$billed.Quantity } else { 0 }
    $activeQty = if ($actual) { [int]$actual.Active } else { 0 }

    # Status explains a zero, which is the whole point: a live site with nothing
    # billed is an anomaly, unless the line simply has not started yet.
    $status =
        if ($billedQty -gt 0) {
            'Billing'
        }
        elseif ($billed -and $billed.Future.Count -gt 0) {
            $next = $billed.Future | Sort-Object EffectiveDate | Select-Object -First 1
            $pending = ($billed.Future | Measure-Object Quantity -Sum).Sum
            'Starts {0:yyyy-MM-dd} (qty {1})' -f $next.EffectiveDate, [int]$pending
        }
        elseif ($billed -and $billed.Stranded.Count -gt 0) {
            $line = $billed.Stranded | Select-Object -First 1
            '{0} agreement (qty {1})' -f $line.AgreementStatus, [int](($billed.Stranded | Measure-Object Quantity -Sum).Sum)
        }
        elseif ($billed -and $billed.Cancelled.Count -gt 0) {
            'Cancelled'
        }
        elseif ($activeQty -gt 0) {
            'Not billed'
        }
        else {
            'No billing, no agents'
        }

    if ($billedQty -gt 0 -and -not $actual) { $status = 'No S1 site' }

    # Drop only rows that carry no signal at all: nothing billed, nothing deployed,
    # and no pending or stranded line to explain why.
    if (-not $IncludeZeroRows -and $billedQty -eq 0 -and $activeQty -eq 0 -and $status -eq 'No billing, no agents') { continue }

    # Prefer the CWM company name: it is the billing name of record.
    $rows += [pscustomobject]@{
        'Company Name'    = if ($billed) { $billed.Name } else { $actual.Name }
        'Billed Quantity' = $billedQty
        'Active Licenses' = $activeQty
        'Status'          = $status
        OnBothSides       = [bool]($billed -and $actual)
    }
}

#endregion

#region Report

# Buckets are keyed off Status, not off which side the row came from: a future-dated
# line puts a company on both sides while billing nothing, and comparing 0 against a
# live agent count there would manufacture a discrepancy that does not exist.
# Status alone decides comparability -- do NOT also require agents > 0 here. A company
# billed for seats with zero agents deployed is exactly the overbilling this report
# exists to surface, and filtering on agent count would hide it.
$matched = @($rows | Where-Object { $_.Status -eq 'Billing' })
$billedOnly = @($rows | Where-Object { $_.Status -eq 'No S1 site' })
$actualOnly = @($rows | Where-Object { $_.Status -eq 'Not billed' })
$future = @($rows | Where-Object { $_.Status -like 'Starts *' })
$stranded = @($rows | Where-Object { $_.Status -like '*agreement (qty*' })
$cancelledRows = @($rows | Where-Object { $_.Status -eq 'Cancelled' })
$overBilled = @($matched | Where-Object { $_.'Billed Quantity' -gt $_.'Active Licenses' })
$underBilled = @($matched | Where-Object { $_.'Billed Quantity' -lt $_.'Active Licenses' })

Write-Host ''
Write-Step '--- Summary ---'
Write-Host ("  Companies in the report : {0}" -f $rows.Count)
Write-Host ("  On both sides           : {0}" -f $matched.Count) -ForegroundColor Green
Write-Host ("    billed > actual       : {0}" -f $overBilled.Count) -ForegroundColor $(if ($overBilled.Count) { 'Yellow' } else { 'Gray' })
Write-Host ("    billed < actual       : {0}" -f $underBilled.Count) -ForegroundColor $(if ($underBilled.Count) { 'Yellow' } else { 'Gray' })
Write-Host ("    exact match           : {0}" -f ($matched.Count - $overBilled.Count - $underBilled.Count)) -ForegroundColor Gray
Write-Host ("  Billed, no S1 site      : {0}" -f $billedOnly.Count) -ForegroundColor $(if ($billedOnly.Count) { 'Red' } else { 'Gray' })
Write-Host ("  S1 site, not billed     : {0}" -f $actualOnly.Count) -ForegroundColor $(if ($actualOnly.Count) { 'Red' } else { 'Gray' })
Write-Host ("  Scheduled to start      : {0}" -f $future.Count) -ForegroundColor $(if ($future.Count) { 'Cyan' } else { 'Gray' })
Write-Host ("  On a dead agreement     : {0}" -f $stranded.Count) -ForegroundColor $(if ($stranded.Count) { 'Yellow' } else { 'Gray' })
Write-Host ''
Write-Host ("  Total billed  : {0}" -f (($rows | Measure-Object 'Billed Quantity' -Sum).Sum)) -ForegroundColor Gray
Write-Host ("  Total active  : {0}" -f (($rows | Measure-Object 'Active Licenses' -Sum).Sum)) -ForegroundColor Gray

if ($billedOnly.Count -gt 0) {
    Write-Host ''
    Write-Step '--- Billed, but no site in the SentinelOne console ---' 'Red'
    $billedOnly | Sort-Object 'Company Name' |
        Format-Table -AutoSize 'Company Name', 'Billed Quantity' | Out-String -Width 120 | Write-Host
}

if ($actualOnly.Count -gt 0) {
    Write-Host ''
    Write-Step '--- Deployed in SentinelOne, but nothing billed ---' 'Red'
    $actualOnly | Sort-Object -Descending { $_.'Active Licenses' } |
        Format-Table -AutoSize 'Company Name', 'Active Licenses' | Out-String -Width 120 | Write-Host
}

if ($future.Count -gt 0) {
    Write-Host ''
    Write-Step '--- Not billing yet, but scheduled (no action needed) ---' 'Cyan'
    $future | Sort-Object 'Company Name' |
        Format-Table -AutoSize 'Company Name', 'Active Licenses', 'Status' | Out-String -Width 120 | Write-Host
}

if ($stranded.Count -gt 0) {
    Write-Host ''
    Write-Step '--- Addition sits on an agreement that is no longer active ---' 'Yellow'
    $stranded | Sort-Object 'Company Name' |
        Format-Table -AutoSize 'Company Name', 'Active Licenses', 'Status' | Out-String -Width 120 | Write-Host
}

if ($overBilled.Count -gt 0 -or $underBilled.Count -gt 0) {
    Write-Host ''
    Write-Step '--- Quantity mismatches ---' 'Yellow'
    @($overBilled + $underBilled) |
        Sort-Object -Descending { [Math]::Abs($_.'Billed Quantity' - $_.'Active Licenses') } |
        Format-Table -AutoSize 'Company Name', 'Billed Quantity', 'Active Licenses',
            @{ n = 'Delta'; e = { $_.'Billed Quantity' - $_.'Active Licenses' } } |
        Out-String -Width 120 | Write-Host
}

#endregion

#region Output

if (-not (Test-Path $OutputPath)) { $null = New-Item -ItemType Directory -Path $OutputPath -Force }
$stamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$csvPath = Join-Path $OutputPath "${stamp}_S1_BilledVsActual_$ProductIdentifier.csv"

$rows | Sort-Object 'Company Name' |
    Select-Object 'Company Name', 'Billed Quantity', 'Active Licenses', 'Status' |
    Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host ''
Write-Step '--- File ---'
Write-Host ("  {0}" -f $csvPath) -ForegroundColor Gray
Write-Host ''

#endregion
