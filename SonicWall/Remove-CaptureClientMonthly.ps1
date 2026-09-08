<#
.SYNOPSIS
    Removes Capture Client subscriptions from MSSP Monthly billing in MySonicWall,
    replicating the per-product "Remove from MSSP Monthly" action in the GUI.

.DESCRIPTION
    The MySonicWall GUI offers "Remove from MSSP Monthly" one product at a time. This
    batches it across every tenant under the MSSP.

    ############################################################################
    #  THIS SCRIPT CHANGES BILLING ON PRODUCTION CUSTOMER SUBSCRIPTIONS.       #
    #  It is a DRY RUN unless -Apply is supplied.                              #
    #                                                                          #
    #  It touches Capture Client (ProductID 401) ONLY. Cloud App Security      #
    #  (402) and every other product are excluded by construction and the      #
    #  exclusion is re-asserted immediately before each call.                  #
    ############################################################################

    Scope is discovered live on every run rather than read from a stale file: a
    subscription removed yesterday must not be attempted again today.

    Safety model:
      - Dry run by default; -Apply performs the removals.
      - Hard ProductID filter, re-checked per row immediately before the call. A row
        that is not the requested product aborts the whole run rather than being
        skipped, because its presence means the filter is broken.
      - Full inventory is snapshotted to CSV before the first change.
      - Rows the API does not offer removal for (removeMsspMonthlyAllowed != YES) are
        reported and skipped.
      - -MinUsedNodeCount / -SkipInUse keep subscriptions with live endpoints out of
        the batch.
      - -TenantId limits the run to specific tenants for a pilot.
      - Every removal is verified by re-reading the inventory afterwards; a call that
        reports success without the subscription actually leaving monthly billing is
        reported as VERIFY_FAILED, not as success.

.NOTES
    Run under Windows PowerShell 5.1 (powershell.exe), not PowerShell 7.

    Credentials come from the store written by Get-CaptureClientTenantMap.ps1
    (Config\CaptureClient-Credentials.xml), which holds the MySonicWall API key.

    ==========================================================================
    THE REMOVAL PAYLOAD IS NOT YET CONFIRMED -- see Invoke-MsspMonthlyRemoval.
    -Apply refuses to run until $script:PayloadConfirmed is set to $true.
    ==========================================================================

.PARAMETER Apply
    Perform the removals. Without it the script reports exactly what it would do.

.PARAMETER ProductId
    MySonicWall product to act on. Default '401' (Capture Client). Changing this is a
    deliberate act; 402 is Cloud App Security.

.PARAMETER TenantId
    Limit the run to these MySonicWall tenant ids (ProductGroupIDs).

.PARAMETER SkipInUse
    Exclude subscriptions whose usedNodeCount is greater than zero, i.e. those with
    endpoints currently consuming licences.

.PARAMETER MinUsedNodeCount
    Exclude subscriptions at or above this consumed-node count. -SkipInUse is the
    same as -MinUsedNodeCount 1.

.PARAMETER OutputDir
    Where the CSVs are written. Defaults to Reports\ beside this script.

.EXAMPLE
    .\Remove-CaptureClientMonthly.ps1

    Dry run. Lists every Capture Client monthly subscription and what would happen.

.EXAMPLE
    .\Remove-CaptureClientMonthly.ps1 -TenantId 2434364 -Apply

    Pilot against a single tenant.

.EXAMPLE
    .\Remove-CaptureClientMonthly.ps1 -SkipInUse -Apply

    Removes only the subscriptions with no endpoints consuming licences.
#>

[CmdletBinding()]
param(
    [switch]$Apply,

    [string]$ProductId = '401',

    [string[]]$TenantId,

    [switch]$SkipInUse,

    [int]$MinUsedNodeCount,

    [string]$OutputDir
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

. (Join-Path $PSScriptRoot 'MySonicWallApi.ps1')

$ConfigPath = Join-Path $PSScriptRoot 'Config\CaptureClient-Credentials.xml'
if (-not $OutputDir) { $OutputDir = Join-Path $PSScriptRoot 'Reports' }

function Write-Step { param([string]$Message, [string]$Color = 'Cyan') Write-Host $Message -ForegroundColor $Color }

#region The removal call

# CONFIRMED against a real "Remove from MSSP Monthly" click captured from the
# MySonicWall GUI (DevTools > Network, 2026-08-25):
#
#     POST https://api.mysonicwall.com/api/mssp/allocatemsspproductservices
#     {"serialNumber":"CC000000A319","Mode":"DEROLL"}
#
# The body really is just those two fields -- no product group, product id, service
# id, node count or expiry. The serial alone identifies what to de-enrol.
#
# The GUI authenticates with an OAuth bearer token, but the endpoint accepts the
# MySonicWall API key just as well: probed with Mode=DEROLL and a bogus serial, it
# clears mode validation and fails on ownership (MSWAPP:PRODUCTGROUPNOTBELONGTOUSER)
# rather than on the mode.
#
# `expireService` and Mode=EDIT, which the OpenAPI schema suggests, are a different
# operation -- editing an existing provision, not removing it. Do not substitute them.
$script:PayloadConfirmed = $true

function Invoke-MsspMonthlyRemoval {
    param(
        [Parameter(Mandatory)][System.Security.SecureString]$ApiKey,
        [Parameter(Mandatory)][object]$Row
    )

    $body = @{
        serialNumber = $Row.SerialNumber
        Mode         = 'DEROLL'
    }

    try {
        $response = Invoke-RestMethod -Uri 'https://api.mysonicwall.com/api/mssp/allocatemsspproductservices' `
            -Headers (New-MswHeader -ApiKey $ApiKey) -Method Post `
            -Body (ConvertTo-Json $body) -ContentType 'application/json'
    }
    catch {
        return [pscustomobject]@{ Ok = $false; Detail = $_.Exception.Message }
    }

    if ($response.operationStatus -eq 'OK') { return [pscustomobject]@{ Ok = $true; Detail = '' } }
    return [pscustomobject]@{ Ok = $false; Detail = "$($response.msgCode) $($response.msgDesc)".Trim() }
}

#endregion

#region Configuration

if (-not (Test-Path $ConfigPath)) {
    Write-Host "ERROR: No credential store at $ConfigPath" -ForegroundColor Red
    Write-Host '       Run Get-CaptureClientTenantMap.ps1 once first to create it.' -ForegroundColor Red
    exit 2
}
$config = Import-Clixml -Path $ConfigPath
if (-not $config.MySonicWall -or -not $config.MySonicWall.ApiKey) {
    Write-Host 'ERROR: The credential store holds no MySonicWall API key.' -ForegroundColor Red
    exit 2
}
$apiKey = $config.MySonicWall.ApiKey

if ($SkipInUse -and -not $PSBoundParameters.ContainsKey('MinUsedNodeCount')) { $MinUsedNodeCount = 1 }

#endregion

Write-Host ''
if ($Apply) { Write-Step '=== Remove from MSSP Monthly -- APPLY (changes billing) ===' 'Yellow' }
else { Write-Step '=== Remove from MSSP Monthly -- DRY RUN (no changes) ===' 'Cyan' }
Write-Host ("  Product   : {0}{1}" -f $ProductId, $(if ($ProductId -eq '401') { ' (Capture Client)' } elseif ($ProductId -eq '402') { ' (Cloud App Security)' } else { '' })) -ForegroundColor Gray
if ($TenantId) { Write-Host ("  Tenants   : {0}" -f ($TenantId -join ', ')) -ForegroundColor Gray }
if ($MinUsedNodeCount -gt 0) { Write-Host ("  Skipping  : subscriptions with {0}+ consumed nodes" -f $MinUsedNodeCount) -ForegroundColor Gray }
Write-Host ''

#region Inventory

try {
    Write-Step 'Reading MSSP monthly inventory...'
    $content = Get-MswMsspProductService -ApiKey $apiKey -TenantId $TenantId
    $all = @(Get-MswMonthlyProductRow -Content $content -ProductId $ProductId)
}
catch {
    Write-Host "ERROR: Could not read the MSSP inventory: $($_.Exception.Message)" -ForegroundColor Red
    exit 2
}

Write-Host ("  Tenants under the MSSP        : {0}" -f @($content.mssplist).Count) -ForegroundColor Gray
Write-Host ("  Monthly subscriptions (id {0}) : {1}" -f $ProductId, $all.Count) -ForegroundColor Gray

if ($all.Count -eq 0) {
    Write-Host ''
    Write-Step 'Nothing on MSSP monthly billing for this product. Nothing to do.' 'Green'
    exit 0
}

if (-not (Test-Path $OutputDir)) { $null = New-Item -ItemType Directory -Path $OutputDir -Force }
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$inventoryPath = Join-Path $OutputDir "${stamp}_MSW_MonthlyInventory_$ProductId.csv"
$all | Sort-Object TenantName | Export-Csv -Path $inventoryPath -NoTypeInformation -Encoding UTF8
Write-Host ("  Inventory snapshot            : {0}" -f $inventoryPath) -ForegroundColor Gray

#endregion

#region Select

$notRemovable = @($all | Where-Object { $_.RemoveAllowed -ne 'YES' })
$inUse = @($all | Where-Object { $MinUsedNodeCount -gt 0 -and $_.UsedNodeCount -ge $MinUsedNodeCount })
$targets = @($all | Where-Object {
    $_.RemoveAllowed -eq 'YES' -and -not ($MinUsedNodeCount -gt 0 -and $_.UsedNodeCount -ge $MinUsedNodeCount)
})

Write-Host ''
Write-Step '--- Selection ---'
Write-Host ("  Eligible for removal : {0}" -f $targets.Count) -ForegroundColor Green
if ($inUse.Count) { Write-Host ("  Skipped, in use      : {0}" -f $inUse.Count) -ForegroundColor Yellow }
if ($notRemovable.Count) { Write-Host ("  Removal not offered  : {0}" -f $notRemovable.Count) -ForegroundColor Yellow }
Write-Host ("  Licensed nodes total : {0}" -f (($targets | Measure-Object NodeCount -Sum).Sum)) -ForegroundColor Gray

if ($inUse.Count -gt 0) {
    Write-Host ''
    Write-Step '--- Skipped: endpoints are consuming licences ---' 'Yellow'
    $inUse | Sort-Object -Descending UsedNodeCount |
        Format-Table -AutoSize TenantName, TenantId, NodeCount, UsedNodeCount | Out-String -Width 200 | Write-Host
}

if ($notRemovable.Count -gt 0) {
    Write-Host ''
    Write-Step '--- Skipped: removeMsspMonthlyAllowed is not YES ---' 'Yellow'
    $notRemovable | Sort-Object TenantName |
        Format-Table -AutoSize TenantName, TenantId, SerialNumber, MsspStatus, RemoveAllowed | Out-String -Width 200 | Write-Host
}

if ($targets.Count -eq 0) {
    Write-Host ''
    Write-Step 'Nothing eligible after filtering. Nothing to do.' 'Green'
    exit 0
}

#endregion

#region Dry run

if (-not $Apply) {
    Write-Host ''
    Write-Step '--- Would remove from MSSP Monthly ---'
    $targets | Sort-Object TenantName |
        Format-Table -AutoSize TenantName, TenantId, SerialNumber, ServiceName, NodeCount, UsedNodeCount |
        Out-String -Width 200 | Write-Host

    Write-Host ''
    Write-Step ("DRY RUN complete. {0} subscription(s) would be removed. Nothing was changed." -f $targets.Count) 'Cyan'
    if (-not $script:PayloadConfirmed) {
        Write-Host ''
        Write-Host '  NOTE: the removal payload is still unconfirmed, so -Apply will refuse to run.' -ForegroundColor Yellow
        Write-Host '        See the "removal call" region at the top of this script.' -ForegroundColor Yellow
    }
    Write-Host ''
    exit 0
}

#endregion

#region Apply

if (-not $script:PayloadConfirmed) {
    Write-Host ''
    Write-Host 'ERROR: The "Remove from MSSP Monthly" payload has not been confirmed.' -ForegroundColor Red
    Write-Host '       Refusing to change billing on live customer subscriptions using a guess.' -ForegroundColor Red
    Write-Host ''
    Write-Host '       To confirm it: remove one product in the MySonicWall GUI with browser' -ForegroundColor Yellow
    Write-Host '       DevTools open on the Network tab, then compare the captured request with' -ForegroundColor Yellow
    Write-Host '       Invoke-MsspMonthlyRemoval in this script and set $script:PayloadConfirmed.' -ForegroundColor Yellow
    exit 3
}

Write-Host ''
Write-Step ("Removing {0} subscription(s) from MSSP Monthly..." -f $targets.Count) 'Yellow'

$results = @()
$done = 0
foreach ($row in $targets) {
    $done++
    Write-Progress -Activity 'Remove from MSSP Monthly' -Status $row.TenantName `
        -PercentComplete ([int](100 * $done / $targets.Count))

    # Re-assert the product filter immediately before the call. If a row of another
    # product reached this point the filter is broken, and continuing could remove
    # Cloud App Security subscriptions -- so stop the whole run.
    if ($row.ProductId -ne $ProductId) {
        Write-Progress -Activity 'Remove from MSSP Monthly' -Completed
        Write-Host ''
        Write-Host ("FATAL: row for {0} carries ProductID {1}, expected {2}. Aborting." -f $row.TenantName, $row.ProductId, $ProductId) -ForegroundColor Red
        Write-Host '       No further calls will be made. Results so far are logged below.' -ForegroundColor Red
        break
    }

    $call = Invoke-MsspMonthlyRemoval -ApiKey $apiKey -Row $row
    $results += [pscustomobject]@{
        TenantName   = $row.TenantName
        TenantId     = $row.TenantId
        SerialNumber = $row.SerialNumber
        ProductId    = $row.ProductId
        ServiceName  = $row.ServiceName
        NodeCount    = $row.NodeCount
        Result       = if ($call.Ok) { 'CALLED' } else { 'FAILED' }
        Detail       = $call.Detail
    }

    Write-Host ("  {0,-42} {1}" -f $row.TenantName, $(if ($call.Ok) { 'CALLED' } else { "FAILED - $($call.Detail)" })) `
        -ForegroundColor $(if ($call.Ok) { 'Green' } else { 'Red' })
}
Write-Progress -Activity 'Remove from MSSP Monthly' -Completed

# Verify against a fresh read rather than trusting the response codes.
Write-Host ''
Write-Step 'Verifying against a fresh inventory read...'
try {
    $after = Get-MswMsspProductService -ApiKey $apiKey -TenantId $TenantId

    # DEROLL takes the whole serial off monthly billing, so the serial disappearing
    # from the product's monthly inventory is the thing to assert.
    $stillMonthly = @{}
    foreach ($row in @(Get-MswMonthlyProductRow -Content $after -ProductId $ProductId)) {
        $stillMonthly[$row.SerialNumber] = $true
    }
    foreach ($result in $results) {
        if ($result.Result -ne 'CALLED') { continue }
        $result.Result = if ($stillMonthly.ContainsKey($result.SerialNumber)) { 'VERIFY_FAILED' } else { 'REMOVED' }
    }
}
catch {
    Write-Warning "Could not re-read the inventory to verify: $($_.Exception.Message)"
}

$resultPath = Join-Path $OutputDir "${stamp}_MSW_MonthlyRemoval_Results.csv"
$results | Export-Csv -Path $resultPath -NoTypeInformation -Encoding UTF8

$removed = @($results | Where-Object Result -eq 'REMOVED')
$failed = @($results | Where-Object Result -ne 'REMOVED')

Write-Host ''
Write-Step '--- Result ---'
Write-Host ("  Removed : {0}" -f $removed.Count) -ForegroundColor Green
if ($failed.Count) { Write-Host ("  Failed  : {0}" -f $failed.Count) -ForegroundColor Red }
Write-Host ("  Before  : {0}" -f $inventoryPath) -ForegroundColor Gray
Write-Host ("  Log     : {0}" -f $resultPath) -ForegroundColor Gray
Write-Host ''

exit $(if ($failed.Count) { 1 } else { 0 })

#endregion
