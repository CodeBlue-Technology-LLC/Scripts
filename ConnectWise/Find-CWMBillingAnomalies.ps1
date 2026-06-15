<#
.SYNOPSIS
    Audits ConnectWise Manage agreement additions for billing inconsistencies.

.DESCRIPTION
    READ-ONLY. Connects to ConnectWise Manage with stored credentials, pulls every
    addition on active agreements of the chosen types, and reports likely billing
    anomalies:

      - Superseding-part duplication ... a NEW part number that replaced two (or
        more) OLD parts billed alongside the old parts it replaced (e.g.
        automate-patch billed together with automate-patch-workstation +
        automate-patch-server) -- the client is charged twice for the same thing.
      - Duplicate line items ........... the same product on one agreement twice.
      - Coverage gaps .................. within one named book of agreements
        (default: PeopleFirst Support), a product most of those clients have but
        some are missing (possible under-billing).
      - Billable but free ............. billable additions with a $0 unit price.
      - Zero quantity ................. active billable additions with quantity 0.
      - Negative margin ............... unit cost greater than unit price.
      - Price outliers ................ a client billed well off the typical PER-UNIT
        price for a product (often stale pricing).

    It also produces a separate month-over-month report: each IT Services agreement's
    recurring total reconstructed as of the 1st of last month vs this month, flagging
    any that moved by at least $100 or 5% (configurable).

    This script NEVER writes to ConnectWise Manage. It only reads and reports.

.NOTES
    Credentials: reuses an existing Config\Credentials.xml that contains the
    CWMConnectionInfo block (the same file the Veeam / SentinelOne / Cove sync
    scripts create). Point -ConfigPath at one of those files, or let the script
    auto-discover a sibling Config\Credentials.xml under the repo.

    Prerequisites: ConnectWiseManageAPI module (installed automatically if missing).

.EXAMPLE
    .\Find-CWMBillingAnomalies.ps1
    .\Find-CWMBillingAnomalies.ps1 -ConfigPath ..\Veeam\Config\Credentials.xml
    .\Find-CWMBillingAnomalies.ps1 -AgreementTypes 'IT Services Agreement','Data Center Agreement'
#>

[CmdletBinding()]
param(
    # Path to an existing Credentials.xml holding a CWMConnectionInfo block.
    [string]$ConfigPath,

    # Agreement types to audit.
    [string[]]$AgreementTypes = @('IT Services Agreement'),

    # Where CSV reports are written (defaults to .\Reports next to this script).
    [string]$OutputPath,

    # Coverage gap: a product adopted by at least this fraction of companies (but
    # not all) is flagged for the companies that are missing it. 0.5 = 50%.
    [double]$CoverageThreshold = 0.5,

    # Coverage gap is scoped to agreements whose name contains this text, and each
    # such company is compared only against other agreements that also match. This
    # keeps the comparison apples-to-apples within one standardized book of business.
    [string]$CoverageAgreementNameLike = 'IT Services Agreement - PeopleFirst Support',

    # Price outlier: flag a line whose per-unit price is at least this multiple
    # away from the product's median, in either direction. 1.0 = 100%, i.e. flag
    # only prices at/above 2x the median (above) or at/below half the median (below).
    [double]$PriceOutlierTolerance = 1.0,

    # Minimum number of companies billed for a product before price-outlier and
    # coverage analysis run for it (avoids noise on rare/custom SKUs).
    [int]$MinSamplesForPricing = 4,

    # Month-over-month agreement-total check: flag an agreement whose recurring total
    # changed by at least this many dollars OR this fraction between last month and
    # this month. Defaults: $100 or 5%.
    [double]$MoMDollarThreshold = 100,
    [double]$MoMPercentThreshold = 0.05,

    # Only this agreement type is included in the month-over-month check.
    [string]$MoMAgreementType = 'IT Services Agreement'
)

$ErrorActionPreference = 'Stop'

# $PSScriptRoot is not populated in param defaults when [CmdletBinding()] is set
# (Windows PowerShell quirk), so default OutputPath here in the body instead.
if (-not $OutputPath) { $OutputPath = Join-Path $PSScriptRoot 'Reports' }

#region ----- Run under Windows PowerShell 5.1 -----------------------------------
# The ConnectWiseManageAPI module calls Invoke-RestMethod internally. On this host's
# PowerShell 7 build, Invoke-RestMethod throws a NullReferenceException on every call
# (a pwsh bug, not a credentials problem), so the module cannot connect. Invoke-
# RestMethod works correctly under Windows PowerShell 5.1, so re-launch there.
if ($PSVersionTable.PSEdition -eq 'Core' -and $IsWindows) {
    $winPS = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path $winPS) {
        Write-Host "Relaunching under Windows PowerShell 5.1 (ConnectWiseManageAPI is unreliable under PowerShell 7 on this host)..." -ForegroundColor Yellow
        $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
        foreach ($kv in $PSBoundParameters.GetEnumerator()) {
            if ($kv.Value -is [System.Management.Automation.SwitchParameter]) {
                if ($kv.Value.IsPresent) { $argList += "-$($kv.Key)" }
            }
            elseif ($kv.Value -is [Array]) {
                $argList += "-$($kv.Key)"; $argList += ($kv.Value -join ',')
            }
            else {
                $argList += "-$($kv.Key)"; $argList += "$($kv.Value)"
            }
        }
        & $winPS @argList
        exit $LASTEXITCODE
    }
    Write-Warning "Running under PowerShell 7 and could not find Windows PowerShell 5.1 to relaunch. If Connect-CWM fails with a NullReferenceException, run this script with powershell.exe instead of pwsh."
}
#endregion

#region ----- Tunable rule config -------------------------------------------------

# NEW part numbers (from recent syncs) and the OLD parts each one replaced. When a
# sync introduces a new part it does not always remove the old parts, so a company
# can end up billed for both -- charging twice for the same coverage. If a company
# is billed the NewPart AND one or more of the parts it Replaces at the same time,
# that's a duplication. Identifiers are compared case-insensitively. Add new
# mappings here as more sync re-mappings come to light.
$SupersededParts = @(
    @{ NewPart = 'automate-patch'; Replaces = @('automate-patch-workstation', 'automate-patch-server') }
    @{ NewPart = 'cbt-ms-av-edr';  Replaces = @('cbt-ms-av-basic-server', 'cbt-ms-av-basic-workstation') }
)

# Groups of product identifiers whose billed quantity should match per company
# (e.g. seat counts that should track each other). Any company billed differing
# quantities across a group is flagged. Leave empty if you don't want this check.
# Example:
#   @{ Name = 'Workstation seats'; Identifiers = @('cbt-ms-av-basic-workstation','automate-patch-workstation') }
$CorrelatedQuantityGroups = @(
    # @{ Name = '...'; Identifiers = @('...','...') }
)

#endregion

#region ----- Locate + load credentials ------------------------------------------

if (-not $ConfigPath) {
    # Prefer this script's own Config, then any sibling Config\Credentials.xml
    # under the repo root that actually contains a CWMConnectionInfo block.
    $own = Join-Path $PSScriptRoot 'Config\Credentials.xml'
    if (Test-Path $own) {
        $ConfigPath = $own
    }
    else {
        $repoRoot = Split-Path $PSScriptRoot -Parent
        $found = Get-ChildItem -Path $repoRoot -Recurse -Filter 'Credentials.xml' -ErrorAction SilentlyContinue |
            Where-Object {
                try { (Import-Clixml -Path $_.FullName).CWMConnectionInfo } catch { $false }
            } |
            Select-Object -First 1
        if ($found) { $ConfigPath = $found.FullName }
    }
}

if (-not $ConfigPath -or -not (Test-Path $ConfigPath)) {
    Write-Host "No Credentials.xml with a CWMConnectionInfo block was found." -ForegroundColor Red
    Write-Host "Pass -ConfigPath pointing at one of your sync scripts' Config\Credentials.xml." -ForegroundColor Yellow
    exit 1
}

$Config = Import-Clixml -Path $ConfigPath
$CWMConnectionInfo = $Config.CWMConnectionInfo
if (-not $CWMConnectionInfo) {
    Write-Host "Credentials file '$ConfigPath' has no CWMConnectionInfo block." -ForegroundColor Red
    exit 1
}
Write-Host "Using CWM credentials from: $ConfigPath" -ForegroundColor DarkGray

#endregion

#region ----- Connect -------------------------------------------------------------

try {
    if (-not (Get-Module -ListAvailable -Name ConnectWiseManageAPI)) {
        Install-Module 'ConnectWiseManageAPI' -Force -Scope CurrentUser
    }
    Import-Module ConnectWiseManageAPI
    Connect-CWM @CWMConnectionInfo
}
catch {
    Write-Host "Failed to connect to ConnectWise Manage: $_" -ForegroundColor Red
    exit 1
}

#endregion

#region ----- Pull agreements + additions ----------------------------------------

$typeClause = ($AgreementTypes | ForEach-Object { "type/name=`"$_`"" }) -join ' or '
$condition  = "($typeClause) and agreementStatus=`"Active`""

Write-Host "`nFetching active agreements ($($AgreementTypes -join ', '))..." -ForegroundColor Cyan
$agreements = Get-CWMAgreement -condition $condition -all
Write-Host "  $($agreements.Count) agreements" -ForegroundColor Gray

$now = Get-Date

# Flatten every addition into one record set we can slice for each analysis.
$records = New-Object System.Collections.Generic.List[object]
$agInfo  = @{}   # agreementId -> agreement (for company/type lookups)

$i = 0
foreach ($ag in $agreements) {
    $i++
    Write-Progress -Activity 'Reading agreement additions' -Status "$i / $($agreements.Count) - $($ag.company.name)" -PercentComplete ($i / [Math]::Max($agreements.Count,1) * 100)
    $agInfo[$ag.id] = $ag

    $additions = Get-CWMAgreementAddition -parentId $ag.id -all
    foreach ($add in $additions) {
        $ident = if ($add.product) { [string]$add.product.identifier } else { '' }

        # "Currently billing": effective on/before today and not yet cancelled.
        $eff = $add.effectiveDate
        $can = $add.cancelledDate
        $effOk = (-not $eff) -or ([datetime]$eff -le $now)
        $canOk = (-not $can) -or ([datetime]$can -gt $now)
        $active = $effOk -and $canOk

        $qty       = [double]($add.quantity)
        $unitPrice = [double]($add.unitPrice)
        $extPrice  = [double]($add.extPrice)

        # Effective PER-UNIT price for cross-client comparison: prefer the explicit
        # unit price; only when it is blank do we derive per-unit from the extended
        # (total) price. Never compare on the extended/total figure.
        $unitPriceEff = if ($unitPrice -gt 0) { $unitPrice }
                        elseif ($qty -gt 0 -and $extPrice -ne 0) { [Math]::Round($extPrice / $qty, 4) }
                        else { $unitPrice }

        $records.Add([PSCustomObject]@{
            CompanyId     = $ag.company.id
            Company       = $ag.company.name
            AgreementId   = $ag.id
            Agreement     = $ag.name
            AgreementType = $ag.type.name
            AdditionId    = $add.id
            Identifier    = $ident
            IdentLower    = $ident.ToLower()
            Description   = $add.description
            Quantity      = $qty
            UnitPrice     = $unitPrice
            UnitPriceEff  = $unitPriceEff
            ExtPrice      = $extPrice
            UnitCost      = [double]($add.unitCost)
            BillCustomer  = $add.billCustomer
            EffectiveDate = $eff
            CancelledDate = $can
            Active        = $active
        })
    }
}
Write-Progress -Activity 'Reading agreement additions' -Completed

$activeRecords = $records | Where-Object { $_.Active }
$companyIds    = $agreements | ForEach-Object { $_.company.id } | Select-Object -Unique
$totalCompanies = $companyIds.Count

Write-Host "  $($records.Count) additions ($($activeRecords.Count) currently billing) across $totalCompanies companies`n" -ForegroundColor Gray

#endregion

#region ----- Findings collector --------------------------------------------------

$findings = New-Object System.Collections.Generic.List[object]
function Add-Finding {
    param($Category, $Company, $Agreement, $Product, $Detail)
    $findings.Add([PSCustomObject]@{
        Category  = $Category
        Company   = $Company
        Agreement = $Agreement
        Product   = $Product
        Detail    = $Detail
    })
}

#endregion

#region ----- Analyses ------------------------------------------------------------

# Group active records once, per company, for the per-company checks.
$byCompany = $activeRecords | Group-Object CompanyId

# 1) Superseding-part duplication (new part billed alongside the parts it replaced)
foreach ($grp in $byCompany) {
    $idents = $grp.Group | Select-Object -ExpandProperty IdentLower -Unique
    $companyName = $grp.Group[0].Company
    foreach ($map in $SupersededParts) {
        if ($idents -contains $map.NewPart.ToLower()) {
            $overlap = $map.Replaces | Where-Object { $idents -contains $_.ToLower() }
            if ($overlap) {
                Add-Finding 'Superseding-part duplication' $companyName $grp.Group[0].Agreement $map.NewPart `
                    "Billed for new part '$($map.NewPart)' AND old parts it replaced: $($overlap -join ', ')"
            }
        }
    }
}

# 2) Duplicate line items (same product twice on one agreement) --------------------
$activeRecords | Group-Object AgreementId, IdentLower | Where-Object { $_.Count -gt 1 -and $_.Group[0].Identifier } | ForEach-Object {
    $r = $_.Group[0]
    Add-Finding 'Duplicate line item' $r.Company $r.Agreement $r.Identifier `
        "Product appears $($_.Count) times on the same agreement (qty: $(( $_.Group | ForEach-Object { $_.Quantity }) -join ', '))"
}

# 3) Coverage gaps -- scoped to one named book of agreements -----------------------
# Only audit coverage within agreements whose name matches $CoverageAgreementNameLike
# (default: PeopleFirst Support), and compare each such company only against the
# other matching agreements. Additions on a company's *other* agreements don't count.
$pfAgreements   = $agreements | Where-Object { $_.name -like "*$CoverageAgreementNameLike*" }
$pfAgreementIds = @($pfAgreements | ForEach-Object { $_.id })
$pfCompanyIds   = @($pfAgreements | ForEach-Object { $_.company.id } | Select-Object -Unique)
$pfTotal        = $pfCompanyIds.Count

if ($pfTotal -eq 0) {
    Write-Warning "Coverage gap: no agreements matched name like '*$CoverageAgreementNameLike*' -- skipping this check."
}
else {
    # company id -> a matching agreement (for reporting the company that's missing it)
    $pfAgByCompany = @{}
    foreach ($a in $pfAgreements) { if (-not $pfAgByCompany.ContainsKey($a.company.id)) { $pfAgByCompany[$a.company.id] = $a } }

    # Old parts that newer parts replaced (from $SupersededParts) are being phased
    # out, so a company NOT having one is correct -- it migrated to the new part.
    # Exclude them as coverage-gap targets.
    $oldParts = @($SupersededParts | ForEach-Object { $_.Replaces } | ForEach-Object { $_.ToLower() })

    $pfRecords = $activeRecords | Where-Object { $pfAgreementIds -contains $_.AgreementId }
    $pfRecords | Where-Object { $_.Identifier } | Group-Object IdentLower | ForEach-Object {
        if ($oldParts -contains $_.Name) { return }   # skip superseded old parts
        $companiesWith = @($_.Group | Select-Object -ExpandProperty CompanyId -Unique)
        $coverage = $companiesWith.Count / [Math]::Max($pfTotal, 1)
        if ($coverage -ge $CoverageThreshold -and $coverage -lt 1.0 -and $companiesWith.Count -ge $MinSamplesForPricing) {
            $ident = $_.Group[0].Identifier
            $missing = $pfCompanyIds | Where-Object { $_ -notin $companiesWith }
            foreach ($cid in $missing) {
                $ag = $pfAgByCompany[$cid]
                Add-Finding 'Coverage gap' $ag.company.name $ag.name $ident `
                    ("{0:P0} of '{1}' clients are billed '{2}' but this one is not" -f $coverage, $CoverageAgreementNameLike, $ident)
            }
        }
    }
}

# 4) Billable but $0 price ---------------------------------------------------------
$activeRecords | Where-Object { $_.BillCustomer -eq 'Billable' -and $_.UnitPrice -le 0 -and $_.Quantity -gt 0 } | ForEach-Object {
    Add-Finding 'Billable but free' $_.Company $_.Agreement $_.Identifier `
        "Marked Billable but unit price is $($_.UnitPrice) (qty $($_.Quantity))"
}

# 5) Zero quantity (active billable line billing nothing) --------------------------
$activeRecords | Where-Object { $_.BillCustomer -eq 'Billable' -and $_.Quantity -eq 0 } | ForEach-Object {
    Add-Finding 'Zero quantity' $_.Company $_.Agreement $_.Identifier `
        "Billable addition with quantity 0"
}

# 6) Negative margin (cost > price) ------------------------------------------------
$activeRecords | Where-Object { $_.UnitPrice -gt 0 -and $_.UnitCost -gt $_.UnitPrice } | ForEach-Object {
    Add-Finding 'Negative margin' $_.Company $_.Agreement $_.Identifier `
        "Unit cost $($_.UnitCost) exceeds unit price $($_.UnitPrice)"
}

# 7) Price outliers (client off the typical PER-UNIT price for a product) ----------
# Only active, Billable lines with a per-unit price AND a non-zero quantity feed both
# the median and the flagging (a zero-quantity line bills nothing, so its price is
# moot -- it's covered by the Zero quantity check instead). Flag a price at/above
# (1+tol)x the median (above) or at/below the median / (1+tol) (below). With the
# default tol of 1.0 that's >=2x or <=half the median.
$activeRecords | Where-Object { $_.Identifier -and $_.BillCustomer -eq 'Billable' -and $_.UnitPriceEff -gt 0 -and $_.Quantity -gt 0 } |
    Group-Object IdentLower | Where-Object { ($_.Group | Select-Object -ExpandProperty CompanyId -Unique).Count -ge $MinSamplesForPricing } | ForEach-Object {
        $prices = $_.Group | Select-Object -ExpandProperty UnitPriceEff | Sort-Object
        $mid = [int][Math]::Floor($prices.Count / 2)
        $median = if ($prices.Count % 2) { $prices[$mid] } else { ($prices[$mid - 1] + $prices[$mid]) / 2 }
        if ($median -le 0) { return }
        $threshHigh = $median * (1 + $PriceOutlierTolerance)
        $threshLow  = $median / (1 + $PriceOutlierTolerance)
        foreach ($r in $_.Group) {
            $p = $r.UnitPriceEff
            if ($p -gt $threshHigh) {
                Add-Finding 'Price outlier' $r.Company $r.Agreement $r.Identifier `
                    ("Unit price {0:C2} is {1:P0} above median {2:C2}" -f $p, (($p - $median) / $median), $median)
            }
            elseif ($p -lt $threshLow) {
                Add-Finding 'Price outlier' $r.Company $r.Agreement $r.Identifier `
                    ("Unit price {0:C2} is {1:P0} below median {2:C2}" -f $p, (($median - $p) / $median), $median)
            }
        }
    }

# 8) Correlated quantity mismatch (configurable) -----------------------------------
foreach ($grp in $byCompany) {
    $companyName = $grp.Group[0].Company
    foreach ($cg in $CorrelatedQuantityGroups) {
        $present = $grp.Group | Where-Object { $cg.Identifiers -contains $_.Identifier }
        $qtys = $present | Select-Object -ExpandProperty Quantity -Unique
        if ($present.Count -ge 2 -and $qtys.Count -gt 1) {
            $detail = ($present | ForEach-Object { "$($_.Identifier)=$($_.Quantity)" }) -join ', '
            Add-Finding 'Quantity mismatch' $companyName $grp.Group[0].Agreement $cg.Name $detail
        }
    }
}

#endregion

#region ----- Month-over-month agreement totals -----------------------------------
# Reconstruct each agreement's recurring total ("what would be billed") as of the
# 1st of last month and the 1st of this month, by summing the Billable additions
# that were active on each date, then flag agreements whose total moved by at least
# $MoMDollarThreshold or $MoMPercentThreshold. Only $MoMAgreementType is included.
#
# This sees additions that were added or cancelled between the two months (cancelled
# lines stay on the record with a cancelledDate). It cannot see a line that was hard-
# deleted from the agreement, nor an in-place quantity/price edit -- those need saved
# monthly snapshots, a possible future enhancement.

$thisMonthStart = (Get-Date -Day 1).Date
$lastMonthStart = $thisMonthStart.AddMonths(-1)

function Test-ActiveOn {
    param($Record, [datetime]$On)
    $eff = $Record.EffectiveDate
    $can = $Record.CancelledDate
    $effOk = (-not $eff) -or ([datetime]$eff -le $On)
    $canOk = (-not $can) -or ([datetime]$can -gt $On)
    return ($effOk -and $canOk)
}

$recordsByAgreement = $records | Group-Object AgreementId -AsHashTable -AsString
$momAgreements = $agreements | Where-Object { $_.type.name -eq $MoMAgreementType }

$momFindings = New-Object System.Collections.Generic.List[object]
foreach ($ag in $momAgreements) {
    $recs = $recordsByAgreement[[string]$ag.id]
    if (-not $recs) { continue }
    $billable = $recs | Where-Object { $_.BillCustomer -eq 'Billable' }

    $lastTotal = [double](($billable | Where-Object { Test-ActiveOn $_ $lastMonthStart } | Measure-Object ExtPrice -Sum).Sum)
    $thisTotal = [double](($billable | Where-Object { Test-ActiveOn $_ $thisMonthStart } | Measure-Object ExtPrice -Sum).Sum)

    $delta = $thisTotal - $lastTotal
    $pct   = if ($lastTotal -ne 0) { $delta / $lastTotal } else { $null }

    $hitDollar = [Math]::Abs($delta) -ge $MoMDollarThreshold
    $hitPct    = ($null -ne $pct) -and ([Math]::Abs($pct) -ge $MoMPercentThreshold)
    if (-not ($hitDollar -or $hitPct)) { continue }

    $momFindings.Add([PSCustomObject]@{
        Company        = $ag.company.name
        Agreement      = $ag.name
        PercentChange  = if ($null -ne $pct) { [Math]::Round($pct * 100, 1) } else { 'n/a (was $0)' }
        Delta          = [Math]::Round($delta, 2)
        Direction      = if ($delta -gt 0) { 'Increase' } elseif ($delta -lt 0) { 'Decrease' } else { 'None' }
        LastMonth      = $lastMonthStart.ToString('yyyy-MM')
        LastMonthTotal = [Math]::Round($lastTotal, 2)
        ThisMonth      = $thisMonthStart.ToString('yyyy-MM')
        ThisMonthTotal = [Math]::Round($thisTotal, 2)
    })
}

#endregion

#region ----- Output --------------------------------------------------------------

if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }
$stamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'

# Console summary by category.
Write-Host "===== Billing anomaly summary =====" -ForegroundColor Cyan
if ($findings.Count -eq 0) {
    Write-Host "No anomalies found." -ForegroundColor Green
}
else {
    $findings | Group-Object Category | Sort-Object Name | ForEach-Object {
        Write-Host ("  {0,-22} {1,4}" -f $_.Name, $_.Count) -ForegroundColor Yellow
    }
    Write-Host ("  {0,-22} {1,4}" -f 'TOTAL', $findings.Count) -ForegroundColor White
}

Write-Host ("`nAgreement total changes {0} -> {1} (>= `${2:N0} or {3:P0}): {4}" -f `
    $lastMonthStart.ToString('yyyy-MM'), $thisMonthStart.ToString('yyyy-MM'), $MoMDollarThreshold, $MoMPercentThreshold, $momFindings.Count) -ForegroundColor Cyan

# Findings CSV (the actionable report).
$findingsCsv = Join-Path $OutputPath "${stamp}_CWM_BillingAnomalies.csv"
$findings | Sort-Object Category, Company | Export-Csv -Path $findingsCsv -NoTypeInformation -Encoding UTF8
Write-Host "`nFindings report: $findingsCsv" -ForegroundColor Green

# Month-over-month agreement total changes (separate report).
$momCsv = Join-Path $OutputPath "${stamp}_CWM_AgreementMoMChanges.csv"
$momFindings | Sort-Object @{e='Delta';Descending=$false} | Export-Csv -Path $momCsv -NoTypeInformation -Encoding UTF8
Write-Host "MoM changes:     $momCsv" -ForegroundColor Green

# Raw additions CSV (full dataset for ad-hoc review / pivoting).
$rawCsv = Join-Path $OutputPath "${stamp}_CWM_Additions_Raw.csv"
$records | Select-Object Company, Agreement, AgreementType, Identifier, Description, Quantity, UnitPrice, UnitPriceEff, ExtPrice, UnitCost, BillCustomer, EffectiveDate, CancelledDate, Active |
    Sort-Object Company, Identifier | Export-Csv -Path $rawCsv -NoTypeInformation -Encoding UTF8
Write-Host "Raw additions:   $rawCsv" -ForegroundColor Green

#endregion
