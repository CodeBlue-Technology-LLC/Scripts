<#
.SYNOPSIS
    Writes each Capture Client tenant's authentication code into the Automate client
    EDF "SWCCuninstall".

.DESCRIPTION
    Fills the client-level Extra Data Field "SWCCuninstall" with the per-tenant
    "Authentication Code for Tenant" collected by Get-CaptureClientTenantMap.ps1 --
    the 64-hex-character value the console shows under Management > Client Installers
    > Utilities, which SonicWallCaptureClientCleanerUtility.exe takes as -authCode.
    An Automate script can then remove the agent using a single client-scoped
    variable:

        "SonicWallCaptureClientCleanerUtility.exe" -authCode %SWCCuninstall%

    ############################################################################
    #  THIS SCRIPT WRITES TO PRODUCTION AUTOMATE.                              #
    #  It is a DRY RUN unless -Apply is supplied. Nothing is written to        #
    #  SonicWall in either mode -- Capture Client is not contacted at all.     #
    ############################################################################

    Input is the tenant map CSV produced by Get-CaptureClientTenantMap.ps1 (newest in
    Reports\ by default). That is deliberate: only tenants that resolved 1:1 to an
    Automate client are in that file, so an ambiguous or unmatched tenant can never be
    stamped onto a guessed client.

    Safety model:
      - Dry run by default; -Apply performs writes.
      - Every current EDF value is snapshotted to CSV before the first write.
      - Values already correct are skipped, so re-runs converge and cost nothing.
      - Rows whose existing value is non-empty and does NOT look like an
        authentication code are reported as OVERWRITE and skipped unless
        -OverwriteForeignValues is given. See the note below.
      - Every write is verified by reading the value back; a write that reports
        success but does not stick is reported as VERIFY_FAILED, not as success.
      - -ClientId limits the run to specific clients for a pilot.

    Note on existing values: SWCCuninstall has historically held short uppercase
    strings (7 characters, e.g. "WEVLMFL") on some clients. Those are not
    authentication codes -- they look like agent uninstall passwords. Replacing one
    with the 64-character authentication code destroys it, so the script refuses by
    default and lists every case. Decide what those values are before overriding.

.NOTES
    Run under Windows PowerShell 5.1 (powershell.exe), not PowerShell 7.

    Prerequisites:
      - AutomateAPI module, and the credential store written by
        Get-CaptureClientTenantMap.ps1 (Config\CaptureClient-Credentials.xml).
      - An Automate account with permission to EDIT CLIENTS. Read-only integration
        accounts return HTTP 403 on client EDF writes while still being able to read
        them, and can even write computer records -- so this is easy to misdiagnose.
        The script preflights this explicitly and tells you before touching anything.
      - The EDF must already exist (Clients tab). It is located by title, not by a
        hardcoded definition id.

    The Automate REST API expects RFC 6902 JSON Patch for EDF updates:
        PATCH /cwa/api/v1/clients/{clientId}/extrafields/{extraFieldDefinitionId}
        Content-Type: application/json          (NOT application/json-patch+json -> 415)
        Body: [ { "op":"replace", "path":"/TextFieldSettings/Value", "value":"..." } ]
    Any other body shape returns 500, so a 500 here means the body, a 403 means rights.

.PARAMETER Apply
    Perform the writes. Without it the script reports exactly what it would change.

.PARAMETER TenantMapCsv
    Tenant map from Get-CaptureClientTenantMap.ps1. Defaults to the newest
    *_SWCC_TenantMap.csv in the Reports directory.

.PARAMETER ClientId
    Limit the run to these Automate client ids. Use for a pilot before the full batch.

.PARAMETER AuthCodeFieldTitle
    Title of the EDF receiving the authentication code. Default 'SWCCuninstall'.

.PARAMETER WriteTenantId
    Also write the Capture Client console tenant GUID into the EDF named by
    -TenantIdFieldTitle. Off by default. Worth turning on once: it gives future runs
    an exact key to join on instead of relying on company-name matching.

.PARAMETER TenantIdFieldTitle
    Title of the EDF receiving the console tenant GUID. Default 'SWCCtenantId'.
    Only used with -WriteTenantId.

.PARAMETER OverwriteForeignValues
    Replace existing SWCCuninstall values that do not look like authentication codes.
    Read the note in the description before using this.

.PARAMETER AutomateCredential
    Override the stored Automate credential. Use this to run under an account that has
    client-edit rights without changing the stored integration credential.

.PARAMETER MaxAgeHours
    Refuse to use a tenant map older than this many hours, so stale codes are not
    written. Default 24. Set 0 to disable the check.

.EXAMPLE
    .\Set-AutomateCaptureClientAuthCode.ps1

    Dry run against the newest tenant map. Shows every intended change.

.EXAMPLE
    .\Set-AutomateCaptureClientAuthCode.ps1 -ClientId 447 -Apply

    Pilot: writes one client and verifies by read-back.

.EXAMPLE
    .\Set-AutomateCaptureClientAuthCode.ps1 -Apply -AutomateCredential (Get-Credential)

    Full batch under an account with client-edit rights.
#>

[CmdletBinding()]
param(
    [switch]$Apply,

    [string]$TenantMapCsv,

    [int[]]$ClientId,

    [string]$AuthCodeFieldTitle = 'SWCCuninstall',

    [switch]$WriteTenantId,

    [string]$TenantIdFieldTitle = 'SWCCtenantId',

    [switch]$OverwriteForeignValues,

    [System.Management.Automation.PSCredential]$AutomateCredential,

    [int]$MaxAgeHours = 24
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ConfigPath = Join-Path $PSScriptRoot 'Config\CaptureClient-Credentials.xml'
$ReportsDir = Join-Path $PSScriptRoot 'Reports'

function Write-Step { param([string]$Message, [string]$Color = 'Cyan') Write-Host $Message -ForegroundColor $Color }

# What a Capture Client authentication code looks like. Used both to sanity-check the
# map and to recognize a value already in Automate as one of ours.
$AuthCodePattern = '^[0-9A-Fa-f]{32,}$'

#region Load configuration

if (-not (Test-Path $ConfigPath)) {
    Write-Host "ERROR: No credential store at $ConfigPath" -ForegroundColor Red
    Write-Host '       Run Get-CaptureClientTenantMap.ps1 once first to create it.' -ForegroundColor Red
    exit 2
}
$config = Import-Clixml -Path $ConfigPath

if (-not $AutomateCredential) {
    $AutomateCredential = New-Object System.Management.Automation.PSCredential(
        $config.Automate.Username, $config.Automate.Password)
}

#endregion

#region Resolve the tenant map

if (-not $TenantMapCsv) {
    if (-not (Test-Path $ReportsDir)) {
        Write-Host 'ERROR: No Reports directory. Run Get-CaptureClientTenantMap.ps1 first.' -ForegroundColor Red
        exit 2
    }
    $newest = Get-ChildItem -Path $ReportsDir -Filter '*_SWCC_TenantMap.csv' |
        Sort-Object LastWriteTime | Select-Object -Last 1
    if (-not $newest) {
        Write-Host "ERROR: No *_SWCC_TenantMap.csv in $ReportsDir." -ForegroundColor Red
        Write-Host '       Run Get-CaptureClientTenantMap.ps1 to produce one.' -ForegroundColor Red
        exit 2
    }
    $TenantMapCsv = $newest.FullName
}

if (-not (Test-Path $TenantMapCsv)) {
    Write-Host "ERROR: Tenant map not found: $TenantMapCsv" -ForegroundColor Red
    exit 2
}

$mapAgeHours = ((Get-Date) - (Get-Item $TenantMapCsv).LastWriteTime).TotalHours
if ($MaxAgeHours -gt 0 -and $mapAgeHours -gt $MaxAgeHours) {
    Write-Host ("ERROR: Tenant map is {0:N1} hours old (limit {1}). Re-run Get-CaptureClientTenantMap.ps1," -f $mapAgeHours, $MaxAgeHours) -ForegroundColor Red
    Write-Host '       or pass -MaxAgeHours 0 to override.' -ForegroundColor Red
    exit 2
}

$mapRows = @(Import-Csv -Path $TenantMapCsv)
if ($mapRows.Count -eq 0) {
    Write-Host "ERROR: Tenant map $TenantMapCsv is empty." -ForegroundColor Red
    exit 2
}

# A malformed code must never reach production, even if the map was hand-edited.
$badCode = @($mapRows | Where-Object { $_.CleanerAuthCode -notmatch $AuthCodePattern })
if ($badCode.Count -gt 0) {
    Write-Host "ERROR: $($badCode.Count) row(s) in the map do not carry a valid authentication code:" -ForegroundColor Red
    $badCode | Select-Object -First 10 | ForEach-Object {
        Write-Host ("       {0} -> '{1}'" -f $_.AutomateClientName, $_.CleanerAuthCode) -ForegroundColor Red
    }
    exit 2
}

if ($ClientId) {
    $wanted = @{}
    foreach ($id in $ClientId) { $wanted["$id"] = $true }
    $mapRows = @($mapRows | Where-Object { $wanted.ContainsKey([string]$_.AutomateClientId) })
    if ($mapRows.Count -eq 0) {
        Write-Host 'ERROR: None of the requested -ClientId values are in the tenant map.' -ForegroundColor Red
        exit 2
    }
}

#endregion

Write-Host ''
if ($Apply) { Write-Step '=== Capture Client auth-code writer -- APPLY (writes to Automate) ===' 'Yellow' }
else { Write-Step '=== Capture Client auth-code writer -- DRY RUN (no changes) ===' 'Cyan' }
Write-Host ("  Tenant map : {0}" -f (Split-Path -Leaf $TenantMapCsv)) -ForegroundColor Gray
Write-Host ("  Age        : {0:N1} h" -f $mapAgeHours) -ForegroundColor Gray
Write-Host ("  Clients    : {0}" -f $mapRows.Count) -ForegroundColor Gray
Write-Host ("  Fields     : {0}{1}" -f $AuthCodeFieldTitle, $(if ($WriteTenantId) { " + $TenantIdFieldTitle" } else { '' })) -ForegroundColor Gray
Write-Host ''

#region Connect to Automate

try {
    if (-not (Get-Module -ListAvailable -Name AutomateAPI)) {
        throw 'The AutomateAPI module is not installed. Run: Install-Module AutomateAPI -Scope CurrentUser'
    }
    Import-Module AutomateAPI -ErrorAction Stop

    Write-Step 'Connecting to Automate...'
    Connect-AutomateAPI -Server $config.Automate.Server -Credential $AutomateCredential `
        -ClientID $config.Automate.ClientId -Force -Quiet | Out-Null
}
catch {
    Write-Host "ERROR: Automate connection failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 2
}

# The module exposes no EDF setter, so raw REST is required for the writes. Reuse the
# module's own session rather than minting a second token against the same account.
$automateModule = Get-Module AutomateAPI
$tokenHeader = & $automateModule { $Script:CWAToken }
$automateServer = & $automateModule { $Script:CWAServer }

$restHeaders = @{}
foreach ($key in $tokenHeader.Keys) { $restHeaders[$key] = $tokenHeader[$key] }
$restHeaders['clientId'] = $config.Automate.ClientId

function Set-ClientExtraField {
    <#
        RFC 6902 JSON Patch against the client EDF route. Returns a status string:
        OK / FORBIDDEN / BADREQUEST / ERROR:<detail>.
    #>
    param(
        [int]$TargetClientId,
        [int]$DefinitionId,
        [string]$Value
    )

    $uri = "$automateServer/cwa/api/v1/clients/$TargetClientId/extrafields/$DefinitionId"
    $body = ConvertTo-Json @(@{ op = 'replace'; path = '/TextFieldSettings/Value'; value = $Value }) -Depth 5

    try {
        $null = Invoke-RestMethod -Uri $uri -Method Patch -Headers $restHeaders -Body $body -ContentType 'application/json'
        return 'OK'
    }
    catch {
        $code = $null
        try { $code = $_.Exception.Response.StatusCode.value__ } catch {}
        switch ($code) {
            403 { return 'FORBIDDEN' }
            500 { return 'BADREQUEST' }
            default { return "ERROR:$code $($_.Exception.Message)" }
        }
    }
}

function Resolve-FieldDefinition {
    param([object[]]$Fields, [string]$Title)

    $hit = @($Fields | Where-Object { $_.Title -eq $Title })
    if ($hit.Count -eq 0) {
        $hit = @($Fields | Where-Object { $_.Title -and $_.Title.Trim() -eq $Title.Trim() })
    }
    if ($hit.Count -eq 0) { return $null }
    return $hit[0]
}

#endregion

#region Resolve EDF definition ids by title

Write-Step 'Resolving EDF definitions...'
$probeClientId = [int]$mapRows[0].AutomateClientId
try {
    $probeFields = @(Get-AutomateClientExtraFields -ClientId $probeClientId)
}
catch {
    Write-Host "ERROR: Could not read client EDFs: $($_.Exception.Message)" -ForegroundColor Red
    exit 2
}

$authCodeField = Resolve-FieldDefinition -Fields $probeFields -Title $AuthCodeFieldTitle
$tenantIdField = if ($WriteTenantId) { Resolve-FieldDefinition -Fields $probeFields -Title $TenantIdFieldTitle } else { $null }

if (-not $authCodeField) {
    Write-Host "ERROR: No client EDF titled '$AuthCodeFieldTitle'." -ForegroundColor Red
    Write-Host "       Client EDF titles available: $((($probeFields | ForEach-Object { $_.Title }) | Sort-Object) -join ', ')" -ForegroundColor Yellow
    exit 2
}
if ($WriteTenantId -and -not $tenantIdField) {
    Write-Host "ERROR: No client EDF titled '$TenantIdFieldTitle'." -ForegroundColor Red
    exit 2
}

foreach ($field in @($authCodeField, $tenantIdField)) {
    if (-not $field) { continue }
    if ($field.IsReadOnly) {
        Write-Host "ERROR: EDF '$($field.Title)' is marked read-only in Automate; it cannot be written." -ForegroundColor Red
        exit 2
    }
    if (-not $field.PSObject.Properties['TextFieldSettings']) {
        Write-Host "ERROR: EDF '$($field.Title)' is not a text field (display format $($field.DisplayFormat.Name))." -ForegroundColor Red
        exit 2
    }
    Write-Host ("  {0} -> definition {1}" -f $field.Title, $field.ExtraFieldDefinitionId) -ForegroundColor Gray
}

#endregion

#region Preflight: can this account actually write a client EDF?

if ($Apply) {
    Write-Step 'Preflighting write permission...'
    # Write the current value back to itself: succeeds without changing anything,
    # fails with the same 403 a real write would hit.
    $currentProbe = [string](Resolve-FieldDefinition -Fields $probeFields -Title $authCodeField.Title).TextFieldSettings.Value
    $probeResult = Set-ClientExtraField -TargetClientId $probeClientId -DefinitionId $authCodeField.ExtraFieldDefinitionId -Value $currentProbe

    if ($probeResult -eq 'FORBIDDEN') {
        Write-Host ''
        Write-Host 'ERROR: HTTP 403 writing a client EDF. The Automate account cannot edit clients.' -ForegroundColor Red
        Write-Host ("       Account in use: {0}" -f $AutomateCredential.UserName) -ForegroundColor Red
        Write-Host ''
        Write-Host '       This is a permissions boundary, not a bad request. The same account can read' -ForegroundColor Yellow
        Write-Host '       client EDFs and can write COMPUTER records, so the token itself is fine.' -ForegroundColor Yellow
        Write-Host ''
        Write-Host '       Fix either way:' -ForegroundColor Yellow
        Write-Host '         a) Grant the account client-edit rights in Automate' -ForegroundColor Yellow
        Write-Host '            (System > Users > User Classes -> Clients: Add/Edit), or' -ForegroundColor Yellow
        Write-Host '         b) Re-run under an admin account:' -ForegroundColor Yellow
        Write-Host '            .\Set-AutomateCaptureClientAuthCode.ps1 -Apply -AutomateCredential (Get-Credential)' -ForegroundColor Yellow
        exit 3
    }
    elseif ($probeResult -ne 'OK') {
        Write-Host "ERROR: Preflight write failed: $probeResult" -ForegroundColor Red
        exit 3
    }
    Write-Host '  Write permission confirmed.' -ForegroundColor Green
}

#endregion

#region Read current state

Write-Step 'Reading current field values...'
$plan = @()
$done = 0

foreach ($row in $mapRows) {
    $done++
    Write-Progress -Activity 'Reading Automate client EDFs' -Status $row.AutomateClientName `
        -PercentComplete ([int](100 * $done / $mapRows.Count))

    $targetId = [int]$row.AutomateClientId
    try { $fields = @(Get-AutomateClientExtraFields -ClientId $targetId) }
    catch {
        $plan += [pscustomobject]@{
            AutomateClientId = $targetId; AutomateClientName = $row.AutomateClientName
            TenantName = $row.TenantName; TenantId = $row.TenantId
            FieldTitle = $AuthCodeFieldTitle; CurrentValue = ''; DesiredValue = $row.CleanerAuthCode
            Action = 'READ_FAILED'; Detail = $_.Exception.Message
        }
        continue
    }

    $work = @(, @($authCodeField, $row.CleanerAuthCode))
    if ($WriteTenantId) { $work += , @($tenantIdField, $row.TenantId) }

    foreach ($item in $work) {
        $definition = $item[0]
        $desired = [string]$item[1]

        $current = ''
        $live = Resolve-FieldDefinition -Fields $fields -Title $definition.Title
        if ($live -and $live.PSObject.Properties['TextFieldSettings']) { $current = [string]$live.TextFieldSettings.Value }

        $action =
            if ($current -eq $desired) { 'ALREADY_SET' }
            elseif (-not $current) { 'WRITE' }
            elseif ($definition.Title -eq $AuthCodeFieldTitle -and $current -notmatch $AuthCodePattern -and -not $OverwriteForeignValues) { 'OVERWRITE_BLOCKED' }
            else { 'REPLACE' }

        $plan += [pscustomobject]@{
            AutomateClientId   = $targetId
            AutomateClientName = $row.AutomateClientName
            TenantName         = $row.TenantName
            TenantId           = $row.TenantId
            FieldTitle         = $definition.Title
            DefinitionId       = $definition.ExtraFieldDefinitionId
            CurrentValue       = $current
            DesiredValue       = $desired
            Action             = $action
            Detail             = ''
        }
    }
}
Write-Progress -Activity 'Reading Automate client EDFs' -Completed

#endregion

#region Report the plan

$blocked = @($plan | Where-Object { $_.Action -eq 'OVERWRITE_BLOCKED' })
$readFailed = @($plan | Where-Object { $_.Action -eq 'READ_FAILED' })
$toWrite = @($plan | Where-Object { $_.Action -in @('WRITE', 'REPLACE') })
$alreadySet = @($plan | Where-Object { $_.Action -eq 'ALREADY_SET' })

Write-Host ''
Write-Step '--- Plan ---'
Write-Host ("  Already correct   : {0}" -f $alreadySet.Count) -ForegroundColor Gray
Write-Host ("  To write (empty)  : {0}" -f @($plan | Where-Object Action -eq 'WRITE').Count) -ForegroundColor Green
Write-Host ("  To replace        : {0}" -f @($plan | Where-Object Action -eq 'REPLACE').Count) -ForegroundColor Yellow
if ($blocked.Count) { Write-Host ("  Blocked overwrite : {0}" -f $blocked.Count) -ForegroundColor Red }
if ($readFailed.Count) { Write-Host ("  Read failed       : {0}" -f $readFailed.Count) -ForegroundColor Red }

if ($blocked.Count -gt 0) {
    Write-Host ''
    Write-Step '--- BLOCKED: existing value is not an authentication code ---' 'Red'
    Write-Host '    These clients already hold something else in this field. Find out what it is' -ForegroundColor Yellow
    Write-Host '    before replacing it; pass -OverwriteForeignValues once you have decided.' -ForegroundColor Yellow
    Write-Host ''
    $blocked | Sort-Object AutomateClientName |
        Format-Table -AutoSize AutomateClientId, AutomateClientName, FieldTitle, CurrentValue |
        Out-String -Width 200 | Write-Host
}

if ($readFailed.Count -gt 0) {
    Write-Host ''
    Write-Step '--- READ FAILED ---' 'Red'
    $readFailed | Format-Table -AutoSize AutomateClientId, AutomateClientName, Detail | Out-String -Width 200 | Write-Host
}

if (-not $Apply) {
    if ($toWrite.Count -gt 0) {
        Write-Host ''
        Write-Step '--- Would write ---'
        $toWrite | Sort-Object AutomateClientName |
            Format-Table -AutoSize AutomateClientId, AutomateClientName, FieldTitle, @{ n = 'New'; e = { $_.DesiredValue } } |
            Out-String -Width 200 | Write-Host
    }
    Write-Host ''
    Write-Step 'DRY RUN complete. Nothing was written. Re-run with -Apply to commit.' 'Cyan'
    Write-Host ''
    exit $(if ($blocked.Count -or $readFailed.Count) { 1 } else { 0 })
}

#endregion

#region Snapshot, then write

if (-not (Test-Path $ReportsDir)) { $null = New-Item -ItemType Directory -Path $ReportsDir -Force }
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'

$snapshotPath = Join-Path $ReportsDir "${stamp}_SWCC_EDF_Snapshot.csv"
$plan | Select-Object AutomateClientId, AutomateClientName, FieldTitle, DefinitionId, CurrentValue, DesiredValue, Action |
    Export-Csv -Path $snapshotPath -NoTypeInformation -Encoding UTF8
Write-Host ''
Write-Host ("  Snapshot of current values: {0}" -f $snapshotPath) -ForegroundColor Gray

if ($toWrite.Count -eq 0) {
    Write-Host ''
    Write-Step 'Nothing to write; every field is already correct.' 'Green'
    Write-Host ''
    exit $(if ($blocked.Count -or $readFailed.Count) { 1 } else { 0 })
}

Write-Host ''
Write-Step ("Writing {0} field value(s)..." -f $toWrite.Count) 'Yellow'

$results = @()
$done = 0
foreach ($item in $toWrite) {
    $done++
    Write-Progress -Activity 'Writing Automate client EDFs' -Status $item.AutomateClientName `
        -PercentComplete ([int](100 * $done / $toWrite.Count))

    $status = Set-ClientExtraField -TargetClientId $item.AutomateClientId -DefinitionId $item.DefinitionId -Value $item.DesiredValue

    if ($status -eq 'OK') {
        # A 200 is not proof the value stuck; read it back.
        try {
            $after = @(Get-AutomateClientExtraFields -ClientId $item.AutomateClientId)
            $live = Resolve-FieldDefinition -Fields $after -Title $item.FieldTitle
            $readBack = if ($live -and $live.PSObject.Properties['TextFieldSettings']) { [string]$live.TextFieldSettings.Value } else { '' }
            $status = if ($readBack -eq $item.DesiredValue) { 'WRITTEN' } else { 'VERIFY_FAILED' }
        }
        catch { $status = 'VERIFY_FAILED' }
    }

    $results += [pscustomobject]@{
        AutomateClientId   = $item.AutomateClientId
        AutomateClientName = $item.AutomateClientName
        TenantName         = $item.TenantName
        FieldTitle         = $item.FieldTitle
        PreviousValue      = $item.CurrentValue
        NewValue           = $item.DesiredValue
        Result             = $status
    }

    $color = switch ($status) { 'WRITTEN' { 'Green' } 'VERIFY_FAILED' { 'Red' } default { 'Red' } }
    Write-Host ("  {0,-40} {1,-16} {2}" -f $item.AutomateClientName, $item.FieldTitle, $status) -ForegroundColor $color
}
Write-Progress -Activity 'Writing Automate client EDFs' -Completed

$resultPath = Join-Path $ReportsDir "${stamp}_SWCC_EDF_WriteResults.csv"
$results | Export-Csv -Path $resultPath -NoTypeInformation -Encoding UTF8

$written = @($results | Where-Object Result -eq 'WRITTEN')
$failed = @($results | Where-Object Result -ne 'WRITTEN')

Write-Host ''
Write-Step '--- Result ---'
Write-Host ("  Written : {0}" -f $written.Count) -ForegroundColor Green
if ($failed.Count) { Write-Host ("  Failed  : {0}" -f $failed.Count) -ForegroundColor Red }
Write-Host ("  Log     : {0}" -f $resultPath) -ForegroundColor Gray
Write-Host ''

exit $(if ($failed.Count -or $blocked.Count -or $readFailed.Count) { 1 } else { 0 })

#endregion
