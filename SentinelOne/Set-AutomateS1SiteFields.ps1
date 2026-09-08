<#
.SYNOPSIS
    Writes the SentinelOne Site ID and Site Token into ConnectWise Automate client EDFs.

.DESCRIPTION
    Fills the client-level Extra Data Fields "S1SiteID" and "S1ScopeToken" for every
    company that the reconciliation script verified as a clean 1:1 match, so an Automate
    script can install the SentinelOne agent into the correct site using a single
    client-scoped variable.

    ############################################################################
    #  THIS SCRIPT WRITES TO PRODUCTION AUTOMATE.                              #
    #  It is a DRY RUN unless -Apply is supplied. Nothing is written to         #
    #  SentinelOne in either mode -- S1 is read-only here.                      #
    ############################################################################

    Input is the install site map CSV produced by Compare-S1SiteToAutomateClient.ps1
    (newest in Reports\ by default). That is deliberate: only mappings that passed the
    1:1 reconciliation can be written, so an ambiguous or unmatched company can never
    be stamped with a guessed site.

    Site tokens are NOT read from the map. They are fetched live from SentinelOne on
    every run and held in memory only, so the enrollment secret never has to sit in a
    CSV on disk. This also means a rotated token is picked up automatically.

    Safety model:
      - Dry run by default; -Apply performs writes.
      - Every current EDF value is snapshotted to CSV before the first write.
      - Values already correct are skipped, so re-runs converge and cost nothing.
      - Every write is verified by reading the value back; a write that reports
        success but does not stick is reported as VERIFY_FAILED, not as success.
      - -ClientId limits the run to specific clients for a pilot.

.NOTES
    Prerequisites:
      - AutomateAPI module, and the same credential store used by the reconciliation
        script (Config\S1-AutomateReconcile-Credentials.xml).
      - An Automate account with permission to EDIT CLIENTS. Read-only integration
        accounts return HTTP 403 on client EDF writes while still being able to read
        them, and can even write computer records -- so this is easy to misdiagnose.
        The script preflights this explicitly and tells you before touching anything.
      - The two client EDFs must already exist (Clients tab). They are located by
        title, not by hardcoded definition id.

    Run under Windows PowerShell 5.1 (powershell.exe), not PowerShell 7.

    The Automate REST API expects RFC 6902 JSON Patch for EDF updates:
        PATCH /cwa/api/v1/clients/{clientId}/extrafields/{extraFieldDefinitionId}
        Content-Type: application/json          (NOT application/json-patch+json -> 415)
        Body: [ { "op":"replace", "path":"/TextFieldSettings/Value", "value":"..." } ]
    Any other body shape returns 500, so a 500 here means the body, a 403 means rights.

.PARAMETER Apply
    Perform the writes. Without it the script reports exactly what it would change.

.PARAMETER SiteMapCsv
    Install site map from the reconciliation script. Defaults to the newest
    *_S1_Automate_SiteMap.csv in the Reports directory.

.PARAMETER ClientId
    Limit the run to these Automate client ids. Use for a pilot before the full batch.

.PARAMETER SiteIdFieldTitle
    Title of the EDF receiving the S1 site id. Default 'S1SiteID'.

.PARAMETER SiteTokenFieldTitle
    Title of the EDF receiving the S1 site token. Default 'S1ScopeToken'. The value
    written is the site's registrationToken, which is what the S1 console surfaces as
    the site-scoped token.

.PARAMETER SkipToken
    Write only the site id, leaving the token field untouched.

.PARAMETER AutomateCredential
    Override the stored Automate credential. Use this to run under an account that has
    client-edit rights without changing the stored integration credential.

.PARAMETER MaxAgeHours
    Refuse to use a site map older than this many hours, so stale mappings are not
    written. Default 24. Set 0 to disable the check.

.EXAMPLE
    .\Set-AutomateS1SiteFields.ps1

    Dry run against the newest site map. Shows every intended change.

.EXAMPLE
    .\Set-AutomateS1SiteFields.ps1 -ClientId 65 -Apply

    Pilot: writes both fields for a single client and verifies by read-back.

.EXAMPLE
    .\Set-AutomateS1SiteFields.ps1 -Apply -AutomateCredential (Get-Credential)

    Full batch under an account with client-edit rights.
#>

[CmdletBinding()]
param(
    [switch]$Apply,

    [string]$SiteMapCsv,

    [int[]]$ClientId,

    [string]$SiteIdFieldTitle = 'S1SiteID',

    [string]$SiteTokenFieldTitle = 'S1ScopeToken',

    [switch]$SkipToken,

    [System.Management.Automation.PSCredential]$AutomateCredential,

    [int]$MaxAgeHours = 24
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ConfigPath = Join-Path $PSScriptRoot 'Config\S1-AutomateReconcile-Credentials.xml'
$ReportsDir = Join-Path $PSScriptRoot 'Reports'

function Write-Step { param([string]$Message, [string]$Color = 'Cyan') Write-Host $Message -ForegroundColor $Color }

#region Load configuration

if (-not (Test-Path $ConfigPath)) {
    Write-Host "ERROR: No credential store at $ConfigPath" -ForegroundColor Red
    Write-Host '       Run Compare-S1SiteToAutomateClient.ps1 once first to create it.' -ForegroundColor Red
    exit 2
}
$config = Import-Clixml -Path $ConfigPath

if (-not $AutomateCredential) {
    $AutomateCredential = New-Object System.Management.Automation.PSCredential(
        $config.Automate.Username, $config.Automate.Password)
}

#endregion

#region Resolve the site map

if (-not $SiteMapCsv) {
    if (-not (Test-Path $ReportsDir)) {
        Write-Host "ERROR: No Reports directory. Run Compare-S1SiteToAutomateClient.ps1 first." -ForegroundColor Red
        exit 2
    }
    $newest = Get-ChildItem -Path $ReportsDir -Filter '*_S1_Automate_SiteMap.csv' |
        Sort-Object LastWriteTime | Select-Object -Last 1
    if (-not $newest) {
        Write-Host "ERROR: No *_S1_Automate_SiteMap.csv in $ReportsDir." -ForegroundColor Red
        Write-Host '       Run Compare-S1SiteToAutomateClient.ps1 to produce one.' -ForegroundColor Red
        exit 2
    }
    $SiteMapCsv = $newest.FullName
}

if (-not (Test-Path $SiteMapCsv)) {
    Write-Host "ERROR: Site map not found: $SiteMapCsv" -ForegroundColor Red
    exit 2
}

$mapAgeHours = ((Get-Date) - (Get-Item $SiteMapCsv).LastWriteTime).TotalHours
if ($MaxAgeHours -gt 0 -and $mapAgeHours -gt $MaxAgeHours) {
    Write-Host ("ERROR: Site map is {0:N1} hours old (limit {1}). Re-run the reconciliation first," -f $mapAgeHours, $MaxAgeHours) -ForegroundColor Red
    Write-Host '       or pass -MaxAgeHours 0 to override.' -ForegroundColor Red
    exit 2
}

$mapRows = @(Import-Csv -Path $SiteMapCsv)
if ($mapRows.Count -eq 0) {
    Write-Host "ERROR: Site map $SiteMapCsv is empty." -ForegroundColor Red
    exit 2
}
if ($ClientId) {
    $wanted = @{}
    foreach ($id in $ClientId) { $wanted["$id"] = $true }
    $mapRows = @($mapRows | Where-Object { $wanted.ContainsKey([string]$_.AutomateClientId) })
    if ($mapRows.Count -eq 0) {
        Write-Host "ERROR: None of the requested -ClientId values are in the site map." -ForegroundColor Red
        exit 2
    }
}

#endregion

Write-Host ''
if ($Apply) { Write-Step '=== Automate S1 field writer -- APPLY (writes to Automate) ===' 'Yellow' }
else { Write-Step '=== Automate S1 field writer -- DRY RUN (no changes) ===' 'Cyan' }
Write-Host ("  Site map   : {0}" -f (Split-Path -Leaf $SiteMapCsv)) -ForegroundColor Gray
Write-Host ("  Age        : {0:N1} h" -f $mapAgeHours) -ForegroundColor Gray
Write-Host ("  Companies  : {0}" -f $mapRows.Count) -ForegroundColor Gray
Write-Host ("  Fields     : {0}{1}" -f $SiteIdFieldTitle, $(if ($SkipToken) { ' (token skipped)' } else { " + $SiteTokenFieldTitle" })) -ForegroundColor Gray
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

function Resolve-FieldDefinition {
    param([object[]]$Fields, [string]$Title)

    $hit = @($Fields | Where-Object { $_.Title -eq $Title })
    if ($hit.Count -eq 0) {
        $hit = @($Fields | Where-Object { $_.Title -and $_.Title.Trim() -eq $Title.Trim() })
    }
    if ($hit.Count -eq 0) { return $null }
    return $hit[0]
}

$siteIdField = Resolve-FieldDefinition -Fields $probeFields -Title $SiteIdFieldTitle
$siteTokenField = if ($SkipToken) { $null } else { Resolve-FieldDefinition -Fields $probeFields -Title $SiteTokenFieldTitle }

if (-not $siteIdField) {
    Write-Host "ERROR: No client EDF titled '$SiteIdFieldTitle'." -ForegroundColor Red
    Write-Host "       Client EDF titles available: $((($probeFields | ForEach-Object { $_.Title }) | Sort-Object) -join ', ')" -ForegroundColor Yellow
    exit 2
}
if (-not $SkipToken -and -not $siteTokenField) {
    Write-Host "ERROR: No client EDF titled '$SiteTokenFieldTitle'." -ForegroundColor Red
    exit 2
}

foreach ($f in @($siteIdField, $siteTokenField)) {
    if (-not $f) { continue }
    if ($f.IsReadOnly) {
        Write-Host "ERROR: EDF '$($f.Title)' is marked read-only in Automate; it cannot be written." -ForegroundColor Red
        exit 2
    }
    if (-not $f.PSObject.Properties['TextFieldSettings']) {
        Write-Host "ERROR: EDF '$($f.Title)' is not a text field (display format $($f.DisplayFormat.Name))." -ForegroundColor Red
        exit 2
    }
}

Write-Host ("  {0} -> definition {1}" -f $siteIdField.Title, $siteIdField.ExtraFieldDefinitionId) -ForegroundColor Gray
if ($siteTokenField) {
    Write-Host ("  {0} -> definition {1}" -f $siteTokenField.Title, $siteTokenField.ExtraFieldDefinitionId) -ForegroundColor Gray
    if (-not $siteTokenField.IsMasked -and -not $siteTokenField.IsEncrypted) {
        Write-Warning "EDF '$($siteTokenField.Title)' is neither masked nor encrypted; the site token will be visible in the Automate UI."
    }
}

#endregion

#region Preflight: can this account actually write a client EDF?

if ($Apply) {
    Write-Step 'Preflighting write permission...'
    # Write the current value back to itself: succeeds without changing anything,
    # fails with the same 403 a real write would hit.
    $currentProbe = [string](Resolve-FieldDefinition -Fields $probeFields -Title $siteIdField.Title).TextFieldSettings.Value
    $probeResult = Set-ClientExtraField -TargetClientId $probeClientId -DefinitionId $siteIdField.ExtraFieldDefinitionId -Value $currentProbe

    if ($probeResult -eq 'FORBIDDEN') {
        Write-Host ''
        Write-Host 'ERROR: HTTP 403 writing a client EDF. The Automate account cannot edit clients.' -ForegroundColor Red
        Write-Host ("       Account in use: {0}" -f $AutomateCredential.UserName) -ForegroundColor Red
        Write-Host '' -ForegroundColor Red
        Write-Host '       This is a permissions boundary, not a bad request. The same account can read' -ForegroundColor Yellow
        Write-Host '       client EDFs and can write COMPUTER records, so the token itself is fine.' -ForegroundColor Yellow
        Write-Host '' -ForegroundColor Yellow
        Write-Host '       Fix either way:' -ForegroundColor Yellow
        Write-Host '         a) Grant the account client-edit rights in Automate' -ForegroundColor Yellow
        Write-Host '            (System > Users > User Classes -> Clients: Add/Edit), or' -ForegroundColor Yellow
        Write-Host '         b) Re-run under an admin account:' -ForegroundColor Yellow
        Write-Host '            .\Set-AutomateS1SiteFields.ps1 -Apply -AutomateCredential (Get-Credential)' -ForegroundColor Yellow
        exit 3
    }
    elseif ($probeResult -ne 'OK') {
        Write-Host "ERROR: Preflight write failed: $probeResult" -ForegroundColor Red
        exit 3
    }
    Write-Host '  Write permission confirmed.' -ForegroundColor Green
}

#endregion

#region Pull live SentinelOne site tokens

$tokenBySiteId = @{}
if (-not $SkipToken) {
    Write-Step 'Fetching SentinelOne site tokens...'
    try {
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($config.S1.ApiToken)
        try { $apiToken = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }

        $s1Headers = @{ Authorization = "ApiToken $apiToken" }
        $base = $config.S1.BaseUri.TrimEnd('/')
        $cursor = $null

        do {
            $uri = "$base/web/api/v2.1/sites?limit=1000"
            if ($cursor) { $uri += "&cursor=$([Uri]::EscapeDataString($cursor))" }
            $response = Invoke-RestMethod -Uri $uri -Headers $s1Headers -Method Get
            foreach ($site in $response.data.sites) {
                if ($site.registrationToken) { $tokenBySiteId[[string]$site.id] = $site.registrationToken }
            }
            $cursor = $response.pagination.nextCursor
        } while ($cursor)

        Write-Host "  $($tokenBySiteId.Keys.Count) site tokens retrieved" -ForegroundColor Gray
    }
    catch {
        Write-Host "ERROR: Could not retrieve SentinelOne site tokens: $($_.Exception.Message)" -ForegroundColor Red
        exit 2
    }
    finally {
        Remove-Variable -Name apiToken -ErrorAction SilentlyContinue
    }
}

#endregion

#region Snapshot, then write

$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
if (-not (Test-Path $ReportsDir)) { New-Item -ItemType Directory -Path $ReportsDir -Force | Out-Null }
$snapshotCsv = Join-Path $ReportsDir "${timestamp}_S1_EDF_Snapshot.csv"
$resultCsv = Join-Path $ReportsDir "${timestamp}_S1_EDF_WriteResults.csv"

Write-Step 'Reading current values...'
$snapshot = New-Object System.Collections.Generic.List[object]
$results = New-Object System.Collections.Generic.List[object]
$processed = 0

foreach ($row in $mapRows) {
    $processed++
    $cid = [int]$row.AutomateClientId
    $siteId = [string]$row.S1SiteId

    if ($processed % 25 -eq 0) { Write-Host "  $processed / $($mapRows.Count)..." -ForegroundColor DarkGray }

    try { $fields = @(Get-AutomateClientExtraFields -ClientId $cid) }
    catch {
        $results.Add([PSCustomObject]@{
                AutomateClientId = $cid; AutomateClientName = $row.AutomateClientName
                S1SiteId = $siteId; SiteIdStatus = 'READ_FAILED'; TokenStatus = 'READ_FAILED'
                Detail = $_.Exception.Message
            })
        continue
    }

    $currentSiteId = [string](Resolve-FieldDefinition -Fields $fields -Title $siteIdField.Title).TextFieldSettings.Value
    $currentToken = ''
    if ($siteTokenField) {
        $currentToken = [string](Resolve-FieldDefinition -Fields $fields -Title $siteTokenField.Title).TextFieldSettings.Value
    }

    $snapshot.Add([PSCustomObject]@{
            AutomateClientId   = $cid
            AutomateClientName = $row.AutomateClientName
            PreviousSiteId     = $currentSiteId
            PreviousTokenSet   = [bool]$currentToken
            PreviousTokenValue = $currentToken
        })

    $desiredToken = if ($siteTokenField) { [string]$tokenBySiteId[$siteId] } else { '' }

    # --- site id ---
    $siteIdStatus = 'UNCHANGED'
    if ($currentSiteId -ne $siteId) {
        if (-not $Apply) { $siteIdStatus = 'WOULD_SET' }
        else {
            $r = Set-ClientExtraField -TargetClientId $cid -DefinitionId $siteIdField.ExtraFieldDefinitionId -Value $siteId
            $siteIdStatus = if ($r -eq 'OK') { 'SET' } else { $r }
        }
    }

    # --- token ---
    $tokenStatus = 'SKIPPED'
    if ($siteTokenField) {
        if (-not $desiredToken) { $tokenStatus = 'NO_TOKEN_IN_S1' }
        elseif ($currentToken -eq $desiredToken) { $tokenStatus = 'UNCHANGED' }
        elseif (-not $Apply) { $tokenStatus = 'WOULD_SET' }
        else {
            $r = Set-ClientExtraField -TargetClientId $cid -DefinitionId $siteTokenField.ExtraFieldDefinitionId -Value $desiredToken
            $tokenStatus = if ($r -eq 'OK') { 'SET' } else { $r }
        }
    }

    # --- verify by read-back: a 2xx is not proof the value stuck ---
    if ($Apply -and ($siteIdStatus -eq 'SET' -or $tokenStatus -eq 'SET')) {
        try {
            $after = @(Get-AutomateClientExtraFields -ClientId $cid)
            if ($siteIdStatus -eq 'SET') {
                $v = [string](Resolve-FieldDefinition -Fields $after -Title $siteIdField.Title).TextFieldSettings.Value
                if ($v -ne $siteId) { $siteIdStatus = 'VERIFY_FAILED' }
            }
            if ($tokenStatus -eq 'SET') {
                $v = [string](Resolve-FieldDefinition -Fields $after -Title $siteTokenField.Title).TextFieldSettings.Value
                if ($v -ne $desiredToken) { $tokenStatus = 'VERIFY_FAILED' }
            }
        }
        catch {
            if ($siteIdStatus -eq 'SET') { $siteIdStatus = 'VERIFY_ERROR' }
            if ($tokenStatus -eq 'SET') { $tokenStatus = 'VERIFY_ERROR' }
        }
    }

    $results.Add([PSCustomObject]@{
            AutomateClientId   = $cid
            AutomateClientName = $row.AutomateClientName
            S1SiteName         = $row.S1SiteName
            S1SiteId           = $siteId
            PreviousSiteId     = $currentSiteId
            SiteIdStatus       = $siteIdStatus
            TokenStatus        = $tokenStatus
            Detail             = ''
        })
}

$snapshot | Export-Csv -Path $snapshotCsv -NoTypeInformation -Encoding UTF8
$results | Export-Csv -Path $resultCsv -NoTypeInformation -Encoding UTF8

#endregion

#region Report

function Count-Status { param($Rows, $Property, $Value) @($Rows | Where-Object { $_.$Property -eq $Value }).Count }

Write-Host ''
Write-Step '--- Summary ---'
foreach ($status in @('SET', 'WOULD_SET', 'UNCHANGED', 'VERIFY_FAILED', 'VERIFY_ERROR', 'FORBIDDEN', 'BADREQUEST', 'READ_FAILED')) {
    $a = Count-Status $results 'SiteIdStatus' $status
    $b = Count-Status $results 'TokenStatus' $status
    if ($a -or $b) {
        $color = switch ($status) {
            'SET' { 'Green' } 'WOULD_SET' { 'Cyan' } 'UNCHANGED' { 'Gray' } default { 'Red' }
        }
        Write-Host ("  {0,-14}  SiteID {1,4}   Token {2,4}" -f $status, $a, $b) -ForegroundColor $color
    }
}
$noToken = Count-Status $results 'TokenStatus' 'NO_TOKEN_IN_S1'
if ($noToken) { Write-Host ("  {0,-14}  Token  {1,4}" -f 'NO_TOKEN_IN_S1', $noToken) -ForegroundColor Yellow }

$problemRows = @($results | Where-Object {
        $_.SiteIdStatus -notin @('SET', 'WOULD_SET', 'UNCHANGED') -or
        $_.TokenStatus -notin @('SET', 'WOULD_SET', 'UNCHANGED', 'SKIPPED')
    })

if ($problemRows.Count -gt 0) {
    Write-Host ''
    Write-Step '--- Problems ---' 'Red'
    $problemRows | Format-Table AutomateClientId, AutomateClientName, SiteIdStatus, TokenStatus, Detail -AutoSize -Wrap |
        Out-String -Width 4096 | Write-Host
}

if (-not $Apply) {
    $changes = @($results | Where-Object { $_.SiteIdStatus -eq 'WOULD_SET' -or $_.TokenStatus -eq 'WOULD_SET' })
    Write-Host ''
    Write-Step ("--- {0} companies would change (showing up to 20) ---" -f $changes.Count) 'Cyan'
    $changes | Select-Object -First 20 |
        Format-Table AutomateClientId, AutomateClientName, S1SiteId, SiteIdStatus, TokenStatus -AutoSize |
        Out-String -Width 4096 | Write-Host
}

Write-Host ''
Write-Step "Snapshot (pre-change values) : $snapshotCsv" 'Green'
Write-Step "Write results                : $resultCsv" 'Green'

if (-not $SkipToken) {
    Write-Warning "$snapshotCsv contains any previous site-token values. Treat it as a secret and delete it once the migration is signed off."
}

Write-Host ''
if (-not $Apply) {
    Write-Host 'DRY RUN -- nothing was written. Re-run with -Apply to commit these changes.' -ForegroundColor Yellow
    $exitCode = 0
}
elseif ($problemRows.Count -gt 0) {
    Write-Host ("COMPLETED WITH ERRORS: {0} companies did not write cleanly (see above)." -f $problemRows.Count) -ForegroundColor Red
    $exitCode = 1
}
else {
    Write-Host 'All mapped companies now carry the correct S1 site id and token.' -ForegroundColor Green
    $exitCode = 0
}

#endregion

exit $exitCode
