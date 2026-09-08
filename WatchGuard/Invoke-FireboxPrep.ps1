<#
    Invoke-FireboxPrep.ps1

    Basic prep / onboarding steps against a single LOCALLY-MANAGED WatchGuard Firebox
    over the Fireware CLI (SSH, port 4118), then records the device in IT Glue.

    What it does
      1. Audit      - always. Collects sysinfo, feature key, NTP, log settings, mgmt
                      accounts, global settings and WatchGuard Cloud status. Read-only.
      2. Admin      - creates/updates a standard Device Administrator account, and can
                      disable named leftover accounts.
      3. Hardening  - enables Device Management account lockout, tunes the thresholds,
                      and optionally disables concurrent Device Administrator logins.
      4. Cloud prep - syncs the feature key, enables NTP + NTP servers, sets the
                      timezone. These are the prerequisites for adding the box to
                      WatchGuard Cloud as a locally-managed device.
      5. IT Glue    - creates (or updates) a Configuration from the collected sysinfo:
                      name, serial, model, firmware, IP/MAC, install date.

    NOTHING IS CHANGED unless you pass -Apply. Without it the script connects, audits,
    and prints the exact command list it would run.

    IMPORTANT - what the CLI cannot do
      Fireware exposes 'watchguard-cloud' as a SHOW component only. There is no CLI
      command to tick "Enable WatchGuard Cloud". After this script reports all
      prerequisites green, finish in Fireware Web UI (System > WatchGuard Cloud) or
      Policy Manager (Setup > WatchGuard Cloud) and paste the verification code from
      WatchGuard Cloud > Configure > Devices > Add Device > Local Management.

    Requirements
      - Windows PowerShell 5.1 (not pwsh 7 - Invoke-RestMethod is broken there for us)
      - Posh-SSH module:  Install-Module Posh-SSH -Scope CurrentUser
      - Config\firebox-prep-credentials.xml holding the IT Glue key (see -InitConfig)

    Examples
      # Audit only, no changes:
      .\Invoke-FireboxPrep.ps1 -FireboxHost 10.0.1.1

      # Full prep + write the config into IT Glue org 9044976:
      .\Invoke-FireboxPrep.ps1 -FireboxHost 10.0.1.1 -Apply `
          -NewAdminUser cbtadmin -TimeZone 15 -ITGlueOrganizationId 9044976

    !! Transcript output can contain configuration detail. Do not commit it. !!
#>
[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Run')]
    [Alias('Firebox', 'Target')]
    [string]$FireboxHost,

    [Parameter(ParameterSetName = 'Run')]
    [int]$Port = 4118,

    [Parameter(ParameterSetName = 'Run')]
    [string]$UserName = 'admin',

    [Parameter(ParameterSetName = 'Run')]
    [pscredential]$Credential,

    # --- change gate -------------------------------------------------------
    [Parameter(ParameterSetName = 'Run')]
    [switch]$Apply,

    # --- 2. admin account --------------------------------------------------
    [Parameter(ParameterSetName = 'Run')]
    [string]$NewAdminUser,

    [Parameter(ParameterSetName = 'Run')]
    [securestring]$NewAdminPassword,

    [Parameter(ParameterSetName = 'Run')]
    [string[]]$DisableAccount,

    # --- 3. hardening ------------------------------------------------------
    [Parameter(ParameterSetName = 'Run')]
    [switch]$SkipHardening,

    [Parameter(ParameterSetName = 'Run')]
    [ValidateRange(1, 10)]
    [int]$LockoutAttempts = 5,

    [Parameter(ParameterSetName = 'Run')]
    [ValidateRange(1, 1440)]
    [int]$LockoutDuration = 15,

    [Parameter(ParameterSetName = 'Run')]
    [ValidateRange(0, 10)]
    [int]$LockoutTempLockouts = 3,

    [Parameter(ParameterSetName = 'Run')]
    [switch]$DisableConcurrentAdmins,

    # --- 4. cloud prep -----------------------------------------------------
    [Parameter(ParameterSetName = 'Run')]
    [switch]$SkipCloudPrep,

    [Parameter(ParameterSetName = 'Run')]
    [string[]]$NtpServer = @('time.watchguard.com', 'pool.ntp.org'),

    # 00-74. Run with -ListTimeZones to print the device's list.
    [Parameter(ParameterSetName = 'Run')]
    [ValidateRange(0, 74)]
    [int]$TimeZone = -1,

    [Parameter(ParameterSetName = 'Run')]
    [switch]$ListTimeZones,

    [Parameter(ParameterSetName = 'Run')]
    [switch]$SkipFeatureKeySync,

    # --- 5. IT Glue --------------------------------------------------------
    [Parameter(ParameterSetName = 'Run')]
    [string]$ITGlueOrganizationId,

    [Parameter(ParameterSetName = 'Run')]
    [string]$ITGlueConfigName,

    [Parameter(ParameterSetName = 'Run')]
    [string]$ITGlueStatus = 'Active',

    [Parameter(ParameterSetName = 'Run')]
    [string]$ITGlueConfigType = 'Managed Network Firewall',

    # --- output ------------------------------------------------------------
    [Parameter(ParameterSetName = 'Run')]
    [string]$OutputPath,

    # seed Config\firebox-prep-credentials.xml from the dashboard creds file
    [Parameter(Mandatory = $true, ParameterSetName = 'Init')]
    [switch]$InitConfig
)

$ErrorActionPreference = 'Stop'
$script:StartedAt = Get-Date

# Fireware config type id fallback, used if the type cannot be resolved by name.
$script:ITGlueFirewallTypeIdFallback = 619408

# ===========================================================================
# Config / credentials
# ===========================================================================
$ConfigDir  = Join-Path $PSScriptRoot 'Config'
$ConfigFile = Join-Path $ConfigDir 'firebox-prep-credentials.xml'

if ($InitConfig) {
    # Seed this script's own credential file from the dashboard one. We read the
    # dashboard file but never write to it.
    $seed = Join-Path (Split-Path -Parent $PSScriptRoot) 'IT Glue\Dashboard\Config\dashboard-credentials.xml'
    if (-not (Test-Path $seed)) { throw "Seed credentials not found: $seed" }
    if (-not (Test-Path $ConfigDir)) { New-Item -ItemType Directory -Path $ConfigDir | Out-Null }
    if (Test-Path $ConfigFile) { throw "Already exists, refusing to overwrite: $ConfigFile" }
    $src = Import-Clixml $seed
    @{ ITGlue = $src['ITGlue'] } | Export-Clixml -Path $ConfigFile
    Write-Host "Wrote $ConfigFile (IT Glue key only)." -ForegroundColor Green
    return
}

if ($PSVersionTable.PSVersion.Major -ge 6) {
    Write-Warning 'Run this under Windows PowerShell 5.1. Invoke-RestMethod behaves badly for us under pwsh 7.'
}

# ===========================================================================
# SSH plumbing
# ===========================================================================
if (-not (Get-Module -ListAvailable -Name Posh-SSH)) {
    throw "Posh-SSH is not installed. Run: Install-Module Posh-SSH -Scope CurrentUser"
}
Import-Module Posh-SSH -ErrorAction Stop

if (-not $Credential) {
    $Credential = Get-Credential -UserName $UserName -Message "Fireware Device Administrator on $FireboxHost"
}

$script:Transcript = New-Object System.Collections.ArrayList

function Write-Step { param([string]$Text) Write-Host "`n=== $Text ===" -ForegroundColor Cyan }
function Write-Ok   { param([string]$Text) Write-Host "  [ ok ] $Text" -ForegroundColor Green }
function Write-Warn { param([string]$Text) Write-Host "  [warn] $Text" -ForegroundColor Yellow }
function Write-Info { param([string]$Text) Write-Host "  $Text" -ForegroundColor Gray }

function Read-FbOutput {
    <#
      Read from the Fireware shell stream until the prompt comes back.
      Fireware pages long output and sometimes asks y/n, so handle both.
    #>
    param(
        [Parameter(Mandatory)] $Stream,
        [int]$TimeoutSec = 25,
        [switch]$AnswerYes
    )
    $sb       = New-Object System.Text.StringBuilder
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $lastData = Get-Date

    while ((Get-Date) -lt $deadline) {
        if ($Stream.DataAvailable) {
            $chunk = $Stream.Read()
            if ($chunk) { [void]$sb.Append($chunk); $lastData = Get-Date }

            $tail = $sb.ToString()
            if ($tail.Length -gt 200) { $tail = $tail.Substring($tail.Length - 200) }

            if ($tail -match '(--\s*More\s*--|<--- More --->|\(END\))\s*$') {
                $Stream.Write(' ')
                continue
            }
            if ($tail -match '\(\s*y\s*/\s*n\s*\)|\[\s*y\s*/\s*n\s*\]|\(yes/no\)') {
                $Stream.WriteLine($(if ($AnswerYes) { 'y' } else { 'n' }))
                continue
            }
            continue
        }

        Start-Sleep -Milliseconds 150
        $idleMs = ((Get-Date) - $lastData).TotalMilliseconds
        $txt    = $sb.ToString()

        # prompt is the last line, ending in > (main) or # (config/policy)
        if ($idleMs -gt 400 -and $txt -match '(?m)^[^\r\n]*[>#]\s*$') { break }
        if ($idleMs -gt 6000 -and $txt.Length -gt 0) { break }
    }
    $sb.ToString()
}

function Invoke-FbCommand {
    <#
      Send one CLI command and return its output. Honours -Apply: commands flagged
      -Change are only printed, not sent, unless $Apply is set.
    #>
    param(
        [Parameter(Mandatory)] $Stream,
        [Parameter(Mandatory)] [string]$Command,
        [switch]$Change,
        [int]$TimeoutSec = 25
    )

    if ($Change -and -not $Apply) {
        Write-Host "  [dry ] $Command" -ForegroundColor DarkYellow
        [void]$script:Transcript.Add([pscustomobject]@{ When = Get-Date; Command = $Command; Mode = 'dry-run'; Output = '' })
        return ''
    }

    $Stream.WriteLine($Command)
    Start-Sleep -Milliseconds 200
    $out = Read-FbOutput -Stream $Stream -TimeoutSec $TimeoutSec -AnswerYes:$Change

    # strip the echoed command and the trailing prompt line
    $lines = @($out -split "`r?`n")
    if ($lines.Count -gt 1 -and $lines[0].Trim() -eq $Command.Trim()) { $lines = @($lines[1..($lines.Count - 1)]) }
    if ($lines.Count -gt 1 -and $lines[-1] -match '[>#]\s*$')         { $lines = @($lines[0..($lines.Count - 2)]) }
    elseif ($lines.Count -eq 1 -and $lines[0] -match '[>#]\s*$')      { $lines = @() }
    $clean = ($lines -join "`n").Trim()

    [void]$script:Transcript.Add([pscustomobject]@{
        When = Get-Date; Command = $Command; Mode = $(if ($Change) { 'applied' } else { 'read' }); Output = $clean
    })

    if ($Change) { Write-Host "  [sent] $Command" -ForegroundColor Magenta }
    if ($clean -match '(?im)^\s*(invalid|unknown command|syntax error|%\s*error|command failed)') {
        Write-Warn "Device rejected: $Command"
        Write-Info ($clean -split "`n" | Select-Object -First 2 | Out-String).Trim()
    }
    return $clean
}

function Get-FbField {
    # pull "Label: value" out of CLI output, tolerant of spacing and label variants
    param([string]$Text, [string[]]$Labels)
    foreach ($l in $Labels) {
        $m = [regex]::Match("$Text", "(?im)^\s*$([regex]::Escape($l))\s*[:=]\s*(.+?)\s*$")
        if ($m.Success) { return $m.Groups[1].Value.Trim() }
    }
    return ''
}

# ===========================================================================
# 1. CONNECT + AUDIT
# ===========================================================================
Write-Step "Connecting to $FireboxHost`:$Port as $($Credential.UserName)"

$session = $null
try {
    $session = New-SSHSession -ComputerName $FireboxHost -Port $Port -Credential $Credential -AcceptKey -ConnectionTimeout 20
    $stream  = New-SSHShellStream -SSHSession $session -TerminalName 'vt100' -Columns 200 -Rows 5000 -Width 1600 -Height 8000
    Start-Sleep -Milliseconds 800
    $banner = Read-FbOutput -Stream $stream -TimeoutSec 15
    Write-Ok "Connected. $((($banner -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -First 1))"
} catch {
    throw "SSH connect to $FireboxHost`:$Port failed: $($_.Exception.Message)"
}

$audit = [ordered]@{}
try {
    Write-Step 'Audit (read-only)'

    if ($ListTimeZones) {
        Invoke-FbCommand -Stream $stream -Command 'configure' | Out-Null
        $tzList = Invoke-FbCommand -Stream $stream -Command 'system timezone ?'
        Invoke-FbCommand -Stream $stream -Command 'exit' | Out-Null
        Write-Host $tzList
    }

    $audit.SysInfo    = Invoke-FbCommand -Stream $stream -Command 'sysinfo'
    $audit.FeatureKey = Invoke-FbCommand -Stream $stream -Command 'show feature-key'
    $audit.Ntp        = Invoke-FbCommand -Stream $stream -Command 'show ntp'
    $audit.Cloud      = Invoke-FbCommand -Stream $stream -Command 'show watchguard-cloud'
    $audit.MgmtUsers  = Invoke-FbCommand -Stream $stream -Command 'show device-mgmt-user'
    $audit.Global     = Invoke-FbCommand -Stream $stream -Command 'show global-setting'
    $audit.LogSetting = Invoke-FbCommand -Stream $stream -Command 'show log-setting'
    $audit.Interfaces = Invoke-FbCommand -Stream $stream -Command 'show interface'

    # --- parse the bits we care about -------------------------------------
    $model    = Get-FbField $audit.SysInfo @('Model', 'Device Model', 'Product')
    $serial   = Get-FbField $audit.SysInfo @('Serial Number', 'Serial', 'SerialNumber')
    $firmware = Get-FbField $audit.SysInfo @('Firmware Version', 'Fireware Version', 'Version', 'OS Version')
    $devName  = Get-FbField $audit.SysInfo @('Name', 'Device Name', 'System Name', 'Hostname')
    $uptime   = Get-FbField $audit.SysInfo @('Uptime', 'System Uptime')

    if (-not $serial)   { $m = [regex]::Match($audit.SysInfo, '(?im)serial[^\r\n:]*[:=]\s*(\S+)');   if ($m.Success) { $serial   = $m.Groups[1].Value } }
    if (-not $firmware) { $m = [regex]::Match($audit.SysInfo, '(?im)\b(\d+\.\d+\.\d+(?:\.\d+)?)\b'); if ($m.Success) { $firmware = $m.Groups[1].Value } }

    $ntpEnabled  = $audit.Ntp -match '(?im)^\s*(ntp\s+)?(status|enabled?)\s*[:=]?\s*(enabled|yes|on|true)'
    $cloudOn     = $audit.Cloud -match '(?im)(enabled|connected|registered)' -and $audit.Cloud -notmatch '(?im)not (enabled|connected|registered)'
    $lockoutOn   = $audit.MgmtUsers -match '(?im)lockout[^\r\n]*\b(enabled|on)\b'

    Write-Host ''
    Write-Info ("Device name    : {0}" -f $(if ($devName)  { $devName }  else { '(not parsed)' }))
    Write-Info ("Model          : {0}" -f $(if ($model)    { $model }    else { '(not parsed)' }))
    Write-Info ("Serial         : {0}" -f $(if ($serial)   { $serial }   else { '(not parsed)' }))
    Write-Info ("Firmware       : {0}" -f $(if ($firmware) { $firmware } else { '(not parsed)' }))
    Write-Info ("Uptime         : {0}" -f $(if ($uptime)   { $uptime }   else { '(not parsed)' }))
    Write-Info ("NTP enabled    : {0}" -f $ntpEnabled)
    Write-Info ("Acct lockout   : {0}" -f $lockoutOn)
    Write-Info ("WG Cloud       : {0}" -f $(if ($cloudOn) { 'enabled/connected' } else { 'NOT enabled' }))

    if (-not $serial) {
        Write-Warn 'Serial number not parsed from sysinfo. Check the raw transcript and adjust Get-FbField labels.'
    }

    # =======================================================================
    # 2. ADMIN / API ACCOUNT
    # =======================================================================
    Write-Step 'Admin account setup'
    if ($NewAdminUser) {
        if (-not $NewAdminPassword) {
            $NewAdminPassword = Read-Host -AsSecureString "Passphrase for Device Administrator '$NewAdminUser' (8-32 chars)"
        }
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                     [Runtime.InteropServices.Marshal]::SecureStringToBSTR($NewAdminPassword))
        if ($plain.Length -lt 8 -or $plain.Length -gt 32) {
            throw "Fireware Device Management passphrases must be 8-32 characters (got $($plain.Length))."
        }
        # the passphrase is passed as a bare CLI token, so whitespace would split the command
        if ($plain -match '\s') {
            throw 'Passphrase cannot contain whitespace - it is sent as a single CLI argument.'
        }

        $exists = $audit.MgmtUsers -match "(?im)^\s*$([regex]::Escape($NewAdminUser))\b"
        Write-Info $(if ($exists) { "'$NewAdminUser' already exists - updating passphrase and role." }
                     else         { "'$NewAdminUser' not present - creating." })

        # device-mgmt-user is a MAIN mode command, no 'configure' needed.
        Invoke-FbCommand -Stream $stream -Change `
            -Command "device-mgmt-user $NewAdminUser Firebox-DB password $plain role Device-Administrator" | Out-Null

        $plain = $null
        Write-Ok "Device Administrator '$NewAdminUser' set."
    } else {
        Write-Info 'No -NewAdminUser given, skipping account creation.'
    }

    foreach ($acct in $DisableAccount) {
        if ($acct -eq 'admin') { Write-Warn "Refusing to disable 'admin' - that is your way back in."; continue }
        Invoke-FbCommand -Stream $stream -Change `
            -Command "device-mgmt-user $acct Firebox-DB role Disabled" | Out-Null
        Write-Ok "Account '$acct' set to Disabled."
    }

    # =======================================================================
    # 3. MANAGEMENT ACCESS HARDENING
    # =======================================================================
    if (-not $SkipHardening) {
        Write-Step 'Management access hardening'
        Invoke-FbCommand -Stream $stream -Command 'configure' | Out-Null

        Invoke-FbCommand -Stream $stream -Change -Command 'device-mgmt-user account-lockout enable' | Out-Null
        Invoke-FbCommand -Stream $stream -Change -Command "device-mgmt-user account-lockout attempts $LockoutAttempts" | Out-Null
        Invoke-FbCommand -Stream $stream -Change -Command "device-mgmt-user account-lockout duration $LockoutDuration" | Out-Null
        Invoke-FbCommand -Stream $stream -Change -Command "device-mgmt-user account-lockout lockouts $LockoutTempLockouts" | Out-Null

        if ($DisableConcurrentAdmins) {
            Invoke-FbCommand -Stream $stream -Change -Command 'no global-setting device-admin-connections enable' | Out-Null
        }

        Invoke-FbCommand -Stream $stream -Command 'exit' | Out-Null
        Write-Ok "Account lockout: $LockoutAttempts attempts / ${LockoutDuration}min / $LockoutTempLockouts temp lockouts."

        # Policy-level exposure is REPORTED, never changed - editing the WatchGuard
        # management policy over the same SSH session is how you lock yourself out.
        $wgPolicy = Invoke-FbCommand -Stream $stream -Command 'show policy WatchGuard'
        if (-not $wgPolicy -or $wgPolicy -match '(?im)invalid|unknown') {
            $wgPolicy = Invoke-FbCommand -Stream $stream -Command 'show policy'
        }
        $audit.MgmtPolicy = $wgPolicy
        if ($wgPolicy -match '(?im)\bAny-External\b|\b0\.0\.0\.0/0\b') {
            Write-Warn 'Management policy appears to allow Any-External. Review manually - not changed by this script.'
        } else {
            Write-Ok 'Management policy does not obviously allow Any-External (review the transcript to confirm).'
        }
    }

    # =======================================================================
    # 4. WATCHGUARD CLOUD ONBOARDING PREP
    # =======================================================================
    if (-not $SkipCloudPrep) {
        Write-Step 'WatchGuard Cloud onboarding prep'

        if (-not $SkipFeatureKeySync) {
            # WG Cloud requires the current feature key to be synced on the device.
            Invoke-FbCommand -Stream $stream -Change -Command 'sync feature-key apply' -TimeoutSec 60 | Out-Null
            Write-Ok 'Feature key synced from LiveSecurity (apply = effective immediately).'
        }

        Invoke-FbCommand -Stream $stream -Command 'configure' | Out-Null

        Invoke-FbCommand -Stream $stream -Change -Command 'ntp enable' | Out-Null
        foreach ($srv in $NtpServer) {
            if ($srv -match '^\d{1,3}(\.\d{1,3}){3}$') {
                Invoke-FbCommand -Stream $stream -Change -Command "ntp server ip $srv" | Out-Null
            } else {
                Invoke-FbCommand -Stream $stream -Change -Command "ntp server domain $srv" | Out-Null
            }
        }
        Write-Ok ("NTP enabled with: {0}" -f ($NtpServer -join ', '))

        if ($TimeZone -ge 0) {
            Invoke-FbCommand -Stream $stream -Change -Command "system timezone $TimeZone" | Out-Null
            Write-Ok "Timezone set to index $TimeZone."
        } else {
            Write-Info 'No -TimeZone given, leaving timezone alone. Use -ListTimeZones to see the device list.'
        }

        Invoke-FbCommand -Stream $stream -Command 'exit' | Out-Null

        # re-read post-change state
        $audit.NtpAfter   = Invoke-FbCommand -Stream $stream -Command 'show ntp'
        $audit.CloudAfter = Invoke-FbCommand -Stream $stream -Command 'show watchguard-cloud'
    }

} finally {
    if ($session) {
        try { $stream.WriteLine('exit') } catch { }
        Remove-SSHSession -SSHSession $session -ErrorAction SilentlyContinue | Out-Null
    }
}

# ===========================================================================
# 5. IT GLUE CONFIGURATION
# ===========================================================================
$igResult = $null
if ($ITGlueOrganizationId) {
    Write-Step "IT Glue configuration (org $ITGlueOrganizationId)"

    if (-not (Test-Path $ConfigFile)) {
        throw "Missing $ConfigFile. Run: .\Invoke-FireboxPrep.ps1 -InitConfig"
    }
    $cfg  = Import-Clixml $ConfigFile
    $ig   = $cfg['ITGlue']
    $base = $ig.BaseUri
    $H    = @{ 'x-api-key' = $ig.ApiKey; 'Content-Type' = 'application/vnd.api+json' }

    function Invoke-IG {
        param([string]$Method, [string]$Path, $Body)
        $uri = "$base$Path"
        if ($Body) {
            Invoke-RestMethod -Uri $uri -Method $Method -Headers $H -Body ($Body | ConvertTo-Json -Depth 10)
        } else {
            Invoke-RestMethod -Uri $uri -Method $Method -Headers $H
        }
    }

    # -- resolve type / status / manufacturer / model ----------------------
    $typeId = $null
    try {
        $t = Invoke-IG GET "/configuration_types?filter[name]=$([uri]::EscapeDataString($ITGlueConfigType))"
        if ($t.data.Count -gt 0) { $typeId = $t.data[0].id }
    } catch { }
    if (-not $typeId) {
        $typeId = $script:ITGlueFirewallTypeIdFallback
        Write-Warn "Config type '$ITGlueConfigType' not resolved by name, using fallback id $typeId."
    }

    $statusId = $null
    try {
        $s = Invoke-IG GET "/configuration_statuses?filter[name]=$([uri]::EscapeDataString($ITGlueStatus))"
        if ($s.data.Count -gt 0) { $statusId = $s.data[0].id }
    } catch { }
    if (-not $statusId) { Write-Warn "Status '$ITGlueStatus' not resolved - config will be created without a status." }

    $manuId = $null; $modelId = $null
    try {
        $mf = Invoke-IG GET "/manufacturers?filter[name]=WatchGuard"
        if ($mf.data.Count -gt 0) {
            $manuId = $mf.data[0].id
            if ($model) {
                $md = Invoke-IG GET "/models?filter[manufacturer_id]=$manuId&filter[name]=$([uri]::EscapeDataString($model))"
                if ($md.data.Count -gt 0) { $modelId = $md.data[0].id }
            }
        }
    } catch { }
    if (-not $manuId)  { Write-Warn 'Manufacturer "WatchGuard" not found in IT Glue - leaving blank.' }
    if ($model -and -not $modelId) { Write-Warn "Model '$model' not found under WatchGuard - recorded in notes instead." }

    # -- does a config with this serial already exist in the org? ----------
    $existing = $null
    if ($serial) {
        try {
            $q = Invoke-IG GET "/organizations/$ITGlueOrganizationId/relationships/configurations?filter[serial_number]=$([uri]::EscapeDataString($serial))"
            if ($q.data.Count -gt 0) { $existing = $q.data[0] }
        } catch { }
    }

    $cfgName = if ($ITGlueConfigName) { $ITGlueConfigName }
               elseif ($devName)      { $devName }
               elseif ($model)        { "WatchGuard $model" }
               else                   { "WatchGuard Firebox ($FireboxHost)" }

    $notes = @(
        "Prepped by Invoke-FireboxPrep.ps1 on $(Get-Date -Format 'yyyy-MM-dd HH:mm')."
        "Management: SSH $FireboxHost`:$Port (Fireware CLI), locally managed."
        $(if ($model)    { "Model: $model" })
        $(if ($firmware) { "Fireware: $firmware" })
        $(if ($uptime)   { "Uptime at audit: $uptime" })
        "WatchGuard Cloud: $(if ($cloudOn) { 'enabled' } else { 'NOT enabled - finish in Web UI System > WatchGuard Cloud' })"
    ) | Where-Object { $_ }

    $attrs = @{
        'organization-id'        = $ITGlueOrganizationId
        'configuration-type-id'  = $typeId
        'name'                   = $cfgName
        'hostname'               = $(if ($devName) { $devName } else { $cfgName })
        'primary-ip'             = $FireboxHost
        'serial-number'          = $serial
        'operating-system-notes' = $(if ($firmware) { "Fireware $firmware" } else { '' })
        'notes'                  = ($notes -join "`n")
        'installed-at'           = (Get-Date -Format 'yyyy-MM-dd')
    }
    if ($statusId) { $attrs['configuration-status-id'] = $statusId }
    if ($manuId)   { $attrs['manufacturer-id']         = $manuId }
    if ($modelId)  { $attrs['model-id']                = $modelId }
    # drop empties so we never blank an existing field on update
    $clean = @{}
    foreach ($kv in $attrs.GetEnumerator()) {
        if ("$($kv.Value)".Trim() -ne '') { $clean[$kv.Key] = $kv.Value }
    }
    $attrs = $clean

    if (-not $Apply) {
        Write-Host '  [dry ] Would write this IT Glue configuration:' -ForegroundColor DarkYellow
        $attrs.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Info ("  {0,-24} {1}" -f $_.Key, $_.Value) }
        Write-Info $(if ($existing) { "  (would UPDATE existing config id $($existing.id))" } else { '  (would CREATE a new config)' })
    } else {
        $body = @{ data = @{ type = 'configurations'; attributes = $attrs } }
        if ($existing) {
            $id = $existing.id
            # IT Glue can return 422 on a partial PATCH yet still apply it, so we
            # verify by read-back rather than trusting the response.
            try   { Invoke-IG PATCH "/configurations/$id" $body | Out-Null }
            catch { Write-Warn "PATCH returned $($_.Exception.Message) - verifying by read-back." }
            $check = Invoke-IG GET "/configurations/$id"
            $igResult = $check.data
            if ($check.data.attributes.'serial-number' -eq $serial) { Write-Ok "Updated IT Glue configuration $id ($cfgName)." }
            else { Write-Warn "Read-back did not match. Check configuration $id manually." }
        } else {
            $new = Invoke-IG POST "/configurations" $body
            $igResult = $new.data
            Write-Ok "Created IT Glue configuration $($new.data.id) ($cfgName)."
        }
    }
} else {
    Write-Info "`nNo -ITGlueOrganizationId given, skipping IT Glue."
}

# ===========================================================================
# Summary + punchlist
# ===========================================================================
Write-Step 'Summary'
Write-Host ("  Mode           : {0}" -f $(if ($Apply) { 'APPLY (changes were sent)' } else { 'DRY RUN (nothing changed)' })) `
    -ForegroundColor $(if ($Apply) { 'Magenta' } else { 'Yellow' })
Write-Info ("  Firebox        : {0} ({1} {2})" -f $FireboxHost, $model, $firmware)
Write-Info ("  Serial         : {0}" -f $serial)
Write-Info ("  Commands sent  : {0}" -f @($script:Transcript | Where-Object Mode -eq 'applied').Count)
Write-Info ("  Elapsed        : {0:n1}s" -f ((Get-Date) - $script:StartedAt).TotalSeconds)

Write-Host "`n  Remaining manual steps:" -ForegroundColor Cyan
Write-Host "   1. WatchGuard Cloud > Configure > Devices > Add Device > Local Management." -ForegroundColor Gray
Write-Host "      Copy the verification code (valid 30 days)." -ForegroundColor Gray
Write-Host "   2. Fireware Web UI > System > WatchGuard Cloud > Enable WatchGuard Cloud," -ForegroundColor Gray
Write-Host "      paste the code, Save. (No CLI equivalent exists - 'watchguard-cloud'" -ForegroundColor Gray
Write-Host "      is a show-only component in Fireware.)" -ForegroundColor Gray
Write-Host "   3. Confirm the box reports in under Monitor > Devices." -ForegroundColor Gray

if ($OutputPath) {
    $dir = Split-Path -Parent $OutputPath
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [pscustomobject]@{
        RanAt        = $script:StartedAt
        FireboxHost  = $FireboxHost
        Applied      = [bool]$Apply
        DeviceName   = $devName
        Model        = $model
        Serial       = $serial
        Firmware     = $firmware
        CloudEnabled = $cloudOn
        Audit        = $audit
        Transcript   = $script:Transcript
        ITGlue       = $igResult
    } | ConvertTo-Json -Depth 8 | Set-Content -Path $OutputPath -Encoding UTF8
    Write-Host "`n  Transcript: $OutputPath" -ForegroundColor Cyan
    Write-Host "  Contains device configuration detail - do not commit." -ForegroundColor Red
}
