<#
.SYNOPSIS
    Installs the SentinelOne agent into the correct site using a site scope token.

.DESCRIPTION
    Payload for the ConnectWise Automate script "SentinelOne Deploy"
    (Scripts > Antivirus > SentinelOne).

    The Automate script reads the CLIENT-level EDF "S1ScopeToken" and passes it here as
    -ScopeToken. That token is the SentinelOne site's registrationToken, which encodes
    both the console URL and the site key, so it alone determines which site the agent
    enrolls into. No API credentials are needed on the endpoint.

    The EDFs are populated by Set-AutomateS1SiteFields.ps1, which only writes companies
    that Compare-S1SiteToAutomateClient.ps1 verified as a clean 1:1 Automate-to-S1 match.
    That chain is what makes it safe to run this unattended: a client with no verified
    site has an empty EDF, and this script refuses to run rather than guessing.

    Designed to be safely re-runnable. If the agent is already installed it reports and
    exits 0 without touching anything, so it can be attached to a monitor or run against
    a whole client repeatedly.

.PARAMETER ScopeToken
    The site scope token (S1 registrationToken). From Automate, pass the client EDF.

.PARAMETER InstallerPath
    Full path to the SentinelOne installer already on the endpoint. The Automate script
    should download it from the LTShare first -- letting the agent (running as SYSTEM)
    pull from a UNC path is unreliable, so download-then-execute is the correct split.

.PARAMETER ExpectedConsoleUrl
    If supplied, the token's embedded console URL must match this or the script aborts.
    Guards against a stale or wrong-tenant token being pushed to endpoints.

.PARAMETER DetectOnly
    Report what would happen -- detection, token validation, installer check -- and make
    no changes. Use this for the first run against a pilot machine.

.PARAMETER LogPath
    Log file. Defaults to $env:windir\Temp\SentinelOne-Deploy.log.

.PARAMETER TimeoutSeconds
    How long to wait for the agent service to come up after install. Default 300.

.OUTPUTS
    Writes a single RESULT= line as the last line of output for Automate to parse.

    Exit codes:
      0 = success (installed now, or already present)
      1 = missing or malformed scope token
      2 = installer not found
      3 = installation failed
      4 = installed but the service did not start in time
      5 = unsupported platform

.NOTES
    The scope token is a secret. It is never written to the log in full -- only a short
    fingerprint -- because Automate script logs are widely readable.

.EXAMPLE
    .\Install-SentinelOneAgent.ps1 -ScopeToken '<token>' -DetectOnly

    Validates the token and reports install state without changing anything.

.EXAMPLE
    .\Install-SentinelOneAgent.ps1 -ScopeToken '<token>' -InstallerPath 'C:\Windows\Temp\SentinelInstaller.exe'
#>

[CmdletBinding()]
param(
    [string]$ScopeToken,

    [string]$ScopeTokenFile,

    [string]$InstallerPath,

    [string]$ExpectedConsoleUrl,

    [switch]$DetectOnly,

    [string]$LogPath = "$env:windir\Temp\SentinelOne-Deploy.log",

    [int]$TimeoutSeconds = 300
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Automate parses the last line; keep a single machine-readable result.
$script:Result = 'UNKNOWN'

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    try { Add-Content -Path $LogPath -Value $line -ErrorAction SilentlyContinue } catch {}
}

function Exit-With {
    param([int]$Code, [string]$ResultText, [string]$Message, [string]$Level = 'INFO')
    Write-Log $Message $Level
    Write-Host "RESULT=$ResultText"
    exit $Code
}

function Get-TokenFingerprint {
    # Enough to correlate two tokens in a log without disclosing either.
    param([string]$Token)
    try {
        $sha = [Security.Cryptography.SHA256]::Create()
        $hash = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Token))
        return ([BitConverter]::ToString($hash) -replace '-').Substring(0, 12)
    }
    catch { return 'unavailable' }
}

function Test-SentinelInstalled {
    <#
        Multiple independent signals: a partial or broken install can leave the service
        registered but the product directory gone, or vice versa. Report all of them so
        a half-installed machine is visible rather than silently "already installed".
    #>
    $services = @(Get-Service -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -in @('SentinelAgent', 'SentinelStaticEngine', 'SentinelHelperService', 'LogProcessorService') })

    # Verified against a live 24.1.x agent: there is NO HKLM\SOFTWARE\SentinelOne key.
    # The product registers only a standard Uninstall entry named 'Sentinel Agent'.
    $uninstall = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    ) | Where-Object { Test-Path $_ } | ForEach-Object {
        Get-ChildItem $_ -ErrorAction SilentlyContinue | ForEach-Object {
            Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
        }
    } | Where-Object { $_.DisplayName -match '^Sentinel Agent' } | Select-Object -First 1

    $ctl = $null
    $root = Join-Path $env:ProgramFiles 'SentinelOne'
    if (Test-Path $root) {
        $ctl = Get-ChildItem -Path $root -Filter 'sentinelctl.exe' -Recurse -ErrorAction SilentlyContinue |
            Select-Object -First 1
    }

    $version = $null
    if ($uninstall -and $uninstall.DisplayVersion) { $version = [string]$uninstall.DisplayVersion }
    elseif ($ctl -and $ctl.DirectoryName -match 'Sentinel Agent\s+([\d\.]+)') { $version = $Matches[1] }

    # Signals are weighted, not merely counted. Services and sentinelctl are HARD
    # evidence that the product is really present. A leftover Add/Remove Programs entry
    # is not -- S1 uninstalls routinely leave one behind, and treating that as "installed"
    # would block a legitimate deploy on a machine a tech has just cleaned off.
    $hardEvidence = ($services.Count -gt 0 -or $null -ne $ctl)

    [PSCustomObject]@{
        ServicesFound   = @($services | ForEach-Object { $_.Name })
        AgentService    = ($services | Where-Object { $_.Name -eq 'SentinelAgent' } | Select-Object -First 1)
        UninstallName   = $(if ($uninstall) { $uninstall.DisplayName } else { $null })
        SentinelCtl     = $(if ($ctl) { $ctl.FullName } else { $null })
        Version         = $version
        IsInstalled     = $hardEvidence
        LooksComplete   = ($services.Count -gt 0 -and $null -ne $ctl)
        UninstallRemnant = ($null -ne $uninstall -and -not $hardEvidence)
    }
}

function Resolve-ScopeToken {
    <#
        A site token is base64 of {"url":"https://...","site_key":"..."}. Decoding it
        validates the token locally before any install is attempted, and lets us confirm
        it points at the expected console -- catching a stale or wrong-tenant EDF value
        before it enrolls an endpoint into someone else's site.
    #>
    param([string]$Token)

    $clean = ($Token -replace '\s', '')
    if ([string]::IsNullOrWhiteSpace($clean)) { return $null }

    try {
        $pad = $clean.Length % 4
        if ($pad -eq 2) { $clean += '==' } elseif ($pad -eq 3) { $clean += '=' } elseif ($pad -eq 1) { return $null }
        $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($clean))
        $obj = $json | ConvertFrom-Json
    }
    catch { return $null }

    if (-not $obj.url -or -not $obj.site_key) { return $null }

    [PSCustomObject]@{
        Url         = [string]$obj.url
        SiteKeyLen  = ([string]$obj.site_key).Length
        Fingerprint = Get-TokenFingerprint $Token
    }
}

# ------------------------------------------------------------------ main

Write-Log '=============================================================='
Write-Log ("SentinelOne Deploy starting on {0} (DetectOnly={1})" -f $env:COMPUTERNAME, [bool]$DetectOnly)

if ([Environment]::OSVersion.Platform -ne 'Win32NT') {
    Exit-With 5 'UNSUPPORTED_PLATFORM' 'Not a Windows platform.' 'ERROR'
}

# Prefer the file form. A token passed as an argument is visible in the process command
# line to any local user and is captured by EDR telemetry and the Automate command log;
# the file is read once and shredded immediately.
if ($ScopeTokenFile) {
    if (-not (Test-Path $ScopeTokenFile)) {
        Exit-With 1 'BAD_TOKEN' "ScopeTokenFile '$ScopeTokenFile' not found." 'ERROR'
    }
    try {
        $ScopeToken = (Get-Content -Path $ScopeTokenFile -Raw).Trim()
        Write-Log 'Scope token read from file.'
    }
    catch {
        Exit-With 1 'BAD_TOKEN' "Could not read ScopeTokenFile: $($_.Exception.Message)" 'ERROR'
    }
    finally {
        try {
            # Overwrite before delete so the value is not recoverable from free space.
            Set-Content -Path $ScopeTokenFile -Value (' ' * 512) -ErrorAction SilentlyContinue
            Remove-Item -Path $ScopeTokenFile -Force -ErrorAction SilentlyContinue
        }
        catch {}
    }
}

if ([string]::IsNullOrWhiteSpace($ScopeToken)) {
    Exit-With 1 'BAD_TOKEN' 'No scope token supplied. Pass -ScopeToken or -ScopeTokenFile.' 'ERROR'
}

# --- 1. validate the token before anything else ---
$tokenInfo = Resolve-ScopeToken -Token $ScopeToken
if (-not $tokenInfo) {
    Exit-With 1 'BAD_TOKEN' @'
Scope token is missing or malformed. Expected the SentinelOne site registrationToken
(base64 of {"url":...,"site_key":...}). Check the client EDF S1ScopeToken in Automate --
an empty value means this client was never matched to an S1 site by the reconciliation.
'@ 'ERROR'
}
Write-Log ("Token OK -- console {0}, site_key {1} chars, fingerprint {2}" -f $tokenInfo.Url, $tokenInfo.SiteKeyLen, $tokenInfo.Fingerprint)

if ($ExpectedConsoleUrl) {
    $a = $tokenInfo.Url.TrimEnd('/').ToLowerInvariant()
    $b = $ExpectedConsoleUrl.TrimEnd('/').ToLowerInvariant()
    if ($a -ne $b) {
        Exit-With 1 'WRONG_CONSOLE' ("Token points at '$($tokenInfo.Url)' but '$ExpectedConsoleUrl' was required. Refusing to enroll.") 'ERROR'
    }
    Write-Log 'Console URL matches the expected tenant.'
}

# --- 2. already installed? ---
$state = Test-SentinelInstalled
if ($state.IsInstalled) {
    Write-Log ("Detected: services=[{0}] version={1} sentinelctl={2}" -f
        ($state.ServicesFound -join ','), $state.Version, $state.SentinelCtl)

    if ($state.LooksComplete) {
        $svc = $state.AgentService
        if ($svc -and $svc.Status -ne 'Running') {
            Write-Log ("SentinelAgent present but status is '{0}'." -f $svc.Status) 'WARN'
            Exit-With 0 'ALREADY_INSTALLED_NOT_RUNNING' 'Agent installed but not running; investigate rather than reinstalling.' 'WARN'
        }
        Exit-With 0 'ALREADY_INSTALLED' 'SentinelOne agent already installed and running. Nothing to do.'
    }

    # Partial install: reinstalling over this usually fails, so surface it instead.
    Exit-With 3 'PARTIAL_INSTALL' 'SentinelOne is partially installed (some artifacts present). Clean up manually before deploying.' 'ERROR'
}
if ($state.UninstallRemnant) {
    # Expected after a manual uninstall, which is how state-B machines are prepared.
    # The unified .exe installs cleanly over this, so warn and continue.
    Write-Log ("Leftover uninstall entry '{0}' but no services or binaries -- treating as not installed." -f $state.UninstallName) 'WARN'
}
Write-Log 'SentinelOne not detected.'

# --- 3. installer ---
if (-not $InstallerPath) {
    if ($DetectOnly) {
        Write-Log 'No -InstallerPath supplied; skipping installer check in detect-only mode.' 'WARN'
        Exit-With 0 'WOULD_INSTALL' 'Detect-only: token valid, agent absent. Would install if an installer were supplied.'
    }
    Exit-With 2 'NO_INSTALLER' 'No -InstallerPath supplied. The Automate script must download the installer to the endpoint first.' 'ERROR'
}

if (-not (Test-Path $InstallerPath)) {
    Exit-With 2 'NO_INSTALLER' "Installer not found at '$InstallerPath'. Check the Automate File Download step." 'ERROR'
}

$installer = Get-Item $InstallerPath
Write-Log ("Installer: {0} ({1:N1} MB)" -f $installer.FullName, ($installer.Length / 1MB))

if ($DetectOnly) {
    Exit-With 0 'WOULD_INSTALL' 'Detect-only: token valid, agent absent, installer present. Would install now.'
}

# --- 4. install ---
# The unified .exe takes the token directly; the .msi takes it as a property.
$isMsi = $installer.Extension -ieq '.msi'
if ($isMsi) {
    $exe = "$env:windir\System32\msiexec.exe"
    $argList = @('/i', "`"$($installer.FullName)`"", '/qn', '/norestart', "SITE_TOKEN=$ScopeToken")
}
else {
    $exe = $installer.FullName
    $argList = @('-t', $ScopeToken, '-q')
}

Write-Log ("Launching installer ({0})..." -f $(if ($isMsi) { 'msi' } else { 'exe' }))
try {
    $proc = Start-Process -FilePath $exe -ArgumentList $argList -Wait -PassThru -NoNewWindow
    $code = $proc.ExitCode
}
catch {
    Exit-With 3 'INSTALL_FAILED' "Failed to launch installer: $($_.Exception.Message)" 'ERROR'
}

# 3010 = success, reboot required. S1 does not normally need one, but treat it as success.
if ($code -ne 0 -and $code -ne 3010) {
    Exit-With 3 'INSTALL_FAILED' "Installer returned exit code $code." 'ERROR'
}
Write-Log ("Installer completed with exit code {0}." -f $code)

# --- 5. verify the service actually comes up ---
Write-Log "Waiting up to $TimeoutSeconds s for SentinelAgent..."
$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
$running = $false
while ((Get-Date) -lt $deadline) {
    $svc = Get-Service -Name 'SentinelAgent' -ErrorAction SilentlyContinue
    if ($svc -and $svc.Status -eq 'Running') { $running = $true; break }
    Start-Sleep -Seconds 5
}

if (-not $running) {
    Exit-With 4 'SERVICE_NOT_RUNNING' 'Install reported success but SentinelAgent did not reach Running. Check the endpoint.' 'ERROR'
}

$final = Test-SentinelInstalled
Write-Log ("Verified: version={0} services=[{1}]" -f $final.Version, ($final.ServicesFound -join ','))
Exit-With 0 'INSTALLED' 'SentinelOne agent installed and running.'
