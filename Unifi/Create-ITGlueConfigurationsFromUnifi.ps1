#Requires -Version 5.1
<#
.SYNOPSIS
    Creates or updates IT Glue Configurations from Unifi network devices.

.DESCRIPTION
    Connects to a Unifi Network Controller, lets you select a site, resolves the
    matching IT Glue organization, then creates or updates IT Glue Configurations
    for all switches, gateways, and access points on that site. Port connectivity
    tables (built from LLDP neighbor data) are written to each configuration's
    Notes field.

    Credentials are stored securely with DPAPI encryption in Config\Credentials.xml.
    Site-to-org mappings are saved so you only select once per site.

.PARAMETER Reset
    Clears all stored credentials and prompts for new ones.

.EXAMPLE
    .\Create-ITGlueConfigurationsFromUnifi.ps1

.EXAMPLE
    .\Create-ITGlueConfigurationsFromUnifi.ps1 -Reset
#>
[CmdletBinding()]
param(
    [switch]$Reset
)

$ErrorActionPreference = 'Stop'

$ScriptDir       = $PSScriptRoot
$CredentialsPath = Join-Path $ScriptDir 'Config\Credentials.xml'

# ---------------------------------------------------------------------------
# TLS / Certificate handling (works on PS 5.1 and PS 7+)
# ---------------------------------------------------------------------------
if ($PSVersionTable.PSVersion.Major -ge 7) {
    # -SkipCertificateCheck is used per-call below
    $script:SkipCertParam = @{ SkipCertificateCheck = $true }
} else {
    $script:SkipCertParam = @{}
    # Bypass self-signed cert validation globally for PS 5.1
    if (-not ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type) {
        Add-Type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate cert,
        WebRequest request, int certificateProblem) { return true; }
}
"@
    }
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}

# ---------------------------------------------------------------------------
# Module setup
# ---------------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name ITGlueAPI)) {
    Write-Host "ITGlueAPI module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name ITGlueAPI -Scope CurrentUser -Force -AllowClobber
        Write-Host "ITGlueAPI installed successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to install ITGlueAPI: $($_.Exception.Message)"
    }
}
Import-Module ITGlueAPI -Force -ErrorAction Stop

# ---------------------------------------------------------------------------
# Helper: display banner
# ---------------------------------------------------------------------------
function Write-Banner {
    param([string]$Text)
    Write-Host ""
    Write-Host ("=" * 65) -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host ("=" * 65) -ForegroundColor Cyan
    Write-Host ""
}

# ---------------------------------------------------------------------------
# Credential management
# ---------------------------------------------------------------------------
function Initialize-Credentials {
    param([switch]$Reset)

    $configDir = Split-Path $CredentialsPath -Parent
    if (-not (Test-Path $configDir)) {
        New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    }

    $creds = $null
    if ((Test-Path $CredentialsPath) -and (-not $Reset)) {
        try {
            $creds = Import-Clixml -Path $CredentialsPath
        }
        catch {
            Write-Warning "Failed to load stored credentials; will prompt for new ones."
        }
    }

    if (-not $creds) { $creds = @{} }
    $needsSave = $false

    # --- Unifi ---
    if (-not $creds.Unifi) {
        Write-Banner "Unifi Cloud API Credentials"
        Write-Host "Generate an API key at: unifi.ui.com -> Settings -> API Keys" -ForegroundColor Yellow
        Write-Host "The API calls go to api.ui.com (cloud) - no local controller URL needed." -ForegroundColor Yellow
        Write-Host ""
        $apiKeyInput = Read-Host "Unifi API Key"
        $creds.Unifi = @{
            ApiKey = ($apiKeyInput | ConvertTo-SecureString -AsPlainText -Force)
        }
        $needsSave = $true
    }

    # --- ITGlue ---
    if (-not $creds.ITGlue) {
        Write-Banner "IT Glue Credentials"
        $defaultItgUri = 'https://api.itglue.com'
        $itgUri = Read-Host "ITGlue API Base URL [$defaultItgUri]"
        if ([string]::IsNullOrWhiteSpace($itgUri)) { $itgUri = $defaultItgUri }
        $itgKey = Read-Host "ITGlue API Key"
        $creds.ITGlue = @{
            BaseUri = $itgUri
            ApiKey  = $itgKey
        }
        $needsSave = $true
    }

    # --- SiteMappings (initialise if absent) ---
    if (-not $creds.SiteMappings) {
        $creds.SiteMappings = @{}
        $needsSave = $true
    }

    if ($needsSave) {
        try {
            $creds | Export-Clixml -Path $CredentialsPath -Force
            Write-Host "Credentials saved to $CredentialsPath" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to save credentials: $($_.Exception.Message)"
        }
    }

    return $creds
}

function Save-Credentials {
    param($Config)
    try {
        $Config | Export-Clixml -Path $CredentialsPath -Force
    }
    catch {
        Write-Warning "Could not save credentials: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Unifi Cloud API functions (api.ui.com - Official Network API)
# ---------------------------------------------------------------------------
$script:UnifiApiBase = 'https://api.ui.com'
$script:UnifiHeaders = $null

function Initialize-UnifiApi {
    param($Config)

    $plainKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Config.Unifi.ApiKey)
    )
    $script:UnifiHeaders = @{
        'X-API-KEY' = $plainKey
        'Accept'    = 'application/json'
    }

    Write-Host "Connecting to Unifi cloud API (api.ui.com) ..." -ForegroundColor Yellow
    try {
        $test = Invoke-RestMethod -Uri "$($script:UnifiApiBase)/v1/sites" -Headers $script:UnifiHeaders
        $count = if ($test.data) { $test.data.Count } else { 0 }
        Write-Host "Connected. Found $count site(s)." -ForegroundColor Green
    }
    catch {
        $sc = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { '?' }
        $errBody = ''
        if ($_.Exception.Response) {
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                $errBody = (New-Object System.IO.StreamReader($stream)).ReadToEnd()
            } catch { }
        }
        Write-Host "  [$sc] $($_.Exception.Message)" -ForegroundColor Red
        if ($errBody) { Write-Host "  Response: $errBody" -ForegroundColor DarkYellow }
        Write-Error "Could not connect to Unifi cloud API. Verify the API key at unifi.ui.com -> Settings -> API Keys. Run with -Reset to re-enter."
    }
}

function Disconnect-UnifiController {
    # Cloud API is stateless - nothing to disconnect
}

function Invoke-UnifiApi {
    param(
        [string]$Path,
        [string]$Method = 'GET',
        [object]$Body   = $null
    )
    $uri = "$($script:UnifiApiBase)$Path"
    $params = @{
        Uri     = $uri
        Method  = $Method
        Headers = $script:UnifiHeaders
    }

    if ($Body) {
        $params.Body        = ($Body | ConvertTo-Json -Depth 10)
        $params.ContentType = 'application/json'
    }

    $response = Invoke-RestMethod @params
    # Cloud API wraps results in .data
    if ($response.PSObject.Properties.Name -contains 'data') {
        return $response.data
    }
    return $response
}

function Get-UnifiSites {
    # Fetch sites and hosts, then join so each site has a human-readable label
    $sites = Invoke-UnifiApi -Path '/v1/sites'
    if (-not $sites) { return @() }

    # Build host lookup: hostId -> host name
    $hostMap = @{}
    try {
        $hosts = Invoke-UnifiApi -Path '/v1/hosts'
        if ($hosts) {
            foreach ($h in $hosts) {
                $hName = if ($h.reportedState -and $h.reportedState.hostname) { $h.reportedState.hostname }
                         elseif ($h.hostname)     { $h.hostname }
                         elseif ($h.name)         { $h.name }
                         elseif ($h.hardwareId)   { $h.hardwareId }
                         else                     { $h.id }
                $hostMap[$h.id] = $hName
            }
        }
    } catch { }

    # Annotate each site with a display label and sort alphabetically
    foreach ($s in $sites) {
        $siteName = if ($s.meta -and $s.meta.name) { $s.meta.name }
                    elseif ($s.name)                { $s.name }
                    else                            { 'Default' }
        $hostName = if ($s.hostId -and $hostMap[$s.hostId]) { $hostMap[$s.hostId] } else { '' }
        $label = if ($hostName -and $hostName -ne $siteName) { "$hostName - $siteName" } else { $siteName }
        $s | Add-Member -NotePropertyName '_displayName' -NotePropertyValue $label -Force
    }

    return $sites | Sort-Object { $_._displayName }
}

function Get-UnifiDevices {
    param([string]$SiteId)

    $allDevices = Invoke-UnifiApi -Path "/v1/sites/$SiteId/devices"

    # Cloud API device type field may be 'productType', 'type', or inferred from model
    # Filter to switches, gateways, APs
    return $allDevices | Where-Object {
        $t = $_.type
        if (-not $t) { $t = $_.productType }
        if (-not $t) { $t = '' }
        $t = $t.ToLower()
        $t -match 'switch|gateway|firewall|router|ap|access.?point|wireless|usw|ugw|udm|uap|usg|ucg'
    }
}

# ---------------------------------------------------------------------------
# IT Glue helpers
# ---------------------------------------------------------------------------
function Get-ITGlueOrgForSite {
    param(
        [string]$SiteName,
        [string]$SiteId,
        $Config
    )

    # Check for saved mapping
    $savedOrgId = $Config.SiteMappings[$SiteId]
    if ($savedOrgId) {
        # Verify the org still exists
        try {
            $existing = Get-ITGlueOrganizations -id $savedOrgId
            if ($existing.data) {
                $orgName = $existing.data.attributes.name
                Write-Host "Saved IT Glue organization: " -NoNewline -ForegroundColor Cyan
                Write-Host "$orgName (ID: $savedOrgId)" -ForegroundColor White
                $confirm = Read-Host "Use this organization? (Y/n)"
                if ($confirm -ne 'n' -and $confirm -ne 'N') {
                    return @{ Id = $savedOrgId; Name = $orgName }
                }
            }
        }
        catch {
            Write-Warning "Saved org ID $savedOrgId could not be resolved. Re-selecting."
        }
    }

    # Try name-based search first
    Write-Host "Searching IT Glue for organizations matching '$SiteName'..." -ForegroundColor Yellow
    $orgs = @()
    try {
        $result = Get-ITGlueOrganizations -filter_name $SiteName -page_size 50
        if ($result.data) { $orgs = $result.data }
    }
    catch {
        Write-Warning "Name-filtered search failed: $($_.Exception.Message)"
    }

    if ($orgs.Count -eq 0) {
        Write-Host "No close matches found. Loading all organizations..." -ForegroundColor Yellow
        $page = 1
        do {
            $result = Get-ITGlueOrganizations -page_size 100 -page_number $page
            if ($result.data) { $orgs += $result.data }
            $page++
        } while ($result.data -and $result.data.Count -eq 100)
    }

    if ($orgs.Count -eq 0) {
        Write-Error "No IT Glue organizations found."
    }

    Write-Host ""
    Write-Host "Select the IT Glue organization for Unifi site '$SiteName':" -ForegroundColor Cyan
    for ($i = 0; $i -lt $orgs.Count; $i++) {
        Write-Host "  [$($i + 1)] $($orgs[$i].attributes.name)" -ForegroundColor White
    }
    Write-Host ""
    do {
        $sel = Read-Host "Enter number (1-$($orgs.Count))"
    } while (-not ($sel -match '^\d+$') -or [int]$sel -lt 1 -or [int]$sel -gt $orgs.Count)

    $chosen = $orgs[[int]$sel - 1]
    $orgId   = [int]$chosen.id
    $orgName = $chosen.attributes.name

    # Save mapping
    $Config.SiteMappings[$SiteId] = $orgId
    Save-Credentials -Config $Config
    Write-Host "Mapping saved: site '$SiteName' -> '$orgName'" -ForegroundColor Green

    return @{ Id = $orgId; Name = $orgName }
}

function Get-ITGlueConfigTypeMap {
    # Returns a hashtable mapping Unifi type strings → IT Glue config type IDs
    $typeNames = @{
        Switch   = 'Managed Network Switch'
        Firewall = 'Managed Network Firewall'
        Wifi     = 'Managed Network WiFi Access'
        Device   = 'Managed Network Device'
    }

    Write-Host "Resolving IT Glue configuration types..." -ForegroundColor Yellow
    $allTypes = Get-ITGlueConfigurationTypes
    $typeData = $allTypes.data

    $resolved = @{}
    foreach ($key in $typeNames.Keys) {
        $name   = $typeNames[$key]
        $match  = $typeData | Where-Object { $_.attributes.name -eq $name }
        if (-not $match) {
            Write-Error "IT Glue configuration type '$name' not found. Ensure it exists (synced from ConnectWise)."
        }
        $resolved[$key] = [int]$match.id
    }

    # Build Unifi-type → resolved key map
    return @{
        usw     = $resolved.Switch
        ugw     = $resolved.Firewall
        udm     = $resolved.Firewall
        udmpro  = $resolved.Firewall
        usg     = $resolved.Firewall
        usg3p   = $resolved.Firewall
        ucg     = $resolved.Firewall
        uxg     = $resolved.Firewall
        uap     = $resolved.Wifi
    }
}

# ---------------------------------------------------------------------------
# Port table builder
# ---------------------------------------------------------------------------
function Build-PortTable {
    param($Device)

    $html = @()

    $isSwitch  = $Device.type -in @('usw')
    $isGateway = $Device.type -in @('ugw','udm','udmpro','usg','usg3p','ucg','uxg')
    $isAP      = $Device.type -in @('uap')

    if ($isAP) {
        # Access Point: show uplink and radio info
        $html += "<h3>Uplink</h3>"
        $html += "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;font-family:monospace;font-size:12px'>"
        $html += "<tr style='background:#eee'><th>Interface</th><th>Uplink Device</th><th>Uplink MAC</th><th>Uplink Port</th><th>IP</th></tr>"

        $uplinkDevice = if ($Device.uplink) { $Device.uplink.'uplink_device_name' } else { '-' }
        $uplinkMac    = if ($Device.uplink) { $Device.uplink.uplink_mac }           else { '-' }
        $uplinkPort   = if ($Device.uplink) { $Device.uplink.uplink_remote_port }   else { '-' }
        $uplinkIp     = if ($Device.uplink) { $Device.uplink.ip }                   else { $Device.ip }

        $html += "<tr><td>eth0</td><td>$uplinkDevice</td><td>$uplinkMac</td><td>$uplinkPort</td><td>$uplinkIp</td></tr>"
        $html += "</table>"

        # Radios
        if ($Device.radio_table) {
            $html += "<h3>Radios</h3>"
            $html += "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;font-family:monospace;font-size:12px'>"
            $html += "<tr style='background:#eee'><th>Radio</th><th>Band</th><th>Channel</th><th>TX Power</th><th>Clients</th></tr>"
            foreach ($radio in $Device.radio_table) {
                $band    = $radio.radio
                $channel = $radio.channel
                $txPower = if ($radio.tx_power) { "$($radio.tx_power) dBm" } else { 'Auto' }
                # Client count comes from radio_table_stats if present
                $clients = '-'
                if ($Device.radio_table_stats) {
                    $stat = $Device.radio_table_stats | Where-Object { $_.radio -eq $band }
                    if ($stat) { $clients = $stat.'num_sta' }
                }
                $html += "<tr><td>$band</td><td>$($radio.radio)</td><td>$channel</td><td>$txPower</td><td>$clients</td></tr>"
            }
            $html += "</table>"
        }

        return $html -join "`n"
    }

    if ($isSwitch -or $isGateway) {
        # Build LLDP lookup: local_port_idx → neighbor info
        $lldpMap = @{}
        if ($Device.lldp_table) {
            foreach ($entry in $Device.lldp_table) {
                $idx = $entry.local_port_idx
                if ($null -ne $idx) {
                    $lldpMap[$idx] = $entry
                }
            }
        }

        $html += "<h3>Port Connectivity</h3>"
        $html += "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;font-family:monospace;font-size:12px'>"
        $html += "<tr style='background:#eee'><th>Port</th><th>Name</th><th>Status</th><th>Speed (Mbps)</th><th>PoE</th><th>Connected Device</th><th>Connected MAC</th><th>Connected Port</th></tr>"

        $ports = if ($Device.port_table) { $Device.port_table } else { @() }
        # Sort by port index
        $ports = $ports | Sort-Object { [int]($_.port_idx) }

        foreach ($port in $ports) {
            $idx    = $port.port_idx
            $name   = if ($port.name)       { $port.name }       else { "Port $idx" }
            $up     = $port.up
            $status = if ($up) { 'Up' } else { 'Down' }
            $speed  = if ($up -and $port.speed) { $port.speed } else { '-' }

            # PoE
            $poe = '-'
            if ($port.poe_mode -and $port.poe_mode -ne 'off') {
                $poe = if ($port.poe_power) { "$($port.poe_power)W" } else { $port.poe_mode }
            }

            # LLDP neighbor
            $neighborName = '-'
            $neighborMac  = '-'
            $neighborPort = '-'
            if ($lldpMap.ContainsKey($idx)) {
                $n            = $lldpMap[$idx]
                $neighborName = if ($n.chassis_id)  { $n.chassis_id }  else { '-' }
                # Prefer system name if available
                if ($n.system_name) { $neighborName = $n.system_name }
                $neighborMac  = if ($n.chassis_id)  { $n.chassis_id }  else { '-' }
                $neighborPort = if ($n.port_id)     { $n.port_id }     else { '-' }
            }
            # Also check uplink port
            if ($Device.uplink -and $Device.uplink.port_idx -eq $idx) {
                if ($neighborName -eq '-') { $neighborName = $Device.uplink.'uplink_device_name' }
                if ($neighborMac  -eq '-') { $neighborMac  = $Device.uplink.uplink_mac }
            }

            $rowBg = if ($up) { '' } else { " style='color:#999'" }
            $html += "<tr$rowBg><td>$idx</td><td>$name</td><td>$status</td><td>$speed</td><td>$poe</td><td>$neighborName</td><td>$neighborMac</td><td>$neighborPort</td></tr>"
        }

        $html += "</table>"
        return $html -join "`n"
    }

    return ''
}

# ---------------------------------------------------------------------------
# Configuration sync
# ---------------------------------------------------------------------------
function Get-UnifiDeviceType {
    param([string]$Type)
    switch -Regex ($Type) {
        '^usw'           { return 'switch' }
        '^(ugw|usg|udm|ucg|uxg)' { return 'gateway' }
        '^uap'           { return 'ap' }
        default          { return 'unknown' }
    }
}

function Sync-DeviceToITGlue {
    param(
        $Device,
        [int]$OrgId,
        [hashtable]$ConfigTypeMap
    )

    $mac      = $Device.mac
    $model    = if ($Device.model_name) { $Device.model_name } elseif ($Device.model) { $Device.model } else { '' }

    # Build display name: prefer user-set name, then model+MAC suffix, then MAC
    if (-not [string]::IsNullOrWhiteSpace($Device.name)) {
        $name = $Device.name
    } else {
        $macClean  = ($mac -replace '[:\-]','').ToUpper()
        $macSuffix = if ($macClean.Length -ge 6) { $macClean.Substring($macClean.Length - 6) } else { $macClean }
        $name = if ($model) { "$model-$macSuffix" } else { $mac }
    }

    $ip       = if ($Device.ip)       { $Device.ip }       else { '' }
    $firmware = if ($Device.version)  { $Device.version }  else { '' }

    # Only use hostname if it looks like a real device name (not the controller FQDN)
    $rawHostname = if ($Device.hostname) { $Device.hostname } else { '' }
    $hostname = if ($rawHostname -and $rawHostname -notmatch 'unifi\.|\.ui\.com$') { $rawHostname } else { '' }

    # Resolve config type ID
    $configTypeId = $ConfigTypeMap[$Device.type]
    if (-not $configTypeId) {
        Write-Warning "  No config type mapped for Unifi type '$($Device.type)' - skipping $name"
        return 'skipped'
    }

    # Build notes (port table + model/firmware header)
    $portTable = Build-PortTable -Device $Device
    $notesHeader = "<p><strong>Model:</strong> $model | <strong>Firmware:</strong> $firmware | <strong>MAC:</strong> $mac</p>"
    $notes = $notesHeader + "`n" + $portTable

    # Check for existing configuration by MAC
    $existing = $null
    try {
        $searchResult = Get-ITGlueConfigurations -organization_id $OrgId -filter_mac_address $mac
        if ($searchResult.data -and $searchResult.data.Count -gt 0) {
            $existing = $searchResult.data[0]
        }
    }
    catch {
        Write-Warning "  Could not search for existing config (MAC: $mac): $($_.Exception.Message)"
    }

    $configAttribs = @{
        'organization-id'       = $OrgId
        'configuration-type-id' = $configTypeId
        'name'                  = $name
        'hostname'              = $hostname
        'primary-ip'            = $ip
        'mac-address'           = $mac
        'notes'                 = $notes
    }

    if ($existing) {
        # Update
        $configId = [int]$existing.id
        try {
            Set-ITGlueConfigurations -id $configId -data @{
                type       = 'configurations'
                attributes = $configAttribs
            } | Out-Null
            Write-Host "  Updated : $name ($mac)" -ForegroundColor Green
        }
        catch {
            Write-Host "  ERROR updating $name : $($_.Exception.Message)" -ForegroundColor Red
            return 'error'
        }
    }
    else {
        # Create
        try {
            $created = New-ITGlueConfigurations -data @{
                type       = 'configurations'
                attributes = $configAttribs
            }
            $configId = [int]$created.data.id
            Write-Host "  Created : $name ($mac)" -ForegroundColor Cyan
        }
        catch {
            Write-Host "  ERROR creating $name : $($_.Exception.Message)" -ForegroundColor Red
            return 'error'
        }
    }

    # Sync primary management interface
    try {
        $existingIfaces = Get-ITGlueConfigurationInterfaces -configuration_id $configId
        $mgmtIface = $existingIfaces.data | Where-Object { $_.attributes.primary -eq $true } | Select-Object -First 1

        $ifaceAttribs = @{
            'configuration-id' = $configId
            'name'             = 'Management'
            'ip-address'       = $ip
            'mac-address'      = $mac
            'primary'          = $true
            'notes'            = "Model: $model | Firmware: $firmware"
        }

        if ($mgmtIface) {
            Set-ITGlueConfigurationInterfaces -configuration_id $configId -id ([int]$mgmtIface.id) -data @{
                type       = 'configuration-interfaces'
                attributes = $ifaceAttribs
            } | Out-Null
        }
        else {
            New-ITGlueConfigurationInterfaces -configuration_id $configId -data @{
                type       = 'configuration-interfaces'
                attributes = $ifaceAttribs
            } | Out-Null
        }
    }
    catch {
        Write-Warning "  Could not sync interface for $name : $($_.Exception.Message)"
    }

    return if ($existing) { 'updated' } else { 'created' }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
try {
    Write-Banner "Unifi → IT Glue Configuration Sync"

    # 1. Credentials
    $creds = Initialize-Credentials -Reset:$Reset

    # 2. Connect to IT Glue
    Add-ITGlueBaseURI -base_uri $creds.ITGlue.BaseUri
    Add-ITGlueAPIKey  -Api_Key  $creds.ITGlue.ApiKey

    # 3. Connect to Unifi
    Initialize-UnifiApi -Config $creds

    # 4. List sites
    Write-Host "Fetching Unifi sites..." -ForegroundColor Yellow
    $sites = Get-UnifiSites
    if (-not $sites -or $sites.Count -eq 0) {
        Write-Error "No sites found on the controller."
    }

    Write-Host ""
    Write-Host "Select a Unifi site:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $sites.Count; $i++) {
        Write-Host "  [$($i + 1)] $($sites[$i]._displayName)" -ForegroundColor White
    }
    Write-Host ""
    do {
        $sel = Read-Host "Enter number (1-$($sites.Count))"
    } while (-not ($sel -match '^\d+$') -or [int]$sel -lt 1 -or [int]$sel -gt $sites.Count)

    $selectedSite = $sites[[int]$sel - 1]
    $siteId   = $selectedSite.siteId
    $siteName = $selectedSite._displayName
    Write-Host "Selected site: $siteName ($siteId)" -ForegroundColor Green

    # 5. Resolve IT Glue org
    $org = Get-ITGlueOrgForSite -SiteName $siteName -SiteId $siteId -Config $creds

    Write-Banner "Syncing: $siteName → $($org.Name)"

    # 6. Config type map
    $configTypeMap = Get-ITGlueConfigTypeMap

    # 7. Fetch devices
    Write-Host "Fetching Unifi devices for site '$siteName'..." -ForegroundColor Yellow
    $devices = Get-UnifiDevices -SiteId $siteId

    if (-not $devices -or $devices.Count -eq 0) {
        Write-Host "No switches, gateways, or APs found on this site." -ForegroundColor Yellow
        return
    }

    Write-Host "Found $($devices.Count) device(s) to sync." -ForegroundColor White
    Write-Host ""

    # 8. Sync each device
    $counts = @{ created = 0; updated = 0; skipped = 0; error = 0 }

    foreach ($device in $devices) {
        $displayName = if ($device.name) { $device.name }
                       elseif ($device.mac)  { $device.mac }
                       elseif ($device.macAddress) { $device.macAddress }
                       else { 'unknown' }
        $devType = if ($device.type) { $device.type } elseif ($device.productType) { $device.productType } else { '?' }
        Write-Host "Processing: $displayName [$devType]" -ForegroundColor Yellow

        $result = Sync-DeviceToITGlue -Device $device -OrgId $org.Id -ConfigTypeMap $configTypeMap
        $counts[$result]++
    }

    # 9. Summary
    Write-Host ""
    Write-Host ("=" * 65) -ForegroundColor Cyan
    Write-Host "  Sync complete for: $siteName → $($org.Name)" -ForegroundColor Cyan
    Write-Host "  Created : $($counts.created)" -ForegroundColor Cyan
    Write-Host "  Updated : $($counts.updated)" -ForegroundColor Cyan
    Write-Host "  Skipped : $($counts.skipped)" -ForegroundColor Cyan
    Write-Host "  Errors  : $($counts.error)"  -ForegroundColor $(if ($counts.error -gt 0) { 'Red' } else { 'Cyan' })
    Write-Host ("=" * 65) -ForegroundColor Cyan
    Write-Host ""
}
finally {
    Disconnect-UnifiController
}
