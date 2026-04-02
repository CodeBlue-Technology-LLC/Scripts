<#
.SYNOPSIS
    Analyzes User Profile Disk (VHDX) files to identify stale disks based on last write date.

.DESCRIPTION
    Scans a UPD share for UVHD-*.vhdx files named by SID and identifies those
    not written within a specified number of days. Resolves SIDs to AD user accounts
    using DirectoryServices - NO ActiveDirectory module required.

.PARAMETER UPDPath
    Path to the UPD share containing .vhdx files (e.g., "\\server\UPDs")

.PARAMETER DaysInactive
    Number of days since last write to consider a UPD stale

.PARAMETER ExportPath
    Optional CSV export path (default: current directory with timestamp)

.EXAMPLE
    .\Analyze-StaleUPDProfiles-NoModule.ps1 -UPDPath "\\fileserver\UPDs" -DaysInactive 180

.EXAMPLE
    .\Analyze-StaleUPDProfiles-NoModule.ps1 -UPDPath "D:\UPDs" -DaysInactive 365 -ExportPath "C:\Reports\StaleUPDs.csv"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$UPDPath,

    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 7300)]
    [int]$DaysInactive,

    [Parameter(Mandatory = $false)]
    [string]$ExportPath = ""
)

$CutoffDate = (Get-Date).AddDays(-$DaysInactive)
$StartTime  = Get-Date

Write-Host "===================================================================" -ForegroundColor Cyan
Write-Host "  Stale UPD (User Profile Disk) Analysis" -ForegroundColor Cyan
Write-Host "===================================================================" -ForegroundColor Cyan
Write-Host "UPD Path:        $UPDPath" -ForegroundColor White
Write-Host "Days Inactive:   $DaysInactive days" -ForegroundColor White
Write-Host "Cutoff Date:     $($CutoffDate.ToString('yyyy-MM-dd'))" -ForegroundColor White
Write-Host "Start Time:      $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor White
Write-Host "===================================================================" -ForegroundColor Cyan
Write-Host ""

function Format-FileSize {
    param([long]$Bytes)
    if ($Bytes -ge 1TB)     { return "{0:N2} TB" -f ($Bytes / 1TB) }
    elseif ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    elseif ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    else                    { return "{0} Bytes" -f $Bytes }
}

function Get-ADInfoBySID {
    param([string]$SID)
    try {
        $SIDObj    = New-Object System.Security.Principal.SecurityIdentifier($SID)
        $NTAccount = $SIDObj.Translate([System.Security.Principal.NTAccount])
        $Username  = $NTAccount.Value.Split('\')[-1]

        $Domain   = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        $DomainDN = ($Domain.Name.Split('.') | ForEach-Object { "DC=$_" }) -join ','

        $Searcher = New-Object System.DirectoryServices.DirectorySearcher
        $Searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomainDN")
        $Searcher.Filter = "(&(objectCategory=User)(sAMAccountName=$Username))"
        $Searcher.PropertiesToLoad.Add("lastLogon")          | Out-Null
        $Searcher.PropertiesToLoad.Add("userAccountControl") | Out-Null
        $Searcher.PropertiesToLoad.Add("displayName")        | Out-Null

        $Result = $Searcher.FindOne()

        if ($null -eq $Result) {
            return @{ Exists = $false; Username = $Username; DisplayName = "N/A"; LastLogon = "N/A"; Enabled = "N/A"; Error = $null }
        }

        $LastLogon = "Never"
        if ($Result.Properties["lastLogon"].Count -gt 0) {
            $Val = $Result.Properties["lastLogon"][0]
            try {
                $dt = [DateTime]::FromFileTime([Int64]$Val)
                if ($dt.Year -gt 1900) { $LastLogon = $dt.ToString('yyyy-MM-dd') }
            } catch {}
        }

        $Enabled = $true
        if ($Result.Properties["userAccountControl"].Count -gt 0) {
            $Enabled = -not ($Result.Properties["userAccountControl"][0] -band 2)
        }

        $DisplayName = if ($Result.Properties["displayName"].Count -gt 0) { $Result.Properties["displayName"][0] } else { $Username }

        return @{ Exists = $true; Username = $Username; DisplayName = $DisplayName; LastLogon = $LastLogon; Enabled = $Enabled; Error = $null }

    } catch [System.Security.Principal.IdentityNotMappedException] {
        return @{ Exists = $false; Username = "Orphaned SID"; DisplayName = "N/A"; LastLogon = "N/A"; Enabled = "N/A"; Error = "SID not resolvable" }
    } catch {
        return @{ Exists = "Error"; Username = "Unknown"; DisplayName = "N/A"; LastLogon = "N/A"; Enabled = "N/A"; Error = $_.Exception.Message }
    }
}

Write-Host "Testing Active Directory connectivity..." -ForegroundColor Yellow
$ADAvailable = $false
try {
    $TestDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
    Write-Host "[OK] Connected to domain: $($TestDomain.Name)" -ForegroundColor Green
    $ADAvailable = $true
} catch {
    Write-Host "[WARNING] AD unavailable - SID resolution will be skipped" -ForegroundColor Yellow
}
Write-Host ""

Write-Host "Scanning for UVHD VHDX files in: $UPDPath" -ForegroundColor Yellow

$AllVHDX = Get-ChildItem -Path $UPDPath -Filter "UVHD-*.vhdx" -ErrorAction SilentlyContinue |
           Where-Object { $_.BaseName -match '^UVHD-S-1-\d+-\d+(-\d+)+$' }

if ($AllVHDX.Count -eq 0) {
    Write-Host "[ERROR] No UVHD SID-named .vhdx files found in $UPDPath" -ForegroundColor Red
    exit
}

Write-Host "Found $($AllVHDX.Count) UPD files to evaluate" -ForegroundColor Cyan
Write-Host ""

$Results    = @()
$Processed  = 0
$ErrorCount = 0

foreach ($VHD in $AllVHDX) {
    $Processed++
    $Pct = [math]::Round(($Processed / $AllVHDX.Count) * 100, 1)
    Write-Progress -Activity "Analyzing UPD Files" -Status "$($VHD.Name) - $Processed of $($AllVHDX.Count)" -PercentComplete $Pct

    $IsStale = $VHD.LastWriteTime -lt $CutoffDate
    if (-not $IsStale) { continue }

    $LastWriteStr = $VHD.LastWriteTime.ToString('yyyy-MM-dd')
    Write-Host "[->] $($VHD.Name) - Last write: $LastWriteStr" -ForegroundColor Gray

    $SID = $VHD.BaseName -replace '^UVHD-', ''

    if ($ADAvailable) {
        $ADInfo = Get-ADInfoBySID -SID $SID
        if ($ADInfo.Error) {
            Write-Host "     AD Note: $($ADInfo.Error)" -ForegroundColor DarkYellow
        }
    } else {
        $ADInfo = @{ Exists = "N/A"; Username = "N/A"; DisplayName = "N/A"; LastLogon = "N/A"; Enabled = "N/A"; Error = $null }
    }

    $DaysSince    = (New-TimeSpan -Start $VHD.LastWriteTime -End (Get-Date)).Days
    $LastWriteFull = $VHD.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')

    $Result = [PSCustomObject]@{
        SID                = $SID
        VHDXFile           = $VHD.Name
        VHDXPath           = $VHD.FullName
        LastWriteDate      = $LastWriteFull
        DaysSinceLastWrite = $DaysSince
        SizeBytes          = $VHD.Length
        SizeFormatted      = Format-FileSize -Bytes $VHD.Length
        ADUserExists       = $ADInfo.Exists
        Username           = $ADInfo.Username
        DisplayName        = $ADInfo.DisplayName
        ADLastLogon        = $ADInfo.LastLogon
        ADEnabled          = $ADInfo.Enabled
    }

    $Results += $Result

    $StatusColor = if ($ADInfo.Exists -eq $false)      { "Red" }
                   elseif ($ADInfo.Exists -eq "Error")  { "DarkYellow" }
                   elseif ($ADInfo.Enabled -eq $false)  { "Yellow" }
                   else                                 { "White" }

    Write-Host "     User: $($ADInfo.Username) ($($ADInfo.DisplayName)) | Size: $(Format-FileSize -Bytes $VHD.Length) | AD Enabled: $($ADInfo.Enabled)" -ForegroundColor $StatusColor
}

Write-Progress -Activity "Analyzing UPD Files" -Completed

Write-Host ""
Write-Host "===================================================================" -ForegroundColor Cyan
Write-Host "  ANALYSIS SUMMARY" -ForegroundColor Cyan
Write-Host "===================================================================" -ForegroundColor Cyan

$TotalSize     = ($Results | Measure-Object -Property SizeBytes -Sum).Sum
$NoADUser      = @($Results | Where-Object { $_.ADUserExists -eq $false })
$DisabledUsers = @($Results | Where-Object { $_.ADEnabled -eq $false })
$ADErrors      = @($Results | Where-Object { $_.ADUserExists -eq "Error" })
$NoADSize      = ($NoADUser      | Measure-Object -Property SizeBytes -Sum).Sum
$DisabledSize  = ($DisabledUsers | Measure-Object -Property SizeBytes -Sum).Sum

Write-Host "Total VHDX Files Scanned:        $($AllVHDX.Count)" -ForegroundColor White
Write-Host "Stale UPDs Found:                $($Results.Count)" -ForegroundColor Yellow
Write-Host "Orphaned / No AD User:           $($NoADUser.Count)" -ForegroundColor Red
Write-Host "Disabled AD Account:             $($DisabledUsers.Count)" -ForegroundColor Yellow
if ($ADErrors.Count -gt 0) {
    Write-Host "AD Lookup Errors:                $($ADErrors.Count)" -ForegroundColor DarkYellow
}
Write-Host ""
Write-Host "Total Stale UPD Size:            $(Format-FileSize -Bytes $TotalSize)" -ForegroundColor Cyan
Write-Host "Size - Orphaned SIDs:            $(Format-FileSize -Bytes $NoADSize)" -ForegroundColor Red
Write-Host "Size - Disabled Accounts:        $(Format-FileSize -Bytes $DisabledSize)" -ForegroundColor Yellow
Write-Host ""

if ($Results.Count -gt 0) {
    Write-Host "===================================================================" -ForegroundColor Cyan
    Write-Host "  TOP 10 LARGEST STALE UPDs" -ForegroundColor Cyan
    Write-Host "===================================================================" -ForegroundColor Cyan
    Write-Host ""

    $Top10 = $Results | Sort-Object SizeBytes -Descending | Select-Object -First 10
    $n = 1
    foreach ($Item in $Top10) {
        $Tag = if ($Item.ADUserExists -eq $false)      { "[ORPHANED]" }
               elseif ($Item.ADUserExists -eq "Error")  { "[AD ERROR]" }
               elseif ($Item.ADEnabled -eq $false)      { "[DISABLED]" }
               else                                     { "[ACTIVE]" }

        Write-Host "$n. $($Item.Username) / $($Item.DisplayName) $Tag" -ForegroundColor White
        Write-Host "   Size: $($Item.SizeFormatted) | Last Write: $($Item.LastWriteDate) | Days: $($Item.DaysSinceLastWrite)" -ForegroundColor Gray
        Write-Host "   SID:  $($Item.SID)" -ForegroundColor DarkGray
        Write-Host ""
        $n++
    }
}

if ($Results.Count -gt 0) {
    if ([string]::IsNullOrEmpty($ExportPath)) {
        $Ts = Get-Date -Format "yyyyMMdd_HHmmss"
        $ExportPath = Join-Path (Get-Location) "StaleUPDs_$Ts.csv"
    }
    try {
        $Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
        Write-Host "===================================================================" -ForegroundColor Cyan
        Write-Host "[SUCCESS] Results exported to: $ExportPath" -ForegroundColor Green
        Write-Host "===================================================================" -ForegroundColor Cyan
    } catch {
        Write-Host "[ERROR] Export failed: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "[INFO] No stale UPDs found - nothing to export" -ForegroundColor Yellow
}

$Duration = New-TimeSpan -Start $StartTime -End (Get-Date)
$DurationStr = $Duration.ToString('mm\:ss')
Write-Host ""
Write-Host "Completed in $DurationStr" -ForegroundColor Cyan
Write-Host ""
