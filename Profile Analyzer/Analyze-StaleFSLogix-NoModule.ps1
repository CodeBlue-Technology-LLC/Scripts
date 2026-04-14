<#
.SYNOPSIS
    Analyzes FSLogix profile containers to identify stale profiles based on VHDX last write date.

.DESCRIPTION
    Scans an FSLogix profile share for folders named SID_Username, locates the Profile_*.VHDX
    inside each, and flags those not written within a specified number of days.
    Resolves SIDs to AD for enabled/lastlogon status using DirectoryServices.
    NO ActiveDirectory module required.

.PARAMETER ProfilePath
    Path to the FSLogix profile share (e.g., "\\server\FSLogix\User Profiles")

.PARAMETER DaysInactive
    Number of days since last VHDX write to consider a profile stale

.PARAMETER ExportPath
    Optional CSV export path (default: current directory with timestamp)

.EXAMPLE
    .\Analyze-StaleFSLogix-NoModule.ps1 -ProfilePath "E:\Shares\FSLogix\User Profiles" -DaysInactive 180

.EXAMPLE
    .\Analyze-StaleFSLogix-NoModule.ps1 -ProfilePath "\\fileserver\FSLogix\Profiles" -DaysInactive 365 -ExportPath "C:\Reports\StaleFSLogix.csv"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$ProfilePath,

    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 7300)]
    [int]$DaysInactive,

    [Parameter(Mandatory = $false)]
    [string]$ExportPath = ""
)

$CutoffDate = (Get-Date).AddDays(-$DaysInactive)
$StartTime  = Get-Date

Write-Host "===================================================================" -ForegroundColor Cyan
Write-Host "  Stale FSLogix Profile Container Analysis" -ForegroundColor Cyan
Write-Host "===================================================================" -ForegroundColor Cyan
Write-Host "Profile Path:    $ProfilePath" -ForegroundColor White
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

        # Bind to domain root using machine context - avoids UPN vs AD domain name mismatch
        $Searcher = New-Object System.DirectoryServices.DirectorySearcher
        $Searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry
        $Searcher.Filter = "(&(objectCategory=User)(sAMAccountName=$Username))"
        $Searcher.PropertiesToLoad.Add("lastLogon")          | Out-Null
        $Searcher.PropertiesToLoad.Add("userAccountControl") | Out-Null
        $Searcher.PropertiesToLoad.Add("displayName")        | Out-Null

        $Result = $Searcher.FindOne()

        if ($null -eq $Result) {
            return @{ Exists = $false; DisplayName = "N/A"; LastLogon = "N/A"; Enabled = "N/A"; Error = $null }
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

        return @{ Exists = $true; DisplayName = $DisplayName; LastLogon = $LastLogon; Enabled = $Enabled; Error = $null }

    } catch [System.Security.Principal.IdentityNotMappedException] {
        return @{ Exists = $false; DisplayName = "N/A"; LastLogon = "N/A"; Enabled = "N/A"; Error = "SID not resolvable" }
    } catch {
        return @{ Exists = "Error"; DisplayName = "N/A"; LastLogon = "N/A"; Enabled = "N/A"; Error = $_.Exception.Message }
    }
}

# AD connectivity test
Write-Host "Testing Active Directory connectivity..." -ForegroundColor Yellow
$ADAvailable = $false
try {
    $TestEntry = New-Object System.DirectoryServices.DirectoryEntry
    if ($TestEntry.Name) {
        Write-Host "[OK] Connected to AD: $($TestEntry.Name)" -ForegroundColor Green
        $ADAvailable = $true
    } else {
        Write-Host "[WARNING] AD unavailable - SID resolution will be skipped" -ForegroundColor Yellow
    }
} catch {
    Write-Host "[WARNING] AD unavailable - SID resolution will be skipped" -ForegroundColor Yellow
}
Write-Host ""

# Enumerate FSLogix profile folders - pattern: SID_Username
Write-Host "Scanning FSLogix profile folders in: $ProfilePath" -ForegroundColor Yellow

$AllFolders = Get-ChildItem -Path $ProfilePath -Directory -ErrorAction SilentlyContinue |
              Where-Object { $_.Name -match '^S-1-\d+-\d+(-\d+)+_\S+' }

if ($AllFolders.Count -eq 0) {
    Write-Host "[ERROR] No FSLogix profile folders found matching SID_Username pattern in $ProfilePath" -ForegroundColor Red
    exit
}

Write-Host "Found $($AllFolders.Count) profile folders to evaluate" -ForegroundColor Cyan
Write-Host ""

$Results    = @()
$Processed  = 0
$ErrorCount = 0

foreach ($Folder in $AllFolders) {
    $Processed++
    $Pct = [math]::Round(($Processed / $AllFolders.Count) * 100, 1)
    Write-Progress -Activity "Analyzing FSLogix Profiles" -Status "$($Folder.Name) - $Processed of $($AllFolders.Count)" -PercentComplete $Pct

    # Parse SID and username - split on first underscore only
    $SplitIdx = $Folder.Name.IndexOf('_')
    $SID      = $Folder.Name.Substring(0, $SplitIdx)
    $Username = $Folder.Name.Substring($SplitIdx + 1)

    # Find the VHDX inside - skip metadata files
    $VHDX = Get-ChildItem -Path $Folder.FullName -Filter "Profile_*.VHDX" -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -ieq ".vhdx" } |
            Select-Object -First 1

    if ($null -eq $VHDX) {
        Write-Host "[WARN] No Profile VHDX found in: $($Folder.Name)" -ForegroundColor DarkYellow
        $ErrorCount++
        continue
    }

    $IsStale = $VHDX.LastWriteTime -lt $CutoffDate
    if (-not $IsStale) { continue }

    $LastWriteStr = $VHDX.LastWriteTime.ToString('yyyy-MM-dd')
    Write-Host "[->] $($Folder.Name) - Last write: $LastWriteStr" -ForegroundColor Gray

    if ($ADAvailable) {
        $ADInfo = Get-ADInfoBySID -SID $SID
        if ($ADInfo.Error) {
            Write-Host "     AD Note: $($ADInfo.Error)" -ForegroundColor DarkYellow
        }
    } else {
        $ADInfo = @{ Exists = "N/A"; DisplayName = "N/A"; LastLogon = "N/A"; Enabled = "N/A"; Error = $null }
    }

    $DaysSince     = (New-TimeSpan -Start $VHDX.LastWriteTime -End (Get-Date)).Days
    $LastWriteFull = $VHDX.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')

    $MetaFile  = Get-ChildItem -Path $Folder.FullName -Filter "*.metadata" -File -ErrorAction SilentlyContinue | Select-Object -First 1
    $TotalSize = $VHDX.Length + $(if ($MetaFile) { $MetaFile.Length } else { 0 })

    $Result = [PSCustomObject]@{
        FolderName         = $Folder.Name
        SID                = $SID
        Username           = $Username
        DisplayName        = $ADInfo.DisplayName
        VHDXFile           = $VHDX.Name
        VHDXPath           = $VHDX.FullName
        LastWriteDate      = $LastWriteFull
        DaysSinceLastWrite = $DaysSince
        SizeBytes          = $TotalSize
        SizeFormatted      = Format-FileSize -Bytes $TotalSize
        ADUserExists       = $ADInfo.Exists
        ADLastLogon        = $ADInfo.LastLogon
        ADEnabled          = $ADInfo.Enabled
    }

    $Results += $Result

    $StatusColor = if ($ADInfo.Exists -eq $false)      { "Red" }
                   elseif ($ADInfo.Exists -eq "Error")  { "DarkYellow" }
                   elseif ($ADInfo.Enabled -eq $false)  { "Yellow" }
                   else                                 { "White" }

    Write-Host "     User: $Username ($($ADInfo.DisplayName)) | Size: $(Format-FileSize -Bytes $TotalSize) | AD Enabled: $($ADInfo.Enabled)" -ForegroundColor $StatusColor
}

Write-Progress -Activity "Analyzing FSLogix Profiles" -Completed

Write-Host ""
Write-Host "===================================================================" -ForegroundColor Cyan
Write-Host "  ANALYSIS SUMMARY" -ForegroundColor Cyan
Write-Host "===================================================================" -ForegroundColor Cyan

$TotalSize     = ($Results | Measure-Object -Property SizeBytes -Sum).Sum
$NoADUser      = @($Results | Where-Object { $_.ADUserExists -is [bool] -and $_.ADUserExists -eq $false })
$DisabledUsers = @($Results | Where-Object { $_.ADEnabled -is [bool] -and $_.ADEnabled -eq $false })
$ADErrors      = @($Results | Where-Object { $_.ADUserExists -is [string] -and $_.ADUserExists -eq "Error" })
$NoADSize      = ($NoADUser      | Measure-Object -Property SizeBytes -Sum).Sum
$DisabledSize  = ($DisabledUsers | Measure-Object -Property SizeBytes -Sum).Sum

Write-Host "Total Profile Folders Scanned:   $($AllFolders.Count)" -ForegroundColor White
Write-Host "Stale Profiles Found:            $($Results.Count)" -ForegroundColor Yellow
Write-Host "Orphaned / No AD User:           $($NoADUser.Count)" -ForegroundColor Red
Write-Host "Disabled AD Account:             $($DisabledUsers.Count)" -ForegroundColor Yellow
if ($ADErrors.Count -gt 0) {
    Write-Host "AD Lookup Errors:                $($ADErrors.Count)" -ForegroundColor DarkYellow
}
Write-Host "Warnings (no VHDX found):        $ErrorCount" -ForegroundColor $(if ($ErrorCount -gt 0) { "DarkYellow" } else { "Green" })
Write-Host ""
Write-Host "Total Stale Profile Size:        $(Format-FileSize -Bytes $TotalSize)" -ForegroundColor Cyan
Write-Host "Size - Orphaned SIDs:            $(Format-FileSize -Bytes $NoADSize)" -ForegroundColor Red
Write-Host "Size - Disabled Accounts:        $(Format-FileSize -Bytes $DisabledSize)" -ForegroundColor Yellow
Write-Host ""

if ($Results.Count -gt 0) {
    Write-Host "===================================================================" -ForegroundColor Cyan
    Write-Host "  TOP 10 LARGEST STALE PROFILES" -ForegroundColor Cyan
    Write-Host "===================================================================" -ForegroundColor Cyan
    Write-Host ""

    $Top10 = $Results | Sort-Object SizeBytes -Descending | Select-Object -First 10
    $n = 1
    foreach ($Item in $Top10) {
        $Tag = if ($Item.ADUserExists -is [bool] -and $Item.ADUserExists -eq $false) { "[ORPHANED]" }
               elseif ($Item.ADUserExists -is [string] -and $Item.ADUserExists -eq "Error") { "[AD ERROR]" }
               elseif ($Item.ADEnabled -is [bool] -and $Item.ADEnabled -eq $false) { "[DISABLED]" }
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
        $ExportPath = Join-Path (Get-Location) "StaleFSLogix_$Ts.csv"
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
    Write-Host "[INFO] No stale profiles found - nothing to export" -ForegroundColor Yellow
}

$Duration    = New-TimeSpan -Start $StartTime -End (Get-Date)
$DurationStr = $Duration.ToString('mm\:ss')
Write-Host ""
Write-Host "Completed in $DurationStr" -ForegroundColor Cyan
Write-Host ""
