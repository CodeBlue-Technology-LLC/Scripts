<#
.SYNOPSIS
    Analyzes user profile directories to identify stale folders based on file modification dates.

.DESCRIPTION
    This script recursively scans a directory structure (typically user profiles) and identifies
    parent folders where NO files have been modified within a specified number of days.
    Uses .NET DirectoryServices to query AD - NO ActiveDirectory module required!

.PARAMETER RootPath
    The root directory to scan (e.g., "C:\Shares\Users")

.PARAMETER DaysInactive
    Number of days to consider a folder stale if no files have been modified

.PARAMETER ExportPath
    Optional path to export results to CSV (default: current directory with timestamp)

.EXAMPLE
    .\Analyze-StaleProfiles-NoModule.ps1 -RootPath "C:\Shares\Users" -DaysInactive 365

.EXAMPLE
    .\Analyze-StaleProfiles-NoModule.ps1 -RootPath "C:\Shares\Users" -DaysInactive 730 -ExportPath "C:\Reports\StaleProfiles.csv"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$RootPath,

    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 7300)]
    [int]$DaysInactive,

    [Parameter(Mandatory = $false)]
    [string]$ExportPath = ""
)

$CutoffDate = (Get-Date).AddDays(-$DaysInactive)
$StartTime  = Get-Date

Write-Host "===================================================================" -ForegroundColor Cyan
Write-Host "  Stale Profile Folder Analysis" -ForegroundColor Cyan
Write-Host "===================================================================" -ForegroundColor Cyan
Write-Host "Root Path:       $RootPath" -ForegroundColor White
Write-Host "Days Inactive:   $DaysInactive days" -ForegroundColor White
Write-Host "Cutoff Date:     $($CutoffDate.ToString('yyyy-MM-dd'))" -ForegroundColor White
Write-Host "Start Time:      $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor White
Write-Host "===================================================================" -ForegroundColor Cyan
Write-Host ""

$Results       = @()
$ProcessedCount = 0
$ErrorCount    = 0

function Get-FolderSize {
    param([string]$Path)
    try {
        $Size = (Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
                 Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        if ($null -eq $Size) { return 0 }
        return $Size
    } catch { return 0 }
}

function Format-FileSize {
    param([long]$Bytes)
    if ($Bytes -ge 1TB)     { return "{0:N2} TB" -f ($Bytes / 1TB) }
    elseif ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    elseif ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    else                    { return "{0} Bytes" -f $Bytes }
}

function Get-ADUserInfo {
    param([string]$Username)
    try {
        $Searcher = New-Object System.DirectoryServices.DirectorySearcher
        $Searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry
        $Searcher.Filter = "(&(objectCategory=User)(sAMAccountName=$Username))"
        $Searcher.PropertiesToLoad.Add("lastLogon")          | Out-Null
        $Searcher.PropertiesToLoad.Add("userAccountControl") | Out-Null
        $Searcher.PropertiesToLoad.Add("distinguishedName")  | Out-Null

        $Result = $Searcher.FindOne()

        if ($null -eq $Result) {
            return @{ Exists = $false; LastLogon = "N/A"; Enabled = "N/A"; Error = $null }
        }

        $LastLogon = "Never"
        if ($Result.Properties["lastLogon"].Count -gt 0) {
            $Val = $Result.Properties["lastLogon"][0]
            try {
                $dt = [DateTime]::FromFileTime([Int64]$Val)
                if ($dt.Year -gt 1900) { $LastLogon = $dt.ToString('yyyy-MM-dd') }
            } catch { $LastLogon = "Never" }
        }

        $Enabled = $true
        if ($Result.Properties["userAccountControl"].Count -gt 0) {
            $Enabled = -not ($Result.Properties["userAccountControl"][0] -band 2)
        }

        return @{ Exists = $true; LastLogon = $LastLogon; Enabled = $Enabled; Error = $null }

    } catch {
        return @{ Exists = "Error"; LastLogon = "N/A"; Enabled = "N/A"; Error = $_.Exception.Message }
    }
}

Write-Host "Testing Active Directory connectivity..." -ForegroundColor Yellow
$ADAvailable = $false
try {
    $TestEntry = New-Object System.DirectoryServices.DirectoryEntry
    if ($TestEntry.Name) {
        Write-Host "[OK] Connected to AD: $($TestEntry.Name)" -ForegroundColor Green
        $ADAvailable = $true
    } else {
        Write-Host "[WARNING] AD unavailable - AD lookups will be skipped" -ForegroundColor Yellow
    }
} catch {
    Write-Host "[WARNING] AD unavailable - AD lookups will be skipped" -ForegroundColor Yellow
}
Write-Host ""

Write-Host "Scanning user profile directories..." -ForegroundColor Yellow
$UserFolders = Get-ChildItem -Path $RootPath -Directory -ErrorAction SilentlyContinue

if ($UserFolders.Count -eq 0) {
    Write-Host "[ERROR] No subdirectories found in $RootPath" -ForegroundColor Red
    exit
}

Write-Host "Found $($UserFolders.Count) user folders to analyze" -ForegroundColor Cyan
Write-Host ""

foreach ($Folder in $UserFolders) {
    $ProcessedCount++
    $PercentComplete = [math]::Round(($ProcessedCount / $UserFolders.Count) * 100, 1)
    Write-Progress -Activity "Analyzing Folders" -Status "Processing: $($Folder.Name) - $ProcessedCount of $($UserFolders.Count)" -PercentComplete $PercentComplete

    try {
        $Files = Get-ChildItem -Path $Folder.FullName -File -Recurse -ErrorAction SilentlyContinue

        if ($Files.Count -eq 0) {
            $MostRecentDate = $Folder.LastWriteTime
            $IsStale = $true
        } else {
            $MostRecentDate = ($Files | Sort-Object LastWriteTime -Descending | Select-Object -First 1).LastWriteTime
            $IsStale = $MostRecentDate -lt $CutoffDate
        }

        if ($IsStale) {
            $LastModStr = $MostRecentDate.ToString('yyyy-MM-dd')
            Write-Host "[->] Analyzing: $($Folder.Name) - Last modified: $LastModStr" -ForegroundColor Gray

            $FolderSizeBytes     = Get-FolderSize -Path $Folder.FullName
            $FolderSizeFormatted = Format-FileSize -Bytes $FolderSizeBytes

            if ($ADAvailable) {
                $ADInfo = Get-ADUserInfo -Username $Folder.Name
                if ($ADInfo.Error) {
                    Write-Host "     AD Lookup Error: $($ADInfo.Error)" -ForegroundColor DarkYellow
                }
            } else {
                $ADInfo = @{ Exists = "N/A"; LastLogon = "N/A"; Enabled = "N/A"; Error = $null }
            }

            $DaysSinceActivity = (New-TimeSpan -Start $MostRecentDate -End (Get-Date)).Days
            $LastModFull       = $MostRecentDate.ToString('yyyy-MM-dd HH:mm:ss')

            $Result = [PSCustomObject]@{
                FolderName            = $Folder.Name
                FolderPath            = $Folder.FullName
                LastModifiedDate      = $LastModFull
                DaysSinceLastActivity = $DaysSinceActivity
                SizeBytes             = $FolderSizeBytes
                SizeFormatted         = $FolderSizeFormatted
                FileCount             = $Files.Count
                ADUserExists          = $ADInfo.Exists
                ADLastLogon           = $ADInfo.LastLogon
                ADEnabled             = $ADInfo.Enabled
            }

            $Results += $Result

            $StatusColor = if ($ADInfo.Exists -is [bool] -and $ADInfo.Exists -eq $false)       { "Red" }
                           elseif ($ADInfo.Exists -is [string] -and $ADInfo.Exists -eq "Error") { "DarkYellow" }
                           elseif ($ADInfo.Enabled -is [bool] -and $ADInfo.Enabled -eq $false)  { "Yellow" }
                           else                                                                  { "White" }

            Write-Host "     Size: $FolderSizeFormatted | Files: $($Files.Count) | AD User: $($ADInfo.Exists) | Enabled: $($ADInfo.Enabled)" -ForegroundColor $StatusColor
        }

    } catch {
        Write-Host "[ERROR] Error processing $($Folder.Name): $($_.Exception.Message)" -ForegroundColor Red
        $ErrorCount++
    }
}

Write-Progress -Activity "Analyzing Folders" -Completed

Write-Host ""
Write-Host "===================================================================" -ForegroundColor Cyan
Write-Host "  ANALYSIS SUMMARY" -ForegroundColor Cyan
Write-Host "===================================================================" -ForegroundColor Cyan
Write-Host ""

$TotalSize        = ($Results | Measure-Object -Property SizeBytes -Sum).Sum
$NoADUser         = @($Results | Where-Object { $_.ADUserExists -is [bool] -and $_.ADUserExists -eq $false })
$DisabledUsers    = @($Results | Where-Object { $_.ADEnabled -is [bool] -and $_.ADEnabled -eq $false })
$ADErrors         = @($Results | Where-Object { $_.ADUserExists -is [string] -and $_.ADUserExists -eq "Error" })
$NoADUserSize     = ($NoADUser      | Measure-Object -Property SizeBytes -Sum).Sum
$DisabledUserSize = ($DisabledUsers | Measure-Object -Property SizeBytes -Sum).Sum

Write-Host "Total Folders Scanned:           $($UserFolders.Count)" -ForegroundColor White
Write-Host "Stale Folders Found:             $($Results.Count)" -ForegroundColor Yellow
Write-Host "Folders with No AD User:         $($NoADUser.Count)" -ForegroundColor Red
Write-Host "Folders with Disabled AD User:   $($DisabledUsers.Count)" -ForegroundColor Yellow
if ($ADErrors.Count -gt 0) {
    Write-Host "Folders with AD Lookup Errors:   $($ADErrors.Count)" -ForegroundColor DarkYellow
}
Write-Host "Errors Encountered:              $ErrorCount" -ForegroundColor $(if ($ErrorCount -gt 0) { "Red" } else { "Green" })
Write-Host ""
Write-Host "Total Size of Stale Folders:     $(Format-FileSize -Bytes $TotalSize)" -ForegroundColor Cyan
Write-Host "Size of No AD User Folders:      $(Format-FileSize -Bytes $NoADUserSize)" -ForegroundColor Red
Write-Host "Size of Disabled User Folders:   $(Format-FileSize -Bytes $DisabledUserSize)" -ForegroundColor Yellow
Write-Host ""
Write-Host "Potential Space Savings:         $(Format-FileSize -Bytes $TotalSize)" -ForegroundColor Green
Write-Host ""

if ($Results.Count -gt 0) {
    Write-Host "===================================================================" -ForegroundColor Cyan
    Write-Host "  TOP 10 LARGEST STALE FOLDERS" -ForegroundColor Cyan
    Write-Host "===================================================================" -ForegroundColor Cyan
    Write-Host ""

    $Top10   = $Results | Sort-Object SizeBytes -Descending | Select-Object -First 10
    $Counter = 1
    foreach ($Item in $Top10) {
        $ADStatus = if ($Item.ADUserExists -is [bool] -and $Item.ADUserExists -eq $false)       { "[NO AD USER]" }
                    elseif ($Item.ADUserExists -is [string] -and $Item.ADUserExists -eq "Error") { "[AD ERROR]" }
                    elseif ($Item.ADEnabled -is [bool] -and $Item.ADEnabled -eq $false)          { "[DISABLED]" }
                    else                                                                          { "[ACTIVE]" }

        Write-Host "$Counter. $($Item.FolderName) $ADStatus" -ForegroundColor White
        Write-Host "   Size: $($Item.SizeFormatted) | Last Activity: $($Item.LastModifiedDate) | Days: $($Item.DaysSinceLastActivity)" -ForegroundColor Gray
        Write-Host ""
        $Counter++
    }
}

if ($Results.Count -gt 0) {
    if ([string]::IsNullOrEmpty($ExportPath)) {
        $Timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
        $ExportPath = Join-Path (Get-Location) "StaleProfiles_$Timestamp.csv"
    }
    try {
        $Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
        Write-Host "===================================================================" -ForegroundColor Cyan
        Write-Host "[SUCCESS] Results exported to: $ExportPath" -ForegroundColor Green
        Write-Host "===================================================================" -ForegroundColor Cyan
    } catch {
        Write-Host "[ERROR] Failed to export results: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "[INFO] No stale folders found - nothing to export" -ForegroundColor Yellow
}

$Duration    = New-TimeSpan -Start $StartTime -End (Get-Date)
$DurationStr = $Duration.ToString('mm\:ss')
Write-Host ""
Write-Host "Analysis completed in $DurationStr" -ForegroundColor Cyan
Write-Host ""
