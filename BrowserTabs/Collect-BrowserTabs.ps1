<#
    Collect-BrowserTabs.ps1
    One-time collector: walks every user profile on an RDS host and stages the
    browser session/history files needed to reconstruct open tabs.

    Run in an ELEVATED PowerShell session on the RDS host (Local Admin).
    Then run parse_tabs.py against the -Staging folder to get a CSV.

    Firefox recovery.jsonlz4  -> reliable "current tabs"
    Chrome/Edge Current Tabs  -> best-effort "current tabs" (SNSS, undocumented)
    Chrome/Edge History       -> reliable visited-URL fallback

    Files are copied via a Volume Shadow Copy snapshot so LOCKED files
    (open browsers) can still be read. Falls back to a raw copy if VSS fails.

    NOTE: Employee monitoring is usually a policy/legal matter. Make sure this
    pull is authorized and disclosed before running it.
#>

[CmdletBinding()]
param(
    [string]$Staging = "C:\Temp\browser_pull",
    [string]$UsersRoot = "C:\Users"
)

$ErrorActionPreference = "Continue"

function New-Snapshot {
    param([string]$Drive = "C:\")
    try {
        $r = (Get-WmiObject -List Win32_ShadowCopy).Create($Drive, "ClientAccessible")
        if ($r.ReturnValue -ne 0) { return $null }
        $sc = Get-WmiObject Win32_ShadowCopy | Where-Object { $_.ID -eq $r.ShadowID }
        return $sc  # has .DeviceObject like \\?\GLOBALROOT\Device\HarddiskVolumeShadowCopyN
    } catch { return $null }
}

function Copy-FromSnapshot {
    param([string]$Snapshot, [string]$SourceFull, [string]$Dest)
    # SourceFull like C:\Users\jdoe\...  -> map to snapshot device path
    if ($Snapshot) {
        $rel = $SourceFull.Substring(2)                     # strip "C:"
        $src = ($Snapshot.TrimEnd('\')) + $rel              # device + \Users\...
    } else {
        $src = $SourceFull
    }
    try {
        Copy-Item -LiteralPath $src -Destination $Dest -Force -ErrorAction Stop
        return $true
    } catch {
        # last resort: raw copy of the live (possibly locked) file
        try { Copy-Item -LiteralPath $SourceFull -Destination $Dest -Force -ErrorAction Stop; return $true }
        catch { Write-Warning "  copy failed: $SourceFull -> $($_.Exception.Message)"; return $false }
    }
}

# --- setup ---
New-Item -ItemType Directory -Force -Path $Staging | Out-Null
$snap = New-Snapshot -Drive "C:\"
if ($snap) { Write-Host "VSS snapshot: $($snap.DeviceObject)" -ForegroundColor Green }
else       { Write-Host "VSS unavailable - using raw copy (locked files may be skipped)" -ForegroundColor Yellow }
$snapDev = if ($snap) { $snap.DeviceObject } else { $null }

$profiles = Get-ChildItem -LiteralPath $UsersRoot -Directory |
    Where-Object { $_.Name -notin @("Public","Default","Default User","All Users") }

foreach ($p in $profiles) {
    $user = $p.Name
    Write-Host "`n=== $user ===" -ForegroundColor Cyan
    $appData  = Join-Path $p.FullName "AppData\Roaming"
    $localApp = Join-Path $p.FullName "AppData\Local"

    # ---------- Firefox ----------
    $ffRoot = Join-Path $appData "Mozilla\Firefox\Profiles"
    if (Test-Path $ffRoot) {
        $ffDest = Join-Path $Staging "$user\firefox"
        New-Item -ItemType Directory -Force -Path $ffDest | Out-Null
        Get-ChildItem -LiteralPath $ffRoot -Directory | ForEach-Object {
            $prof = $_.Name
            foreach ($rel in @("sessionstore-backups\recovery.jsonlz4",
                               "sessionstore-backups\recovery.baklz4",
                               "sessionstore.jsonlz4")) {
                $srcFull = Join-Path $_.FullName $rel
                if (Test-Path $srcFull) {
                    $leaf = ($rel -replace '.*\\','')
                    $dest = Join-Path $ffDest "${prof}_$leaf"
                    if (Copy-FromSnapshot $snapDev $srcFull $dest) { Write-Host "  FF  $prof\$leaf" }
                }
            }
        }
    }

    # ---------- Chrome & Edge ----------
    $chromiumTargets = @{
        "chrome" = Join-Path $localApp "Google\Chrome\User Data"
        "edge"   = Join-Path $localApp "Microsoft\Edge\User Data"
    }
    foreach ($br in $chromiumTargets.Keys) {
        $udRoot = $chromiumTargets[$br]
        if (-not (Test-Path $udRoot)) { continue }
        $brDest = Join-Path $Staging "$user\$br"
        # each Chromium profile folder: Default, Profile 1, etc.
        Get-ChildItem -LiteralPath $udRoot -Directory |
            Where-Object { $_.Name -eq "Default" -or $_.Name -like "Profile *" } |
            ForEach-Object {
                $prof = $_.Name
                $pDest = Join-Path $brDest ($prof -replace '\s','_')
                New-Item -ItemType Directory -Force -Path $pDest | Out-Null
                # current tabs / session (SNSS) live under \Sessions\ in newer builds,
                # or as loose files in older ones - grab both.
                $sessDir = Join-Path $_.FullName "Sessions"
                if (Test-Path $sessDir) {
                    Get-ChildItem -LiteralPath $sessDir -File |
                        Where-Object { $_.Name -like "Session_*" -or $_.Name -like "Tabs_*" } |
                        Sort-Object LastWriteTime -Descending | Select-Object -First 4 |
                        ForEach-Object { Copy-FromSnapshot $snapDev $_.FullName (Join-Path $pDest $_.Name) | Out-Null }
                }
                foreach ($loose in @("Current Session","Current Tabs","Last Session","Last Tabs")) {
                    $srcFull = Join-Path $_.FullName $loose
                    if (Test-Path $srcFull) { Copy-FromSnapshot $snapDev $srcFull (Join-Path $pDest $loose) | Out-Null }
                }
                # History DB (+ wal) as reliable fallback
                foreach ($h in @("History","History-wal")) {
                    $srcFull = Join-Path $_.FullName $h
                    if (Test-Path $srcFull) { Copy-FromSnapshot $snapDev $srcFull (Join-Path $pDest $h) | Out-Null }
                }
                Write-Host "  $($br.ToUpper())  $prof"
            }
    }
}

# --- flatten multi-profile Chromium so the parser sees one dir per browser ---
# (parser reads <user>\chrome\{Current Tabs,History}; if multiple Chromium
#  profiles exist they are under <user>\chrome\<Profile>\ - copy Default up)
Get-ChildItem -LiteralPath $Staging -Directory | ForEach-Object {
    foreach ($br in @("chrome","edge")) {
        $brDir = Join-Path $_.FullName $br
        $def   = Join-Path $brDir "Default"
        if ((Test-Path $def) -and -not (Test-Path (Join-Path $brDir "History"))) {
            Get-ChildItem -LiteralPath $def -File | ForEach-Object {
                Copy-Item $_.FullName (Join-Path $brDir $_.Name) -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

if ($snap) { try { $snap.Delete() } catch {} }
Write-Host "`nDone. Staged under: $Staging" -ForegroundColor Green
Write-Host "Next: python parse_tabs.py --input `"$Staging`" --output browser_tabs.csv" -ForegroundColor Green