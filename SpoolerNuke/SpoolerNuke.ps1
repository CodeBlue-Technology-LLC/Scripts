#Requires -RunAsAdministrator
<# 
    Spooler Nuclear Reset
    - Stops spooler + printfilterpipelinesvc
    - Clears spool queue
    - Removes all third-party drivers from registry (keeps Microsoft/inbox)
    - Clears driver files from x64\3
    - Clears print processors (keeps winprint)
    - Clears port monitors (keeps inbox Windows ones)
    - Restarts spooler
#>

$ErrorActionPreference = "Continue"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = "C:\Temp\SpoolerReset_$timestamp.log"

# Inbox Microsoft monitors to keep
$keepMonitors = @(
    "Local Port",
    "Standard TCP/IP Port",
    "USB Monitor",
    "WSD Port"
)

# Inbox Microsoft drivers to keep (partial match)
$keepDrivers = @(
    "Microsoft",
    "Remote Desktop Easy Print",
    "Microsoft enhanced Point and Print compatibility driver"
)

function Write-Log {
    param([string]$msg)
    $line = "$(Get-Date -Format 'HH:mm:ss') - $msg"
    Write-Host $line
    Add-Content -Path $logFile -Value $line
}

# Create log dir
if (!(Test-Path "C:\Temp")) { New-Item -ItemType Directory -Path "C:\Temp" | Out-Null }

Write-Log "=== Spooler Nuclear Reset Started ==="

# --- Stop Services ---
Write-Log "Stopping PrintFilterPipelineSvc..."
Stop-Service -Name "PrintFilterPipelineSvc" -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

Write-Log "Stopping Spooler..."
Stop-Service -Name Spooler -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3

# Confirm dead
$svc = Get-Service -Name Spooler
if ($svc.Status -ne "Stopped") {
    Write-Log "WARNING: Spooler did not stop cleanly - attempting taskkill"
    taskkill /F /IM spoolsv.exe 2>$null
    Start-Sleep -Seconds 2
}
Write-Log "Spooler status: $((Get-Service Spooler).Status)"

# --- Clear Spool Queue ---
Write-Log "Clearing spool queue..."
$spoolPath = "C:\Windows\System32\spool\PRINTERS"
$queueFiles = Get-ChildItem -Path $spoolPath -ErrorAction SilentlyContinue
foreach ($f in $queueFiles) {
    try {
        Remove-Item $f.FullName -Force
        Write-Log "  Deleted queue file: $($f.Name)"
    } catch {
        Write-Log "  FAILED to delete: $($f.Name) - $_"
    }
}

# --- Registry: Remove Third-Party Drivers ---
Write-Log "Cleaning driver registry entries..."
$driverRegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-3"
$drivers = Get-ChildItem -Path $driverRegPath -ErrorAction SilentlyContinue

foreach ($driver in $drivers) {
    $name = $driver.PSChildName
    $keep = $false
    foreach ($k in $keepDrivers) {
        if ($name -like "*$k*") { $keep = $true; break }
    }
    if ($keep) {
        Write-Log "  KEEPING driver: $name"
    } else {
        try {
            Remove-Item -Path $driver.PSPath -Recurse -Force
            Write-Log "  REMOVED driver: $name"
        } catch {
            Write-Log "  FAILED to remove driver: $name - $_"
        }
    }
}

# --- Registry: Remove Third-Party Print Processors ---
Write-Log "Cleaning print processor registry entries..."
$ppRegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Print Processors"
$processors = Get-ChildItem -Path $ppRegPath -ErrorAction SilentlyContinue

foreach ($pp in $processors) {
    $name = $pp.PSChildName
    if ($name -eq "winprint") {
        Write-Log "  KEEPING print processor: $name"
    } else {
        try {
            Remove-Item -Path $pp.PSPath -Recurse -Force
            Write-Log "  REMOVED print processor: $name"
        } catch {
            Write-Log "  FAILED to remove print processor: $name - $_"
        }
    }
}

# --- Registry: Remove Third-Party Port Monitors ---
Write-Log "Cleaning port monitor registry entries..."
$monitorRegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors"
$monitors = Get-ChildItem -Path $monitorRegPath -ErrorAction SilentlyContinue

foreach ($mon in $monitors) {
    $name = $mon.PSChildName
    if ($keepMonitors -contains $name) {
        Write-Log "  KEEPING monitor: $name"
    } else {
        try {
            Remove-Item -Path $mon.PSPath -Recurse -Force
            Write-Log "  REMOVED monitor: $name"
        } catch {
            Write-Log "  FAILED to remove monitor: $name - $_"
        }
    }
}

# --- Clear Driver Files ---
Write-Log "Clearing driver files from x64\3..."
$driverFilePath = "C:\Windows\System32\spool\DRIVERS\x64\3"
$driverFiles = Get-ChildItem -Path $driverFilePath -ErrorAction SilentlyContinue
foreach ($f in $driverFiles) {
    try {
        Remove-Item $f.FullName -Force
        Write-Log "  Deleted: $($f.Name)"
    } catch {
        Write-Log "  FAILED to delete: $($f.Name) - $_"
    }
}

# --- Start Spooler ---
Write-Log "Starting Spooler..."
Start-Sleep -Seconds 2
Start-Service -Name Spooler
Start-Sleep -Seconds 5

$finalStatus = (Get-Service Spooler).Status
Write-Log "Spooler final status: $finalStatus"

if ($finalStatus -eq "Running") {
    Write-Log "=== SUCCESS - Spooler is running ==="
} else {
    Write-Log "=== FAILED - Spooler did not start, check event logs ==="
}

Write-Log "Log saved to: $logFile"
Write-Host "`nDone. Log at $logFile" -ForegroundColor Cyan
