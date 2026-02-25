# Quick Run - Stale Profile Analysis
# This is a simplified version for immediate use

Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Quick Profile Cleanup Helper" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# Prompt for directory
$RootPath = Read-Host "Enter the root path to scan (e.g., C:\Shares\Users)"

# Validate path
if (-not (Test-Path $RootPath -PathType Container)) {
    Write-Host "[✗] Error: Path does not exist or is not a directory" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit
}

# Prompt for days
Write-Host ""
Write-Host "How old should profiles be to flag them?" -ForegroundColor Yellow
Write-Host "  [1] 90 days (3 months)" -ForegroundColor White
Write-Host "  [2] 180 days (6 months)" -ForegroundColor White
Write-Host "  [3] 365 days (1 year)" -ForegroundColor Green
Write-Host "  [4] 730 days (2 years)" -ForegroundColor Green
Write-Host "  [5] Custom" -ForegroundColor White
Write-Host ""

$Choice = Read-Host "Select option (1-5)"

switch ($Choice) {
    "1" { $DaysInactive = 90 }
    "2" { $DaysInactive = 180 }
    "3" { $DaysInactive = 365 }
    "4" { $DaysInactive = 730 }
    "5" { 
        $DaysInactive = Read-Host "Enter number of days"
        if ($DaysInactive -notmatch '^\d+$' -or [int]$DaysInactive -lt 1) {
            Write-Host "[✗] Invalid number. Using 365 days as default." -ForegroundColor Yellow
            $DaysInactive = 365
        }
    }
    default {
        Write-Host "[i] Invalid choice. Using 365 days as default." -ForegroundColor Yellow
        $DaysInactive = 365
    }
}

Write-Host ""
Write-Host "Starting analysis..." -ForegroundColor Green
Write-Host ""

# Check if the full script exists in the same directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$MainScript = Join-Path $ScriptDir "Analyze-StaleProfiles-NoModule.ps1"

if (Test-Path $MainScript) {
    # Run the main script
    & $MainScript -RootPath $RootPath -DaysInactive $DaysInactive
} else {
    Write-Host "[✗] Error: Could not find Analyze-StaleProfiles-NoModule.ps1" -ForegroundColor Red
    Write-Host "    Make sure both scripts are in the same directory." -ForegroundColor Yellow
}

Write-Host ""
Read-Host "Press Enter to exit"
