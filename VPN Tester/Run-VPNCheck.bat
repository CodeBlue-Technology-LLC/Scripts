@echo off
REM ---------------------------------------------------------------------
REM  VPN Pre-Flight Check - double-click launcher
REM  Keep this file in the same folder as VPNCheck.ps1
REM  No install, no admin rights, no execution-policy change needed.
REM ---------------------------------------------------------------------

cd /d "%~dp0"

if not exist "%~dp0VPNCheck.ps1" (
    echo.
    echo   ERROR: VPNCheck.ps1 not found in this folder.
    echo   Keep Run-VPNCheck.bat and VPNCheck.ps1 together.
    echo.
    pause
    exit /b 1
)

REM Prefer PowerShell 7 if it's installed, otherwise fall back to Windows PowerShell 5.1
where pwsh.exe >nul 2>&1
if %errorlevel%==0 (
    pwsh.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0VPNCheck.ps1" %*
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0VPNCheck.ps1" %*
)

if %errorlevel% neq 0 (
    echo.
    echo   Script exited with code %errorlevel%
    pause
)
