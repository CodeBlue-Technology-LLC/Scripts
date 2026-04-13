<#
.SYNOPSIS
    Migrates Outlook autocomplete cache to a new Office 365 profile.

.DESCRIPTION
    Exports the autocomplete cache (stream_autocomplete) from an existing Outlook
    profile and imports it into a new Office 365 profile using nk2edit. Configures
    Outlook registry settings for M365 autodiscover and modern authentication.

    Designed to run on individual workstations as either SYSTEM (for scripted/RMM
    deployments) or under the target user's profile.

    When running as SYSTEM:
      - Export: accesses the user's RoamCache directly (no ACL hacks needed)
      - Import: loads the user's registry hive, stages the NK2, and creates a
        RunOnce entry so Outlook imports it on next login.

    When running as the user:
      - Export: may require elevation if the user's own RoamCache has restrictive ACLs
      - Import: writes to HKCU directly and launches Outlook to import the NK2

.PARAMETER Mode
    Export  - Backs up autocomplete cache from the user's RoamCache to a staging folder
    Import  - Creates the Office 365 profile and imports the autocomplete cache via nk2edit
    Full    - Runs Export then Import sequentially

.PARAMETER TargetUser
    Username to target. Defaults to the currently logged-in user (detected via explorer.exe
    when running as SYSTEM, or $env:USERNAME otherwise).

.PARAMETER StagingPath
    Folder to stage exported autocomplete files. Default: C:\cbt\m365

.PARAMETER Nk2EditPath
    Path to nk2edit.exe. Default: <StagingPath>\nk2edit.exe (C:\cbt\m365\nk2edit.exe)
    If not present, the script will automatically download and extract it from NirSoft.

.PARAMETER ProfileName
    Name for the new Outlook profile. Default: "Office 365"

.PARAMETER OutlookWaitSeconds
    Seconds to wait for Outlook to initialize before closing it. Default: 30

.EXAMPLE
    # Full migration as SYSTEM (e.g., via RMM tool)
    .\Migrate-AutocompleteCache.ps1 -Mode Full

.EXAMPLE
    # Export only, targeting a specific user
    .\Migrate-AutocompleteCache.ps1 -Mode Export -TargetUser jsmith

.EXAMPLE
    # Import only, running as the logged-in user
    .\Migrate-AutocompleteCache.ps1 -Mode Import

.EXAMPLE
    # Full migration with custom staging path and profile name
    .\Migrate-AutocompleteCache.ps1 -Mode Full -StagingPath "D:\Migration\autocomplete" -ProfileName "Microsoft 365"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet("Export", "Import", "Full")]
    [string]$Mode,

    [string]$TargetUser,

    [string]$StagingPath = "C:\cbt\m365",

    [string]$Nk2EditPath,

    [string]$ProfileName = "Office 365",

    [int]$OutlookWaitSeconds = 30
)

$ErrorActionPreference = "Stop"

if (-not $Nk2EditPath) {
    $Nk2EditPath = Join-Path $StagingPath "nk2edit.exe"
}

# Ensure staging folder exists
if (-not (Test-Path $StagingPath)) {
    New-Item -Path $StagingPath -ItemType Directory -Force | Out-Null
    Write-Host "Created staging folder: $StagingPath"
}

# Download nk2edit if not present
if (-not (Test-Path $Nk2EditPath)) {
    Write-Host "nk2edit.exe not found - downloading from NirSoft..."
    $nk2ZipUrl  = "https://www.nirsoft.net/utils/nk2edit-32-64.zip"
    $nk2ZipFile = Join-Path $StagingPath "nk2edit.zip"
    try {
        Invoke-WebRequest -Uri $nk2ZipUrl -OutFile $nk2ZipFile -UseBasicParsing
        Expand-Archive -Path $nk2ZipFile -DestinationPath $StagingPath -Force
        Remove-Item $nk2ZipFile -Force
        if (-not (Test-Path $Nk2EditPath)) {
            throw "nk2edit.exe was not found in the archive after extraction."
        }
        Write-Host "  nk2edit.exe downloaded and extracted to: $Nk2EditPath"
    } catch {
        throw "Failed to download/extract nk2edit.exe: $_"
    }
}

#region Helper Functions

function Test-RunningAsSystem {
    return ([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value -eq "S-1-5-18")
}

function Test-IsAdmin {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-LoggedInUser {
    # Try explorer process first (most reliable for interactive sessions)
    try {
        $explorer = Get-Process -Name explorer -IncludeUserName -ErrorAction SilentlyContinue |
            Select-Object -First 1
        if ($explorer) {
            return $explorer.UserName.Split('\')[-1]
        }
    } catch { }

    # Fallback to WMI
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem
        if ($cs.UserName) {
            return $cs.UserName.Split('\')[-1]
        }
    } catch { }

    return $null
}

function Get-UserProfilePath {
    param([string]$Username)

    # Check ProfileList registry for exact match on leaf folder name
    $profileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    foreach ($sid in (Get-ChildItem $profileListPath -ErrorAction SilentlyContinue)) {
        $profilePath = (Get-ItemProperty $sid.PSPath -ErrorAction SilentlyContinue).ProfileImagePath
        if ($profilePath -and (Split-Path $profilePath -Leaf) -eq $Username) {
            return $profilePath
        }
    }

    # Fallback: look for folder with domain suffix (e.g., jsmith.DOMAIN)
    $match = Get-ChildItem "C:\Users" -Directory |
        Where-Object { $_.Name -eq $Username -or $_.Name -like "$Username.*" } |
        Select-Object -First 1
    if ($match) { return $match.FullName }

    return $null
}

function Get-UserSID {
    param([string]$Username)

    # Try translating the account name directly
    try {
        $account = New-Object System.Security.Principal.NTAccount($Username)
        return $account.Translate([System.Security.Principal.SecurityIdentifier]).Value
    } catch { }

    # Fallback: search ProfileList by profile path leaf
    $profileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    foreach ($sid in (Get-ChildItem $profileListPath -ErrorAction SilentlyContinue)) {
        $profilePath = (Get-ItemProperty $sid.PSPath -ErrorAction SilentlyContinue).ProfileImagePath
        if ($profilePath -and (Split-Path $profilePath -Leaf) -like "$Username*") {
            return (Split-Path $sid.PSPath -Leaf)
        }
    }

    return $null
}

function Get-OutlookExePath {
    $appPath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE" -ErrorAction SilentlyContinue).'(Default)'
    if ($appPath -and (Test-Path $appPath)) { return $appPath }
    return $null
}

function Set-RegistryValue {
    param(
        [string]$Path,
        [string]$Name,
        $Value,
        [Microsoft.Win32.RegistryValueKind]$Type = [Microsoft.Win32.RegistryValueKind]::DWord
    )

    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }
    Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type
}

#endregion

#region Resolve Target User

if (-not $TargetUser) {
    if (Test-RunningAsSystem) {
        $TargetUser = Get-LoggedInUser
        if (-not $TargetUser) {
            throw "Running as SYSTEM but could not detect a logged-in user. Specify -TargetUser."
        }
        Write-Host "Detected logged-in user: $TargetUser"
    } else {
        $TargetUser = $env:USERNAME
    }
}

#endregion

#region Export

function Export-AutocompleteCache {
    Write-Host "=== Exporting autocomplete cache for '$TargetUser' ==="

    if (-not (Test-IsAdmin)) {
        throw "Export mode requires administrator privileges. Run elevated or as SYSTEM."
    }

    $userProfile = Get-UserProfilePath -Username $TargetUser
    if (-not $userProfile) {
        throw "Could not find a profile folder for '$TargetUser' under C:\Users."
    }
    Write-Host "  User profile path: $userProfile"

    $roamCache = Join-Path $userProfile "AppData\Local\Microsoft\Outlook\RoamCache"
    if (-not (Test-Path $roamCache)) {
        throw "RoamCache not found at '$roamCache'. The user may not have an existing Outlook profile."
    }

    # Grant ourselves access if needed, and track it for cleanup
    $aclModified = $false
    $currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $accessRule = $null
    try {
        $null = Get-ChildItem $roamCache -ErrorAction Stop
    } catch {
        Write-Host "  Granting temporary access to RoamCache..."
        $acl = Get-Acl -Path $roamCache
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $currentIdentity, "FullControl",
            "ContainerInherit,ObjectInherit", "None", "Allow"
        )
        $acl.SetAccessRule($accessRule)
        $acl | Set-Acl $roamCache
        $aclModified = $true
    }

    try {
        $streamFiles = Get-ChildItem $roamCache -Filter "stream_autocomplete*" -ErrorAction SilentlyContinue
        if (-not $streamFiles) {
            Write-Warning "No stream_autocomplete files found in RoamCache. Nothing to export."
            return
        }

        $destFolder = Join-Path $StagingPath $TargetUser
        if (-not (Test-Path $destFolder)) {
            New-Item -Path $destFolder -ItemType Directory -Force | Out-Null
        }

        foreach ($file in $streamFiles) {
            $destFile = Join-Path $destFolder $file.Name
            if (Test-Path $destFile) {
                Write-Host "  Skipping (already staged): $($file.Name)"
            } else {
                Copy-Item $file.FullName -Destination $destFolder
                Write-Host "  Exported: $($file.Name) ($([math]::Round($file.Length / 1KB, 1)) KB)"
            }
        }

        Write-Host "  Export complete -> $destFolder"
    } finally {
        if ($aclModified -and $accessRule) {
            Write-Host "  Removing temporary ACL..."
            $acl = Get-Acl -Path $roamCache
            $acl.RemoveAccessRule($accessRule) | Out-Null
            $acl | Set-Acl $roamCache
        }
    }
}

#endregion

#region Import

function Import-AutocompleteCache {
    Write-Host "=== Importing autocomplete cache for '$TargetUser' ==="

    $userStaging = Join-Path $StagingPath $TargetUser
    if (-not (Test-Path $userStaging)) {
        throw "No staged files at '$userStaging'. Run with -Mode Export first."
    }

    # Pick the largest stream_autocomplete file (most complete cache)
    $backupCache = Get-ChildItem $userStaging -Recurse -Filter "stream_autocomplete*" |
        Sort-Object -Descending -Property Length |
        Select-Object -First 1

    if (-not $backupCache) {
        throw "No stream_autocomplete files found in '$userStaging'."
    }
    Write-Host "  Using: $($backupCache.Name) ($([math]::Round($backupCache.Length / 1KB, 1)) KB)"

    # Determine registry base path - HKCU for user context, HKU\<SID> for SYSTEM
    $isSystem = Test-RunningAsSystem
    $hiveLoaded = $false
    $regBase = "HKCU:"
    $userSID = $null

    if ($isSystem) {
        Write-Host "  Running as SYSTEM - resolving user registry hive..."

        $userSID = Get-UserSID -Username $TargetUser
        if (-not $userSID) {
            throw "Could not determine SID for '$TargetUser'."
        }

        if (Test-Path "Registry::HKU\$userSID") {
            Write-Host "  User hive already loaded (user is logged in)."
            $regBase = "Registry::HKU\$userSID"
        } else {
            $userProfile = Get-UserProfilePath -Username $TargetUser
            $ntUserDat = Join-Path $userProfile "NTUSER.DAT"
            if (-not (Test-Path $ntUserDat)) {
                throw "NTUSER.DAT not found at '$ntUserDat'."
            }
            $regLoadResult = reg load "HKU\$userSID" $ntUserDat 2>&1
            if ($LASTEXITCODE -ne 0) {
                throw "Failed to load registry hive: $regLoadResult"
            }
            $hiveLoaded = $true
            $regBase = "Registry::HKU\$userSID"
            Write-Host "  Loaded hive from: $ntUserDat"
        }
    }

    try {
        $outlookBase = "$regBase\Software\Microsoft\Office\16.0\Outlook"
        $identityBase = "$regBase\Software\Microsoft\Office\16.0\Common\Identity"

        # Configure Outlook profile
        Write-Host "  Configuring Outlook registry for profile '$ProfileName'..."
        Set-RegistryValue -Path "$outlookBase\AutoDiscover" -Name "ZeroConfigExchange" -Value 1
        Set-RegistryValue -Path $outlookBase -Name "DefaultProfile" -Value $ProfileName -Type String

        # Ensure profile key exists
        if (-not (Test-Path "$outlookBase\Profiles\$ProfileName")) {
            New-Item -Path "$outlookBase\Profiles\$ProfileName" -Force | Out-Null
        }

        # Enable modern authentication
        Set-RegistryValue -Path $identityBase -Name "EnableADAL" -Value 1

        # Allow Office 365 autodiscover endpoint
        Set-RegistryValue -Path "$outlookBase\AutoDiscover" -Name "ExcludeExplicitO365Endpoint" -Value 0

        # Re-apply signature defaults - Outlook resets these per-account when a new profile is
        # created, so we read the existing values and write them back after profile setup to
        # ensure they stick as global fallback defaults.
        $mailSettingsPath = "$regBase\Software\Microsoft\Office\16.0\Common\MailSettings"
        $existingSettings = Get-ItemProperty -Path $mailSettingsPath -ErrorAction SilentlyContinue
        $newSig   = $existingSettings.NewSignature
        $replySig = $existingSettings.ReplySignature

        if ($newSig -or $replySig) {
            Write-Host "  Re-applying signature defaults..."
            if ($newSig) {
                Set-RegistryValue -Path $mailSettingsPath -Name "NewSignature" -Value $newSig -Type String
                Write-Host "    New email  : $newSig"
            }
            if ($replySig) {
                Set-RegistryValue -Path $mailSettingsPath -Name "ReplySignature" -Value $replySig -Type String
                Write-Host "    Reply/forward: $replySig"
            }
        } else {
            Write-Host "  No signature defaults found in MailSettings - skipping."
        }

        Write-Host "  Registry configured."

        # Convert stream_autocomplete -> nk2 via nk2edit
        $nk2TextFile = Join-Path $backupCache.DirectoryName "$ProfileName.nk2"

        if ($isSystem) {
            $userProfile = Get-UserProfilePath -Username $TargetUser
            $nk2FinalDir = Join-Path $userProfile "AppData\Roaming\Microsoft\Outlook"
        } else {
            $nk2FinalDir = Join-Path $env:APPDATA "Microsoft\Outlook"
        }
        $nk2FinalFile = Join-Path $nk2FinalDir "$ProfileName.nk2"

        if (Test-Path $nk2FinalFile) {
            Write-Host "  NK2 already exists at '$nk2FinalFile' - skipping conversion."
        } else {
            if (-not (Test-Path $nk2FinalDir)) {
                New-Item -Path $nk2FinalDir -ItemType Directory -Force | Out-Null
            }

            Write-Host "  Converting stream_autocomplete -> text..."
            & $Nk2EditPath /nk2_to_text $backupCache.FullName $nk2TextFile
            if ($LASTEXITCODE -ne 0) {
                throw "nk2edit /nk2_to_text failed (exit code $LASTEXITCODE)."
            }

            Write-Host "  Converting text -> NK2 binary..."
            & $Nk2EditPath /text_to_nk2 $nk2TextFile $nk2FinalFile
            if ($LASTEXITCODE -ne 0) {
                throw "nk2edit /text_to_nk2 failed (exit code $LASTEXITCODE)."
            }

            Write-Host "  NK2 staged: $nk2FinalFile"
        }

        # Launch Outlook or schedule import for next login
        if ($isSystem) {
            Write-Host "  Running as SYSTEM - scheduling NK2 import for next user login..."

            $outlookExe = Get-OutlookExePath
            if ($outlookExe) {
                $runOnceKey = "$regBase\Software\Microsoft\Windows\CurrentVersion\RunOnce"
                $importCmd = "`"$outlookExe`" /importnk2 /profile `"$ProfileName`""
                Set-RegistryValue -Path $runOnceKey -Name "OutlookNK2Import" -Value $importCmd -Type String
                Write-Host "  RunOnce entry created. Outlook will import the NK2 on next login."
            } else {
                Write-Warning "Could not find Outlook executable. The user will need to run Outlook manually with:"
                Write-Warning "  outlook.exe /importnk2 /profile `"$ProfileName`""
            }
        } else {
            $outlookExe = Get-OutlookExePath
            if (-not $outlookExe) {
                Write-Warning "Could not find Outlook. Run it manually with: outlook.exe /importnk2 /profile `"$ProfileName`""
                return
            }

            Write-Host "  Launching Outlook to import NK2 (waiting $OutlookWaitSeconds seconds)..."
            $proc = Start-Process -FilePath $outlookExe `
                -ArgumentList "/importnk2", "/profile `"$ProfileName`"" -PassThru

            Start-Sleep -Seconds $OutlookWaitSeconds

            if (-not $proc.HasExited) {
                $proc.Kill()
                Write-Host "  Outlook closed after NK2 import."
            } else {
                Write-Host "  Outlook exited on its own."
            }
        }

        Write-Host "  Import complete."
    } finally {
        if ($hiveLoaded -and $userSID) {
            Write-Host "  Unloading user registry hive..."
            [gc]::Collect()
            [gc]::WaitForPendingFinalizers()
            Start-Sleep -Seconds 2
            reg unload "HKU\$userSID" 2>&1 | Out-Null
        }
    }
}

#endregion

#region Main

switch ($Mode) {
    "Export" {
        Export-AutocompleteCache
    }
    "Import" {
        Import-AutocompleteCache
    }
    "Full" {
        Export-AutocompleteCache
        Import-AutocompleteCache
    }
}

Write-Host "`nDone. Migration ($Mode) complete for '$TargetUser'."

#endregion
