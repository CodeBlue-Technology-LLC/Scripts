<#
.SYNOPSIS
    Migrates Outlook autocomplete cache to a new Office 365 profile.

.DESCRIPTION
    Exports the autocomplete cache (stream_autocomplete) from an existing Outlook
    profile and imports it into a new Office 365 profile using nk2edit. Configures
    Outlook registry settings for M365 autodiscover and modern authentication.

    Two modes are available depending on the migration scenario:

    System  - Run the night before as SYSTEM (e.g., via RMM). Exports the cache,
              configures registry via the user's hive, stages the NK2, and creates
              a RunOnce entry so Outlook imports it on next login.

    User    - Run onsite as admin in the user's session. Exports the cache, configures
              HKCU, closes Outlook, and relaunches it with /importnk2 so the user
              doesn't need to log off.

.PARAMETER Mode
    System  - Remote/RMM deployment the night before. User logs in fresh the next day.
    User    - Onsite migration. Outlook is closed and relaunched automatically.

.PARAMETER TargetUser
    Username to target. In System mode, defaults to the currently logged-in user
    (detected via explorer.exe). In User mode, defaults to $env:USERNAME.

.PARAMETER StagingPath
    Folder to stage exported autocomplete files. Default: C:\cbt\m365

.PARAMETER Nk2EditPath
    Path to nk2edit.exe. Default: <StagingPath>\nk2edit.exe (C:\cbt\m365\nk2edit.exe)
    If not present, the script will automatically download and extract it from NirSoft.

.PARAMETER ProfileName
    Name for the new Outlook profile. Default: "Office 365"

.EXAMPLE
    # Remote deployment as SYSTEM (e.g., via RMM tool the night before)
    .\Migrate-AutocompleteCache.ps1 -Mode System

.EXAMPLE
    # Onsite migration as admin in the user's session
    .\Migrate-AutocompleteCache.ps1 -Mode User

.EXAMPLE
    # Onsite migration targeting a specific user
    .\Migrate-AutocompleteCache.ps1 -Mode User -TargetUser jsmith

.EXAMPLE
    # Remote deployment with custom staging path and profile name
    .\Migrate-AutocompleteCache.ps1 -Mode System -StagingPath "D:\Migration\autocomplete" -ProfileName "Microsoft 365"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet("System", "User")]
    [string]$Mode,

    [string]$TargetUser,

    [string]$StagingPath = "C:\cbt\m365",

    [string]$Nk2EditPath,

    [string]$ProfileName = "Office 365"
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
    try {
        $explorer = Get-Process -Name explorer -IncludeUserName -ErrorAction SilentlyContinue |
            Select-Object -First 1
        if ($explorer) {
            return $explorer.UserName.Split('\')[-1]
        }
    } catch { }

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

    $profileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    foreach ($sid in (Get-ChildItem $profileListPath -ErrorAction SilentlyContinue)) {
        $profilePath = (Get-ItemProperty $sid.PSPath -ErrorAction SilentlyContinue).ProfileImagePath
        if ($profilePath -and (Split-Path $profilePath -Leaf) -eq $Username) {
            return $profilePath
        }
    }

    $match = Get-ChildItem "C:\Users" -Directory |
        Where-Object { $_.Name -eq $Username -or $_.Name -like "$Username.*" } |
        Select-Object -First 1
    if ($match) { return $match.FullName }

    return $null
}

function Get-UserSID {
    param([string]$Username)

    try {
        $account = New-Object System.Security.Principal.NTAccount($Username)
        return $account.Translate([System.Security.Principal.SecurityIdentifier]).Value
    } catch { }

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

function Export-AutocompleteCache {
    param(
        [string]$Username,
        [string]$Staging
    )

    Write-Host "--- Exporting autocomplete cache ---"

    $userProfile = Get-UserProfilePath -Username $Username
    if (-not $userProfile) {
        throw "Could not find a profile folder for '$Username' under C:\Users."
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
            return $null
        }

        $destFolder = Join-Path $Staging $Username
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
        return $destFolder
    } finally {
        if ($aclModified -and $accessRule) {
            Write-Host "  Removing temporary ACL..."
            $acl = Get-Acl -Path $roamCache
            $acl.RemoveAccessRule($accessRule) | Out-Null
            $acl | Set-Acl $roamCache
        }
    }
}

function Configure-OutlookRegistry {
    param(
        [string]$RegBase,
        [string]$Profile
    )

    $outlookBase  = "$RegBase\Software\Microsoft\Office\16.0\Outlook"
    $identityBase = "$RegBase\Software\Microsoft\Office\16.0\Common\Identity"

    Write-Host "  Configuring Outlook registry for profile '$Profile'..."
    Set-RegistryValue -Path "$outlookBase\AutoDiscover" -Name "ZeroConfigExchange" -Value 1
    Set-RegistryValue -Path $outlookBase -Name "DefaultProfile" -Value $Profile -Type String

    if (-not (Test-Path "$outlookBase\Profiles\$Profile")) {
        New-Item -Path "$outlookBase\Profiles\$Profile" -Force | Out-Null
    }

    Set-RegistryValue -Path $identityBase -Name "EnableADAL" -Value 1
    Set-RegistryValue -Path "$outlookBase\AutoDiscover" -Name "ExcludeExplicitO365Endpoint" -Value 0

    # Re-apply signature defaults
    $mailSettingsPath = "$RegBase\Software\Microsoft\Office\16.0\Common\MailSettings"
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
}

function Convert-AutocompleteToNk2 {
    param(
        [string]$StagedFolder,
        [string]$Nk2Destination,
        [string]$Profile
    )

    $backupCache = Get-ChildItem $StagedFolder -Recurse -Filter "stream_autocomplete*" |
        Sort-Object -Descending -Property Length |
        Select-Object -First 1

    if (-not $backupCache) {
        throw "No stream_autocomplete files found in '$StagedFolder'."
    }
    Write-Host "  Using: $($backupCache.Name) ($([math]::Round($backupCache.Length / 1KB, 1)) KB)"

    # Use space-free temp paths in the staging folder - nk2edit parses its own
    # command line and fails silently when paths contain spaces.
    $nk2TextTemp  = Join-Path $StagingPath "autocomplete_export.txt"
    $nk2BinTemp   = Join-Path $StagingPath "autocomplete_import.nk2"
    $nk2FinalFile = Join-Path $Nk2Destination "$Profile.nk2"

    if (Test-Path $nk2FinalFile) {
        Write-Host "  NK2 already exists at '$nk2FinalFile' - skipping conversion."
        return $nk2FinalFile
    }

    if (-not (Test-Path $Nk2Destination)) {
        New-Item -Path $Nk2Destination -ItemType Directory -Force | Out-Null
    }

    # nk2edit must run from its own directory
    Push-Location (Split-Path $Nk2EditPath -Parent)
    try {
        Write-Host "  Converting stream_autocomplete -> text..."
        & $Nk2EditPath /nk2_to_text $backupCache.FullName $nk2TextTemp
        if (-not (Test-Path $nk2TextTemp)) {
            throw "nk2edit /nk2_to_text failed - output file not created."
        }
        $textSize = (Get-Item $nk2TextTemp).Length
        Write-Host "  Text file size: $textSize bytes"
        if ($textSize -eq 0) {
            throw "nk2edit /nk2_to_text produced an empty file - stream_autocomplete format may not be supported."
        }

        Write-Host "  Converting text -> NK2 binary..."
        & $Nk2EditPath /text_to_nk2 $nk2TextTemp $nk2BinTemp
        if (-not (Test-Path $nk2BinTemp)) {
            throw "nk2edit /text_to_nk2 failed - output file not created."
        }
    } finally {
        Pop-Location
    }

    Copy-Item $nk2BinTemp $nk2FinalFile
    Remove-Item $nk2TextTemp, $nk2BinTemp -Force
    Write-Host "  NK2 staged: $nk2FinalFile"
    return $nk2FinalFile
}

#endregion

#region Resolve Target User

if (-not $TargetUser) {
    if ($Mode -eq "System") {
        $TargetUser = Get-LoggedInUser
        if (-not $TargetUser) {
            throw "Running in System mode but could not detect a logged-in user. Specify -TargetUser."
        }
        Write-Host "Detected logged-in user: $TargetUser"
    } else {
        $TargetUser = $env:USERNAME
    }
}

#endregion

#region System Mode

function Invoke-SystemMigration {
    Write-Host "=== System Mode: Migrating autocomplete for '$TargetUser' ==="

    if (-not (Test-RunningAsSystem)) {
        if (-not (Test-IsAdmin)) {
            throw "System mode requires running as SYSTEM or at least as Administrator."
        }
        Write-Warning "Not running as SYSTEM - proceeding as Administrator. Hive loading may fail if the user is logged in."
    }

    # Export
    $stagedFolder = Export-AutocompleteCache -Username $TargetUser -Staging $StagingPath
    if (-not $stagedFolder) { return }

    # Load user registry hive
    Write-Host "--- Configuring registry via user hive ---"
    $userSID = Get-UserSID -Username $TargetUser
    if (-not $userSID) {
        throw "Could not determine SID for '$TargetUser'."
    }

    $hiveLoaded = $false
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

    try {
        Configure-OutlookRegistry -RegBase $regBase -Profile $ProfileName

        # Convert and stage NK2
        Write-Host "--- Converting autocomplete to NK2 ---"
        $userProfile = Get-UserProfilePath -Username $TargetUser
        $nk2Dir = Join-Path $userProfile "AppData\Roaming\Microsoft\Outlook"
        Convert-AutocompleteToNk2 -StagedFolder $stagedFolder -Nk2Destination $nk2Dir -Profile $ProfileName

        # Schedule import for next login via RunOnce
        Write-Host "--- Scheduling NK2 import for next login ---"
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
    } finally {
        if ($hiveLoaded) {
            Write-Host "  Unloading user registry hive..."
            [gc]::Collect()
            [gc]::WaitForPendingFinalizers()
            Start-Sleep -Seconds 2
            reg unload "HKU\$userSID" 2>&1 | Out-Null
        }
    }
}

#endregion

#region User Mode

function Invoke-UserMigration {
    Write-Host "=== User Mode: Migrating autocomplete for '$TargetUser' ==="

    # Export
    $stagedFolder = Export-AutocompleteCache -Username $TargetUser -Staging $StagingPath
    if (-not $stagedFolder) { return }

    # Configure registry via HKCU
    Write-Host "--- Configuring registry ---"
    Configure-OutlookRegistry -RegBase "HKCU:" -Profile $ProfileName

    # Convert and stage NK2
    Write-Host "--- Converting autocomplete to NK2 ---"
    $nk2Dir = Join-Path $env:APPDATA "Microsoft\Outlook"
    Convert-AutocompleteToNk2 -StagedFolder $stagedFolder -Nk2Destination $nk2Dir -Profile $ProfileName

    # Close Outlook and relaunch with NK2 import
    Write-Host "--- Launching Outlook with NK2 import ---"
    $outlookExe = Get-OutlookExePath
    if (-not $outlookExe) {
        throw "Could not find Outlook executable. Install Outlook or specify the path manually."
    }

    $outlookProc = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
    if ($outlookProc) {
        Write-Host "  Closing Outlook..."
        $outlookProc | ForEach-Object { $_.CloseMainWindow() | Out-Null }
        Start-Sleep -Seconds 3
        # Force-kill if still running
        $outlookProc = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
        if ($outlookProc) {
            Write-Host "  Outlook didn't close gracefully - forcing..."
            $outlookProc | Stop-Process -Force
            Start-Sleep -Seconds 2
        }
    }

    Write-Host "  Launching: $outlookExe /importnk2 /profile `"$ProfileName`""
    Start-Process -FilePath $outlookExe -ArgumentList "/importnk2", "/profile", "`"$ProfileName`""
    Write-Host "  Outlook launched. Autocomplete cache will be imported."
}

#endregion

#region Main

switch ($Mode) {
    "System" { Invoke-SystemMigration }
    "User"   { Invoke-UserMigration }
}

Write-Host "`nDone. Migration ($Mode mode) complete for '$TargetUser'."

#endregion
