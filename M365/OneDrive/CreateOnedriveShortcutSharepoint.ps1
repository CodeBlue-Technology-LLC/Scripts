param(
    [Parameter(Mandatory=$false)]
    [string[]]$UserEmails,

    [Parameter(Mandatory=$false)]
    [string]$CsvPath,

    [switch]$AutoMap,

    [switch]$Account
)

# Build unified user list
$users = @()
if ($Account) {
    # -Account flag defers user discovery until after Graph connects
} elseif ($UserEmails) {
    $users += $UserEmails
}
if ($CsvPath) {
    if (-not (Test-Path $CsvPath)) {
        Write-Host "CSV file not found: $CsvPath" -ForegroundColor Red
        exit 1
    }
    $csvData = Import-Csv -Path $CsvPath
    $users += $csvData.Email | Where-Object { $_ }
}
if (-not $Account -and $users.Count -eq 0) {
    $inputEmail = Read-Host "Enter user email address"
    if (-not $inputEmail) {
        Write-Host "No email address provided. Exiting." -ForegroundColor Red
        exit 1
    }
    $users = @($inputEmail)
}

if (-not $Account) {
    Write-Host "Processing $($users.Count) user(s)..." -ForegroundColor Cyan
}

# Check and install required modules if not installed
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Sites", "Microsoft.Graph.Users", "Microsoft.Online.SharePoint.PowerShell")
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing $module..." -ForegroundColor Yellow
        Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
    }
    # Import the module if not already loaded
    if (-not (Get-Module -Name $module)) {
        if ($module -eq "Microsoft.Online.SharePoint.PowerShell" -and $PSVersionTable.PSVersion.Major -ge 7) {
            Import-Module -Name $module -UseWindowsPowerShell -ErrorAction SilentlyContinue
        } else {
            Import-Module -Name $module -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        }
    }
}

# Connect to Microsoft Graph
$graphScopes = @("Files.ReadWrite.All", "Sites.ReadWrite.All", "User.Read.All")
$context = Get-MgContext
if (-not $context) {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    try {
        Connect-MgGraph -Scopes $graphScopes -NoWelcome -ErrorAction Stop
    } catch {
        Write-Host "Graph authentication failed: $_" -ForegroundColor Red
        exit 1
    }
}

# If -Account, discover all licensed users with a SharePoint/OneDrive service plan
if ($Account) {
    Write-Host "`nRetrieving all licensed users with OneDrive..." -ForegroundColor Cyan

    # Known service plan names that include OneDrive/SharePoint access
    $sharepointPlanNames = @(
        "SHAREPOINTSTANDARD",
        "SHAREPOINTENTERPRISE",
        "SHAREPOINTONLINE_MULTIGEO",
        "SHAREPOINTSTANDARD_EDU",
        "SHAREPOINTENTERPRISE_EDU",
        "SHAREPOINTENTERPRISE_MIDMARKET",
        "SHAREPOINTDESKLESS",
        "SHAREPOINTWAC"
    )

    $oneDriveUsers = @()
    $noOneDriveUsers = @()
    $uri = "https://graph.microsoft.com/v1.0/users?`$filter=assignedLicenses/`$count ne 0 and accountEnabled eq true&`$select=userPrincipalName,displayName,assignedPlans&`$top=999&`$count=true"
    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers @{ "ConsistencyLevel" = "eventual" }
        foreach ($u in $response.value) {
            # Skip guest/external users
            if ($u.userPrincipalName -match '#EXT#') { continue }

            # Check if user has an enabled SharePoint service plan (which provides OneDrive)
            $hasOneDrive = $false
            foreach ($plan in $u.assignedPlans) {
                if ($plan.capabilityStatus -eq "Enabled" -and $plan.service -eq "SharePoint") {
                    $hasOneDrive = $true
                    break
                }
            }

            if ($hasOneDrive) {
                $oneDriveUsers += [PSCustomObject]@{ displayName = $u.displayName; upn = $u.userPrincipalName }
            } else {
                $noOneDriveUsers += [PSCustomObject]@{ displayName = $u.displayName }
            }
        }
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    # Display sorted results
    foreach ($u in ($oneDriveUsers | Sort-Object displayName)) {
        Write-Host "  $($u.displayName) ($($u.upn))" -ForegroundColor DarkGray
    }
    foreach ($u in ($noOneDriveUsers | Sort-Object displayName)) {
        Write-Host "  $($u.displayName) - no OneDrive license, skipping" -ForegroundColor DarkGray
    }

    $users = ($oneDriveUsers | Sort-Object displayName).upn
    if ($users.Count -eq 0) {
        Write-Host "No licensed users with OneDrive found." -ForegroundColor Red
        exit 1
    }
    Write-Host "Found $($users.Count) user(s) with OneDrive" -ForegroundColor Green
}

# Get the admin's email for granting OneDrive access later
$adminUser = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/me?`$select=userPrincipalName").userPrincipalName
Write-Host "Signed in as: $adminUser" -ForegroundColor Gray

# Detect SharePoint domain from the tenant's root site
Write-Host "Detecting SharePoint domain..." -ForegroundColor Cyan
try {
    $rootSite = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/root"
    $rootUrl = $rootSite.webUrl

    if ($rootUrl) {
        $uri = [System.Uri]$rootUrl
        $subdomain = $uri.Host -replace '\.sharepoint\.com$', ''
        Write-Host "Detected SharePoint domain: $subdomain.sharepoint.com" -ForegroundColor Green
    } else {
        throw "Could not retrieve root site URL"
    }
} catch {
    Write-Host "Could not detect SharePoint domain via Graph, using email domain as fallback..." -ForegroundColor Yellow
    $emailDomain = $users[0].Split('@')[1]
    $subdomain = $emailDomain.Split('.')[0]
    Write-Host "Using SharePoint subdomain: $subdomain (from $emailDomain)" -ForegroundColor Cyan
}

# Construct full SharePoint domain
$domain = "$subdomain.sharepoint.com"
$adminUrl = "https://$subdomain-admin.sharepoint.com"

Write-Host "Using SharePoint domain: $domain" -ForegroundColor Cyan

# Connect to SharePoint Online
Write-Host "`nConnecting to SharePoint Online..." -ForegroundColor Cyan
Connect-SPOService -Url $adminUrl

# Get all tenant sites once (used for non-Teams site discovery)
Write-Host "`nRetrieving all SharePoint sites in tenant..." -ForegroundColor Cyan
$allTenantSites = Get-SPOSite -Limit All | Where-Object {
    $_.Url -like "*/sites/*" -and $_.Template -ne "RedirectSite#0"
}
Write-Host "Found $($allTenantSites.Count) sites in tenant" -ForegroundColor Green

# Process each user
foreach ($currentUser in $users) {
    Write-Host "`n################################################################" -ForegroundColor Cyan
    Write-Host "Processing user: $currentUser" -ForegroundColor Yellow
    Write-Host "################################################################" -ForegroundColor Cyan

    # Grant admin temporary access to user's OneDrive
    $userOneDriveUrl = "https://$subdomain-my.sharepoint.com/personal/" + ($currentUser -replace '[@.]', '_')
    try {
        Set-SPOUser -Site $userOneDriveUrl -LoginName $adminUser -IsSiteCollectionAdmin $true -ErrorAction Stop | Out-Null
        Write-Host "Granted admin access to $currentUser's OneDrive" -ForegroundColor Gray
    } catch {
        Write-Host "Warning: Could not grant admin access to OneDrive: $_" -ForegroundColor Yellow
    }

    $allUserSites = @()

    # Discover SharePoint sites the user has access to
    Write-Host "`nChecking SharePoint sites for $currentUser..." -ForegroundColor Cyan

    # 1) Get user ID and all transitive group memberships (M365 groups + security groups)
    Write-Host "  Checking group memberships..." -ForegroundColor Gray
    $userId = $null
    $userGroupIds = @{}
    try {
        $userId = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$currentUser`?`$select=id").id

        # transitiveMemberOf resolves nested groups — catches security groups inside other groups
        $transitiveUri = "https://graph.microsoft.com/v1.0/users/$userId/transitiveMemberOf?`$select=id,displayName,groupTypes&`$top=999"
        do {
            $transitiveResponse = Invoke-MgGraphRequest -Method GET -Uri $transitiveUri
            foreach ($grp in $transitiveResponse.value) {
                if ($grp.'@odata.type' -eq '#microsoft.graph.group') {
                    $userGroupIds[$grp.id] = $grp.displayName
                }
            }
            $transitiveUri = $transitiveResponse.'@odata.nextLink'
        } while ($transitiveUri)
        Write-Host "  Found $($userGroupIds.Count) group memberships" -ForegroundColor Gray
    } catch {
        Write-Host "  Warning: Could not retrieve group memberships: $_" -ForegroundColor Yellow
    }

    # 2) M365 group-connected sites: directly look up SharePoint sites for Unified groups
    $groupSiteUrls = @{}
    foreach ($gid in $userGroupIds.Keys) {
        try {
            $groupSite = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$gid/sites/root?`$select=id,displayName,webUrl" -ErrorAction Stop
            if ($groupSite.webUrl -and -not $groupSiteUrls.ContainsKey($groupSite.webUrl)) {
                $groupSiteUrls[$groupSite.webUrl] = $true
                $siteObj = [PSCustomObject]@{
                    displayName = if ($groupSite.displayName) { $groupSite.displayName } else { $userGroupIds[$gid] }
                    webUrl = $groupSite.webUrl
                    id = $groupSite.id
                }
                $allUserSites += $siteObj
                Write-Host "  $($siteObj.displayName) (via M365 group)" -ForegroundColor DarkGray
            }
        } catch {
            # Not a Unified group or no site — expected for security groups
        }
    }

    # 3) Direct + security group access: check each tenant site
    Write-Host "  Checking site permissions (direct + security group)..." -ForegroundColor Gray
    $knownUrls = @{}
    foreach ($s in $allUserSites) { $knownUrls[$s.webUrl] = $true }

    foreach ($tenantSite in $allTenantSites) {
        if ($knownUrls.ContainsKey($tenantSite.Url)) { continue }
        $found = $false

        # Fast path: check if user is directly listed
        try {
            $spoUser = Get-SPOUser -Site $tenantSite.Url -LoginName $currentUser -ErrorAction SilentlyContinue
            if ($spoUser) { $found = $true }
        } catch { }

        # Slow path: check if any of the user's security groups are listed as site users
        if (-not $found -and $userGroupIds.Count -gt 0) {
            try {
                $siteUsers = Get-SPOUser -Site $tenantSite.Url -ErrorAction SilentlyContinue
                foreach ($siteUser in $siteUsers) {
                    # SharePoint stores Azure AD groups as claims containing the group GUID
                    foreach ($gid in $userGroupIds.Keys) {
                        if ($siteUser.LoginName -match [regex]::Escape($gid)) {
                            $found = $true
                            break
                        }
                    }
                    if ($found) { break }
                }
            } catch { }
        }

        if ($found) {
            $siteName = $tenantSite.Url.Split('/')[-1]
            $siteObj = [PSCustomObject]@{
                displayName = if ($tenantSite.Title) { $tenantSite.Title } else { $siteName }
                webUrl = $tenantSite.Url
                id = $null
            }
            $allUserSites += $siteObj
            Write-Host "  $($siteObj.displayName) (direct/security group)" -ForegroundColor DarkGray
        }
    }

    # Filter and deduplicate
    $userSites = $allUserSites | Where-Object {
        $_.webUrl -like "*/sites/*" -and $_.displayName -ne "All Company"
    } | Sort-Object -Property webUrl -Unique

    $siteCount = @($userSites).Count
    if ($siteCount -eq 0) {
        Write-Host "`nNo sites found for $currentUser" -ForegroundColor Yellow
        continue
    }

    Write-Host "`nFound $siteCount site(s) accessible by $currentUser" -ForegroundColor Green
    Write-Host "Processing each site...`n" -ForegroundColor Cyan

    foreach ($spoSite in $userSites) {
        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host "Site: $($spoSite.displayName)" -ForegroundColor Yellow
        Write-Host "URL: $($spoSite.webUrl)" -ForegroundColor Gray

        # Compute the shortcut name up front so we can show it before prompting
        $shortcutName = $spoSite.displayName
        if ($shortcutName -match '^(.+?) - \1$') {
            $shortcutName = $Matches[1]
        }
        Write-Host "Will map as: $shortcutName" -ForegroundColor White

        if (-not $AutoMap) {
            $mapSite = Read-Host "Map (Y) / Remove (N) / Skip (S)?"
            if ($mapSite -eq "S" -or $mapSite -eq "s") {
                Write-Host "Skipping $($spoSite.displayName)..." -ForegroundColor Gray
                continue
            }
            if ($mapSite -eq "N" -or $mapSite -eq "n") {
                # Remove existing shortcut from OneDrive
                $removed = $false
                $namesToCheck = @($shortcutName, $spoSite.displayName, "General") | Select-Object -Unique
                foreach ($nameToRemove in $namesToCheck) {
                    try {
                        $encodedName = [Uri]::EscapeDataString($nameToRemove)
                        $existingItem = Invoke-MgGraphRequest -Method GET `
                            -Uri "https://graph.microsoft.com/v1.0/users/$currentUser/drive/root:/$encodedName" `
                            -ErrorAction Stop
                        if ($existingItem) {
                            Invoke-MgGraphRequest -Method DELETE `
                                -Uri "https://graph.microsoft.com/v1.0/users/$currentUser/drive/items/$($existingItem.id)" | Out-Null
                            Write-Host "Removed shortcut '$nameToRemove' from OneDrive." -ForegroundColor Green
                            $removed = $true
                            break
                        }
                    } catch {
                        # 404 = doesn't exist, try next name
                    }
                }
                if (-not $removed) {
                    Write-Host "No existing shortcut found to remove for $($spoSite.displayName)." -ForegroundColor Yellow
                }
                continue
            }
        }

        try {
            # Extract site name from URL
            $siteName = $spoSite.webUrl.Split('/')[-1]
            $sitePath = "/sites/$siteName"

            Write-Host "Getting site information..." -ForegroundColor Cyan
            $site = Get-MgSite -SiteId "$domain`:$sitePath"
            $drives = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($site.Id)/drives"
            $drive = $drives.value | Where-Object { $_.name -eq "Documents" }

            if (-not $drive) {
                Write-Host "No Documents library found. Skipping..." -ForegroundColor Yellow
                continue
            }

            # Try /General first (Teams sites), fall back to drive root (non-Teams sites)
            $folderItem = $null
            try {
                $folderItem = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root:/General"
                Write-Host "Mapping /General folder..." -ForegroundColor Gray
            } catch {
                Write-Host "/General folder not found, mapping Documents library root..." -ForegroundColor Gray
                $folderItem = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root"
            }

            # Check if shortcut already exists with the correct name (direct path lookup)
            $existingCorrect = $null
            try {
                $encodedCorrect = [Uri]::EscapeDataString($shortcutName)
                $existingCorrect = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/users/$currentUser/drive/root:/$encodedCorrect" `
                    -ErrorAction Stop
            } catch { }
            if ($existingCorrect) {
                Write-Host "Already mapped - '$shortcutName' exists in OneDrive. Skipping." -ForegroundColor Yellow
                continue
            }

            # Check for existing shortcuts with old bad names and rename them
            $possibleOldNames = @(
                $spoSite.displayName                          # raw display name (e.g. "3 Amigos - 3 Amigos")
                "$shortcutName - $shortcutName"               # duplicated form if displayName was already clean
                "General"                                     # old script used folder name for Teams sites
            ) | Select-Object -Unique | Where-Object { $_ -ne $shortcutName }

            $renamed = $false
            foreach ($oldName in $possibleOldNames) {
                try {
                    $encodedOld = [Uri]::EscapeDataString($oldName)
                    $oldItem = Invoke-MgGraphRequest -Method GET `
                        -Uri "https://graph.microsoft.com/v1.0/users/$currentUser/drive/root:/$encodedOld" `
                        -ErrorAction Stop
                    if ($oldItem) {
                        Write-Host "Renaming '$oldName' -> '$shortcutName'..." -ForegroundColor Yellow
                        Invoke-MgGraphRequest -Method PATCH `
                            -Uri "https://graph.microsoft.com/v1.0/users/$currentUser/drive/items/$($oldItem.id)" `
                            -Body (@{ name = $shortcutName } | ConvertTo-Json) -ContentType "application/json" | Out-Null
                        Write-Host "Successfully renamed to '$shortcutName'!" -ForegroundColor Green
                        $renamed = $true
                        break
                    }
                } catch {
                    # 404 = doesn't exist, expected
                }
            }
            if ($renamed) { continue }

            # No existing shortcut found — create new one
            $body = @{
                name       = $shortcutName
                remoteItem = @{
                    id              = $folderItem.id
                    parentReference = @{
                        driveId = $drive.id
                    }
                }
            }
            $newItem = Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/users/$currentUser/drive/root/children" `
                -Body ($body | ConvertTo-Json -Depth 10) -ContentType "application/json"

            # Graph ignores the name field on shortcut creation — rename if needed
            if ($newItem.name -ne $shortcutName) {
                Invoke-MgGraphRequest -Method PATCH `
                    -Uri "https://graph.microsoft.com/v1.0/users/$currentUser/drive/items/$($newItem.id)" `
                    -Body (@{ name = $shortcutName } | ConvertTo-Json) -ContentType "application/json" | Out-Null
            }
            Write-Host "Successfully mapped as '$shortcutName'!" -ForegroundColor Green
        }
        catch {
            $errMsg = $_.ToString()
            if ($errMsg -match "shortcutAlreadyExists" -or $errMsg -match "That shortcut already exists") {
                Write-Host "Already mapped - $($spoSite.displayName) shortcut already exists in OneDrive. Skipping." -ForegroundColor Yellow
            } else {
                Write-Host "Error mapping $($spoSite.displayName): $_" -ForegroundColor Red
            }
        }
    }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Site mapping complete for all users!" -ForegroundColor Green
