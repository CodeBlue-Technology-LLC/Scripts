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

    foreach ($tenantSite in $allTenantSites) {
        try {
            $spoUser = Get-SPOUser -Site $tenantSite.Url -LoginName $currentUser -ErrorAction SilentlyContinue
            if ($spoUser) {
                $siteName = $tenantSite.Url.Split('/')[-1]
                $siteObj = [PSCustomObject]@{
                    displayName = if ($tenantSite.Title) { $tenantSite.Title } else { $siteName }
                    webUrl = $tenantSite.Url
                    id = $null
                }
                $allUserSites += $siteObj
                Write-Host "  $($siteObj.displayName)" -ForegroundColor DarkGray
            }
        } catch {
            # User doesn't have access, skip
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

        if (-not $AutoMap) {
            $mapSite = Read-Host "Map this site to OneDrive? (Y/n)"
            if ($mapSite -eq "N" -or $mapSite -eq "n") {
                Write-Host "Skipping $($spoSite.displayName)..." -ForegroundColor Gray
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

            $shortcutName = if ($folderItem.name -eq "root") { $spoSite.displayName } else { $folderItem.name }

            # Check if shortcut already exists in user's OneDrive
            $existingItem = $null
            try {
                $encodedName = [Uri]::EscapeDataString($shortcutName)
                $existingItem = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/users/$currentUser/drive/root:/${encodedName}" `
                    -ErrorAction Stop
            } catch {
                # 404 means it doesn't exist, which is expected
            }
            if ($existingItem) {
                Write-Host "Already mapped - '$shortcutName' already exists in OneDrive. Skipping." -ForegroundColor Yellow
                continue
            }

            $body = @{
                name       = $shortcutName
                remoteItem = @{
                    id              = $folderItem.id
                    parentReference = @{
                        driveId = $drive.id
                    }
                }
            }
            Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/users/$currentUser/drive/root/children" `
                -Body ($body | ConvertTo-Json -Depth 10) -ContentType "application/json"

            Write-Host "Successfully mapped $($spoSite.displayName)!" -ForegroundColor Green
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
