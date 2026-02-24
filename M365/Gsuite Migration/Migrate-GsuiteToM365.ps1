#Requires -Version 5.1

<#
.SYNOPSIS
    GSuite to Microsoft 365 email migration tool.
.DESCRIPTION
    Interactive menu-driven script that walks through the full GSuite-to-M365
    email migration process: GCP service account setup, API enablement,
    subdomain creation, user export/import, and migration batch creation.
.NOTES
    Google Drive to SharePoint/OneDrive migration will be added in a future version.
#>

# ── Module checks ──────────────────────────────────────────────────────────────
foreach ($mod in @('ExchangeOnlineManagement', 'Microsoft.Graph.Users')) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "Installing $mod..." -ForegroundColor Yellow
        Install-Module $mod -Force -Scope CurrentUser
    }
}

# ── Config helpers ─────────────────────────────────────────────────────────────
$script:ConfigDir  = Join-Path $PSScriptRoot "Config"
$script:ConfigFile = Join-Path $script:ConfigDir "credentials.xml"

function Save-MigrationConfig {
    param([hashtable]$Config)
    if (-not (Test-Path $script:ConfigDir)) {
        New-Item -ItemType Directory -Path $script:ConfigDir -Force | Out-Null
    }
    $Config | Export-Clixml -Path $script:ConfigFile -Force
}

function Get-MigrationConfig {
    if (Test-Path $script:ConfigFile) {
        return Import-Clixml -Path $script:ConfigFile
    }
    return @{}
}

# ── Google OAuth helper ────────────────────────────────────────────────────────
# Uses browser-based OAuth 2.0 flow with a loopback redirect.
# Requires a GCP OAuth Client ID — if not already created, Step 1 will guide
# the user to create one in the GCP console (or use an existing one).

function Get-GoogleAccessToken {
    param(
        [string]$ClientId,
        [string]$ClientSecret,
        [string[]]$Scopes
    )

    $redirectUri = "http://localhost:8642"
    $scopeString = ($Scopes -join " ")
    $state = [guid]::NewGuid().ToString()

    $authUrl = "https://accounts.google.com/o/oauth2/v2/auth?" +
        "client_id=$ClientId" +
        "&redirect_uri=$([uri]::EscapeDataString($redirectUri))" +
        "&response_type=code" +
        "&scope=$([uri]::EscapeDataString($scopeString))" +
        "&access_type=offline" +
        "&state=$state" +
        "&prompt=consent"

    # Start local HTTP listener
    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add("$redirectUri/")
    $listener.Start()

    Write-Host "`nOpening browser for Google authentication..." -ForegroundColor Cyan
    Start-Process $authUrl

    # Wait for the callback
    $context = $listener.GetContext()
    $code = $context.Request.QueryString["code"]

    # Send response to browser
    $response = $context.Response
    $html = [System.Text.Encoding]::UTF8.GetBytes("<html><body><h2>Authentication successful!</h2><p>You can close this tab.</p></body></html>")
    $response.ContentLength64 = $html.Length
    $response.OutputStream.Write($html, 0, $html.Length)
    $response.Close()
    $listener.Stop()

    if (-not $code) {
        throw "Failed to receive authorization code from Google."
    }

    # Exchange code for tokens
    $tokenBody = @{
        code          = $code
        client_id     = $ClientId
        client_secret = $ClientSecret
        redirect_uri  = $redirectUri
        grant_type    = "authorization_code"
    }
    $tokenResponse = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method Post -Body $tokenBody
    return $tokenResponse.access_token
}

function Invoke-GoogleApi {
    param(
        [string]$Uri,
        [string]$Method = "GET",
        [string]$AccessToken,
        [object]$Body
    )
    $headers = @{ Authorization = "Bearer $AccessToken" }
    $params = @{
        Uri         = $Uri
        Method      = $Method
        Headers     = $headers
        ContentType = "application/json"
    }
    if ($Body) {
        $params.Body = ($Body | ConvertTo-Json -Depth 10)
    }
    Invoke-RestMethod @params
}

# ── Step functions ─────────────────────────────────────────────────────────────

function Step1-CreateServiceAccount {
    param([string]$Domain)

    $config = Get-MigrationConfig

    Write-Host "`n═══════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Step 1: Create GCP Project & Service Account" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════`n" -ForegroundColor Cyan

    # Get Google admin email
    $adminEmail = Read-Host "Enter Google Workspace admin email"

    Write-Host "`nTo automate GCP setup, we need an OAuth Client ID." -ForegroundColor Yellow
    Write-Host "If you don't have one, create it in GCP Console:" -ForegroundColor Yellow
    Write-Host "  1. Go to https://console.cloud.google.com/apis/credentials" -ForegroundColor White
    Write-Host "  2. Create Credentials > OAuth Client ID" -ForegroundColor White
    Write-Host "  3. Application type: Desktop app" -ForegroundColor White
    Write-Host "  4. Copy the Client ID and Client Secret`n" -ForegroundColor White

    $clientId     = Read-Host "Enter OAuth Client ID"
    $clientSecret = Read-Host "Enter OAuth Client Secret"

    $scopes = @(
        "https://www.googleapis.com/auth/cloud-platform"
        "https://www.googleapis.com/auth/iam"
        "https://www.googleapis.com/auth/admin.directory.domain"
        "https://www.googleapis.com/auth/admin.directory.user.readonly"
    )
    $token = Get-GoogleAccessToken -ClientId $clientId -ClientSecret $clientSecret -Scopes $scopes

    # Create GCP project
    $projectId = "m365-migration-" + ($Domain -replace '\.', '-').Substring(0, [Math]::Min(($Domain -replace '\.', '-').Length, 20))
    Write-Host "`nCreating GCP project '$projectId'..." -ForegroundColor Cyan

    try {
        $project = Invoke-GoogleApi -Uri "https://cloudresourcemanager.googleapis.com/v1/projects" `
            -Method POST -AccessToken $token -Body @{
                projectId = $projectId
                name      = "M365 Migration - $Domain"
            }
        Write-Host "  Project creation initiated. Waiting for provisioning..." -ForegroundColor Gray
        Start-Sleep -Seconds 5
    }
    catch {
        if ($_.Exception.Response.StatusCode -eq 409) {
            Write-Host "  Project '$projectId' already exists, continuing..." -ForegroundColor Yellow
        }
        else { throw }
    }

    # Create service account
    $saName = "m365-migration"
    Write-Host "Creating service account '$saName'..." -ForegroundColor Cyan

    try {
        $sa = Invoke-GoogleApi -Uri "https://iam.googleapis.com/v1/projects/$projectId/serviceAccounts" `
            -Method POST -AccessToken $token -Body @{
                accountId      = $saName
                serviceAccount = @{
                    displayName = "M365 Migration Service Account"
                    description = "Used for GSuite to M365 email migration"
                }
            }
    }
    catch {
        if ($_.Exception.Response.StatusCode -eq 409) {
            Write-Host "  Service account already exists, fetching..." -ForegroundColor Yellow
            $sa = Invoke-GoogleApi -Uri "https://iam.googleapis.com/v1/projects/$projectId/serviceAccounts/$saName@$projectId.iam.gserviceaccount.com" `
                -AccessToken $token
        }
        else { throw }
    }

    $saEmail    = $sa.email
    $saUniqueId = $sa.uniqueId

    Write-Host "  Service Account Email: $saEmail" -ForegroundColor Green
    Write-Host "  Unique ID (Client ID): $saUniqueId" -ForegroundColor Green

    # Create JSON key
    Write-Host "Creating JSON key..." -ForegroundColor Cyan
    $keyResponse = Invoke-GoogleApi -Uri "https://iam.googleapis.com/v1/projects/$projectId/serviceAccounts/$saEmail/keys" `
        -Method POST -AccessToken $token -Body @{ keyAlgorithm = "KEY_ALG_RSA_2048"; privateKeyType = "TYPE_GOOGLE_CREDENTIALS_FILE" }

    $keyJson = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($keyResponse.privateKeyData))
    $keyFilePath = Join-Path $script:ConfigDir "service-account-key.json"

    if (-not (Test-Path $script:ConfigDir)) {
        New-Item -ItemType Directory -Path $script:ConfigDir -Force | Out-Null
    }
    $keyJson | Out-File -FilePath $keyFilePath -Encoding UTF8 -Force

    Write-Host "  Key saved to: $keyFilePath" -ForegroundColor Green

    # Save config
    $config.Domain          = $Domain
    $config.AdminEmail      = $adminEmail
    $config.KeyFilePath     = $keyFilePath
    $config.ProjectId       = $projectId
    $config.SAEmail         = $saEmail
    $config.SAUniqueId      = $saUniqueId
    $config.OAuthClientId   = $clientId
    $config.OAuthClientSecret = $clientSecret
    Save-MigrationConfig $config

    Write-Host "`n┌─────────────────────────────────────────────────────┐" -ForegroundColor Green
    Write-Host "│  Service Account Created Successfully!              │" -ForegroundColor Green
    Write-Host "│                                                     │" -ForegroundColor Green
    Write-Host "│  Client ID: $saUniqueId" -ForegroundColor Green
    Write-Host "│  Email:     $saEmail" -ForegroundColor Green
    Write-Host "│                                                     │" -ForegroundColor Green
    Write-Host "│  Save the Client ID — you'll need it for Step 2.    │" -ForegroundColor Green
    Write-Host "└─────────────────────────────────────────────────────┘" -ForegroundColor Green
}

function Step2-ManualInstructions {
    $config = Get-MigrationConfig

    Write-Host "`n═══════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Step 2: Manual Configuration Required" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════`n" -ForegroundColor Cyan

    $clientId = if ($config.SAUniqueId) { $config.SAUniqueId } else { "<Client ID from Step 1>" }
    $saEmail  = if ($config.SAEmail)    { $config.SAEmail }    else { "<service account email>" }

    Write-Host "Two manual steps are required that cannot be automated:`n" -ForegroundColor Yellow

    Write-Host "── 1. Enable Domain-Wide Delegation ──────────────────────" -ForegroundColor White
    Write-Host "  a. Go to: https://console.cloud.google.com/iam-admin/serviceaccounts" -ForegroundColor White
    Write-Host "  b. Select project and click on service account: $saEmail" -ForegroundColor White
    Write-Host "  c. Click 'Show advanced settings' or 'Show domain-wide delegation'" -ForegroundColor White
    Write-Host "  d. Check 'Enable Google Workspace Domain-wide Delegation'" -ForegroundColor White
    Write-Host "  e. Set a product name for consent screen and Save`n" -ForegroundColor White

    Write-Host "── 2. Authorize OAuth Scopes ─────────────────────────────" -ForegroundColor White
    Write-Host "  a. Go to: https://admin.google.com" -ForegroundColor White
    Write-Host "  b. Navigate to: Security > Access and data control > API Controls" -ForegroundColor White
    Write-Host "  c. Click 'Manage Domain Wide Delegation'" -ForegroundColor White
    Write-Host "  d. Click 'Add new'" -ForegroundColor White
    Write-Host "  e. Client ID: $clientId" -ForegroundColor White
    Write-Host "  f. Paste the following scopes:`n" -ForegroundColor White

    $scopes = "https://mail.google.com/,https://www.google.com/m8/feeds,https://www.googleapis.com/auth/contacts.readonly,https://www.googleapis.com/auth/calendar,https://www.googleapis.com/auth/calendar.readonly,https://www.googleapis.com/auth/admin.directory.group.readonly,https://www.googleapis.com/auth/admin.directory.user.readonly,https://www.googleapis.com/auth/drive,https://sites.google.com/feeds/,https://www.googleapis.com/auth/gmail.settings.sharing,https://www.googleapis.com/auth/gmail.settings.basic,https://www.googleapis.com/auth/contacts.other.readonly"

    Write-Host $scopes -ForegroundColor Cyan
    Write-Host ""

    # Copy to clipboard
    $scopes | Set-Clipboard
    Write-Host "  (Scopes copied to clipboard)" -ForegroundColor Gray

    Write-Host "`n  g. Click 'Authorize'" -ForegroundColor White
    Write-Host "`nNote: It may take 15 minutes to 24 hours for these settings to propagate.`n" -ForegroundColor Yellow

    Read-Host "Press Enter when both steps are complete"
    Write-Host "Step 2 complete." -ForegroundColor Green
}

function Step3-EnableAPIs {
    param([string]$Domain)

    $config = Get-MigrationConfig

    Write-Host "`n═══════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Step 3: Enable Required Google APIs" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════`n" -ForegroundColor Cyan

    $token = Get-GoogleAccessToken -ClientId $config.OAuthClientId -ClientSecret $config.OAuthClientSecret `
        -Scopes @("https://www.googleapis.com/auth/cloud-platform", "https://www.googleapis.com/auth/service.management")

    $projectId = $config.ProjectId
    $apis = @(
        "gmail.googleapis.com"
        "calendar-json.googleapis.com"
        "contacts.googleapis.com"
        "people.googleapis.com"
        "admin.googleapis.com"
    )

    Write-Host "Enabling APIs for project '$projectId'..." -ForegroundColor Cyan

    try {
        Invoke-GoogleApi -Uri "https://serviceusage.googleapis.com/v1/projects/$projectId/services:batchEnable" `
            -Method POST -AccessToken $token -Body @{ serviceIds = $apis }
        Write-Host "  Batch enable request submitted." -ForegroundColor Green
    }
    catch {
        Write-Host "  Batch enable failed, trying individually..." -ForegroundColor Yellow
        foreach ($api in $apis) {
            try {
                Invoke-GoogleApi -Uri "https://serviceusage.googleapis.com/v1/projects/$projectId/services/${api}:enable" `
                    -Method POST -AccessToken $token
                Write-Host "  Enabled: $api" -ForegroundColor Green
            }
            catch {
                Write-Host "  Failed to enable ${api}: $_" -ForegroundColor Red
            }
        }
    }

    Write-Host "`nStep 3 complete. APIs enabled." -ForegroundColor Green
}

function Step4-CreateSubdomains {
    param([string]$Domain)

    $config = Get-MigrationConfig

    Write-Host "`n═══════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Step 4: Create Subdomains (Google + M365)" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════`n" -ForegroundColor Cyan

    $gsuiteSub = "gsuite.$Domain"
    $m365Sub   = "m365.$Domain"

    # ── Google side ──
    Write-Host "── Google Workspace ──" -ForegroundColor White
    $token = Get-GoogleAccessToken -ClientId $config.OAuthClientId -ClientSecret $config.OAuthClientSecret `
        -Scopes @("https://www.googleapis.com/auth/admin.directory.domain")

    foreach ($sub in @($gsuiteSub, $m365Sub)) {
        Write-Host "Adding domain alias '$sub' to Google Workspace..." -ForegroundColor Cyan
        try {
            Invoke-GoogleApi -Uri "https://admin.googleapis.com/admin/directory/v1/customer/my_customer/domains" `
                -Method POST -AccessToken $token -Body @{ domainName = $sub }
            Write-Host "  Added: $sub" -ForegroundColor Green
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 409) {
                Write-Host "  Already exists: $sub" -ForegroundColor Yellow
            }
            else {
                Write-Host "  Failed: $sub - $_" -ForegroundColor Red
            }
        }
    }

    # ── M365 side ──
    Write-Host "`n── Microsoft 365 ──" -ForegroundColor White

    if (-not (Get-ConnectionInformation)) {
        Connect-ExchangeOnline
    }

    $existing = Get-AcceptedDomain | Where-Object { $_.DomainName -eq $m365Sub }
    if (-not $existing) {
        Write-Host "Adding accepted domain '$m365Sub' to M365..." -ForegroundColor Cyan
        try {
            New-AcceptedDomain -Name $m365Sub -DomainName $m365Sub -DomainType InternalRelay
            Write-Host "  Added: $m365Sub" -ForegroundColor Green
        }
        catch {
            Write-Host "  Failed: $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "  Already exists: $m365Sub" -ForegroundColor Yellow
    }

    # ── DNS instructions ──
    Write-Host "`n── DNS Records Required ──────────────────────────────────" -ForegroundColor Yellow
    Write-Host "Add the following DNS records at your DNS provider:`n" -ForegroundColor White
    Write-Host "  1. $gsuiteSub" -ForegroundColor Cyan
    Write-Host "     Type: TXT (or CNAME) — Google verification record" -ForegroundColor White
    Write-Host "     Check Google Admin > Domains for the exact value`n" -ForegroundColor Gray
    Write-Host "  2. $m365Sub (Google verification)" -ForegroundColor Cyan
    Write-Host "     Type: TXT (or CNAME) — Google verification record" -ForegroundColor White
    Write-Host "     Check Google Admin > Domains for the exact value`n" -ForegroundColor Gray
    Write-Host "  3. $m365Sub (M365 verification)" -ForegroundColor Cyan
    Write-Host "     Type: TXT — Microsoft verification record" -ForegroundColor White
    Write-Host "     Check M365 Admin Center > Settings > Domains for the exact value`n" -ForegroundColor Gray
    Write-Host "─────────────────────────────────────────────────────────" -ForegroundColor Yellow

    Read-Host "`nPress Enter when DNS records are added and verified"
    Write-Host "Step 4 complete." -ForegroundColor Green
}

function Step5-ExportUsersAndGenerateCSV {
    param([string]$Domain)

    $config = Get-MigrationConfig

    Write-Host "`n═══════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Step 5: Export Users from Google & Generate CSV" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════`n" -ForegroundColor Cyan

    $token = Get-GoogleAccessToken -ClientId $config.OAuthClientId -ClientSecret $config.OAuthClientSecret `
        -Scopes @("https://www.googleapis.com/auth/admin.directory.user.readonly")

    # Fetch all users (including archived/suspended)
    Write-Host "Fetching users from Google Workspace..." -ForegroundColor Cyan
    $allUsers = @()
    $pageToken = $null

    do {
        $uri = "https://admin.googleapis.com/admin/directory/v1/users?domain=$Domain&maxResults=500&projection=full"
        if ($pageToken) { $uri += "&pageToken=$pageToken" }

        $response = Invoke-GoogleApi -Uri $uri -AccessToken $token
        if ($response.users) {
            $allUsers += $response.users
        }
        $pageToken = $response.nextPageToken
    } while ($pageToken)

    Write-Host "  Found $($allUsers.Count) users" -ForegroundColor Green

    # Generate unique passwords from DinoPass
    Write-Host "Generating passwords from DinoPass..." -ForegroundColor Cyan
    $passwords = @()
    $attempts = 0

    for ($i = 0; $i -lt $allUsers.Count; $i++) {
        $pw = (Invoke-RestMethod -Uri "https://www.dinopass.com/password/strong?row=$i" -Method Get).Trim()
        $passwords += $pw
        Start-Sleep -Milliseconds 200  # rate limiting
    }

    # Uniqueness check — re-fetch any duplicates
    $maxRetries = 10
    $retryCount = 0
    do {
        $dupes = $passwords | Group-Object | Where-Object { $_.Count -gt 1 }
        if ($dupes) {
            Write-Host "  Found $($dupes.Count) duplicate password(s), re-fetching..." -ForegroundColor Yellow
            foreach ($dupe in $dupes) {
                $indices = for ($i = 0; $i -lt $passwords.Count; $i++) {
                    if ($passwords[$i] -eq $dupe.Name) { $i }
                }
                # Keep the first, re-fetch the rest
                foreach ($idx in ($indices | Select-Object -Skip 1)) {
                    $newPw = (Invoke-RestMethod -Uri "https://www.dinopass.com/password/strong?retry=$idx&r=$([guid]::NewGuid())" -Method Get).Trim()
                    $passwords[$idx] = $newPw
                    Start-Sleep -Milliseconds 200
                }
            }
        }
        $retryCount++
    } while ($dupes -and $retryCount -lt $maxRetries)

    if ($dupes) {
        Write-Host "  Warning: Could not eliminate all duplicate passwords after $maxRetries attempts" -ForegroundColor Red
    }
    else {
        Write-Host "  All passwords are unique" -ForegroundColor Green
    }

    # Build CSV data
    $csvData = for ($i = 0; $i -lt $allUsers.Count; $i++) {
        $u = $allUsers[$i]
        $email     = $u.primaryEmail
        $username  = ($email -split '@')[0]
        $firstName = $u.name.givenName
        $lastName  = $u.name.familyName

        # Get phone — try work phone first, then any phone
        $phone = ""
        if ($u.phones) {
            $workPhone = $u.phones | Where-Object { $_.type -eq "work" } | Select-Object -First 1
            if ($workPhone) { $phone = $workPhone.value }
            elseif ($u.phones[0].value) { $phone = $u.phones[0].value }
        }

        # Determine archived status
        $archived = if ($u.archived -eq $true -or $u.suspended -eq $true) { "TRUE" } else { "FALSE" }

        [PSCustomObject]@{
            FirstName                  = $firstName
            LastName                   = $lastName
            ExternalEmailAddress       = "$username@gsuite.$Domain"
            MicrosoftOnlineServicesID  = $email
            Password                   = $passwords[$i]
            ProxyAddress               = "$username@m365.$Domain"
            Phone                      = $phone
            Archived                   = $archived
        }
    }

    # Save CSV
    $csvPath = Join-Path $PSScriptRoot "m365userimport.csv"
    $csvData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nCSV saved to: $csvPath" -ForegroundColor Green

    # Display for review
    Write-Host "`n── User List ──────────────────────────────────────" -ForegroundColor Yellow
    $csvData | Format-Table FirstName, LastName, MicrosoftOnlineServicesID, Archived -AutoSize
    Write-Host "Total: $($csvData.Count) users ($($csvData | Where-Object { $_.Archived -eq 'TRUE' } | Measure-Object | Select-Object -ExpandProperty Count) archived)" -ForegroundColor Cyan

    Read-Host "`nReview the CSV. Press Enter to continue or Ctrl+C to abort"
    Write-Host "Step 5 complete." -ForegroundColor Green
}

function Step6-CreateMailUsers {
    param([string]$Domain)

    $config = Get-MigrationConfig

    Write-Host "`n═══════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Step 6: Create M365 Mail Users" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════`n" -ForegroundColor Cyan

    # Connect to Exchange Online
    if (-not (Get-ConnectionInformation)) {
        Connect-ExchangeOnline
    }

    # Connect to MgGraph — auto-detect tenant
    $tenantDomain = (Get-AcceptedDomain | Where-Object {
        $_.DomainName -like "*.onmicrosoft.com" -and $_.DomainName -notlike "*mail.onmicrosoft.com"
    }).DomainName
    Write-Host "Tenant: $tenantDomain" -ForegroundColor Cyan

    $mgContext = Get-MgContext
    if (-not $mgContext -or $mgContext.TenantId -ne (Get-OrganizationConfig).Guid) {
        if ($mgContext) { Disconnect-MgGraph }
        Connect-MgGraph -Scopes "User.ReadWrite.All" -TenantId $tenantDomain
    }

    # Import CSV
    $csvPath = Join-Path $PSScriptRoot "m365userimport.csv"
    if (-not (Test-Path $csvPath)) {
        Write-Host "CSV not found at $csvPath. Run Step 5 first." -ForegroundColor Red
        return
    }
    $users = Import-Csv $csvPath

    foreach ($user in $users) {
        $displayName = "$($user.FirstName) $($user.LastName)"
        $password = ConvertTo-SecureString $user.Password -AsPlainText -Force

        try {
            $existing = Get-MailUser -Identity $user.MicrosoftOnlineServicesID -ErrorAction SilentlyContinue
            if (-not $existing) {
                New-MailUser -Name $displayName `
                             -DisplayName $displayName `
                             -FirstName $user.FirstName `
                             -LastName $user.LastName `
                             -ExternalEmailAddress $user.ExternalEmailAddress `
                             -MicrosoftOnlineServicesID $user.MicrosoftOnlineServicesID `
                             -Password $password `
                             -ResetPasswordOnNextLogon $true

                Set-MailUser -Identity $user.MicrosoftOnlineServicesID `
                             -EmailAddresses @{Add = $user.ProxyAddress}

                Write-Host "Created: $displayName" -ForegroundColor Green
            }
            else {
                Write-Host "Skipped: $displayName (already exists)" -ForegroundColor Yellow
            }

            # Remove default MRM policy
            Set-MailUser -Identity $user.MicrosoftOnlineServicesID `
                         -RetentionPolicy $null `
                         -ErrorAction SilentlyContinue

            # Set phone via Graph (with retry for sync delay)
            if ($user.Phone) {
                $phoneSet = $false
                for ($i = 1; $i -le 6; $i++) {
                    try {
                        Update-MgUser -UserId $user.MicrosoftOnlineServicesID `
                                      -BusinessPhones @($user.Phone)
                        Write-Host "  Updated phone for $displayName" -ForegroundColor Cyan
                        $phoneSet = $true
                        break
                    }
                    catch {
                        Write-Host "  Waiting for $displayName to sync to Azure AD... (attempt $i/6)" -ForegroundColor Gray
                        Start-Sleep -Seconds 10
                    }
                }
                if (-not $phoneSet) {
                    Write-Host "  Could not set phone for $displayName after 6 attempts" -ForegroundColor Red
                }
            }
        }
        catch {
            Write-Host "Failed: $displayName - $_" -ForegroundColor Red
        }
    }

    Write-Host "`nStep 6 complete." -ForegroundColor Green
}

function Step7-CreateMigrationBatch {
    param([string]$Domain)

    $config = Get-MigrationConfig

    Write-Host "`n═══════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Step 7: Create Migration Batch" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════`n" -ForegroundColor Cyan

    # Connect to Exchange Online
    if (-not (Get-ConnectionInformation)) {
        Connect-ExchangeOnline
    }

    # Load config
    $adminEmail  = $config.AdminEmail
    $keyFilePath = $config.KeyFilePath

    if (-not $adminEmail -or -not $keyFilePath) {
        Write-Host "Missing admin email or key file path in config. Run Step 1 first." -ForegroundColor Red
        return
    }
    if (-not (Test-Path $keyFilePath)) {
        Write-Host "Service account key file not found at: $keyFilePath" -ForegroundColor Red
        return
    }

    # Import user CSV to generate migration CSV
    $csvPath = Join-Path $PSScriptRoot "m365userimport.csv"
    if (-not (Test-Path $csvPath)) {
        Write-Host "User CSV not found at $csvPath. Run Step 5 first." -ForegroundColor Red
        return
    }
    $users = Import-Csv $csvPath

    # Generate migration batch CSV
    $migrationCsvPath = Join-Path $PSScriptRoot "migration-batch.csv"
    $migrationData = $users | ForEach-Object {
        [PSCustomObject]@{ EmailAddress = $_.MicrosoftOnlineServicesID }
    }
    $migrationData | Export-Csv -Path $migrationCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Migration batch CSV saved to: $migrationCsvPath" -ForegroundColor Green
    Write-Host "  Users in batch: $($migrationData.Count)" -ForegroundColor Cyan

    # Test migration endpoint
    Write-Host "`nTesting migration server availability..." -ForegroundColor Cyan
    try {
        Test-MigrationServerAvailability -Gmail `
            -ServiceAccountKeyFileData $([System.IO.File]::ReadAllBytes($keyFilePath)) `
            -EmailAddress $adminEmail
        Write-Host "  Connection test successful!" -ForegroundColor Green
    }
    catch {
        Write-Host "  Connection test failed: $_" -ForegroundColor Red
        Write-Host "  Ensure Step 2 manual steps are complete and scopes have propagated (up to 24 hours)." -ForegroundColor Yellow
        $continue = Read-Host "Continue anyway? (y/N)"
        if ($continue -ne 'y') { return }
    }

    # Create migration endpoint
    $endpointName = "GsuiteEndpoint-$($Domain -replace '\.', '-')"
    Write-Host "`nCreating migration endpoint '$endpointName'..." -ForegroundColor Cyan
    try {
        $existingEndpoint = Get-MigrationEndpoint -Identity $endpointName -ErrorAction SilentlyContinue
        if ($existingEndpoint) {
            Write-Host "  Endpoint already exists, skipping creation." -ForegroundColor Yellow
        }
        else {
            New-MigrationEndpoint -Gmail `
                -ServiceAccountKeyFileData $([System.IO.File]::ReadAllBytes($keyFilePath)) `
                -EmailAddress $adminEmail `
                -Name $endpointName
            Write-Host "  Endpoint created." -ForegroundColor Green
        }
    }
    catch {
        Write-Host "  Failed to create endpoint: $_" -ForegroundColor Red
        return
    }

    # Create migration batch
    $batchName = "GsuiteBatch-$($Domain -replace '\.', '-')"
    $m365Sub   = "m365.$Domain"
    Write-Host "`nCreating migration batch '$batchName'..." -ForegroundColor Cyan
    try {
        New-MigrationBatch -SourceEndpoint $endpointName `
            -Name $batchName `
            -CSVData $([System.IO.File]::ReadAllBytes($migrationCsvPath)) `
            -TargetDeliveryDomain $m365Sub
        Write-Host "  Batch created." -ForegroundColor Green
    }
    catch {
        Write-Host "  Failed to create batch: $_" -ForegroundColor Red
        return
    }

    # Confirm and start
    Write-Host "`nReady to start the migration batch." -ForegroundColor Yellow
    Write-Host "Note: When the batch starts, mail users will be converted to mailboxes." -ForegroundColor Yellow
    Write-Host "Exchange licenses must be assigned within 30 days after this point.`n" -ForegroundColor Yellow

    $confirm = Read-Host "Start migration batch now? (y/N)"
    if ($confirm -eq 'y') {
        Start-MigrationBatch -Identity $batchName
        Write-Host "`nMigration batch started!" -ForegroundColor Green
        Write-Host "Monitor progress with: Get-MigrationBatch -Identity '$batchName'" -ForegroundColor Cyan
    }
    else {
        Write-Host "Batch created but not started. Start it later with:" -ForegroundColor Yellow
        Write-Host "  Start-MigrationBatch -Identity '$batchName'" -ForegroundColor White
    }

    Write-Host "`nStep 7 complete." -ForegroundColor Green
}

# ── Main menu ──────────────────────────────────────────────────────────────────

function Show-Menu {
    param([string]$Domain)

    while ($true) {
        Write-Host "`n╔═══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║     GSuite to M365 Migration — $Domain" -ForegroundColor Cyan
        Write-Host "╠═══════════════════════════════════════════════════════╣" -ForegroundColor Cyan
        Write-Host "║  1. Create GCP Project & Service Account             ║" -ForegroundColor White
        Write-Host "║  2. Manual Steps (Instructions)                      ║" -ForegroundColor White
        Write-Host "║  3. Enable Required Google APIs                      ║" -ForegroundColor White
        Write-Host "║  4. Create Subdomains (Google + M365)                ║" -ForegroundColor White
        Write-Host "║  5. Export Users from Google & Generate CSV           ║" -ForegroundColor White
        Write-Host "║  6. Create M365 Mail Users                           ║" -ForegroundColor White
        Write-Host "║  7. Create Migration Batch                           ║" -ForegroundColor White
        Write-Host "║                                                       ║" -ForegroundColor White
        Write-Host "║  Q. Quit                                              ║" -ForegroundColor White
        Write-Host "╚═══════════════════════════════════════════════════════╝" -ForegroundColor Cyan

        $choice = Read-Host "`nSelect step"

        switch ($choice) {
            '1' { Step1-CreateServiceAccount -Domain $Domain }
            '2' { Step2-ManualInstructions }
            '3' { Step3-EnableAPIs -Domain $Domain }
            '4' { Step4-CreateSubdomains -Domain $Domain }
            '5' { Step5-ExportUsersAndGenerateCSV -Domain $Domain }
            '6' { Step6-CreateMailUsers -Domain $Domain }
            '7' { Step7-CreateMigrationBatch -Domain $Domain }
            'Q' { return }
            'q' { return }
            default { Write-Host "Invalid selection." -ForegroundColor Red }
        }
    }
}

# ── Entry point ────────────────────────────────────────────────────────────────
$config = Get-MigrationConfig
$domain = if ($config.Domain) {
    $useSaved = Read-Host "Use saved domain '$($config.Domain)'? (Y/n)"
    if ($useSaved -eq 'n') { Read-Host "Enter domain (e.g. customkitchensinc.net)" }
    else { $config.Domain }
}
else {
    Read-Host "Enter domain (e.g. customkitchensinc.net)"
}

if (-not $domain) {
    Write-Host "Domain is required." -ForegroundColor Red
    exit 1
}

Show-Menu -Domain $domain
