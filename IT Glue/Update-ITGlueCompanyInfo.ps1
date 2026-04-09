<#
.SYNOPSIS
    Identifies IT Glue companies missing domains and/or logos and fixes them.

.DESCRIPTION
    - Finds organizations without domains tracked in IT Glue.
    - Infers the correct domain from contact email addresses (requires 2+ contacts
      with the same email domain, excluding common providers).
    - Adds missing domains via Chrome-cookie web scraping (no official API exists).
    - Finds organizations without logos.
    - Fetches logos from Logo.dev by domain and attempts to upload via API,
      falling back to Chrome-cookie web scraping if needed.

.PARAMETER Test
    Specify a company name to run the full process against a single org.
    The script will query IT Glue for that company, check its domains/contacts/logo,
    and apply changes to just that one org. Use this to verify everything works
    before running against all companies.

.PARAMETER Reset
    Re-prompt for all stored credentials.

.PARAMETER ReportOnly
    Show what would be changed without making any modifications.

.PARAMETER SkipDomains
    Skip domain processing (only process logos).

.PARAMETER SkipLogos
    Skip logo processing (only process domains).

.PARAMETER DomainThreshold
    Minimum number of contacts sharing an email domain before it's added (default: 2).

.EXAMPLE
    .\Update-ITGlueCompanyInfo.ps1 -ReportOnly
    # Scans IT Glue and saves results to Config\scan-results.json

.EXAMPLE
    .\Update-ITGlueCompanyInfo.ps1
    # If a cached scan exists, loads it and applies changes (no re-scan)

.EXAMPLE
    .\Update-ITGlueCompanyInfo.ps1 -SkipLogos

.EXAMPLE
    .\Update-ITGlueCompanyInfo.ps1 -DomainThreshold 3
#>
[CmdletBinding()]
param(
    [string]$Test,
    [switch]$Reset,
    [switch]$ReportOnly,
    [switch]$SkipDomains,
    [switch]$SkipLogos,
    [switch]$ClearCache,
    [int]$DomainThreshold = 2
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# ── Common email domains to exclude ──────────────────────────────────────────
$ExcludedDomains = @(
    # Free email providers
    'gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'aol.com',
    'icloud.com', 'me.com', 'mac.com', 'live.com', 'msn.com',
    'protonmail.com', 'proton.me', 'ymail.com', 'googlemail.com',
    # ISP email domains
    'comcast.net', 'xfinity.com', 'att.net', 'sbcglobal.net',
    'verizon.net', 'charter.net', 'cox.net', 'spectrum.net',
    'centurylink.net', 'centurytel.net', 'windstream.net',
    'frontier.com', 'frontiernet.net', 'earthlink.net',
    'suddenlink.net', 'mediacombb.net', 'optimum.net',
    'optonline.net', 'twc.com', 'roadrunner.com', 'rr.com',
    'bellsouth.net', 'embarqmail.com', 'cableone.net',
    # Internal
    'codebluetechnology.com'
)

# ── Non-public TLDs to exclude ──────────────────────────────────────────────
$ExcludedTLDs = @(
    '.local', '.lan', '.internal', '.corp', '.home', '.localdomain',
    '.intranet', '.private', '.test', '.example', '.invalid', '.localhost',
    '.onmicrosoft.com'
)

# ── Module check ─────────────────────────────────────────────────────────────
if (-not (Get-Module -ListAvailable -Name ITGlueAPI)) {
    Write-Host "ITGlueAPI module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name ITGlueAPI -Scope CurrentUser -Force -AllowClobber
        Write-Host "ITGlueAPI installed successfully." -ForegroundColor Green
    }
    catch {
        throw "Failed to install ITGlueAPI module: $($_.Exception.Message)"
    }
}
Import-Module ITGlueAPI -Force -ErrorAction Stop

# ── Credential management ────────────────────────────────────────────────────
$CredentialsPath = "$ScriptDir\Config\credentials.xml"

function Initialize-Credentials {
    param([switch]$Reset)

    $configDir = Split-Path $CredentialsPath -Parent
    if (-not (Test-Path $configDir)) {
        New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    }

    $creds = $null
    if ((Test-Path $CredentialsPath) -and (-not $Reset)) {
        try { $creds = Import-Clixml -Path $CredentialsPath }
        catch { Write-Warning "Failed to load stored credentials, will prompt for new ones." }
    }

    $needsSave = $false
    if (-not $creds) { $creds = @{} }

    # ITGlue
    if (-not $creds.ITGlue) {
        $creds.ITGlue = @{}
        $needsSave = $true
    }
    if (-not $creds.ITGlue.BaseUri) {
        $itgBaseUri = Read-Host "ITGlue API Base URL [https://api.itglue.com]"
        if ([string]::IsNullOrWhiteSpace($itgBaseUri)) { $itgBaseUri = 'https://api.itglue.com' }
        $creds.ITGlue.BaseUri = $itgBaseUri
        $needsSave = $true
    }
    if (-not $creds.ITGlue.ApiKey) {
        $creds.ITGlue.ApiKey = Read-Host "ITGlue API Key"
        $needsSave = $true
    }
    if (-not $creds.ITGlue.Subdomain) {
        $creds.ITGlue.Subdomain = Read-Host "ITGlue Subdomain (e.g., yourcompany - used for web UI domain/logo upload)"
        $needsSave = $true
    }

    # Logo.dev
    if (-not $creds.LogoDev) {
        $creds.LogoDev = @{}
        $needsSave = $true
    }
    if (-not $creds.LogoDev.Token) {
        $creds.LogoDev.Token = Read-Host "Logo.dev API Token"
        $needsSave = $true
    }

    if ($needsSave) {
        $creds | Export-Clixml -Path $CredentialsPath
        Write-Host "Credentials saved to $CredentialsPath" -ForegroundColor Green
    }

    return $creds
}

# ── Selenium-based domain creation (handles SSO) ────────────────────────────
$script:seleniumDriver = $null

function Initialize-ITGlueBrowser {
    <#
    .SYNOPSIS
        Opens Edge via Selenium, navigates to ITGlue, and waits for user to complete SSO login.
        Returns the WebDriver instance.
    #>
    param([string]$ITGlueUrl)

    Import-Module Selenium -Force -ErrorAction Stop

    $edgeDriverDir = "c:\cbt\EdgeDriver"
    $edgeOptions = New-Object OpenQA.Selenium.Edge.EdgeOptions

    Write-Host "  Opening Edge browser..." -ForegroundColor Cyan
    $driver = New-Object OpenQA.Selenium.Edge.EdgeDriver($edgeDriverDir, $edgeOptions)

    Write-Host "  Navigating to ITGlue..." -ForegroundColor Cyan
    $driver.Navigate().GoToUrl("https://$ITGlueUrl")

    Write-Host ""
    Write-Host "  *** Please log in to ITGlue in the browser window ***" -ForegroundColor Yellow
    Write-Host "  Press ENTER here once you are logged in and see the ITGlue dashboard..." -ForegroundColor Yellow
    $null = Read-Host

    return $driver
}

function New-ITGDomain {
    <#
    .SYNOPSIS
        Adds a domain to an IT Glue organization via Selenium browser automation.
    #>
    param(
        [string]$OrgID,
        [string]$DomainName,
        [string]$ITGlueUrl,
        [OpenQA.Selenium.IWebDriver]$Driver
    )

    try {
        # Navigate to domain creation page
        $Driver.Navigate().GoToUrl("https://$ITGlueUrl/$OrgID/domains/new")
        Start-Sleep -Seconds 2

        # Find the domain name input field and fill it
        $nameField = $null
        foreach ($selector in @('domain_name', 'domain[name]')) {
            try {
                $nameField = $Driver.FindElement([OpenQA.Selenium.By]::Name($selector))
                if ($nameField) { break }
            } catch {}
        }
        if (-not $nameField) {
            try {
                $nameField = $Driver.FindElement([OpenQA.Selenium.By]::CssSelector('input[type="text"]'))
            } catch {}
        }

        if (-not $nameField) {
            Write-Warning "  Could not find domain name input field on page"
            return $false
        }

        $nameField.Clear()
        $nameField.SendKeys($DomainName)

        # Find and click the submit button
        $submitBtn = $null
        foreach ($selector in @('input[type="submit"]', 'button[type="submit"]', '.btn-primary')) {
            try {
                $submitBtn = $Driver.FindElement([OpenQA.Selenium.By]::CssSelector($selector))
                if ($submitBtn) { break }
            } catch {}
        }

        if (-not $submitBtn) {
            Write-Warning "  Could not find submit button on page"
            return $false
        }

        $submitBtn.Click()
        Start-Sleep -Seconds 3

        # Check for success
        $pageSource = $Driver.PageSource
        if ($pageSource -like "*Domain has been created successfully*" -or $pageSource -like "*$DomainName*") {
            return $true
        }
        else {
            Write-Warning "  Domain creation may have failed for '$DomainName' - check ITGlue manually"
            return $false
        }
    }
    catch {
        Write-Warning "  Failed to create domain '$DomainName': $($_.Exception.Message)"
        return $false
    }
}

# ── Logo functions ───────────────────────────────────────────────────────────
function Get-LogoForDomain {
    <#
    .SYNOPSIS
        Fetches a logo from Logo.dev for the given domain. Returns byte array or $null.
    #>
    param(
        [string]$Domain,
        [string]$Token
    )

    $url = "https://img.logo.dev/${Domain}?token=${Token}&size=256&format=png&fallback=404"

    try {
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing -ErrorAction Stop
        if ($response.StatusCode -eq 200 -and $response.Content.Length -gt 0) {
            return $response.Content
        }
    }
    catch {
        # 404 = no logo available, other errors = log
        if ($_.Exception.Response -and $_.Exception.Response.StatusCode.value__ -ne 404) {
            Write-Warning "  Logo.dev error for '$Domain': $($_.Exception.Message)"
        }
    }
    return $null
}

function Get-LogoFromWebsite {
    <#
    .SYNOPSIS
        Scrapes a company's website for images with "logo" in the URL/filename.
        Downloads all matches and returns the largest one as a byte array, or $null.
    #>
    param(
        [string]$Domain
    )

    $siteUrl = "https://$Domain"
    try {
        $page = Invoke-WebRequest -Uri $siteUrl -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
    }
    catch {
        Write-Verbose "  Could not fetch $siteUrl : $($_.Exception.Message)"
        return $null
    }

    # Find all image URLs from <img src="..."> and CSS/inline url(...) references
    $imgUrls = @()

    # img tags
    if ($page.Images) {
        $imgUrls += $page.Images | ForEach-Object { $_.src } | Where-Object { $_ }
    }

    # Also check srcset attributes and any other image-like URLs in the HTML
    $matches_ = [regex]::Matches($page.Content, '(?:src|srcset|href|content)\s*=\s*["'']([^"'']+?\.(?:png|jpg|jpeg|webp|svg))[^"'']*["'']', 'IgnoreCase')
    foreach ($m in $matches_) {
        $imgUrls += $m.Groups[1].Value
    }

    # Also check CSS url() references
    $cssMatches = [regex]::Matches($page.Content, 'url\(\s*["'']?([^"''\)]+?\.(?:png|jpg|jpeg|webp|svg))["'']?\s*\)', 'IgnoreCase')
    foreach ($m in $cssMatches) {
        $imgUrls += $m.Groups[1].Value
    }

    # Deduplicate all found image URLs
    $allImgUrls = $imgUrls | Select-Object -Unique

    # Priority 1: URLs containing logo-related keywords
    $logoKeywords = 'logo|brand|branding'
    $logoUrls = @($allImgUrls | Where-Object { $_ -imatch $logoKeywords })

    # Priority 2: apple-touch-icon and large favicons (usually clean square logos)
    $touchIcons = @()
    $touchIconMatches = [regex]::Matches($page.Content, '<link[^>]+rel\s*=\s*["''](?:apple-touch-icon|icon)["''][^>]+href\s*=\s*["'']([^"'']+)["'']', 'IgnoreCase')
    foreach ($m in $touchIconMatches) { $touchIcons += $m.Groups[1].Value }
    # Also check href before rel
    $touchIconMatches2 = [regex]::Matches($page.Content, '<link[^>]+href\s*=\s*["'']([^"'']+)["''][^>]+rel\s*=\s*["''](?:apple-touch-icon|icon)["'']', 'IgnoreCase')
    foreach ($m in $touchIconMatches2) { $touchIcons += $m.Groups[1].Value }
    $touchIcons = @($touchIcons | Select-Object -Unique)

    # Priority 3: Open Graph image (og:image)
    $ogImages = @()
    $ogMatches = [regex]::Matches($page.Content, '<meta[^>]+property\s*=\s*["'']og:image["''][^>]+content\s*=\s*["'']([^"'']+)["'']', 'IgnoreCase')
    foreach ($m in $ogMatches) { $ogImages += $m.Groups[1].Value }
    $ogMatches2 = [regex]::Matches($page.Content, '<meta[^>]+content\s*=\s*["'']([^"'']+)["''][^>]+property\s*=\s*["'']og:image["'']', 'IgnoreCase')
    foreach ($m in $ogMatches2) { $ogImages += $m.Groups[1].Value }
    $ogImages = @($ogImages | Select-Object -Unique)

    # Try each priority group in order
    $sourceLabel = ''
    $candidateUrls = @()
    if ($logoUrls.Count -gt 0) {
        $candidateUrls = $logoUrls
        $sourceLabel = 'logo/brand images'
    }
    elseif ($touchIcons.Count -gt 0) {
        $candidateUrls = $touchIcons
        $sourceLabel = 'apple-touch-icon/favicon'
    }
    elseif ($ogImages.Count -gt 0) {
        $candidateUrls = $ogImages
        $sourceLabel = 'og:image'
    }

    if ($candidateUrls.Count -eq 0) {
        Write-Verbose "  No logo images found on $siteUrl"
        return $null
    }

    Write-Host "  Found $($candidateUrls.Count) candidate $sourceLabel on website" -ForegroundColor White

    # Resolve relative URLs and download each, keep the largest
    $bestBytes = $null
    $bestSize = 0
    $bestUrl = ''

    foreach ($imgUrl in $candidateUrls) {
        # Skip SVGs - they don't convert well to raster for upload
        if ($imgUrl -imatch '\.svg(\?|$)') { continue }

        # Resolve relative URLs
        if ($imgUrl -match '^//') {
            $imgUrl = "https:$imgUrl"
        }
        elseif ($imgUrl -notmatch '^https?://') {
            $imgUrl = [System.Uri]::new([System.Uri]$siteUrl, $imgUrl).AbsoluteUri
        }

        try {
            $imgResponse = Invoke-WebRequest -Uri $imgUrl -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            if ($imgResponse.Content.Length -gt $bestSize) {
                $bestSize = $imgResponse.Content.Length
                $bestBytes = $imgResponse.Content
                $bestUrl = $imgUrl
            }
        }
        catch {
            Write-Verbose "  Failed to download $imgUrl : $($_.Exception.Message)"
        }
    }

    if ($bestBytes) {
        Write-Host "  Best logo from website ($sourceLabel): $bestUrl ($bestSize bytes)" -ForegroundColor White
        return $bestBytes
    }

    return $null
}

function Set-ITGlueLogo {
    <#
    .SYNOPSIS
        Attempts to set an organization's logo via IT Glue API PATCH.
        Returns $true if successful, $false if the API doesn't support it.
    #>
    param(
        [string]$OrgId,
        [byte[]]$LogoBytes,
        [string]$BaseUri,
        [string]$ApiKey
    )

    $base64 = [Convert]::ToBase64String($LogoBytes)

    # Try format 1: base64 string directly
    $data = @{
        type       = 'organizations'
        attributes = @{
            logo = $base64
        }
    }

    $headers = @{
        'x-api-key'    = $ApiKey
        'Content-Type' = 'application/vnd.api+json'
    }

    try {
        $body = @{ data = $data } | ConvertTo-Json -Depth 5
        $null = Invoke-RestMethod -Uri "$BaseUri/organizations/$OrgId" `
            -Method PATCH -Headers $headers -Body $body
        return $true
    }
    catch {
        Write-Verbose "  API logo upload (format 1) failed: $($_.Exception.Message)"
    }

    # Try format 2: content + file_name (like user avatars)
    $data.attributes.logo = @{
        content   = $base64
        file_name = "logo.png"
    }

    try {
        $body = @{ data = $data } | ConvertTo-Json -Depth 5
        $null = Invoke-RestMethod -Uri "$BaseUri/organizations/$OrgId" `
            -Method PATCH -Headers $headers -Body $body
        return $true
    }
    catch {
        Write-Verbose "  API logo upload (format 2) failed: $($_.Exception.Message)"
    }

    return $false
}

function Set-ITGlueLogoViaSelenium {
    <#
    .SYNOPSIS
        Uploads a logo to an IT Glue organization via Selenium browser automation.
        Saves logo to a temp file, navigates to org edit page, and uploads via file input.
    #>
    param(
        [string]$OrgID,
        [byte[]]$LogoBytes,
        [string]$ITGlueUrl,
        [OpenQA.Selenium.IWebDriver]$Driver
    )

    try {
        # Save logo to temp file for file input
        $tempLogo = "$env:TEMP\itglue_logo_$OrgID.png"
        [System.IO.File]::WriteAllBytes($tempLogo, $LogoBytes)

        # Navigate to org edit page
        $editUrl = "https://$ITGlueUrl/organizations/$OrgID/edit"
        $Driver.Navigate().GoToUrl($editUrl)

        # Wait for the logo file input, retrying if the page is slow to load
        $fileInput = $null
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            Start-Sleep -Seconds 5

            try {
                $fileInput = $Driver.FindElement([OpenQA.Selenium.By]::Id('organization_logo'))
            } catch {}

            if (-not $fileInput) {
                foreach ($selector in @('input.uploaded_file_input', 'input[type="file"][file_type="image/*"]', 'input[type="file"]')) {
                    try {
                        $fileInput = $Driver.FindElement([OpenQA.Selenium.By]::CssSelector($selector))
                        if ($fileInput) { break }
                    } catch {}
                }
            }

            if ($fileInput) { break }
            if ($attempt -lt 3) {
                Write-Host "  Page still loading, retrying... ($attempt/3)" -ForegroundColor Gray
            }
        }

        if (-not $fileInput) {
            Write-Host "  DEBUG: Current URL: $($Driver.Url)" -ForegroundColor Magenta
            $inputCount = 0
            try { $inputCount = $Driver.FindElements([OpenQA.Selenium.By]::TagName('input')).Count } catch {}
            Write-Host "  DEBUG: Total inputs on page: $inputCount" -ForegroundColor Magenta
            $pageTitle = ''
            try { $pageTitle = $Driver.Title } catch {}
            Write-Host "  DEBUG: Page title: $pageTitle" -ForegroundColor Magenta
            Write-Warning "  Could not find logo file input on org edit page"
            return $false
        }

        $fileInput.SendKeys($tempLogo)
        Start-Sleep -Seconds 2

        # Find and click save/submit
        $submitBtn = $null
        foreach ($selector in @('input[type="submit"]', 'button[type="submit"]', '.btn-primary', 'input[name="commit"]')) {
            try {
                $submitBtn = $Driver.FindElement([OpenQA.Selenium.By]::CssSelector($selector))
                if ($submitBtn) { break }
            } catch {}
        }

        if ($submitBtn) {
            # Scroll to button and use JavaScript click to avoid overlay interception
            $Driver.ExecuteScript("arguments[0].scrollIntoView(true);", $submitBtn)
            Start-Sleep -Milliseconds 500
            $Driver.ExecuteScript("arguments[0].click();", $submitBtn)
            Start-Sleep -Seconds 3
        }

        # Cleanup temp file
        Remove-Item $tempLogo -Force -ErrorAction SilentlyContinue

        return $true
    }
    catch {
        Write-Warning "  Selenium logo upload failed for org ${OrgID}: $($_.Exception.Message)"
        return $false
    }
}

# ── SSL tracker removal ─────────────────────────────────────────────────────
function Remove-ITGlueSSLTrackers {
    <#
    .SYNOPSIS
        Removes auto-created SSL certificate trackers for a given org/domain.
        IT Glue automatically adds SSL trackers when a domain is created.
    #>
    param(
        [string]$OrgId,
        [string]$DomainName,
        [string]$BaseUri,
        [string]$ApiKey
    )

    $headers = @{
        'x-api-key'    = $ApiKey
        'Content-Type' = 'application/vnd.api+json'
    }

    try {
        $response = Invoke-RestMethod -Uri "$BaseUri/ssl_certificates?filter[organization_id]=$OrgId&page[size]=100" `
            -Method GET -Headers $headers -ErrorAction Stop

        if (-not $response.data -or $response.data.Count -eq 0) {
            Write-Host "    No SSL trackers found" -ForegroundColor Gray
            return 0
        }

        # Match only auto-created trackers: exact domain match, no subdomains, no notes
        $matching = @($response.data | Where-Object {
            $cert = $_
            $certDomain = $cert.attributes.'common-name'
            if (-not $certDomain) { $certDomain = $cert.attributes.name }
            $hasNotes = $cert.attributes.notes -and $cert.attributes.notes.Trim() -ne ''
            ($certDomain -eq $DomainName -or $certDomain -eq "www.$DomainName" -or $certDomain -eq "*.$DomainName") -and -not $hasNotes
        })

        if ($matching.Count -eq 0) {
            Write-Host "    No matching SSL trackers for '$DomainName'" -ForegroundColor Gray
            return 0
        }

        $removed = 0
        foreach ($cert in $matching) {
            try {
                $certName = $cert.attributes.'common-name'
                if (-not $certName) { $certName = $cert.attributes.name }
                Invoke-RestMethod -Uri "$BaseUri/ssl_certificates/$($cert.id)" `
                    -Method DELETE -Headers $headers -ErrorAction Stop
                Write-Host "    Removed SSL tracker: $certName (ID: $($cert.id))" -ForegroundColor Green
                $removed++
            }
            catch {
                Write-Warning "    Failed to remove SSL tracker $($cert.id): $($_.Exception.Message)"
            }
        }
        return $removed
    }
    catch {
        Write-Warning "    Failed to query SSL trackers: $($_.Exception.Message)"
        return 0
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

Write-Host "`n=== IT Glue Company Info Updater ===" -ForegroundColor Cyan
if ($ReportOnly) { Write-Host "[REPORT ONLY MODE - no changes will be made]" -ForegroundColor Yellow }

$creds = Initialize-Credentials -Reset:$Reset

# Configure ITGlueAPI module
Add-ITGlueBaseURI -base_uri $creds.ITGlue.BaseUri
Add-ITGlueAPIKey -Api_Key $creds.ITGlue.ApiKey

$itgUrl = "$($creds.ITGlue.Subdomain).itglue.com"

# ── Step 1: Get all organizations + check domains/logos ──────────────────────
# (Skipped in -Test mode, which queries a single org directly)
$ScanFile = "$ScriptDir\Config\scan-results.json"
$loadedFromCache = $false

if ($ClearCache) {
    if (Test-Path $ScanFile) {
        Remove-Item $ScanFile -Force
        Write-Host "Scan cache cleared." -ForegroundColor Green
    } else {
        Write-Host "No cache file found." -ForegroundColor Gray
    }
    return
}

if ($Test) {
    # Test mode handles its own org lookup after this block
    $allOrgs = @()
    $orgsWithoutDomains = @()
    $orgsWithoutLogos = @()
    $cachedEmailDomains = @{}
    $loadedFromCache = $true
}
elseif (Test-Path $ScanFile) {
    Write-Host "`nLoading cached scan from $ScanFile..." -ForegroundColor Cyan
    try {
        $scanData = Get-Content $ScanFile -Raw | ConvertFrom-Json

        # Rebuild org-like objects from cached data
        $allOrgs = @($scanData.allOrgs)
        $orgsWithoutDomains = @($scanData.orgsWithoutDomains)
        $orgsWithoutLogos = @($scanData.orgsWithoutLogos)
        $cachedEmailDomains = @{}
        if ($scanData.emailDomainsByOrg) {
            foreach ($entry in $scanData.emailDomainsByOrg.PSObject.Properties) {
                $cachedEmailDomains[$entry.Name] = @{}
                foreach ($dp in $entry.Value.PSObject.Properties) {
                    $cachedEmailDomains[$entry.Name][$dp.Name] = [int]$dp.Value
                }
            }
        }

        $loadedFromCache = $true
        Write-Host "  Loaded: $($allOrgs.Count) orgs, $($orgsWithoutDomains.Count) without domains, $($orgsWithoutLogos.Count) without logos" -ForegroundColor Green
        Write-Host "  Scan date: $($scanData.scanDate)" -ForegroundColor Gray
    }
    catch {
        Write-Warning "Failed to load scan file, will re-scan: $($_.Exception.Message)"
    }
}

if (-not $loadedFromCache) {
    Write-Host "`nFetching all organizations..." -ForegroundColor Cyan
    $allOrgs = @()
    $page = 1
    do {
        $result = Get-ITGlueOrganizations -page_size 1000 -page_number $page
        if ($result.data) { $allOrgs += $result.data }
        $page++
    } while ($result.data -and $result.data.Count -eq 1000)

    Write-Host "  Found $($allOrgs.Count) organizations" -ForegroundColor White

    # Check domains and logos for each org
    Write-Host "Checking domains and logos for each organization..." -ForegroundColor Cyan
    $orgsWithoutDomains = @()
    $orgsWithoutLogos = @()

    $i = 0
    foreach ($org in $allOrgs) {
        $i++
        Write-Progress -Activity "Checking organizations" -Status "$i / $($allOrgs.Count) - $($org.attributes.name)" -PercentComplete (($i / $allOrgs.Count) * 100)

        # Check domains
        if (-not $SkipDomains) {
            try {
                $domains = Get-ITGlueDomains -filter_organization_id $org.id
                if (-not $domains.data -or $domains.data.Count -eq 0) {
                    $orgsWithoutDomains += $org
                }
            }
            catch {
                Write-Warning "  Could not check domains for '$($org.attributes.name)': $($_.Exception.Message)"
            }
        }

        # Check logo
        if (-not $SkipLogos) {
            if (-not $org.attributes.logo -or $org.attributes.logo -eq '') {
                $orgsWithoutLogos += $org
            }
        }

        # Small delay to avoid rate limiting
        Start-Sleep -Milliseconds 200
    }
    Write-Progress -Activity "Checking organizations" -Completed

    # Pre-scan email domains for orgs without domains (avoids re-querying contacts later)
    $cachedEmailDomains = @{}
    if (-not $SkipDomains -and $orgsWithoutDomains.Count -gt 0) {
        Write-Host "Pre-scanning contact email domains..." -ForegroundColor Cyan
        $i = 0
        foreach ($org in $orgsWithoutDomains) {
            $i++
            Write-Progress -Activity "Scanning contacts" -Status "$i / $($orgsWithoutDomains.Count) - $($org.attributes.name)" -PercentComplete (($i / $orgsWithoutDomains.Count) * 100)

            try {
                $contacts = @()
                $cPage = 1
                do {
                    $cResult = Get-ITGlueContacts -organization_id $org.id -page_size 1000 -page_number $cPage
                    if ($cResult.data) { $contacts += $cResult.data }
                    $cPage++
                } while ($cResult.data -and $cResult.data.Count -eq 1000)
            }
            catch {
                Write-Warning "  Could not get contacts for '$($org.attributes.name)': $($_.Exception.Message)"
                continue
            }

            # Count unique contacts per domain (not duplicate emails)
            $emailDomains = @{}
            foreach ($contact in $contacts) {
                $emails = $contact.attributes.'contact-emails'
                if (-not $emails) { continue }
                $contactDomains = @{}
                foreach ($email in $emails) {
                    $addr = $email.value
                    if (-not $addr -or $addr -notmatch '@(.+)$') { continue }
                    $domain = $Matches[1].ToLower().Trim()
                    if ($domain -in $ExcludedDomains) { continue }
                    if ($ExcludedTLDs | Where-Object { $domain.EndsWith($_) }) { continue }
                    $contactDomains[$domain] = $true
                }
                foreach ($d in $contactDomains.Keys) {
                    if (-not $emailDomains.ContainsKey($d)) { $emailDomains[$d] = 0 }
                    $emailDomains[$d]++
                }
            }
            if ($emailDomains.Count -gt 0) {
                $cachedEmailDomains[$org.id.ToString()] = $emailDomains
            }

            Start-Sleep -Milliseconds 200
        }
        Write-Progress -Activity "Scanning contacts" -Completed
    }

    # Save scan results for reuse
    Write-Host "Saving scan results to $ScanFile..." -ForegroundColor Cyan
    $scanData = @{
        scanDate           = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        domainThreshold    = $DomainThreshold
        allOrgs            = $allOrgs
        orgsWithoutDomains = $orgsWithoutDomains
        orgsWithoutLogos   = $orgsWithoutLogos
        emailDomainsByOrg  = $cachedEmailDomains
    }
    $scanData | ConvertTo-Json -Depth 10 | Set-Content $ScanFile -Encoding UTF8
    Write-Host "  Scan saved to $ScanFile" -ForegroundColor Green
}

# ── Test mode: run against a single named company ────────────────────────────
if ($Test) {
    Write-Host "`n--- TEST MODE: '$Test' ---" -ForegroundColor Magenta

    # Look up the org by name
    $searchResult = Get-ITGlueOrganizations -filter_name $Test -page_size 50
    $testOrg = $searchResult.data | Where-Object { $_.attributes.name -eq $Test } | Select-Object -First 1
    if (-not $testOrg) {
        # Try partial match
        $testOrg = $searchResult.data | Select-Object -First 1
    }
    if (-not $testOrg) {
        Write-Host "No organization found matching '$Test'" -ForegroundColor Red
        return
    }

    Write-Host "Found: $($testOrg.attributes.name) (ID: $($testOrg.id))" -ForegroundColor Magenta

    # Check domains
    $testDomains = Get-ITGlueDomains -filter_organization_id $testOrg.id
    $hasDomains = $testDomains.data -and $testDomains.data.Count -gt 0
    Write-Host "  Has domains: $hasDomains" -ForegroundColor White
    if ($hasDomains) {
        foreach ($d in $testDomains.data) { Write-Host "    - $($d.attributes.name)" -ForegroundColor Gray }
    }

    # Check logo
    $hasLogo = [bool]$testOrg.attributes.logo
    Write-Host "  Has logo:    $hasLogo" -ForegroundColor White

    # Check contacts / email domains
    $contacts = @()
    $cPage = 1
    do {
        $cResult = Get-ITGlueContacts -organization_id $testOrg.id -page_size 1000 -page_number $cPage
        if ($cResult.data) { $contacts += $cResult.data }
        $cPage++
    } while ($cResult.data -and $cResult.data.Count -eq 1000)

    # Count unique contacts per domain
    $emailDomains = @{}
    foreach ($contact in $contacts) {
        $emails = $contact.attributes.'contact-emails'
        if (-not $emails) { continue }
        $contactDomains = @{}
        foreach ($email in $emails) {
            $addr = $email.value
            if (-not $addr -or $addr -notmatch '@(.+)$') { continue }
            $domain = $Matches[1].ToLower().Trim()
            if ($domain -in $ExcludedDomains) { continue }
            if ($ExcludedTLDs | Where-Object { $domain.EndsWith($_) }) { continue }
            $contactDomains[$domain] = $true
        }
        foreach ($d in $contactDomains.Keys) {
            if (-not $emailDomains.ContainsKey($d)) { $emailDomains[$d] = 0 }
            $emailDomains[$d]++
        }
    }

    Write-Host "  Contacts: $($contacts.Count)" -ForegroundColor White
    $qualifyingDomains = $emailDomains.GetEnumerator() | Where-Object { $_.Value -ge $DomainThreshold } | Sort-Object Value -Descending
    if ($qualifyingDomains) {
        foreach ($d in $qualifyingDomains) {
            Write-Host "  Email domain: $($d.Key) ($($d.Value) contacts)" -ForegroundColor White
        }
    }
    else {
        Write-Host "  No qualifying email domains (need $DomainThreshold+ contacts)" -ForegroundColor Gray
    }

    $confirm = Read-Host "`nProceed with test on '$($testOrg.attributes.name)'? (y/n)"
    if ($confirm -ne 'y') {
        Write-Host "Test cancelled." -ForegroundColor Yellow
        return
    }

    # Set up cached data and narrow lists to just this org
    $cachedEmailDomains = @{}
    if ($emailDomains.Count -gt 0) {
        $cachedEmailDomains[$testOrg.id.ToString()] = $emailDomains
    }
    $orgsWithoutDomains = @()
    $orgsWithoutLogos = @()
    if (-not $hasDomains) { $orgsWithoutDomains = @($testOrg) }
    if (-not $hasLogo) { $orgsWithoutLogos = @($testOrg) }
    $SkipDomains = $false
    $SkipLogos = $false
    Write-Host ""
}

if (-not $Test) {
    Write-Host "  Organizations without domains: $($orgsWithoutDomains.Count)" -ForegroundColor $(if ($orgsWithoutDomains.Count -gt 0) { 'Yellow' } else { 'Green' })
    Write-Host "  Organizations without logos:   $($orgsWithoutLogos.Count)" -ForegroundColor $(if ($orgsWithoutLogos.Count -gt 0) { 'Yellow' } else { 'Green' })
}

# ── Step 2: Fix missing domains ──────────────────────────────────────────────
$domainsAdded = 0
$domainsFailed = 0
$domainsAddedList = @()  # Track added domains for SSL tracker cleanup

if (-not $SkipDomains -and $orgsWithoutDomains.Count -gt 0) {
    Write-Host "`n--- Processing Missing Domains ---" -ForegroundColor Cyan

    $i = 0
    foreach ($org in $orgsWithoutDomains) {
        $i++
        Write-Host "`n[$i/$($orgsWithoutDomains.Count)] $($org.attributes.name) (ID: $($org.id))" -ForegroundColor White

        # Use cached email domains from scan
        $orgId = $org.id.ToString()
        if (-not $cachedEmailDomains.ContainsKey($orgId) -or $cachedEmailDomains[$orgId].Count -eq 0) {
            Write-Host "  No qualifying email domains found" -ForegroundColor Gray
            continue
        }

        $emailDomains = $cachedEmailDomains[$orgId]

        # Filter domains meeting threshold
        $qualifyingDomains = $emailDomains.GetEnumerator() |
            Where-Object { $_.Value -ge $DomainThreshold } |
            Sort-Object Value -Descending

        if (-not $qualifyingDomains) {
            Write-Host "  No qualifying email domains found (need $DomainThreshold+ contacts)" -ForegroundColor Gray
            continue
        }

        # Check which domains already exist for this org
        $existingDomains = @()
        try {
            $domResult = Get-ITGlueDomains -filter_organization_id $org.id
            if ($domResult.data) {
                $existingDomains = $domResult.data | ForEach-Object { $_.attributes.name.ToLower().Trim() }
            }
        } catch {}

        foreach ($d in $qualifyingDomains) {
            if ($d.Key -in $existingDomains) {
                Write-Host "  Domain: $($d.Key) - already exists, skipping" -ForegroundColor Gray
                continue
            }

            Write-Host "  Domain: $($d.Key) ($($d.Value) contacts)" -ForegroundColor White

            if ($ReportOnly) {
                Write-Host "    [REPORT] Would add domain '$($d.Key)'" -ForegroundColor Yellow
                $domainsAdded++
                continue
            }

            # Initialize Selenium browser if needed
            if (-not $script:seleniumDriver) {
                $script:seleniumDriver = Initialize-ITGlueBrowser -ITGlueUrl $itgUrl
                if (-not $script:seleniumDriver) {
                    Write-Warning "  Cannot proceed with domain creation - browser session unavailable."
                    break
                }
            }

            $success = New-ITGDomain -OrgID $org.id -DomainName $d.Key `
                -ITGlueUrl $itgUrl -Driver $script:seleniumDriver

            if ($success) {
                Write-Host "    Added domain '$($d.Key)'" -ForegroundColor Green
                $domainsAdded++
                $domainsAddedList += @{ OrgId = $org.id; OrgName = $org.attributes.name; Domain = $d.Key }
            }
            else {
                $domainsFailed++
            }

            Start-Sleep -Milliseconds 500
        }
    }
}

# ── Step 3: Fix missing logos ────────────────────────────────────────────────
$logosSet = 0
$logosFailed = 0

if (-not $SkipLogos -and $orgsWithoutLogos.Count -gt 0) {
    Write-Host "`n--- Processing Missing Logos ---" -ForegroundColor Cyan

    $i = 0
    foreach ($org in $orgsWithoutLogos) {
        $i++
        Write-Host "`n[$i/$($orgsWithoutLogos.Count)] $($org.attributes.name) (ID: $($org.id))" -ForegroundColor White

        # Get domains for this org (may have just been added)
        $domainName = $null
        try {
            $domains = Get-ITGlueDomains -filter_organization_id $org.id
            if ($domains.data -and $domains.data.Count -gt 0) {
                $domainName = $domains.data[0].attributes.name
            }
        }
        catch {
            Write-Warning "  Could not check ITGlue domains: $($_.Exception.Message)"
        }

        # Fall back to email domain from contact scan (only if it met the threshold)
        if (-not $domainName) {
            $orgId = $org.id.ToString()
            if ($cachedEmailDomains.ContainsKey($orgId)) {
                $topDomain = $cachedEmailDomains[$orgId].GetEnumerator() |
                    Where-Object { $_.Value -ge $DomainThreshold } |
                    Sort-Object Value -Descending | Select-Object -First 1
                if ($topDomain) { $domainName = $topDomain.Key }
            }
        }

        if (-not $domainName) {
            Write-Host "  No domains found and no email domains to use, skipping logo" -ForegroundColor Gray
            continue
        }
        Write-Host "  Using domain: $domainName" -ForegroundColor White

        if ($ReportOnly) {
            Write-Host "    [REPORT] Would fetch and set logo from Logo.dev for '$domainName'" -ForegroundColor Yellow
            $logosSet++
            continue
        }

        # Fetch logo from Logo.dev
        $logoBytes = Get-LogoForDomain -Domain $domainName -Token $creds.LogoDev.Token
        if (-not $logoBytes) {
            Write-Host "  No logo on Logo.dev, checking website..." -ForegroundColor Yellow
            $logoBytes = Get-LogoFromWebsite -Domain $domainName
        }
        if (-not $logoBytes) {
            Write-Host "  No logo available for '$domainName'" -ForegroundColor Gray
            continue
        }

        Write-Host "  Logo fetched ($($logoBytes.Length) bytes)" -ForegroundColor White

        $uploaded = $false

        # IT Glue API silently ignores logo fields on PATCH, so use Selenium directly
        if (-not $script:seleniumDriver) {
            $script:seleniumDriver = Initialize-ITGlueBrowser -ITGlueUrl $itgUrl
            if (-not $script:seleniumDriver) {
                Write-Warning "  Cannot proceed with logo upload - browser session unavailable."
                break
            }
        }

        $uploaded = Set-ITGlueLogoViaSelenium -OrgID $org.id -LogoBytes $logoBytes `
            -ITGlueUrl $itgUrl -Driver $script:seleniumDriver

        if ($uploaded) {
            Write-Host "  Logo set successfully" -ForegroundColor Green
            $logosSet++
        }
        else {
            Write-Warning "  Failed to set logo"
            $logosFailed++
        }

        Start-Sleep -Milliseconds 500
    }
}

# ── Step 4: Remove auto-created SSL trackers ────────────────────────────────
$sslRemoved = 0

if ($domainsAddedList.Count -gt 0 -and -not $ReportOnly) {
    Write-Host "`n--- Removing Auto-Created SSL Trackers ---" -ForegroundColor Cyan

    $i = 0
    foreach ($entry in $domainsAddedList) {
        $i++
        Write-Host "  [$i/$($domainsAddedList.Count)] $($entry.OrgName) - $($entry.Domain)" -ForegroundColor White
        $removed = Remove-ITGlueSSLTrackers -OrgId $entry.OrgId -DomainName $entry.Domain `
            -BaseUri $creds.ITGlue.BaseUri -ApiKey $creds.ITGlue.ApiKey
        $sslRemoved += $removed
    }
}
elseif ($domainsAddedList.Count -gt 0 -and $ReportOnly) {
    Write-Host "`n  [REPORT] Would remove SSL trackers for $($domainsAddedList.Count) newly added domain(s)" -ForegroundColor Yellow
}

# ── Summary ──────────────────────────────────────────────────────────────────
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
if ($ReportOnly) {
    Write-Host "[REPORT ONLY - no changes were made]" -ForegroundColor Yellow
    Write-Host "Scan results cached. Run without -ReportOnly to apply changes." -ForegroundColor Yellow
}
else {
    Write-Host "Scan cache preserved. Delete $ScanFile manually to force a fresh scan." -ForegroundColor Gray
}
Write-Host "  Organizations scanned:    $($allOrgs.Count)"
if (-not $SkipDomains) {
    Write-Host "  Orgs missing domains:     $($orgsWithoutDomains.Count)"
    Write-Host "  Domains added:            $domainsAdded" -ForegroundColor $(if ($domainsAdded -gt 0) { 'Green' } else { 'White' })
    if ($domainsFailed -gt 0) {
        Write-Host "  Domains failed:           $domainsFailed" -ForegroundColor Red
    }
    if ($sslRemoved -gt 0) {
        Write-Host "  SSL trackers removed:     $sslRemoved" -ForegroundColor Green
    }
}
if (-not $SkipLogos) {
    Write-Host "  Orgs missing logos:       $($orgsWithoutLogos.Count)"
    Write-Host "  Logos set:                $logosSet" -ForegroundColor $(if ($logosSet -gt 0) { 'Green' } else { 'White' })
    if ($logosFailed -gt 0) {
        Write-Host "  Logos failed:             $logosFailed" -ForegroundColor Red
    }
}

# Close Selenium browser if open
if ($script:seleniumDriver) {
    try {
        $script:seleniumDriver.Quit()
        Write-Host "(Browser closed)" -ForegroundColor Gray
    }
    catch {}
    $script:seleniumDriver = $null
}
Write-Host ""
