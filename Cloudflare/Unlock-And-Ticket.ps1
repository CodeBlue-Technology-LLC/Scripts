<#
.SYNOPSIS
    Standalone script to unlock a domain at GoDaddy, retrieve auth code,
    and create a ConnectWise ticket for registrar transfer follow-up.
.DESCRIPTION
    For domains already migrated to Cloudflare. Looks up the Cloudflare
    zone automatically to get account/zone IDs for the ticket.
    Use -TicketOnly with -AuthCode to skip the unlock/auth retrieval steps.
.PARAMETER Domain
    The domain to unlock and create a ticket for.
.PARAMETER AuthCode
    Provide the auth code directly instead of retrieving it from GoDaddy.
.PARAMETER TicketOnly
    Skip unlock and auth code retrieval; only create the ConnectWise ticket.
.EXAMPLE
    .\Unlock-And-Ticket.ps1 -Domain "example.com"
.EXAMPLE
    .\Unlock-And-Ticket.ps1 -Domain "example.com" -TicketOnly -AuthCode "abc123"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Domain,

    [Parameter()]
    [string]$AuthCode,

    [Parameter()]
    [switch]$TicketOnly
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Import modules
Import-Module "$ScriptDir\Modules\GoDaddy.psm1" -Force
Import-Module "$ScriptDir\Modules\Cloudflare.psm1" -Force

# Load credentials
$Credentials = Import-Clixml "$ScriptDir\Config\credentials.xml"
$GoDaddyCreds = $Credentials.GoDaddy
$CloudflareCreds = $Credentials.Cloudflare
$ConnectWiseCreds = $Credentials.ConnectWise

# Step 1: Look up Cloudflare zone
Write-Host "`n=== Looking up Cloudflare zone for $Domain ===" -ForegroundColor Cyan
$zones = Get-CloudflareZones -Credentials $CloudflareCreds
$cfZone = $zones | Where-Object { $_.name -eq $Domain }
if (-not $cfZone) {
    Write-Error "Zone '$Domain' not found in Cloudflare. Has the domain been migrated?"
    exit 1
}
$cfAccountId = $cfZone.account.id
$cfAccountName = $cfZone.account.name
Write-Host "  Zone ID:      $($cfZone.id)" -ForegroundColor Green
Write-Host "  Account ID:   $cfAccountId" -ForegroundColor Green
Write-Host "  Account Name: $cfAccountName" -ForegroundColor Green

if (-not $TicketOnly) {
    # Step 2: Unlock domain
    Write-Host "`n=== Unlocking $Domain at GoDaddy ===" -ForegroundColor Cyan
    Unlock-GoDaddyDomain -Credentials $GoDaddyCreds -Domain $Domain -Confirm:$false

    # Step 3: Retrieve auth code (if not provided)
    if (-not $AuthCode) {
        Write-Host "`n=== Retrieving auth code ===" -ForegroundColor Cyan
        $AuthCode = Get-GoDaddyAuthCode -Credentials $GoDaddyCreds -Domain $Domain
        if ($AuthCode) {
            Write-Host "  Auth code retrieved successfully." -ForegroundColor Green
        }
        else {
            Write-Warning "Could not retrieve auth code. You may need to get it manually from GoDaddy."
        }
    }
}

$authCode = $AuthCode

# Step 4: Create ConnectWise ticket
Write-Host "`n=== Creating ConnectWise ticket ===" -ForegroundColor Cyan
if (-not $ConnectWiseCreds) {
    Write-Warning "No ConnectWise credentials configured. Skipping ticket creation."
    Write-Host "`nAuth Code: $authCode" -ForegroundColor Yellow
    exit 0
}

if (-not (Get-Module -ListAvailable -Name ConnectWiseManageAPI)) {
    Write-Host "Installing ConnectWiseManageAPI module..." -ForegroundColor Yellow
    Install-Module -Name ConnectWiseManageAPI -Scope CurrentUser -Force -AllowClobber
}

Import-Module ConnectWiseManageAPI -Force
Connect-CWM @ConnectWiseCreds

# Search for company using Cloudflare subaccount name
$companyName = $cfAccountName
Write-Host "  Searching for company: $companyName" -ForegroundColor Cyan
$companies = Get-CWMCompany -condition "name like '%$companyName%'"

$normalizedSearch = $companyName -replace '[^\w\s]', '' -replace '\s+', ' '
$matchingCompany = $companies | Where-Object {
    $normalizedName = $_.name -replace '[^\w\s]', '' -replace '\s+', ' '
    $normalizedName -eq $normalizedSearch
} | Select-Object -First 1

if (-not $matchingCompany) {
    Write-Warning "Company '$companyName' not found in ConnectWise."
    Write-Host "`nAuth Code: $authCode" -ForegroundColor Yellow
    exit 0
}

Write-Host "  Found company: $($matchingCompany.name) (ID: $($matchingCompany.identifier))" -ForegroundColor Cyan

# Build ticket description
$ticketDescription = "Transfer domain $Domain from GoDaddy to Cloudflare.`n`n"
$ticketDescription += "Domain: $Domain`n"
$ticketDescription += "Cloudflare Account ID: $cfAccountId`n"
$ticketDescription += "Cloudflare Zone ID: $($cfZone.id)`n"
if ($authCode) {
    $ticketDescription += "`nAuth Code: $authCode`n"
}
else {
    $ticketDescription += "`nAuth Code: COULD NOT RETRIEVE - obtain manually from GoDaddy`n"
}
$ticketDescription += "`nDomain has been unlocked at GoDaddy and is ready for transfer."

$ticket = New-CWMTicket -summary "Transfer $Domain to Cloudflare" `
    -company @{identifier = $matchingCompany.identifier} `
    -initialDescription $ticketDescription

Write-Host "`n=== Done ===" -ForegroundColor Green
Write-Host "  Ticket #$($ticket.id): $($ticket.summary)" -ForegroundColor Green
if ($authCode) {
    Write-Host "  Auth Code: $authCode" -ForegroundColor Green
}
