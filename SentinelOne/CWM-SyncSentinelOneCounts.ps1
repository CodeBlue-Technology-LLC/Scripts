<#
.SYNOPSIS
    Syncs SentinelOne endpoint counts to ConnectWise Manage agreement additions.

.DESCRIPTION
    Retrieves active site endpoint counts from two SentinelOne instances (EDR and MDR),
    matches them to CWM agreements by company name, and updates existing additions with
    the current activeLicenses count. Generates a CSV report of all changes.

.NOTES
    On first run, the script prompts for all connection details and stores them
    in Config\Credentials.xml (DPAPI-encrypted, user/machine-specific).

    Prerequisites:
      - ConnectWiseManageAPI module (installed automatically if missing)

    Additions must already exist on agreements before running this script.
    This script only updates quantities — it does not create new additions.
#>

$ErrorActionPreference = 'Stop'

#region Configuration
$ConfigPath = Join-Path $PSScriptRoot "Config\Credentials.xml"
if (-not (Test-Path $ConfigPath)) {
    Write-Host "No configuration found. Please enter the following details." -ForegroundColor Yellow

    $EDRURL    = Read-Host -Prompt "Enter SentinelOne EDR URL (e.g., https://usea1-cw04edr.sentinelone.net)"
    $EDRAPIKey = Read-Host -Prompt "Enter SentinelOne EDR API Token"
    $MDRURL    = Read-Host -Prompt "Enter SentinelOne MDR URL (e.g., https://usea1-cw04mdr.sentinelone.net)"
    $MDRAPIKey = Read-Host -Prompt "Enter SentinelOne MDR API Token"

    $CWMServer     = Read-Host -Prompt "Enter CWM server URL (e.g., https://api-na.myconnectwise.net)"
    $CWMCompany    = Read-Host -Prompt "Enter CWM company name"
    $CWMPubKey     = Read-Host -Prompt "Enter CWM public key"
    $CWMPrivateKey = Read-Host -Prompt "Enter CWM private key"
    $CWMClientId   = Read-Host -Prompt "Enter CWM client ID"

    $configDir = Join-Path $PSScriptRoot "Config"
    if (-not (Test-Path $configDir)) { New-Item -ItemType Directory -Path $configDir | Out-Null }

    @{
        EDR = @{
            URL               = $EDRURL.TrimEnd('/')
            APIKey            = $EDRAPIKey
            ProductIdentifier = "SENT-ONE-CTRL"
        }
        MDR = @{
            URL               = $MDRURL.TrimEnd('/')
            APIKey            = $MDRAPIKey
            ProductIdentifier = "SENT-ONE-MDR"
        }
        CWMConnectionInfo = @{
            Server     = $CWMServer
            Company    = $CWMCompany
            pubKey     = $CWMPubKey
            privateKey = $CWMPrivateKey
            clientId   = $CWMClientId
        }
    } | Export-Clixml -Path $ConfigPath

    Write-Host "Configuration saved to $ConfigPath" -ForegroundColor Green
}
$Config = Import-Clixml -Path $ConfigPath
$CWMConnectionInfo = $Config.CWMConnectionInfo
#endregion

#region Functions
function Get-S1Sites {
    param(
        [string]$BaseURL,
        [string]$APIKey
    )

    $allSites = @()
    $cursor = $null
    $headers = @{ 'Authorization' = "ApiToken $APIKey" }

    do {
        $uri = "$BaseURL/web/api/v2.1/sites?states=active&limit=1000"
        if ($cursor) { $uri += "&cursor=$cursor" }

        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
        $allSites += $response.data.sites
        $cursor = $response.pagination.nextCursor
    } while ($cursor)

    return $allSites
}
#endregion

#region Connect
try {
    if (Get-Module -ListAvailable -Name ConnectWiseManageAPI) {
        Connect-CWM @CWMConnectionInfo
    } else {
        Install-Module 'ConnectWiseManageAPI' -Force
        Connect-CWM @CWMConnectionInfo
    }
}
catch {
    Write-Host "Failed to connect to ConnectWise Manage: $_" -ForegroundColor Red
    exit 1
}
#endregion

#region Main
try {
    # Fetch all active IT Services and Data Center agreements
    $agreements = Get-CWMAgreement -condition '(type/name="IT Services Agreement" or type/name="Data Center Agreement") and agreementStatus="Active"' -all

    # Build lookup: company name -> agreement, preferring IT Services Agreement
    $agreementByCompany = @{}
    foreach ($agreement in $agreements) {
        $companyName = $agreement.company.name
        if ($agreementByCompany.ContainsKey($companyName)) {
            # Prefer IT Services Agreement over Data Center Agreement
            if ($agreement.type.name -eq "IT Services Agreement") {
                $agreementByCompany[$companyName] = $agreement
            }
        } else {
            $agreementByCompany[$companyName] = $agreement
        }
    }

    $reportData = @()
    $date = ((Get-Date).AddMonths(1)) | Get-Date -Format yyyy-MM-01
    $effectiveDate = "${date}T00:00:00Z"

    foreach ($instanceKey in @('EDR', 'MDR')) {
        $instance = $Config.$instanceKey
        Write-Host "`nProcessing SentinelOne $instanceKey ($($instance.URL))..." -ForegroundColor Cyan

        try {
            $sites = Get-S1Sites -BaseURL $instance.URL -APIKey $instance.APIKey
        }
        catch {
            Write-Host "Failed to connect to SentinelOne $instanceKey`: $_" -ForegroundColor Red
            continue
        }

        Write-Host "  Found $($sites.Count) active sites" -ForegroundColor Gray

        foreach ($site in $sites) {
            try {
                # Match site name to CWM agreement (case-insensitive via hashtable default)
                $agreement = $agreementByCompany[$site.name]

                if (-not $agreement) {
                    Write-Warning "No CWM agreement match for S1 site: '$($site.name)' [$instanceKey]"
                    continue
                }

                # Find existing addition for this product
                $additions = Get-CWMAgreementAddition -AgreementID $agreement.id -all
                $existing = $additions | Where-Object { $_.product.identifier -eq $instance.ProductIdentifier }

                if (-not $existing) {
                    Write-Warning "No existing '$($instance.ProductIdentifier)' addition on agreement for '$($site.name)' [$instanceKey] — skipping (additions must be created manually)"
                    continue
                }

                $qtyBefore = $existing.quantity
                $qtyAfter = $site.activeLicenses

                # Update quantity if changed
                if ($qtyBefore -ne $qtyAfter) {
                    $Update = @{
                        AgreementID = $agreement.id
                        AdditionID  = $existing.id
                        Operation   = 'replace'
                        Path        = 'quantity'
                    }
                    Update-CWMAgreementAddition @Update -Value $qtyAfter | Out-Null
                    Write-Host "  Updated '$($site.name)' [$instanceKey]: $qtyBefore -> $qtyAfter" -ForegroundColor Green
                } else {
                    Write-Host "  No change for '$($site.name)' [$instanceKey]: $qtyBefore" -ForegroundColor Gray
                }

                $reportData += [PSCustomObject]@{
                    'Company Name'    = $site.name
                    'Instance'        = $instanceKey
                    'Product'         = $instance.ProductIdentifier
                    'Quantity Before' = $qtyBefore
                    'Quantity After'  = $qtyAfter
                }
            }
            catch {
                Write-Warning "Failed to sync site '$($site.name)' [$instanceKey]: $_"
                continue
            }
        }
    }

    # Export CSV report
    if ($reportData.Count -gt 0) {
        $reportPath = Join-Path $PSScriptRoot "Reports"
        if (-not (Test-Path $reportPath)) { New-Item -ItemType Directory -Path $reportPath | Out-Null }

        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $csvFile = Join-Path $reportPath "${timestamp}_SentinelOne_Sync_Report.csv"

        $reportData | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
        Write-Host "`nReport saved to $csvFile" -ForegroundColor Green
    } else {
        Write-Host "`nNo additions were found to update." -ForegroundColor Yellow
    }

    Write-Host "`nSync complete." -ForegroundColor Cyan
}
catch {
    Write-Host "Error during sync: $_" -ForegroundColor Red
    exit 1
}
#endregion
