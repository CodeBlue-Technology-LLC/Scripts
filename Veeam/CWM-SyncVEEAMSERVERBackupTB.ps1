<#
.SYNOPSIS
    Syncs Veeam backup storage usage to ConnectWise Manage agreement additions.

.DESCRIPTION
    Retrieves backup sizes from Veeam B&R, matches them to CWM agreements by
    company name, and creates or updates specified addition with
    the total TB used. Effective date is set to the 1st of next month.

.NOTES
    On first run, the script prompts for all connection details and stores them
    in Config\Credentials.xml (DPAPI-encrypted, user/machine-specific).

    Prerequisites:
      - Veeam.Backup.PowerShell module (installed on the Veeam B&R server)
      - ConnectWiseManageAPI module (installed automatically if missing)
#>

$ErrorActionPreference = 'Stop'

#import configuration, prompt to create if it doesn't exist
$ConfigPath = "Config\Credentials.xml"
if (-not (Test-Path $ConfigPath)) {
    Write-Host "No configuration found. Please enter the following details." -ForegroundColor Yellow

    $ServerName = Read-Host -Prompt "Enter the Veeam B&R server name (e.g., servername.domain.local)"
    $Credential = Get-Credential -Message "Enter credentials for the Veeam B&R server"

    $CWMServer     = Read-Host -Prompt "Enter CWM server URL (e.g., https://api-na.myconnectwise.net)"
    $CWMCompany    = Read-Host -Prompt "Enter CWM company name"
    $CWMPubKey     = Read-Host -Prompt "Enter CWM public key"
    $CWMPrivateKey = Read-Host -Prompt "Enter CWM private key"
    $CWMClientId   = Read-Host -Prompt "Enter CWM client ID"
    $ProductId     = Read-Host -Prompt "Enter CWM product ID for Veeam addition"

    @{
        Credential        = $Credential
        ServerName        = $ServerName
        ProductId         = [int]$ProductId
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
$Credential = $Config.Credential
$ServerName = $Config.ServerName
$ProductId = $Config.ProductId

try {
    #connection for ConnectWise powershell module
    if (Get-Module -ListAvailable -Name ConnectWiseManageAPI) { Connect-CWM @CWMConnectionInfo } else { Install-Module 'ConnectWiseManageAPI'; Connect-CWM @CWMConnectionInfo }

    #connection for Veeam B&R powershell module
    Import-Module Veeam.Backup.PowerShell
    Disconnect-VBRServer -ErrorAction SilentlyContinue
    Connect-VBRServer -Server $ServerName -Credential $Credential
}
catch {
    Write-Host "Failed to connect: $_" -ForegroundColor Red
    exit 1
}

try {
    #get all backups
    $backups = Get-VBRBackup

    #find space for all non-zero GB backups
    $usages = foreach ($backup in $backups | Where-Object {$_.vmcount -ne 0}) {
        $storages = $backup.GetAllChildrenStorages()
        $backupSize = 0
        foreach ($storage in $storages) {
            $backupSize += [Math]::Round($storage.Stats.BackupSize / 1GB, 1)
        }
        $backupSizeTB = $backupSize / 1024
        $backup | Select-Object Name, @{n='SizeTB';e={$backupSizeTB}}
    }

    #find all agreements
    $agreements = Get-CWMAgreement -condition '(type/name="IT Services Agreement" or type/name="Data Center Agreement") and agreementStatus="Active"' -all

    #find VEEAM-VM additions in agreements
    $additions = foreach ($agreement in $agreements) {
        Get-CWMAgreementAddition -parentid $agreement.id -condition 'product/identifier like "VEEAM-VM"' -all
    }

    #set effective date variable to start of next month
    $date = ((Get-Date).AddMonths(1)) | Get-Date -Format yyyy-MM-01
    $effectivedate = "$($date)T00:00:00Z"

    #build agreement lookup to avoid redundant API calls
    $agreementLookup = @{}
    $agreements | ForEach-Object { $agreementLookup[$_.id] = $_ }

    #loop through additions
    foreach ($addition in $additions) {
        $agreement = $agreementLookup[$addition.agreementId]
        foreach ($usage in $usages) {
            if ($usage.name -ilike "*$($agreement.company.name)*" -or $usage.name -ilike "*$($agreement.company.identifier)*") {
                #build the object
                $CreateParam = @{
                    AgreementID = $Agreement.id
                    product = @{id = $ProductId}
                    billCustomer = 'Billable'
                    quantity = $usage.sizeTB
                    effectivedate = $effectivedate
                }
                $existingaddition = Get-CWMAgreementAddition -parentId $agreement.id -condition 'product/identifier like "VEEAM-COLO-STORAGE-TB"'
                #if it doesn't exist already, create it
                if ([string]::IsNullOrEmpty($existingaddition)) {
                    New-CWMAgreementAddition @CreateParam
                }
                #otherwise edit it
                else {
                    Update-CWMAgreementAddition -parentid $agreement.id -id $existingaddition.id -Operation 'replace' -Path 'quantity' -Value $usage.sizeTB
                }
                $CreateParam = $null
            }
        }
    }
}
catch {
    Write-Host "Error during sync: $_" -ForegroundColor Red
    exit 1
}