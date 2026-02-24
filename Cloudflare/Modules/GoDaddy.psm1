# GoDaddy API Module
# Read operations are safe; write operations require -Confirm

$script:GoDaddyApiBase = "https://api.godaddy.com/v1"

function Get-GoDaddyAuthHeader {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials
    )

    @{
        Authorization = "sso-key $($Credentials.ApiKey):$($Credentials.ApiSecret)"
    }
}

function Get-GoDaddyDomains {
    <#
    .SYNOPSIS
        Lists all domains in the GoDaddy account.
    .DESCRIPTION
        Read-only operation to retrieve all domains managed by the GoDaddy account.
    .PARAMETER Credentials
        Hashtable containing ApiKey and ApiSecret for GoDaddy.
    .EXAMPLE
        $creds = (Import-Clixml .\Config\credentials.xml).GoDaddy
        Get-GoDaddyDomains -Credentials $creds
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials
    )

    $headers = Get-GoDaddyAuthHeader -Credentials $Credentials

    try {
        $response = Invoke-RestMethod -Uri "$script:GoDaddyApiBase/domains" `
            -Headers $headers `
            -Method Get `
            -ContentType "application/json"

        return $response
    }
    catch {
        Write-Error "Failed to retrieve GoDaddy domains: $($_.Exception.Message)"
        throw
    }
}

function Get-GoDaddyDNSRecords {
    <#
    .SYNOPSIS
        Retrieves DNS records for a specific domain.
    .DESCRIPTION
        Read-only operation to get all DNS records for a domain in GoDaddy.
    .PARAMETER Credentials
        Hashtable containing ApiKey and ApiSecret for GoDaddy.
    .PARAMETER Domain
        The domain name to query (e.g., "example.com").
    .PARAMETER Type
        Optional. Filter by record type (A, AAAA, CNAME, MX, TXT, NS, SRV).
    .EXAMPLE
        $creds = (Import-Clixml .\Config\credentials.xml).GoDaddy
        Get-GoDaddyDNSRecords -Credentials $creds -Domain "example.com"
    .EXAMPLE
        Get-GoDaddyDNSRecords -Credentials $creds -Domain "example.com" -Type "MX"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials,

        [Parameter(Mandatory)]
        [string]$Domain,

        [Parameter()]
        [ValidateSet("A", "AAAA", "CNAME", "MX", "TXT", "NS", "SRV", "SOA", "CAA")]
        [string]$Type
    )

    $headers = Get-GoDaddyAuthHeader -Credentials $Credentials

    $uri = "$script:GoDaddyApiBase/domains/$Domain/records"
    if ($Type) {
        $uri += "/$Type"
    }

    try {
        $response = Invoke-RestMethod -Uri $uri `
            -Headers $headers `
            -Method Get `
            -ContentType "application/json"

        return $response
    }
    catch {
        Write-Error "Failed to retrieve DNS records for $Domain : $($_.Exception.Message)"
        throw
    }
}

function Set-GoDaddyNameservers {
    <#
    .SYNOPSIS
        Updates the nameservers for a domain at GoDaddy.
    .DESCRIPTION
        Sets custom nameservers for a domain. Requires confirmation.
        Used to point a GoDaddy domain to Cloudflare nameservers.
    .PARAMETER Credentials
        Hashtable containing ApiKey and ApiSecret for GoDaddy.
    .PARAMETER Domain
        The domain name to update (e.g., "example.com").
    .PARAMETER NameServers
        Array of nameserver hostnames (e.g., @("anna.ns.cloudflare.com", "bob.ns.cloudflare.com")).
    .EXAMPLE
        $creds = (Import-Clixml .\Config\credentials.xml).GoDaddy
        Set-GoDaddyNameservers -Credentials $creds -Domain "example.com" -NameServers @("anna.ns.cloudflare.com", "bob.ns.cloudflare.com")
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials,

        [Parameter(Mandatory)]
        [string]$Domain,

        [Parameter(Mandatory)]
        [string[]]$NameServers
    )

    $headers = Get-GoDaddyAuthHeader -Credentials $Credentials

    $body = @{
        nameServers = $NameServers
    } | ConvertTo-Json

    if ($PSCmdlet.ShouldProcess($Domain, "Update nameservers to: $($NameServers -join ', ')")) {
        try {
            $response = Invoke-RestMethod -Uri "$script:GoDaddyApiBase/domains/$Domain" `
                -Headers $headers `
                -Method Patch `
                -Body $body `
                -ContentType "application/json"

            Write-Host "Nameservers updated successfully for $Domain" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Error "Failed to update nameservers for $Domain : $($_.Exception.Message)"
            throw
        }
    }
}

function Unlock-GoDaddyDomain {
    <#
    .SYNOPSIS
        Unlocks a domain at GoDaddy for transfer.
    .DESCRIPTION
        Removes the registrar lock from a domain, allowing it to be transferred
        to another registrar. Requires confirmation.
    .PARAMETER Credentials
        Hashtable containing ApiKey and ApiSecret for GoDaddy.
    .PARAMETER Domain
        The domain name to unlock (e.g., "example.com").
    .EXAMPLE
        $creds = (Import-Clixml .\Config\credentials.xml).GoDaddy
        Unlock-GoDaddyDomain -Credentials $creds -Domain "example.com"
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials,

        [Parameter(Mandatory)]
        [string]$Domain
    )

    $headers = Get-GoDaddyAuthHeader -Credentials $Credentials

    if ($PSCmdlet.ShouldProcess($Domain, "Unlock domain for transfer")) {
        $maxAttempts = 3
        $retryDelaySec = 30

        for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
            try {
                # First retrieve current domain details (GoDaddy PATCH requires nameServers + subaccountId)
                $domainDetails = Invoke-RestMethod -Uri "$script:GoDaddyApiBase/domains/$Domain" `
                    -Headers $headers `
                    -Method Get `
                    -ContentType "application/json"

                # Check if domain is already unlocked
                if ($domainDetails.locked -eq $false) {
                    Write-Host "Domain '$Domain' is already unlocked." -ForegroundColor Green
                    return $true
                }

                # Check for transfer eligibility issues
                if ($domainDetails.transferProtected) {
                    Write-Warning "Domain '$Domain' is transfer-protected (may be within 60 days of registration or a prior transfer). Unlock may fail."
                }

                $body = @{
                    locked      = $false
                    nameServers = $domainDetails.nameServers
                    renewAuto   = $domainDetails.renewAuto
                }
                if ($domainDetails.subaccountId) {
                    $body.subaccountId = $domainDetails.subaccountId
                }
                $body = $body | ConvertTo-Json

                Invoke-RestMethod -Uri "$script:GoDaddyApiBase/domains/$Domain" `
                    -Headers $headers `
                    -Method Patch `
                    -Body $body `
                    -ContentType "application/json"

                Write-Host "Domain '$Domain' unlocked for transfer." -ForegroundColor Green
                return $true
            }
            catch {
                $statusCode = $null
                if ($_.Exception.Response) {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                }

                if ($statusCode -eq 422 -and $attempt -lt $maxAttempts) {
                    Write-Warning "Unlock attempt $attempt of $maxAttempts failed with 422. This can happen after a recent nameserver change. Retrying in ${retryDelaySec}s..."
                    Start-Sleep -Seconds $retryDelaySec
                    continue
                }

                if ($statusCode -eq 422) {
                    Write-Error "Failed to unlock domain $Domain after $maxAttempts attempts: GoDaddy returned 422 (Unprocessable Entity). The domain may be transfer-protected (within 60 days of registration/transfer), or cannot be unlocked via API for this TLD."
                }
                else {
                    Write-Error "Failed to unlock domain $Domain : $($_.Exception.Message)"
                }
                throw
            }
        }
    }
}

function Get-GoDaddyAuthCode {
    <#
    .SYNOPSIS
        Retrieves the EPP/auth code for a domain at GoDaddy.
    .DESCRIPTION
        Gets the authorization code needed to transfer a domain to another registrar.
        The domain should be unlocked before retrieving the auth code.
    .PARAMETER Credentials
        Hashtable containing ApiKey and ApiSecret for GoDaddy.
    .PARAMETER Domain
        The domain name to get the auth code for (e.g., "example.com").
    .EXAMPLE
        $creds = (Import-Clixml .\Config\credentials.xml).GoDaddy
        Get-GoDaddyAuthCode -Credentials $creds -Domain "example.com"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials,

        [Parameter(Mandatory)]
        [string]$Domain
    )

    $headers = Get-GoDaddyAuthHeader -Credentials $Credentials

    try {
        $response = Invoke-RestMethod -Uri "$script:GoDaddyApiBase/domains/$Domain" `
            -Headers $headers `
            -Method Get `
            -ContentType "application/json"

        if ($response.authCode) {
            return $response.authCode
        }
        else {
            Write-Warning "No auth code returned for $Domain. The domain may need to be unlocked first."
            return $null
        }
    }
    catch {
        Write-Error "Failed to retrieve auth code for $Domain : $($_.Exception.Message)"
        throw
    }
}

function Remove-GoDaddyPrivacy {
    <#
    .SYNOPSIS
        Removes domain privacy (Domains By Proxy) from a GoDaddy domain.
    .DESCRIPTION
        Disables the privacy/proxy service on a domain, exposing the real registrant
        contact information. This may be required before transferring a domain to
        another registrar, as the receiving registrar needs to see the actual
        registrant details.
    .PARAMETER Credentials
        Hashtable containing ApiKey and ApiSecret for GoDaddy.
    .PARAMETER Domain
        The domain name to remove privacy from (e.g., "example.com").
    .EXAMPLE
        $creds = (Import-Clixml .\Config\credentials.xml).GoDaddy
        Remove-GoDaddyPrivacy -Credentials $creds -Domain "example.com"
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials,

        [Parameter(Mandatory)]
        [string]$Domain
    )

    $headers = Get-GoDaddyAuthHeader -Credentials $Credentials

    if ($PSCmdlet.ShouldProcess($Domain, "Remove domain privacy (Domains By Proxy)")) {
        # First check if privacy is enabled
        try {
            $domainDetails = Invoke-RestMethod -Uri "$script:GoDaddyApiBase/domains/$Domain" `
                -Headers $headers `
                -Method Get `
                -ContentType "application/json"

            if ($domainDetails.privacy -eq $false) {
                Write-Host "Domain '$Domain' does not have privacy enabled." -ForegroundColor Green
                return $true
            }
        }
        catch {
            Write-Warning "Could not check privacy status for $Domain : $($_.Exception.Message)"
        }

        try {
            Invoke-RestMethod -Uri "$script:GoDaddyApiBase/domains/$Domain/privacy" `
                -Headers $headers `
                -Method Delete `
                -ContentType "application/json"

            Write-Host "Privacy removed from '$Domain'." -ForegroundColor Green
            return $true
        }
        catch {
            Write-Error "Failed to remove privacy from $Domain : $($_.Exception.Message)"
            throw
        }
    }
}

function Get-GoDaddyDomainDetails {
    <#
    .SYNOPSIS
        Retrieves domain details from GoDaddy, including nameservers and lock status.
    .PARAMETER Credentials
        Hashtable containing ApiKey and ApiSecret for GoDaddy.
    .PARAMETER Domain
        The domain name to query (e.g., "example.com").
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Credentials,

        [Parameter(Mandatory)]
        [string]$Domain
    )

    $headers = Get-GoDaddyAuthHeader -Credentials $Credentials

    try {
        $response = Invoke-RestMethod -Uri "$script:GoDaddyApiBase/domains/$Domain" `
            -Headers $headers `
            -Method Get `
            -ContentType "application/json"

        return $response
    }
    catch {
        Write-Error "Failed to retrieve domain details for $Domain : $($_.Exception.Message)"
        throw
    }
}

Export-ModuleMember -Function Get-GoDaddyDomains, Get-GoDaddyDNSRecords, Set-GoDaddyNameservers, Unlock-GoDaddyDomain, Get-GoDaddyAuthCode, Get-GoDaddyDomainDetails, Remove-GoDaddyPrivacy
