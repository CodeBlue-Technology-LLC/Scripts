<#
.SYNOPSIS
    Shared helpers for the MySonicWall (api.mysonicwall.com) REST API.

.DESCRIPTION
    Dot-source this file. Authentication is a single `X-API-KEY` header carrying the
    MySonicWall API key (MySonicWall > My Account > API Key) -- unlike the Capture
    Client console, there is no token exchange.

    The live OpenAPI definition is served, unauthenticated, at
        https://api.mysonicwall.com/swagger/v1/swagger.json
    and is the only complete endpoint reference. Note that PowerShell 5.1's
    ConvertFrom-Json cannot parse it: it contains both `totalprice` and `totalPrice`,
    which collide in a case-insensitive dictionary. Read it with another tool.

.NOTES
    Run under Windows PowerShell 5.1 (powershell.exe), not PowerShell 7.

    Product IDs seen in this tenant:
        401 = Capture Client   (services: Advanced 6732, Premier 6759)
        402 = Cloud App Security
#>

$script:MswApiBaseUri = 'https://api.mysonicwall.com'

function New-MswHeader {
    param([Parameter(Mandatory)][System.Security.SecureString]$ApiKey)

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ApiKey)
    try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }

    return @{ 'X-API-KEY' = $plain; 'Accept' = 'application/json' }
}

function Get-MswMsspProductService {
    <#
        POST /api/mssp/getmsspproductservices

        Every tenant under the MSSP, each with a serialList of registered products and,
        per product, a serviceList carrying node counts and the monthly-billing flags:

            msspStatus                 ACTIVE when the product is on MSSP monthly
            removeMsspMonthlyAllowed   YES when "Remove from MSSP Monthly" is offered
            showConverttoMSSP          YES when the reverse action is offered
            nodeCount / usedNodeCount  licensed vs actually consumed

        A tenant with no monthly products comes back with no serialList at all. That is
        authoritative -- querying such a tenant individually returns the same thing --
        so absence here means "not on monthly billing", not "not fetched".

        -TenantId narrows the call to specific MySonicWall tenants (ProductGroupIDs).
    #>
    param(
        [Parameter(Mandatory)][System.Security.SecureString]$ApiKey,
        [string[]]$TenantId,
        [int]$PageSize = 500
    )

    $body = @{ pageno = '1'; pagesize = "$PageSize" }
    if ($TenantId) { $body['tenantids'] = ($TenantId -join ',') }

    $response = Invoke-RestMethod -Uri "$script:MswApiBaseUri/api/mssp/getmsspproductservices" `
        -Headers (New-MswHeader -ApiKey $ApiKey) -Method Post `
        -Body (ConvertTo-Json $body) -ContentType 'application/json'

    if ($response.operationStatus -eq 'NOK') {
        throw "MySonicWall rejected the product-service query: $($response.msgCode) $($response.msgDesc)"
    }
    return $response.content
}

function Get-MswMonthlyProductRow {
    <#
        Flattens Get-MswMsspProductService into one row per tenant/product/service.

        -ProductId is a hard filter, not a convenience: every caller in this repo is
        expected to name the product it means. Capture Client is 401; Cloud App
        Security is 402 and must never be swept up by a Capture Client operation.
    #>
    param(
        [Parameter(Mandatory)][object]$Content,
        [Parameter(Mandatory)][string]$ProductId
    )

    $rows = @()
    foreach ($tenant in @($Content.mssplist)) {
        if (-not $tenant.PSObject.Properties['serialList']) { continue }
        foreach ($product in @($tenant.serialList)) {
            if ([string]$product.productID -ne $ProductId) { continue }
            foreach ($service in @($product.serviceList)) {
                $rows += [pscustomobject]@{
                    TenantName    = $tenant.tenantName
                    TenantId      = [string]$tenant.tenantId
                    MsspName      = $tenant.msspName
                    MsspId        = [string]$tenant.msspId
                    SerialNumber  = [string]$product.serialNumber
                    FriendlyName  = [string]$product.friendlyName
                    ProductName   = [string]$product.ProductName
                    ProductId     = [string]$product.productID
                    ServiceName   = [string]$service.serviceName
                    ServiceId     = [string]$service.serviceID
                    NodeCount     = [int]$service.nodeCount
                    UsedNodeCount = [int]$service.usedNodeCount
                    MsspStatus    = [string]$service.msspStatus
                    StartDate     = ([string]$service.StartDate).Trim()
                    ExpiryDate    = [string]$service.expiryDate
                    RemoveAllowed = [string]$service.removeMsspMonthlyAllowed
                    EditAllowed   = [string]$service.editAllowed
                }
            }
        }
    }
    return $rows
}
