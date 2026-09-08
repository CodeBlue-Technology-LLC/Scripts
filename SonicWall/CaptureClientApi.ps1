<#
.SYNOPSIS
    Shared helpers for the SonicWall Capture Client management API.

.DESCRIPTION
    Dot-source this file. It provides authentication against the Capture Client
    console and the handful of read calls the reporting scripts need.

    Authentication is a three-legged handshake, because the console itself has no
    long-lived API key -- console logins are OAUTH via MySonicWall:

      1. GET  https://api.mysonicwall.com/api/hgms/get-cloud-tenants
         Header: X-API-KEY: <MySonicWall API key>
         Returns every MySonicWall tenant and which cloud services each has.
         Used once to find a ProductGroupID that has Capture Client enabled.

      2. POST https://api.mysonicwall.com/api/generate-cscaccesscode
         Header: X-API-KEY: <MySonicWall API key>
         Body:   { "tenantId": "<ProductGroupID>", "tileName": "ISNSMSAFEENABLED" }
         Returns a single-use access code that is valid for FIVE MINUTES.

      3. POST https://<console>/api/auth/getApiToken?cscaccesscode=<code>
         Returns a JWT valid for SEVEN DAYS.

    The JWT is then sent on every console call as a bare `Authorization: <token>`
    header -- NOT `Bearer <token>`. Bearer happens to be accepted today, but the
    vendor documents the bare form, so that is what is used here.

    Scope note: the JWT inherits the scope of the MySonicWall user behind the API
    key, not the ProductGroupID passed in step 2. Passing a different customer's
    ProductGroupID does not mint a token for that customer's console tenant -- it
    returns the same token every time. Cross-tenant reads are done instead by adding
    ?tenantId=<console tenant GUID> to each call. That GUID is a different identifier
    from the MySonicWall ProductGroupID and comes from /api/tenant/list/verbose.

.NOTES
    Run under Windows PowerShell 5.1 (powershell.exe), not PowerShell 7.

    Region hosts: captureclient-36.sonicwall.com (US), captureclient-36eu.sonicwall.com (EU).

    Endpoint reference: https://<console>/api/docs/
#>

$script:MswBaseUri = 'https://api.mysonicwall.com'

function ConvertTo-PlainText {
    <# SecureString -> String, for use as a request header only. #>
    param([Parameter(Mandatory)][System.Security.SecureString]$Secure)

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

function Get-MswCloudTenant {
    <#
        Every MySonicWall tenant visible to the key, with its cloud service tiles.
        This is a large response (roughly 600 KB); call it once per run.
    #>
    param([Parameter(Mandatory)][System.Security.SecureString]$ApiKey)

    $headers = @{ 'X-API-KEY' = (ConvertTo-PlainText -Secure $ApiKey); 'Accept' = 'application/json' }
    $response = Invoke-RestMethod -Uri "$script:MswBaseUri/api/hgms/get-cloud-tenants" -Headers $headers -Method Get
    return @($response.content.arrTenants)
}

function Get-CaptureClientProductGroupId {
    <#
        A ProductGroupID that has the Capture Client tile enabled. Any one will do --
        it only has to satisfy the access-code call. Returns $null if the key can see
        no Capture Client tenant at all, which means the key is scoped too narrowly.
    #>
    param([Parameter(Mandatory)][object[]]$Tenant)

    foreach ($item in $Tenant) {
        if (-not $item.PSObject.Properties['cloudServices']) { continue }
        foreach ($service in @($item.cloudServices)) {
            if ($service.serviceName -eq 'Capture Client' -and $service.isEnabled) {
                return [string]$item.productGroupID
            }
        }
    }
    return $null
}

function Connect-CaptureClient {
    <#
        Runs legs 2 and 3 of the handshake and returns a session object:
            BaseUri  - console base URI, no trailing slash
            Headers  - hashtable to splat onto every subsequent call
            Expires  - [datetime] the JWT stops working (7 days out)

        The access code expires in five minutes, so the two calls are made back to
        back and the code is never persisted.
    #>
    param(
        [Parameter(Mandatory)][System.Security.SecureString]$ApiKey,
        [Parameter(Mandatory)][string]$ProductGroupId,
        [string]$BaseUri = 'https://captureclient-36.sonicwall.com'
    )

    $BaseUri = $BaseUri.TrimEnd('/')
    $mswHeaders = @{ 'X-API-KEY' = (ConvertTo-PlainText -Secure $ApiKey) }
    $body = ConvertTo-Json @{ tenantId = $ProductGroupId; tileName = 'ISNSMSAFEENABLED' }

    $codeResponse = Invoke-RestMethod -Uri "$script:MswBaseUri/api/generate-cscaccesscode" `
        -Headers $mswHeaders -Method Post -Body $body -ContentType 'application/json'

    if (-not $codeResponse.content.accessCode) {
        throw "MySonicWall returned no access code for ProductGroupID $ProductGroupId ($($codeResponse.msgCode))."
    }

    # The code is base64 and routinely carries '=' padding, which has to survive the
    # query string intact.
    $encoded = [Uri]::EscapeDataString([string]$codeResponse.content.accessCode)
    $tokenResponse = Invoke-RestMethod -Uri "$BaseUri/api/auth/getApiToken?cscaccesscode=$encoded" -Method Post

    if (-not $tokenResponse.token) { throw 'Capture Client returned no API token.' }

    return [pscustomobject]@{
        BaseUri = $BaseUri
        Headers = @{ Authorization = [string]$tokenResponse.token }
        Email   = [string]$tokenResponse.email
        Expires = (Get-Date).AddDays(7)
    }
}

function Get-CaptureClientTenant {
    <#
        Every console tenant the authenticated user can see, with per-tenant stats.

        The page size is fixed at 10 server-side; sizePerPage is accepted and then
        ignored, so the only way through the list is to walk pages until the reported
        `count` is satisfied.
    #>
    param([Parameter(Mandatory)][object]$Session)

    $tenants = @()
    $page = 1
    $expected = $null

    do {
        $uri = "$($Session.BaseUri)/api/tenant/list/verbose?page=$page&sizePerPage=10"
        $response = Invoke-RestMethod -Uri $uri -Headers $Session.Headers -Method Get

        if ($null -eq $expected) { $expected = [int]$response.count }
        $batch = @($response.data)
        if ($batch.Count -eq 0) { break }

        $tenants += $batch
        $page++
    }
    while ($tenants.Count -lt $expected -and $page -le 200)

    if ($tenants.Count -ne $expected) {
        Write-Warning "Tenant list returned $($tenants.Count) of $expected rows."
    }
    return $tenants
}

function Get-CaptureClientCleanerAuthCode {
    <#
        The "Authentication Code for Tenant" shown under
        Management > Client Installers > Utilities. It is the 64-hex-character value
        passed to SonicWallCaptureClientCleanerUtility.exe -authCode, and it is
        per-tenant, not per-device.
    #>
    param(
        [Parameter(Mandatory)][object]$Session,
        [Parameter(Mandatory)][string]$TenantId
    )

    $uri = "$($Session.BaseUri)/api/tenant/cleanerAuthCode?tenantId=$([Uri]::EscapeDataString($TenantId))"
    $response = Invoke-RestMethod -Uri $uri -Headers $Session.Headers -Method Get
    return [string]$response.cleanerAuthCode
}

function Get-NormalizedCompanyName {
    <#
        Company-name normalization, applied identically to both sides of the join.
        Deliberately conservative: over-aggressive stripping collapses genuinely
        different companies onto one string and shows up as false ambiguity.

          - lowercase
          - drop apostrophes outright, so O'Neal matches ONeal and Specialty's
            matches Specialtys (turning them into spaces would split the token)
          - remaining punctuation becomes a separator
          - merge runs of two or more single-letter tokens (P.C. -> pc)
          - drop standalone legal-suffix tokens, word-boundary matched
          - drop a leading article
    #>
    param(
        [string]$Name,
        [string[]]$Suffix = @('inc', 'llc', 'ltd', 'corp', 'co', 'pc', 'pllc', 'plc', 'llp', 'lp'),
        [string[]]$Article = @('the'),
        [bool]$MergeInitials = $true
    )

    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }

    $text = $Name.ToLowerInvariant()
    # Straight apostrophe (doubled to escape it) plus both curly variants. Written as
    # \u escapes so the file stays pure ASCII -- PowerShell 5.1 reads a BOM-less UTF-8
    # script as ANSI and would otherwise mangle the literals.
    $text = $text -replace '[''\u2018\u2019]', ''
    $text = $text -replace '[^\w\s]', ' '
    $text = ($text -replace '\s+', ' ').Trim()

    if ($MergeInitials) {
        $tokens = @($text -split ' ' | Where-Object { $_ })
        $merged = New-Object System.Collections.Generic.List[string]
        $run = New-Object System.Collections.Generic.List[string]

        $flushRun = {
            if ($run.Count -ge 2) { $merged.Add(($run -join '')) }
            elseif ($run.Count -eq 1) { $merged.Add($run[0]) }
            $run.Clear()
        }

        foreach ($token in $tokens) {
            if ($token.Length -eq 1 -and $token -match '^[a-z]$') { $run.Add($token) }
            else { & $flushRun; $merged.Add($token) }
        }
        & $flushRun
        $text = $merged -join ' '
    }

    if ($Suffix -and $Suffix.Count -gt 0) {
        $pattern = ($Suffix | Where-Object { $_ } | ForEach-Object { [regex]::Escape($_.ToLowerInvariant()) }) -join '|'
        if ($pattern) { $text = $text -replace "\b(?:$pattern)\b", ' ' }
    }
    $text = ($text -replace '\s+', ' ').Trim()

    if ($Article -and $Article.Count -gt 0) {
        $pattern = ($Article | Where-Object { $_ } | ForEach-Object { [regex]::Escape($_.ToLowerInvariant()) }) -join '|'
        if ($pattern) { $text = $text -replace "^(?:$pattern)\b", ' ' }
    }

    return ($text -replace '\s+', ' ').Trim()
}

function Get-BareTenantName {
    <#
        Console tenant names carry a provisioning suffix -- "Acme Corp-CC000002A95E".
        It never appears in Automate, so it is stripped before the join.
    #>
    param([string]$TenantName)

    if (-not $TenantName) { return '' }
    return ($TenantName -replace '-CC[0-9A-Fa-f]{6,}$', '').Trim()
}
