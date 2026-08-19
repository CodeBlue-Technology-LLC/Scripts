<#
.SYNOPSIS
    Repairs GoDaddy / Starfield "R1" certificate chains on Windows servers.

.DESCRIPTION
    GoDaddy moved DV TLS issuance to new "R1" root hierarchies (GoDaddy TLS Root
    CA - R1 and the Starfield equivalent). Those roots were created in 2025 and are
    not present in older client trust stores. A Windows server that builds a chain
    terminating at the self-signed R1 root will fail for legacy clients: Win7/8.1,
    unpatched 2012 R2, thin clients (Wyse/IGEL), older mobile RD/browser clients,
    and anything with Automatic Root Update disabled.

    Fix: install the published R1->G2 cross certificate into Intermediate
    Certification Authorities, and remove the self-signed R1 root from Trusted Root
    so Schannel stops preferring the short path. The chain then anchors at the
    G2 root, trusted since 2011.

    v2 changes:
      - Scans LocalMachine\My AND LocalMachine\WebHosting (IIS/SNI/CCS)
      - Detects the affected root structurally instead of by a single hardcoded
        thumbprint, so it handles both the GoDaddy and Starfield hierarchies
      - Validates any downloaded cross certificate before importing it
      - Reports IIS (netsh) and RD bindings

    Idempotent. Run with -CheckOnly first.

.PARAMETER CheckOnly
    Report only. Make no changes.

.PARAMETER CrossCertPath
    Use a local copy of the cross certificate instead of downloading. Required on
    servers with no outbound internet. Still validated before import.

.PARAMETER RemoveLegacyG1
    Also remove the G2->G1 cross cert and expired ValiCert-era intermediates from
    the Intermediate store so the chain terminates at G2 rather than continuing to
    the old G1 root. Only relevant if someone previously imported GoDaddy's
    all_intermediate_ca_certificates.p7b. Off by default.

.EXAMPLE
    .\Fix-GoDaddyR1Chain.ps1 -CheckOnly

.EXAMPLE
    .\Fix-GoDaddyR1Chain.ps1

.EXAMPLE
    Invoke-Command -ComputerName SRV1,SRV2 -FilePath .\Fix-GoDaddyR1Chain.ps1

.NOTES
    Reboot required after a change. Schannel caches built chains and a service
    restart is not reliably sufficient.
#>

[CmdletBinding()]
param(
    [switch]$CheckOnly,
    [string]$CrossCertPath,
    [switch]$RemoveLegacyG1
)

$ErrorActionPreference = 'Stop'

# --- Hierarchy definitions ---------------------------------------------------
# Match is on the self-signed root's subject. CrossUrl is where the R1->G2 cross
# certificate lives. KnownCrossThumbprint is pinned where verified; where $null,
# the download is validated structurally instead (subject/issuer must line up and
# the issuing root must already be trusted locally).

$Hierarchies = @(
    @{
        Name                 = 'GoDaddy'
        RootSubjectMatch     = 'CN=GoDaddy TLS Root CA - R1'
        ExpectedCrossIssuer  = 'Go Daddy Root Certificate Authority - G2'
        CrossUrl             = 'https://certs.godaddy.com/repository/gd_tls_root-r1-cross-g2.crt'
        KnownCrossThumbprint = '9708542D763B1AAFC12CD5B65B3BC9E182F71FC4'   # verified 2026-08-18
    },
    @{
        Name                 = 'Starfield'
        RootSubjectMatch     = 'CN=Starfield TLS Root CA - R1'
        ExpectedCrossIssuer  = 'Starfield Root Certificate Authority - G2'
        CrossUrl             = 'https://certs.starfieldtech.com/repository/sf_tls_root-r1-cross-g2.crt'
        KnownCrossThumbprint = $null   # URL verified 2026-08-19; thumbprint not yet pinned.
                                       # Validated structurally at runtime. To pin it, run
                                       # certutil -dump on the downloaded file and paste the SHA1 here.
    }
)

# Legacy G1-era certs, removed only with -RemoveLegacyG1
$LEGACY_G1 = @(
    '841D4A9FC9D3B2F0CA5FAB95525AB2066ACF8322',  # G2 root cross-signed by G1 Class 2
    'DE70F4E2116F7FDCE75F9D13012B7E687A3B2C62',  # Go Daddy Class 2 CA via ValiCert  (expired 2024-06-29)
    '363E4734F757BDEB89868EFE94907774A327695E',  # Starfield Class 2 CA via ValiCert (expired 2024-06-29)
    '446A2A00C1BBA36D59D1C178A67A27C50E6D03DF'   # Starfield Secure CA via ValiCert  (expired 2024-01-09)
)

$CertStores = @('My','WebHosting')

function Write-Step { param($m) Write-Host "`n== $m" -ForegroundColor Cyan }
function Write-Ok   { param($m) Write-Host "   [ok]   $m" -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "   [warn] $m" -ForegroundColor Yellow }
function Write-Info { param($m) Write-Host "   [info] $m" }

function Get-Chain {
    param([Security.Cryptography.X509Certificates.X509Certificate2]$Cert)
    $chain = New-Object Security.Cryptography.X509Certificates.X509Chain
    $null = $chain.Build($Cert)
    $anchor = $chain.ChainElements[$chain.ChainElements.Count - 1].Certificate
    [pscustomobject]@{
        Depth        = $chain.ChainElements.Count
        Anchor       = $anchor
        SelfAnchored = ($anchor.Subject -eq $anchor.Issuer)
        Elements     = @($chain.ChainElements | ForEach-Object { $_.Certificate })
        StatusText   = if ($chain.ChainStatus.Count) { ($chain.ChainStatus | ForEach-Object { $_.Status }) -join ', ' } else { 'OK' }
    }
}

function Get-ServerCerts {
    foreach ($s in $CertStores) {
        $path = "Cert:\LocalMachine\$s"
        if (-not (Test-Path $path)) { continue }
        Get-ChildItem $path -EA SilentlyContinue | Where-Object {
            $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date)
        } | ForEach-Object {
            $_ | Add-Member -NotePropertyName SourceStore -NotePropertyValue $s -Force -PassThru
        }
    }
}

# --- 1. Find certs whose chain dead-ends at a self-signed R1 root -------------
Write-Step "Scanning certificate stores: $($CertStores -join ', ')"

$affected = @()
foreach ($c in Get-ServerCerts) {
    $chain = Get-Chain $c
    $match = $Hierarchies | Where-Object {
        $chain.SelfAnchored -and $chain.Anchor.Subject -like "$($_.RootSubjectMatch)*"
    } | Select-Object -First 1

    if ($match) {
        $affected += [pscustomobject]@{
            Cert      = $c
            Store     = $c.SourceStore
            Root      = $chain.Anchor
            Hierarchy = $match
        }
        Write-Warn ("{0}  [{1}]  exp {2:yyyy-MM-dd}" -f $c.Subject, $c.SourceStore, $c.NotAfter)
        Write-Warn ("  chain dead-ends at self-signed {0}" -f $chain.Anchor.Subject)
    }
    else {
        Write-Info ("{0}  [{1}]  anchor: {2}" -f $c.Subject, $c.SourceStore, $chain.Anchor.Subject)
    }
}

if (-not $affected) {
    Write-Ok "No certificates chaining to a self-signed R1 root. Nothing to do on $env:COMPUTERNAME."
    return
}

$targets = $affected | Group-Object { $_.Hierarchy.Name }

# --- 2. Install cross certificate(s) ----------------------------------------
foreach ($g in $targets) {
    $h    = ($g.Group | Select-Object -First 1).Hierarchy
    $root = ($g.Group | Select-Object -First 1).Root

    Write-Step "$($h.Name): checking for R1->G2 cross certificate"

    # Already present? Subject matches the R1 root, but cross-signed (issuer differs).
    $have = Get-ChildItem Cert:\LocalMachine\CA -EA SilentlyContinue | Where-Object {
        $_.Subject -eq $root.Subject -and $_.Issuer -ne $_.Subject
    }

    if ($have) {
        Write-Ok "Cross certificate present (issued by $($have[0].Issuer))."
        continue
    }
    if ($CheckOnly) {
        Write-Warn "Cross certificate MISSING - would download and import from $($h.CrossUrl)"
        continue
    }

    if ($CrossCertPath) {
        if (-not (Test-Path $CrossCertPath)) { throw "CrossCertPath not found: $CrossCertPath" }
        $file = $CrossCertPath
        Write-Info "Using local copy: $file"
    }
    else {
        $file = Join-Path $env:TEMP ("cross_{0}.crt" -f $h.Name)
        Write-Info "Downloading $($h.CrossUrl)"
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest $h.CrossUrl -OutFile $file -UseBasicParsing
        }
        catch {
            Write-Warn "Download failed: $($_.Exception.Message)"
            Write-Warn "Obtain the $($h.Name) R1->G2 cross certificate manually and re-run with -CrossCertPath."
            continue
        }
    }

    $cand = New-Object Security.Cryptography.X509Certificates.X509Certificate2 $file

    # Validate before trusting anything
    if ($h.KnownCrossThumbprint -and $cand.Thumbprint -ne $h.KnownCrossThumbprint) {
        throw "Thumbprint mismatch for $($h.Name). Expected $($h.KnownCrossThumbprint), got $($cand.Thumbprint). Not importing."
    }
    if ($cand.Subject -ne $root.Subject) {
        throw "Subject mismatch. Expected '$($root.Subject)', got '$($cand.Subject)'. Not importing."
    }
    if ($cand.Subject -eq $cand.Issuer) {
        throw "Downloaded certificate is self-signed - this is the native root, not the cross certificate. Not importing."
    }
    if ($cand.Issuer -notlike "*$($h.ExpectedCrossIssuer)*") {
        throw "Unexpected issuer '$($cand.Issuer)'. Expected '$($h.ExpectedCrossIssuer)'. Not importing."
    }

    $anchorTrusted = Get-ChildItem Cert:\LocalMachine\Root -EA SilentlyContinue |
                     Where-Object { $_.Subject -like "*$($h.ExpectedCrossIssuer)*" }
    if (-not $anchorTrusted) {
        Write-Warn "$($h.ExpectedCrossIssuer) is NOT in Trusted Root. Import it before relying on this fix."
    }

    Write-Ok "Validated: '$($cand.Subject)' issued by '$($cand.Issuer)' ($($cand.Thumbprint))"

    $store = New-Object Security.Cryptography.X509Certificates.X509Store 'CA','LocalMachine'
    $store.Open('ReadWrite'); $store.Add($cand); $store.Close()
    Write-Ok "Imported into Intermediate Certification Authorities."
}

# --- 3. Remove the self-signed R1 root(s) from Trusted Root -------------------
Write-Step "Checking Trusted Root for self-signed R1 root(s)"
$rootStore = New-Object Security.Cryptography.X509Certificates.X509Store 'Root','LocalMachine'
$rootStore.Open('ReadWrite')

foreach ($tp in ($affected.Root.Thumbprint | Select-Object -Unique)) {
    $bad = $rootStore.Certificates | Where-Object { $_.Thumbprint -eq $tp }
    if (-not $bad)       { Write-Ok   "Not present: $tp" }
    elseif ($CheckOnly)  { Write-Warn "PRESENT - would remove: $tp ($($bad.Subject))" }
    else                 { $rootStore.Remove($bad); Write-Ok "Removed: $tp ($($bad.Subject))" }
}
$rootStore.Close()

# --- 4. Optional legacy G1 cleanup ------------------------------------------
if ($RemoveLegacyG1) {
    Write-Step "Removing legacy G1 cross / expired intermediates"
    $caStore = New-Object Security.Cryptography.X509Certificates.X509Store 'CA','LocalMachine'
    $caStore.Open('ReadWrite')
    foreach ($tp in $LEGACY_G1) {
        $f = $caStore.Certificates | Where-Object { $_.Thumbprint -eq $tp }
        if (-not $f)        { Write-Info "not present: $tp" }
        elseif ($CheckOnly) { Write-Warn "would remove: $tp ($($f.Subject))" }
        else                { $caStore.Remove($f); Write-Ok "removed: $tp" }
    }
    $caStore.Close()
}

# --- 5. Post-check -----------------------------------------------------------
Write-Step "Resulting chains"
$stillBroken = $false
foreach ($a in $affected) {
    $chain = Get-Chain $a.Cert
    Write-Info ("{0}  [{1}]" -f $a.Cert.Subject, $a.Store)
    $i = 0
    foreach ($e in $chain.Elements) { Write-Info ("  [{0}] {1}" -f $i, $e.Subject); $i++ }
    Write-Info ("  status: {0}" -f $chain.StatusText)
    if ($chain.SelfAnchored -and $chain.Anchor.Subject -like "*TLS Root CA - R1*") {
        Write-Warn "  STILL anchored at self-signed R1."
        $stillBroken = $true
    }
    else {
        Write-Ok "  anchored at: $($chain.Anchor.Subject)"
    }
}

# --- 6. Bindings -------------------------------------------------------------
Write-Step "Certificate bindings on this host"
try {
    Get-WmiObject -Namespace root\cimv2\TerminalServices -Class Win32_TSGeneralSetting -EA Stop |
        ForEach-Object { Write-Info "RD (TSGeneralSetting): $($_.SSLCertificateSHA1Hash)" }
} catch { Write-Info "No RD Session Host / Gateway settings on this host." }

try {
    netsh http show sslcert 2>$null |
        Select-String 'IP:port|Hostname:port|Certificate Hash|Certificate Store Name' |
        ForEach-Object { Write-Info ($_.ToString().Trim()) }
} catch { }

# --- 7. Summary --------------------------------------------------------------
Write-Step "Summary"
if     ($CheckOnly)    { Write-Info "CheckOnly - no changes made. Re-run without -CheckOnly to apply." }
elseif ($stillBroken)  { Write-Warn "Chain NOT fully repaired on $env:COMPUTERNAME. Review output above." }
else {
    Write-Ok   "Chain repaired on $env:COMPUTERNAME."
    Write-Warn "REBOOT REQUIRED. Schannel caches built chains; a service restart is not reliably enough."
}
