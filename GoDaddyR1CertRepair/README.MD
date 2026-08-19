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
