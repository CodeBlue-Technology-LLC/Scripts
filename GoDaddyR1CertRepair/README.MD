# Fix-GoDaddyR1Chain

Repairs GoDaddy / Starfield **R1** certificate chains on Windows servers.

GoDaddy moved DV TLS issuance to new "R1" root hierarchies — `GoDaddy TLS Root CA - R1`
and the Starfield equivalent — created in August 2025. Those roots are not yet present in
older client trust stores. A Windows server that builds its chain to the *self-signed* R1
root will fail for legacy clients even though the certificate itself is perfectly valid.

This script installs the published **R1 → G2 cross certificate**, so the chain anchors at a
G2 root that has been trusted since 2011.

---

## Symptoms

Appears immediately after an otherwise-normal certificate renewal:

- RD Gateway / RDWeb clients cannot connect
- Browsers on older devices report an untrusted root or incomplete chain
- Affects Windows 7 / 8.1, unpatched Server 2012 R2, thin clients (Wyse, IGEL),
  older mobile RD and browser clients, and any machine with Automatic Root Update disabled
- Modern, patched clients work fine — which is why this often gets reported as
  "only some users can't connect"

Diagnostic tools show a three-certificate chain ending in:

```
CN=GoDaddy TLS Root CA - R1   ->   CN=GoDaddy TLS Root CA - R1     (self-signed)
```

After the fix, four certificates ending in:

```
CN=Go Daddy Root Certificate Authority - G2   (self-signed, trusted since 2011)
```

**Not a certificate-binding problem, and not a TLS/cipher problem.** If the error mentions
socket or network failures rather than trust, look elsewhere.

---

## Usage

Run elevated on the affected server.

```powershell
# report only, make no changes
.\Fix-GoDaddyR1Chain.ps1 -CheckOnly

# apply
.\Fix-GoDaddyR1Chain.ps1

# fleet sweep
Invoke-Command -ComputerName SRV1,SRV2,SRV3 -FilePath .\Fix-GoDaddyR1Chain.ps1 -ArgumentList $true
```

The script self-skips any host with no affected certificate, so it is safe to run broadly.

**A reboot is required after a change.** Schannel caches built chains and a service restart
is not reliably sufficient.

### Parameters

| Parameter | Description |
|---|---|
| `-CheckOnly` | Report current state, change nothing. Always run this first. |
| `-CrossCertPath <path>` | Use a local copy of the cross certificate instead of downloading. Required on servers with no outbound internet. Still validated before import. |
| `-RemoveLegacyG1` | Also remove the G2→G1 cross cert and expired ValiCert-era intermediates so the chain terminates at G2 rather than continuing to the old G1 root. Only relevant if someone previously imported GoDaddy's `all_intermediate_ca_certificates.p7b`. Off by default. |

---

## What it changes

1. Scans `LocalMachine\My` and `LocalMachine\WebHosting` for certificates with a private key
   whose chain dead-ends at a self-signed `*TLS Root CA - R1*`.
2. Downloads and validates the matching R1 → G2 cross certificate, then imports it into
   **Intermediate Certification Authorities**.
3. Removes the self-signed R1 root from **Trusted Root**, so Schannel stops preferring the
   short path and builds through the cross certificate instead.
4. Re-checks every affected chain and reports the result.
5. Reports RD and IIS (`netsh http`) certificate bindings for reference.

Nothing is written unless a chain is actually broken.

### Validation before import

A downloaded certificate is only imported if **all** of the following hold:

- Thumbprint matches the pinned value for that hierarchy
- Subject matches the root being replaced
- It is **not** self-signed (i.e. it is genuinely a cross certificate)
- Issuer is the expected G2 root

Any mismatch aborts without touching the trust store.

### Pinned thumbprints

| Hierarchy | Cross certificate (SHA1) | Verified |
|---|---|---|
| GoDaddy | `9708542D763B1AAFC12CD5B65B3BC9E182F71FC4` | 2026-08-18 |
| Starfield | `0D257EB08FCBD99AD93E9124313D05FA439BF499` | 2026-08-19 |

Sourced from:

- <https://certs.godaddy.com/repository/gd_tls_root-r1-cross-g2.crt>
- <https://certs.starfieldtech.com/repository/sf_tls_root-r1-cross-g2.crt>

---

## Offline / air-gapped servers

DMZ hosts frequently cannot reach the GoDaddy repository. Stage the file on a share:

```powershell
.\Fix-GoDaddyR1Chain.ps1 -CrossCertPath \\share\certs\gd_tls_root-r1-cross-g2.crt
```

Validation still applies, so a wrong or tampered file will be rejected.

---

## Known limitations

- **Windows only.** Linux appliances (VMware UAG, NetScaler, F5, nginx) do not use the
  Windows certificate store. See *Non-Windows platforms* below.
- **This recurs on every renewal.** If a renewal re-imports the self-signed R1 root into
  Trusted Root, the chain breaks again. Add a post-renewal check to the runbook.
- **The G2 hierarchy is scheduled to end around April 2027.** The cross certificate is a
  bridge, not a permanent fix. Before then, legacy clients need the R1 root distributed
  (GPO / Intune) or need replacing.
- **R1 leaf certificates have no ClientAuth EKU.** Fine for server TLS. Anything relying on
  a GoDaddy DV cert for mTLS will break, and this script will not help with that.

---

## Non-Windows platforms

The same principle applies — the cross certificate must be present in the served chain —
but it goes in the chain file, not a certificate store. Order matters, leaf first:

```
1. your leaf certificate
2. GoDaddy TLS Intermediate CA DV - R1v1     (gd_tls_issuing_dv-r1v1)
3. GoDaddy TLS Root CA - R1, issued by G2    (gd_tls_root-r1-cross-g2)
```

Do **not** append the self-signed R1 root — that reintroduces the original problem.

Verify the assembled chain:

```powershell
(Select-String -Path chain.pem -Pattern "BEGIN CERTIFICATE").Count   # expect 3
certutil -dump gd_tls_root-r1-cross-g2.crt | Select-String "Subject:|Issuer:" -Context 0,1
```

Subject must be `GoDaddy TLS Root CA - R1`; Issuer must be
`Go Daddy Root Certificate Authority - G2`.

If PEM files were extracted from a PFX with `openssl pkcs12`, strip the `Bag Attributes`
metadata before uploading — some appliances reject the file without a useful error.

---

## References

- [GoDaddy — R1 root hierarchy transition and ClientAuth EKU removal](https://www.godaddy.com/help/why-is-godaddy-removing-clientauth-eku-and-transitioning-to-the-r1-root-hierarchy-for-dv-tls-issuance-42783)
- [GoDaddy — Fix an untrusted root or certificate chain error](https://www.godaddy.com/help/fix-an-untrusted-root-or-certificate-chain-error-42951)
- [GoDaddy certificate repository](https://certs.godaddy.com/repository)
- [Starfield certificate repository](https://certs.starfieldtech.com/repository)
