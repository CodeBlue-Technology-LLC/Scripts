# VPN Pre-Flight Check

Tells you in about 20 seconds whether a network will let your VPN connect —
before you install a client, before you haul gear in, before you burn an hour
on the phone with a hotel front desk.

Tests SSL VPN (TCP), IKEv2, and L2TP/IPsec (IKEv1) by sending real protocol
packets, not just port knocks. No install, no admin rights, no execution-policy
change. Windows PowerShell 5.1 or PowerShell 7.

## Files

| File | What it is |
|---|---|
| `VPNCheck.ps1` | The tester |
| `Run-VPNCheck.bat` | Double-click launcher |
| `vpncheck.config.json` | Your gateway list — edit this once |

Keep all three together. USB stick, network share, RMM package, whatever.

## Setup

Edit `vpncheck.config.json`:

```json
{
  "gateways": [
    { "host": "vpn.example.com", "sslPort": 1443, "expectedCertSubject": "example.com" }
  ]
}
```

`expectedCertSubject` is optional — it catches TLS interception even from a CA
the script doesn't recognize.

## Running it

```powershell
.\VPNCheck.ps1                                     # uses the config file
.\VPNCheck.ps1 -Gateway vpn.client.com             # one-off
.\VPNCheck.ps1 -Gateway vpn.client.com -SslPort 1443
.\VPNCheck.ps1 -Gateway vpn.client.com -Gentle     # see warning below
.\VPNCheck.ps1 -Quick                              # skip control probes
.\VPNCheck.ps1 -Gateway vpn.x.com -NoReport        # no log file
```

Every run drops a timestamped `.txt` report next to the script — useful for a
tech to email back from site.

## ⚠ Use `-Gentle` against gear with IPS scan detection

The tool hits 1443, 500 and 4500 back to back in ~15 seconds. To an IPS that
looks like a port scan, and the response is often to **auto-block your source
IP** — at which point the tool has caused the exact failure it's measuring, and
you're debugging a block you created.

`-Gentle` spaces the probes out (default 8 s, tune with `-GentleDelaySec`).
Adds about 40 s per gateway. Use it any time you're testing against a firewall
you don't want to annoy — which is most of them.

## What it tests

**Baseline** — captive portal (NCSI probe), DNS hijacking (does the network
answer NXDOMAIN with an IP?), active adapter.

**UDP egress control** — probes UDP/53 and a STUN server on UDP/19302. This is
the part that settles arguments. If those fail, the network blocks UDP broadly
and no gateway config will fix IKEv2/WireGuard/OpenVPN-UDP. If those pass but
your gateway is silent, the problem is specific to that destination.

It also parses the **public IP** out of the STUN response and flags CGNAT
(100.64.0.0/10). Behind carrier NAT the user has no idea what their public
address is, and that's exactly what you need to look for in a far-end firewall's
block list. A CGNAT address is shared — a block may have been earned by someone
else entirely.

**SSL VPN** — TCP connect, full TLS handshake, pulls the certificate, compares
the issuer against ~20 known inspection/proxy CAs (Fortinet, Zscaler, Netskope,
Palo Alto, Meraki, Barracuda…). This matters: a venue running SSL inspection
passes a plain port check and still breaks a cert-pinning client like
GlobalProtect or FortiClient. The script prints the actual issuer and thumbprint.

**IKEv2** — sends a real `IKE_SA_INIT` with three proposals (AES-256/SHA2-256/
MODP-2048, AES-128/SHA1/MODP-1024, 3DES/SHA1/MODP-1024). Three rather than one
so a healthy gateway matches something and answers with an actual SA.

**L2TP/IPsec** — sends a real IKEv1 Main Mode SA proposal. **L2TP/IPsec is
IKEv1, not IKEv2.** An IKEv2-only probe misreports a perfectly healthy L2TP
gateway, so both are tested.

Notify types are decoded by name. `SA ACCEPTED` means the gateway agreed to
crypto terms — the strongest result short of authenticating.
`NO_PROPOSAL_CHOSEN` means it's listening but rejected you.
`proposal MATCHED, wants DH MODP-1024` means policy overlap exists and it just
wants a different group.

L2TP itself (UDP/1701) is deliberately **not** probed — a correct gateway only
accepts it inside the established IPsec SA, so a probe there fails on healthy
boxes and proves nothing.

## Reading the output

```
  vpn.client.com
    SSL VPN (TCP)   : WILL CONNECT
    IKEv2 (IPsec)   : WILL CONNECT (gateway accepted a proposal)
    L2TP/IPsec (v1) : BLOCKED / NO ANSWER
```

### Refused vs Filtered — the distinction that matters

| What you see | What happened | Who's at fault |
|---|---|---|
| `RST in 3 ms` → **WARN** | Host answered "nothing listening here" | Gateway / wrong port. **Network is fine.** |
| `Timed out after 5000 ms` → **FAIL** | Packet discarded in silence | A firewall — see below |

A fast RST means the packet reached the far end and came back. That's proof the
venue is *not* filtering that port. Standing in a hotel arguing with the front
desk, this is the line that settles it.

### 🔴 A timeout does NOT mean the local network blocked you

This is the most important thing in this README.

A silent drop means **nobody answered**. It cannot tell you *who* discarded the
packet. The far-end firewall dropping you looks byte-for-byte identical to the
venue blocking you, from the client's side.

Earlier versions of this tool printed `BLOCKED BY THIS NETWORK` on a timeout.
That was wrong, and in a real investigation it sent a competent team chasing a
theory about apartment-complex DPI filtering for an afternoon. The actual cause
was a **failed-login auto-block on the destination firewall** — techs retrying
bad credentials had gotten the user's public IP blocked for 24 hours.

So when every port to one gateway goes silent while the UDP control probes pass,
the tool now prints a diagnosis block and tells you to suspect the far end
**first**. On WatchGuard specifically, the auto-block list lives in two places
and the obvious one is the wrong one:

```
Firewall > Blocked Sites       = permanent, manually configured
System Status > Blocked Sites  = TEMPORARY auto-blocks   <-- this one
```

Reasons you'll see there: `block failed logins`, `Port scan attack`.

### Confirming it from the far end

If you control the destination firewall, capture on its external interface while
the client retries. On WatchGuard: **System → Diagnostics → Network → TCP Dump**,
tick Advanced Options (you must then specify `-i` yourself; the dropdown is
ignored), and tick *Stream data to a file* for a .pcap.

```
-i eth2 -n -c 40 host <client.public.ip>
```

- **SYNs arrive, no SYN-ACK leaves** → the firewall is dropping. Auto-block,
  geolocation, botnet/reputation, or a policy From restriction.
- **Nothing arrives at all** → it never got there. Upstream of the firewall.

Two traps that cost real time:

1. `tcp[13] = 2` matches SYN **only** — it excludes SYN-ACK, so you can't see
   whether the firewall replied. Good for finding connection attempts on a busy
   port, useless for judging the response.
2. A `-c` limit on a busy port gets exhausted by *other* users' traffic in
   seconds, ending the capture before your client ever retries.

## Honest limitations

- **A timeout can't identify who dropped the packet.** See above. The control
  probes narrow it; only a far-end capture settles it.
- **UDP silence is ambiguous.** Blocked, gateway down, or not listening on 500.
  That's why the STUN/DNS controls exist — don't read a UDP FAIL in isolation.
- **ESP (IP protocol 50) is not tested.** Raw protocol-50 sockets need admin and
  a driver on Windows. Nearly everything behind NAT uses UDP/4500 encapsulation,
  which *is* tested — but a gateway doing native ESP without NAT-T isn't covered.
- **Self-signed ≠ safe.** A self-signed cert is also what interception looks
  like when the CA isn't in the hint list. Record the thumbprint from a known-good
  network and compare — that's the only check that works against an unknown CA.
- **It tests reachability and interception, not authentication.** A full PASS
  means packets flow and nothing is decrypting your TLS. It says nothing about
  whether the credentials, certificate, or client profile are right.
- **Running as SYSTEM** (via RMM) hides anything user-scoped: per-user proxy
  settings, user-context firewall rules, the client's own profile.
- **Rate limiting.** Some gateways silently drop repeated IKE_SA_INIT from one
  source. If results get worse in a tight loop, that's why. Use `-Gentle`.

## Building a real .exe

The `.bat` wrapper needs nothing installed and is the practical answer. For a
single-file `.exe`:

```powershell
Install-Module ps2exe -Scope CurrentUser
Invoke-PS2EXE .\VPNCheck.ps1 .\VPNCheck.exe -title "VPN Pre-Flight Check"
```

Unsigned, so expect SmartScreen the first time.

## Changelog

- **v1.5** — Stopped claiming "BLOCKED BY THIS NETWORK" on a timeout. Added
  `-Gentle` so the tool doesn't trip IPS scan detection and cause the failure
  it's measuring.
- **v1.4** — Public IP via STUN + CGNAT flag. ICMP datapoint. Destination-specific
  blackhole diagnosis.
- **v1.3** — IKEv2 offers three proposals. Verdict separates "agreed to crypto
  terms" from "listening but rejected us". Timeout 3000 → 5000 ms.
- **v1.2** — IKEv1 Main Mode probing for L2TP/IPsec. Notify types decoded by name.
- **v1.1** — TCP failures classified Refused (RST) vs Filtered (silence).
- **v1.0** — Initial.
