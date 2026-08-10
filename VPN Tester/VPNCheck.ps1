<#
    VPNCheck.ps1 -- Pre-flight VPN reachability tester   (v1.5)
    Tests whether a network (hotel / apartment / coffee shop / client site)
    will let an SSL VPN (TCP 443) or IKEv2 IPsec VPN (UDP 500 / 4500) connect,
    WITHOUT installing or configuring the VPN client.

    Usage:
        .\VPNCheck.ps1                          # reads vpncheck.config.json next to the script
        .\VPNCheck.ps1 -Gateway vpn.corp.com    # one-off target
        .\VPNCheck.ps1 -Gateway vpn.corp.com -SslPort 10443
        .\VPNCheck.ps1 -Quick                   # skip the slower control tests
        .\VPNCheck.ps1 -Gateway vpn.corp.com -Gentle   # space probes out; use
                                                       # when the target has IPS
                                                       # scan detection enabled

    FIELD NOTE - a silent timeout does NOT mean the local network blocked you.
    It means nobody answered. The far end discarding your packets looks exactly
    the same from the client. This tool was once confidently wrong about that
    and cost an afternoon: the real cause was a failed-login auto-block on the
    destination firewall, sitting in System Status > Blocked Sites. When every
    port to one gateway goes silent, suspect the far end first.

    v1.1 - TCP failures are now classified as Refused (RST, host answered) vs
           Filtered (silence, firewall dropped it). These mean opposite things
           and the verdict now reports them differently.
    v1.2 - Added IKEv1 Main Mode probing. L2TP/IPsec is IKEv1, NOT IKEv2, so
           an IKEv2-only probe misreports a healthy L2TP gateway. Also decodes
           IKE notify types by name (NO_PROPOSAL_CHOSEN, INVALID_MAJOR_VERSION,
           etc.) instead of just printing a byte count.
    v1.3 - IKEv2 now offers 3 proposals (was 1), so a healthy gateway answers
           SA ACCEPTED instead of NO_PROPOSAL_CHOSEN. Verdict separates "agreed
           to crypto terms" from "listening but rejected us". Default timeout
           raised 3000 -> 5000 ms after observing a 2077 ms RST in the field.
    v1.4 - Reports the public IP the far end sees, parsed out of the STUN
           response, and flags CGNAT. Adds ICMP as a (non-conclusive) datapoint.
           New "destination-specific blackhole" diagnosis for the case where the
           UDP control passes but every port on one gateway is silent - that is
           a far-end or path problem, not a venue blocking VPNs.
    v1.5 - Stopped claiming "BLOCKED BY THIS NETWORK" on a timeout; the tool
           cannot know that and the wording sent a real investigation the wrong
           way. Added -Gentle to space probes so the tool does not trip IPS
           port-scan auto-block and cause the failure it is measuring.

    No admin rights required. No install. PowerShell 5.1 or 7+.
#>

[CmdletBinding()]
param(
    [string]   $Gateway,
    [string[]] $Gateways,
    [int]      $SslPort      = 443,
    [string]   $ExpectedCertSubject,
    [switch]   $Quick,

    # Probing 1443, 500 and 4500 back to back looks like a port scan to an
    # IPS. WatchGuard Default Packet Handling (and equivalents on other gear)
    # will auto-block the source, which means the tool causes the very failure
    # it is trying to measure - and you then debug a block you created.
    # -Gentle spaces the probes out to stay under scan-detection thresholds.
    [switch]   $Gentle,
    [int]      $GentleDelaySec = 8,
    [switch]   $NoReport,
    [string]   $ReportPath,
    # 5s, not 3s. A real gateway was observed returning a TCP RST in 2077 ms on
    # a clean wired LAN. At a 3000 ms timeout that is a 900 ms margin - on hotel
    # wifi the same RST lands past the deadline and gets misreported as a
    # firewall block. Do not lower this without a good reason.
    [int]      $TimeoutMs    = 5000
)

$ErrorActionPreference = 'Stop'
# Deliberately 1.0, not 2.0: this is a field tool. Catching typo'd variables is
# worth it; hard-failing on a missing property mid-probe is not.
Set-StrictMode -Version 1.0

$script:Results    = New-Object System.Collections.ArrayList
$script:PublicIp   = $null
$script:GentleMode = [bool]$Gentle
$script:GentleDelay = $GentleDelaySec

$script:ScriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($script:ScriptDir)) {
    $script:ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if ([string]::IsNullOrWhiteSpace($script:ScriptDir)) {
    $script:ScriptDir = (Get-Location).Path
}

# ---------------------------------------------------------------------------
# Output helpers
# ---------------------------------------------------------------------------

function Write-Head {
    param([string]$Text)
    Write-Host ""
    Write-Host ("=" * 74) -ForegroundColor DarkCyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host ("=" * 74) -ForegroundColor DarkCyan
}

function Write-Sub {
    param([string]$Text)
    Write-Host ""
    Write-Host "-- $Text" -ForegroundColor White
}

function Add-Result {
    param(
        [string] $Category,
        [string] $Test,
        [ValidateSet('PASS','FAIL','WARN','INFO','SKIP')]
        [string] $Status,
        [string] $Detail
    )

    $null = $script:Results.Add([pscustomobject]@{
        Category = $Category
        Test     = $Test
        Status   = $Status
        Detail   = $Detail
    })

    $color = switch ($Status) {
        'PASS' { 'Green'  }
        'FAIL' { 'Red'    }
        'WARN' { 'Yellow' }
        'SKIP' { 'DarkGray' }
        default { 'Gray'  }
    }

    $tag = '[{0}]' -f $Status.PadRight(4)
    Write-Host "  $tag " -ForegroundColor $color -NoNewline
    Write-Host ("{0,-34} " -f $Test) -NoNewline
    Write-Host $Detail -ForegroundColor DarkGray
}

# Pause between probes when -Gentle is set, so a rapid sequence of connections
# to different ports does not register as a port scan on the target's IPS.
function Invoke-ProbePause {
    if (-not $script:GentleMode) { return }
    Write-Host ("       ...pausing {0}s (-Gentle)" -f $script:GentleDelay) -ForegroundColor DarkGray
    Start-Sleep -Seconds $script:GentleDelay
}

# ---------------------------------------------------------------------------
# Packet builders
# ---------------------------------------------------------------------------

function New-RandomBytes {
    param([int]$Count)
    $b = New-Object byte[] $Count
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try   { $rng.GetBytes($b) }
    finally { $rng.Dispose() }
    return $b
}

function ConvertTo-BE16 {
    param([int]$Value)
    return [byte[]]@( [byte](($Value -shr 8) -band 0xFF), [byte]($Value -band 0xFF) )
}

function ConvertTo-BE32 {
    param([uint32]$Value)
    return [byte[]]@(
        [byte](($Value -shr 24) -band 0xFF),
        [byte](($Value -shr 16) -band 0xFF),
        [byte](($Value -shr  8) -band 0xFF),
        [byte]( $Value          -band 0xFF)
    )
}

<#
    Builds a well-formed IKEv2 IKE_SA_INIT request (RFC 7296).

    IKE header (28) + SA payload (132) + KE payload (264) + Nonce (36) = 460

    Any reply at all proves the UDP path is open end to end - even a short
    NO_PROPOSAL_CHOSEN notify (~36 bytes). But three proposals are offered
    rather than one so that a gateway with a normal policy answers with an
    actual SA, which additionally proves IKEv2 is USABLE and not merely
    listening. Those are different findings and the verdict reports them so.
#>
function New-IkeV2SaInitPacket {

    # ---- Transform / proposal helpers -------------------------------------
    # Transform substructure: LastSubstruc(1) Reserved(1) Length(2)
    #                         Type(1) Reserved(1) ID(2) [attributes]
    function New-Tr {
        param([bool]$Last, [int]$Type, [int]$Id, [byte[]]$Attrs = @())
        $ln = 8 + $Attrs.Length
        $b  = [byte[]]@( $(if ($Last) { 0 } else { 3 }), 0 ) + (ConvertTo-BE16 $ln) `
              + [byte[]]@($Type, 0) + (ConvertTo-BE16 $Id)
        if ($Attrs.Length -gt 0) { $b = $b + $Attrs }
        return $b
    }
    # Key length attribute: AF=1 (TV) | type 14 => 0x800E
    function New-KeyLen { param([int]$Bits) return [byte[]]@(0x80,0x0E) + (ConvertTo-BE16 $Bits) }

    # Proposal substructure: LastSubstruc(1) Reserved(1) Length(2)
    #                        Num(1) ProtoID(1)=1(IKE) SPISize(1)=0 NumTransforms(1)
    function New-Prop {
        param([int]$Num, [bool]$Last, [byte[]]$Transforms, [int]$Count)
        $ln = 8 + $Transforms.Length
        return [byte[]]@( $(if ($Last) { 0 } else { 2 }), 0 ) + (ConvertTo-BE16 $ln) `
               + [byte[]]@($Num, 1, 0, $Count) + $Transforms
    }

    # Three proposals, strongest first. One proposal is a coin flip against an
    # unknown policy - three means a real gateway almost always matches
    # something and answers with SA ACCEPTED instead of NO_PROPOSAL_CHOSEN.
    # Types: 1=ENCR 2=PRF 3=INTEG 4=DH

    # P1: AES-256-CBC / PRF-SHA2-256 / HMAC-SHA2-256-128 / MODP-2048
    #     DH group 14 deliberately matches our KE payload below.
    $p1t = (New-Tr $false 1 12 (New-KeyLen 256)) + (New-Tr $false 2 5) `
           + (New-Tr $false 3 12) + (New-Tr $true 4 14)
    $p1  = New-Prop 1 $false $p1t 4

    # P2: AES-128-CBC / PRF-SHA1 / HMAC-SHA1-96 / MODP-1024
    $p2t = (New-Tr $false 1 12 (New-KeyLen 128)) + (New-Tr $false 2 2) `
           + (New-Tr $false 3 2) + (New-Tr $true 4 2)
    $p2  = New-Prop 2 $false $p2t 4

    # P3: 3DES / PRF-SHA1 / HMAC-SHA1-96 / MODP-1024  (legacy gear)
    $p3t = (New-Tr $false 1 3) + (New-Tr $false 2 2) `
           + (New-Tr $false 3 2) + (New-Tr $true 4 2)
    $p3  = New-Prop 3 $true $p3t 4

    $proposal = $p1 + $p2 + $p3                         # 44 + 44 + 40 = 128

    # ---- SA payload (type 33) --------------------------------------------
    $saLen = 4 + $proposal.Length                       # 132
    $sa    = [byte[]]@(34,0) + (ConvertTo-BE16 $saLen) + $proposal

    # ---- KE payload (type 34) --------------------------------------------
    $keData = New-RandomBytes 256
    $keLen  = 8 + $keData.Length                        # 264
    $ke     = [byte[]]@(40,0) + (ConvertTo-BE16 $keLen) + (ConvertTo-BE16 14) `
              + [byte[]]@(0,0) + $keData

    # ---- Nonce payload (type 40) -----------------------------------------
    $nonceData = New-RandomBytes 32
    $nonceLen  = 4 + $nonceData.Length                  # 36
    $nonce     = [byte[]]@(0,0) + (ConvertTo-BE16 $nonceLen) + $nonceData

    # ---- Header -----------------------------------------------------------
    $initSpi = New-RandomBytes 8
    if (($initSpi | Measure-Object -Sum).Sum -eq 0) { $initSpi[0] = 1 }   # must be non-zero
    $respSpi = New-Object byte[] 8

    $body   = $sa + $ke + $nonce
    $total  = 28 + $body.Length                         # 376

    $header = $initSpi + $respSpi `
              + [byte[]]@(33, 0x20, 34, 0x08) `
              + (ConvertTo-BE32 0) `
              + (ConvertTo-BE32 ([uint32]$total))

    return @{
        Packet       = $header + $body
        InitiatorSpi = $initSpi
    }
}

# STUN Binding Request (RFC 5389) -- 20 bytes, used as a UDP egress control.
function New-StunBindingRequest {
    $txid = New-RandomBytes 12
    return @{
        Packet = (ConvertTo-BE16 1) + (ConvertTo-BE16 0) `
                 + [byte[]]@(0x21,0x12,0xA4,0x42) + $txid
        TxId   = $txid
    }
}

<#
    Pulls the public (post-NAT) address out of a STUN Binding Response.

    Why this matters: behind CGNAT the user has no way to know their own public
    IP, and that is exactly the value you need to check against a firewall's
    Blocked Sites / geo-restriction / allow list at the far end. It is also the
    address the far end sees, which is not necessarily one per user.

    Handles XOR-MAPPED-ADDRESS (0x0020, current) and MAPPED-ADDRESS (0x0001,
    legacy), including the 4-byte attribute padding rule.
#>
function Get-StunMappedAddress {
    param([byte[]]$Bytes)

    if ($null -eq $Bytes -or $Bytes.Length -lt 20)       { return $null }
    if ($Bytes[0] -ne 0x01 -or $Bytes[1] -ne 0x01)       { return $null }   # Binding Success
    if ($Bytes[4] -ne 0x21 -or $Bytes[5] -ne 0x12 -or
        $Bytes[6] -ne 0xA4 -or $Bytes[7] -ne 0x42)       { return $null }   # magic cookie

    $msgLen = ($Bytes[2] -shl 8) -bor $Bytes[3]
    $p      = 20
    $end    = [Math]::Min($Bytes.Length, 20 + $msgLen)

    while (($p + 4) -le $end) {
        $atype = ($Bytes[$p]     -shl 8) -bor $Bytes[$p + 1]
        $alen  = ($Bytes[$p + 2] -shl 8) -bor $Bytes[$p + 3]
        $v     = $p + 4
        if (($v + $alen) -gt $Bytes.Length) { break }

        if (($atype -eq 0x0020 -or $atype -eq 0x0001) -and $alen -ge 8 -and $Bytes[$v + 1] -eq 0x01) {
            $port = ($Bytes[$v + 2] -shl 8) -bor $Bytes[$v + 3]
            $o    = @($Bytes[$v + 4], $Bytes[$v + 5], $Bytes[$v + 6], $Bytes[$v + 7])

            if ($atype -eq 0x0020) {
                $port = $port -bxor 0x2112
                $o[0] = $o[0] -bxor 0x21
                $o[1] = $o[1] -bxor 0x12
                $o[2] = $o[2] -bxor 0xA4
                $o[3] = $o[3] -bxor 0x42
            }
            return [pscustomobject]@{
                Address = ($o -join '.')
                Port    = $port
            }
        }
        # attributes are padded to a 4-byte boundary
        $p = $v + $alen + ((4 - ($alen % 4)) % 4)
    }
    return $null
}

# Is an address in CGNAT space (RFC 6598, 100.64.0.0/10)? Tells you the client
# is behind carrier NAT and shares a public IP with other subscribers.
function Test-CgnatAddress {
    param([string]$Address)
    $parts = $Address -split '\.'
    if ($parts.Count -ne 4) { return $false }
    return ([int]$parts[0] -eq 100) -and ([int]$parts[1] -ge 64) -and ([int]$parts[1] -le 127)
}

<#
    Builds an IKEv1 (ISAKMP) Main Mode SA proposal -- RFC 2408 / 2409.

    THIS is what L2TP/IPsec speaks. L2TP/IPsec is IKEv1, not IKEv2, so an
    IKEv2-only probe can come back NO_PROPOSAL_CHOSEN or INVALID_MAJOR_VERSION
    against a perfectly healthy L2TP gateway.

    ISAKMP header (28) + SA payload (124) = 152 bytes.
    Three transforms are offered to maximise the chance of a real Main Mode
    reply rather than a rejection:
        1) 3DES-CBC  / SHA1     / PSK / MODP-1024
        2) AES-128   / SHA1     / PSK / MODP-1024
        3) AES-256   / SHA2-256 / PSK / MODP-2048
#>
function New-IkeV1MainModePacket {

    # Basic (TV-format) ISAKMP attribute: high bit set on the type field.
    function New-Attr {
        param([int]$Type, [int]$Value)
        return [byte[]]@(
            [byte](0x80 -bor (($Type -shr 8) -band 0x7F)),
            [byte]($Type -band 0xFF)
        ) + (ConvertTo-BE16 $Value)
    }

    # Attribute types: 1=Enc 2=Hash 3=Auth 4=Group 11=LifeType 12=LifeDuration 14=KeyLength
    $a1 = (New-Attr 1 5) + (New-Attr 2 2) + (New-Attr 3 1) + (New-Attr 4 2) `
          + (New-Attr 11 1) + (New-Attr 12 28800)                                  # 3DES/SHA1
    $a2 = (New-Attr 1 7) + (New-Attr 14 128) + (New-Attr 2 2) + (New-Attr 3 1) `
          + (New-Attr 4 2) + (New-Attr 11 1) + (New-Attr 12 28800)                 # AES128/SHA1
    $a3 = (New-Attr 1 7) + (New-Attr 14 256) + (New-Attr 2 4) + (New-Attr 3 1) `
          + (New-Attr 4 14) + (New-Attr 11 1) + (New-Attr 12 28800)                # AES256/SHA2-256

    # Transform payload: NextPayload(1) RESERVED(1) Length(2)
    #                    TransformNum(1) TransformID(1)=KEY_IKE RESERVED2(2)
    $t1 = [byte[]]@(3,0) + (ConvertTo-BE16 (8 + $a1.Length)) + [byte[]]@(1,1,0,0) + $a1  # 32
    $t2 = [byte[]]@(3,0) + (ConvertTo-BE16 (8 + $a2.Length)) + [byte[]]@(2,1,0,0) + $a2  # 36
    $t3 = [byte[]]@(0,0) + (ConvertTo-BE16 (8 + $a3.Length)) + [byte[]]@(3,1,0,0) + $a3  # 36

    $transforms = $t1 + $t2 + $t3                                                        # 104

    # Proposal payload: NextPayload(1) RESERVED(1) Length(2)
    #                   ProposalNum(1) ProtoID(1)=ISAKMP SPISize(1)=0 NumTransforms(1)=3
    $propLen  = 8 + $transforms.Length                                                   # 112
    $proposal = [byte[]]@(0,0) + (ConvertTo-BE16 $propLen) + [byte[]]@(1,1,0,3) + $transforms

    # SA payload: NextPayload(1) RESERVED(1) Length(2) DOI(4)=IPSEC Situation(4)=IDENTITY
    $saLen = 12 + $proposal.Length                                                       # 124
    $sa    = [byte[]]@(0,0) + (ConvertTo-BE16 $saLen) `
             + (ConvertTo-BE32 1) + (ConvertTo-BE32 1) + $proposal

    # ISAKMP header: iCookie(8) rCookie(8) NextPayload(1)=1(SA) Version(1)=0x10
    #                ExchType(1)=2(Identity Protection/Main Mode) Flags(1)=0
    #                MessageID(4)=0 Length(4)
    $iCookie = New-RandomBytes 8
    if (($iCookie | Measure-Object -Sum).Sum -eq 0) { $iCookie[0] = 1 }
    $rCookie = New-Object byte[] 8

    $total  = 28 + $sa.Length                                                            # 152
    $header = $iCookie + $rCookie + [byte[]]@(1, 0x10, 2, 0) `
              + (ConvertTo-BE32 0) + (ConvertTo-BE32 ([uint32]$total))

    return @{
        Packet          = $header + $sa
        InitiatorCookie = $iCookie
    }
}

# Describes an IKEv1 reply. A Main Mode response carrying an SA payload means
# the gateway ACCEPTED one of our proposals -- that is a live L2TP/IPsec head end.
function Get-IkeV1ReplyDescription {
    param([byte[]]$Bytes, [int]$Offset = 0)

    if ($Bytes.Length -lt ($Offset + 28)) { return 'truncated' }

    $nextPayload = $Bytes[$Offset + 16]
    $exch        = $Bytes[$Offset + 18]

    # IKEv1 has no "response" flag in the header. The tell is the responder
    # cookie: zero in a request, filled in by the gateway in a reply. Guards
    # against a middlebox echoing our own packet back at us.
    $respCookieSet = $false
    for ($i = 8; $i -lt 16; $i++) {
        if ($Bytes[$Offset + $i] -ne 0) { $respCookieSet = $true; break }
    }
    if (-not $respCookieSet) { return 'echoed request (responder cookie empty)' }

    $exchName = switch ($exch) {
        1 { 'Base' }
        2 { 'Main Mode' }
        4 { 'Aggressive' }
        5 { 'Informational' }
        default { "exchange $exch" }
    }

    # Walk payloads looking for Notification (type 11)
    $p   = $Offset + 28
    $nxt = $nextPayload
    while ($nxt -ne 0 -and ($p + 4) -le $Bytes.Length) {
        $thisType = $nxt
        $nxt      = $Bytes[$p]
        $len      = ($Bytes[$p + 2] -shl 8) -bor $Bytes[$p + 3]
        if ($len -le 0) { break }

        if ($thisType -eq 11 -and ($p + 12) -le $Bytes.Length) {
            # Notification: generic(4) DOI(4) ProtoID(1) SPISize(1) NotifyType(2)
            $ntype = ($Bytes[$p + 10] -shl 8) -bor $Bytes[$p + 11]
            $name  = switch ($ntype) {
                1  { 'INVALID_PAYLOAD_TYPE' }
                5  { 'INVALID_MAJOR_VERSION' }
                7  { 'INVALID_FLAGS' }
                9  { 'INVALID_PROTOCOL_ID' }
                14 { 'NO_PROPOSAL_CHOSEN' }
                24 { 'AUTHENTICATION_FAILED' }
                29 { 'ATTRIBUTES_NOT_SUPPORTED' }
                default { "notify $ntype" }
            }
            return "$exchName / $name"
        }
        $p += $len
    }

    if ($exch -eq 2 -and $nextPayload -eq 1) { return "$exchName / SA ACCEPTED" }
    return $exchName
}

# Decodes the IKEv2 payload chain enough to name the notify type, so a short
# reply can be reported as what it actually is instead of just "36 bytes".
function Get-IkeReplyDescription {
    param([byte[]]$Bytes, [int]$Offset = 0)

    if ($Bytes.Length -lt ($Offset + 28)) { return 'truncated' }

    $nextPayload = $Bytes[$Offset + 16]
    $flags       = $Bytes[$Offset + 19]
    $isResponse  = ($flags -band 0x20) -ne 0

    $desc = if ($isResponse) { 'response' } else { 'request' }

    # Walk to a Notify payload (type 41) if present
    $p   = $Offset + 28
    $nxt = $nextPayload
    while ($nxt -ne 0 -and $p + 4 -le $Bytes.Length) {
        $thisType = $nxt
        $nxt      = $Bytes[$p]
        $len      = ($Bytes[$p + 2] -shl 8) -bor $Bytes[$p + 3]
        if ($len -le 0) { break }

        if ($thisType -eq 41 -and ($p + 8) -le $Bytes.Length) {
            # Notify: NextPayload(1) Crit(1) Len(2) ProtoID(1) SPISize(1) Type(2)
            $ntype = ($Bytes[$p + 6] -shl 8) -bor $Bytes[$p + 7]

            # INVALID_KE_PAYLOAD means a proposal DID match - the gateway just
            # wants a different DH group. Its 2 bytes of notify data name it.
            # That is a stronger result than NO_PROPOSAL_CHOSEN, so say so.
            if ($ntype -eq 17 -and ($p + 10) -le $Bytes.Length) {
                $grp = ($Bytes[$p + 8] -shl 8) -bor $Bytes[$p + 9]
                $gname = switch ($grp) {
                    1  { 'MODP-768' }  2  { 'MODP-1024' } 5  { 'MODP-1536' }
                    14 { 'MODP-2048' } 15 { 'MODP-3072' } 16 { 'MODP-4096' }
                    19 { 'ECP-256' }   20 { 'ECP-384' }   21 { 'ECP-521' }
                    31 { 'Curve25519' }
                    default { "group $grp" }
                }
                return "$desc / proposal MATCHED, wants DH $gname"
            }
            $name  = switch ($ntype) {
                1     { 'UNSUPPORTED_CRITICAL_PAYLOAD' }
                4     { 'INVALID_IKE_SPI' }
                5     { 'INVALID_MAJOR_VERSION' }
                7     { 'INVALID_SYNTAX' }
                8     { 'INVALID_MESSAGE_ID' }
                11    { 'INVALID_SPI' }
                14    { 'NO_PROPOSAL_CHOSEN' }
                17    { 'INVALID_KE_PAYLOAD' }
                24    { 'AUTHENTICATION_FAILED' }
                34    { 'INVALID_SELECTORS' }
                16388 { 'NAT_DETECTION_SOURCE_IP' }
                16389 { 'NAT_DETECTION_DESTINATION_IP' }
                16390 { 'COOKIE' }
                default { "notify $ntype" }
            }
            return "$desc / $name"
        }
        $p += $len
    }

    if ($nextPayload -eq 33) { return "$desc / SA ACCEPTED" }
    return $desc
}

# ---------------------------------------------------------------------------
# Probes
# ---------------------------------------------------------------------------

function Test-UdpProbe {
    param(
        [string] $TargetHost,
        [int]    $Port,
        [byte[]] $Payload,
        [int]    $Timeout = 3000,
        [int]    $Retries = 2
    )

    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        $client = $null
        try {
            $client = New-Object System.Net.Sockets.UdpClient
            $client.Client.ReceiveTimeout = $Timeout
            $client.Client.SendTimeout    = $Timeout
            $client.Connect($TargetHost, $Port)

            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            [void]$client.Send($Payload, $Payload.Length)

            $remote = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
            $data   = $client.Receive([ref]$remote)
            $sw.Stop()

            return [pscustomobject]@{
                Responded = $true
                Bytes     = $data
                Length    = $data.Length
                Ms        = [int]$sw.ElapsedMilliseconds
                Reason    = 'Reply'
                Error     = $null
            }
        }
        catch [System.Net.Sockets.SocketException] {
            $code = $_.Exception.SocketErrorCode
            # ConnectionReset on a connected UDP socket == ICMP port unreachable.
            # That still proves the packet reached the host and came back.
            if ($code -eq 'ConnectionReset') {
                return [pscustomobject]@{
                    Responded = $false; Bytes = $null; Length = 0; Ms = 0
                    Reason = 'Refused'
                    Error  = 'ICMP port unreachable - host reachable, nothing listening on this port'
                }
            }
            if ($attempt -eq $Retries) {
                return [pscustomobject]@{
                    Responded = $false; Bytes = $null; Length = 0; Ms = 0
                    Reason = 'Filtered'; Error = "$code"
                }
            }
        }
        catch {
            if ($attempt -eq $Retries) {
                return [pscustomobject]@{
                    Responded = $false; Bytes = $null; Length = 0; Ms = 0
                    Reason = 'Error'; Error = $_.Exception.Message
                }
            }
        }
        finally {
            if ($null -ne $client) { $client.Close() }
        }
    }
}

<#
    TCP probe that distinguishes WHY a connect failed. This matters more than
    the pass/fail itself:

      Open        -> service is there and reachable
      Refused     -> RST came back. The packet reached the host and the host
                     answered. THE NETWORK IS NOT BLOCKING THIS PORT - the
                     service simply isn't listening. Wrong port, or gateway down.
      Filtered    -> silence until timeout. This is what an actual block looks
                     like: a firewall dropping your SYN on the floor.
      Unreachable -> ICMP host/net unreachable, i.e. no route.
#>
function Test-TcpProbe {
    param([string]$TargetHost, [int]$Port, [int]$Timeout = 3000)

    $client = New-Object System.Net.Sockets.TcpClient
    $sw     = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $iar = $client.BeginConnect($TargetHost, $Port, $null, $null)
        $ok  = $iar.AsyncWaitHandle.WaitOne($Timeout, $false)

        if (-not $ok) {
            return [pscustomobject]@{
                Connected = $false; Ms = $Timeout; Reason = 'Filtered'
                Error = "No response in $Timeout ms - SYN discarded, no RST. Some firewall dropped it; from here you cannot tell whether that was local or far-end"
            }
        }

        $client.EndConnect($iar)
        $sw.Stop()
        return [pscustomobject]@{
            Connected = $true; Ms = [int]$sw.ElapsedMilliseconds
            Reason = 'Open'; Error = $null
        }
    }
    catch {
        $sw.Stop()
        $ms = [int]$sw.ElapsedMilliseconds

        # PowerShell wraps the SocketException in a MethodInvocationException.
        # Walk the inner-exception chain to find the real socket error code.
        $sockEx = $null
        $ex     = $_.Exception
        while ($null -ne $ex) {
            if ($ex -is [System.Net.Sockets.SocketException]) { $sockEx = $ex; break }
            $ex = $ex.InnerException
        }
        $code = if ($null -ne $sockEx) { $sockEx.SocketErrorCode.ToString() } else { '' }

        if ($code -eq 'ConnectionRefused') {
            return [pscustomobject]@{
                Connected = $false; Ms = $ms; Reason = 'Refused'
                Error = "RST in $ms ms - host reachable, nothing listening on this port"
            }
        }
        if ($code -eq 'TimedOut') {
            return [pscustomobject]@{
                Connected = $false; Ms = $ms; Reason = 'Filtered'
                Error = "Timed out after $ms ms - SYN discarded, no RST. Some firewall dropped it; from here you cannot tell whether that was local or far-end"
            }
        }
        if ('HostUnreachable','NetworkUnreachable','HostNotFound','NetworkDown' -contains $code) {
            return [pscustomobject]@{
                Connected = $false; Ms = $ms; Reason = 'Unreachable'
                Error = "$code - no route to the host"
            }
        }

        $msg = if ($null -ne $sockEx) { $sockEx.Message } else { $_.Exception.Message }
        return [pscustomobject]@{
            Connected = $false; Ms = $ms; Reason = 'Error'; Error = $msg
        }
    }
    finally {
        $client.Close()
    }
}

function Get-TlsCertificate {
    param([string]$TargetHost, [int]$Port, [int]$Timeout = 5000)

    $client = New-Object System.Net.Sockets.TcpClient
    $ssl    = $null
    try {
        $iar = $client.BeginConnect($TargetHost, $Port, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($Timeout, $false)) {
            return [pscustomobject]@{ Ok = $false; Error = 'TCP timeout' }
        }
        $client.EndConnect($iar)

        # Accept anything -- we want to SEE the cert, not validate it.
        $validator = [System.Net.Security.RemoteCertificateValidationCallback]{ param($s,$c,$ch,$e) $true }
        $ssl = New-Object System.Net.Security.SslStream($client.GetStream(), $false, $validator)
        $ssl.AuthenticateAsClient($TargetHost)

        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($ssl.RemoteCertificate)

        return [pscustomobject]@{
            Ok         = $true
            Subject    = $cert.Subject
            Issuer     = $cert.Issuer
            NotAfter   = $cert.NotAfter
            NotBefore  = $cert.NotBefore
            Thumbprint = $cert.Thumbprint
            Protocol   = $ssl.SslProtocol.ToString()
            Error      = $null
        }
    }
    catch {
        return [pscustomobject]@{ Ok = $false; Error = $_.Exception.Message }
    }
    finally {
        if ($null -ne $ssl) { $ssl.Dispose() }
        $client.Close()
    }
}

# ---------------------------------------------------------------------------
# Test sections
# ---------------------------------------------------------------------------

function Invoke-NetworkBaseline {
    Write-Sub "Network baseline"

    try {
        $ipcfg = Get-NetIPConfiguration -ErrorAction Stop |
                 Where-Object { $_.IPv4DefaultGateway -and $_.NetAdapter.Status -eq 'Up' } |
                 Select-Object -First 1
        if ($ipcfg) {
            $detail = "{0} | IP {1} | GW {2}" -f `
                $ipcfg.InterfaceAlias,
                $ipcfg.IPv4Address.IPAddress,
                $ipcfg.IPv4DefaultGateway.NextHop
            Add-Result 'Baseline' 'Active adapter' 'INFO' $detail
        }
    } catch {
        Add-Result 'Baseline' 'Active adapter' 'SKIP' 'Get-NetIPConfiguration unavailable'
    }

    # --- Captive portal (Microsoft NCSI probe)
    try {
        $resp = Invoke-WebRequest -Uri 'http://www.msftconnecttest.com/connecttest.txt' `
                                  -UseBasicParsing -TimeoutSec 8 `
                                  -MaximumRedirection 0 -ErrorAction Stop
        if ($resp.Content.Trim() -eq 'Microsoft Connect Test') {
            Add-Result 'Baseline' 'Captive portal' 'PASS' 'None detected - internet is open'
        } else {
            Add-Result 'Baseline' 'Captive portal' 'FAIL' `
                'Response was rewritten - you are behind a portal. Log in first, then re-run.'
        }
    }
    catch {
        Add-Result 'Baseline' 'Captive portal' 'FAIL' `
            ('Probe failed/redirected - likely captive portal. ' + $_.Exception.Message)
    }

    # --- DNS integrity: a known-bogus name must NOT resolve
    $bogus = 'nxdomain-probe-' + (Get-Random) + '.invalid'
    $resolved = $null
    try   { $resolved = [System.Net.Dns]::GetHostAddresses($bogus) }
    catch { $resolved = $null }

    if ($resolved -and $resolved.Count -gt 0) {
        Add-Result 'Baseline' 'DNS integrity' 'WARN' `
            ('NXDOMAIN answered with ' + ($resolved[0].IPAddressToString) +
             ' - DNS is hijacked here. VPN hostname lookups may point at the wrong box.')
    } else {
        Add-Result 'Baseline' 'DNS integrity' 'PASS' 'NXDOMAIN honored - DNS not hijacked'
    }
}

function Invoke-UdpEgressControl {
    param([int]$Timeout)

    Write-Sub "UDP egress control (is UDP blocked in general, or just VPN?)"

    # DNS over UDP 53: A record query for example.com
    $dnsQuery = [byte[]]@(
        0xAB,0xCD, 0x01,0x00, 0x00,0x01, 0x00,0x00, 0x00,0x00, 0x00,0x00,
        0x07,0x65,0x78,0x61,0x6D,0x70,0x6C,0x65,          # 7 "example"
        0x03,0x63,0x6F,0x6D,                              # 3 "com"
        0x00,
        0x00,0x01, 0x00,0x01
    )
    $dns = Test-UdpProbe -TargetHost '8.8.8.8' -Port 53 -Payload $dnsQuery -Timeout $Timeout
    if ($dns.Responded) {
        Add-Result 'UDP control' 'UDP/53 to 8.8.8.8' 'PASS' `
            ("Reply in {0} ms - outbound UDP works" -f $dns.Ms)
    } else {
        Add-Result 'UDP control' 'UDP/53 to 8.8.8.8' 'WARN' `
            ('No reply (' + $dns.Error + ') - DNS is likely forced to the local resolver')
    }

    # STUN on a high, non-standard UDP port -- the closest analogue to UDP 4500.
    $stun = New-StunBindingRequest
    $sr   = Test-UdpProbe -TargetHost 'stun.l.google.com' -Port 19302 `
                          -Payload $stun.Packet -Timeout $Timeout
    if ($sr.Responded -and $sr.Length -ge 20 -and $sr.Bytes[0] -eq 0x01 -and $sr.Bytes[1] -eq 0x01) {
        Add-Result 'UDP control' 'UDP/19302 STUN' 'PASS' `
            ("Binding response in {0} ms - arbitrary high UDP ports are open" -f $sr.Ms)

        # The public IP the far end actually sees. Essential when the client is
        # behind NAT you do not control - this is the address to look for in a
        # firewall's Blocked Sites list, geo rules, or allow list.
        $mapped = Get-StunMappedAddress -Bytes $sr.Bytes
        if ($null -ne $mapped) {
            $script:PublicIp = $mapped.Address
            if (Test-CgnatAddress -Address $mapped.Address) {
                Add-Result 'UDP control' 'Public IP (seen by server)' 'WARN' `
                    ($mapped.Address + ' - CGNAT range. This address is SHARED with other ' +
                     'subscribers, so a far-end block may not have been caused by this user.')
            } else {
                Add-Result 'UDP control' 'Public IP (seen by server)' 'INFO' $mapped.Address
            }
        }
    }
    elseif ($sr.Responded) {
        Add-Result 'UDP control' 'UDP/19302 STUN' 'WARN' `
            ("Got {0} bytes but not a STUN response - something is mangling UDP" -f $sr.Length)
    }
    else {
        Add-Result 'UDP control' 'UDP/19302 STUN' 'FAIL' `
            ('No reply (' + $sr.Error + ') - high-port UDP appears blocked. IKEv2/WireGuard will not work here.')
    }
}

function Invoke-GatewayTests {
    param([string]$Target, [int]$SslPort, [string]$ExpectedSubject, [int]$Timeout)

    Write-Sub "Gateway: $Target"

    # ---- DNS resolution ---------------------------------------------------
    $ips = @()
    try {
        $ips = [System.Net.Dns]::GetHostAddresses($Target) |
               Where-Object { $_.AddressFamily -eq 'InterNetwork' } |
               ForEach-Object { $_.IPAddressToString }
        if ($ips.Count -gt 0) {
            Add-Result $Target 'DNS resolution' 'PASS' ($ips -join ', ')
        } else {
            Add-Result $Target 'DNS resolution' 'FAIL' 'No A record returned'
            return
        }
    }
    catch {
        Add-Result $Target 'DNS resolution' 'FAIL' $_.Exception.Message
        Add-Result $Target 'All gateway tests' 'SKIP' 'Cannot resolve name - is DNS filtered here?'
        return
    }

    # ---- ICMP ----------------------------------------------------------
    # Informational only. Plenty of gateways drop ping by policy, so a failure
    # here proves nothing on its own. But combined with "every port is silent"
    # it separates a whole-host blackhole from per-port filtering.
    try {
        $ping  = New-Object System.Net.NetworkInformation.Ping
        $reply = $ping.Send($Target, 2000)
        if ($reply.Status -eq 'Success') {
            Add-Result $Target 'ICMP echo' 'INFO' `
                ("Reply in {0} ms, TTL {1}" -f $reply.RoundtripTime, $reply.Options.Ttl)
        } else {
            Add-Result $Target 'ICMP echo' 'INFO' `
                ("$($reply.Status) - not conclusive, many gateways drop ping by design")
        }
    }
    catch {
        Add-Result $Target 'ICMP echo' 'INFO' 'Ping unavailable'
    }

    # ---- SSL VPN: TCP reachability ---------------------------------------
    Invoke-ProbePause
    $tcp = Test-TcpProbe -TargetHost $Target -Port $SslPort -Timeout $Timeout

    if ($tcp.Connected) {
        Add-Result $Target "TCP/$SslPort connect" 'PASS' ("Handshake in {0} ms" -f $tcp.Ms)
    }
    elseif ($tcp.Reason -eq 'Refused') {
        # NOT a network block. The host answered. Report it as such.
        Add-Result $Target "TCP/$SslPort connect" 'WARN' `
            ($tcp.Error + ' - network is NOT blocking this port; the gateway is not serving it')
    }
    elseif ($tcp.Reason -eq 'Filtered') {
        Add-Result $Target "TCP/$SslPort connect" 'FAIL' `
            ($tcp.Error + ' - see the diagnosis under VERDICT')
    }
    else {
        Add-Result $Target "TCP/$SslPort connect" 'FAIL' $tcp.Error
    }

    # ---- SSL VPN: TLS + certificate inspection ---------------------------
    if ($tcp.Connected) {
        $tls = Get-TlsCertificate -TargetHost $Target -Port $SslPort -Timeout ($Timeout * 2)
        if (-not $tls.Ok) {
            Add-Result $Target "TLS/$SslPort handshake" 'FAIL' `
                ($tls.Error + ' - port is open but TLS fails. Transparent proxy?')
        }
        else {
            Add-Result $Target "TLS/$SslPort handshake" 'PASS' `
                ("{0} | expires {1:yyyy-MM-dd}" -f $tls.Protocol, $tls.NotAfter)

            Add-Result $Target 'Cert subject' 'INFO' $tls.Subject
            Add-Result $Target 'Cert issuer'  'INFO' $tls.Issuer

            $inspectionHints = @(
                'Fortinet','FortiGate','Zscaler','Blue Coat','Bluecoat','Netskope',
                'Cisco Umbrella','Sophos','Untangle','Kaspersky','ESET','Bitdefender',
                'Avast','Palo Alto','McAfee','Symantec Web','iboss','Lightspeed',
                'Cloudflare Gateway','Forcepoint','Smoothwall','Barracuda','Meraki'
            )
            $hit = $inspectionHints | Where-Object { $tls.Issuer -like "*$_*" } | Select-Object -First 1

            if ($hit) {
                Add-Result $Target 'TLS inspection' 'FAIL' `
                    ("Certificate is issued by '$hit' - this network is decrypting TLS. " +
                     "Cert-pinning VPN clients will refuse to connect.")
            }
            elseif ($ExpectedSubject -and ($tls.Subject -notlike "*$ExpectedSubject*")) {
                Add-Result $Target 'TLS inspection' 'FAIL' `
                    ("Cert subject does not contain '$ExpectedSubject' - traffic is being intercepted.")
            }
            elseif ($tls.Subject -eq $tls.Issuer) {
                Add-Result $Target 'TLS inspection' 'WARN' `
                    'Self-signed certificate. Normal for some appliances, but verify the thumbprint.'
            }
            else {
                Add-Result $Target 'TLS inspection' 'PASS' `
                    'Certificate chain looks like the real gateway (no known interception CA)'
            }

            Add-Result $Target 'Cert thumbprint' 'INFO' $tls.Thumbprint
        }
    }

    # ---- IKEv2: UDP 500 ---------------------------------------------------
    Invoke-ProbePause
    $ike  = New-IkeV2SaInitPacket
    $r500 = Test-UdpProbe -TargetHost $Target -Port 500 -Payload $ike.Packet -Timeout $Timeout

    if ($r500.Responded) {
        $isIke = ($r500.Length -ge 28) -and ($r500.Bytes[17] -eq 0x20) -and ($r500.Bytes[18] -eq 34)
        if ($isIke) {
            $what = Get-IkeReplyDescription -Bytes $r500.Bytes -Offset 0
            Add-Result $Target 'UDP/500 IKEv2' 'PASS' `
                ("IKEv2 {0}, {1} bytes in {2} ms - path is open" -f $what, $r500.Length, $r500.Ms)
        } else {
            Add-Result $Target 'UDP/500 IKEv2' 'WARN' `
                ("Got {0} bytes but not a valid IKEv2 reply - path is open, responder may not be IKEv2" -f $r500.Length)
        }
    }
    elseif ($r500.Reason -eq 'Refused') {
        Add-Result $Target 'UDP/500 IKEv2' 'WARN' `
            ($r500.Error + ' - network is NOT blocking UDP/500; the gateway is not answering on it')
    }
    else {
        Add-Result $Target 'UDP/500 IKEv2' 'FAIL' `
            ('No IKE response (' + $r500.Error + '). Blocked here, or gateway silent on 500.')
    }

    # ---- IKEv2 NAT-T: UDP 4500 (4-byte non-ESP marker prefix) -------------
    Invoke-ProbePause
    $ikeNat  = New-IkeV2SaInitPacket
    $payload = [byte[]]@(0,0,0,0) + $ikeNat.Packet
    $r4500   = Test-UdpProbe -TargetHost $Target -Port 4500 -Payload $payload -Timeout $Timeout

    if ($r4500.Responded) {
        $off = 0
        if ($r4500.Length -ge 4 -and $r4500.Bytes[0] -eq 0 -and $r4500.Bytes[1] -eq 0 `
            -and $r4500.Bytes[2] -eq 0 -and $r4500.Bytes[3] -eq 0) { $off = 4 }
        $isIke = ($r4500.Length -ge ($off + 28)) `
                 -and ($r4500.Bytes[$off + 17] -eq 0x20) -and ($r4500.Bytes[$off + 18] -eq 34)
        if ($isIke) {
            $what = Get-IkeReplyDescription -Bytes $r4500.Bytes -Offset $off
            Add-Result $Target 'UDP/4500 IKEv2 NAT-T' 'PASS' `
                ("IKEv2 {0}, {1} bytes in {2} ms - path is open" -f $what, $r4500.Length, $r4500.Ms)
        } else {
            Add-Result $Target 'UDP/4500 IKEv2 NAT-T' 'WARN' `
                ("Got {0} bytes but not a valid IKEv2 reply" -f $r4500.Length)
        }
    }
    elseif ($r4500.Reason -eq 'Refused') {
        Add-Result $Target 'UDP/4500 IKEv2 NAT-T' 'WARN' `
            ($r4500.Error + ' - network is NOT blocking UDP/4500; the gateway is not answering on it')
    }
    else {
        Add-Result $Target 'UDP/4500 IKEv2 NAT-T' 'FAIL' `
            ('No NAT-T response (' + $r4500.Error + '). This is the port that matters most behind NAT.')
    }

    # ---- IKEv1 Main Mode: what L2TP/IPsec actually speaks ------------------
    Invoke-ProbePause
    $v1   = New-IkeV1MainModePacket
    $r1a  = Test-UdpProbe -TargetHost $Target -Port 500 -Payload $v1.Packet -Timeout $Timeout

    if ($r1a.Responded) {
        $isV1 = ($r1a.Length -ge 28) -and ($r1a.Bytes[17] -eq 0x10)
        if ($isV1) {
            $what = Get-IkeV1ReplyDescription -Bytes $r1a.Bytes -Offset 0
            $st   = if ($what -like '*SA ACCEPTED*') { 'PASS' } else { 'PASS' }
            Add-Result $Target 'UDP/500 IKEv1 (L2TP)' $st `
                ("IKEv1 {0}, {1} bytes in {2} ms" -f $what, $r1a.Length, $r1a.Ms)
        }
        elseif ($r1a.Length -ge 28 -and $r1a.Bytes[17] -eq 0x20) {
            Add-Result $Target 'UDP/500 IKEv1 (L2TP)' 'WARN' `
                ("Gateway answered in IKEv2 ({0} bytes) - it is IKEv2-only, not L2TP/IPsec" -f $r1a.Length)
        }
        else {
            Add-Result $Target 'UDP/500 IKEv1 (L2TP)' 'WARN' `
                ("Got {0} bytes, not recognisable ISAKMP - path open, responder unclear" -f $r1a.Length)
        }
    }
    elseif ($r1a.Reason -eq 'Refused') {
        Add-Result $Target 'UDP/500 IKEv1 (L2TP)' 'WARN' `
            ($r1a.Error + ' - network is NOT blocking UDP/500')
    }
    else {
        Add-Result $Target 'UDP/500 IKEv1 (L2TP)' 'FAIL' `
            ('No IKEv1 response (' + $r1a.Error + ')')
    }

    # IKEv1 over NAT-T (UDP 4500, non-ESP marker) -- the usual path for a
    # roaming L2TP client sitting behind hotel NAT.
    Invoke-ProbePause
    $v1n     = New-IkeV1MainModePacket
    $payload = [byte[]]@(0,0,0,0) + $v1n.Packet
    $r1b     = Test-UdpProbe -TargetHost $Target -Port 4500 -Payload $payload -Timeout $Timeout

    if ($r1b.Responded) {
        $off = 0
        if ($r1b.Length -ge 4 -and $r1b.Bytes[0] -eq 0 -and $r1b.Bytes[1] -eq 0 `
            -and $r1b.Bytes[2] -eq 0 -and $r1b.Bytes[3] -eq 0) { $off = 4 }

        if ($r1b.Length -ge ($off + 28) -and $r1b.Bytes[$off + 17] -eq 0x10) {
            $what = Get-IkeV1ReplyDescription -Bytes $r1b.Bytes -Offset $off
            Add-Result $Target 'UDP/4500 IKEv1 NAT-T' 'PASS' `
                ("IKEv1 {0}, {1} bytes in {2} ms" -f $what, $r1b.Length, $r1b.Ms)
        }
        else {
            Add-Result $Target 'UDP/4500 IKEv1 NAT-T' 'WARN' `
                ("Got {0} bytes, not IKEv1 - path is open" -f $r1b.Length)
        }
    }
    elseif ($r1b.Reason -eq 'Refused') {
        Add-Result $Target 'UDP/4500 IKEv1 NAT-T' 'WARN' `
            ($r1b.Error + ' - network is NOT blocking UDP/4500')
    }
    else {
        Add-Result $Target 'UDP/4500 IKEv1 NAT-T' 'FAIL' `
            ('No IKEv1 NAT-T response (' + $r1b.Error + ')')
    }

    # L2TP itself (UDP 1701) is deliberately NOT probed. A correctly configured
    # gateway only accepts L2TP inside the established IPsec SA and drops bare
    # UDP/1701, so a probe there fails on healthy boxes and proves nothing.
}

# ---------------------------------------------------------------------------
# Verdict
# ---------------------------------------------------------------------------

function Write-Verdict {
    param([string[]]$Targets)

    Write-Head "VERDICT"

    $portalFail = $script:Results | Where-Object { $_.Test -eq 'Captive portal' -and $_.Status -eq 'FAIL' }
    if ($portalFail) {
        Write-Host "  !! CAPTIVE PORTAL DETECTED." -ForegroundColor Yellow
        Write-Host "     Accept the terms in a browser, then run this again. Everything" -ForegroundColor Yellow
        Write-Host "     below is unreliable until you do." -ForegroundColor Yellow
        Write-Host ""
    }

    foreach ($t in $Targets) {
        $rows = $script:Results | Where-Object { $_.Category -eq $t }
        if (-not $rows) { continue }

        $tcpOk      = [bool]($rows | Where-Object { $_.Test -like 'TCP/*' -and $_.Status -eq 'PASS' })
        $tlsOk      = [bool]($rows | Where-Object { $_.Test -like 'TLS/*' -and $_.Status -eq 'PASS' })
        $mitm       = [bool]($rows | Where-Object { $_.Test -eq 'TLS inspection' -and $_.Status -eq 'FAIL' })
        $tcpRefused = [bool]($rows | Where-Object { $_.Test -like 'TCP/*' -and $_.Detail -like '*RST in*' })
        $sslOk      = $tcpOk -and $tlsOk -and (-not $mitm)

        $ikeOk = [bool](($rows | Where-Object { $_.Test -like 'UDP/4500 IKEv2*' -and $_.Status -eq 'PASS' }) -or
                        ($rows | Where-Object { $_.Test -like 'UDP/500 IKEv2*'  -and $_.Status -eq 'PASS' }))

        $v1Ok  = [bool](($rows | Where-Object { $_.Test -like '*IKEv1*' -and $_.Status -eq 'PASS' }))
        $v1Sa  = [bool](($rows | Where-Object { $_.Test -like '*IKEv1*' -and $_.Detail -like '*SA ACCEPTED*' }))

        # Distinguish "the gateway agreed to crypto terms" from "the gateway is
        # merely listening and rejected us". Both prove the path is open, but
        # only the first says IKEv2 is actually usable here.
        $v2Sa  = [bool](($rows | Where-Object { $_.Test -like '*IKEv2*' -and
                                    ($_.Detail -like '*SA ACCEPTED*' -or $_.Detail -like '*MATCHED*') }))
        $v2Rej = [bool](($rows | Where-Object { $_.Test -like '*IKEv2*' -and $_.Detail -like '*NO_PROPOSAL_CHOSEN*' }))

        Write-Host "  $t" -ForegroundColor White

        if ($sslOk) {
            Write-Host "    SSL VPN (TCP)   : WILL CONNECT" -ForegroundColor Green
        }
        elseif ($mitm) {
            Write-Host "    SSL VPN (TCP)   : REACHABLE BUT INTERCEPTED" -ForegroundColor Yellow
        }
        elseif ($tcpRefused) {
            Write-Host "    SSL VPN (TCP)   : NOT OFFERED (network is open, gateway refused)" -ForegroundColor Yellow
        }
        else {
            # NOT "blocked by this network". A silent drop is indistinguishable
            # from the far end discarding you - see the field note in the header.
            Write-Host "    SSL VPN (TCP)   : BLOCKED (local network OR far-end firewall)" -ForegroundColor Red
        }

        if ($v2Sa) {
            Write-Host "    IKEv2 (IPsec)   : WILL CONNECT (gateway accepted a proposal)" -ForegroundColor Green
        }
        elseif ($ikeOk -and $v2Rej) {
            Write-Host "    IKEv2 (IPsec)   : PATH OPEN, but gateway rejected all proposals" -ForegroundColor Yellow
            Write-Host "                      (network is fine - IKEv2 may not be configured here)" -ForegroundColor DarkGray
        }
        elseif ($ikeOk) {
            Write-Host "    IKEv2 (IPsec)   : PATH OPEN" -ForegroundColor Green
        }
        else {
            Write-Host "    IKEv2 (IPsec)   : BLOCKED / NO ANSWER" -ForegroundColor Red
        }

        if ($v1Sa) {
            Write-Host "    L2TP/IPsec (v1) : WILL CONNECT (gateway accepted a proposal)" -ForegroundColor Green
        }
        elseif ($v1Ok) {
            Write-Host "    L2TP/IPsec (v1) : WILL CONNECT" -ForegroundColor Green
        }
        else {
            Write-Host "    L2TP/IPsec (v1) : BLOCKED / NO ANSWER" -ForegroundColor Red
        }

        if ($mitm) {
            Write-Host "    -> TLS is being decrypted. Expect cert errors or a silent" -ForegroundColor Cyan
            Write-Host "       failure from a pinning client. Try IKEv2 or tether." -ForegroundColor Cyan
        }
        elseif (($ikeOk -or $v1Ok) -and $tcpRefused) {
            $use = if ($v1Sa) { 'L2TP/IPsec' } elseif ($ikeOk) { 'IKEv2' } else { 'L2TP/IPsec' }
            Write-Host "    -> Network is clean. Use $use - this gateway simply does not" -ForegroundColor Cyan
            Write-Host "       run an SSL VPN on that port." -ForegroundColor Cyan
        }
        elseif ($sslOk -and -not ($ikeOk -or $v1Ok)) {
            Write-Host "    -> Use the SSL VPN profile at this location." -ForegroundColor Cyan
        }
        elseif (($ikeOk -or $v1Ok) -and -not $sslOk) {
            $use = if ($v1Sa) { 'L2TP/IPsec' } elseif ($ikeOk) { 'IKEv2' } else { 'L2TP/IPsec' }
            Write-Host "    -> Use the $use profile at this location." -ForegroundColor Cyan
        }
        elseif (-not ($ikeOk -or $v1Ok) -and -not $sslOk) {
            Write-Host "    -> Nothing gets out to this gateway." -ForegroundColor Cyan
        }

        # ---- Blackhole discrimination -------------------------------------
        # If the UDP control probes passed but EVERY probe to this gateway went
        # silent, the network is not filtering VPNs by port - traffic to this
        # one destination is being dropped. That is a completely different
        # problem and points somewhere completely different.
        $controlOk = [bool]($script:Results | Where-Object {
            $_.Category -eq 'UDP control' -and $_.Test -like '*STUN*' -and $_.Status -eq 'PASS' })

        $allSilent = -not $sslOk -and -not $ikeOk -and -not $v1Ok -and -not $tcpRefused -and
                     [bool]($rows | Where-Object { $_.Detail -like '*TimedOut*' -or $_.Detail -like '*dropped*' })

        if ($controlOk -and $allSilent) {
            Write-Host ""
            Write-Host "    DESTINATION-SPECIFIC BLACKHOLE" -ForegroundColor Yellow
            Write-Host "    High-port UDP leaves this network fine, but every port on this" -ForegroundColor Yellow
            Write-Host "    gateway is silent. A network blocking VPNs by policy would kill" -ForegroundColor Yellow
            Write-Host "    500/4500 and leave the STUN control dead too. This pattern means" -ForegroundColor Yellow
            Write-Host "    traffic to this host specifically is being discarded." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "    Check, in order:" -ForegroundColor Cyan
            Write-Host "      1. FAR-END AUTO-BLOCK. Checked this first for a reason: a" -ForegroundColor Cyan
            Write-Host "         blocked source gets silent drops on EVERY port with no log" -ForegroundColor Cyan
            Write-Host "         entry - exactly this fingerprint. Causes are usually failed" -ForegroundColor Cyan
            Write-Host "         logins (a tech retrying bad creds will do it) or scan" -ForegroundColor Cyan
            Write-Host "         detection. On WatchGuard these live in TWO places and the" -ForegroundColor Cyan
            Write-Host "         obvious one is the wrong one:" -ForegroundColor Cyan
            Write-Host "           Firewall > Blocked Sites      = permanent, configured" -ForegroundColor DarkGray
            Write-Host "           System Status > Blocked Sites = TEMPORARY auto-blocks <-- this one" -ForegroundColor White
            if ($script:PublicIp) {
                Write-Host "         Look for this address: $script:PublicIp" -ForegroundColor White
                if (Test-CgnatAddress -Address $script:PublicIp) {
                    Write-Host "         (CGNAT - shared. Another subscriber may have tripped it.)" -ForegroundColor DarkGray
                }
            }
            Write-Host "      2. Geo-restriction or country block at the far end." -ForegroundColor Cyan
            Write-Host "      3. Local endpoint security / EDR with a per-destination rule." -ForegroundColor Cyan
            Write-Host "      4. Upstream routing blackhole - run a traceroute to see where" -ForegroundColor Cyan
            Write-Host "         it dies (at the ISP, or at the gateway's upstream hop)." -ForegroundColor Cyan
            Write-Host ""
            Write-Host "    Fastest test: tether to a phone and re-run. If it passes on" -ForegroundColor Cyan
            Write-Host "    cellular, the block is this network or its public IP, not the PC." -ForegroundColor Cyan
        }
        Write-Host ""
    }

    if (-not $Quick) {
        $stunFail = $script:Results | Where-Object { $_.Test -like '*STUN*' -and $_.Status -eq 'FAIL' }
        if ($stunFail) {
            Write-Host "  NOTE: high-port UDP is blocked network-wide, not just for your" -ForegroundColor Yellow
            Write-Host "        gateway. That rules out IKEv2, WireGuard and OpenVPN/UDP." -ForegroundColor Yellow
            Write-Host ""
        }
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

Write-Head "VPN PRE-FLIGHT CHECK   $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "  Host: $env:COMPUTERNAME    User: $env:USERNAME"
Write-Host "  PowerShell: $($PSVersionTable.PSVersion)"
if ($script:GentleMode) {
    Write-Host ("  Gentle mode: {0}s between probes (avoids IPS port-scan auto-block)" -f $script:GentleDelay) -ForegroundColor DarkGray
}

$targets = @()
if ($Gateway)  { $targets += $Gateway }
if ($Gateways) { $targets += $Gateways }

$configPath = Join-Path $script:ScriptDir 'vpncheck.config.json'
if ($targets.Count -eq 0 -and (Test-Path $configPath)) {
    try {
        $cfg = Get-Content $configPath -Raw | ConvertFrom-Json
        if (-not ($cfg.PSObject.Properties.Name -contains 'gateways')) {
            throw "config file has no 'gateways' array"
        }
        foreach ($g in $cfg.gateways) {
            $targets += $g.host
            if ($g.PSObject.Properties.Name -contains 'sslPort' -and $g.sslPort) { $SslPort = $g.sslPort }
            if ($g.PSObject.Properties.Name -contains 'expectedCertSubject' -and $g.expectedCertSubject) {
                $ExpectedCertSubject = $g.expectedCertSubject
            }
        }
        Write-Host "  Config: $configPath" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "  Config file found but could not be parsed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

if ($targets.Count -eq 0) {
    Write-Host ""
    Write-Host "  No gateway specified." -ForegroundColor Yellow
    $entered = Read-Host "  Enter VPN gateway hostname or IP"
    if ([string]::IsNullOrWhiteSpace($entered)) {
        Write-Host "  Nothing to test. Exiting." -ForegroundColor Red
        exit 1
    }
    $targets += $entered.Trim()
}

Invoke-NetworkBaseline

if (-not $Quick) {
    Invoke-UdpEgressControl -Timeout $TimeoutMs
} else {
    Write-Sub "UDP egress control"
    Add-Result 'UDP control' 'Control probes' 'SKIP' '-Quick specified'
}

foreach ($t in $targets) {
    Invoke-GatewayTests -Target $t -SslPort $SslPort `
                        -ExpectedSubject $ExpectedCertSubject -Timeout $TimeoutMs
}

Write-Verdict -Targets $targets

# ---- Report ---------------------------------------------------------------
if (-not $NoReport) {
    if (-not $ReportPath) {
        $stamp      = Get-Date -Format 'yyyyMMdd-HHmmss'
        $ReportPath = Join-Path $script:ScriptDir "vpncheck-$env:COMPUTERNAME-$stamp.txt"
    }

    $lines = @()
    $lines += "VPN Pre-Flight Check"
    $lines += "Date     : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')"
    $lines += "Host     : $env:COMPUTERNAME"
    $lines += "User     : $env:USERNAME"
    $lines += "Targets  : $($targets -join ', ')"
    $lines += ""
    $lines += ($script:Results | Format-Table Category, Status, Test, Detail -AutoSize -Wrap |
               Out-String -Width 200)

    try {
        $lines | Set-Content -Path $ReportPath -Encoding UTF8
        Write-Host "  Report saved: $ReportPath" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "  Could not write report: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Write-Host ""
if ($Host.Name -eq 'ConsoleHost' -and -not $env:VPNCHECK_NOPAUSE) {
    Write-Host "  Press any key to close..." -ForegroundColor DarkGray
    [void]$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}
