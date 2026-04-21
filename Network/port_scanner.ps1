param(
    [Parameter(Mandatory=$true)]
    [string]$Target,

    [int]$StartPort = 1,
    [int]$EndPort   = 65535,
    [int]$Timeout   = 100,
    [int]$Threads   = 200
)

Write-Host "`nScanning $Target  ports $StartPort-$EndPort  ($Threads threads, ${Timeout}ms timeout)" -ForegroundColor Cyan
$sw = [System.Diagnostics.Stopwatch]::StartNew()

$pool = [RunspaceFactory]::CreateRunspacePool(1, $Threads)
$pool.Open()

$scriptBlock = {
    param($ip, $port, $timeout)
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $ar  = $tcp.BeginConnect($ip, $port, $null, $null)
        $ok  = $ar.AsyncWaitHandle.WaitOne($timeout, $false)
        if ($ok -and $tcp.Connected) { return $port }
        $tcp.Close()
    } catch {}
    return $null
}

# Spin up all jobs
$jobs = foreach ($port in $StartPort..$EndPort) {
    $ps = [PowerShell]::Create()
    $ps.RunspacePool = $pool
    [void]$ps.AddScript($scriptBlock).AddArgument($Target).AddArgument($port).AddArgument($Timeout)
    [PSCustomObject]@{ PS = $ps; Handle = $ps.BeginInvoke() }
}

Write-Host "All jobs queued, waiting for results..." -ForegroundColor DarkGray

# Collect results
$open = @()
foreach ($job in $jobs) {
    $result = $job.PS.EndInvoke($job.Handle) | Where-Object { $_ -ne $null }
    if ($result) { $open += [int]$result }
    $job.PS.Dispose()
}

$pool.Close()
$pool.Dispose()
$sw.Stop()

Write-Host ""
if ($open.Count -eq 0) {
    Write-Host "No open ports found on $Target." -ForegroundColor Yellow
} else {
    Write-Host "Open ports on ${Target}:" -ForegroundColor Green
    $open | Sort-Object | ForEach-Object {
        $svc = switch ($_) {
            21   {"FTP"}     22   {"SSH"}     23   {"Telnet"}
            25   {"SMTP"}    53   {"DNS"}     80   {"HTTP"}
            110  {"POP3"}    135  {"RPC"}     139  {"NetBIOS"}
            143  {"IMAP"}    389  {"LDAP"}    443  {"HTTPS"}
            445  {"SMB"}     636  {"LDAPS"}   993  {"IMAPS"}
            995  {"POP3S"}   1433 {"MSSQL"}   1723 {"PPTP"}
            3306 {"MySQL"}   3389 {"RDP"}     5985 {"WinRM-HTTP"}
            5986 {"WinRM-HTTPS"} 8080 {"HTTP-Alt"} 8443 {"HTTPS-Alt"}
            default {""}
        }
        $label = if ($svc) { "  $_ ($svc)" } else { "  $_" }
        Write-Host $label -ForegroundColor Green
    }
}

Write-Host "`nDone in $([math]::Round($sw.Elapsed.TotalSeconds,1))s`n" -ForegroundColor Cyan
