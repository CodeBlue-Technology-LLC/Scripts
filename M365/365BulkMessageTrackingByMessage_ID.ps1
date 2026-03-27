# ============================================================
# Exchange Online - Message ID Lookup from CSV
# Requires: ExchangeOnlineManagement module
# Install:  Install-Module -Name ExchangeOnlineManagement
# Script is rate limited for large csv's
# ============================================================
#
# Usage
#	.\365BulkMessageTrackByMessage_ID.ps1 -InputCsv "C:\temp\SmtpCSReport_Custom Report - SmtpCSReport - 3262026_8644ff59-c6e5-4886-8c81-0481d2206665.csv" -OutputCsv "C:\temp\results.csv"
#

param(
    [Parameter(Mandatory)]
    [string]$InputCsv,

    [Parameter(Mandatory)]
    [string]$OutputCsv,

    [int]$DaysBack       = 10,
    [int]$BatchSize      = 90,
    [int]$BatchPauseSec  = 320
)

# ---- Connect ----
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline -ShowBanner:$false

# ---- Import CSV ----
if (-not (Test-Path $InputCsv)) {
    Write-Error "Input CSV not found: $InputCsv"
    exit 1
}

$records = Import-Csv -Path $InputCsv -Encoding Unicode

if (-not ($records | Get-Member -Name "message_id" -ErrorAction SilentlyContinue)) {
    Write-Error "Column 'message_id' not found in CSV. Please verify headers."
    exit 1
}

Write-Host "Loaded $($records.Count) records from CSV." -ForegroundColor Cyan

$endDate      = Get-Date
$startDate    = $endDate.AddDays(-$DaysBack)
$totalBatches = [math]::Ceiling($records.Count / $BatchSize)
$results      = [System.Collections.Generic.List[PSCustomObject]]::new()
$totalDone    = 0
$batchNum     = 0

Write-Host "Processing in $totalBatches batches of $BatchSize with ${BatchPauseSec}s pause between batches." -ForegroundColor Cyan

# ---- Helper function with retry ----
function Invoke-MessageTrace {
    param(
        [string]$Id,
        [datetime]$Start,
        [datetime]$End,
        [int]$MaxRetries = 3
    )
    $attempt = 0
    while ($attempt -lt $MaxRetries) {
        try {
            $r = Get-MessageTraceV2 -MessageId $Id -StartDate $Start -EndDate $End -ErrorAction Stop
            return $r
        }
        catch {
            $errMsg = $_.ToString()
            if ($errMsg -match "surpassed the permitted limit") {
                $attempt++
                if ($attempt -lt $MaxRetries) {
                    Write-Warning "Rate limited on $Id - waiting 60s before retry $attempt of $MaxRetries..."
                    Start-Sleep -Seconds 60
                }
                else {
                    throw
                }
            }
            else {
                throw
            }
        }
    }
}

# ---- Process batches ----
for ($batchStart = 0; $batchStart -lt $records.Count; $batchStart += $BatchSize) {
    $batchNum++
    $batchEnd = [math]::Min($batchStart + $BatchSize - 1, $records.Count - 1)
    $batch    = $records[$batchStart..$batchEnd]
    $batchMsg = "Batch $batchNum of $totalBatches - $($batch.Count) records"
    Write-Host $batchMsg -ForegroundColor Yellow

    foreach ($rec in $batch) {
        $totalDone++
        $msgId = $rec.message_id.Trim()

        if ([string]::IsNullOrWhiteSpace($msgId)) {
            Write-Warning "Blank message_id at record $totalDone, skipping."
            continue
        }

        $pct    = [math]::Round(($totalDone / $records.Count) * 100)
        $status = "Batch $batchNum/$totalBatches | Record $totalDone/$($records.Count)"
        Write-Progress -Activity "Looking up messages" -Status $status -PercentComplete $pct

        if ($msgId -match "^<.*>$") {
            $id = $msgId
        }
        else {
            $id = "<$msgId>"
        }

        try {
            $trace = Invoke-MessageTrace -Id $id -Start $startDate -End $endDate

            if ($trace) {
                foreach ($t in $trace) {
                    $results.Add([PSCustomObject]@{
                        OriginalDate         = $rec.date
                        OriginalSenderName   = $rec.senderName
                        OriginalSenderDomain = $rec.senderDomain
                        AuthMethod           = $rec.authMethod
                        OriginalClientIP     = $rec.original_client_ip
                        TLS_Version          = $rec.tls_version
                        TLS_Cipher           = $rec.tls_cipher
                        MessageCount         = $rec.message_count
                        MessageId            = $t.MessageId
                        Subject              = $t.Subject
                        Sender               = $t.SenderAddress
                        Recipient            = $t.RecipientAddress
                        Received             = $t.Received
                        Status               = $t.Status
                        FromIP               = $t.FromIP
                    })
                }
            }
            else {
                $results.Add([PSCustomObject]@{
                    OriginalDate         = $rec.date
                    OriginalSenderName   = $rec.senderName
                    OriginalSenderDomain = $rec.senderDomain
                    AuthMethod           = $rec.authMethod
                    OriginalClientIP     = $rec.original_client_ip
                    TLS_Version          = $rec.tls_version
                    TLS_Cipher           = $rec.tls_cipher
                    MessageCount         = $rec.message_count
                    MessageId            = $msgId
                    Subject              = "NOT FOUND"
                    Sender               = ""
                    Recipient            = ""
                    Received             = ""
                    Status               = "NOT FOUND"
                    FromIP               = ""
                })
            }
        }
        catch {
            Write-Warning "Error looking up $msgId : $_"
        }
    }

    # Pause between batches (skip after last batch)
    if ($batchNum -lt $totalBatches) {
        Write-Host "Batch $batchNum complete. Pausing ${BatchPauseSec}s to reset throttle window..." -ForegroundColor Cyan
        for ($s = $BatchPauseSec; $s -gt 0; $s--) {
            $coolPct = [math]::Round((($BatchPauseSec - $s) / $BatchPauseSec) * 100)
            Write-Progress -Activity "Throttle cooldown" -Status "Resuming in $s seconds..." -PercentComplete $coolPct
            Start-Sleep -Seconds 1
        }
        Write-Progress -Activity "Throttle cooldown" -Completed
    }
}

Write-Progress -Activity "Looking up messages" -Completed

# ---- Export ----
$results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
Write-Host "Done! $($results.Count) rows written to: $OutputCsv" -ForegroundColor Green

$found    = ($results | Where-Object { $_.Status -ne "NOT FOUND" }).Count
$notFound = ($results | Where-Object { $_.Status -eq "NOT FOUND" }).Count
Write-Host "  Found    : $found" -ForegroundColor Green
Write-Host "  Not Found: $notFound" -ForegroundColor Yellow

Disconnect-ExchangeOnline -Confirm:$false
