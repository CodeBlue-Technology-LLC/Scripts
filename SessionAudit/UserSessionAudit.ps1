$events = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    Id = 506, 507
    StartTime = (Get-Date).AddDays(-14)
} -ErrorAction SilentlyContinue | Sort-Object TimeCreated

$results = @()
for ($i = 0; $i -lt ($events.Count - 1); $i++) {
    if ($events[$i].Id -eq 506 -and $events[$i+1].Id -eq 507) {
        $standbyStart = $events[$i].TimeCreated
        $standbyEnd   = $events[$i+1].TimeCreated
        $duration     = $standbyEnd - $standbyStart

        # Only flag during business hours (7am-6pm) and over 20 minutes
        $hour = $standbyStart.Hour
        if ($hour -ge 7 -and $hour -le 18 -and $duration.TotalMinutes -gt 20) {
            $results += [PSCustomObject]@{
                Date       = $standbyStart.ToString("MM/dd/yyyy")
                DayOfWeek  = $standbyStart.DayOfWeek
                WentIdle   = $standbyStart.ToString("hh:mm tt")
                CameBack   = $standbyEnd.ToString("hh:mm tt")
                Duration   = $duration.ToString("hh\:mm\:ss")
                Minutes    = [math]::Round($duration.TotalMinutes, 0)
            }
        }
    }
}

if ($results) {
    $results | Export-Csv "C:\Windows\Temp\idle_gaps.csv" -NoTypeInformation
    $results | Format-Table -AutoSize
} else {
    Write-Host "No significant idle periods found during business hours."
}
