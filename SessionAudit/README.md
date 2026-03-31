# Get-IdleGaps.ps1

A PowerShell script that audits user idle periods on Windows machines by parsing **Modern Standby** events (506/507) from the Windows System Event Log. Useful for identifying unexplained absences during business hours without requiring any changes to audit policy.

---

## How It Works

Windows logs two event IDs to the **System** log whenever the screen times out or the machine enters/exits standby:

| Event ID | Meaning |
|----------|---------|
| `506` | System entering Modern Standby (screen off / idle) |
| `507` | System exiting Modern Standby (user returned) |

The script pairs these events, calculates the duration between them, and filters to only show gaps during business hours that exceed a configurable threshold.

---

## Requirements

- Windows 10/11 (Modern Standby must be supported by the hardware)
- PowerShell 5.1 or later
- Local admin rights (to read the Security/System event log)
- No audit policy changes required — these events are logged by default

---

## Usage

Run locally on the target machine:

```powershell
.\Get-IdleGaps.ps1
```

Or push via RMM tool (e.g. ConnectWise Automate) to run silently and drop a CSV:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Get-IdleGaps.ps1
```

Results are saved to:
```
C:\Windows\Temp\idle_gaps.csv
```

---

## Configuration

Edit these variables at the top of the script to suit your needs:

| Variable | Default | Description |
|----------|---------|-------------|
| `$daysBack` | `14` | How many days of history to pull |
| `$businessStart` | `7` | Start of business hours (24hr) |
| `$businessEnd` | `18` | End of business hours (24hr) |
| `$minIdleMinutes` | `20` | Minimum idle duration to report (minutes) |

---

## Output

The script outputs a table (and CSV) with the following columns:

| Column | Description |
|--------|-------------|
| `Date` | Date of the idle period |
| `DayOfWeek` | Day name |
| `WentIdle` | Time the screen went idle |
| `CameBack` | Time activity resumed |
| `Duration` | Gap in `hh:mm:ss` format |
| `Minutes` | Gap in total minutes |

### Example Output

```
Date       DayOfWeek WentIdle CameBack Duration Minutes
----       --------- -------- -------- -------- -------
02/11/2026 Wednesday 12:36 PM 01:15 PM 00:39:08      39
02/12/2026  Thursday 12:58 PM 01:29 PM 00:31:24      31
02/16/2026    Monday 12:35 PM 01:13 PM 00:38:08      38
```

---

## Limitations

- Only logs gaps **going forward** from when the script is first run — no retroactive data beyond what's in the event log (typically 1–4 weeks depending on log size)
- Does not distinguish between a locked screen and a powered-off monitor — both show as standby
- Modern Standby must be supported; older machines using traditional S3 sleep may log differently
- Evening/weekend gaps are filtered out by default but can be adjusted via the config variables

---

## Deploying via ConnectWise Automate

1. Upload the script to your Automate script library
2. Target the machine or group you want to audit
3. Execute with local system context (admin rights required)
4. Retrieve `C:\Windows\Temp\idle_gaps.csv` via file transfer or include an upload step in your script

---

## Related Event IDs

If you later enable audit policy (`auditpol /set /subcategory:"Other Logon/Logoff Events" /success:enable`), you can supplement this with:

| Event ID | Log | Meaning |
|----------|-----|---------|
| `4800` | Security | Workstation locked |
| `4801` | Security | Workstation unlocked |
| `4624` | Security | Logon |
| `4634` | Security | Logoff |

---

## License

MIT
