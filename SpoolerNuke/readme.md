# SpoolerNuke.ps1

Nukes and rebuilds the Windows Print Spooler environment on a print server or RDS host. Removes all third-party drivers, print processors, port monitors, and driver files, then restarts the spooler clean. Printers re-map automatically via Group Policy after cleanup.

---

## When to Use This

- Print Spooler crashing immediately on start with no useful events
- Spooler enters a crash loop after a Windows Update (known issue with cumulative updates touching pscript5/spooler components)
- Corrupt or mismatched third-party driver DLLs in `spool\DRIVERS\x64\3`
- RDS/Terminal Server with an accumulation of old vendor drivers causing instability
- You've already tried removing the offending driver and the spooler is still dying

**Not for:** Single workstation printer issues, stuck print jobs, or spooler that starts but behaves badly. This is for when the spooler won't stay running at all.

---

## What It Does

1. Stops `PrintFilterPipelineSvc` and `Spooler`
2. Clears all pending jobs from `spool\PRINTERS`
3. Removes all third-party driver registry entries under `Version-3` — keeps Microsoft inbox drivers and Remote Desktop Easy Print
4. Removes all third-party print processors — keeps `winprint` only
5. Removes all third-party port monitors — keeps `Local Port`, `Standard TCP/IP Port`, `USB Monitor`, `WSD Port`
6. Deletes all files from `spool\DRIVERS\x64\3`
7. Starts the Spooler and reports status

Everything is logged to `C:\Temp\SpoolerReset_<timestamp>.log`

---

## What It Keeps

| Type | Kept |
|------|------|
| Drivers | Microsoft enhanced Point and Print compatibility driver, Remote Desktop Easy Print |
| Print Processors | winprint |
| Port Monitors | Local Port, Standard TCP/IP Port, USB Monitor, WSD Port |

Everything else gets removed.

---

## Requirements

- Must be run as **Administrator**
- Spooler does **not** need to be running beforehand — script handles a dead spooler
- Intended for **print servers and RDS hosts** where printers are mapped via GPO
- Not intended for standalone workstations where drivers aren't managed centrally

---

## After Running

- Spooler should start clean within a few seconds
- If printers are mapped via GPO, they will repopulate automatically as users log in or on the next GP refresh (`gpupdate /force` to speed it up)
- Drivers will be pulled fresh from the DC/print server as printers are reconnected
- Check `C:\Temp\SpoolerReset_<timestamp>.log` to confirm what was removed

---

## If the Spooler Still Dies After Running

The script logs final spooler status. If it shows `FAILED`:

1. Check `Microsoft-Windows-PrintService/Admin` event log immediately after attempting to start
2. Check `C:\ProgramData\Microsoft\Windows\WER\ReportArchive` for fresh crash dumps
3. Consider whether a recent Windows Update is the culprit — check `Get-HotFix | Sort-Object InstalledOn -Descending`
4. If a KB is suspect, test uninstalling it in a maintenance window

---

## Background / War Story

This script came out of a March 2026 incident where KB5078752 (Server 2019 cumulative update) caused the print spooler on an RDS host to crash immediately on start. Root cause was a combination of an ancient Lexmark Universal PS3 driver (2016 vintage) with a mismatched `PS5UI.DLL` version, plus accumulated third-party driver garbage from years of printer sprawl. Nuking the driver environment and letting GPO repopulate it cleanly resolved the issue in under 2 minutes.

---

## Author

CodeBlue Technology LLC  
Internal tooling — not for distribution
