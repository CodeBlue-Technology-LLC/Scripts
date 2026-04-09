# Stale Profile & UPD Analysis Scripts

## Overview
Two PowerShell scripts for identifying stale user profiles and User Profile Disks (UPDs). Both cross-reference Active Directory using .NET DirectoryServices — no ActiveDirectory module required.

---

## Scripts

### 1. Analyze-StaleProfiles-NoModule.ps1
Targets traditional user profile **folder** shares (e.g., VMware Horizon roaming profiles, home drives). Recursively scans all files in each user folder to find the true last activity date, since parent folder timestamps are often unreliable.

### 2. Analyze-StaleUPDProfiles-NoModule.ps1
Targets **User Profile Disk** shares containing `UVHD-*.vhdx` files named by SID. Uses VHDX `LastWriteTime` as the activity indicator — login activity is sufficient to update this timestamp. Resolves SIDs to AD accounts automatically.

---

## Requirements
- Windows PowerShell 5.1 or PowerShell 7+
- Domain-joined computer (for AD lookups)
- Read access to the target share
- Run as Administrator recommended

---

## Usage

### Stale Profile Folders
```powershell
.\Analyze-StaleProfiles-NoModule.ps1 -RootPath "C:\Shares\Users" -DaysInactive 365
.\Analyze-StaleProfiles-NoModule.ps1 -RootPath "C:\Shares\Users" -DaysInactive 730 -ExportPath "C:\Reports\StaleProfiles.csv"
```

### Stale UPD Disks
```powershell
.\Analyze-StaleUPDProfiles-NoModule.ps1 -UPDPath "C:\Shares\UPD" -DaysInactive 180
.\Analyze-StaleUPDProfiles-NoModule.ps1 -UPDPath "\\fileserver\UPDs" -DaysInactive 365 -ExportPath "C:\Reports\StaleUPDs.csv"
```

---

## Parameters

### Analyze-StaleProfiles-NoModule.ps1

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `RootPath` | String | Yes | Root directory containing user folders |
| `DaysInactive` | Integer | Yes | Days without file activity to flag as stale (1-7300) |
| `ExportPath` | String | No | CSV export path (default: current directory with timestamp) |

### Analyze-StaleUPDProfiles-NoModule.ps1

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `UPDPath` | String | Yes | Path to UPD share containing UVHD-*.vhdx files |
| `DaysInactive` | Integer | Yes | Days since last VHDX write to flag as stale (1-7300) |
| `ExportPath` | String | No | CSV export path (default: current directory with timestamp) |

---

## Output

### Console
Color-coded real-time progress:
- **Red** — No AD user found / orphaned SID
- **Yellow** — Disabled AD account
- **DarkYellow** — AD lookup error
- **White** — Active account
- **Cyan** — Headers and summary

### CSV Export
Both scripts export a CSV with per-entry details. UPD script includes SID, resolved username, and display name. Profile script includes folder file count.

### Summary Statistics
- Total files/disks scanned
- Stale count
- Orphaned (no AD user) count and size
- Disabled account count and size
- AD lookup error count
- Top 10 largest stale entries

---

## AD Lookup Notes

Both scripts use .NET `DirectoryServices` — no RSAT or ActiveDirectory module needed.

The UPD script resolves `SID → NTAccount → sAMAccountName` via LDAP. AD lookup errors in the results typically indicate a GC connectivity issue or cross-domain trust, not orphaned accounts — verify before treating as safe to delete.

`lastLogon` is used for last logon reporting. This attribute is per-DC and does not replicate. In multi-DC environments this may not reflect the true last logon. `lastLogonTimestamp` replicates but lags up to 14 days by design — either is acceptable for stale detection purposes.

---

## Recommended Workflow

1. Run the appropriate script against your share
2. Open the CSV in Excel — sort by `DaysSinceLastWrite` or `SizeBytes`
3. Prioritize **Orphaned SIDs** and **Disabled accounts** for removal — lowest risk
4. Verify **Active** flagged entries with HR or the user's manager before touching
5. Archive before deleting — move to a holding location and wait 30 days
6. Keep the CSV for audit/compliance documentation

---

## Troubleshooting

**Execution policy error**
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

**No UVHD files found**
Confirm files follow the `UVHD-S-1-5-21-*.vhdx` naming convention. The script skips `UVHD-template.vhdx` and any non-SID named files automatically.

**AD lookup errors in results**
Script is domain-joined and AD is reachable, but individual SID resolution failed. Usually a GC or trust issue. Do not treat as orphaned without manual verification.

**Access denied errors**
Run PowerShell as Administrator. Ensure the account has read access to the share.

**Slow performance (profile folder script)**
Expected on large shares — the script must enumerate every file recursively. Run during off-hours or test on a subset first. The UPD script is significantly faster since it only reads VHDX file metadata.

---

## Important Notes

> ⚠️ Always archive before deleting — moves are recoverable, deletes often aren't  
> ⚠️ AD lookup errors ≠ orphaned accounts — verify manually  
> ⚠️ Check your data retention policy before any deletion  
> ⚠️ Test on a non-production share first  

---

## License
Provided as-is for system administration use.
