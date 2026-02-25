# Stale Profile Folder Analysis Script

## Overview
This PowerShell script helps identify and analyze stale user profile folders (like those created by VMware Horizon) by checking when files were last modified, cross-referencing with Active Directory, and calculating potential space savings.

## The Problem It Solves
VMware Horizon (and similar profile management systems) create user folders on file servers. Sometimes the parent folder's "Date Modified" doesn't reflect recent activity inside subdirectories. For example:
- `C:\Shares\Users\chris\` shows last modified 2018
- But `C:\Shares\Users\chris\Documents\report.docx` was modified yesterday

This script **recursively checks ALL files** in each user folder to find the true last activity date.

## Features
✅ Recursively scans all files in subdirectories  
✅ Identifies folders where NO files have been touched in X days  
✅ Cross-references usernames with Active Directory  
✅ Shows AD user's last logon date and enabled status  
✅ Calculates folder sizes for space analysis  
✅ Generates detailed CSV report  
✅ Provides summary statistics  
✅ Color-coded console output  

## Requirements
- Windows PowerShell 5.1 or PowerShell 7+
- Domain-joined computer (for AD lookups)
- Read access to the profile directories

**Two versions available:**
1. **Analyze-StaleProfiles-NoModule.ps1** ✅ **RECOMMENDED** - Uses .NET DirectoryServices, no module required
2. **Analyze-StaleProfiles.ps1** - Requires ActiveDirectory PowerShell module (RSAT)

## Installation
1. Save the script to your desired location
2. **Use the NoModule version** if you're getting module import errors
3. Ensure your machine is domain-joined for AD lookups to work

## Usage

### Basic Usage (No Module Required)
```powershell
.\Analyze-StaleProfiles-NoModule.ps1 -RootPath "C:\Shares\Users" -DaysInactive 365
```

### Specify Custom Export Path
```powershell
.\Analyze-StaleProfiles-NoModule.ps1 -RootPath "C:\Shares\Users" -DaysInactive 730 -ExportPath "C:\Reports\StaleProfiles.csv"
```

### Find Very Stale Profiles (2+ years)
```powershell
.\Analyze-StaleProfiles-NoModule.ps1 -RootPath "C:\Shares\Users" -DaysInactive 730
```

### Run with Elevated Privileges (Recommended)
```powershell
# Open PowerShell as Administrator, then:
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Analyze-StaleProfiles-NoModule.ps1 -RootPath "C:\Shares\Users" -DaysInactive 365
```

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `RootPath` | String | Yes | Root directory containing user folders (e.g., "C:\Shares\Users") |
| `DaysInactive` | Integer | Yes | Number of days to consider a folder stale (1-7300) |
| `ExportPath` | String | No | Custom path for CSV export (default: current directory with timestamp) |

## Output

### Console Output
The script provides real-time progress with color-coded status:
- **Green**: Active Directory module loaded, successful operations
- **Yellow**: Warnings, disabled users
- **Red**: No AD user found, errors
- **Gray**: Processing details
- **Cyan**: Headers and summaries

### CSV Export
The script automatically exports results to a CSV file with these columns:
- `FolderName` - Name of the user folder
- `FolderPath` - Full path to the folder
- `LastModifiedDate` - Most recent file modification date
- `DaysSinceLastActivity` - Days since last file was modified
- `SizeBytes` - Folder size in bytes
- `SizeFormatted` - Human-readable size (KB, MB, GB, TB)
- `FileCount` - Number of files in the folder
- `ADUserExists` - Whether user exists in Active Directory
- `ADLastLogon` - Last logon date from AD
- `ADEnabled` - Whether AD account is enabled

### Summary Statistics
The script provides a comprehensive summary including:
- Total folders scanned
- Number of stale folders found
- Folders with no AD user
- Folders with disabled AD users
- Total size of stale folders
- Potential space savings
- Top 10 largest stale folders

## Example Output

```
═══════════════════════════════════════════════════════════════
  Stale Profile Folder Analysis
═══════════════════════════════════════════════════════════════
Root Path:       C:\Shares\Users
Days Inactive:   365 days
Cutoff Date:     2025-02-05
Start Time:      2026-02-05 14:30:22
═══════════════════════════════════════════════════════════════

[✓] Active Directory module loaded successfully

Scanning user profile directories...
Found 150 user folders to analyze

[→] Analyzing: jsmith (Last modified: 2023-05-12)
    ├─ Size: 2.34 GB | Files: 1,245 | AD User: False | Enabled: N/A

[→] Analyzing: obsolete_user (Last modified: 2022-11-03)
    ├─ Size: 850.23 MB | Files: 523 | AD User: False | Enabled: N/A

═══════════════════════════════════════════════════════════════
  ANALYSIS SUMMARY
═══════════════════════════════════════════════════════════════

Total Folders Scanned:           150
Stale Folders Found:             23
Folders with No AD User:         15
Folders with Disabled AD User:   5
Errors Encountered:              0

Total Size of Stale Folders:     45.67 GB
Size of No AD User Folders:      28.34 GB
Size of Disabled User Folders:   12.89 GB

Potential Space Savings:         45.67 GB

═══════════════════════════════════════════════════════════════
  TOP 10 LARGEST STALE FOLDERS
═══════════════════════════════════════════════════════════════

1. old_contractor [NO AD USER]
   Size: 8.45 GB | Last Activity: 2023-01-15 14:22:33 | Days: 1,117

2. jdoe [DISABLED]
   Size: 5.23 GB | Last Activity: 2023-08-22 09:15:11 | Days: 897

...

═══════════════════════════════════════════════════════════════
[✓] Results exported to: StaleProfiles_20260205_143045.csv
═══════════════════════════════════════════════════════════════

Analysis completed in 03:45
```

## Common Use Cases

### 1. Identify Old Profiles for Archival (1 year)
```powershell
.\Analyze-StaleProfiles-NoModule.ps1 -RootPath "C:\Shares\Users" -DaysInactive 365
```

### 2. Find Ancient Profiles for Deletion (2 years)
```powershell
.\Analyze-StaleProfiles-NoModule.ps1 -RootPath "C:\Shares\Users" -DaysInactive 730
```

### 3. Quarterly Cleanup Check (90 days)
```powershell
.\Analyze-StaleProfiles-NoModule.ps1 -RootPath "C:\Shares\Users" -DaysInactive 90
```

### 4. Generate Monthly Report
```powershell
$ReportDate = Get-Date -Format "yyyy-MM"
.\Analyze-StaleProfiles-NoModule.ps1 -RootPath "C:\Shares\Users" -DaysInactive 180 -ExportPath "C:\Reports\StaleProfiles_$ReportDate.csv"
```

## Workflow Recommendations

1. **Run Analysis**: Execute the script with appropriate days threshold
2. **Review CSV**: Open the CSV in Excel to sort and filter results
3. **Verify with IT**: Cross-check with HR or IT for any false positives
4. **Archive First**: Move stale folders to archive location before deletion
5. **Document**: Keep the CSV report for compliance/audit purposes
6. **Test Deletion**: Start with non-AD users or disabled accounts
7. **Monitor**: Wait 30 days after archival before permanent deletion

## Troubleshooting

### "Execution Policy" Error
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

### "ActiveDirectory module not found"
**Solution:** Use the **Analyze-StaleProfiles-NoModule.ps1** version instead! It uses .NET DirectoryServices and doesn't require any modules.

### "Not connected to a domain or AD unavailable"
The script must run on a domain-joined computer to perform AD lookups. The script will still work but will skip AD user information.

### "Access Denied" Errors
Run PowerShell as Administrator and ensure you have read permissions on the target directories.

### Script Runs Slowly
This is normal for large file servers. The script must scan every file recursively. Consider:
- Running during off-hours
- Testing on a smaller subset first
- Running from a server closer to the file share

## Important Notes

⚠️ **Before Deleting**: Always archive folders before permanent deletion  
⚠️ **Test First**: Run on a test/sample directory before production  
⚠️ **Verify Results**: Some users may have legitimate old profiles  
⚠️ **Check Regulations**: Ensure compliance with data retention policies  

## License
This script is provided as-is for system administration purposes.

## Support
For issues or questions, consult your IT department or system administrator.
