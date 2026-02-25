# CreateOnedriveShortcutSharepoint.ps1

Creates OneDrive shortcuts to SharePoint sites for Microsoft 365 users. Automatically discovers which SharePoint sites a user has access to and maps them as shortcuts in their OneDrive root.

## How It Works

1. Connects to Microsoft Graph and SharePoint Online using admin credentials
2. Discovers SharePoint sites each user has access to
3. For each site, creates a OneDrive shortcut pointing to the `/General` folder (Teams sites) or the Documents library root (non-Teams sites)
4. Skips sites that are already mapped

## Prerequisites

- SharePoint Admin or Global Admin permissions
- The following PowerShell modules (auto-installed if missing):
  - `Microsoft.Graph.Authentication`
  - `Microsoft.Graph.Sites`
  - `Microsoft.Graph.Users`
  - `Microsoft.Online.SharePoint.PowerShell`

## Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `-UserEmails` | `string[]` | One or more user email addresses to process |
| `-CsvPath` | `string` | Path to a CSV file with an `Email` column |
| `-Account` | `switch` | Process all licensed users in the tenant that have a OneDrive (determined by SharePoint license) |
| `-AutoMap` | `switch` | Skip confirmation prompts and map all discovered sites automatically |

## Usage

### Single user (interactive prompts per site)
```powershell
.\CreateOnedriveShortcutSharepoint.ps1 -UserEmails user@domain.com
```

### Multiple users
```powershell
.\CreateOnedriveShortcutSharepoint.ps1 -UserEmails user1@domain.com, user2@domain.com
```

### From a CSV file
```powershell
.\CreateOnedriveShortcutSharepoint.ps1 -CsvPath .\users.csv
```

CSV format:
```csv
Email
user1@domain.com
user2@domain.com
```

### All licensed users with OneDrive
```powershell
.\CreateOnedriveShortcutSharepoint.ps1 -Account
```

### All licensed users, no prompts
```powershell
.\CreateOnedriveShortcutSharepoint.ps1 -Account -AutoMap
```

### No arguments (prompts for email)
```powershell
.\CreateOnedriveShortcutSharepoint.ps1
```

## Notes

- The script grants the signed-in admin temporary Site Collection Admin access to each user's OneDrive in order to create shortcuts
- Sites named "All Company" and redirect sites are excluded automatically
- If a shortcut already exists for a site, it is detected and skipped
- Guest/external users are excluded when using `-Account`
