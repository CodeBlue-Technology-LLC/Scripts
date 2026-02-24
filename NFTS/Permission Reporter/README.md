# NTFS Permissions Exporter

A PowerShell script that exports NTFS permissions to a beautiful, interactive HTML report with a tree view structure. Perfect for auditing file share permissions, security reviews, and documentation.

![NTFS Permissions Report](https://img.shields.io/badge/PowerShell-5.1%2B-blue) ![License](https://img.shields.io/badge/license-MIT-green)

## Features

✨ **Interactive Tree View** - Collapsible folder hierarchy for easy navigation  
🔍 **Live Search** - Filter by folder names, paths, or user accounts  
📊 **Permission Decoding** - Translates cryptic numeric permissions to human-readable text  
🎨 **Clean UI** - Modern, responsive design that works in any browser  
⚡ **Fast Performance** - Handles tens of thousands of folders efficiently  
🎯 **Smart Output** - Auto-detects directories and generates timestamped files  

## Screenshots

### Tree View
Navigate through nested folder structures with expandable/collapsible nodes:

```
[+] AppData
    [+] Local
        [+] Microsoft
            [-] Windows
                • 13 permissions
```

### Permission Details
See decoded permissions with color-coded access types:
- **Green** = Allow
- **Red** = Deny
- *Gray* = Inherited

Numeric permissions like `268435456` are automatically decoded to **"FullControl (All)"**

## Requirements

- Windows PowerShell 5.1 or later
- Appropriate permissions to read ACLs on target directories
- Any modern web browser to view the report

## Installation

1. Download `NTFSPermissionsExporter.ps1`
2. (Optional) Unblock the script:
   ```powershell
   Unblock-File -Path .\NTFSPermissionsExporter.ps1
   ```

## Usage

### Basic Usage

```powershell
.\NTFSPermissionsExporter.ps1 -Path "C:\Shares\Department"
```

This creates a report on your Desktop with a timestamp: `NTFS_Permissions_20260126_143022.html`

### Specify Output Location

```powershell
# Save to specific directory (auto-generates timestamped filename)
.\NTFSPermissionsExporter.ps1 -Path "D:\FileShare" -OutputPath "C:\Reports"

# Save with specific filename
.\NTFSPermissionsExporter.ps1 -Path "D:\FileShare" -OutputPath "C:\Reports\ShareAudit.html"
```

### Control Scan Depth

```powershell
# Limit to 3 levels deep
.\NTFSPermissionsExporter.ps1 -Path "C:\Data" -MaxDepth 3

# Only scan immediate subfolders (fastest)
.\NTFSPermissionsExporter.ps1 -Path "C:\Shares" -TopLevelOnly
```

### Skip System Folders

```powershell
# Automatically skip Windows, Program Files, etc.
.\NTFSPermissionsExporter.ps1 -Path "C:\" -ExcludeSystemFolders -MaxDepth 2
```

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Path` | String | *Required* | Root directory to scan |
| `-OutputPath` | String | Desktop | Where to save the HTML report |
| `-MaxDepth` | Integer | 5 | Maximum folder depth to scan (prevents runaway scans) |
| `-ExcludeSystemFolders` | Switch | False | Skip common system folders |
| `-TopLevelOnly` | Switch | False | Only scan immediate subfolders (depth = 1) |

## Examples

### Audit a File Share
```powershell
.\NTFSPermissionsExporter.ps1 -Path "\\fileserver\departments" -OutputPath "C:\Audits" -MaxDepth 10
```

### Quick Department Scan
```powershell
.\NTFSPermissionsExporter.ps1 -Path "C:\Shares\HR" -OutputPath ".\hr_permissions.html"
```

### Scan C:\ Drive (Safe)
```powershell
.\NTFSPermissionsExporter.ps1 -Path "C:\" -ExcludeSystemFolders -MaxDepth 3 -OutputPath "C:\Temp"
```

### Large Share with Top-Level Overview
```powershell
.\NTFSPermissionsExporter.ps1 -Path "D:\CompanyData" -TopLevelOnly
```

## Understanding the Report

### Tree Navigation
- **[+]** - Click to expand folder and show subfolders
- **[-]** - Click to collapse folder
- **Perms** button - Show/hide permission details for that folder
- **Number badge** - Shows count of permission entries

### Control Buttons
- **Expand All** - Opens all folders in the tree
- **Collapse All** - Closes all folders and permission panels
- **Level 1** - Expands only the first level
- **Level 2** - Expands first two levels

### Permission Decoding

The script automatically translates numeric permissions:

| Raw Value | Decoded Meaning |
|-----------|-----------------|
| `268435456` | FullControl (All) |
| `-1610612736` | ReadAndExecute, Synchronize |
| `-536805376` | Modify, Synchronize |
| `1179785` | Read |
| `1245631` | ReadAndExecute |

Complex numeric permissions are broken down into individual rights like:
- ReadData/ListDirectory
- WriteData/CreateFiles
- ExecuteFile/Traverse
- Delete
- ChangePermissions
- TakeOwnership

## Performance Tips

### For Large Directory Structures (100,000+ files)

1. **Limit depth**: Use `-MaxDepth 5` or lower
2. **Use TopLevelOnly**: Great for getting an overview of major folders
3. **Exclude system folders**: Saves significant time on system drives
4. **Target specific paths**: Scan only what you need

### Expected Performance

| Folders Scanned | Typical Time | File Size |
|-----------------|--------------|-----------|
| 1,000 | 10-30 seconds | 2-5 MB |
| 10,000 | 1-3 minutes | 15-30 MB |
| 50,000+ | 5-15 minutes | 50-100 MB |

*Times vary based on disk speed, network latency (for shares), and folder depth*

## Troubleshooting

### "Access Denied" Errors
- These are normal and counted in the error statistics
- Run PowerShell as Administrator for system folders
- You'll see the count in the final summary

### Script Hangs on Large Directories
- Press `Ctrl+C` to cancel
- Reduce `-MaxDepth` or use `-TopLevelOnly`
- Consider scanning subdirectories individually

### HTML File Won't Open
- Check that the file has `.html` extension
- Try a different browser
- Verify the file isn't corrupted (check file size > 0)

### Numeric Permissions Not Decoded
- Some complex/custom permissions may still show as numbers
- The raw value is always displayed for reference
- Common patterns are automatically decoded

## Use Cases

- 📋 **Security Audits** - Review who has access to sensitive data
- 📝 **Documentation** - Create snapshots of permission structures
- 🔄 **Migration Planning** - Document current state before moving data
- 👥 **Access Reviews** - Identify overly permissive or inherited permissions
- 🔍 **Troubleshooting** - Diagnose permission issues on file shares
- 📊 **Compliance** - Generate evidence for audit requirements

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## License

MIT License - feel free to use and modify for your needs.

## Credits

Created for system administrators who need a better way to visualize and audit NTFS permissions.

---

**Pro Tip**: Bookmark the generated HTML files - they're completely self-contained and can be shared with others or archived for historical comparison!
