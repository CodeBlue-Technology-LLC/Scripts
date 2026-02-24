<#
.SYNOPSIS
    Export NTFS permissions to an interactive HTML report (FAST TREE VIEW)
.DESCRIPTION
    Scans a directory and exports NTFS permissions in a nested tree structure
.PARAMETER Path
    The root path to scan for permissions
.PARAMETER OutputPath
    Where to save the HTML report (default: Desktop)
.PARAMETER MaxDepth
    Maximum folder depth to scan (default: 5 for safety)
.PARAMETER ExcludeSystemFolders
    Skip common system folders (Windows, ProgramData, etc.)
.PARAMETER TopLevelOnly
    Only scan immediate subfolders of root (FAST for large shares)
.EXAMPLE
    .\Export-NTFSPermissions.ps1 -Path "C:\Shares\Department" -MaxDepth 10
.EXAMPLE
    .\Export-NTFSPermissions.ps1 -Path "D:\FileShare" -TopLevelOnly
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$Path,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "$env:USERPROFILE\Desktop\NTFS_Permissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').html",
    
    [Parameter(Mandatory=$false)]
    [int]$MaxDepth = 5,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExcludeSystemFolders,
    
    [Parameter(Mandatory=$false)]
    [switch]$TopLevelOnly
)

# Validate path
if (!(Test-Path $Path)) {
    Write-Error "Path does not exist: $Path"
    exit 1
}

# Check if OutputPath is a directory, if so add filename
if (Test-Path $OutputPath -PathType Container) {
    $OutputPath = Join-Path $OutputPath "NTFS_Permissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    Write-Host "Output directory detected, using: $OutputPath" -ForegroundColor Cyan
}
elseif ((Test-Path (Split-Path $OutputPath -Parent)) -and -not $OutputPath.EndsWith('.html')) {
    # Path exists but no .html extension, add it
    $OutputPath = "$OutputPath.html"
    Write-Host "Added .html extension: $OutputPath" -ForegroundColor Cyan
}

# System folders to skip
$systemFolders = @(
    'Windows', 'Program Files', 'Program Files (x86)', 'ProgramData',
    '$Recycle.Bin', 'System Volume Information', '$WINDOWS.~BT',
    'Recovery', 'PerfLogs', 'MSOCache'
)

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "      NTFS Permissions Export - Fast Tree View                 " -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Scanning: $Path" -ForegroundColor Yellow
Write-Host "Max Depth: $(if($TopLevelOnly){'1 (Top Level Only)'}else{$MaxDepth})" -ForegroundColor Yellow
Write-Host "Exclude System: $ExcludeSystemFolders" -ForegroundColor Yellow
Write-Host ""

$script:folderCount = 0
$script:permCount = 0
$script:errorCount = 0
$script:startTime = Get-Date

# Simple flat structure - we'll build the tree in JavaScript
$allFolders = New-Object System.Collections.ArrayList

function Get-FolderPermissions {
    param(
        [string]$FolderPath,
        [string]$ParentPath = "",
        [int]$CurrentDepth = 0
    )
    
    # Check depth limit
    if ($TopLevelOnly -and $CurrentDepth -gt 1) { return }
    if (!$TopLevelOnly -and $MaxDepth -gt 0 -and $CurrentDepth -ge $MaxDepth) { return }
    
    # Skip system folders if requested
    if ($ExcludeSystemFolders) {
        $folderName = Split-Path $FolderPath -Leaf
        if ($folderName -in $systemFolders) {
            Write-Host "Skipping: $folderName" -ForegroundColor DarkGray
            return
        }
    }
    
    try {
        $script:folderCount++
        if ($script:folderCount % 10 -eq 0) {
            $msg = "Scanning: $script:folderCount folders | $script:permCount perms | $script:errorCount errors"
            Write-Host "`r$msg" -NoNewline -ForegroundColor Cyan
        }
        
        $acl = Get-Acl -Path $FolderPath -ErrorAction Stop
        
        $permissions = @()
        foreach ($access in $acl.Access) {
            $script:permCount++
            
            # Decode numeric rights to readable format
            $rightsValue = [int]$access.FileSystemRights
            $rightsDisplay = $access.FileSystemRights.ToString()
            
            # Common permission patterns
            $decoded = switch ($rightsValue) {
                268435456 { "FullControl (All)" }
                -1610612736 { "ReadAndExecute, Synchronize" }
                -536805376 { "Modify, Synchronize" }
                1179785 { "Read" }
                1180063 { "Read, Write" }
                1180095 { "Read, Write, Delete" }
                1245631 { "ReadAndExecute" }
                1179817 { "Read, WriteAttributes" }
                default { 
                    if ($rightsDisplay -match '^\d+$' -or $rightsDisplay -match '^-\d+$') {
                        # Still numeric, decode the flags
                        $flags = @()
                        if ($rightsValue -band 1) { $flags += "ReadData/ListDirectory" }
                        if ($rightsValue -band 2) { $flags += "WriteData/CreateFiles" }
                        if ($rightsValue -band 4) { $flags += "AppendData/CreateDirectories" }
                        if ($rightsValue -band 8) { $flags += "ReadExtendedAttributes" }
                        if ($rightsValue -band 16) { $flags += "WriteExtendedAttributes" }
                        if ($rightsValue -band 32) { $flags += "ExecuteFile/Traverse" }
                        if ($rightsValue -band 64) { $flags += "DeleteSubdirectoriesAndFiles" }
                        if ($rightsValue -band 128) { $flags += "ReadAttributes" }
                        if ($rightsValue -band 256) { $flags += "WriteAttributes" }
                        if ($rightsValue -band 65536) { $flags += "Delete" }
                        if ($rightsValue -band 131072) { $flags += "ReadPermissions" }
                        if ($rightsValue -band 262144) { $flags += "ChangePermissions" }
                        if ($rightsValue -band 524288) { $flags += "TakeOwnership" }
                        if ($rightsValue -band 1048576) { $flags += "Synchronize" }
                        if ($flags.Count -gt 0) { $flags -join ", " } else { $rightsDisplay }
                    } else {
                        $rightsDisplay
                    }
                }
            }
            
            $permissions += @{
                i = $access.IdentityReference.ToString()
                r = $rightsDisplay
                d = $decoded
                a = $access.AccessControlType.ToString()
                h = $access.IsInherited
                f = $access.InheritanceFlags.ToString()
            }
        }
        
        # Add to flat list
        $script:allFolders.Add(@{
            p = $FolderPath
            n = (Split-Path $FolderPath -Leaf)
            pr = $ParentPath
            d = $CurrentDepth
            pm = $permissions
        }) | Out-Null
        
        # Get subdirectories
        $subfolders = Get-ChildItem -Path $FolderPath -Directory -ErrorAction SilentlyContinue -Force
        foreach ($folder in $subfolders) {
            Get-FolderPermissions -FolderPath $folder.FullName -ParentPath $FolderPath -CurrentDepth ($CurrentDepth + 1)
        }
    }
    catch {
        $script:errorCount++
    }
}

# Scan the directory
Get-FolderPermissions -FolderPath $Path

$elapsed = (Get-Date) - $script:startTime
$elapsedMinutes = [math]::Floor($elapsed.TotalMinutes)
$elapsedSeconds = $elapsed.Seconds

Write-Host "`n"
Write-Host "Scan complete!" -ForegroundColor Green
Write-Host "   Folders: $script:folderCount" -ForegroundColor White
Write-Host "   Permissions: $script:permCount" -ForegroundColor White
Write-Host "   Errors (Access Denied): $script:errorCount" -ForegroundColor Yellow
Write-Host "   Time: $elapsedMinutes`:$elapsedSeconds" -ForegroundColor White
Write-Host ""
Write-Host "Generating HTML (this should be quick)..." -ForegroundColor Cyan

# Convert to JSON - much faster with flat structure
$jsonString = ($allFolders | ConvertTo-Json -Compress -Depth 5)

Write-Host "Writing HTML file..." -ForegroundColor Cyan

# Pre-calculate values for here-string
$reportDate = Get-Date -Format "yyyy-MM-dd HH:mm"
$rootPath = $Path

# Create HTML content as here-string
$htmlContent = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>NTFS Permissions - $rootPath</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: Segoe UI, Tahoma, Geneva, Verdana, sans-serif; background: #f5f5f5; }
.header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
h1 { font-size: 28px; margin-bottom: 10px; }
.stats { display: flex; gap: 30px; margin-top: 15px; flex-wrap: wrap; }
.stat { background: rgba(255,255,255,0.2); padding: 10px 15px; border-radius: 8px; }
.stat-label { font-size: 12px; opacity: 0.9; }
.stat-value { font-size: 20px; font-weight: bold; }
.controls { background: white; padding: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
.search-container { display: flex; gap: 10px; margin-bottom: 15px; flex-wrap: wrap; }
input[type="text"] { flex: 1; min-width: 250px; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px; }
input[type="text"]:focus { outline: none; border-color: #667eea; }
button { padding: 10px 20px; background: #667eea; color: white; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 600; transition: all 0.2s; }
button:hover { background: #5568d3; transform: translateY(-1px); }
.container { padding: 20px; max-width: 1800px; margin: 0 auto; }
.tree-container { background: white; border-radius: 8px; padding: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.08); }
.folder-node { margin: 2px 0; }
.folder-header { display: flex; align-items: center; padding: 8px 10px; cursor: pointer; border-radius: 4px; transition: background 0.15s; }
.folder-header:hover { background: #f0f0f0; }
.expand-icon { width: 18px; height: 18px; display: inline-flex; align-items: center; justify-content: center; margin-right: 6px; font-size: 13px; font-weight: bold; color: #667eea; user-select: none; flex-shrink: 0; }
.expand-icon.has-children.collapsed::before { content: "[+]"; }
.expand-icon.has-children.expanded::before { content: "[-]"; }
.expand-icon.no-children { visibility: hidden; }
.folder-name { flex: 1; font-weight: 500; color: #2c3e50; font-size: 14px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.perm-count { background: #e3e8ff; color: #667eea; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; margin-left: 8px; flex-shrink: 0; }
.perm-button { background: #667eea; color: white; padding: 4px 10px; border-radius: 4px; font-size: 11px; border: none; cursor: pointer; margin-left: 8px; flex-shrink: 0; }
.perm-button:hover { background: #5568d3; }
.children-container { margin-left: 24px; display: none; }
.children-container.expanded { display: block; }
.permissions-panel { margin: 6px 0 8px 44px; background: #f8f9fa; border-left: 3px solid #667eea; border-radius: 4px; display: none; overflow-x: auto; }
.permissions-panel.visible { display: block; }
table { width: 100%; border-collapse: collapse; font-size: 12px; }
th { background: #e9ecef; padding: 8px; text-align: left; font-weight: 600; font-size: 11px; color: #495057; white-space: nowrap; }
td { padding: 8px; border-bottom: 1px solid #dee2e6; }
tr:last-child td { border-bottom: none; }
tr:hover { background: #fff; }
.rights-cell { max-width: 400px; }
.rights-raw { font-family: monospace; color: #6c757d; font-size: 10px; display: block; }
.rights-decoded { color: #495057; font-weight: 500; }
.allow { color: #28a745; font-weight: 600; }
.deny { color: #dc3545; font-weight: 600; }
.inherited { color: #6c757d; }
.loading { text-align: center; padding: 40px; color: #6c757d; }
.path-hint { font-size: 10px; color: #adb5bd; margin-left: 44px; font-family: monospace; margin-top: -4px; margin-bottom: 4px; }
</style>
</head>
<body>
<div class="header">
<h1>NTFS Permissions Report - Tree View</h1>
<div class="stats">
<div class="stat"><div class="stat-label">Root Path</div><div class="stat-value" style="font-size: 14px;">$rootPath</div></div>
<div class="stat"><div class="stat-label">Folders</div><div class="stat-value">$($script:folderCount)</div></div>
<div class="stat"><div class="stat-label">Permissions</div><div class="stat-value">$($script:permCount)</div></div>
<div class="stat"><div class="stat-label">Generated</div><div class="stat-value" style="font-size: 14px;">$reportDate</div></div>
</div></div>
<div class="controls">
<div class="search-container">
<input type="text" id="searchBox" placeholder="Search folder names or paths...">
<button onclick="expandAll()">Expand All</button>
<button onclick="collapseAll()">Collapse All</button>
<button onclick="expandToLevel(1)">Level 1</button>
<button onclick="expandToLevel(2)">Level 2</button>
</div>
</div>
<div class="container">
<div id="loading" class="loading">Building tree structure...</div>
<div class="tree-container" id="treeContainer" style="display:none;"></div>
</div>
<script>
const flatData = $jsonString;
console.log("Loaded " + flatData.length + " folders");
const folderMap = new Map();
flatData.forEach(f => { f.children = []; folderMap.set(f.p, f); });
flatData.forEach(f => { if (f.pr && folderMap.has(f.pr)) { folderMap.get(f.pr).children.push(f); } });
const rootFolder = flatData.find(f => f.pr === "" || f.pr === f.p);
function createNode(folder, parentEl) {
const nodeDiv = document.createElement("div");
nodeDiv.className = "folder-node";
nodeDiv.dataset.path = folder.p;
nodeDiv.dataset.depth = folder.d;
const hasKids = folder.children.length > 0;
const header = document.createElement("div");
header.className = "folder-header";
const icon = document.createElement("span");
icon.className = hasKids ? "expand-icon has-children collapsed" : "expand-icon no-children";
const name = document.createElement("span");
name.className = "folder-name";
name.textContent = folder.n || folder.p;
name.title = folder.p;
const count = document.createElement("span");
count.className = "perm-count";
count.textContent = folder.pm.length;
const btn = document.createElement("button");
btn.className = "perm-button";
btn.textContent = "Perms";
btn.onclick = e => { e.stopPropagation(); togglePerms(nodeDiv); };
header.appendChild(icon);
header.appendChild(name);
header.appendChild(count);
header.appendChild(btn);
if (hasKids) header.onclick = () => toggleKids(nodeDiv, icon);
const panel = document.createElement("div");
panel.className = "permissions-panel";
let html = "<table><thead><tr><th>Identity</th><th>Rights</th><th>Access</th><th>Inherited</th><th>Flags</th></tr></thead><tbody>";
folder.pm.forEach(p => {
const ac = p.a === "Allow" ? "allow" : "deny";
const ih = p.h ? "inherited" : "";
const rightsDisplay = p.d && p.d !== p.r ? "<span class='rights-decoded'>" + p.d + "</span><span class='rights-raw'>" + p.r + "</span>" : p.r;
html += "<tr><td class='" + ih + "'>" + p.i + "</td><td class='rights-cell'>" + rightsDisplay + "</td><td class='" + ac + "'>" + p.a + "</td><td>" + p.h + "</td><td>" + p.f + "</td></tr>";
});
html += "</tbody></table>";
panel.innerHTML = html;
nodeDiv.appendChild(header);
nodeDiv.appendChild(panel);
if (hasKids) {
const kidsCont = document.createElement("div");
kidsCont.className = "children-container";
folder.children.forEach(kid => createNode(kid, kidsCont));
nodeDiv.appendChild(kidsCont);
}
parentEl.appendChild(nodeDiv);
}
function toggleKids(node, icon) {
const cont = node.querySelector(".children-container");
if (!cont) return;
const isExp = cont.classList.toggle("expanded");
icon.classList.toggle("collapsed", !isExp);
icon.classList.toggle("expanded", isExp);
}
function togglePerms(node) { node.querySelector(".permissions-panel").classList.toggle("visible"); }
function expandAll() { document.querySelectorAll(".children-container").forEach(e => e.classList.add("expanded")); document.querySelectorAll(".expand-icon.has-children").forEach(e => { e.classList.remove("collapsed"); e.classList.add("expanded"); }); }
function collapseAll() { document.querySelectorAll(".children-container").forEach(e => e.classList.remove("expanded")); document.querySelectorAll(".expand-icon.has-children").forEach(e => { e.classList.add("collapsed"); e.classList.remove("expanded"); }); document.querySelectorAll(".permissions-panel").forEach(e => e.classList.remove("visible")); }
function expandToLevel(lvl) { collapseAll(); document.querySelectorAll(".folder-node").forEach(n => { const d = parseInt(n.dataset.depth); if (d < lvl) { const cont = n.querySelector(".children-container"); const icon = n.querySelector(".expand-icon.has-children"); if (cont && icon) { cont.classList.add("expanded"); icon.classList.remove("collapsed"); icon.classList.add("expanded"); } } }); }
document.getElementById("searchBox").addEventListener("input", e => {
const term = e.target.value.toLowerCase();
document.querySelectorAll(".folder-node").forEach(n => {
const txt = n.textContent.toLowerCase();
const path = n.dataset.path.toLowerCase();
n.style.display = (term === "" || txt.includes(term) || path.includes(term)) ? "block" : "none";
});
});
document.getElementById("loading").textContent = "Rendering tree...";
setTimeout(() => {
createNode(rootFolder, document.getElementById("treeContainer"));
document.getElementById("loading").style.display = "none";
document.getElementById("treeContainer").style.display = "block";
console.log("Tree rendered!");
}, 100);
</script>
</body>
</html>
"@

$htmlContent | Out-File -FilePath $OutputPath -Encoding UTF8

Write-Host "Report saved: $OutputPath" -ForegroundColor Green
Write-Host "File size: $([math]::Round((Get-Item $OutputPath).Length / 1MB, 2)) MB" -ForegroundColor White
Write-Host ""
Write-Host "Opening in browser..." -ForegroundColor Cyan

Start-Process $OutputPath

Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "                    COMPLETE!                                   " -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
