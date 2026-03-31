#Requires -Modules ActiveDirectory
<#
.SYNOPSIS
    Audits all AD users and groups, exports to XLSX (or CSV fallback).
.NOTES
    XLSX output requires the ImportExcel module:
    Install-Module ImportExcel -Scope CurrentUser
#>

[CmdletBinding()]
param(
    [string]$OutputPath = "$env:USERPROFILE\Desktop",
    [string]$Server     = $env:USERDNSDOMAIN   # Override with specific DC if needed
)

$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$xlsxFile   = Join-Path $OutputPath "AD_Audit_$timestamp.xlsx"
$csvUsers   = Join-Path $OutputPath "AD_Users_$timestamp.csv"
$csvGroups  = Join-Path $OutputPath "AD_Groups_$timestamp.csv"

$hasExcel = Get-Module -ListAvailable -Name ImportExcel

# ── USERS ────────────────────────────────────────────────────────────────────
Write-Host "[*] Pulling AD users..." -ForegroundColor Cyan

$users = Get-ADUser -Filter * -Server $Server -Properties `
    DisplayName, SamAccountName, UserPrincipalName, Enabled,
    LastLogonDate, PasswordLastSet, PasswordNeverExpires,
    PasswordNotRequired, LockedOut, DistinguishedName |
Select-Object @{N="DisplayName";         E={ $_.DisplayName }},
              @{N="Username";             E={ $_.SamAccountName }},
              @{N="UPN";                  E={ $_.UserPrincipalName }},
              @{N="Enabled";              E={ $_.Enabled }},
              @{N="Locked Out";           E={ $_.LockedOut }},
              @{N="Last Logon";           E={ if ($_.LastLogonDate) { $_.LastLogonDate } else { "Never" } }},
              @{N="Password Last Set";    E={ if ($_.PasswordLastSet) { $_.PasswordLastSet } else { "Never" } }},
              @{N="Password Age (Days)";  E={
                  if ($_.PasswordLastSet) {
                      (New-TimeSpan -Start $_.PasswordLastSet -End (Get-Date)).Days
                  } else { "N/A" }
              }},
              @{N="Pwd Never Expires";    E={ $_.PasswordNeverExpires }},
              @{N="Pwd Not Required";     E={ $_.PasswordNotRequired }},
              @{N="OU";                   E={ ($_.DistinguishedName -replace '^CN=[^,]+,','') }}

Write-Host "[+] Found $($users.Count) users." -ForegroundColor Green

# ── GROUPS ───────────────────────────────────────────────────────────────────
Write-Host "[*] Pulling AD groups..." -ForegroundColor Cyan

$rawGroups = Get-ADGroup -Filter * -Server $Server -Properties Members, Description, GroupScope, GroupCategory

$groups = foreach ($g in $rawGroups) {
    $members = if ($g.Members.Count -gt 0) {
        ($g.Members | ForEach-Object {
            try { (Get-ADObject $_ -Server $Server).Name } catch { $_ }
        }) -join "; "
    } else { "(empty)" }

    [PSCustomObject]@{
        "Group Name"     = $g.Name
        "Scope"          = $g.GroupScope
        "Category"       = $g.GroupCategory
        "Description"    = $g.Description
        "Member Count"   = $g.Members.Count
        "Members"        = $members
    }
}

Write-Host "[+] Found $($groups.Count) groups." -ForegroundColor Green

# ── EXPORT ───────────────────────────────────────────────────────────────────
if ($hasExcel) {
    Write-Host "[*] Exporting to XLSX: $xlsxFile" -ForegroundColor Cyan

    $users  | Export-Excel -Path $xlsxFile -WorksheetName "Users"  -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    $groups | Export-Excel -Path $xlsxFile -WorksheetName "Groups" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

    Write-Host "[+] Done → $xlsxFile" -ForegroundColor Green
} else {
    Write-Warning "ImportExcel not found. Falling back to CSV."
    $users  | Export-CSV -Path $csvUsers  -NoTypeInformation -Encoding UTF8
    $groups | Export-CSV -Path $csvGroups -NoTypeInformation -Encoding UTF8
    Write-Host "[+] Done → $csvUsers | $csvGroups" -ForegroundColor Green
}
