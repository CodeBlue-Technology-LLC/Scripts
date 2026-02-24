<#
.SYNOPSIS
    Pre-provisions OneDrive for all licensed users in a Microsoft 365 tenant.

.DESCRIPTION
    Connects to Microsoft Graph and SharePoint Online, retrieves all licensed users,
    and pre-provisions their OneDrive in batches of 199 (the cmdlet limit).

.PARAMETER SharepointURL
    The SharePoint Online admin center URL (e.g., https://contoso-admin.sharepoint.com)

.PARAMETER TenantID
    The Microsoft 365 Tenant ID.

.EXAMPLE
    .\Pre-Provision-OneDrive.ps1 -SharepointURL "https://contoso-admin.sharepoint.com" -TenantID "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

.NOTES
    Requirements:
    - Microsoft.Graph PowerShell module (Install-Module Microsoft.Graph.Users)
    - SharePoint Online Management Shell (Install-Module Microsoft.Online.SharePoint.PowerShell)
    - Must be a SharePoint Administrator with a SharePoint license assigned
    - Users must be allowed to sign in and have a SharePoint license
#>

Param(
    [Parameter(Mandatory = $True)]
    [String]$SharepointURL,

    [Parameter(Mandatory = $True)]
    [String]$TenantID
)

# Connect to Microsoft Graph and SharePoint Online
Connect-MgGraph -TenantId $TenantID -Scopes 'User.Read.All'
Connect-SPOService -Url $SharepointURL

$list = @()
$TotalUsers = 0

# Get all licensed users
$users = Get-MgUser -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -CountVariable licensedUserCount -All -Select UserPrincipalName

foreach ($u in $users) {
    $TotalUsers++
    Write-Host "Queued: $($u.UserPrincipalName)"
    $list += $u.UserPrincipalName

    if ($list.Count -eq 199) {
        Write-Host "Batch limit reached, requesting provision for the current batch"
        Request-SPOPersonalSite -UserEmails $list -NoWait
        Start-Sleep -Milliseconds 655
        $list = @()
    }
}

if ($list.Count -gt 0) {
    Request-SPOPersonalSite -UserEmails $list -NoWait
}

Disconnect-SPOService
Disconnect-MgGraph
Write-Host "Completed OneDrive provisioning for $TotalUsers users"
