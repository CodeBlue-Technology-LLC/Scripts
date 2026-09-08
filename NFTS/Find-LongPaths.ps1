<#
.SYNOPSIS
    Lists files and folders whose full path exceeds the classic Windows MAX_PATH limit.

.DESCRIPTION
    MUST be run under PowerShell 7 (pwsh.exe). .NET Core enumerates paths longer than
    260 characters transparently; Windows PowerShell 5.1 throws PathTooLongException
    and silently skips the very subtrees you are looking for.

.EXAMPLE
    pwsh -File .\Find-LongPaths.ps1

.EXAMPLE
    .\Find-LongPaths.ps1 -Path '\\server\share' | Export-Csv C:\tmp\longpaths.csv -NoTypeInformation
#>
[CmdletBinding()]
param(
    [string]$Path  = '\\192.168.16.14\elc\Shared',
    [int]   $Limit = 259   # MAX_PATH is 260 including the terminating null
)

if ($PSVersionTable.PSVersion.Major -lt 6) {
    throw "Run this under PowerShell 7 (pwsh.exe). Windows PowerShell 5.1 cannot enumerate paths over $Limit characters."
}

$errs   = @()
$result =
    Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue -ErrorVariable +errs |
        Where-Object { $_.FullName.Length -gt $Limit } |
        ForEach-Object {
            [pscustomobject]@{
                PathLength = $_.FullName.Length
                Type       = if ($_.PSIsContainer) { 'Folder' } else { 'File' }
                SizeMB     = if ($_.PSIsContainer) { $null } else { [math]::Round($_.Length / 1MB, 2) }
                FullName   = $_.FullName
            }
        } |
        Sort-Object PathLength -Descending

if ($errs.Count) {
    Write-Warning "$($errs.Count) path(s) could not be read (access denied, offline, etc.) - see `$errs in the caller scope."
}

$result
