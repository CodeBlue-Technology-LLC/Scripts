# CWM-SyncVEEAMSERVERBackupTB

Syncs Veeam Backup & Replication storage usage to ConnectWise Manage agreement additions.

## What it does

1. Connects to a Veeam B&R server and retrieves all backup jobs
2. Calculates the total storage used (in TB) for each backup with VMs
3. Matches backups to active ConnectWise Manage agreements by company name
4. Creates or updates the specified agreement addition with the current TB usage
5. Sets the effective date to the 1st of next month

## Prerequisites

- **Veeam.Backup.PowerShell** module (available on the Veeam B&R server)
- **ConnectWiseManageAPI** module (installed automatically if missing)
- PowerShell 5.1 or later

## Setup

On first run, the script prompts for:

- Veeam B&R server name
- Veeam server credentials (username and password)
- ConnectWise Manage connection details (server URL, company, public key, private key, client ID)
- CWM product ID for the Veeam storage addition

These are saved to `Config/Credentials.xml` using `Export-Clixml`. Credentials are encrypted with DPAPI and can only be decrypted by the same user on the same machine.

To reconfigure, delete `Config/Credentials.xml` and run the script again.

## Usage

```powershell
.\CWM-SyncVEEAMSERVERBackupTB.ps1
```
