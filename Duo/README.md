# Duo MSP Scripts

PowerShell scripts for managing Duo Security subaccounts as an MSP, built by Code Blue Technology.

## Scripts

### ProvisionDuoAccount.ps1

Provisions new Duo subaccounts for ConnectWise Manage companies. The script searches ConnectWise for the company, creates a Duo subaccount with the matching name, and applies standard settings (timezone, universal prompt, helpdesk enrollment email).

```powershell
# Interactive - prompts for company name
.\ProvisionDuoAccount.ps1

# Specify company name directly
.\ProvisionDuoAccount.ps1 -CompanyName "Contoso Ltd"

# Skip automatic settings configuration
.\ProvisionDuoAccount.ps1 -CompanyName "Contoso Ltd" -SkipSettings

# Re-enter stored credentials
.\ProvisionDuoAccount.ps1 -ResetCredentials
```

### Get-DuoBypassUsers.ps1

Reports all Duo users in bypass mode across every subclient. Results are displayed grouped by subclient in the console and exported to a timestamped CSV. Optionally creates ConnectWise tickets — either one per client or all under a single company.

```powershell
# Run with defaults (CSV saved to script directory)
.\Get-DuoBypassUsers.ps1

# Specify CSV output path
.\Get-DuoBypassUsers.ps1 -OutputPath "C:\Reports\bypass.csv"

# Re-enter stored credentials
.\Get-DuoBypassUsers.ps1 -ResetCredentials
```

## Prerequisites

- PowerShell 5.1+
- **DuoSecurity** module (auto-installed if missing)
- **ConnectWiseManageAPI** module (auto-installed if missing; only required for provisioning and ticket creation)

## Credentials

Both scripts share a single credential file at `Config\credentials.xml`, encrypted with DPAPI (machine- and user-bound). On first run you'll be prompted for:

| Key | Purpose |
|-----|---------|
| Duo API Host | Accounts API hostname (MSP level) |
| Duo Integration Key | Accounts API integration key |
| Duo Secret Key | Accounts API secret key |
| ConnectWise Server | e.g. `na.myconnectwise.net` |
| ConnectWise Company ID | CWM company identifier |
| ConnectWise Client ID | API client ID |
| ConnectWise Public Key | API public key |
| ConnectWise Private Key | API private key |

Use the `-ResetCredentials` switch on either script to re-enter credentials. Each script only overwrites its own section — Duo-only resets from `Get-DuoBypassUsers.ps1` preserve ConnectWise keys and vice versa.
