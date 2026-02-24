# Cloudflare DNS Migration Tool - Implementation Plan

## SAFETY REQUIREMENT
**NO CHANGES to Cloudflare or GoDaddy without explicit user approval.**
All write operations will prompt for confirmation before execution.

## Workflow Order
1. **Query GoDaddy** for DNS records (read-only, safe)
2. **Create Cloudflare subtenant** (account) ← Requires approval
3. **Add zone to subtenant** ← Requires approval (captures nameservers)
4. **Preview transfer** - Show what will be created
5. **Transfer records** to Cloudflare ← Requires approval
6. **Update GoDaddy nameservers** to Cloudflare ← Requires approval
7. **Unlock domain and retrieve auth code** for registrar transfer (optional) ← Requires approval
8. **Create ConnectWise ticket** with auth code for transfer follow-up (automatic if domain unlocked)
9. **Create ITGlue DNS/Registrar asset** for documentation (automatic)

## Project Structure
```
Cloudflare/
├── Migrate-DNS.ps1                # Main migration script
├── Modules/
│   ├── Cloudflare.psm1            # Cloudflare API (accounts, zones, DNS)
│   └── GoDaddy.psm1               # GoDaddy API (read-only queries)
└── Config/
    └── credentials.xml            # Encrypted credentials (Export-Clixml)
```

## Implementation Steps

### Step 1: Credential Setup (Config/credentials.xml)
Using PowerShell's `Export-Clixml` pattern (matching Duo project).
**Credentials provided** - will be stored encrypted during implementation.

```powershell
# Structure (values stored securely, not in plan):
@{
    Cloudflare = @{
        Email = "your-email@example.com"
        ApiKey = "[YOUR-KEY]"
    }
    GoDaddy = @{
        ApiKey = "[YOUR-KEY]"
        ApiSecret = "[YOUR-SECRET]"
    }
}
```

### Step 2: GoDaddy Module (Modules/GoDaddy.psm1)
Functions:
- `Get-GoDaddyDomains` - List all domains in account (read-only)
- `Get-GoDaddyDNSRecords -Domain "example.com"` - Get DNS records (read-only)
- `Set-GoDaddyNameservers -Domain "example.com" -NameServers @(...)` - Update nameservers (requires -Confirm)
- `Unlock-GoDaddyDomain -Domain "example.com"` - Unlock for transfer (requires -Confirm)
- `Get-GoDaddyAuthCode -Domain "example.com"` - Retrieve EPP/auth code for transfer out

API: `https://api.godaddy.com/v1/domains/{domain}`
Auth: `Authorization: sso-key {key}:{secret}`

### Step 3: Cloudflare Module (Modules/Cloudflare.psm1)
All write functions include `-Confirm` parameter:
- `New-CloudflareAccount -Name "Customer"` - Create subtenant
- `Get-CloudflareAccounts` - List accounts (read-only)
- `New-CloudflareZone -AccountId X -Domain "example.com"` - Add zone
- `Get-CloudflareZones` - List zones (read-only)
- `New-CloudflareDNSRecord` - Create single DNS record
- `Import-CloudflareDNS -Records $records` - Bulk import with preview

API Base: `https://api.cloudflare.com/client/v4/`
Auth: `X-Auth-Email` + `X-Auth-Key` headers

### Step 4: Migration Script (Migrate-DNS.ps1)
Interactive workflow with confirmations:
```powershell
# Full migration with prompts at each step
.\Migrate-DNS.ps1 -Domain "example.com" -CustomerName "Acme Corp"

# Safe read-only operations (no confirmation needed)
.\Migrate-DNS.ps1 -ListGoDaddyDomains
.\Migrate-DNS.ps1 -PreviewRecords "example.com"
.\Migrate-DNS.ps1 -ListCloudflareAccounts
```

## API Endpoints

### Cloudflare (Tenant Admin)
| Operation | Method | Endpoint |
|-----------|--------|----------|
| Create Account | POST | `/accounts` |
| List Accounts | GET | `/accounts` |
| Create Zone | POST | `/zones` |
| List Zones | GET | `/zones` |
| Create DNS Record | POST | `/zones/{zone_id}/dns_records` |

### GoDaddy
| Operation | Method | Endpoint |
|-----------|--------|----------|
| List Domains | GET | `/v1/domains` |
| Get DNS Records | GET | `/v1/domains/{domain}/records` |
| Update Nameservers | PATCH | `/v1/domains/{domain}` |
| Unlock Domain | PATCH | `/v1/domains/{domain}` (set `locked: false`) |
| Get Auth Code | GET | `/v1/domains/{domain}` (returns `authCode` field) |

## Files to Create
1. `Cloudflare/Config/credentials.xml` - Encrypted credentials
2. `Cloudflare/Modules/GoDaddy.psm1` - GoDaddy read-only module
3. `Cloudflare/Modules/Cloudflare.psm1` - Cloudflare API module
4. `Cloudflare/Migrate-DNS.ps1` - Main migration script

## Verification
1. Test GoDaddy connection: `.\Migrate-DNS.ps1 -ListGoDaddyDomains`
2. Test Cloudflare connection: `.\Migrate-DNS.ps1 -ListCloudflareAccounts`
3. Preview records: `.\Migrate-DNS.ps1 -PreviewRecords "example.com"`
4. Run full migration with user approval at each step
