# IT Glue Services Dashboard - Follow-up Notes (for Claude)

Working notes to resume this project cold. Read this first, then skim
`New-ITGlueServicesDashboard.ps1`. Full design rationale is in the approved plan:
`C:\Users\akolhoff.EPIPHANYPLX\.claude\plans\i-would-like-to-expressive-karp.md`.

## What this is
A per-client gavsto-style "All Services" Bootstrap-panel dashboard for IT Glue. First build
renders a **local HTML preview** from **live, read-only** vendor data. Go-live later = push the
same HTML into an IT Glue flexible asset (see "Go-live" below).

## Files
- `IT Glue/New-ITGlueServicesDashboard.ps1` - main script (committed, 9aae81a).
- `IT Glue/lib/ITGlue-BootStrapHelpers.ps1` - vendored gavsto panel functions, GPL-3.0 (committed).
- `IT Glue/Config/dashboard-credentials.xml` - THIS script's OWN secrets, git-ignored. Seeded from
  the shared credentials.xml on first run, but never written to it. (Shared `credentials.xml` is
  owned by Update-ITGlueCompanyInfo.ps1: ITGlue + LogoDev keys.)
- `IT Glue/Config/service-map.json` - auto-built client->vendor-id cache, git-ignored.
- `IT Glue/Output/*.html` - generated dashboards, git-ignored (Output added to .gitignore).
- Real path on this box is `c:\cbt\Scripts\github\Scripts\IT Glue` (not `...\Scripts\Scripts\...`).

## Current status (as of 2026-06-15)
- Both scripts parse clean (PS 5.1), pure ASCII. Read-only verified (no ITGlue writes; only auth +
  JSON-RPC read POSTs to vendors).
- Prereqs: `ITGlueAPI` v2.2.0 present. **`DuoSecurity` is NOT actually installed** despite the old
  note - the script auto-installs it on first Duo-configured run.
- **IT Glue auth: WORKING.** Verified 2026-06-15 with the (rotated) stored key. Subdomain
  `codebluetechnology`, US region.
- **Dedicated creds: DONE.** Script now uses its own `Config\dashboard-credentials.xml`, seeded
  from the shared file's ITGlue section. Shared `credentials.xml` confirmed untouched.
- **Full IT Glue-only path: DRY-RUN VERIFIED** end-to-end against "Christine S. Rausch, MD, PC"
  (org 9044976). Renders valid Bootstrap-3 HTML; real Domains (3) + M365 Licenses (5) tiles.
- **M365 Licenses tile: VERIFIED + FIXED.** Real trait keys are `license-name`, `active`,
  `consumed`, `unused`, and counts are DECIMALS ("70.0"). Old `^\d+$` probe never matched -> now
  accepts decimals, casts to int, and skips `unused`. Renders e.g. "M365 Business Premium - 69/70".
- **Vendor tiles: STILL NOT tested against live APIs** (need interactive secret entry; can't be
  driven from the agent's non-interactive shell). Next: user runs interactively, one vendor at a
  time. SentinelOne/Duo first (proven), then the unverified Bitdefender/Veeam/Cove.

## New tiles added 2026-06-15 (verified vs Christine Rausch)
- Page title/H1 renamed to **"<OrgName> - PeopleFirst Services"** (was "All <OrgName> Services").
- **Users tile** (info panel, first tile): two linked lines - "IT Glue: <n>" (org contact count; IT
  Glue contacts have NO active/inactive flag so it's the total) linking to `<linkBase>/contacts`,
  and "ConnectWise: <n>" (ACTIVE contacts, inactiveFlag=false) linking to the CWM company record
  (`/v4_6_release/services/system_io/router/openrecord.rails?recordType=CompanyFV&recordId=<id>`).
- **Workstations / Servers tiles**: IT Glue configuration counts filtered by type ("Managed
  Workstation" / "Managed Server") + status "Active", via meta total-count. Type/status ids resolved
  by name at runtime. Linked to `<linkBase>/configurations` (IT Glue has no reliable URL filter, so
  the link is the full configs list, not pre-filtered).
- **ConnectWise Manage**: creds SEEDED from any sibling repo `*/Config/Credentials.xml` that has a
  `CWMConnectionInfo` block (Server connect.codebluetechnology.com). Stored in our own
  dashboard-credentials.xml; never written back. Company matched by normalized name via
  Resolve-OrgServiceId, cached under service-map key 'connectwise'. IMPORTANT: must filter
  `deletedFlag=false` - there are duplicate same-named CW companies (deleted/prospect) with 0
  contacts; the query uses deletedFlag=false. Module: ConnectWiseManageAPI v0.4.16 (runs under 5.1).

## Refinements 2026-06-15 (round 2) - all verified live
- **Tile model = CATEGORY tiles** (user decision). Each category lists EVERY active product as a
  "Product: value" line via New-CategoryTile (info panel; green if any present, red 'None' if not):
  - **Antivirus**: SentinelOne and/or Bitdefender (both resolve independently now - a client can have
    both, e.g. Christine shows "SentinelOne: 55" + "Bitdefender: 2").
  - **MFA**: Duo (renamed from the old 'Duo' tile).
  - **Backup**: Veeam ("Veeam: N srv, M wks, K VM") + Cove endpoint ("Cove: T TB (...)") + Cove M365
    ("Cove 365: T TB (n tenants)"). The standalone **M365 Backup tile was REMOVED** - Cove 365 lives
    in Backup, explicitly labelled "365" so it is not mistaken for server/workstation backup.
- **M365 Licenses hide rule**: skip SKUs whose seat total (denominator) is <= 0 or >= 1000 - hides
  free/unlimited/empty Microsoft SKUs (Windows Store /1000000, Power Automate /10000, Teams
  Exploratory & Exchange Online Essentials 0/0). Real paid licenses (e.g. 69/70) remain.
- **Workstation/Server links** now deep-link with the IT Glue UI filter:
  `.../configurations#partial=&sortBy=name:asc&filters=%5BType:Managed%20Workstation%5D`.
- **ConnectWise contacts link**: param is **`recid`** (NOT `recordId` - that opened a blank "New
  company"). CW Manage has NO deep-link for a company's Contacts TAB (that state is an opaque
  LZMA-compressed URL fragment, not templatable). So the Users tile's ConnectWise line links to the
  company's BILLING CONTACT record: `openrecord.rails?recordType=ContactFV&recid=<billingContact.id>`
  (from `Get-CWMCompany -id <co>`.billingContact.id), falling back to `CompanyFV&recid=<id>` if the
  company has no billing contact.
- Singles row order: Users, Workstations, Servers, Antivirus, MFA, Backup. Info row: M365 Licenses,
  Domains.

## Redesign 2026-06-15 (round 3) - new card UI (from Claude design mockup)
- Replaced the Bootstrap-3 panel rendering with a custom inline-styled **card grid** (matches the
  user's Claude-design mockup). All card CSS is INLINE (IT Glue go-live keeps inline styles); only
  the local-preview page chrome (body bg `#eef0f5`, font) uses a `<style>` block.
- Data model: tiles now return `New-Card -Title -Items`, where each item is `New-CardItem` with
  `Label, Sub, Link, Bg, IconDomain, Kind` (Kind 'pill' or 'number'). Renderers: ConvertTo-CardItemHtml,
  ConvertTo-CardHtml, New-DashboardHtml (single responsive grid), New-PreviewDocument.
- Cards are navy `#051554` (var `$script:CardColor`), `aspect-ratio:1/1`, grid
  `repeat(auto-fit, minmax(220px,1fr))`. Brand pills use Google favicon service
  `https://www.google.com/s2/favicons?domain=<d>&sz=64`. Brand colors: IT Glue #00A2B3,
  ConnectWise #54BEED, SentinelOne #6B0AEA, Bitdefender #ED1C24, Duo #6BBE4A, Veeam #00B336,
  Cove(endpoint) #0E9BA6; muted items (Cove 365, M365 licenses, Domains, Workstation/Server number
  cards) use rgba(255,255,255,0.12). Empty category renders a muted 'None' pill.
- `lib/ITGlue-BootStrapHelpers.ps1` is NO LONGER USED (dot-source removed). Left in repo for now;
  safe to delete. The .SYNOPSIS still says "gavsto Bootstrap-panel" - cosmetic, not updated.
- Favicons load from google.com (fine for preview; for IT Glue go-live confirm external <img> is
  allowed or swap to inline/data-URI icons).

## Round 4 2026-06-15 - Devices, Brandfetch, Users/agreement
- **Devices card**: merged Workstations + Servers into one 'Devices' card with a Workstations pill +
  a Servers pill (Get-DashDevices / Get-ConfigCountPill). Card order now: Users, Devices, Antivirus,
  MFA, Backup, M365 Licenses, Domains. FUTURE: add a stale count per type (devices not seen in
  ConnectWise Automate for 90+ days) - needs Automate API/DB (separate integration).
- **Brand styling via Brandfetch**: New-CardItem now takes `-Brand <domain>` (+ `-Muted`) instead of
  -Bg/-IconDomain. `$script:BrandMap` bakes accent color + logo URL per brand (from Brandfetch);
  `Resolve-Brand` returns baked entry or live-fetches+caches unmapped domains. **API key is hard-coded**
  in `$script:BrandfetchApiKey` (per user request) - SECURITY: visible in repo; restrict in Brandfetch
  (domain allowlist) or rotate. Logos load from cdn.brandfetch.io (replaced Google favicons). Colors:
  itglue #3860be, connectwise #5ea4de, sentinelone #6b0aea, bitdefender #EB0000, duo #74bf4b,
  veeam #03D15F, n-able/Cove #c046ff, microsoft #00A4EF. CAVEAT: logos are Brandfetch icon/symbol
  assets (mostly theme=dark) on a WHITE square - not visually verified; if any render white-on-white,
  switch that brand to a light-theme asset or favicon.
- **Users card restructure**: ConnectWise now ABOVE IT Glue, with TWO CW pills: "CWM Contacts: <n>"
  (active contacts, links to billing contact) and "CWM Agreement: <qty>" = quantity of part
  **CBT-PF-MEMBER** on the company's agreement name 'IT Services Agreement - PeopleFirst Support',
  type 'IT Services Agreement' (Get-CwmMemberQty). Skipped entirely if the company has no such
  agreement (Christine = skipped; Beaty & Brown = 6). Agreement TYPE field is `type.name` in CWM.

## Pending: M365 card (blocked on CIPP API)
User wants an 'M365' card = move the M365 Licenses list into it + add **MFA status** + **AD/Entra
sync** pills, pulled via the **CIPP API** (CIPP + GDAP to all clients - see [[m365-via-cipp]]).
CIPP instance: https://m365.codebluetechnology.com. API client app id 70ac3930-842e-4e4c-a1e7-f40ec5b81076.
BLOCKED: CIPP "Save Azure Configuration" failed ("Failed to save allowed API clients to Azure...") -
the CIPP function app lacks rights to edit its Authentication settings; the API client isn't yet an
allowed client so token auth will fail. Resolve that (CIPP managed identity needs Contributor on the
function app / resource group, OR manually add the client app id under Function App > Authentication >
allowed client applications), then build. Token: POST login.microsoftonline.com/<partnerTenantId>/oauth2/v2.0/token,
grant_type=client_credentials, scope=api://<appid>/.default (ensure that app's Application ID URI is set under Expose an API).

## M365 card 2026-06-16 - DONE (via CIPP)
- Renamed M365 Licenses card -> "M365": MFA pill + AD-sync pill on top, then the license list.
- CIPP creds in dashboard-credentials.xml ($creds.CIPP: ApiUrl https://cippz2zp4.azurewebsites.net,
  TenantId, ClientId, ClientSecret, Scope api://<clientid>/.default). Token = client_credentials.
- MFA: ListGraphRequest Endpoint=reports/authenticationMethods/userRegistrationDetails -> "MFA:
  <isMfaRegistered>/<total>". AD sync: Endpoint=organization -> onPremisesSyncEnabled +
  onPremisesLastSyncDateTime -> "AD Sync: Synced (Nh ago)" or "Cloud-only".
- TENANT MATCH: by DOMAIN. **GOTCHA**: bare /api/ListTenants intermittently returns a single
  AGGREGATED row (all 152 domains joined by spaces in one object) instead of 152 rows - this poisoned
  the match. FIX: query **/api/ListTenants?TenantFilter=<domain>** per org IT Glue domain (returns a
  clean single tenant; the tenant's DEFAULT domain is the one that works - secondary domains 400 in
  Graph with "not under scope"). Misses are NOT cached (re-resolve next run) to avoid poisoning.
- Verified: Christine -> MFA 30/94, AD Sync Synced. CIPP client secret was pasted in chat -> ROTATE.

## Automate links on Devices card 2026-06-16 - DONE (icon pending)
- Devices Workstations/Servers pills now link to the client's Automate computers page:
  `<AutomateUrl>/automate/browse/companies/computers?companyId=<id>`. Falls back to the IT Glue
  configs deep-link if the org has no Automate match.
- Automate creds in dashboard-credentials.xml ($creds.Automate: Url https://automate.codebluetechnology.com,
  Username **ITGlue** (NOT Itglueadmin - that user has local login disabled), Password, ClientId
  b25df2bd-...). AUTH: POST /cwa/api/v1/apitoken {UserName,Password} with **ClientId header** ->
  AccessToken; then Bearer + **ClientId header on every call** (data calls 401 without ClientId; the
  token POST tolerates its absence). Server is ConnectWise-SSO-bound (IsLocalLoginEnabled=false), so
  the API user must be local-login capable (ITGlue is; Itglueadmin wasn't).
- TENANT/CLIENT MATCH: query `/cwa/api/v1/clients?condition=Name contains '<token>'` then
  normalized-match (Resolve-AutomateClient). The bulk clients?pagesize=N list has the SAME aggregation
  flakiness as CIPP ListTenants (intermittent single merged row) - so we use the filtered query and
  do NOT cache misses. Verified: Christine -> companyId 131, Beaty & Brown -> 8.
- ICON: DONE. No Brandfetch entry / favicon 404s, so BrandMap['automate'].Logo is an inlined
  base64 PNG data URI (downscaled to 64x64 from the supplied IT Glue/automate.png). Source PNG can be
  deleted; the data URI is self-contained in the script.
- Type filter: the computers URL has no device-type param, so both pills link to the full computers
  list (the pill label still shows the workstation/server counts).

## Open items / next steps
0. **VEEAM: FIXED + verified 2026-06-15** (Dominion Leasing Software -> "3 Protected, 3 VMs").
   Confirmed live on veeam.codebluetechnology.com:1280:
   - Root cause of the 400: `limit=1000` exceeds VSPC's max of **500**. All collections now page via
     Get-VspcPaged at limit=500 (offset/`meta.pagingInfo.total`).
   - Auth: API key works DIRECTLY as `Authorization: Bearer <key>` (GET /api/v3/about = 200). No
     /token exchange. (No-auth /about = 401.)
   - Workload endpoints: `/api/v3/protectedWorkloads/virtualMachines` (VMs),
     `/protectedWorkloads/computersManagedByConsole` (Veeam Agents; `operationMode` =
     Server/Workstation), `/protectedWorkloads/computersManagedByBackupServer` (B&R-managed; also
     has operationMode). The wrong `managedVirtualMachines`/`computers` paths return the SPA HTML
     with 200. The `?filter=` query param BREAKS the computers endpoints (SPA fallthrough) though it
     works on virtualMachines - so we page each list and filter client-side by `organizationUid`.
   - KNOWN LIMITATIONS (future refinement): (a) no reliable used-storage/TB endpoint found -
     `/organizations/companies/<uid>/backupResources` returns the SPA - so the Veeam tile shows a
     protected-workload COUNT, not TB (Cove still shows TB). (b) server/workstation split currently
     only counts `computersManagedByConsole`; B&R-managed servers/hosts (computersManagedByBackupServer,
     e.g. DLS's Hyper-V host) are NOT counted as servers - so a VM-only client shows "0 servers,
     0 workstations, N VMs". Decide whether to fold computersManagedByBackupServer servers in.
1. **Run with vendors enabled** for a real client and verify each tile:
   - SentinelOne: `Get-S1Sites` proven pattern (from CWM-SyncSentinelOneCounts.ps1). Should work.
   - Duo: uses DuoSecurity module (`Set-DuoApiAuth` with `Type='Accounts'`, `Get-DuoAccounts`,
     `Select-DuoAccount`, `Get-DuoUsers`). Proven pattern. Counts users with status 'active'.
   - **Bitdefender: VERIFIED + FIXED 2026-06-15.** `getCompaniesList` lives on the NETWORK service
     (`/api/v1.0/jsonrpc/network`), NOT `companies` (old guess -> "Method not found"); it returns a
     FLAT {id,name} array, no pagination, no params. Endpoint count uses `getEndpointsList` with
     `parentId`=companyId + `filters.depth.allItemsRecursively=true`, paged (perPage max 100),
     counting items where `managedWithBest=true` ('total' over-counts unmanaged/stale inventory).
     Validated live: A-1 Door Company -> 29 protected (of 33 inventory). The managedWithBest *filter*
     does NOT work (returns 0) - count client-side instead.
   - **Veeam VSPC: endpoint shapes UNVERIFIED.** Auth is solid now (API key as permanent
     `Authorization: Bearer <key>`, validated via GET /api/v3/about). But
     `Get-VspcCompanies`/`Get-VspcCompanyBackup` paths (companies, protectedWorkloads,
     backupResources) need confirming against the live v3 API. Wrapped in try/catch so they degrade.
   - Cove: now uses API-USER auth (2026-06-15) per https://developer.n-able.com/n-able-cove/docs/authorization
     -> Login params are partner + username (API user) + password (the API user's generated token).
     The 'partner' param is REQUIRED; the old code omitted it, which is why it returned "Unknown
     partner/username or bad password". Creds now store PartnerName + ApiUser + ApiToken (the old
     PSCredential shape auto-triggers a re-prompt). EnumeratePartners + EnumerateAccountStatistics
     still mirror CWM-SyncNableCoveData.ps1; PartnerName also resolves numeric PartnerIds.
     M365-backup detection: VERIFIED + FIXED 2026-06-15. Cove M365 cloud accounts are NOT identified
     by DataSources/Product text (those are coded, e.g. AP="GJ", PN="All-In"). They are identified
     by **Physicality (I81) == "Undefined"** (matches CWM-SyncNableCoveData.ps1 line ~895). Get-DashBackup
     now EXCLUDES Physicality=='Undefined' (endpoint backup only) and the Backup tile is titled just
     "Backup" (provider shown in detail: "Cove - N servers, M workstations" / "Veeam - ..."). M365
     cloud backup now always renders on the separate M365 Backup tile (Physicality=='Undefined'),
     regardless of whether IT Glue has an M365 license asset. NOTE: server-vs-workstation split for
     real endpoints still uses OSType -match 'server'; Settings.OT came back numeric ("0") for the
     M365 account, so verify the server/workstation split once a client WITH Cove endpoints is run.
2. **M365 Licenses tile**: DONE/verified (see status). Possible polish: free/unlimited SKUs render
   noisily (e.g. "0/0", "0/1000000", "26/10000"). Could suppress the count suffix when active is 0
   or absurdly large, or hide zero-consumed SKUs.
   - **Resolver shorthand gap (NEW):** typing "Christine Rausch" did NOT resolve - the real name has
     a middle initial ("Christine S. Rausch, MD, PC") so neither normalized nor `*contains*` match
     hit. Had to use the org ID. Consider token-subset matching (all query words present) in
     `Resolve-ItgOrganization` so reasonable shorthands work.
3. **Domains -> registrar/DNS asset link**: currently links each domain to its IT Glue domain page
   (`/<orgId>/domains/<id>`). User confirmed the registrar relationship IS documented; enhance to
   follow related items to the registrar/DNS asset when feasible.
4. Consider adding a short README in `IT Glue/` for the dashboard (not yet created).

## Gotchas already solved (do not re-discover)
- **pwsh 7 vs Duo module conflict (RESOLVED 2026-06-15):** REST stack must run under Windows
  PowerShell 5.1 (pwsh 7's Invoke-RestMethod is broken here), but the `DuoSecurity` module v1.4.3
  manifest REQUIRES PowerShell 7.0 -> they can't coexist. Fixed by dropping the module and calling
  the Duo Accounts/Admin API directly via `Invoke-DuoApi` (HMAC-SHA1 signed REST). Script now has a
  hard guard that refuses to run under pwsh 7.
- **Duo 401 (RESOLVED 2026-06-15):** the native HMAC signing was correct, but Windows PowerShell 5.1's
  Invoke-RestMethod silently refuses to set the restricted `Date` header from -Headers, so Duo never
  saw the signed date -> 401. Fix: send the date in **`X-Duo-Date`** instead (Duo supports it for
  exactly this). Verified live: 25 Accounts API subaccounts enumerated; Christine Rausch -> 65 active
  Duo users. Key type IS Accounts API (MSP-level), as required.
- **Encoding**: PS 5.1 misparses non-ASCII (em-dash, middot, box-drawing) in string literals when
  the file has no BOM. Keep the script **pure ASCII**. Verify after edits:
  `([regex]'[^\x00-\x7F]').Matches((gc -Raw file)).Count` should be 0.
- **IT Glue `filter[name]` treats commas as a value list** -> names like
  "Christine S. Rausch, MD, PC" return nothing. Fixed: `Resolve-ItgOrganization` pages all orgs and
  matches client-side (id -> case-insensitive exact -> normalized -> unique contains -> pick-list).
- **VSPC has API keys now** (Simple Key). It's a permanent bearer credential used directly; NOT
  username/password and NOT a /token exchange. Ref:
  https://helpcenter.veeam.com/docs/vac/rest/api_keys.html
- **`Add-ITGlueAPIKey`** (module v2.2.0) stores the key as a global ReadOnly SecureString; passing a
  plain string works. A 401 = bad key or wrong region (EU = https://api.eu.itglue.com).

## Client matching design
No manual mapping. `Resolve-OrgServiceId` matches each vendor's candidate list to the org by
normalized name (case/punctuation/legal-suffix insensitive). Confident single match = silent;
ambiguous = pick-list; none = one-time prompt. All answers (incl. "not managed") cached in
service-map.json. `-Remap` clears and re-resolves.

## Go-live (out of scope now, planned)
Staged, no rework: the `Get-Dash*` extraction functions feed the renderer today; at go-live they
feed a sync step that writes a per-org **"Managed Services Summary"** flexible asset holding both
structured traits (counts, TB, LastSynced) AND an HTML Textbox trait with `$dashboardHtml`. Add a
`-Publish` switch / `Sync-ITGlueServicesDashboard.ps1`. service-map.json can also move into a
flexible asset. IT Glue sanitizes saved HTML (strips script/style/iframe) but keeps Bootstrap
classes, inline styles, and <a href> - which is why the panels render via IT Glue's own Bootstrap 3.

## SECURITY
User pasted a live IT Glue API key in chat on 2026-06-13. **Rotation confirmed by user 2026-06-15.**
Git identity: use **akolhoff-cbt** (user confirmed 2026-06-15), not the epiphanyplx identity that
auto-detected on commit 9aae81a.

## How to run
```powershell
cd "c:\cbt\Scripts\github\Scripts\IT Glue"
.\New-ITGlueServicesDashboard.ps1 -Organization "<client name>"        # uses cached creds/map
.\New-ITGlueServicesDashboard.ps1 -Organization "<client>" -Reset      # re-enter all credentials
.\New-ITGlueServicesDashboard.ps1 -Organization "<client>" -Remap      # re-resolve vendor matches
.\New-ITGlueServicesDashboard.ps1 -Organization "<client>" -NoOpen     # don't auto-open HTML
```
