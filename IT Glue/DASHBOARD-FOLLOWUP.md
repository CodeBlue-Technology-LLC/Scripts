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
- `IT Glue/Config/credentials.xml` - secrets, git-ignored (Config/ is ignored).
- `IT Glue/Config/service-map.json` - auto-built client->vendor-id cache, git-ignored.
- `IT Glue/Output/*.html` - generated dashboards, git-ignored (Output added to .gitignore).

## Current status (as of 2026-06-13)
- Both scripts parse clean (PS 5.1), pure ASCII. Read-only verified (no ITGlue writes; only auth +
  JSON-RPC read POSTs to vendors).
- Prereqs installed on this machine: `ITGlueAPI` v2.2.0, `DuoSecurity`. PSGallery trusted.
- **IT Glue auth: WORKING.** Org lookup rewritten and working (see gotcha below).
- **Vendor tiles: NOT yet tested against live APIs.** User has only run the IT Glue-only path so
  far (vendors answered N). Next dry run should enable vendors one at a time.

## Open items / next steps
1. **Run with vendors enabled** for a real client and verify each tile:
   - SentinelOne: `Get-S1Sites` proven pattern (from CWM-SyncSentinelOneCounts.ps1). Should work.
   - Duo: uses DuoSecurity module (`Set-DuoApiAuth` with `Type='Accounts'`, `Get-DuoAccounts`,
     `Select-DuoAccount`, `Get-DuoUsers`). Proven pattern. Counts users with status 'active'.
   - **Bitdefender: UNVERIFIED.** `getCompaniesList` + `getMonthlyUsagePerProductType` for endpoint
     count are best-effort guesses. Confirm method names/endpoints against GravityZone API docs.
   - **Veeam VSPC: endpoint shapes UNVERIFIED.** Auth is solid now (API key as permanent
     `Authorization: Bearer <key>`, validated via GET /api/v3/about). But
     `Get-VspcCompanies`/`Get-VspcCompanyBackup` paths (companies, protectedWorkloads,
     backupResources) need confirming against the live v3 API. Wrapped in try/catch so they degrade.
   - Cove: auth + EnumeratePartners + EnumerateAccountStatistics mirror the working
     CWM-SyncNableCoveData.ps1. Needs the MSP's own Partner Name (collected in creds) to resolve
     numeric PartnerIds. M365-backup detection keys off DataSources/Product containing 365/Exchange/
     OneDrive/SharePoint - confirm against real Cove data.
2. **M365 Licenses tile**: reads IT Glue's native "Microsoft Licenses" (or legacy "Office 365
   Licenses") flexible asset. Trait-name probing is best-effort (regex over trait keys for
   name/active/consumed). Verify against a real synced asset and tighten key names.
3. **Domains -> registrar/DNS asset link**: currently links each domain to its IT Glue domain page
   (`/<orgId>/domains/<id>`). User confirmed the registrar relationship IS documented; enhance to
   follow related items to the registrar/DNS asset when feasible.
4. Consider adding a short README in `IT Glue/` for the dashboard (not yet created).

## Gotchas already solved (do not re-discover)
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
User pasted a live IT Glue API key in chat on 2026-06-13. Reminder to confirm it was **rotated**
(IT Glue > Account > Settings > API Keys). Git identity on commit 9aae81a auto-detected as
"Andrew <akolhoff@epiphanyplx.com>" - confirm desired identity.

## How to run
```powershell
cd "c:\cbt\Scripts\Scripts\IT Glue"
.\New-ITGlueServicesDashboard.ps1 -Organization "<client name>"        # uses cached creds/map
.\New-ITGlueServicesDashboard.ps1 -Organization "<client>" -Reset      # re-enter all credentials
.\New-ITGlueServicesDashboard.ps1 -Organization "<client>" -Remap      # re-resolve vendor matches
.\New-ITGlueServicesDashboard.ps1 -Organization "<client>" -NoOpen     # don't auto-open HTML
```
