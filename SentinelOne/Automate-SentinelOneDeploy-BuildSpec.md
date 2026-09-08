# Automate script: "SentinelOne Deploy"

Build spec for `Scripts > Antivirus > SentinelOne > SentinelOne Deploy`.

Must be built in the Control Center GUI. The Automate REST API exposes script **metadata
only** (`GET /scripts` returns Id, Folder, Name, Comments... and no steps) and
`POST /scripts` returns 403, so script logic cannot be authored over the API.

Five steps, no PowerShell. The installer takes the site token directly.

## Prerequisites

**1. Create the folder.** `Antivirus` is folder Id **12** and has no subfolders today.
Right-click it → New Folder → `SentinelOne`.

**2. Stage the installer** on the LTShare, e.g.:

```
\\192.169.254.52\LTShare\Transfer\Software\SentinelOne\SentinelOneInstaller_windows_64bit_v26_1_2_177.exe
```

Verified: that share is writable, and Automate serves it to agents over HTTPS at
`https://automate.codebluetechnology.com/LabTech/Transfer/<relative path>` (agents need
no SMB access). Keep the version in the filename — agents cache downloads, so reusing a
name is how a stale installer gets served.

**3. Client EDFs are already populated** for all 129 reconciled clients:

| EDF Title | Definition Id | Location |
|---|---|---|
| `S1SiteID` | 550 | Client (LocationId 3) |
| `S1ScopeToken` | 603 | Client (LocationId 3) |

## Script properties

- **Name:** `SentinelOne Deploy`
- **Folder:** `Antivirus > SentinelOne`
- **Script Type:** Computer Script

### Description / Notes

> Installs the SentinelOne agent and enrolls it into the correct customer site.
>
> The site is determined entirely by the **client-level EDF `S1ScopeToken`**, which this
> script reads for the client that owns the target computer and passes to the installer
> as `-t`. That value is the S1 site's registrationToken — base64 of
> `{"url":<console>,"site_key":<key>}` — so it carries both the console URL and the site
> identity. Nothing else selects the site: no API credentials, no site-name matching, no
> hardcoded token. Repoint the EDF and the agent enrolls somewhere else.
>
> `S1ScopeToken` (and the companion `S1SiteID`, EDF 550, kept for reporting) is populated
> by `Set-AutomateS1SiteFields.ps1`, which only writes clients that
> `Compare-S1SiteToAutomateClient.ps1` verified as a clean 1:1 Automate-to-SentinelOne
> match. If a client was never matched the EDF is empty and this script exits without
> installing rather than guessing a site — an empty token is a reconciliation problem,
> not a deployment bug.
>
> Assumes the machine has no SentinelOne agent. Machines already carrying an agent from
> another console are uninstalled manually first; this script does not re-home an
> existing agent, and no installer can.

## Steps

1. **Get the token into `%S1Token%`** — see the note below.

2. **If `%S1Token%` is empty** → Script Log `No S1ScopeToken for this client - not
   reconciled to an S1 site` → Script Exit.

   *Worth keeping: without it, an unreconciled client runs the installer with `-t ""`,
   which either fails obscurely or enrolls somewhere you didn't intend.*

3. **File Download** —
   `Software\SentinelOne\SentinelOneInstaller_windows_64bit_v26_1_2_177.exe`
   → `%windir%\Temp\SentinelOneInstaller.exe`

4. **Execute Script / Shell command** (as SYSTEM):

   ```
   "%windir%\Temp\SentinelOneInstaller.exe" -q -t "%S1Token%" --dont_fail_on_config_preserving_failures
   ```

5. **Script Log** `%output%` and the exit code. `0` is success.

### The `-q` flag

Add it. Without `-q` the unified installer can go interactive, and under the Automate
agent there is no desktop session to interact with — the process just sits there until
the script times out. `--dont_fail_on_config_preserving_failures` is worth keeping for
machines that had an agent previously.

### Step 1 — reading the client EDF

How a script pulls a client EDF varies by Automate version, and I could not verify the
variable syntax on your instance (the API exposes no script steps, so there was no
working example to read). Use whichever your version supports:

- **Preferred:** the native EDF accessor in the variable picker, if your Control Center
  offers one for client Extra Data Fields.
- **Fallback — SQL Query step** into `%S1Token%`:

  ```sql
  SELECT Value FROM extrafielddata WHERE ExtraFieldID = 603 AND ID = %clientid%
  ```

  `603` is confirmed correct for `S1ScopeToken`. Confirm the table name in your instance.
  Title-based form, if you'd rather not pin the id:

  ```sql
  SELECT d.Value FROM extrafielddata d
    JOIN extrafielddefine f ON f.ExtraFieldID = d.ExtraFieldID
   WHERE f.Title = 'S1ScopeToken' AND f.LocationHint = 3 AND d.ID = %clientid%
  ```

## Testing on CBT-LT25-04

Computer Id **13383**, client **CodeBlue Technology, LLC.** (client 1), Windows 11,
Online. That client's `S1ScopeToken` is populated and points at site
`2422689016592054876` (CodeBlue Technology End Users).

Note this machine's client already has agents in that site, so confirm CBT-LT25-04 itself
has no agent before testing, or you are testing an upgrade rather than a fresh install.

After running, confirm in the S1 console that the endpoint landed in **CodeBlue
Technology End Users** and not Default site — landing in Default site means the token
was empty or wrong.

## Installer choice

Use the `.exe`, not the `.msi`. The EXE wraps that same MSI (79.0 MB vs 73.6 MB) and the
extra weight is prerequisite checking and cleanup logic for machines that had an agent
before. MSI only pays off for GPO/SCCM/Intune-native deployment, which is not this.

## Optional

`Install-SentinelOneAgent.ps1` in this folder does the same install with guardrails —
validates the token decodes and points at the expected console, detects existing or
partial installs, waits for the service, and returns a parseable `RESULT=` line. Not
needed for the five-step version above; it exists if the simple script ever proves too
blunt on a machine that was not cleanly prepared.
