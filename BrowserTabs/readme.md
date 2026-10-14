Collect-BrowserTabs.ps1
One-time collector: walks every user profile on an RDS host and stages the
browser session/history files needed to reconstruct open tabs.

Run in an ELEVATED PowerShell session on the RDS host (Local Admin).
Then run parse_tabs.py against the -Staging folder to get a CSV.

Firefox recovery.jsonlz4  -> reliable "current tabs"
Chrome/Edge Current Tabs  -> best-effort "current tabs" (SNSS, undocumented)
Chrome/Edge History       -> reliable visited-URL fallback

Files are copied via a Volume Shadow Copy snapshot so LOCKED files
(open browsers) can still be read. Falls back to a raw copy if VSS fails.
