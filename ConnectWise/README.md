# ConnectWise Manage billing tools

## Find-CWMBillingAnomalies.ps1

**Read-only** auditor that scans active ConnectWise Manage agreements for likely
billing inconsistencies. It only reads from CWM — it never creates, updates, or
deletes anything.

### What it checks

| Finding | Meaning |
| --- | --- |
| Superseding-part duplication | A new sync part billed **alongside** the old parts it replaced (double billing). |
| Duplicate line item | The same product appears more than once on one agreement. |
| Coverage gap | Within the PeopleFirst Support book only: a product most of those clients are billed for, but some are missing (possible under-billing). Scoped via `-CoverageAgreementNameLike`. Old parts (any in a `$SupersededParts` `Replaces` list) are excluded, since a client missing one has simply moved to the replacement part. |
| Billable but free | An addition marked `Billable` with a `$0` unit price. |
| Zero quantity | An active billable addition with quantity exactly `0` (bills nothing). |
| Negative margin | Unit cost is greater than unit price. |
| Price outlier | A client billed well off the median **per-unit** price for a product (often stale pricing). |
| Quantity mismatch | Configured products whose seat counts should track each other but don't. |

Every check only considers **currently-billing** additions: effective date on or
before today, and not cancelled (or cancelled with a future date). Audited agreement
type defaults to **IT Services Agreement** only.

### Month-over-month agreement totals

A separate report (`*_CWM_AgreementMoMChanges.csv`) recreates a month-over-month
total comparison. For each **IT Services** agreement it reconstructs the recurring
total ("what would be billed") as of the **1st of last month** and the **1st of this
month** — by summing the Billable additions that were active on each date — and flags
any agreement whose total moved by at least **$100** or **5%**
(`-MoMDollarThreshold`, `-MoMPercentThreshold`, `-MoMAgreementType`).

Columns: Company, Agreement, PercentChange, Delta, Direction, LastMonth,
LastMonthTotal, ThisMonth, ThisMonthTotal.

It catches lines **added** or **cancelled** between the months (cancelled lines stay
on the record with a cancelled date). It does **not** see a line that was hard-deleted
from the agreement, nor an in-place quantity/price edit to a line present in both
months — capturing those would require saving a monthly total snapshot on each run.

### Credentials

Reuses an existing `Config\Credentials.xml` that contains the `CWMConnectionInfo`
block created by the Veeam / SentinelOne / Cove sync scripts. The file is
DPAPI-encrypted and tied to the user/machine that created it.

- Pass `-ConfigPath` to point at a specific file, **or**
- Run it with no `-ConfigPath` and it auto-discovers a sibling `Credentials.xml`
  under the repo that has a `CWMConnectionInfo` block.

### Usage

```powershell
# Auto-discover credentials, audit IT Services agreements
.\Find-CWMBillingAnomalies.ps1

# Point at a specific credentials file
.\Find-CWMBillingAnomalies.ps1 -ConfigPath ..\Veeam\Config\Credentials.xml

# Add Data Center agreements back in, looser coverage threshold
.\Find-CWMBillingAnomalies.ps1 -AgreementTypes 'IT Services Agreement','Data Center Agreement' -CoverageThreshold 0.4
```

### Output

- Console: a per-category count summary.
- `Reports\<timestamp>_CWM_BillingAnomalies.csv` — the actionable findings
  (Category, Severity, Company, Agreement, Product, Detail).
- `Reports\<timestamp>_CWM_Additions_Raw.csv` — every addition pulled, for your own
  pivoting / spot-checks.

### Tuning

Edit the config blocks near the top of the script:

- **`$SupersededParts`** — add `@{ NewPart = 'new-sku'; Replaces = @('old-sku-1','old-sku-2') }`
  entries as more sync re-mappings are discovered. This is what catches the
  Barnes & Diehl `automate-patch` / `cbt-ms-av-edr` cases.
- **`$CorrelatedQuantityGroups`** — list products whose quantities should match per
  client (e.g. matching workstation seat counts across AV and patching). Empty by
  default.

Thresholds are parameters: `-CoverageThreshold`, `-PriceOutlierTolerance`,
`-MinSamplesForPricing`.
