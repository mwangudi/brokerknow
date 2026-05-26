# CDS Trade Imports — sample data

This folder contains a real-world sample for testing the **Operations → CDS Trade Imports** screen.

## Files

| File | Purpose |
| --- | --- |
| `MSE_Transaction_Statement_2026-05-12.csv` | A 172-row Crystal Reports export from MSE for **12 May 2026** (participant **CEDAMWMW**, market **MAINBOARD**), covering 11 securities: AIRTEL, FDHB, FMBCH, ILLOVO, NBM, NBS, NICO, NITL, PCL, SUNBIRD, TNM. Converted from the original `.xls` so the parser can read it directly. |
| `MSE_Daily_Prices_2026-05-12.csv` | End-of-day MSE board prices for **12 May 2026** covering all 17 actively-listed counters. Use this with **Operations → Price Imports**. |
| `MSE_Holdings_Statement_2026-05-12.csv` | Sample CSD holdings snapshot for **12 May 2026** (CDSNo / Symbol / Quantity / AccountStatus / BalanceFree). Use this with **Operations → CDS Holdings Imports**. Re-uses the same demo clients seeded by `seed_cds_import_demo.sql` and includes one **UnknownClient** row (`CEDAXXX9999999999999999`) and one **UnknownSecurity** row (`XYZNEW`) so you can exercise both unmatch reasons. |
| `seed_cds_import_demo.sql` | Idempotent T‑SQL script that creates the broker, demo clients, and open released orders so the upload above lands one row in **every** classification (commit-eligible, OrderHeld, OverAllocated, NoPendingOrder, UnknownClient). |

## How to use

1. **Run the seed** against your `TEST_Malawi_2` (or whichever) database — e.g. in SSMS:

   ```text
   :CONNECT YourSqlServer
   USE TEST_Malawi_2;
   :r BK_Files\Samples\seed_cds_import_demo.sql
   ```

   The script:
   - Widens `dbo.Broker.BrokerCode` to `nvarchar(20)` so SWIFT/BIC participant codes fit.
   - Inserts broker `CEDAMWMW` if missing.
   - Inserts five demo clients with the CDS numbers from the CSV.
   - Creates five released open orders (qty/price chosen to produce one of each classification).
   - Picks a sensible default for every NOT NULL FK (`Branch`, `Class`, `Commission`, `Residency`, `OrderType`, `OrderSecType`, `OrderHoldType`) by selecting the first existing row, so it works on any restored Malawi backup.

2. **Open** `Operations → CDS Trade Imports` in the web app.

3. **Upload** `MSE_Transaction_Statement_2026-05-12.csv` (Exchange = MSE, trade date auto-detected as 12 May 2026).

4. After reconciliation runs, switch between the **View Imported / View Unmatched / View Committed** tabs. You'll see:

   | CSV row | Expected outcome |
   | --- | --- |
   | `…7243 / TNM / Buy 33001` | **Ready · Order #** (commit-eligible) |
   | `…0281 / TNM / Sell 33001` | **Ready · Order #** (commit-eligible) |
   | `…7731 / SUNBIRD / Buy 190` | **Ready · Order #** (commit-eligible) |
   | `…4342 / FDHB / Buy 693` | **Order on hold** badge |
   | `…7456 / FDHB / Buy 2171` | **Over-allocated** badge |
   | `…4037 / FDHB / Buy 90` | **No pending order** badge |
   | `…5648 / SUNBIRD / Sell …` | **Unknown client** badge |
   | every other row | **Unknown client** (no order seeded) |

5. **Tick** one or more eligible rows and click **Commit selected**, or click **Commit all eligible** in the action bar — each commit calls `IContractService.CreateContractAsync`, which writes a `Contract`, `Lot`, and placeholder `LevyContract` rows just like the legacy `cont_CreateContract` SP.

6. After committing, the rows move from **View Imported** → **View Committed** with a clickable `Contract #N` link.

## Re-running

The seed script is idempotent — every insert is gated by an `EXISTS` / `NOT EXISTS` check. Running it twice has no extra effect.

To reset the staging table (without touching contracts/lots that were already committed):

```sql
DELETE FROM dbo.CdsImportedTrades WHERE BatchId IN (SELECT BatchId FROM dbo.CdsImportedTrades);
```
