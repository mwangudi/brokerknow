# Demo Import Files — Client Demo (Hosted Env)

This folder contains six CSV files for the PM's client demo on the hosted
environment (46.101.6.131), grouped into three pairs. For each importer
(Prices, CDS Trades, CDS Holdings) there is one **small "test"** file you can
upload first to confirm the importer is working in the new environment, and
one **larger "demo"** file to show the PM during the actual session.

All files target the **MSE** exchange and use the 21 securities currently
seeded on the droplet (`AIRTEL`, `BHL`, `FDHB`, `FMB`, `FMBCH`, `ICON`,
`ILLOVO`, `MPICO`, `NBM`, `NBS`, `NBSR`, `NICO`, `NITL`, `OML`, `OMU`, `PCL`,
`PIM`, `REAL`, `STANDARD`, `SUNBIRD`, `TNM`).

---

## Files

| Pair      | Test (run first)                  | Demo (show to PM)                 |
| --------- | --------------------------------- | --------------------------------- |
| Prices    | `prices_test.csv`                 | `prices_demo.csv`                 |
| Trades    | `cds_trades_test.csv`             | `cds_trades_demo.csv`             |
| Holdings  | `cds_holdings_test.csv`           | `cds_holdings_demo.csv`           |

Upload via:

* **Prices**   → Operations → Price Imports → Upload CSV
* **Trades**   → Operations → CDS Trade Imports → Upload CSV
* **Holdings** → Operations → CDS Holdings Imports → Upload CSV

For Trades & Holdings, pick **Exchange = MSE** and set the **Trade Date** to
today (or whatever the demo date is).

---

## CDSC numbering — important

The droplet currently has **5 219 clients** loaded, of which **only one** has
a `ClientCDSNo` populated. Both the CDS-trade and CDS-holdings importers match
exclusively on `Client.ClientCdsNo`. With CDS missing, every staged row will
land as **Unmatched / `UnknownClient`** with the note
*"Client CDS 'XYZ' is not on file."*

There are two ways to handle this for the demo:

### Option A — Recommended: pre-seed dummy CDS numbers on six demo clients

Run the SQL block in [`seed_demo_cds.sql`](seed_demo_cds.sql) **once** on the
droplet before the demo. It assigns `DEMO001 … DEMO006` to clients
`Client_DPA_` 1-6 (Tony De Castro, Samuel Kalake, Patricia Ramani, Lester
Tandwe, Robert Mdoka, Jellings Chiumia). The same six codes are used as the
`Client Code` / `CDSNo` values in the demo files — so the **Holdings** import
will reconcile cleanly (Matched), and the **Trades** import will progress one
step further to either `Matched` or `NoPendingOrder` (because no pending
orders exist for those clients yet — see *Trade caveat* below).

### Option B — Demo the unmatched-rows review UI

If we don't want to touch client records, leave CDS empty and treat every
trade/holding row as a deliberate "unmatched" case. The Imports page will show
all rows with `UnknownClient` reasons, which lets us walk the PM through the
operations team's normal reconciliation workflow (manual client-link,
re-reconcile, commit). The file format is identical; only the matching
outcome differs.

> **Suggested approach**: do **A** for the Holdings & Prices demos (clean
> "everything matched / committed" story), and use **B** for one Trades file
> to show the unmatched-row resolution flow. That covers both narratives.

### Trade caveat (independent of CDS)

The trade reconciler additionally requires:

1. A **registered broker** matching the file's `Participant` code (the demo
   files use `BK001` — adjust if your droplet uses a different code, see the
   Broker Settings page).
2. An **open order** for the same client + symbol + side. Without pending
   orders, matched rows will be flagged `NoPendingOrder`, which is still a
   useful "we caught it before committing" story for the PM.

If you want a fully-green trade import, create matching pending orders for
the same six clients/symbols before uploading.

---

## Quick sanity check after each upload

In the Imports page, the batch row should show:

* `RowCount` equal to the data rows in the CSV (3 for the test files, 8-12 for
  the demo files — see headers of each file for the exact count).
* `MatchedCount` + `UnmatchedCount` = `RowCount`.
* No 5xx errors in the toast.
