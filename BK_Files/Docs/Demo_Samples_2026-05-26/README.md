# Demo Upload Pack — 26 May 2026

Pre-built CSV files ready to drop into the three import screens on the admin site.
All client CDS numbers, broker codes, and security symbols below are **real values from the live droplet** — they will resolve cleanly during import matching.

---

## Suggested upload order

1. **Prices** → Operations ▸ *Price Imports* ▸ Upload `01_prices_2026-05-26.csv`
   - Sets `SecurityMktPrice` for all 18 tradeable instruments. Quote date is read from the `Date:` header (26 May 2026).
2. **Holdings snapshot** → Operations ▸ *CDS Holdings* ▸ Upload `02_holdings_2026-05-26.csv`
   - Stages 15 holdings rows so every seller in the trade files has enough stock to deliver.
3. **Trade files (one per broker)** → Operations ▸ *CDS Trade Imports* ▸ upload each in turn:
   - `03_trades_SML_2026-05-26.csv` — 5 trades (Stockbrokers Malawi Ltd)
   - `04_trades_CCL_2026-05-26.csv` — 5 trades (CDH Capital Ltd)
   - `05_trades_CEDAR_2026-05-26.csv` — 1 trade (Cedar Capital Limited)
   - `06_trades_FSL_2026-05-26.csv` — 2 trades (FDH Stockbrokers Ltd)
   - `07_trades_TSL_2026-05-26.csv` — 3 trades (Trust Securities Ltd)

Each broker file represents that participant's view of the same matched-trade set. Across all five files there are **8 matched pairs** (16 lots, 8 contracts when committed).

---

## What's in the matched-pair set

| # | Security | Qty | Price (MWK) | Sett. Amount | Buy Broker | Sell Broker |
|---|---|---:|---:|---:|---|---|
| 1 | FMB | 100,000 | 7.00 | 700,000 | CCL | SML |
| 2 | AIRTEL | 50,000 | 17.80 | 890,000 | CEDAR | SML |
| 3 | NBM | 5,000 | 56.25 | 281,250 | TSL | FSL |
| 4 | NBS | 25,000 | 14.70 | 367,500 | SML | CCL |
| 5 | TNM | 200,000 | 1.88 | 376,000 | SML | CCL |
| 6 | NICO | 10,000 | 15.00 | 150,000 | FSL | CCL |
| 7 | ILLOVO | 1,000 | 150.00 | 150,000 | TSL | CCL |
| 8 | STANDARD | 2,000 | 115.00 | 230,000 | SML | TSL |

For full client-name detail (e.g. for the Add Order or Allocate Trade workflows), see [Demo_Sample_Trades_2026-05-26.md](../Demo_Sample_Trades_2026-05-26.md).

---

## Notes for the PM

- **Trade date** field on the upload form can be left blank — each file's `From:` header tells the importer it's 26 May 2026.
- **Quote date** on the price upload form is also read from the `Date:` header.
- All sellers' holdings already exceed the trade quantity (verified against `dbo.Lot` net positions), so the holdings file is for snapshot demonstration; the underlying data will pass holdings checks even if not imported.
- After uploading a CDS Trade file, the batch page will show **Committable** counts — review then *Commit batch* to materialise the lots/contracts.
- If any single row shows "UnknownClient" or "UnknownBroker", the matching will surface it — double-click the row in the batch view for details.
