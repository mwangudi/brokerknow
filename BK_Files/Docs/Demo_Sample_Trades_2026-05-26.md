# BrokerKnow Demo — Sample Trades, Holdings & Prices

**Date:** 26 May 2026
**For:** PM demo of Add Order → Allocate Trade workflow
**Source:** Live Malawi droplet snapshot (real clients, real holdings)

All clients, brokers and securities below exist in the system and can be picked from the dropdowns / lookups. CDS numbers, IDs, and balances are already on file.

---

## 1. Price list (use these prices when allocating)

| Security | Code | Price (MWK) |
|---|---|---:|
| Airtel Malawi PLC | AIRTEL | 17.80 |
| Blantyre Hotels PLC | BHL | 7.00 |
| FDH Bank PLC | FDHB | 10.00 |
| First Merchant Bank | FMB | 7.00 |
| FMB Capital Holdings | FMBCH | 45.01 |
| ICON Properties | ICON | 8.75 |
| Illovo Sugar Malawi PLC | ILLOVO | 150.00 |
| Malawi Property Investment Co | MPICO | 2.50 |
| National Bank of Malawi | NBM | 56.25 |
| NBS Bank | NBS | 14.70 |
| NICO Holdings PLC | NICO | 15.00 |
| National Investment Trust PLC | NITL | 17.50 |
| Old Mutual plc | OML | 520.00 |
| Old Mutual Malawi Limited | OMU | 1,866.01 |
| Press Corporation PLC | PCL | 188.00 |
| Standard Bank PLC | STANDARD | 115.00 |
| Sunbird Tourism PLC | SUNBIRD | 7.00 |
| Telekom Networks Malawi | TNM | 1.88 |

---

## 2. Brokers on the system

| Code | Name |
|---|---|
| SML | Stockbrokers Malawi Ltd |
| CCL | CDH Capital Ltd |
| CEDAR | Cedar Capital Limited |
| FSL | FDH Stockbrokers Ltd |
| TSL | Trust Securities Ltd |
| ASL | Alliance Stockbrokers |

---

## 3. Opening holdings (sellers have stock to deliver)

These are real net-long positions already in the system, so each sell side below will pass the holdings check.

| Seller Client | Security | Holding | Selling |
|---|---|---:|---:|
| Livingstone Exports Ltd | FMB | 16,446,961 | 100,000 |
| Chandrakant Makadia | AIRTEL | 5,885,050 | 50,000 |
| Press Trust | NBM | 1,133,041 | 5,000 |
| SIGELEGE BEACH RESORT & CONFERENCE CENTRE | NBS | 11,791,968 | 25,000 |
| CDH Asset Management Ltd | TNM | 43,684,860 | 200,000 |
| CDH Asset Management Ltd | NICO | 36,915,443 | 10,000 |
| CDH Asset Management Ltd | ILLOVO | 2,935,857 | 1,000 |
| Satemwa Trust | STANDARD | 100,219 | 2,000 |

---

## 4. Matched trade pairs (8 ready-to-allocate trades)

Each row is **one matched trade** — same security, quantity, price and trade date. Settlement is T+3 (Friday 29 May 2026). Use the suggested MSE slip numbers, or any sequence you prefer.

**Trade date:** 26 May 2026 · **Settlement date:** 29 May 2026

| # | Slip | Security | Qty | Price | Gross (MWK) | Buy Broker | Buyer | Sell Broker | Seller |
|---|---|---|---:|---:|---:|---|---|---|---|
| 1 | MSE002150/1 | FMB | 100,000 | 7.00 | 700,000.00 | CCL (CDH) | Dheeraj Dikshit | SML | Livingstone Exports Ltd |
| 2 | MSE002150/2 | AIRTEL | 50,000 | 17.80 | 890,000.00 | CEDAR | Berc Trust | SML | Chandrakant Makadia |
| 3 | MSE002150/3 | NBM | 5,000 | 56.25 | 281,250.00 | TSL | Satemwa Trust | FSL | Press Trust |
| 4 | MSE002150/4 | NBS | 25,000 | 14.70 | 367,500.00 | SML | Benedicto Nkhoma | CCL (CDH) | SIGELEGE BEACH RESORT & CONFERENCE CENTRE |
| 5 | MSE002150/5 | TNM | 200,000 | 1.88 | 376,000.00 | SML | Livingstone Exports Ltd | CCL (CDH) | CDH Asset Management Ltd |
| 6 | MSE002150/6 | NICO | 10,000 | 15.00 | 150,000.00 | FSL | Ramesh Haridas Savjani | CCL (CDH) | CDH Asset Management Ltd |
| 7 | MSE002150/7 | ILLOVO | 1,000 | 150.00 | 150,000.00 | TSL | LIFECO UNRESTRICTED PENSION FUND | CCL (CDH) | CDH Asset Management Ltd |
| 8 | MSE002150/8 | STANDARD | 2,000 | 115.00 | 230,000.00 | SML | National Investment Trust Limited | TSL | Satemwa Trust |

---

## 5. How to use this in the demo

For each row above:

1. **Add Order — Buy side**
   Client = buyer · Broker = buy broker · Security = security · Qty / Price as above.
2. **Add Order — Sell side**
   Client = seller · Broker = sell broker · same Security / Qty / Price.
3. **Allocate Trade** on either order, enter the matching MSE slip number, trade date 26 May 2026, settlement 29 May 2026.
4. Confirm the contract generates, levies populate, and the lot appears under both clients' contracts.

For a single end-to-end sanity check, start with **Trade #5 (TNM)** — same case as the PM's earlier screenshot, now unblocked by today's fix.

---

## 6. Notes & caveats

- Buyer cash balances were not verified — if a buyer is over their credit limit the system will warn but still allow. For the demo, smaller-value trades (#3, #6, #7, #8) are safest.
- All sellers above have **far more** holdings than the trade quantity, so the holdings check will pass cleanly.
- Client names are case-sensitive in the search box but partial-match works — typing "Livingstone" or "Chandrakant" is enough.
- If a security/broker dropdown shows additional codes not in the table above, ignore — the eight listed are confirmed tradeable today.
