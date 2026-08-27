# BrokerKnow — System Status

**As at 2026-08-26.** Figures read from the live droplet, not from notes.
Companion documents: `Ops_Runbook.md` (how to operate it), `PM_Feedback_Backlog.md`
(what the PM asked for), `IPO_And_Rights.md` (the newest feature).

---

## 1. Where the system stands

BrokerKnow is live for **Cedar Capital Malawi** and running in parallel with the legacy
Axis system, which remains the book of record — Cedar send a database dump every week or
two and we replace production with it. Rwanda and Kenya exist as separate tenants on the
same droplet but are not carrying live business.

| Tenant | Status | Database | Port |
| --- | --- | --- | --- |
| **Malawi production** | **Live** | `axis_db_prod` | 5260 |
| Malawi test | Live, for demos/UAT | `BrokerKnow_Malawi0612` | 5261 |
| Rwanda | Provisioned, not trading | `BrokerKnow_RW_Clean` | 5262 |
| Kenya | Provisioned, not trading | `BrokerKnow_KE_Clean` | 5264 |

### Live Malawi book

| | Count |
| --- | ---: |
| Clients | 6,015 (5,990 undeleted) |
| Orders | 20,778 |
| Contracts | 51,035 |
| Payments | 43,792 |
| Staff logins | 39 (33 active staff) |
| External logins (agents + clients) | 2 |
| Clients with a CDS number | 4,367 |
| Clients still without one | 1,623 |

### Surfaces in use

| Surface | Host |
| --- | --- |
| Admin | `cedarcapital.martensafrica.com` |
| Client portal | `cedarclient.martensafrica.com` |
| Agent portal | `cedarclient.martensafrica.com/agent/` |
| Test admin | `cedartest.martensafrica.com/admin/` |
| Kenya | `ke.martensafrica.com` |
| Rwanda test | `rwandatest.martensafrica.com` |

TLS is Let's Encrypt via `certbot.timer` (twice daily, currently healthy). Earliest
expiry is `pharma.localinvestors.co.ke` on 2026-09-30; the Cedar certificates run to
2026-11-09.

---

## 2. What has been delivered

Broadly, the PM's testing feedback (`PM_Feedback_Backlog.md`) is worked through, plus
several larger pieces built on top:

- **Orders → contracts → settlement**, with T+3 settlement dates that skip weekends and
  a maintainable holiday calendar.
- **Contract notes** matching the legacy wording and layout, with a real signature
  stamped on the PDF, plus a maker/checker reversal flow.
- **Payments and receipts** posting directly for back-office staff, with maker/checker
  retained for external agent and client cash requests. Balance checks enforced.
- **Client onboarding** — account opening for Individual / Joint / ITF / Corporate,
  risk scoring, multi-stage approval, and generated CSD1 / account-opening /
  residential-verification PDFs filed against the client.
- **Client portal and agent portal**, including order placement and document upload.
- **Reports** — client listing, statements, contract-note register, agents' commission,
  plus audit reports for client deletions, contract alterations and payment alterations.
- **IPOs and rights issues** (the most recent feature) — offerings, applications from
  both admin and portal, batching, remittance, allotment and refunds. Documented for the
  PM in `IPO_Feature_Summary_PM.md`.
- **Multi-tenant white-labelling** so Rwanda and Kenya run from the same codebase.

---

## 3. How production data is kept current

Every refresh replaces `axis_db_prod` in full, gated on reconciliation, penny-exact
money totals, foreign-key integrity and a smoke test, with the previous database kept
as `axis_db_prod_pre<NNNN>` for rollback. Full procedure in `Ops_Runbook.md §6`.

Two things must be restored after every import, because the legacy dump does not carry
them: **CDS numbers** and the **swapped bank-account name/number** fix.

### CDS numbers

Cedar's August 2026 client list arrived in a new shape — a **Reserve Bank of Malawi
"Client Accounts Report"** rather than Cedar's own listing, with no client code to key
on. Numbers are matched on identity document and name, tiered by how strong the evidence
is, and only the confident tiers are applied:

| Tier | Basis | Count | Applied |
| --- | --- | ---: | --- |
| A | Identity document matches exactly | 4,133 | Yes |
| B | Exact name + near-identical document | 373 | Yes |
| C | Exact name, documents clearly disagree | 315 | **No** |
| D | Name maps to more than one client | 5 | **No** |
| E | No counterpart in our book | 4,649 | n/a |

4,367 clients hold a number; none is duplicated. Two spreadsheets are with Cedar for
review — one of the 1,590 clients without a number (now 1,623), one of the 4,367 to
confirm. Since refresh #10 the numbers are carried across from the previous database
rather than re-derived, so nothing the team corrects in the app is overwritten.

---

## 4. Open items

| # | Item | Blocked on | Notes |
| --- | --- | --- | --- |
| 1 | **Recurring live-only lots** | Investigation | 5 contracts have lot rows on live but not in the legacy dump, so they lose trade detail at each cutover; 4 have survived several refreshes. Data recoverable from `axis_db_prod_pre0826`. Root cause unknown — worth establishing, as it repeats. |
| 2 | **CDS review returns** | Cedar | 1,623 clients need a number; 4,367 need confirming. Both spreadsheets issued 2026-08-19. 267 of the gap rows already carry a suggested number to verify. |
| 3 | **Levy account mapping** | Cedar's accountant | Which levy Entity account maps to which `LevyContract.SystemMaintained` code. Deliberately not guessed — it feeds a financial statement. |
| 4 | **Agent portal DNS** | Cedar's DNS admin | Add A record `cedaragent` → `46.101.6.131` (TTL 300) in cPanel → Zone Editor, then run the staged `/root/cedaragent-setup.sh`. The build is already on the droplet at `/var/www/cedaragent`. |
| 5 | **Support phone number** | Cedar | Needed for `VITE_SUPPORT_PHONE`; portal currently ships without one. |
| 6 | **IPO remittance presentation** | Cedar's accountant | Currently posts one payment per client. The alternative is a single lump sum to the NBSR nominal account. Both are defensible; needs a decision before the next offering. |
| 7 | **IPO phase 3** | Cedar (file formats) | Real bank/registrar file layouts (none exist to copy), posting allotted shares into `Holdings`, and executing cheque/EFT refunds. |
| 8 | **Data quality on external logins** | Cedar | Agent `brownleer@africanalliance` has an email with no TLD, so mail to it can never deliver. Neither remaining agent login has ever been used. |
| 9 | **4,649 register accounts with no client** | Cedar | CDS accounts on the RBM register that match no client in Cedar's book. May be dormant, closed, or held through another broker — worth a view from Cedar. |
| 10 | ~~**Droplet disk hygiene**~~ | **Done 2026-08-27** | 15 GB reclaimed (58% → 38% used, 48 GB free): 100 stale API install backups, 95 stale web roots, 66 deploy tarballs, apt cache, and 5 superseded rollback databases. `pre0818` and `pre0826` kept. Recurrence prevented — `ops/deploy-*.sh` now keep only the newest 3 backups per target, and journald is capped at 200 MB. |
| 11 | **Backups never leave the droplet** | Spaces credentials | Nightly full + 15-minute log backups run correctly to `/var/backups/brokerknow` (8.3 GB, 14-day retention), but the offsite `rclone` step logs `Offsite SKIPPED` every night — rclone is not installed and no `spaces:` remote exists. Backups sit on the same disk as the data, so losing the droplet loses both. |
| 12 | **Both SSH keys on one machine** | Your action | A second key (`id_ed25519_do`) was authorised and tested on 2026-08-27, so `/root/.ssh/authorized_keys` now holds two working entries. Both private keys still sit on the same laptop, so losing the machine still means losing access — keep a copy of one off-machine (password manager or another machine). Password auth remains off; the DigitalOcean console is the last resort. |
| 13 | **`sa` password still in git history** | Decision to rewrite + rotate | Removed from all tracked files on 2026-08-27 — the ops scripts now source `/etc/brokerknow/db.conf` (root-only, mode 600) and abort loudly if it is absent. But it sat in plaintext for months, so it **remains in history on the remote**; clearing that needs `git filter-repo` and a force-push. The value should also be rotated, which touches four `appsettings.json`, `/etc/brokerknow/db.conf` and the `sa` login itself. `sa` is still the only SQL login. |
| 14 | **Client keys collide between BrokerKnow and Axis** | Decision on approach | Root cause of the mis-linked logins, now understood (section 5). A pre-cutover guard catches recurrence, but the underlying clash remains: both systems allocate `Client_DPA_` from `MAX+1` over the same range. Options are to give app-created clients a reserved high block, or to stop creating clients in the app and require them in Axis first. Until then, portal-approved clients are lost at each refresh. |
| 15 | **`linkExisting` has no sanity check** | Small fix | Approving a registration against an existing client verifies only that the client exists — not that the name, email or ID resembles the applicant. That is how login 141 was attached to an unrelated client. A warning on mismatch would have caught it. |

---

## 5. Portal logins

35 logins: 33 staff (legacy-linked) and 2 agent logins, neither yet used.

Four client portal logins were **removed on 2026-08-27**. Each carried a name and
email that did not match the client record it was attached to — for example
`shell123@gmail.com` ("Katherine Client") was linked to client 5923, Harry Kaminjolo.
Anyone signing in would have seen another client's portfolio and statements, and one
of the four had been used on 14 August. The client records themselves are genuine and
were left untouched; only the logins were deleted, with a full-row snapshot retained
in `dbo.PortalUsers_removed_20260827`.

> Whatever allowed those logins to be attached to unrelated clients has now been
> traced — it was not a portal bug. See below.

### Why they pointed at the wrong clients

The links were **correct when they were made**. In the database as it stood before
the 18 August refresh, client 5921 really was Dominic Mutinyu, 5923 really was
Katherine, 5924 really was Ken Ross — each matching its login exactly.

The refresh replaced `dbo.Client` wholesale from the legacy dump while the
app-layer `PortalUsers` table was grafted across intact. Clients created in
BrokerKnow take their key from `MAX(Client_DPA_)+1` **in the app database**, and
Axis independently allocates the same numbers to different people. So after
cutover those ids referred to Madalitso Kadzeya, Harry Kaminjolo and Mayamiko
Kapanda, and the logins silently followed.

Two consequences, both wider than the four logins:

1. **Any client created in BrokerKnow but not entered in Axis is destroyed at the
   next refresh** and its id handed to someone else. Four were lost this way
   (5921–5924). `OnlineRegistration` is `0` on all 6,015 rows, so app-created
   clients cannot even be identified afterwards.
2. **The delta check could not see it** — it compares clients by `Client_DPA_`, so
   a collision looks like a match. It reported “0 clients live-only” while three
   were being overwritten by different people.

A guard now runs before cutover — `BK_Files/refresh_identity_check.sql` compares
identity rather than id and aborts if a client id a portal login points at would
change person. Replayed against the 18 August data it correctly aborts and names
all three logins.

Login 141 was a different fault: it was already attached to an unrelated client
*before* any refresh. The approval path’s `linkExisting` branch only checks that
the chosen client exists — it never checks the client plausibly belongs to the
applicant.

---

## 6. Operational notes worth knowing

- **Access from a new machine** is documented in `Ops_Runbook.md §2`. Two SSH keys are
  authorised, but both private keys are on the same laptop — keep a copy of one
  somewhere else. Secret values live in `Access_Credentials.local.md`, which is git-ignored.
- **Deploys are per-instance.** A change to shared API code needs deploying to all four
  services; a web change needs the correct bundle per root (see `Ops_Runbook.md §1`).
- **Never `git add -A` in this repo.** `BK_Files/` holds client PII and real signatures
  that are deliberately git-ignored. Stage explicitly.
- **The contract-note signature is not in git.** A fresh clone publishes without it and
  contract notes silently fall back to a placeholder — see `Ops_Runbook.md §7.4`.
- **Rollback databases are the safety net** for anything an import drops; keep at least
  the two most recent.
