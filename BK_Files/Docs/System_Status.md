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
| 10 | **Droplet disk hygiene** | Approval to delete | 104 old API install backups (**12 GB**) and 95 stale web-root backups (959 MB) have accumulated in `/opt` and `/var/www` — about a third of the used disk. Also 7 rollback databases; `pre0708` … `pre0810` are safe to drop. 34 GB free, so not urgent, but it grows with every deploy. |
| 11 | **Backups never leave the droplet** | Spaces credentials | Nightly full + 15-minute log backups run correctly to `/var/backups/brokerknow` (8.3 GB, 14-day retention), but the offsite `rclone` step logs `Offsite SKIPPED` every night — rclone is not installed and no `spaces:` remote exists. Backups sit on the same disk as the data, so losing the droplet loses both. |
| 12 | **Single SSH key, no fallback** | Your action | `/root/.ssh/authorized_keys` has one entry and password auth is disabled. Lose that key and the only way back is the DigitalOcean console. Add a second key before it bites — `Ops_Runbook.md §2` has the commands, and an unused `id_ed25519_do` key already exists locally. |
| 13 | **`sa` password committed to git** | Decision to rotate | It appears in plaintext in six tracked files (runbook, `ops/backup-*.sh`, `ops/bootstrap-test.sh`, `ops/install-backups.sh`, `BK_Files/Demo/seed_demo_cds.sql`) and so is in history on the remote. `sa` is also the only SQL login — every API instance uses it. Rotating touches four `appsettings.json`, both backup scripts, and any helper that embeds it. |
| 14 | **Portal logins attached to the wrong client** | Investigation | The four logins removed on 2026-08-27 were each linked to a real but unrelated client record, and one had been used. Whatever allowed that link to be made has not been traced — until it is, the same thing can happen on the next self-registration. See section 5. |

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

> Whatever allowed those logins to be attached to unrelated clients has not been
> traced — open item 14. Until it is, the same thing can recur on self-registration.

---

## 6. Operational notes worth knowing

- **Access from a new machine** is documented in `Ops_Runbook.md §2`. There is only
  one SSH key on the droplet and no password fallback — add a spare before you need it.
  Secret values live in `Access_Credentials.local.md`, which is git-ignored.
- **Deploys are per-instance.** A change to shared API code needs deploying to all four
  services; a web change needs the correct bundle per root (see `Ops_Runbook.md §1`).
- **Never `git add -A` in this repo.** `BK_Files/` holds client PII and real signatures
  that are deliberately git-ignored. Stage explicitly.
- **The contract-note signature is not in git.** A fresh clone publishes without it and
  contract notes silently fall back to a placeholder — see `Ops_Runbook.md §7.4`.
- **Rollback databases are the safety net** for anything an import drops; keep at least
  the two most recent.
