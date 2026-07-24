# BrokerKnow — Code & Platform Audit (2026-07-24)

Scope: `brokerknow-api` (.NET 10), `brokerknow-web` / `brokerknow-portal` (React/Vite),
and the DigitalOcean droplet (SQL Server + nginx + 4 API instances). Advisory only —
findings are prioritized P1 (address soon) → P3 (worth doing). Verified against the
codebase on 2026-07-24 unless noted.

---

## 0. What's already solid (keep)
- **AuthZ**: RBAC v2 `[RequireArea]` / `[RequirePage]` filters guard the data controllers
  (the historical "anonymous-open" controllers are resolved). Auditors read-only filter,
  single-session enforcement, per-page overrides.
- **Secrets**: committed `appsettings.json` holds only placeholders (`REPLACE_WITH…`, empty
  password); real DB password / JWT key are set per-instance on the droplet and preserved
  across deploys.
- **Money hot-paths** (Payment / Order / Contract) use `LegacyKeys.NextIdAsync`
  (UPDLOCK/HOLDLOCK) inside an execution-strategy transaction.
- FK/PK/index hardening applied; structured logging (Serilog) + `UseSerilogRequestLogging`;
  EF `EnableRetryOnFailure`; penny-exact clean-schema migration pipeline; TLS on subdomains;
  raw SQL is static/parameterized (no injection — e.g. `LookupsController.LedgerBankAccounts`).

---

## P1 — Address soon (real risk)

### P1.1 Dual-write data loss (biggest architectural risk)
- **Finding**: the legacy **desktop app and the web app both write to production**, and each
  fresh desktop re-import (done regularly) *full-replaces* the business data, silently
  dropping web-only rows. Observed twice: web-created back-office logins (0713/0722) and, on
  0722, **2 web CDS contracts + 5 receipts**. `BK_Files/migrate_preserve_weblogins.sql` now
  carries web-created `dbo.Users`/`dbo.UserGroups` forward, but **web-created business data
  (Contracts/Lots/Payments/CDS) is still lost**.
- **Risk**: recurring, silent financial-data loss on every import; MAX+1 key collisions between
  the two writers (contract renumbering seen on 0722).
- **Recommendation**: choose **one system of record**. Either (a) make the web authoritative
  and retire desktop data entry, or (b) replace the full-replace import with a **delta/merge**
  and extend the "carry-forward web-only rows" step to business tables. Decision needed before
  the next import.

### P1.2 Backups are local-only  — ⏳ PENDING (scheduled for go-live, Aug 2026)
- **Finding**: nightly DB+filesystem backups (~7.4 GB, 11-day retention) live **only on the
  droplet**. `/etc/cron.d/brokerknow-backup` → `backup-nightly.sh` has an `rclone` hook that is
  not configured.
- **Risk**: droplet loss = total data loss.
- **Status**: DEFERRED to go-live by the client (still pre-go-live; parallel run week of
  2026-07-28, live Aug 2026). No separate server available — **decided approach** (plan in place,
  execute at go-live):
  1. Enable **DigitalOcean Droplet Backups** (console checkbox) — whole-droplet weekly, offsite,
     zero setup.
  2. Add a granular daily sync via the existing `rclone` hook to **Backblaze B2** (cheapest) or
     **Google Drive** (free 15 GB — holds ~3 weeks at ~670 MB/night). Storage choice + credentials
     to be provided at go-live; then wire `rclone config` + verify first upload.

### P1.3 No rate limiting on auth endpoints  — ✅ DONE (2026-07-24)
- **Finding**: `/auth/login`, `/auth/forgot-password`, `/auth/reset-password` have no rate
  limiting (OTP lockout at 5 attempts only partially helps, and only when OTP is enabled).
- **Risk**: credential brute-force / reset abuse.
- **Recommendation**: add ASP.NET rate limiting (fixed/sliding window) on `/auth/*`, and/or
  nginx `limit_req` + fail2ban. Self-contained code change.
- **Implemented**: ASP.NET `AddRateLimiter` "auth" policy — fixed window **30 requests/min
  per client IP** (X-Forwarded-For from nginx), 429 on exceed; `[EnableRateLimiting("auth")]`
  on login, login/verify-otp, forgot-password, reset-password. Verified on test (single login
  OK, burst throttled after 30, other IPs unaffected). Deployed all 4 instances.

---

## P2 — Should fix

### P2.4 MAX+1 key generation still raw in ~12 controllers  — ✅ DONE (2026-07-24)
- **Finding**: `(await db.X.MaxAsync(x => (int?)x.Dpa) ?? 0) + 1` without locking in
  AccountManagers (Owners), Agents, Brokers, Banks, Commissions, Groups, Holidays, Levies,
  Users, PortalUsers (`tbClient`), ChartOfAccounts, and **CdsTradeImports** (BatchId +
  materialize Order/OrdDetail — higher risk, EOD batch).
- **Risk**: concurrent inserts read the same MAX → duplicate key → PK rejects one → 500.
- **Recommendation**: route these through the existing `LegacyKeys.NextIdAsync`
  (`BrokerKnow.Application/Common/LegacyKeys.cs`) as the money paths do.
- **Implemented**: 16 create-paths across 14 controllers now take the key + INSERT in one
  `CreateExecutionStrategy` + `BeginTransactionAsync` + `LegacyKeys.NextIdAsync` (UPDLOCK)
  transaction — AccountManagers, Agents, Banks (Bank + BnkBranch), Brokers, Commissions,
  Groups, Holidays, Levies (Levy + LevySecurity), Securities, Users, PortalUsers (Client),
  Clients (direct), ChartOfAccounts (NominalAccount, 4-way id widened under the lock), and
  CdsTradeImports (BatchId + tbOrder/OrdDetail in both materialize paths).
- **Surfaced two latent bugs** (create was already broken on these, masked by the identity
  error): `Commission`, `Groups`, `Levy`, `LevySecurity` PKs are **SQL IDENTITY** on all four
  DBs (verified) — added `IDENTITY_INSERT` toggling (matching the existing `Users` path) so the
  explicit MAX+1 insert succeeds; and `Commission.SystemMaintained` (NOT NULL) was never set on
  create → now defaults to 0. Runtime-verified on test: create+delete of group/levy/commission
  (identity) and holiday/bank/branch (non-identity) all 201. Deployed all 4 instances.

### P2.5 Money represented as `float`  — ✅ DONE (2026-07-24)
- **Finding**: e.g. `CommissionsController.SaveCommissionRequest` mixes `float`
  (CommissionRate, UpperSecurityCommission, BondCommission) and `decimal` (boundaries,
  minimums).
- **Risk**: rounding drift in a financial system.
- **Recommendation**: standardize all rate/money fields on `decimal`; audit entities +
  request records for `float`/`double`.
- **Implemented**: fixed the actual drift in the C# booking math — every
  `(decimal)(rate / 100)` (float division, then cast) became `((decimal)rate / 100m)`
  (cast, then exact decimal division) across `CommissionCalculator` (3-tier bands + VAT) and
  `ContractService` (WHT, exchange, agent, handling-VAT, user-defined levies). The legacy
  `real` rate columns are shared with the desktop app + SQL procs, so DB storage types were
  left unchanged (a column migration is a separate, riskier change). Added
  `CommissionCalculatorTests` (first coverage of the money engine) and revived the whole test
  suite (stale `IEmailService`/`IPaymentService`/`INotificationService` fakes updated to
  current signatures; `LegacyKeys.NextIdAsync` given a non-relational fallback + the two
  transactional test contexts ignore the in-memory `TransactionIgnoredWarning`) → **46/46
  pass**. Deployed all 4 instances.

### P2.6 No global exception handler  — ✅ DONE (2026-07-24)
- **Finding**: `Program.cs` pipeline has no `UseExceptionHandler` / `AddProblemDetails`.
- **Risk**: inconsistent error responses; unhandled exceptions return bare 500s.
- **Recommendation**: add `AddProblemDetails()` + `UseExceptionHandler` for consistent,
  non-leaking error payloads.
- **Implemented**: `AddProblemDetails()` + `app.UseExceptionHandler()` (first in pipeline) →
  RFC 7807 ProblemDetails for unhandled exceptions. Smoke-tested (test then prod: ping, auth
  pipeline, auth guard, rate limiter, clean 404 — 0 failures) and deployed all 4 instances.

### P2.7 No monitoring / alerting
- **Finding**: issues surface via user complaints (e.g. the 0722 session bug). Serilog logs to
  file/journal only.
- **Recommendation**: add an error sink (Serilog → email/Slack, or Sentry) for proactive
  exception alerting; a simple uptime check on the 4 `/api/ping` ports.

---

## P3 — Worth doing
- **P3.8 Integration tests for the money/levy engine.** The suite fakes
  `PaymentService`/`ContractService`, so real key-gen, transactions, and the Malawi-specific
  levy math are untested. Add tests against a containerized SQL Server, per tenant.
- **P3.9 CI/CD.** Deploys are manual `scp` scripts (one has blanked a web root). A minimal
  build → test → guarded-deploy pipeline removes that footgun.
- **P3.10 JWT signing key per tenant.** Confirm the 4 instances don't share a key (per-DB user
  lookup mitigates cross-tenant use, but it's a defense-in-depth gap).
- **P3.11 Reduce `_DPA_` boilerplate.** Adopt the `app.*` friendly views more widely; factor
  the repeated reference-data CRUD (MAX+1 + validate + apply) into a shared base — also fixes
  P2.4 in one place.
- Weak default password `BrokerKnow@123` / standing `Passw0rd`: ensure `MustChangePassword`
  is set on all admin-created logins; consider a stronger default.

---

## Suggested execution order
1. **P1.3 rate limiting** — self-contained, can start immediately.
2. **P1.2 offsite backups** — quick once a bucket + credentials are provided.
3. **P1.1 dual-write** — needs a system-of-record decision, then design the merge/import change.
4. Then P2.6 (exception handler), P2.4 (key-gen), P2.5 (money types), P2.7 (alerting).
