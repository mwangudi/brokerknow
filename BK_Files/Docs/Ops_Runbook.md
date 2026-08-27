# BrokerKnow Ops Runbook

> **Sensitive document.** Contains droplet credentials and SQL Server `sa` password.
> Keep out of public source-control mirrors. Rotate any value before sharing externally.

---

## 1. Droplet & environment topology

> Verified against the live droplet on **2026-08-26**. Earlier revisions of this
> section described a two-service estate with a `BrokerKnow` / `BrokerKnow_Test`
> database pair — that has not been true since the multi-tenant split. The table
> below is read from `systemctl` and each instance's `appsettings.json`.

Single DigitalOcean droplet hosting **four** API instances, all web bundles, and SQL Server.

| Item | Value |
| --- | --- |
| Host | `46.101.6.131` (Ubuntu 22.04, 2 vCPU / 4 GB, region `lon1`) |
| Hostname | `ubuntu-s-2vcpu-4gb-lon1` |
| SSH | `ssh root@46.101.6.131` (root login — key-based) |
| SQL Server | SQL Server 2022 (RTM-CU24-GDR) on `localhost,1433`, unit `mssql-server.service` |
| SQL `sa` password | `SpringfielD##88` |
| sqlcmd | `/opt/mssql-tools18/bin/sqlcmd` — **requires `-C`** (trust self-signed cert) |
| Process owner | `deploy` user (systemd units run as `deploy`) |
| Disk | 77 GB volume, ~38% used (48 GB free) |

### API services

Each tenant is a **separate systemd unit with its own database**. The unit name is
the only reliable way to tell them apart — never assume a port from a hostname.

| Tenant | systemd unit | Port | Install root | Database |
| --- | --- | --- | --- | --- |
| **Malawi PROD** | `brokerknow-api.service` | `5260` | `/opt/brokerknow/api` | `axis_db_prod` |
| **Malawi TEST** | `brokerknow-api-test.service` | `5261` | `/opt/brokerknow-test/api` | `BrokerKnow_Malawi0612` |
| **Rwanda** | `brokerknow-api-rwanda.service` | `5262` | `/opt/brokerknow-rwanda/api` | `BrokerKnow_RW_Clean` |
| **Kenya** | `brokerknow-api-kenya.service` | `5264` | `/opt/brokerknow-kenya/api` | `BrokerKnow_KE_Clean` |

To confirm which database an instance is really using:

```bash
grep -oE 'Database=[^;]*' /opt/brokerknow/api/appsettings.json
```

Each service: `User=deploy`, `Restart=always`, `Requires=mssql-server.service`,
`ASPNETCORE_ENVIRONMENT=Production`.

Health probe: `curl http://127.0.0.1:5260/api/ping` (200 means up).

### Web bundles (served by nginx)

**Bundles are not interchangeable** — each is built against a different base path
and, for the agent surfaces, a different audience. Copying the wrong `dist` into a
root produces a blank page with 404s on the hashed assets.

| Surface | Web root | Build command (from `brokerknow-web/`) |
| --- | --- | --- |
| Admin host | `/var/www/admin-host` | `VITE_BASENAME=/ npm run build -- --base=/` |
| Admin (sub-path) | `/var/www/admin` | plain `npm run build` (clear `VITE_BASENAME`) |
| Cedar test admin | `/var/www/cedartest-admin` | plain `npm run build` |
| Agent host / test agent | `/var/www/agent-host`, `/var/www/test-agent` | `VITE_AUDIENCE=agent VITE_BASENAME=/agent npm run build -- --base=/agent/` |
| Cedar agent | `/var/www/cedaragent` | `VITE_AUDIENCE=agent VITE_BASENAME=/ npm run build -- --base=/` |
| Kenya admin | `/var/www/kenya-admin` | `npm run build -- --mode kenya --base=/ke/admin/` |
| Rwanda admin | `/var/www/rwanda-admin` | `npm run build -- --mode rwanda --base=/rw/admin/` |
| Rwanda test admin | `/var/www/rwandatest-admin` | `VITE_API_URL=/api VITE_BASENAME=/admin npm run build -- --mode rwanda --base=/admin/` |
| Client portal | `/var/www/portal` | from `brokerknow-portal/` |
| Kenya / Rwanda portal | `/var/www/kenya-portal`, `/var/www/rwanda-portal` | portal, `--mode <tenant> --base=/ke/` or `/rw/` |

Public hostnames: `cedarcapital` (admin), `cedarclient` (portal + `/agent/`),
`cedartest` (`/admin/`), `ke.martensafrica.com`, `rwandatest.martensafrica.com`.

> **`npm run build` is `tsc -b && vite build`.** A TypeScript error short-circuits the
> chain, `vite` never runs, and the **previous** `dist` is still on disk — so a deploy
> will silently ship a stale bundle. Always confirm `✓ built` before tarring.

---

## 2. Access — getting in from another machine

> Secret values live in `BK_Files/Docs/Access_Credentials.local.md`, which is
> git-ignored. This section deliberately contains no passwords or key material.
> Verified against the droplet **2026-08-26**.

### SSH

```bash
ssh root@46.101.6.131
```

| | |
| --- | --- |
| Auth | **Public key only** — `PasswordAuthentication no`, `PermitRootLogin yes` |
| Port | 22 (open in `ufw` as profile `OpenSSH`) |
| Key | ED25519, fingerprint `SHA256:C489SlXgJOOx/mQSb3OM+OO+ICIiqdeXabzPi8lfs+U`, comment `michael.wangudi@patasoko.co.ke` |
| Local path | `~/.ssh/id_ed25519` |
| Other accounts | `deploy` (uid 1000) owns the systemd units; not used interactively |

> ### ⚠ There is exactly one key on this droplet
>
> `/root/.ssh/authorized_keys` holds **a single entry**, and password
> authentication is off. If that private key is lost, SSH access is gone for
> good — the only way back is the DigitalOcean web console (password reset or
> recovery ISO).
>
> **Before working from a new machine, do one of these while you still have
> access:**
>
> ```bash
> # Option A — authorise a second key (preferred; keeps machines independent)
> ssh-keygen -t ed25519 -C "laptop-2" -f ~/.ssh/id_ed25519_bk2
> ssh root@46.101.6.131 "cat >> /root/.ssh/authorized_keys" < ~/.ssh/id_ed25519_bk2.pub
> ssh -i ~/.ssh/id_ed25519_bk2 root@46.101.6.131 whoami   # verify BEFORE relying on it
> ```
>
> ```powershell
> # Option B — copy the existing private key to the new machine
> #   Windows: %USERPROFILE%\.ssh\id_ed25519   (also copy the .pub)
> #   chmod 600 on Linux/macOS, or Windows will refuse it as "too open"
> ```
>
> An unused key `id_ed25519_do` (comment `do-brokerknow`) exists locally but is
> **not** installed on the droplet — it is a ready-made candidate for Option A.

### Reaching the services

Everything except HTTP/HTTPS is bound to loopback, so use an SSH tunnel rather
than opening firewall ports.

| Service | Bind | Reach it from your machine |
| --- | --- | --- |
| API 5260 / 5261 / 5262 / 5264 | `127.0.0.1` | `ssh -L 5260:127.0.0.1:5260 root@46.101.6.131` |
| SQL Server 1433 | `0.0.0.0`, **blocked by `ufw`** | `ssh -L 1433:127.0.0.1:1433 root@46.101.6.131` then connect SSMS to `localhost,1433` |
| nginx 80 / 443 | public | the hostnames below |

`ufw` is active and allows only `OpenSSH` and `Nginx Full`. Default incoming is
deny. **The firewall is the only thing keeping SQL Server off the internet** —
port 1433 itself listens on all interfaces.

### Public hostnames

| Host | Serves |
| --- | --- |
| `cedarcapital.martensafrica.com` | Malawi admin |
| `cedarclient.martensafrica.com` | Malawi client portal, `/agent/` for agents |
| `cedartest.martensafrica.com` | Test admin at `/admin/` |
| `ke.martensafrica.com` | Kenya |
| `rwandatest.martensafrica.com` | Rwanda test |
| `booklab.localinvestors.co.ke`, `pharma.localinvestors.co.ke` | unrelated tenants on the same box |

TLS is Let's Encrypt, renewed by `certbot.timer` (twice daily, healthy). nginx
configs are in `/etc/nginx/sites-enabled/` — the BrokerKnow surfaces are all in
the single `brokerknow` file.

> The box also runs MySQL, MailHog and two Python services on ports 4000/4100 for
> unrelated tenants. Don't assume a process on this droplet is ours.

### Database access

```bash
/opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P '<pwd>' -C -d axis_db_prod -Q "SELECT 1"
```

`-C` (trust the self-signed certificate) is **required** — without it every
connection fails. `sa` is the only SQL login; all four API instances use it.

### Backups

| | |
| --- | --- |
| Schedule | Full nightly 02:00, transaction log every 15 min (`/etc/cron.d/brokerknow-backup`) |
| Scripts | `/usr/local/sbin/backup-nightly.sh`, `/usr/local/sbin/backup-log.sh` |
| Destination | `/var/backups/brokerknow/{sql,fs}` — currently ~8.3 GB |
| Retention | 14 days, rotated locally |
| Log | `/var/log/brokerknow-backup.log` |

Verify the last run with:

```bash
grep -c 'backup-nightly start' /var/log/brokerknow-backup.log
tail -5 /var/log/brokerknow-backup.log
```

> ### ⚠ Backups are not leaving the droplet
>
> `backup-nightly.sh` supports an offsite `rclone sync` to DigitalOcean Spaces,
> but **rclone is not installed and no `spaces:` remote is configured**, so every
> night logs `Offsite SKIPPED`. Backups therefore sit on the same disk as the
> data they protect — losing the droplet loses both. Fixing it is roughly:
>
> ```bash
> apt-get install -y rclone
> rclone config       # create an S3-compatible remote literally named "spaces"
> /usr/local/sbin/backup-nightly.sh   # re-run; the log should show "Offsite sync ->"
> ```

### Recovering a database

Full backups are ordinary `.bak` files:

```bash
ls -t /var/backups/brokerknow/sql/axis_db_prod_FULL_*.bak | head -3
```

Restore with `MOVE` to a new name rather than over the live database, then rename
in — the same pattern the refresh cutover uses (§6).

---

## 3. Deployment

### Web (portal / admin)

```powershell
# from brokerknow-web/
npm run build

# upload (preserves all hashed asset files)
C:\Windows\System32\OpenSSH\scp.exe -O -r .\dist\* root@46.101.6.131:/var/www/portal/

# mirror PROD → TEST web (or build twice with VITE env if they diverge)
ssh root@46.101.6.131 "cp -r /var/www/portal/* /var/www/test-portal/"
```

Verify the new hashed bundle is referenced:

```powershell
ssh root@46.101.6.131 "grep -oE 'index-[A-Za-z0-9_-]+\.(js|css)' /var/www/portal/index.html"
```

### API (.NET 10)

Use the helper script `tmp/deploy_api.sh` (commit if it isn't already). It backs up the
existing install, swaps in the new tarball, **preserves the existing
`appsettings.json`** (which differs per environment — TEST/PROD have different
connection strings) and restarts the unit.

```powershell
# from brokerknow-api/
dotnet publish src/BrokerKnow.Api/BrokerKnow.Api.csproj -c Release -o publish-out

$stamp = Get-Date -Format yyyyMMddHHmmss
tar -czf "api-publish-$stamp.tgz" -C publish-out .

C:\Windows\System32\OpenSSH\scp.exe -O "api-publish-$stamp.tgz" `
    "root@46.101.6.131:/tmp/api-publish-$stamp.tgz"
C:\Windows\System32\OpenSSH\scp.exe -O ..\tmp\deploy_api.sh root@46.101.6.131:/tmp/deploy_api.sh
ssh root@46.101.6.131 "chmod +x /tmp/deploy_api.sh; /tmp/deploy_api.sh $stamp"
```

The script iterates over `test` then `prod`:

1. Snapshots the current install to `/opt/.../api.bak-<stamp>`.
2. Wipes `/opt/.../api/`.
3. Extracts the tarball.
4. Restores the saved `appsettings.json`.
5. Restarts the service and reports `is-active` + `health=`.

**Rollback** = stop the service, restore from `/opt/.../api.bak-<stamp>`, start.

### Critical gotcha: `appsettings.json`

The published tarball contains the developer's `appsettings.json`, which points at
`(localdb)\MSSQLLocalDB`. Overwriting the droplet's file with that **crashes the
service immediately** with `LocalDB is not supported on this platform`. The deploy
script handles this. **Don't deploy without it.**

---

## 4. Sensitive PowerShell quoting traps

Repeatedly bit us today:

- `$(...)` and `(Get-Content ...)` inside an `ssh "..."` double-quoted argument get
  pre-evaluated by PowerShell. Compute timestamps in PS first
  (`$stamp = Get-Date -Format yyyyMMddHHmmss`) and interpolate as `$stamp`.
- Bash variables like `$ROOT` or `$ENV` (the latter is also a PS automatic var) get
  evaluated by PS, not bash. Either:
  - Single-quote the ssh payload, or
  - Write the script to a file and `scp` + `chmod +x` + run remotely (preferred for
    anything non-trivial).
- Each `run_in_terminal` opens a fresh shell. `Set-Location` and env vars do **not**
  persist between calls. Always prefix with `Set-Location 'C:\...'` and chain with `;`.
- `scp` from PowerShell sometimes resolves to nothing — use
  `C:\Windows\System32\OpenSSH\scp.exe -O ...` to be safe.

---

## 5. Useful one-liners

```powershell
# Recent API errors (PROD)
ssh root@46.101.6.131 "journalctl -u brokerknow-api --since '30 minutes ago' --no-pager | grep -iE 'error|exception|fail' | tail -40"

# Live tail
ssh root@46.101.6.131 "journalctl -u brokerknow-api -f"

# Run a SQL script against a database (note: -C is required, and the full sqlcmd path)
C:\Windows\System32\OpenSSH\scp.exe -O .\query.sql root@46.101.6.131:/tmp/q.sql
ssh root@46.101.6.131 "/opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P 'SpringfielD##88' -C -d axis_db_prod -W -i /tmp/q.sql"

# Re-check which bundle each surface is serving
ssh root@46.101.6.131 "grep -oE 'index-[A-Za-z0-9_-]+\.(js|css)' /var/www/portal/index.html"
```

---

## 6. Malawi production data refresh (recurring)

Cedar send a fresh `.rar` of the legacy Malawi database every week or two. The refresh
**replaces `axis_db_prod` wholesale**, so anything that lives only in the app has to be
put back afterwards. Ten refreshes done to date; #10 was 2026-08-26.

### The pipeline

Driver scripts live in `/tmp` on the droplet and are **cleared on reboot** — re-derive
them from the previous run rather than writing new ones:

```bash
sed 's/0818/0826/g; s/180826/260826/g' /tmp/ml0818_<step>.sh > /tmp/ml0826_<step>.sh
```

Steps, in order: `extract` → restore → `user_diff` → `build` → `smoke` → `delta` →
`cutover` → post-cutover fixes → `live_verify` → `cleanup`.

Supporting SQL lives in `/tmp/bk/` (`graft_app_layer_*.sql`, `migrate_*.sql`,
`schema_baseline_BrokerKnow.sql`, `fix_swapped_bankacc_name_number.sql`).

### Gates that must pass before cutover

| Gate | Expected |
| --- | --- |
| Reconcile | 162 tables, 0 real mismatches, 2 intentional skips |
| Money | Payment / ClientBalances / Lot gross / LevyContract all **penny-exact** |
| Foreign keys | 22 trusted, 0 violations (23 once the app has started — it adds one) |
| Smoke on a throwaway instance | 20 endpoints 200, contract-note PDF valid |
| Delta vs live | nothing live-only that matters (see below) |

The cutover script re-checks logins / staff / admins / trusted FKs / client count and
**aborts** if any guard fails. It renames `axis_db_prod` → `axis_db_prod_pre<NNNN>`
(the rollback) and the freshly built database into `axis_db_prod`, so `appsettings.json`
never changes.

**Rollback:** stop `brokerknow-api`, rename `axis_db_prod` out of the way, rename
`axis_db_prod_pre<NNNN>` back to `axis_db_prod`, start the service.

### Things the import destroys — restore them every time

1. **`ClientCDSNo` is wiped.** Since refresh #10 the numbers are **copied across from
   the pre-cutover database**, not re-derived from the register spreadsheet:

   ```sql
   SELECT Client_DPA_, ClientCDSNo INTO dbo.ClientCDS_backup_<NNNN>
   FROM axis_db_prod_pre<NNNN>.dbo.Client WHERE ISNULL(ClientCDSNo,'') <> '';
   -- then UPDATE ... JOIN inside a TRAN, rolling back unless updated = expected
   ```

   Re-running the matcher instead would silently discard any correction the team has
   made in the app since the last refresh.

2. **`BankAcc` name/number come back swapped.** Re-run
   `/tmp/bk/fix_swapped_bankacc_name_number.sql` (~80 rows, takes its own snapshot).

### Known, accepted findings

- **Lots that exist only on live.** A handful of contracts are present in the legacy
  dump while their `Lot` rows are not, so those contracts lose their trade detail at
  cutover (17 at refresh #9, 5 at #10, 4 of which have survived several refreshes).
  They remain in the rollback database. Root cause not yet established.
- **Staff passwords are preserved**, not reset: the app-layer graft copies `PortalUsers`
  verbatim. The user diff only reports genuinely new/removed staff. Verify with
  `lesc`/`Passw0rd` → 200 and `webadmin`/`Passw0rd` → 401.
- `/users/online` → 403 and `/reports/clients` → 404 in the smoke output are script
  artefacts (RBAC gate and a wrong endpoint name), not regressions.

### CDS numbers from the Reserve Bank register

The August 2026 client list is a **Reserve Bank of Malawi "Client Accounts Report"**,
not Cedar's own listing: header on **row 5**, no `Client Code` column, `Client UID` =
`PP` + identity document. Matching therefore runs on document and name, tiered by
strength of evidence (`BK_Files/match_rbm_csd.py`) — exact-document and
near-identical-document matches are applied, name-only matches where the documents
disagree are withheld for a human. See `System_Status.md` for current coverage.

> Generated apply SQL and the client CSVs contain plaintext ID/passport numbers and are
> **git-ignored** (`/BK_Files/apply_cedar_csd*.sql`, `/BK_Files/CDS Numbers To *.csv`).

---

## 7. Changes shipped — 2026-06-04

### 7.1 Receipts + banks dropdown empty (regression)

**Symptom.** "Banks do not pull anymore" + "Failed to save receipt" on the
Add Receipt page.

**Root cause.** Commit `b59590a` (G2 "internal-only banks") added
`?internalOnly=true` to the `useLookup` calls in
`brokerknow-web/src/pages/Payments/AddReceipt.tsx`. The legacy `BankAcc` table
has *zero* rows with `Client_DPA_ IS NULL` (all 5,267 accounts belong to clients),
so the lookup returned an empty list — dropdown empty, can't pick a bank, can't
save. Verified by `SELECT COUNT(*) FROM dbo.BankAcc WHERE Client_DPA_ IS NULL` = 0.

**Fix.** Reverted both `useLookup` calls to unfiltered:

```ts
// brokerknow-web/src/pages/Payments/AddReceipt.tsx
const { data: banks } = useLookup("banks");
const { data: bankAccounts } = useLookup("bank-accounts") as { ... };
```

(matches what `AddPayment.tsx` already does)

**Deployed.** Web bundle `index-Cz8CAVAA.js` to `/var/www/portal` and
`/var/www/test-portal`. No API change.

### 7.2 Add User → "Failed to save" (legacy IDENTITY column)

**Symptom.** Creating a user from the Users page failed silently with
"Failed to save."

**Root cause.** On the Cedar Malawi database, `dbo.Users.UserID` is an
`IDENTITY` column. Our EF config (`SecurityConfigurations.AppUserConfiguration`)
uses `ValueGeneratedNever` and computes IDs as `MAX(UserId) + 1` for legacy
stored-proc parity. SQL Server rejected the insert with
`Cannot insert explicit value for identity column in table 'Users' when
IDENTITY_INSERT is set to OFF` (error 544).

**Fix.** Wrapped the insert in `SET IDENTITY_INSERT dbo.Users ON/OFF` inside the
EF execution strategy in
[brokerknow-api/src/BrokerKnow.Api/Controllers/UsersController.cs](../../brokerknow-api/src/BrokerKnow.Api/Controllers/UsersController.cs)
(`Create` action). Harmless on databases where `UserID` is *not* an identity
column (Cedar dev box) and necessary on legacy production-shape databases.

Verified end-to-end with `POST /api/users` → 201 on both TEST and PROD.

### 7.3 Users listing — newest first

Changed `GET /api/users` from `OrderBy(u => u.UserName)` to
`OrderByDescending(u => u.UserId)`. The legacy `Users` table has no
date-added column, but `UserID` is sequential so highest ID is effectively most
recently added.

### 7.4 Contract-note signature

PM provided a scanned signature; saved to
`brokerknow-api/src/BrokerKnow.Api/wwwroot/signature.png` (13.6 KB).

The F3 PDF code (`BrokerKnow.Reports/Contracts/ContractNotePdf.cs` +
`ContractsController.LoadSignatureImage()`) already looked for that path; the
"[Signature pending]" placeholder is now replaced by the actual signature on
every contract note. Deployed to both TEST and PROD wwwroot.

> **Real person's signature — never commit or push it.** This is a live
> signature provided by the PM. It is deliberately **kept out of git** and lives
> only (a) locally on the developer machine, (b) in this docs folder, and
> (c) on the droplet's API wwwroot. Both are git-ignored so an accidental
> `git add -A` cannot stage them:
>
> - Root repo `.gitignore`: `/BK_Files/Docs/signature.png`, `/BK_Files/Docs/Gina's Signature .pdf`
> - API submodule `.gitignore`: `src/BrokerKnow.Api/wwwroot/signature.png` (+ `.jpg`)
>
> **Reference copies kept in docs** (this folder, also git-ignored):
> - `BK_Files/Docs/Gina's Signature .pdf` — the original source the PM supplied.
> - `BK_Files/Docs/signature.png` — a byte-for-byte copy of the deployed asset
>   (SHA-256 `f3b1516a…b940ef9`).
>
> **Loader resolution order** (`ContractsController.LoadSignatureImage()`):
> 1. env var `BROKERKNOW_SIGNATURE_PATH`
> 2. `wwwroot/signature.png`
> 3. `wwwroot/signature.jpg`
>
> **Fresh-deploy step (manual).** Because the PNG is not in git, a clean
> clone + `dotnet publish` ships **without** it and contract notes silently fall
> back to the "[Signature pending]" placeholder. After any fresh install, copy
> the reference PNG back into the wwwroot before/with the deploy, e.g.:
>
> ```powershell
> # restore the signature onto a freshly-published tree before tarballing
> Copy-Item 'BK_Files\Docs\signature.png' `
>   'brokerknow-api\src\BrokerKnow.Api\wwwroot\signature.png' -Force
> ```
>
> Or set `BROKERKNOW_SIGNATURE_PATH` in the systemd unit to a stable path on the
> droplet that the deploy never wipes.

### 7.5 Deploy artefacts

| Build stamp | Tarball |
| --- | --- |
| API (initial signature + IDENTITY fix) | `api-publish-20260604111510.tgz` |
| API (Users listing sort) | `api-publish-20260604112836.tgz` |

Backups on droplet:

- `/opt/brokerknow/api.bak-20260604111510`
- `/opt/brokerknow/api.bak-20260604112836`
- `/opt/brokerknow-test/api.bak-20260604111510`
- `/opt/brokerknow-test/api.bak-20260604112836`

All services confirmed `active` and `health=200` after each cutover.
