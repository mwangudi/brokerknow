# BrokerKnow Ops Runbook

> **Sensitive document.** Contains droplet credentials and SQL Server `sa` password.
> Keep out of public source-control mirrors. Rotate any value before sharing externally.

---

## 1. Droplet & environment topology

Single DigitalOcean droplet hosting both API instances, web bundles, and SQL Server.

| Item | Value |
| --- | --- |
| Host | `46.101.6.131` (Ubuntu 22.04, 2 vCPU / 4 GB, region `lon1`) |
| Hostname | `ubuntu-s-2vcpu-4gb-lon1` |
| SSH | `ssh root@46.101.6.131` (root login — key-based) |
| SQL Server | `localhost,1433` running as `mssql-server.service` |
| SQL `sa` password | `SpringfielD##88` |
| Process owner | `deploy` user (systemd units run as `deploy`) |

### API services

| Env | systemd unit | Port | Install root | Database |
| --- | --- | --- | --- | --- |
| **PROD** | `brokerknow-api.service` | `5260` | `/opt/brokerknow/api` | `BrokerKnow` |
| **TEST** | `brokerknow-api-test.service` | `5261` | `/opt/brokerknow-test/api` | `BrokerKnow_Test` |

Each service:

- `WorkingDirectory=/opt/brokerknow{,-test}/api`
- `ExecStart=/usr/bin/dotnet /opt/.../BrokerKnow.Api.dll`
- `User=deploy`, `Restart=always`, `Requires=mssql-server.service`
- Environment: `ASPNETCORE_ENVIRONMENT=Production`, `ASPNETCORE_URLS=http://127.0.0.1:526{0,1}`

Health probe: `curl http://localhost:5260/health` (200 means up).

### Web bundles (served by nginx)

| Surface | Web root |
| --- | --- |
| PROD client portal | `/var/www/portal` |
| PROD admin host | `/var/www/admin-host` |
| TEST client portal | `/var/www/test-portal` |
| TEST admin host | `/var/www/test-admin` |

Bundles are plain Vite output (`index.html` + `assets/index-*.js` + `assets/index-*.css`).
Cutover = copy `dist/*` over the web root and the new `index.html` will reference the
new hashed bundle filename.

---

## 2. Deployment

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

## 3. Sensitive PowerShell quoting traps

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

## 4. Useful one-liners

```powershell
# Recent API errors (PROD)
ssh root@46.101.6.131 "journalctl -u brokerknow-api --since '30 minutes ago' --no-pager | grep -iE 'error|exception|fail' | tail -40"

# Live tail
ssh root@46.101.6.131 "journalctl -u brokerknow-api -f"

# Run a SQL script against a database
C:\Windows\System32\OpenSSH\scp.exe -O .\query.sql root@46.101.6.131:/tmp/q.sql
ssh root@46.101.6.131 "sqlcmd -S localhost -U sa -P 'SpringfielD##88' -C -d BrokerKnow -W -i /tmp/q.sql"

# Re-check which bundle each surface is serving
ssh root@46.101.6.131 "grep -oE 'index-[A-Za-z0-9_-]+\.(js|css)' /var/www/portal/index.html"
```

---

## 5. Changes shipped — 2026-06-04

### 5.1 Receipts + banks dropdown empty (regression)

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

### 5.2 Add User → "Failed to save" (legacy IDENTITY column)

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

### 5.3 Users listing — newest first

Changed `GET /api/users` from `OrderBy(u => u.UserName)` to
`OrderByDescending(u => u.UserId)`. The legacy `Users` table has no
date-added column, but `UserID` is sequential so highest ID is effectively most
recently added.

### 5.4 Contract-note signature

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

### 5.5 Deploy artefacts

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
