# BrokerKnow — Production Hosting Runbook

**Droplet:** `46.101.6.131` (DigitalOcean)
**OS:** Ubuntu (current LTS) · **DB:** SQL Server 2022 · **Runtime:** .NET 10
**Last updated:** June 2026

## 1. Hostname plan

| Hostname | Purpose | Backend |
|---|---|---|
| `cedarcapital.mw` / `www.cedarcapital.mw` | Marketing site (rebuild — see §6) | TBD |
| `portal.cedarcapital.mw` | Client portal SPA + `/api/` | `:5260` (prod API) |
| `admin.cedarcapital.mw` | Back-office SPA + `/api/` | `:5260` (prod API) |
| `test-portal.cedarcapital.mw` | Test portal | `:5261` (test API) |
| `test-admin.cedarcapital.mw` | Test admin | `:5261` (test API) |

Bare-IP access (`http://46.101.6.131/`) keeps working during DNS rollout — the
default nginx server still serves portal at `/` and admin at `/admin/` against
the prod API.

## 2. DNS

Create at the registrar (or DigitalOcean DNS):

```
portal       A 46.101.6.131
admin        A 46.101.6.131
test-portal  A 46.101.6.131
test-admin   A 46.101.6.131
@            A <marketing-host-IP>   # only when marketing site is ready
www          A <marketing-host-IP>
```

TTL: 300 during cutover, raise to 3600 once stable.

## 3. TLS

After DNS resolves, on the droplet:

```bash
ssh root@46.101.6.131
bash /tmp/issue-certs.sh ops@cedarcapital.mw
# Add --staging on the end to dry-run against LE staging first.
```

`certbot.timer` handles renewals; verify with `certbot renew --dry-run`.

## 4. Services

| Unit | Port | DB | Uploads root |
|---|---|---|---|
| `brokerknow-api.service` | `127.0.0.1:5260` | `BrokerKnow` | `/opt/brokerknow/uploads/` |
| `brokerknow-api-test.service` | `127.0.0.1:5261` | `BrokerKnow_Test` | `/opt/brokerknow-test/uploads/` |
| `mssql-server.service` | `127.0.0.1:1433` (bound localhost) | — | — |
| `nginx.service` | `80`, `443` | — | — |

Useful commands:

```bash
systemctl status brokerknow-api brokerknow-api-test
journalctl -u brokerknow-api -f
journalctl -u brokerknow-api-test -f
```

## 5. Backups

Installed by `ops/install-backups.sh`. Layout:

| What | When | Where | Retention |
|---|---|---|---|
| SQL FULL (all user DBs) | Daily 02:00 | `/var/backups/brokerknow/sql/*_FULL_*.bak` | 14 d |
| SQL LOG | Every 15 min | `/var/backups/brokerknow/sql/*_LOG_*.trn` | 2 d |
| Uploads + configs | Daily 02:00 | `/var/backups/brokerknow/fs/files_*.tar.gz` | 14 d |
| Log file | — | `/var/log/brokerknow-backup.log` | 8 w (logrotate) |

Both user DBs are in **FULL** recovery model → point-in-time recovery via
`RESTORE LOG ... WITH STOPAT`.

**Offsite (DO Spaces)** — credentials not configured yet. To enable:

```bash
apt-get install -y rclone
rclone config           # create remote name "spaces", type "s3", provider DigitalOcean
# Region: fra1 (or whichever Space you create). Bucket: brokerknow-backups.
```

The nightly script auto-detects the `spaces:` remote and syncs after the
local backup completes.

**Monthly restore drill** (do this — backups you never restore aren't backups):

```bash
# Restore the latest prod FULL into a scratch DB and run a smoke query
LATEST=$(ls -t /var/backups/brokerknow/sql/BrokerKnow_FULL_*.bak | head -1)
sqlcmd -S localhost -U sa -P '***' -C -Q "
RESTORE FILELISTONLY FROM DISK = N'$LATEST';
"
# then RESTORE DATABASE [BrokerKnow_Restore] FROM DISK = '$LATEST' WITH MOVE ...
```

## 6. Marketing site

The marketing site rebuild plan and content inventory live in
`ops/marketing-site-spec.md`. Decide between:

1. **Static rebuild** in this repo (Next.js export, deploys to `/var/www/marketing` on the droplet) — cheapest, no extra infra.
2. **Hand off to designer** (Webflow / Framer) — they host, we just point DNS.

Recommendation: option 1 if we want full design control; option 2 if marketing
copy will change frequently and we don't want to be in the loop for every edit.

## 7. Deploys

API:

```powershell
cd c:\Users\v-mwangudi\source\repos\BrokerKnow\brokerknow-api\src\BrokerKnow.Api
dotnet publish -c Release -o publish
$ts = Get-Date -Format yyyyMMddHHmmss
cd ..\..\..\..
tar -czf api-publish.tgz -C brokerknow-api\src\BrokerKnow.Api\publish .
scp -O api-publish.tgz "root@46.101.6.131:/tmp/api-publish-$ts.tgz"
# then on droplet: stop, replace, preserve appsettings.json, start, /health
```

SPAs are built locally with `VITE_API_URL=/api`, then `tar`/`scp` to
`/var/www/admin-host` (admin), `/var/www/portal` (portal),
`/var/www/test-admin`, `/var/www/test-portal`.

## 8. Security checklist before go-live

- [ ] Rotate `sa` password; move connection string to `User Id=brokerknow_app` (least-priv).
- [ ] UFW: allow `22/80/443` only (`ufw status` should show that).
- [ ] Fail2ban on `sshd`.
- [ ] `admin.cedarcapital.mw` IP allow-list in nginx (see commented block in `ops/nginx-brokerknow.conf`).
- [ ] `appsettings.json` permission `0600 deploy:deploy`.
- [ ] JWT signing key rotated from any dev default; stored only in `appsettings.json`.
- [ ] Disable MailHog `/mail/` route on the prod default_server, or restrict by IP.
- [ ] DO Spaces offsite backup live.
- [ ] At least one restore drill passed.
