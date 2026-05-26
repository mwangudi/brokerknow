# BrokerKnow

Modernisation of the legacy Cedar Capital BrokerKnow stockbroking system for the Malawi Stock Exchange. The umbrella repo hosts the working tree for the .NET 10 API, the React admin web, the client portal, and the legacy artefacts kept for reference.

## Layout

| Path | Purpose |
| --- | --- |
| `brokerknow-api/` | .NET 10 Web API + EF Core 9 (SQL Server). Domain / Application / Infrastructure / Reports / Api projects under `src/`. |
| `brokerknow-web/` | React 19 + Vite + Tailwind admin (TailAdmin fork). Operations, accounts, reporting and admin screens. |
| `brokerknow-portal/` | React + Vite client portal (statements, orders, profile). |
| `BK_Files/` | Reference material — legacy ASP source, schema dumps, demo CSVs, generated PDFs, manuals. |
| `tools/` | Local ops helpers. |
| `smoke_brokers_agents.ps1` | Quick smoke-test script for Brokers + Agents endpoints. |

> `brokerknow-api/` and `brokerknow-web/` are committed as **gitlinks**; clone with `git clone --recurse-submodules` once their remotes are configured. Until then, work in the existing local checkouts.

## Stack

- **API** — .NET 10, EF Core 9, SQL Server 2022, JWT auth, Serilog, QuestPDF for Contract Notes / statements, MailKit for email.
- **Web** — React 19, Vite 6, Tailwind 3, SWR, Axios, React Router 7.
- **Database** — Legacy schema (`Owner`, `Client`, `Contract`, `Lot`, `Order`, `LevyContract`, `Journal`, `JournalEntry`, …) mapped 1:1 in EF; reference SQL & SP definitions under `BK_Files/Legacy_System/`.

## Development

### API
```powershell
cd brokerknow-api/src/BrokerKnow.Api
dotnet run
```
Listens on `http://localhost:5260`. Connection string in `appsettings.Development.json` (default targets local SQL Server).

### Web
```powershell
cd brokerknow-web
npm install
npm run dev
```
Vite dev server at `http://localhost:5173`. Mounts under the `/admin` basename.

### Portal
```powershell
cd brokerknow-portal
npm install
npm run dev
```

## Deployment

The API runs on a DigitalOcean droplet as `systemd` unit `brokerknow-api` on port `5260` (Nginx in front, SQL Server 2022 local). Publish flow:

```powershell
cd brokerknow-api/src/BrokerKnow.Api
dotnet publish -c Release -o publish
cd publish; tar -czf ..\api-publish.tgz *
scp -O ..\api-publish.tgz root@<droplet>:/tmp/
ssh root@<droplet> '<heredoc that snapshots /opt/brokerknow, preserves appsettings.json, untars, restarts, healthchecks :5260/health>'
```

The droplet's production `appsettings.json` is preserved across deploys — never overwrite it with the dev copy from `publish/`.

## Branches

- `main` — historical/legacy baseline.
- `develop` — integration trunk; deploys to droplet.
- `feature/*` — short-lived feature branches off `develop`; merge back with `--no-ff`.

## Conventions

- Soft-delete via `Deleted` flag on most tables (no `Deleted` column on `Client`).
- Contract numbers live on `Lot.ContractNumber`, not `Contract`.
- Journal amounts use `JournalEntryDebit` / `JournalEntryCredit`.
- All commission/levy money in **MWK**; PDF wording must match legacy Cedar Capital Contract Note format.

## Status

See `BK_Files/Docs/` for feature docs (Levy_Setup_And_Usage, Order_Lifecycle, Payments_And_Journals, Client_Portal_Registration). New features live on `feature/*` branches and are recorded in the changelog of merged commits.
