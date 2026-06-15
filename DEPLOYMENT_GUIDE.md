# Account-Opening Workflow - Deployment Guide

## Deployment Architecture

### 3-Environment Setup
```
cedartest.martensafrica.com  ─── API (Test) ──┐
                                                ├─→ Test: Full workflow testing
cedar.martensafrica.com      ─── API (Prod) ──┤
                                                └─→ Prod: Live customer data
cedarclient.martensafrica.com ── Portal ──────→ (connects to Cedar Prod API)
```

## Pre-Deployment Checklist

### Code Readiness
- [ ] All changes committed to `develop` branch
- [ ] Both API and Web builds successful locally
- [ ] No new compilation warnings or errors
- [ ] Git status clean: `git status --short` shows no uncommitted files

### Environment Readiness
- [ ] Both API servers have .NET 10 runtime installed
- [ ] Portal server has Node.js 18+ installed
- [ ] QuestPDF NuGet package configured (no license key needed for v5.x)
- [ ] Attachments storage paths exist and are writeable on both API servers
- [ ] Database migrations can be applied (SQL Server access)

### Database Prerequisites
- [ ] Blantyre branch exists in both cedartest and cedar databases
- [ ] Cedar Capital Malawi tenant configured
- [ ] User roles include: Officer, RiskOfficer, Supervisor

---

## Deployment Procedure

### Step 1: Build Artifacts

#### Build API
```powershell
cd c:\Users\v-mwangudi\source\repos\BrokerKnow\brokerknow-api
dotnet clean
dotnet build BrokerKnow.slnx -c Release
dotnet publish -c Release -o publish-release
```
**Output**: `brokerknow-api/publish-release/` folder ready for deployment

#### Build Portal
```powershell
cd c:\Users\v-mwangudi\source\repos\BrokerKnow\brokerknow-web
npm run build
```
**Output**: `brokerknow-web/dist-host/` folder ready for deployment

---

### Step 2: Deploy Cedartest API

**Target**: `https://cedartest.martensafrica.com`

#### 2a. Prepare Deployment Package
```powershell
# Copy release package to deployment server
# Destination: /var/www/api-test/ or similar on server

scp -r publish-release/* user@cedartest.martensafrica.com:/path/to/deployment/
```

#### 2b. Stop Current Service
```bash
# On cedartest server
sudo systemctl stop brokerknow-api-test
```

#### 2c. Deploy New Build
```bash
# On cedartest server
cd /path/to/deployment
sudo cp -r . /opt/brokerknow-api-test/
sudo chown -R www-data:www-data /opt/brokerknow-api-test
```

#### 2d. Update Configuration (if needed)
Edit `/opt/brokerknow-api-test/appsettings.json`:
```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=cedartest-db;Database=BrokerKnow_Test;..."
  },
  "Branding": {
    "CompanyName": "Cedar Capital Malawi"
  }
}
```

#### 2e. Apply Database Migrations
```bash
# On cedartest server with access to SQL Server
cd /opt/brokerknow-api-test
dotnet BrokerKnow.Api.dll migrate
```

#### 2f. Start Service
```bash
sudo systemctl start brokerknow-api-test
sudo systemctl status brokerknow-api-test
```

#### 2g. Verify Deployment
```bash
curl -s https://cedartest.martensafrica.com/api/health
# Expected: 200 OK response
```

---

### Step 3: Deploy Cedar (Prod) API

**Target**: `https://cedar.martensafrica.com`

Follow the same procedure as Step 2, but:
- Use `cedar` server credentials instead of `cedartest`
- Deployment path: `/opt/brokerknow-api-prod/`
- Database connection: Production database
- Service name: `brokerknow-api-prod`

#### Verification
```bash
curl -s https://cedar.martensafrica.com/api/health
# Expected: 200 OK response
```

---

### Step 4: Deploy Cedarclient Portal

**Target**: `https://cedarclient.martensafrica.com`

#### 4a. Prepare Portal Build
```powershell
cd c:\Users\v-mwangudi\source\repos\BrokerKnow\brokerknow-web
npm run build
```

#### 4b. Upload to Server
```powershell
# Copy built files to portal server
scp -r dist-host/* user@cedarclient.martensafrica.com:/path/to/deployment/
```

#### 4c. Update Environment Configuration
Edit portal `.env` or `config.json` to point to **prod API**:
```json
{
  "VITE_API_BASE_URL": "https://cedar.martensafrica.com",
  "VITE_API_TIMEOUT": "30000"
}
```

#### 4d. Deploy Portal Files
```bash
# On cedarclient portal server
cd /path/to/deployment
sudo cp -r dist-host/* /var/www/html/cedarclient/
sudo chown -R www-data:www-data /var/www/html/cedarclient/
sudo systemctl restart nginx
# or apache, depending on web server
```

#### 4e. Verify Portal
```bash
curl -s https://cedarclient.martensafrica.com
# Expected: HTML response (portal home page)
```

---

## Post-Deployment Validation

### Health Checks

```bash
# API Health Check - Cedartest
curl -i https://cedartest.martensafrica.com/api/health

# API Health Check - Cedar (Prod)
curl -i https://cedar.martensafrica.com/api/health

# Portal Connectivity Check
curl -I https://cedarclient.martensafrica.com
```

### API Smoke Tests (via curl or Postman)

```bash
# 1. List clients (verify API is working)
curl -H "Authorization: Bearer <token>" \
  https://cedar.martensafrica.com/api/clients

# 2. Get specific client account-opening workflow
curl -H "Authorization: Bearer <token>" \
  https://cedar.martensafrica.com/api/clients/{clientId}/account-opening
```

### Portal Smoke Tests (via browser)

1. Open `https://cedarclient.martensafrica.com`
2. Login with test credentials
3. Navigate to a client → "Account Opening" tab
4. Verify form sections load correctly
5. Test Save Draft functionality
6. Verify API communication (check Network tab in browser DevTools)

---

## Rollback Procedure

If issues detected after deployment:

### Immediate Rollback (Cedartest)
```bash
# On cedartest server
sudo systemctl stop brokerknow-api-test

# Restore previous version
sudo cp -r /opt/brokerknow-api-test-backup/* /opt/brokerknow-api-test/
sudo chown -R www-data:www-data /opt/brokerknow-api-test

# Restart with old version
sudo systemctl start brokerknow-api-test
```

### Immediate Rollback (Cedar Prod)
```bash
# On cedar server - DO NOT USE BACKUP, USE PREVIOUS RELEASE BUILD
sudo systemctl stop brokerknow-api-prod

# Restore previous release
sudo cp -r /opt/brokerknow-api-prod-v{PREVIOUS}/* /opt/brokerknow-api-prod/
sudo chown -R www-data:www-data /opt/brokerknow-api-prod

# Restart with previous version
sudo systemctl start brokerknow-api-prod
```

### Immediate Rollback (Portal)
```bash
# On cedarclient server
sudo cp -r /var/www/html/cedarclient-backup/* /var/www/html/cedarclient/
sudo chown -R www-data:www-data /var/www/html/cedarclient/
sudo systemctl restart nginx
```

---

## Troubleshooting

### API Deployment Issues

**Problem**: 502 Bad Gateway or Connection Refused
```bash
# Check if API is running
sudo systemctl status brokerknow-api-prod

# Check logs
sudo journalctl -u brokerknow-api-prod -n 50 --no-pager

# Verify network connectivity
curl -v https://cedar.martensafrica.com/api/health
```

**Problem**: Database Migration Fails
```bash
# Check database connection
sqlcmd -S cedar-db -U sa -Q "SELECT DB_NAME()"

# Run migration manually
cd /opt/brokerknow-api-prod
dotnet BrokerKnow.Api.dll migrate --environment Production
```

**Problem**: QuestPDF License Error
- Ensure NuGet packages are restored: `dotnet restore`
- Check `appsettings.json` has QuestPDF section (or leave empty for free tier)

### Portal Deployment Issues

**Problem**: Blank page or 404
```bash
# Check if files deployed correctly
ls -la /var/www/html/cedarclient/
# Should see: index.html, js/, css/, assets/ directories

# Check web server configuration
sudo cat /etc/nginx/sites-available/cedarclient
```

**Problem**: API Connection Errors in Browser Console
- Verify `VITE_API_BASE_URL` is set to prod API URL
- Check CORS headers on API: `curl -i -X OPTIONS https://cedar.martensafrica.com/api/health`

---

## Deployment Checklist

### Pre-Deployment
- [ ] Code changes committed and pushed to `develop`
- [ ] All builds successful (no errors)
- [ ] Database backups created
- [ ] Deployment windows scheduled
- [ ] Team notified of downtime (if applicable)

### During Deployment
- [ ] Cedartest API deployed and verified
- [ ] Cedar (Prod) API deployed and verified
- [ ] Cedarclient Portal deployed and verified
- [ ] Health checks passing
- [ ] Smoke tests executed successfully

### Post-Deployment
- [ ] All endpoints responding
- [ ] Portal accessible and functional
- [ ] User can access account-opening workflow
- [ ] PDFs generate and download correctly
- [ ] Logs show no errors
- [ ] Database queries running normally

---

## Contact & Escalation

| Component | Owner | Escalation |
|-----------|-------|-----------|
| cedartest.martensafrica.com API | DevOps | |
| cedar.martensafrica.com API | DevOps | |
| cedarclient.martensafrica.com Portal | DevOps | |
| Database Migrations | DBA | |

---

## Notes

- Always test on **cedartest** before deploying to **cedar (prod)**
- Portal (cedarclient) connects to **prod API** (cedar), not cedartest
- Keep dated backups: `/opt/brokerknow-api-prod-v20260615/`
- Monitor logs after deployment for 24+ hours: `sudo tail -f /var/log/brokerknow-api-prod.log`
