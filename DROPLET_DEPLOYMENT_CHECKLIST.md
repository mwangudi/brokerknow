# Account-Opening Workflow - Droplet Deployment Checklist

**Date**: 2026-06-15  
**Status**: Ready for Physical Testing & Deployment  
**Git Branch**: develop  

---

## Pre-Deployment Git Status

✅ **All Changes Staged**
```
Modified: brokerknow-api/ (new account-opening controller + PDF generators)
Modified: brokerknow-web/ (new account-opening page + hooks + routes)
Added: Documentation (DEPLOYMENT_GUIDE.md, SMOKE_TEST_PLAN.md, etc.)
Added: Reference Documents (BK_Files/Docs/Acc-Opening/)
```

**Ready to commit and push**

---

## Physical Testing on Droplet - Workflow

### 1. Deploy to Cedartest First
```bash
# SSH to cedartest server
ssh user@cedartest.martensafrica.com

# Pull latest code
cd /opt/brokerknow-api
git pull origin develop

# Build Release
dotnet publish -c Release -o /opt/brokerknow-api-release

# Stop current service
sudo systemctl stop brokerknow-api-test

# Deploy new build
sudo cp -r /opt/brokerknow-api-release/* /opt/brokerknow-api/
sudo chown -R www-data:www-data /opt/brokerknow-api/

# Apply migrations if needed
cd /opt/brokerknow-api
dotnet EF database update

# Start service
sudo systemctl start brokerknow-api-test
sudo systemctl status brokerknow-api-test
```

### 2. Test Account-Opening Workflow on Cedartest
**Physical Testing Steps**:

1. **Create Test Client**
   - Open portal: https://cedartest.martensafrica.com
   - Create new client with ID
   - Verify Blantyre branch auto-assigned

2. **Test Individual Account Workflow**
   - Navigate to Account Opening tab
   - Fill form: Channel = BRANCH, Type = INDIVIDUAL, Account = INDIVIDUAL
   - Save Draft
   - Officer: Approve Interim
   - Risk Officer: Submit Risk Assessment
   - Supervisor: Approve Final
   - Verify all 3 PDFs download correctly

3. **Test Joint Account Workflow**
   - Create another test client
   - Account Opening → Type = JOINT
   - Add 2nd applicant
   - Go through full approval process
   - Verify PDFs generated with both applicants

4. **Test ITF Account Workflow**
   - Create another test client
   - Account Opening → Type = ITF
   - Add beneficiary info
   - Complete approval workflow
   - Verify PDFs include beneficiary

5. **Verify PDF Quality**
   - Download each PDF
   - Check formatting: Cedar masthead, logo, client info correct
   - Verify all fields populated
   - Check signature lines present

6. **Verify Attachments Storage**
   - Navigate to client attachments
   - Confirm 3 PDFs filed after supervisor approval
   - File names include timestamps

### 3. If All Cedartest Tests Pass → Deploy to Cedar (Prod)
```bash
# SSH to cedar server
ssh user@cedar.martensafrica.com

# Pull latest code
cd /opt/brokerknow-api
git pull origin develop

# Build Release (same as cedartest)
dotnet publish -c Release -o /opt/brokerknow-api-release

# Create backup of current version
sudo cp -r /opt/brokerknow-api /opt/brokerknow-api-backup-20260615

# Stop service
sudo systemctl stop brokerknow-api-prod

# Deploy new build
sudo cp -r /opt/brokerknow-api-release/* /opt/brokerknow-api/
sudo chown -R www-data:www-data /opt/brokerknow-api/

# Start service
sudo systemctl start brokerknow-api-prod
sudo systemctl status brokerknow-api-prod
```

### 4. Test on Cedar (Prod)
- Repeat steps 1-6 with prod database
- Use real client data (or isolated test clients)
- Verify against production attachments storage

### 5. Deploy Portal to Cedarclient
```bash
# SSH to cedarclient server
ssh user@cedarclient.martensafrica.com

# Pull latest code
cd /var/www/html/brokerknow-web
git pull origin develop

# Build production
npm run build

# Backup current version
sudo cp -r dist-host dist-host-backup-20260615

# Deploy new build
sudo cp -r dist-host/* /var/www/html/cedarclient/
sudo chown -R www-data:www-data /var/www/html/cedarclient/

# Verify portal is accessible
curl -I https://cedarclient.martensafrica.com/
```

### 6. Final Portal Testing
- Access https://cedarclient.martensafrica.com
- Navigate to Account Opening for test client
- Verify all form sections load
- Test Save Draft
- Test Approval workflow (if user has roles)
- Download PDFs
- Verify API connectivity (Network tab in DevTools)

---

## Critical Verification Points

### Before Marking as Complete

- [ ] **API Endpoints Working**
  - Account-opening workflow accessible
  - PDF download endpoints respond
  - Attachments are filed correctly

- [ ] **PDF Generation**
  - CSD1 Form PDF has correct formatting and data
  - Account Opening Form PDF displays properly
  - Residential Verification PDF has signature lines

- [ ] **Database**
  - Blantyre branch exists
  - Cedar Capital Malawi tenant configured
  - User roles (Officer, RiskOfficer, Supervisor) present

- [ ] **Attachment Storage**
  - PDFs saved to attachments folder
  - File permissions correct
  - Space available

- [ ] **Portal UI**
  - Form sections render correctly
  - API calls succeed (200/201 responses)
  - No console errors in browser
  - Download buttons work

- [ ] **Workflow State**
  - Workflow status transitions: Draft → InterimApproved → RiskSubmitted → SupervisorApproved
  - Risk scores calculate correctly
  - Data persists across page refresh

---

## Rollback Plan (If Issues Found)

### Quick Rollback
```bash
# On cedartest or cedar server
sudo systemctl stop brokerknow-api-prod  # or brokerknow-api-test

# Restore previous version
sudo cp -r /opt/brokerknow-api-backup-20260615/* /opt/brokerknow-api/
sudo chown -R www-data:www-data /opt/brokerknow-api/

# Restart
sudo systemctl start brokerknow-api-prod
```

### Database Rollback (if migrations applied)
```bash
# If migrations were run, you may need to revert
cd /opt/brokerknow-api
dotnet ef database update --previous-version
```

---

## After Physical Testing Passes

### Commit & Push to Git
```bash
# From local workstation
cd c:\Users\v-mwangudi\source\repos\BrokerKnow

# Check status
git status --short

# All changes already staged, so commit
git commit -m "feat: Add account-opening workflow

- AccountOpeningController with 6 endpoints for workflow stages
- Csd1FormPdf, ClientAccountOpeningWorkflowPdf, ResidentialAddressVerificationPdf generators
- ClientAccountOpeningPage React component with full form workflow
- Support for Individual, Joint, and ITF account types
- Multi-stage approval: interim officer → risk assessment → supervisor
- Risk scoring with auto-calculation
- PDF filing to client attachments
- ClientsController: Auto-assign Blantyre branch for Cedar clients

Smoke tested on cedartest and cedar environments"

# Push to develop
git push origin develop

# Verify push
git status -sb
```

---

## Troubleshooting Guide

### PDF Generation Fails
```
Error: QuestPDF license
Fix: Ensure QuestPDF v5+ (free tier) is configured
Check: appsettings.json has QuestPDF section
```

### Attachments Not Filing
```
Error: File path not accessible
Fix: Verify attachments storage folder exists and is writeable
Check: Folder permissions: sudo chown -R www-data:www-data /opt/attachments/
```

### Authorization Fails
```
Error: User cannot approve workflow
Fix: Verify user has correct roles: Officer, RiskOfficer, Supervisor
Check: User claims include required roles in token
```

### API Responds 500
```
Error: Internal server error
Fix: Check API logs: sudo journalctl -u brokerknow-api-prod -n 50
Check: Database connection string in appsettings.json
Check: EF migrations applied successfully
```

### Portal Shows Blank Page
```
Error: 404 or blank page
Fix: Verify dist-host built correctly
Check: Browser DevTools Network tab for API errors
Check: VITE_API_BASE_URL points to correct API
```

---

## Sign-Off

| Step | Status | Tester | Date | Notes |
|------|--------|--------|------|-------|
| Cedartest Deployment | ⬜ | | | |
| Cedartest Testing | ⬜ | | | |
| Cedar Deployment | ⬜ | | | |
| Cedar Testing | ⬜ | | | |
| Portal Deployment | ⬜ | | | |
| Portal Testing | ⬜ | | | |
| Git Push to Develop | ⬜ | | | |
| All Tests Passing | ⬜ | | | READY FOR RELEASE |

---

## Documentation Reference

- [SMOKE_TEST_PLAN.md](./SMOKE_TEST_PLAN.md) - Complete test scenarios
- [DEPLOYMENT_GUIDE.md](./DEPLOYMENT_GUIDE.md) - Detailed deployment steps
- [DEPLOYMENT_STATUS.md](./DEPLOYMENT_STATUS.md) - Current readiness status
- [SMOKE_TEST_RESULTS.md](./SMOKE_TEST_RESULTS.md) - Local test verification

---

**Next Action**: Deploy to cedartest, run physical tests, then proceed with cedar and portal deployments.

Once all tests pass → Commit & Push → Done ✓
