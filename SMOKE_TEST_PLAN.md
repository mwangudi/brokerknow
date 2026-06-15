# Account-Opening Workflow - Smoke Test Plan

## Overview
Full end-to-end testing of the account-opening workflow across 3 environments:
- **cedartest.martensafrica.com** - Test API
- **cedar.martensafrica.com** - Production API
- **cedarclient.martensafrica.com** - Production Portal (test against cedar prod API)

## Test Environments & Access

### API Environments
- **Test**: `https://cedartest.martensafrica.com` 
- **Prod**: `https://cedar.martensafrica.com`

### Portal Environment
- **Prod Portal**: `https://cedarclient.martensafrica.com` (connects to prod API)

## Pre-Test Checklist

- [ ] Confirm database migrations are applied to both environments
- [ ] Verify QuestPDF license is configured in both API instances
- [ ] Check attachments storage paths are writeable
- [ ] Confirm Blantyre branch exists in both environments (defaulting for Cedar clients)
- [ ] Verify Cedar Capital Malawi tenant is configured

## Test Workflow Stages

### Stage 1: Create Test Client & Initialize Workflow

**Test Case 1.1: Create Cedar Client** (Test API: cedartest)
```
POST /api/clients
{
  "firstName": "Test",
  "lastName": "OpeningSmoke",
  "email": "test.opening@cedar.test",
  "phone": "+265999999999",
  "idNumber": "TEST001",
  "idType": "NationalID"
}
Expected: Client created, auto-assigned Blantyre branch
```

**Test Case 1.2: Verify Account-Opening Workflow Initialize** (Test API: cedartest)
```
GET /api/clients/{clientId}/account-opening
Expected: 200 OK, returns empty/new AccountOpeningWorkflow structure
```

### Stage 2: Draft Form - CSD1 Details & Joint Applicants

**Test Case 2.1: Save Draft - Individual Account with CSD1 Details** (Test API: cedartest)
```
PUT /api/clients/{clientId}/account-opening/draft
{
  "channel": "BRANCH",
  "clientType": "INDIVIDUAL",
  "accountType": "INDIVIDUAL",
  "idDocument": "NATIONAL_ID",
  "csd1Details": {
    "sourceOfFunds": "EMPLOYMENT",
    "expectedMonthlyTurnover": "50000"
  }
}
Expected: 200 OK, workflow saved
```

**Test Case 2.2: Save Draft - Joint Account with 2nd Applicant** (Test API: cedartest)
```
PUT /api/clients/{clientId}/account-opening/draft
{
  "channel": "BRANCH",
  "clientType": "INDIVIDUAL",
  "accountType": "JOINT",
  "idDocument": "NATIONAL_ID",
  "jointApplicants": [
    {
      "firstName": "Joint",
      "lastName": "Owner",
      "idNumber": "JOINT001",
      "idType": "NationalID"
    }
  ],
  "csd1Details": { ... }
}
Expected: 200 OK, joint applicant persisted
```

**Test Case 2.3: Save Draft - ITF Account with Beneficiary** (Test API: cedartest)
```
PUT /api/clients/{clientId}/account-opening/draft
{
  "channel": "BRANCH",
  "clientType": "INDIVIDUAL",
  "accountType": "ITF",
  "idDocument": "PASSPORT",
  "itfBeneficiary": {
    "name": "ITF Beneficiary",
    "relationship": "Heir"
  },
  "csd1Details": { ... }
}
Expected: 200 OK, ITF beneficiary persisted
```

### Stage 3: Interim Officer Approval

**Test Case 3.1: Officer Approves Draft** (Test API: cedartest)
```
POST /api/clients/{clientId}/account-opening/interim-approve
{
  "approvalNotes": "Document review complete, proceeding to risk assessment"
}
Expected: 200 OK, workflow.status = "InterimApproved"
```

### Stage 4: Risk Assessment Submission

**Test Case 4.1: Submit Risk Assessment** (Test API: cedartest)
```
POST /api/clients/{clientId}/account-opening/submit-risk
{
  "riskItems": [
    { "category": "ID_VERIFICATION", "score": 1, "notes": "Verified government ID" },
    { "category": "ADDRESS_VERIFICATION", "score": 2, "notes": "Document on file" },
    { "category": "PEP_SANCTIONS", "score": 1, "notes": "Sanction check clear" },
    { "category": "SOURCE_OF_FUNDS", "score": 2, "notes": "Employment verified" },
    { "category": "EXPECTED_ACTIVITY", "score": 1, "notes": "Activity profile normal" }
  ]
}
Expected: 200 OK, riskTotal = 7, workflow.status = "RiskSubmitted"
```

### Stage 5: Supervisor Final Approval & PDF Generation

**Test Case 5.1: Supervisor Approves & Triggers PDFs** (Test API: cedartest)
```
POST /api/clients/{clientId}/account-opening/supervisor-approve
{
  "approvalNotes": "Risk assessment acceptable, account approved"
}
Expected: 200 OK, workflow.status = "SupervisorApproved"
         3 PDFs generated and saved to attachments
```

### Stage 6: PDF Download & Verification

**Test Case 6.1: Download CSD1 Form PDF** (Test API: cedartest)
```
GET /api/clients/{clientId}/account-opening/download/csd1
Expected: 200 OK, PDF file download
          File contains: Cedar Capital Malawi header, client details, account type, CSD1 details
```

**Test Case 6.2: Download Account Opening Form PDF** (Test API: cedartest)
```
GET /api/clients/{clientId}/account-opening/download/account-opening
Expected: 200 OK, PDF file download
          File contains: Cedar masthead, client info, account type, joint applicants (if applicable)
```

**Test Case 6.3: Download Residential Verification PDF** (Test API: cedartest)
```
GET /api/clients/{clientId}/account-opening/download/residential-verification
Expected: 200 OK, PDF file download
          File contains: Address details, signature lines for officer & supervisor
```

**Test Case 6.4: Verify PDFs Filed as Attachments** (Test API: cedartest)
```
GET /api/clients/{clientId}/attachments
Expected: 3 new attachments present:
          - CSD1_Form_{timestamp}.pdf
          - Account_Opening_Form_{timestamp}.pdf
          - Residential_Verification_{timestamp}.pdf
```

### Stage 7: Portal UI Workflow (Prod Portal against Prod API)

**Test Case 7.1: Access Account-Opening Page in Portal** (Prod Portal: cedarclient)
```
Navigate to: https://cedarclient.martensafrica.com/clients/{clientId}/account-opening
Expected: Page loads, displays form sections:
          - Basics (channel, client type, account type)
          - CSD1 Details
          - Joint Applicants (if applicable)
          - Confirmations
          - Residential Verification
          - Risk Scoring
```

**Test Case 7.2: Complete Workflow via Portal UI** (Prod Portal: cedarclient)
```
1. Fill form → Save Draft
2. Officer: Click "Interim Approve" button
3. Risk Officer: Fill risk items → Submit Risk Assessment
4. Supervisor: Click "Supervisor Approve" button
Expected: UI shows success messages at each stage
          PDF download buttons become enabled after approval
```

**Test Case 7.3: Download PDFs via Portal** (Prod Portal: cedarclient)
```
Click each PDF download button (CSD1, Account Opening, Residential)
Expected: PDFs download with proper formatting
```

## Validation Criteria

### Data Integrity
- [ ] Workflow persists across page refreshes
- [ ] Joint applicants correctly stored and retrieved
- [ ] ITF beneficiary data preserved
- [ ] Risk scores calculated correctly (sum of all item scores)
- [ ] Account type influences form fields correctly (ITF shows beneficiary, Joint shows 2nd applicant)

### PDF Generation
- [ ] All 3 PDFs generate without errors
- [ ] PDF files contain correct client data
- [ ] Cedar Capital Malawi branding visible
- [ ] Forms are readable and properly formatted
- [ ] Signature lines present in residential verification

### Attachment Storage
- [ ] PDFs filed to client attachments after supervisor approval
- [ ] Attachments retrievable via API
- [ ] File names include timestamps
- [ ] No duplicate attachments on re-approval

### API Response Codes
- [ ] 200 OK on successful operations
- [ ] 400 Bad Request on invalid data
- [ ] 404 Not Found on missing client/workflow
- [ ] 401 Unauthorized on auth failures

### Multi-Environment Consistency
- [ ] Test API (cedartest): All tests pass
- [ ] Prod API (cedar): All tests pass (with prod data)
- [ ] Prod Portal (cedarclient): UI/UX works correctly

## Test Execution Steps

### Local Testing (before deployment)
1. Start local BrokerKnow API (`dotnet run`)
2. Start local React portal (`npm run dev`)
3. Run Test Cases 1.1 - 7.3 using Postman/curl and browser
4. Verify all validation criteria

### Cedartest Deployment
1. Build and publish API: `dotnet publish -c Release`
2. Deploy to cedartest.martensafrica.com
3. Run Test Cases 1.1 - 6.4 (API tests)
4. Verify attachments storage is accessible

### Cedar (Prod) Deployment
1. Build and publish API: `dotnet publish -c Release`
2. Deploy to cedar.martensafrica.com
3. Run Test Cases 1.1 - 6.4 with fresh test client
4. Verify attachments storage is accessible

### Cedarclient Portal Deployment
1. Build React: `npm run build`
2. Deploy to cedarclient.martensafrica.com
3. Run Test Cases 7.1 - 7.3 (Portal UI tests)
4. Confirm API connectivity to prod cedar API

## Rollback Plan

If critical issues found:
1. Revert git to previous commit: `git revert <commit-hash>`
2. Rebuild and redeploy to affected environments
3. Notify users of issue and ETA
4. Coordinate database rollback if needed

## Sign-Off

| Environment | Tester | Date | Status | Notes |
|-------------|--------|------|--------|-------|
| Cedartest API | | | | |
| Cedar API | | | | |
| Cedarclient Portal | | | | |

---

## Quick Reference: Test Data Templates

### Individual Account
```json
{
  "channel": "BRANCH",
  "clientType": "INDIVIDUAL",
  "accountType": "INDIVIDUAL",
  "idDocument": "NATIONAL_ID"
}
```

### Joint Account
```json
{
  "channel": "BRANCH",
  "clientType": "INDIVIDUAL",
  "accountType": "JOINT",
  "idDocument": "NATIONAL_ID",
  "jointApplicants": [
    { "firstName": "Second", "lastName": "Owner", "idNumber": "ID002", "idType": "NationalID" }
  ]
}
```

### ITF Account
```json
{
  "channel": "BRANCH",
  "clientType": "INDIVIDUAL",
  "accountType": "ITF",
  "idDocument": "PASSPORT",
  "itfBeneficiary": { "name": "Beneficiary Name", "relationship": "Heir" }
}
```
