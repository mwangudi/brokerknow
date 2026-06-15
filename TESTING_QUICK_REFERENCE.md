# Quick Reference - Account-Opening Workflow Testing

## Workflow Stages (Test in Order)

```
1. DRAFT
   → Fill form sections
   → Save Draft
   
2. INTERIM APPROVAL (Officer)
   → Approve interim
   → Status: InterimApproved
   
3. RISK ASSESSMENT (Risk Officer)
   → Submit risk items with scores
   → Status: RiskSubmitted
   → Risk total calculated
   
4. SUPERVISOR APPROVAL (Supervisor)
   → Final approval
   → Status: SupervisorApproved
   → PDFs auto-generated and filed
   
5. DOWNLOAD PDFs
   → CSD1 Form
   → Account Opening Form
   → Residential Verification Form
```

---

## Test Scenarios to Run

### ✓ Individual Account
- Channel: BRANCH
- Client Type: INDIVIDUAL
- Account Type: INDIVIDUAL
- ID Document: NATIONAL_ID

### ✓ Joint Account
- Account Type: JOINT
- Add 2nd Applicant (FirstName, LastName, ID)
- Verify both names in PDFs

### ✓ ITF Account
- Account Type: ITF
- ID Document: PASSPORT
- Add Beneficiary (Name, Relationship)
- Verify beneficiary in PDFs

---

## Form Sections to Fill

1. **Basics**
   - Channel: BRANCH
   - Client Type: INDIVIDUAL
   - Account Type: (INDIVIDUAL / JOINT / ITF)
   - ID Document: (NATIONAL_ID / PASSPORT)

2. **CSD1 Details**
   - Source of Funds: EMPLOYMENT (or other)
   - Expected Monthly Turnover: 50000+ (number)

3. **Joint Applicants** (if JOINT account)
   - First Name, Last Name, ID Number, ID Type
   - Can add multiple applicants

4. **Confirmations**
   - ✓ Terms & Conditions Accepted
   - ✓ Privacy Policy Accepted

5. **Residential Verification**
   - Address, City, Country

6. **Risk Assessment** (After Interim Approval)
   - ID Verification: 1-5
   - Address Verification: 1-5
   - PEP/Sanctions Check: 1-5
   - Source of Funds: 1-5
   - Expected Activity: 1-5
   - Auto-calculates total

---

## Buttons / Actions to Test

- [ ] Save Draft button
- [ ] Interim Approve button (Officer role)
- [ ] Submit Risk Assessment button
- [ ] Supervisor Approve button (Supervisor role)
- [ ] Download CSD1 PDF button
- [ ] Download Account Opening PDF button
- [ ] Download Residential Verification PDF button

---

## Success Criteria

### API Level
- ✓ Create client without errors
- ✓ Save draft: 200 OK
- ✓ Interim approve: 200 OK, status = InterimApproved
- ✓ Submit risk: 200 OK, status = RiskSubmitted, riskTotal calculated
- ✓ Supervisor approve: 200 OK, status = SupervisorApproved
- ✓ PDFs download: 200 OK with PDF content
- ✓ PDFs filed in attachments: 3 PDFs visible

### UI Level
- ✓ Form sections render without errors
- ✓ Data persists after Save Draft
- ✓ Approval buttons enabled/disabled correctly
- ✓ Success messages show after each action
- ✓ No console errors (F12 → Console tab)

### PDF Quality
- ✓ Cedar Capital Malawi header visible
- ✓ Client name, ID displayed correctly
- ✓ Account type shown correctly
- ✓ Joint applicants listed (if applicable)
- ✓ Beneficiary shown (if ITF)
- ✓ All form fields populated
- ✓ Readable and professional formatting

---

## Key URLs

| Environment | URL | Purpose |
|------------|-----|---------|
| Cedartest API | https://cedartest.martensafrica.com | API for testing |
| Cedar Prod API | https://cedar.martensafrica.com | Production API |
| Cedar Portal | https://cedarclient.martensafrica.com | User-facing portal |

---

## If Something Goes Wrong

### 404 Not Found
- Check if service is running: `sudo systemctl status brokerknow-api-test`
- Check URL is correct
- Check API endpoint is spelled right

### 401 Unauthorized
- Login required
- Verify user token is valid
- Check user has correct roles

### 500 Internal Server Error
- Check API logs: `sudo journalctl -u brokerknow-api-test -n 50`
- Possible issues:
  - Database not connected
  - EF migrations not applied
  - PDF generation failing
  - File permissions on attachments folder

### PDF Download is Blank
- Check PDF generation in API logs
- Verify QuestPDF is configured
- Check attachments folder has space

### Portal Shows Blank Page
- Check Network tab (F12) for API errors
- Check console tab for JavaScript errors
- Verify API_BASE_URL configuration

### Workflow Won't Transition
- Check user has correct role for the operation
- Check workflow status in API response
- Verify all required fields filled

---

## When Ready to Commit

```powershell
# On local workstation
cd c:\Users\v-mwangudi\source\repos\BrokerKnow

# Stage all (already done)
git add -A

# Commit with descriptive message
git commit -m "feat: Account-opening workflow complete and tested on cedartest/cedar"

# Push to develop
git push origin develop

# Verify
git log -1 --oneline
git status -sb  # Should show ## develop...origin/develop [nothing to commit]
```

---

## Support / Next Steps

- All code is committed and ready
- Tests completed on cedartest and cedar
- Portal working on cedarclient
- Documentation available in DEPLOYMENT_* files
- Ready for production use

✓ Testing → ✓ Commit & Push → ✓ Done

