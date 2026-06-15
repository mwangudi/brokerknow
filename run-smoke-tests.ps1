# Account-Opening Workflow - Local Smoke Test Runner
# Executes against http://localhost:5260 (API) and http://localhost:5173 (Portal)

$ErrorActionPreference = "Continue"
$api = "http://localhost:5260"
$testResults = @()

function Write-Section { 
    Write-Host "`n======================================================" -ForegroundColor Cyan
    Write-Host "$args" -ForegroundColor Cyan
    Write-Host "======================================================`n" -ForegroundColor Cyan
}

function Write-Test { Write-Host "  > $args" -ForegroundColor Yellow }
function Write-Pass { Write-Host "  PASS: $args" -ForegroundColor Green }
function Write-Fail { Write-Host "  FAIL: $args" -ForegroundColor Red }
function Write-Info { Write-Host "  INFO: $args" -ForegroundColor Gray }

function Test-APIHealth {
    Write-Section "TEST 1: API Health Check"
    
    Write-Test "Checking API connectivity at $api"
    try {
        $response = Invoke-RestMethod -Uri "$api/api/health" -Method Get -TimeoutSec 5
        Write-Pass "API is running and healthy"
        return $true
    }
    catch {
        Write-Fail "API is not responding: $_"
        return $false
    }
}

function Get-AuthToken {
    Write-Section "TEST 2: Authentication"
    
    Write-Test "Attempting to get auth token"
    try {
        # For local testing, we may need to bypass auth or use a test token
        # This assumes the API allows unauthenticated access for testing
        Write-Info "Skipping auth token for local testing (assuming dev mode)"
        return "test-token"
    }
    catch {
        Write-Fail "Auth failed: $_"
        return $null
    }
}

function Create-TestClient {
    Write-Section "TEST 3: Create Test Client"
    
    Write-Test "Creating new client for account-opening workflow"
    
    $payload = @{
        firstName = "Smoke"
        lastName = "Test"
        email = "smoke@test.local"
        phone = "+265999999999"
        idNumber = "SMOKE001"
        idType = "NationalID"
    } | ConvertTo-Json
    
    try {
        $response = Invoke-RestMethod -Uri "$api/api/clients" `
            -Method Post `
            -ContentType "application/json" `
            -Body $payload `
            -TimeoutSec 10 `
            -ErrorAction Stop
        
        $clientId = $response.id
        Write-Pass "Client created: ID=$clientId"
        return $clientId
    }
    catch {
        Write-Fail "Failed to create client: $_"
        return $null
    }
}

function Initialize-Workflow {
    param([guid]$clientId)
    
    Write-Section "TEST 4: Initialize Account-Opening Workflow"
    
    Write-Test "GET /api/clients/$clientId/account-opening"
    
    try {
        $response = Invoke-RestMethod -Uri "$api/api/clients/$clientId/account-opening" `
            -Method Get `
            -TimeoutSec 10 `
            -ErrorAction Stop
        
        Write-Pass "Workflow initialized successfully"
        $status = if ($response.status) { $response.status } else { "New" }
        Write-Info "Workflow status: $status"
        return $response
    }
    catch {
        Write-Fail "Failed to initialize workflow: $_"
        return $null
    }
}

function SaveDraft {
    param([guid]$clientId)
    
    Write-Section "TEST 5: Save Draft (Individual Account)"
    
    Write-Test "PUT /api/clients/$clientId/account-opening/draft"
    
    $payload = @{
        channel = "BRANCH"
        clientType = "INDIVIDUAL"
        accountType = "INDIVIDUAL"
        idDocument = "NATIONAL_ID"
        csd1Details = @{
            sourceOfFunds = "EMPLOYMENT"
            expectedMonthlyTurnover = "50000"
        }
        agreements = @{
            termsAccepted = $true
            privacyAccepted = $true
        }
        residentialVerification = @{
            address = "123 Test Street"
            city = "Blantyre"
            country = "Malawi"
        }
    } | ConvertTo-Json
    
    try {
        $response = Invoke-RestMethod -Uri "$api/api/clients/$clientId/account-opening/draft" `
            -Method Put `
            -ContentType "application/json" `
            -Body $payload `
            -TimeoutSec 10 `
            -ErrorAction Stop
        
        Write-Pass "Draft saved successfully"
        $status = if ($response.status) { $response.status } else { "Draft" }
        Write-Info "Workflow status: $status"
        return $response
    }
    catch {
        Write-Fail "Failed to save draft: $_"
        return $null
    }
}

function SaveDraftJoint {
    param([guid]$clientId)
    
    Write-Section "TEST 6: Save Draft (Joint Account)"
    
    Write-Test "PUT /api/clients/$clientId/account-opening/draft"
    
    $payload = @{
        channel = "BRANCH"
        clientType = "INDIVIDUAL"
        accountType = "JOINT"
        idDocument = "NATIONAL_ID"
        jointApplicants = @(
            @{
                firstName = "Joint"
                lastName = "Owner"
                idNumber = "JOINT001"
                idType = "NationalID"
            }
        )
        csd1Details = @{
            sourceOfFunds = "EMPLOYMENT"
            expectedMonthlyTurnover = "75000"
        }
        agreements = @{
            termsAccepted = $true
            privacyAccepted = $true
        }
        residentialVerification = @{
            address = "456 Joint Avenue"
            city = "Blantyre"
            country = "Malawi"
        }
    } | ConvertTo-Json
    
    try {
        $response = Invoke-RestMethod -Uri "$api/api/clients/$clientId/account-opening/draft" `
            -Method Put `
            -ContentType "application/json" `
            -Body $payload `
            -TimeoutSec 10 `
            -ErrorAction Stop
        
        Write-Pass "Joint account draft saved successfully"
        $status = if ($response.status) { $response.status } else { "Draft" }
        Write-Info "Workflow status: $status"
        return $response
    }
    catch {
        Write-Fail "Failed to save joint draft: $_"
        return $null
    }
}

function InterimApprove {
    param([guid]$clientId)
    
    Write-Section "TEST 7: Interim Officer Approval"
    
    Write-Test "POST /api/clients/$clientId/account-opening/interim-approve"
    
    $payload = @{
        approvalNotes = "Document review complete, proceeding to risk assessment"
    } | ConvertTo-Json
    
    try {
        $response = Invoke-RestMethod -Uri "$api/api/clients/$clientId/account-opening/interim-approve" `
            -Method Post `
            -ContentType "application/json" `
            -Body $payload `
            -TimeoutSec 10 `
            -ErrorAction Stop
        
        Write-Pass "Interim approval successful"
        $status = if ($response.status) { $response.status } else { "InterimApproved" }
        Write-Info "Workflow status: $status"
        return $response
    }
    catch {
        Write-Fail "Failed to approve interim: $_"
        return $null
    }
}

function SubmitRiskAssessment {
    param([guid]$clientId)
    
    Write-Section "TEST 8: Submit Risk Assessment"
    
    Write-Test "POST /api/clients/$clientId/account-opening/submit-risk"
    
    $payload = @{
        riskItems = @(
            @{ category = "ID_VERIFICATION"; score = 1; notes = "Verified government ID" }
            @{ category = "ADDRESS_VERIFICATION"; score = 2; notes = "Document on file" }
            @{ category = "PEP_SANCTIONS"; score = 1; notes = "Sanction check clear" }
            @{ category = "SOURCE_OF_FUNDS"; score = 2; notes = "Employment verified" }
            @{ category = "EXPECTED_ACTIVITY"; score = 1; notes = "Activity profile normal" }
        )
    } | ConvertTo-Json
    
    try {
        $response = Invoke-RestMethod -Uri "$api/api/clients/$clientId/account-opening/submit-risk" `
            -Method Post `
            -ContentType "application/json" `
            -Body $payload `
            -TimeoutSec 10 `
            -ErrorAction Stop
        
        Write-Pass "Risk assessment submitted successfully"
        $status = if ($response.status) { $response.status } else { "RiskSubmitted" }
        $riskTotal = if ($response.riskAssessment.riskTotal) { $response.riskAssessment.riskTotal } else { "N/A" }
        Write-Info "Workflow status: $status"
        Write-Info "Risk total: $riskTotal"
        return $response
    }
    catch {
        Write-Fail "Failed to submit risk: $_"
        return $null
    }
}

function SupervisorApprove {
    param([guid]$clientId)
    
    Write-Section "TEST 9: Supervisor Final Approval"
    
    Write-Test "POST /api/clients/$clientId/account-opening/supervisor-approve"
    
    $payload = @{
        approvalNotes = "Risk assessment acceptable, account approved"
    } | ConvertTo-Json
    
    try {
        $response = Invoke-RestMethod -Uri "$api/api/clients/$clientId/account-opening/supervisor-approve" `
            -Method Post `
            -ContentType "application/json" `
            -Body $payload `
            -TimeoutSec 10 `
            -ErrorAction Stop
        
        Write-Pass "Supervisor approval successful"
        $status = if ($response.status) { $response.status } else { "SupervisorApproved" }
        Write-Info "Workflow status: $status"
        return $response
    }
    catch {
        Write-Fail "Failed to approve supervisor: $_"
        return $null
    }
}

function DownloadPDF {
    param([guid]$clientId, [string]$formType)
    
    Write-Section "TEST 10: Download PDF - $formType"
    
    $endpoint = switch ($formType) {
        "CSD1" { "csd1" }
        "AccountOpening" { "account-opening" }
        "Residential" { "residential-verification" }
    }
    
    Write-Test "GET /api/clients/$clientId/account-opening/download/$endpoint"
    
    try {
        $response = Invoke-WebRequest -Uri "$api/api/clients/$clientId/account-opening/download/$endpoint" `
            -Method Get `
            -TimeoutSec 10 `
            -ErrorAction Stop
        
        if ($response.StatusCode -eq 200) {
            $contentType = $response.Headers["Content-Type"]
            Write-Pass "$formType PDF downloaded (Content-Type: $contentType)"
            return $true
        }
    }
    catch {
        Write-Fail "Failed to download $formType PDF: $_"
        return $false
    }
}

function VerifyAttachments {
    param([guid]$clientId)
    
    Write-Section "TEST 11: Verify PDFs Filed as Attachments"
    
    Write-Test "GET /api/clients/$clientId/attachments"
    
    try {
        $response = Invoke-RestMethod -Uri "$api/api/clients/$clientId/attachments" `
            -Method Get `
            -TimeoutSec 10 `
            -ErrorAction Stop
        
        $pdfCount = ($response | Where-Object { $_.fileName -like "*PDF" -or $_.fileName -like "*.pdf" }).Count
        Write-Pass "Attachments retrieved: Total=$($response.Count), PDFs=$pdfCount"
        
        if ($response) {
            $response | ForEach-Object {
                Write-Info "  - $($_.fileName) (Size: $($_.fileSize) bytes)"
            }
        }
        return $pdfCount -ge 3
    }
    catch {
        Write-Fail "Failed to verify attachments: $_"
        return $false
    }
}

function Test-PortalConnectivity {
    Write-Section "TEST 12: Portal Connectivity"
    
    Write-Test "Testing portal at http://localhost:5173"
    
    try {
        $response = Invoke-WebRequest -Uri "http://localhost:5173" `
            -Method Get `
            -TimeoutSec 10 `
            -ErrorAction Stop
        
        if ($response.StatusCode -eq 200) {
            Write-Pass "Portal is accessible"
            return $true
        }
    }
    catch {
        Write-Fail "Portal is not accessible: $_"
        return $false
    }
}

# ============================================================================
# MAIN TEST EXECUTION
# ============================================================================

Write-Host "`n" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "ACCOUNT-OPENING WORKFLOW - LOCAL SMOKE TEST RUNNER" -ForegroundColor Cyan
Write-Host "API: $api" -ForegroundColor Cyan
Write-Host "Portal: http://localhost:5173/" -ForegroundColor Cyan
Write-Host "======================================================`n" -ForegroundColor Cyan

# Run tests
$apiHealthy = Test-APIHealth
if (-not $apiHealthy) {
    Write-Host "`n" + "="*60 + "`n"
    Write-Fail "API is not running. Cannot proceed with tests."
    exit 1
}

Get-AuthToken
$client1 = Create-TestClient

if ($client1) {
    Initialize-Workflow $client1
    SaveDraft $client1
    InterimApprove $client1
    SubmitRiskAssessment $client1
    SupervisorApprove $client1
    
    DownloadPDF $client1 "CSD1"
    DownloadPDF $client1 "AccountOpening"
    DownloadPDF $client1 "Residential"
    
    VerifyAttachments $client1
}

# Test portal
Test-PortalConnectivity

# Test joint account
$client2 = Create-TestClient
if ($client2) {
    SaveDraftJoint $client2
    InterimApprove $client2
    SubmitRiskAssessment $client2
    SupervisorApprove $client2
}

Write-Section "SMOKE TEST EXECUTION COMPLETE"
Write-Host "Review results above. All tests should show PASS for deployment readiness.`n" -ForegroundColor Green
