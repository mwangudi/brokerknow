# Account-Opening Workflow - Pre-Deployment Automation Script
# Usage: .\deploy-prep.ps1 -action build|test|package|all

param(
    [ValidateSet("build", "test", "package", "all", "api-only", "web-only")]
    [string]$action = "all"
)

$ErrorActionPreference = "Stop"
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Colors for output
function Write-Section { Write-Host "`n=== $args ===`n" -ForegroundColor Cyan }
function Write-Success { Write-Host "✓ $args" -ForegroundColor Green }
function Write-Error { Write-Host "✗ $args" -ForegroundColor Red }
function Write-Warning { Write-Host "⚠ $args" -ForegroundColor Yellow }

# Verify workspace exists
if (-not (Test-Path "$scriptRoot\brokerknow-api")) {
    Write-Error "brokerknow-api folder not found"
    exit 1
}
if (-not (Test-Path "$scriptRoot\brokerknow-web")) {
    Write-Error "brokerknow-web folder not found"
    exit 1
}

Write-Section "Account-Opening Workflow Deployment Preparation"

# ============================================================================
# BUILD PHASE
# ============================================================================

function Build-API {
    Write-Section "Building API (brokerknow-api)"
    
    try {
        Push-Location "$scriptRoot\brokerknow-api"
        
        Write-Host "1. Cleaning previous builds..."
        & dotnet clean -q
        
        Write-Host "2. Restoring NuGet packages..."
        & dotnet restore -q
        
        Write-Host "3. Building Release configuration..."
        & dotnet build BrokerKnow.slnx -c Release --no-restore
        
        Write-Success "API build completed"
        
        Pop-Location
        return $true
    }
    catch {
        Write-Error "API build failed: $_"
        Pop-Location
        return $false
    }
}

function Build-Web {
    Write-Section "Building Portal (brokerknow-web)"
    
    try {
        Push-Location "$scriptRoot\brokerknow-web"
        
        Write-Host "1. Installing npm dependencies..."
        & npm install --legacy-peer-deps
        
        Write-Host "2. Running build..."
        & npm run build
        
        Write-Success "Portal build completed"
        
        Pop-Location
        return $true
    }
    catch {
        Write-Error "Portal build failed: $_"
        Pop-Location
        return $false
    }
}

# ============================================================================
# TEST PHASE
# ============================================================================

function Test-Builds {
    Write-Section "Testing Build Artifacts"
    
    $buildSuccess = $true
    
    # Check API outputs
    if (Test-Path "$scriptRoot\brokerknow-api\bin\Release") {
        Write-Success "API Release build output exists"
    }
    else {
        Write-Error "API Release build output not found"
        $buildSuccess = $false
    }
    
    # Check Web output
    if (Test-Path "$scriptRoot\brokerknow-web\dist-host") {
        Write-Success "Portal dist-host output exists"
        $fileCount = (Get-ChildItem "$scriptRoot\brokerknow-web\dist-host" -Recurse -File).Count
        Write-Host "  Files: $fileCount" -ForegroundColor Gray
    }
    else {
        Write-Error "Portal dist-host output not found"
        $buildSuccess = $false
    }
    
    return $buildSuccess
}

# ============================================================================
# PACKAGE PHASE
# ============================================================================

function Package-Deployments {
    Write-Section "Creating Deployment Packages"
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $packageDir = "$scriptRoot\deployment-packages-$timestamp"
    
    try {
        New-Item -ItemType Directory -Path $packageDir -Force | Out-Null
        
        # Package API
        Write-Host "1. Packaging API for deployment..."
        Push-Location "$scriptRoot\brokerknow-api"
        
        & dotnet publish -c Release -o "$packageDir\api-package" --no-build
        Write-Success "API packaged: $packageDir\api-package"
        
        Pop-Location
        
        # Package Web
        Write-Host "2. Packaging Portal for deployment..."
        if (Test-Path "$scriptRoot\brokerknow-web\dist-host") {
            Copy-Item "$scriptRoot\brokerknow-web\dist-host\*" "$packageDir\portal-package" -Recurse -Force
            Write-Success "Portal packaged: $packageDir\portal-package"
        }
        else {
            Write-Error "Portal dist-host not found, skipping"
        }
        
        # Create deployment manifest
        Write-Host "3. Creating deployment manifest..."
        $manifest = @{
            timestamp = $timestamp
            buildInfo = @{
                commitHash = & git -C $scriptRoot rev-parse --short HEAD
                branch = & git -C $scriptRoot rev-parse --abbrev-ref HEAD
                timestamp = Get-Date -Format "o"
            }
            contents = @{
                api = "api-package"
                portal = "portal-package"
            }
        }
        
        $manifest | ConvertTo-Json | Out-File "$packageDir\MANIFEST.json"
        Write-Success "Manifest created"
        
        Write-Section "Deployment Packages Ready"
        Write-Host "Location: $packageDir" -ForegroundColor Green
        Write-Host "Contents:" -ForegroundColor Green
        Get-ChildItem $packageDir -Directory | ForEach-Object { Write-Host "  - $_" }
        
        return $true
    }
    catch {
        Write-Error "Packaging failed: $_"
        return $false
    }
}

# ============================================================================
# EXECUTION
# ============================================================================

$buildApiSuccess = $true
$buildWebSuccess = $true
$testSuccess = $false
$packageSuccess = $false

# Execute based on action parameter
switch ($action) {
    "api-only" {
        $buildApiSuccess = Build-API
    }
    "web-only" {
        $buildWebSuccess = Build-Web
    }
    "build" {
        $buildApiSuccess = Build-API
        $buildWebSuccess = Build-Web
    }
    "test" {
        $testSuccess = Test-Builds
    }
    "package" {
        $packageSuccess = Package-Deployments
    }
    "all" {
        $buildApiSuccess = Build-API
        $buildWebSuccess = Build-Web
        $testSuccess = Test-Builds
        $packageSuccess = Package-Deployments
    }
}

# ============================================================================
# SUMMARY
# ============================================================================

Write-Section "Pre-Deployment Summary"

$results = @()
if ($action -in "build", "all", "api-only") {
    $results += "API Build: $(if ($buildApiSuccess) { 'PASS ✓' } else { 'FAIL ✗' })"
}
if ($action -in "build", "all", "web-only") {
    $results += "Web Build: $(if ($buildWebSuccess) { 'PASS ✓' } else { 'FAIL ✗' })"
}
if ($action -in "test", "all") {
    $results += "Build Tests: $(if ($testSuccess) { 'PASS ✓' } else { 'FAIL ✗' })"
}
if ($action -in "package", "all") {
    $results += "Packaging: $(if ($packageSuccess) { 'PASS ✓' } else { 'FAIL ✗' })"
}

$results | ForEach-Object { Write-Host $_ }

Write-Section "Next Steps"

if ($action -in "all", "test", "package") {
    Write-Host "1. Review SMOKE_TEST_PLAN.md for comprehensive test cases" -ForegroundColor Green
    Write-Host "2. Review DEPLOYMENT_GUIDE.md for deployment procedures" -ForegroundColor Green
    Write-Host "3. Start with cedartest environment first" -ForegroundColor Green
    Write-Host "4. Promote to cedar (prod) after validation" -ForegroundColor Green
    Write-Host "5. Deploy portal last (points to cedar prod API)" -ForegroundColor Green
}

$overallSuccess = $buildApiSuccess -and $buildWebSuccess -and (
    if ($action -in "test", "all") { $testSuccess } else { $true }
)

if ($overallSuccess) {
    Write-Success "Pre-deployment preparation complete!"
    exit 0
}
else {
    Write-Error "Pre-deployment preparation incomplete. Fix errors above and retry."
    exit 1
}
