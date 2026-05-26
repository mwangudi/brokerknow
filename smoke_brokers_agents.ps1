$ErrorActionPreference = 'Stop'
$base = 'http://localhost:5260/api'
$pass = 0
$fail = 0
function Step($name, $block) {
  try {
    & $block
    Write-Host "  PASS  $name" -ForegroundColor Green
    $script:pass++
  } catch {
    Write-Host "  FAIL  $name -> $($_.Exception.Message)" -ForegroundColor Red
    if ($_.ErrorDetails) { Write-Host "        $($_.ErrorDetails.Message)" -ForegroundColor DarkGray }
    $script:fail++
  }
}

Write-Host "`n=== BROKERS smoke ===" -ForegroundColor Cyan

$createdBrokerId = $null
$listBefore = @()

Step "GET /brokers (list)" {
  $script:listBefore = Invoke-RestMethod -Uri "$base/brokers" -Method Get
  Write-Host "        $(@($script:listBefore).Count) brokers"
}

$smokeCode = "ZSMK"
Step "POST /brokers (create)" {
  $body = @{
    brokerCode       = $smokeCode
    brokerName       = "Smoke Test Broker"
    brokerAddr       = "PO Box 1, Smoketown"
    brokerOfficeTel  = "+265 1 111 111"
    brokerFax        = "+265 1 111 222"
    brokerOpeningBal = 12345.67
  } | ConvertTo-Json
  $resp = Invoke-RestMethod -Uri "$base/brokers" -Method Post -Body $body -ContentType 'application/json'
  $script:createdBrokerId = $resp.id
  if (-not $script:createdBrokerId) { throw "no id returned" }
  Write-Host "        new BrokerDpa = $($script:createdBrokerId)"
}

Step "POST /brokers (duplicate code rejected)" {
  $body = @{ brokerCode = $smokeCode; brokerName = "Dup Test" } | ConvertTo-Json
  try {
    Invoke-RestMethod -Uri "$base/brokers" -Method Post -Body $body -ContentType 'application/json' | Out-Null
    throw "duplicate was accepted"
  } catch {
    if ($_.Exception.Response.StatusCode.value__ -ne 400) { throw }
  }
}

Step "POST /brokers (long code rejected)" {
  $body = @{ brokerCode = "TOOLONG"; brokerName = "X" } | ConvertTo-Json
  try {
    Invoke-RestMethod -Uri "$base/brokers" -Method Post -Body $body -ContentType 'application/json' | Out-Null
    throw "accepted >5 chars"
  } catch {
    if ($_.Exception.Response.StatusCode.value__ -ne 400) { throw }
  }
}

Step "GET /brokers/{id}" {
  $b = Invoke-RestMethod -Uri "$base/brokers/$script:createdBrokerId" -Method Get
  if ($b.brokerCode -ne $smokeCode) { throw "code mismatch: $($b.brokerCode)" }
  if ($b.brokerOpeningBal -ne 12345.67) { throw "balance mismatch: $($b.brokerOpeningBal)" }
}

Step "GET /brokers?search=Smoke" {
  $hits = Invoke-RestMethod -Uri "$base/brokers?search=Smoke" -Method Get
  if (-not ($hits | Where-Object { $_.brokerDpa -eq $script:createdBrokerId })) { throw "not found via search" }
}

Step "PUT /brokers/{id} (edit)" {
  $body = @{
    brokerCode       = $smokeCode
    brokerName       = "Smoke Test Broker (edited)"
    brokerAddr       = "Updated address"
    brokerOfficeTel  = "+265 9 999 999"
    brokerFax        = $null
    brokerOpeningBal = 99999.99
  } | ConvertTo-Json
  Invoke-RestMethod -Uri "$base/brokers/$script:createdBrokerId" -Method Put -Body $body -ContentType 'application/json' | Out-Null
  $b = Invoke-RestMethod -Uri "$base/brokers/$script:createdBrokerId" -Method Get
  if ($b.brokerName -ne "Smoke Test Broker (edited)") { throw "name not updated" }
  if ($b.brokerOfficeTel -ne "+265 9 999 999") { throw "tel not updated" }
}

Step "DELETE /brokers/{id}" {
  Invoke-RestMethod -Uri "$base/brokers/$script:createdBrokerId" -Method Delete | Out-Null
  try {
    Invoke-RestMethod -Uri "$base/brokers/$script:createdBrokerId" -Method Get | Out-Null
    throw "still exists after delete"
  } catch {
    if ($_.Exception.Response.StatusCode.value__ -ne 404) { throw }
  }
}

Step "DELETE /brokers/{existingWithLots} refused" {
  # find a broker that has lots
  $withLots = $script:listBefore | Where-Object { $_.lotCount -gt 0 } | Select-Object -First 1
  if (-not $withLots) { Write-Host "        (no broker has lots — skipping)" -ForegroundColor DarkYellow; return }
  try {
    Invoke-RestMethod -Uri "$base/brokers/$($withLots.brokerDpa)" -Method Delete | Out-Null
    throw "delete with lots was allowed"
  } catch {
    if ($_.Exception.Response.StatusCode.value__ -ne 400) { throw }
  }
}

Write-Host "`n=== AGENTS smoke ===" -ForegroundColor Cyan

$createdAgentId = $null

Step "GET /agents (list)" {
  $rows = Invoke-RestMethod -Uri "$base/agents" -Method Get
  Write-Host "        $(@($rows).Count) agents"
}

Step "POST /agents (create)" {
  $body = @{
    agentName       = "Smoke Test Agent"
    agentEmail      = "smoke@test.local"
    agentCellTel    = "+265 88 888 888"
    agentIdPass     = "SMK-001"
    agentOpeningBal = 1000
  } | ConvertTo-Json
  $resp = Invoke-RestMethod -Uri "$base/agents" -Method Post -Body $body -ContentType 'application/json'
  $script:createdAgentId = $resp.id
  if (-not $script:createdAgentId) { throw "no id returned" }
  Write-Host "        new AgentDpa = $($script:createdAgentId)"
}

Step "POST /agents (no name rejected)" {
  $body = @{ agentName = "" } | ConvertTo-Json
  try {
    Invoke-RestMethod -Uri "$base/agents" -Method Post -Body $body -ContentType 'application/json' | Out-Null
    throw "blank name accepted"
  } catch {
    if ($_.Exception.Response.StatusCode.value__ -ne 400) { throw }
  }
}

Step "POST /agents (bad email rejected)" {
  $body = @{ agentName = "Bad Mail"; agentEmail = "notanemail" } | ConvertTo-Json
  try {
    Invoke-RestMethod -Uri "$base/agents" -Method Post -Body $body -ContentType 'application/json' | Out-Null
    throw "bad email accepted"
  } catch {
    if ($_.Exception.Response.StatusCode.value__ -ne 400) { throw }
  }
}

Step "GET /agents/{id}" {
  $a = Invoke-RestMethod -Uri "$base/agents/$script:createdAgentId" -Method Get
  if ($a.agentName -ne "Smoke Test Agent") { throw "name mismatch" }
}

Step "GET /agents?search=Smoke" {
  $hits = Invoke-RestMethod -Uri "$base/agents?search=Smoke" -Method Get
  if (-not ($hits | Where-Object { $_.agentDpa -eq $script:createdAgentId })) { throw "not found via search" }
}

Step "PUT /agents/{id} (edit)" {
  $body = @{
    agentName       = "Smoke Test Agent (edited)"
    agentEmail      = "smoke2@test.local"
    agentCellTel    = "+265 11 222 333"
    agentIdPass     = "SMK-001"
    agentOpeningBal = 99999  # should be ignored on edit
  } | ConvertTo-Json
  Invoke-RestMethod -Uri "$base/agents/$script:createdAgentId" -Method Put -Body $body -ContentType 'application/json' | Out-Null
  $a = Invoke-RestMethod -Uri "$base/agents/$script:createdAgentId" -Method Get
  if ($a.agentName -ne "Smoke Test Agent (edited)") { throw "name not updated" }
  if ($a.agentEmail -ne "smoke2@test.local") { throw "email not updated" }
  if ([decimal]$a.agentOpeningBal -ne 1000) { throw "opening balance changed (expected locked at 1000, got $($a.agentOpeningBal))" }
}

Step "DELETE /agents/{id} (soft)" {
  Invoke-RestMethod -Uri "$base/agents/$script:createdAgentId" -Method Delete | Out-Null
  try {
    Invoke-RestMethod -Uri "$base/agents/$script:createdAgentId" -Method Get | Out-Null
    throw "still visible after soft delete"
  } catch {
    if ($_.Exception.Response.StatusCode.value__ -ne 404) { throw }
  }
}

Write-Host "`n=== Result: $pass passed, $fail failed ===`n" -ForegroundColor $(if ($fail -eq 0) {'Green'} else {'Red'})
exit $fail
