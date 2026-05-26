#requires -Version 7
# Generates apply_cds.sql (UPDATE ... in one transaction) and dpa_list.txt
# (the list of Client_DPA_ values to snapshot for rollback) from
# cds_match_proposed.csv.

$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot

$rows = Import-Csv -LiteralPath 'cds_match_proposed.csv'
"Loaded $($rows.Count) proposed matches"

$apply = New-Object System.Collections.Generic.List[string]
$dpas  = New-Object System.Collections.Generic.List[string]

$apply.Add("SET NOCOUNT ON;")
$apply.Add("SET XACT_ABORT ON;")
$apply.Add("BEGIN TRAN;")
$apply.Add("")

$skipped = 0
foreach ($r in $rows) {
  $dpa = "$($r.ClientDpa)".Trim()
  $cds = "$($r.CsdCode)".Trim()
  if (-not $dpa -or $dpa -notmatch '^\d+$') { $skipped++; continue }
  if (-not $cds) { $skipped++; continue }
  $cdsEsc = $cds -replace "'", "''"
  $apply.Add("UPDATE Client SET ClientCDSNo = N'$cdsEsc' WHERE Client_DPA_ = $dpa AND ISNULL(Deleted,0)=0;")
  $dpas.Add($dpa)
}

$apply.Add("")
$apply.Add("COMMIT;")

Set-Content -LiteralPath 'apply_cds.sql' -Value ($apply -join "`r`n") -Encoding UTF8
Set-Content -LiteralPath 'dpa_list.txt'  -Value ($dpas  -join ',')    -Encoding ASCII

"Wrote apply_cds.sql ($($apply.Count) lines)"
"Wrote dpa_list.txt  ($($dpas.Count) DPAs)"
if ($skipped) { "Skipped $skipped malformed rows" }
