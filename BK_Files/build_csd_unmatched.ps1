#requires -Version 7
# Builds csd_unmatched_for_pm.csv by comparing the full CSD source
# (IAN MPUNGA CSD.csv) against the PM's authoritative matched listing
# (Client Listing - Updated CSD.xlsx, column "CSD Number").

$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
Import-Module ImportExcel

$pmRows = Import-Excel -Path '.\Client Listing - Updated CSD.xlsx' -WorksheetName 'Client Listing'
"PM listing total rows: $($pmRows.Count)"

$pmMatched = @{}
foreach ($r in $pmRows) {
  $cds = "$($r.'CSD Number')".Trim()
  if ($cds) { $pmMatched[$cds] = $true }
}
"PM rows with a CSD Number populated: $($pmMatched.Count)"

$csd = Import-Csv -LiteralPath '.\IAN MPUNGA CSD.csv'
$csdSeg = $csd | Where-Object { $_.'Client Type' -eq 'Segregated' }
"Segregated CSDs in source: $($csdSeg.Count)"

$unmatched = $csdSeg | Where-Object {
  $code = "$($_.Code)".Trim()
  -not $pmMatched.ContainsKey($code)
}
"CSDs not present in PM's matched list: $($unmatched.Count)"

$unmatched | Export-Csv -LiteralPath '.\csd_unmatched_for_pm.csv' -NoTypeInformation -Encoding UTF8
"Wrote csd_unmatched_for_pm.csv ($($unmatched.Count) rows)"
