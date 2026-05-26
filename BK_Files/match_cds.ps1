$ErrorActionPreference = 'Stop'
Set-Location 'c:\Users\v-mwangudi\source\repos\BrokerKnow\BK_Files'

$csd = Import-Csv 'IAN MPUNGA CSD.csv' |
       Where-Object { $_.'Client Type' -eq 'Segregated' -and $_.'Account Owner Name' }

$bk = Get-Content 'bk_clients.csv' |
      Where-Object { $_ -match '^\d+\|' } |
      ForEach-Object {
        $p = $_ -split '\|'
        [pscustomobject]@{
          ClientDpa  = [int]$p[0]
          Name       = $p[1]
          IdPass     = $p[2]
          CurrentCds = $p[3]
        }
      }

function Normalise([string]$s) {
  if (-not $s) { return '' }
  $t = $s.ToLowerInvariant().Trim()
  $t = $t -replace "[\u2019']", ''
  $t = $t -replace '[^a-z0-9 ]', ' '
  $t = $t -replace '\s+', ' '
  return $t.Trim()
}

# Token-sort: split on whitespace, drop very short noise tokens (single-letter
# initials), sort alphabetically and rejoin. Lets "Tony De Castro" match
# "DECASTRO Tony", "Daniel Dunga" match "Dunga Daniel", etc.
function TokenSortKey([string]$s) {
  $n = Normalise $s
  if (-not $n) { return '' }
  $tokens = $n -split ' ' | Where-Object { $_.Length -gt 1 } | Sort-Object
  return ($tokens -join ' ')
}

$csdByName = @{}
$csdByTokens = @{}
foreach ($r in $csd) {
  $k = Normalise $r.'Account Owner Name'
  if (-not $k) { continue }
  if (-not $csdByName.ContainsKey($k)) { $csdByName[$k] = New-Object System.Collections.Generic.List[object] }
  $csdByName[$k].Add($r) | Out-Null

  $tk = TokenSortKey $r.'Account Owner Name'
  if ($tk) {
    if (-not $csdByTokens.ContainsKey($tk)) { $csdByTokens[$tk] = New-Object System.Collections.Generic.List[object] }
    $csdByTokens[$tk].Add($r) | Out-Null
  }
}
$csdByUid = @{}
foreach ($r in $csd) {
  $u = "$($r.'Client UID')".Trim()
  if ($u) { $csdByUid[$u] = $r }
}

"`nCSD source:        $($csd.Count) segregated client rows"
"BK target:         $($bk.Count) clients"
"CSD unique names:        $($csdByName.Count)"
"CSD unique token-keys:   $($csdByTokens.Count)"
"CSD unique UIDs:         $($csdByUid.Count)"
"`nMatch results"
"------------------------------------------------------------"

$exactName = 0; $tokenName = 0; $multiNameHits = 0; $uidOnly = 0; $unmatched = 0; $alreadySet = 0
$matched       = New-Object System.Collections.Generic.List[object]
$ambiguous     = New-Object System.Collections.Generic.List[object]
$unmatchedList = New-Object System.Collections.Generic.List[object]

foreach ($c in $bk) {
  if ($c.CurrentCds -and $c.CurrentCds -notmatch '^DEMO\d') { $alreadySet++; continue }
  $key = Normalise $c.Name
  $tkey = TokenSortKey $c.Name
  $hit = $null
  $how = ''

  if ($key -and $csdByName.ContainsKey($key)) {
    $list = $csdByName[$key]
    if ($list.Count -eq 1) {
      $hit = $list[0]; $how = 'Name'; $exactName++
    } else {
      $multiNameHits++
      $ambiguous.Add([pscustomobject]@{
        BkDpa = $c.ClientDpa; BkName = $c.Name; CsdMatches = $list.Count; MatchedBy = 'Name'
      })
      continue
    }
  } elseif ($tkey -and $csdByTokens.ContainsKey($tkey)) {
    $list = $csdByTokens[$tkey]
    if ($list.Count -eq 1) {
      $hit = $list[0]; $how = 'TokenSort'; $tokenName++
    } else {
      $multiNameHits++
      $ambiguous.Add([pscustomobject]@{
        BkDpa = $c.ClientDpa; BkName = $c.Name; CsdMatches = $list.Count; MatchedBy = 'TokenSort'
      })
      continue
    }
  } elseif ($c.IdPass -and $csdByUid.ContainsKey($c.IdPass.Trim())) {
    $hit = $csdByUid[$c.IdPass.Trim()]; $how = 'UID'; $uidOnly++
  } else {
    $unmatched++
    $unmatchedList.Add($c)
    continue
  }

  $matched.Add([pscustomobject]@{
    ClientDpa     = $c.ClientDpa
    BkName        = $c.Name
    BkIdPass      = $c.IdPass
    CsdCode       = $hit.Code
    CsdUID        = $hit.'Client UID'
    CsdOwner      = $hit.'Account Owner Name'
    CsdLegacyCode = $hit.'Legacy Code'
    MatchedBy     = $how
  })
}

"Already has CDS set (skipped):       $alreadySet"
"Exact-name single hit:               $exactName"
"Token-sort name single hit:          $tokenName"
"Ambiguous (name -> multiple):        $multiNameHits"
"UID-only match:                      $uidOnly"
"Unmatched:                           $unmatched"
"Total confident matches:             $($exactName + $tokenName + $uidOnly)"
"`nWriting outputs..."
$matched       | Export-Csv 'cds_match_proposed.csv'  -NoTypeInformation -Encoding UTF8
$ambiguous     | Export-Csv 'cds_match_ambiguous.csv' -NoTypeInformation -Encoding UTF8
$unmatchedList | Export-Csv 'cds_match_unmatched.csv' -NoTypeInformation -Encoding UTF8
"Files written under BK_Files/"
"`nSample matches:"
$matched | Select-Object -First 5 | Format-Table BkDpa,BkName,CsdCode,CsdUID,MatchedBy -AutoSize
