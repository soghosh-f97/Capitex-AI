<#
    Self-contained tests for the v3.0 live-data layer (no Pester, no internet, no AWS needed).
    Run:  pwsh -NoProfile -File Tests/Test-LiveData.ps1        (or powershell.exe on Windows)
    Exit code = number of failed checks.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
foreach ($m in @('LiveDataProviders.ps1', 'PricingIndex.ps1', 'InstanceSizing.ps1', 'OSMapping.ps1', 'MigrationStrategyGenerator.ps1')) {
    . (Join-Path (Join-Path $root 'Modules') $m)
}
$script:fail = 0; $script:pass = 0
function Check {
    param([string]$Name, [bool]$Condition, [string]$Detail = '')
    if ($Condition) { $script:pass++; Write-Host "  PASS  $Name" -ForegroundColor Green }
    else { $script:fail++; Write-Host "  FAIL  $Name  $Detail" -ForegroundColor Red }
}
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('livedata-test-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
$settingsPath = Join-Path (Join-Path $root 'Config') 'AWSMigrationStrategySettings.psd1'
function New-Settings { Get-MigrationStrategyDefaults -SettingsPath $settingsPath }

Write-Host "1. Parser: endoflife.date v1 shape with bool / null / missing values"
$v1 = @'
{ "schema_version":"1.2.0","result":{"name":"windows","releases":[
  {"name":"11-24h2-e","label":"11 24H2 (E)","releaseDate":"2024-10-01","isLts":false,"eoasFrom":null,"eolFrom":"2027-10-12","eoesFrom":null},
  {"name":"11-24h2-w","label":"11 24H2 (W)","releaseDate":"2024-10-01","isLts":false,"eolFrom":"2026-10-13","eoesFrom":false},
  {"name":"10-22h2","label":"10 22H2","isLts":false,"eolFrom":"2025-10-14","eoesFrom":"2028-10-10"},
  {"name":"10-21h2-e-lts","label":"10 21H2 (E) (LTSC)","isLts":true,"eolFrom":true},
  {"name":"","label":"blank name is skipped"} ]}}
'@ | ConvertFrom-Json
$r = @(ConvertFrom-LifecycleJson -Product 'windows' -Json $v1 -Source 'test')
Check 'v1: 4 releases (blank name skipped)' ($r.Count -eq 4) "got $($r.Count)"
Check 'v1: eolFrom date kept' ($r[0].EndOfSupport -eq '2027-10-12')
Check 'v1: null eoesFrom -> empty' ($r[0].ExtendedEnd -eq '')
Check 'v1: false eoesFrom -> empty' ($r[1].ExtendedEnd -eq '')
Check 'v1: eolFrom = true -> empty' ($r[3].EndOfSupport -eq '')
Check 'v1: isLts true' ($r[3].IsLts -eq $true)
Check 'v1: display name E' ($r[0].DisplayName -eq 'Windows 11 24H2 - Enterprise / Education') $r[0].DisplayName
Check 'v1: display name W' ($r[1].DisplayName -eq 'Windows 11 24H2 - Home / Pro') $r[1].DisplayName
Check 'v1: display name 10 22H2' ($r[2].DisplayName -eq 'Windows 10 22H2') $r[2].DisplayName
Check 'v1: display name LTSC' ($r[3].DisplayName -eq 'Windows 10 Enterprise LTSC 2021') $r[3].DisplayName

Write-Host "2. Parser: legacy API shape"
$legacy = '[{"cycle":"2012-r2","releaseDate":"2013-11-25","support":"2018-10-09","eol":"2023-10-10","extendedSupport":"2026-10-13","lts":true},{"cycle":"2019","eol":"2029-01-09","extendedSupport":false,"lts":"2018-11-13"}]' | ConvertFrom-Json
$r = @(ConvertFrom-LifecycleJson -Product 'windows-server' -Json $legacy)
Check 'legacy: 2 releases' ($r.Count -eq 2)
Check 'legacy: eol/extended mapped' ($r[0].EndOfSupport -eq '2023-10-10' -and $r[0].ExtendedEnd -eq '2026-10-13' -and $r[0].MainstreamEnd -eq '2018-10-09')
Check 'legacy: extendedSupport false -> empty' ($r[1].ExtendedEnd -eq '')
Check 'legacy: name 2012 R2' ($r[0].DisplayName -eq 'Windows Server 2012 R2') $r[0].DisplayName
Check 'parser: unknown shape -> no releases' (@(ConvertFrom-LifecycleJson -Product 'x' -Json ('{"foo":1}' | ConvertFrom-Json)).Count -eq 0)
Check 'date: DateTime value' ((ConvertTo-LiveDate ([datetime]'2030-01-02')) -eq '2030-01-02')
Check 'date: timestamp string' ((ConvertTo-LiveDate '2030-01-02T00:00:00Z') -eq '2030-01-02')

Write-Host "3. Offline catalogue = snapshot, and the snapshot changes no settings date (v2.2 results kept)"
$s = New-Settings
$cat = Get-LifecycleCatalog -Settings $s -Offline -CacheDir $tmp
$expected = @('windows-server', 'windows', 'rhel', 'oracle-linux', 'centos', 'debian', 'ubuntu', 'sles', 'rocky-linux')
foreach ($p in $expected) { Check "offline: $p from snapshot" ($cat.Source[$p] -eq 'Snapshot') "source=$($cat.Source[$p])" }
$changes = @(Update-OsLifecycleFromCatalog -Settings $s -Catalog $cat)
Check 'offline: 0 date changes vs settings file' (@($changes | Where-Object { $_.Field -ne 'Live' }).Count -eq 0) (($changes | ForEach-Object { "$($_.Label) $($_.Field) $($_.Old)->$($_.New)" }) -join '; ')
Check 'offline: every Live key found' (@($changes | Where-Object { $_.Field -eq 'Live' }).Count -eq 0) (($changes | ForEach-Object { $_.Source }) -join '; ')
Check 'offline: DateSource set to snapshot' ((@($s.OsLifecycle | Where-Object { $_.Live -and $_.DateSource -notlike 'offline snapshot*' })).Count -eq 0)
Check 'offline: rows without Live -> settings file' ((@($s.OsLifecycle | Where-Object { -not $_.Live -and $_.DateSource -ne 'settings file' })).Count -eq 0)
$w = @($cat.Products['windows'] | Where-Object { $_.Cycle -like '11-*' })
Check 'snapshot: Windows 11 releases for the chooser' ($w.Count -ge 10) "got $($w.Count)"

Write-Host "4. Cache: a newer cached date is applied, the date-only note is rewritten, stale cache is flagged"
$s = New-Settings
$fake = '{"result":{"name":"rhel","releases":[{"name":"6","eolFrom":"2020-11-30","eoesFrom":"2029-06-30"},{"name":"7","eolFrom":"2024-06-30","eoesFrom":"2028-06-30"}]}}'
[System.IO.File]::WriteAllText((Join-Path $tmp 'eol_rhel.json'), $fake)
$cat = Get-LifecycleCatalog -Settings $s -Offline -CacheDir $tmp
Check 'cache: rhel from fresh cache' ($cat.Source['rhel'] -eq 'Cache') $cat.Source['rhel']
$changes = @(Update-OsLifecycleFromCatalog -Settings $s -Catalog $cat)
$rhel6 = $s.OsLifecycle | Where-Object { $_.Label -eq 'RHEL 6' }
Check 'cache: RHEL 6 extended end updated' ($rhel6.ExtendedEnd -eq '2029-06-30') $rhel6.ExtendedEnd
Check 'cache: RHEL 6 note rewritten' ($rhel6.Note -eq 'ELS ends Jun 2029 (verify: Red Hat may offer a longer ELS)') $rhel6.Note
Check 'cache: change reported' (@($changes | Where-Object { $_.Label -eq 'RHEL 6' -and $_.Field -eq 'ExtendedEnd' }).Count -eq 1)
Check 'cache: RHEL 8 missing in cache -> kept + reported' (@($changes | Where-Object { $_.Label -eq 'RHEL 8' -and $_.Field -eq 'Live' }).Count -eq 1)
(Get-Item (Join-Path $tmp 'eol_rhel.json')).LastWriteTime = (Get-Date).AddDays(-30)
$cat = Get-LifecycleCatalog -Settings (New-Settings) -Offline -CacheDir $tmp
Check 'cache: 30-day-old cache older than the shipped snapshot -> Snapshot' ($cat.Source['rhel'] -eq 'Snapshot') $cat.Source['rhel']
Check 'cache: reason names the newer snapshot' ($cat.Reason['rhel'] -like '*replaced by the newer shipped snapshot*') $cat.Reason['rhel']
$oldSnap = Join-Path $tmp 'old_snapshot.json'
[System.IO.File]::WriteAllText($oldSnap, ((Get-Content -LiteralPath (Join-Path (Join-Path $root 'Config') 'LifecycleSnapshot.json') -Raw) -replace '"generated_at": "[^"]+"', '"generated_at": "2025-01-01"'))
$cat = Get-LifecycleCatalog -Settings (New-Settings) -Offline -CacheDir $tmp -SnapshotPath $oldSnap
Check 'cache: 30-day-old cache newer than the snapshot -> StaleCache' ($cat.Source['rhel'] -eq 'StaleCache') $cat.Source['rhel']
Check 'cache: stale warning message' (@($cat.Messages | Where-Object { $_ -like 'rhel: using cached dates*' }).Count -eq 1)
Remove-Item (Join-Path $tmp 'eol_rhel.json')
$sd = New-Settings; $deb = $sd.OsLifecycle | Where-Object { $_.Label -eq 'Debian 12' }
[System.IO.File]::WriteAllText((Join-Path $tmp 'eol_debian.json'), '{"result":{"releases":[{"name":"12","eolFrom":"2026-06-10","eoesFrom":"2028-07-31"}]}}')
$null = Update-OsLifecycleFromCatalog -Settings $sd -Catalog (Get-LifecycleCatalog -Settings $sd -Offline -CacheDir $tmp)
Check "eoes rows: Debian 12 EndOfSupport <- eoesFrom (LTS end)" ($deb.EndOfSupport -eq '2028-07-31' -and $deb.ExtendedEnd -eq '') "$($deb.EndOfSupport)/$($deb.ExtendedEnd)"
Remove-Item (Join-Path $tmp 'eol_debian.json')

Write-Host "5. Online path with the internet unreachable: no exception, falls back to the snapshot"
$s = New-Settings
$s.LiveData = @{ Enabled = $true; CacheDays = 7; LifecycleApiBase = 'https://invalid.invalid/api/v1/products'; LifecycleLegacyBase = 'https://invalid.invalid/api'; TimeoutSec = 5; VmImportDocUrl = 'https://invalid.invalid/x'; VmImportDocKnownHash = '' }
$threw = $false
try { $cat = Get-LifecycleCatalog -Settings $s -CacheDir $tmp } catch { $threw = $true; Write-Host $_ }
Check 'online-fail: no exception' (-not $threw)
Check 'online-fail: windows from snapshot' ($cat.Source['windows'] -eq 'Snapshot')
Check 'online-fail: failure recorded once, host then skipped' (@($cat.Messages | Where-Object { $_ -like 'Could not read*' }).Count -eq 1 -and @($cat.Messages | Where-Object { $_ -like '*is not reachable*' }).Count -eq 1) (($cat.Messages) -join ' | ')
$vm = Test-VmImportDocChanged -Settings $s
Check 'vm-import: unreachable -> NotChecked' ($vm.Status -eq 'NotChecked') $vm.Status
$vm = Test-VmImportDocChanged -Settings $s -Offline
Check 'vm-import: offline -> NotChecked' ($vm.Status -eq 'NotChecked')
$txt = Get-VmImportDocText -Html '<html><nav>menu</nav><div id="main-col-body"><h1>Prereqs</h1><script>x=1</script><p>Up to 16 TiB &amp; 21 volumes</p></div></html>'
Check 'vm-import: page text extracted' ($txt -eq 'Prereqs Up to 16 TiB & 21 volumes') $txt

Write-Host "6. Support cut-off date"
$asOf = [datetime]'2026-11-04'
Check 'cutoff: fixed date' ((Resolve-SupportCutoffDate -Value '2027-12-31' -AsOf $asOf) -eq '2027-12-31')
Check 'cutoff: today+12m' ((Resolve-SupportCutoffDate -Value 'today+12m' -AsOf $asOf) -eq '2027-11-04')
Check 'cutoff: TODAY + 2y' ((Resolve-SupportCutoffDate -Value 'TODAY + 2y' -AsOf $asOf) -eq '2028-11-04')
$threw = $false; try { Resolve-SupportCutoffDate -Value '31.12.2027' -AsOf $asOf | Out-Null } catch { $threw = $true }
Check 'cutoff: invalid text throws' $threw

Write-Host "7. CPU architecture (Graviton) and sizing"
Check 'arm: Graviton processor' (Test-ArmInstanceType -InstanceType 'm7g.large' -PhysicalProcessor 'AWS Graviton3 Processor')
Check 'arm: Intel processor wins over name' (-not (Test-ArmInstanceType -InstanceType 'g6.xlarge' -PhysicalProcessor 'AMD EPYC 7R13 Processor'))
foreach ($t in @('a1.large', 't4g.micro', 'c8gn.large', 'x2gd.large', 'im4gn.large', 'g5g.xlarge')) { Check "arm by name: $t" (Test-ArmInstanceType -InstanceType $t -PhysicalProcessor '') }
foreach ($t in @('g6.xlarge', 'g4dn.xlarge', 'm8i-flex.large', 'u7in-16tb.224xlarge', 'm7i.large', 'inf2.xlarge')) { Check "x86 by name: $t" (-not (Test-ArmInstanceType -InstanceType $t -PhysicalProcessor '')) }
$Buffer = 1.15
$InstanceSpecs = @(
    [pscustomobject]@{ Type = 'm7g.large';  vCPU = 2; RAM = 8;  Family = 'GeneralPurpose'; Architecture = 'arm64' }
    [pscustomobject]@{ Type = 'm7i.xlarge'; vCPU = 4; RAM = 16; Family = 'GeneralPurpose'; Architecture = 'x86_64' }
    [pscustomobject]@{ Type = 't4g.small';  vCPU = 2; RAM = 2;  Family = 'Burstable';      Architecture = 'arm64' }
    [pscustomobject]@{ Type = 't3.medium';  vCPU = 2; RAM = 4;  Family = 'Burstable';      Architecture = 'x86_64' }
    [pscustomobject]@{ Type = 'm7i.4xlarge'; vCPU = 16; RAM = 64; Family = 'GeneralPurpose'; Architecture = 'x86_64' }
)
$AwsLiveFacts = $null
$bf = Get-BestFitInstance -RequiredCPU 1 -RequiredRAM_GiB 1 -PricingOS 'Linux'
Check 'sizing: Linux never gets Graviton' ($bf.Type -eq 't3.medium') $bf.Type
$bf = Get-BestFitInstance -RequiredCPU 1.5 -RequiredRAM_GiB 6 -PricingOS 'Linux'
Check 'sizing: general purpose x86' ($bf.Type -eq 'm7i.xlarge') $bf.Type
$n1 = Get-BestFitInstanceWithNIC -RequiredCPU 1.5 -RequiredRAM_GiB 6 -RequiredNICs 6
$n2 = Get-BestFitInstanceWithNIC -RequiredCPU 1.5 -RequiredRAM_GiB 6 -RequiredNICs 6
Check 'nic: upgrade works twice in a row' ($n1.Type -eq 'm7i.4xlarge' -and $n2.Type -eq 'm7i.4xlarge') "$($n1.Type) / $($n2.Type)"
Check 'nic: approximate table named' ($n1.NIC_Note -like '*approximate table*') $n1.NIC_Note
$AwsLiveFacts = [pscustomobject]@{ InstanceTypes = @{ 'm7i.xlarge' = @{ MaxENI = 6; Architectures = @('x86_64') } }; WindowsAmis = @(); Region = 'eu-central-1' }
Check 'eni: live value wins' ((Get-MaxENI -InstanceType 'm7i.xlarge') -eq 6)
Check 'eni: table otherwise' ((Get-MaxENI -InstanceType 'm7i.4xlarge') -eq 8)
$AwsLiveFacts = $null

Write-Host "8. OS mapping uses the lifecycle table (one source of truth)"
$StrategySettings = New-Settings; $StrategyAsOf = [datetime]'2026-11-04'
$m = Get-OSMapping -OSString 'Microsoft Windows Server 2022 (64-bit)'
Check 'osmap: 2022 date from table' ($m.EndOfLife -eq 'Oct 14, 2031') $m.EndOfLife
$m = Get-OSMapping -OSString 'Microsoft Windows Server 2008 R2 (64-bit)'
Check 'osmap: 2008 EOL' ($m.EndOfLife -eq 'Jan 14, 2020 (EOL)') $m.EndOfLife
$m = Get-OSMapping -OSString 'Red Hat Enterprise Linux 7 (64-bit)'
Check 'osmap: RHEL 7 in extended support' ($m.EndOfLife -eq 'Jun 30, 2024 (EOL; paid extended support to Jun 30, 2028)') $m.EndOfLife
$m = Get-OSMapping -OSString 'FortiGate'
Check 'osmap: vendor appliance' ($m.EndOfLife -eq 'N/A - Vendor appliance')
$StrategySettings.OsLifecycle | Where-Object { $_.Label -eq 'Windows Server 2022' } | ForEach-Object { $_['EndOfSupport'] = '2031-12-31' }
Check 'osmap: follows a live-updated date' ((Get-OSMapping -OSString 'Microsoft Windows Server 2022 (64-bit)').EndOfLife -eq 'Dec 31, 2031')

Write-Host "9. AWS facts and Data_Sources rows"
$f = Get-AwsLiveFacts -Region 'eu-central-1'
Check 'aws: disabled by default' (-not $f.Enabled -and $f.InstanceTypes.Count -eq 0)
$s = New-Settings
$cat = Get-LifecycleCatalog -Settings $s -Offline -CacheDir $tmp
$ch = @(Update-OsLifecycleFromCatalog -Settings $s -Catalog $cat)
$rows = @(New-DataSourceRows -LifecycleCatalog $cat -LifecycleChanges $ch -VmImportCheck (Test-VmImportDocChanged -Settings $s -Offline) -AwsFacts $f -CutoffSetting 'today+12m' -CutoffResolved '2027-11-04' -ArmTypesExcluded 42 -Settings $s)
Check 'datasources: rows built' ($rows.Count -ge 12) "got $($rows.Count)"
Check 'datasources: shape' (@($rows | Where-Object { -not ($_.PSObject.Properties['Area'] -and $_.PSObject.Properties['Source'] -and $_.PSObject.Properties['Status'] -and $_.PSObject.Properties['RetrievedAt'] -and $_.PSObject.Properties['Detail']) }).Count -eq 0)
Check 'datasources: snapshot labelled' (@($rows | Where-Object { $_.Status -eq 'Offline snapshot (verify)' }).Count -eq 9)

Write-Host "10. Live key alternatives (endoflife.date naming may vary) and Data_Sources wording"
$s = New-Settings
$w10 = $s.OsLifecycle | Where-Object { $_.Label -eq 'Windows 10' }
Check 'settings: Windows 10 has LiveAlt' ([string]$w10.LiveAlt -like '*10-22h2-e*') ([string]$w10.LiveAlt)
Check 'settings: Windows Server 2008 has no Live key' (-not ($s.OsLifecycle | Where-Object { $_.Label -like 'Windows Server 2008*' }).ContainsKey('Live'))
$alts = @(Get-LiveKeyAlternatives 'windows:10-22h2|windows:10-22h2-e|10-22h2-w')
Check 'alternatives: three parsed, product carried over' ($alts.Count -eq 3 -and $alts[2].Product -eq 'windows' -and $alts[2].Cycle -eq '10-22h2-w')
Check 'alternatives: garbage skipped' (@(Get-LiveKeyAlternatives '|:x|nocolon').Count -eq 0)
$win = '{"result":{"name":"windows","releases":[{"name":"10-22h2-e","eolFrom":"2025-10-14","eoesFrom":"2028-10-10"},{"name":"11-24h2-e","eolFrom":"2027-10-12"},{"name":"11-24h2-w","eolFrom":"2026-10-13"}]}}'
[System.IO.File]::WriteAllText((Join-Path $tmp 'eol_windows.json'), $win)
$cat = Get-LifecycleCatalog -Settings $s -Offline -CacheDir $tmp
$ch = @(Update-OsLifecycleFromCatalog -Settings $s -Catalog $cat)
Check 'alternatives: Windows 10 dates from 10-22h2-e' ($w10.LiveFound -and $w10.LiveRelease -eq '10-22h2-e' -and $w10.DateSource -like 'endoflife.date (cache*') "$($w10.LiveRelease) / $($w10.DateSource)"
$rows = @(New-DataSourceRows -LifecycleCatalog $cat -LifecycleChanges $ch -Settings $s)
$wrow = $rows | Where-Object { $_.Area -eq 'OS support dates - windows' }
Check 'datasources: one row per product' (@($rows | Where-Object { $_.Area -like 'OS support dates - *' -and $_.Area -notin @('OS support dates - updated', 'OS support dates - note', 'OS support dates - settings file') }).Count -eq 9)
Check 'datasources: windows row says which rows took dates' ($wrow.Detail -like '*taken from this source for: *Windows 10*') $wrow.Detail
Check 'datasources: windows row says which rows kept settings dates' ($wrow.Detail -like '*settings-file dates kept for: *Windows 7*' -and $wrow.Detail -notlike '*kept for: *Windows 10*') $wrow.Detail
Check 'datasources: no per-URL note rows' (@($rows | Where-Object { $_.Area -eq 'OS support dates - note' -and $_.Detail -like 'Could not read*' }).Count -eq 0)
Remove-Item (Join-Path $tmp 'eol_windows.json')
# stale windows cache without any 10-/11- release while the (older) snapshot has them -> snapshot
[System.IO.File]::WriteAllText((Join-Path $tmp 'eol_windows.json'), '{"result":{"releases":[{"name":"7-sp1","eolFrom":"2020-01-14"}]}}')
(Get-Item (Join-Path $tmp 'eol_windows.json')).LastWriteTime = (Get-Date).AddDays(-40)
$cat = Get-LifecycleCatalog -Settings (New-Settings) -Offline -CacheDir $tmp -SnapshotPath $oldSnap
Check 'stale cache missing chooser releases -> snapshot' ($cat.Source['windows'] -eq 'Snapshot' -and $cat.Reason['windows'] -like '*no release for a version-choice family*') "$($cat.Source['windows']) / $($cat.Reason['windows'])"
Remove-Item (Join-Path $tmp 'eol_windows.json')

Write-Host "11. Host unreachable: one try per host, legacy API only after 404 / bad format (mocked Invoke-WebRequest)"
$script:iwr = New-Object System.Collections.Generic.List[string]
$script:mode = 'transport'
function Invoke-WebRequest {
    param([string]$Uri, [switch]$UseBasicParsing, [int]$TimeoutSec, $ErrorAction)
    $script:iwr.Add($Uri)
    switch ($script:mode) {
        'transport' { throw (New-Object System.Net.WebException('The proxy tunnel request to proxy failed with status code 403.')) }
        'http404'   {
            if ($Uri -like '*/v1/*') { $ex = New-Object System.Exception('Response status code does not indicate success: 404 (Not Found).'); $ex | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = 404 }); throw $ex }
            return [pscustomobject]@{ Content = '[{"cycle":"2016","eol":"2027-01-12"},{"cycle":"2019","eol":"2029-01-09"}]' }
        }
        'http500'   { $ex = New-Object System.Exception('500'); $ex | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = 500 }); throw $ex }
        'html'      { if ($Uri -like '*/v1/*') { return [pscustomobject]@{ Content = '<html>maintenance</html>' } }; return [pscustomobject]@{ Content = '[{"cycle":"2019","eol":"2029-01-09"}]' } }
    }
}
$sOn = New-Settings
$sOn.LiveData = @{ Enabled = $true; CacheDays = 7; LifecycleApiBase = 'https://eol.test/api/v1/products'; LifecycleLegacyBase = 'https://eol.test/api'; TimeoutSec = 5; VmImportDocUrl = ''; VmImportDocKnownHash = '' }
$liveTmp = Join-Path $tmp 'live'; New-Item -ItemType Directory -Path $liveTmp | Out-Null
$cat = Get-LifecycleCatalog -Settings $sOn -CacheDir $liveTmp 6>$null
Check 'unreachable: only ONE request for 9 products' ($script:iwr.Count -eq 1) ("requests: " + $script:iwr.Count)
Check 'unreachable: every product from the snapshot' (@($cat.Source.Values | Where-Object { $_ -ne 'Snapshot' }).Count -eq 0)
Check 'unreachable: short reason per product' ($cat.Reason['rhel'] -like 'eol.test not reachable*') $cat.Reason['rhel']
$script:iwr.Clear(); $script:mode = 'http404'
$cat = Get-LifecycleCatalog -Settings $sOn -CacheDir $liveTmp 6>$null
Check 'http 404: legacy API tried and used' ($cat.Source['windows-server'] -eq 'Live' -and $script:iwr.Count -eq 18) ("source=" + $cat.Source['windows-server'] + " requests=" + $script:iwr.Count)
Get-ChildItem $liveTmp | Remove-Item
$script:iwr.Clear(); $script:mode = 'http500'
$cat = Get-LifecycleCatalog -Settings $sOn -CacheDir $liveTmp 6>$null
Check 'http 500: legacy API NOT tried, snapshot used' ($script:iwr.Count -eq 9 -and $cat.Source['rhel'] -eq 'Snapshot' -and $cat.Reason['rhel'] -like 'HTTP 500*') ("requests=" + $script:iwr.Count + " " + $cat.Reason['rhel'])
$script:iwr.Clear(); $script:mode = 'html'
$cat = Get-LifecycleCatalog -Settings $sOn -CacheDir $liveTmp 6>$null
Check 'bad format: legacy API tried' ($script:iwr.Count -eq 18 -and $cat.Source['windows-server'] -eq 'Live') ("requests=" + $script:iwr.Count)
Remove-Item -Path Function:\Invoke-WebRequest
Get-ChildItem $liveTmp | Remove-Item

Write-Host "12. Dates cannot shift with time zones; relative dates; VM Import page footer"
$j = ConvertFrom-LiveJsonText '{"a":"2029-01-09T23:30:00-05:00","b":"2029-01-09T00:30:00+09:00","c":"2029-01-09"}'
Check 'tz: offset timestamp keeps its date' ((ConvertTo-LiveDate $j.a) -eq '2029-01-09' -and (ConvertTo-LiveDate $j.b) -eq '2029-01-09' -and (ConvertTo-LiveDate $j.c) -eq '2029-01-09') "$($j.a) / $($j.b)"
Check 'relative: today-36m' ((Resolve-SupportCutoffDate -Value 'today-36m' -AsOf $asOf) -eq '2023-11-04')
Check 'relative: today - 1y' ((Resolve-SupportCutoffDate -Value 'today - 1y' -AsOf $asOf) -eq '2025-11-04')
$page1 = '<html><nav>menu</nav><div id="main-col-body"><h1>Prereqs</h1><p>Up to 16 TiB</p></div><div id="main-col-footer">Did this page help you? (c) 2025 Amazon</div></html>'
$page2 = '<html><nav>other menu</nav><div id="main-col-body"><h1>Prereqs</h1><p>Up to 16 TiB</p></div><div id="main-col-footer">Did this page help you? (c) 2026 Amazon</div></html>'
$t1 = Get-VmImportDocText -Html $page1
Check 'vm-import: footer not in the text' ($t1 -eq 'Prereqs Up to 16 TiB') $t1
Check 'vm-import: footer change keeps the same text' ($t1 -eq (Get-VmImportDocText -Html $page2))
Check 'vm-import: "Document Conventions" cut' ((Get-VmImportDocText -Html '<main><p>Rules</p><h2>Document Conventions</h2><p>x</p></main>') -eq 'Rules')

Write-Host "13. Windows AMI names (SSM /aws/service/ami-windows-latest)"
$onlyR2 = @('/aws/service/ami-windows-latest/Windows_Server-2012-R2_RTM-English-64Bit-Base', '/aws/service/ami-windows-latest/Windows_Server-2016-English-Full-Base', '/aws/service/ami-windows-latest/EC2LaunchV2-Windows_Server-2022-English-Full-Base')
$onlyRtm = @('/aws/service/ami-windows-latest/Windows_Server-2012-RTM-English-64Bit-Base')
Check 'ami: 2012 R2 found' (Test-WindowsAmiOffered -AmiNames $onlyR2 -Version '2012 R2')
Check 'ami: 2012 does NOT match R2 names' (-not (Test-WindowsAmiOffered -AmiNames $onlyR2 -Version '2012'))
Check 'ami: 2012 RTM found' (Test-WindowsAmiOffered -AmiNames $onlyRtm -Version '2012')
Check 'ami: 2012 R2 does NOT match RTM names' (-not (Test-WindowsAmiOffered -AmiNames $onlyRtm -Version '2012 R2'))
Check 'ami: 2016 / 2022 found, 2019 / 2025 not' ((Test-WindowsAmiOffered -AmiNames $onlyR2 -Version '2016') -and (Test-WindowsAmiOffered -AmiNames $onlyR2 -Version '2022') -and -not (Test-WindowsAmiOffered -AmiNames $onlyR2 -Version '2019') -and -not (Test-WindowsAmiOffered -AmiNames $onlyR2 -Version '2025'))
$StrategySettings = New-Settings; $StrategyAsOf = [datetime]'2026-11-04'
$AwsLiveFacts = [pscustomobject]@{ InstanceTypes = @{}; WindowsAmis = $onlyR2; Region = 'eu-central-1' }
Check 'osmap ami: 2012 R2 VM -> offered' ((Get-OSMapping -OSString 'Microsoft Windows Server 2012 R2 (64-bit)').RecommendedAMI -like '*standard AMI offered*')
Check 'osmap ami: 2012 VM -> not offered' ((Get-OSMapping -OSString 'Microsoft Windows Server 2012 (64-bit)').RecommendedAMI -like '*no standard AMI*')
$ds = @(New-DataSourceRows -AwsFacts ([pscustomobject]@{ Enabled = $true; Region = 'eu-central-1'; InstanceTypes = @{}; RegionServices = @(); WindowsAmis = $onlyR2; Status = @{ InstanceTypes = 'x'; RegionServices = 'x'; WindowsAmis = 'Live' }; Messages = @() }))
$amiRow = $ds | Where-Object { $_.Area -eq 'Windows Server AMIs' }
Check 'datasources ami: 2012 NOT, 2012 R2 offered' ($amiRow.Detail -like '*2012: NOT offered*' -and $amiRow.Detail -like '*2012 R2: offered*') $amiRow.Detail
$AwsLiveFacts = $null

Write-Host "14. Snapshot and settings content (review findings)"
$snap = Get-Content -LiteralPath (Join-Path (Join-Path $root 'Config') 'LifecycleSnapshot.json') -Raw | ConvertFrom-Json
$wr = @($snap.products.windows.result.releases)
Check 'snapshot: LTSB 2015 present' (@($wr | Where-Object { $_.name -eq '10-1507-e-lts' -and $_.eolFrom -eq '2025-10-14' }).Count -eq 1)
Check 'snapshot: 22H2 note says consumer ESU ended' ([string](($wr | Where-Object { $_.name -eq '10-22h2' }).note) -like '*consumer ESU ended 13 Oct 2026*')
$ur = @($snap.products.ubuntu.result.releases | Where-Object { $_.eoesFrom })
Check 'snapshot: Ubuntu ESM ends in April' (@($ur | Where-Object { [string]$_.eoesFrom -notmatch '-04-30$' }).Count -eq 0)
$s = New-Settings
$vc = @($s.VersionChoices)
Check 'settings: Windows 10/11 hide releases ended > 36 months ago' (@($vc | Where-Object { $_.Family -like 'Windows 1*' -and $_.HideEndedBefore -eq 'today-36m' }).Count -eq 2)
Check 'settings: OsTemplate keeps bitness {1}' (@($vc | Where-Object { [string]$_.OsTemplate -notlike '*{1}*' }).Count -eq 0)
Check 'settings: TimeoutSec 20' ([int]$s.LiveData.TimeoutSec -eq 20)

Remove-Item -Recurse -Force $tmp
Write-Host ("{0} passed, {1} failed" -f $script:pass, $script:fail) -ForegroundColor $(if ($script:fail) { 'Red' } else { 'Green' })
exit $script:fail
