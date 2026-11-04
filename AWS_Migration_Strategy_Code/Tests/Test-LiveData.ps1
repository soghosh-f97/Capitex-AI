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
Check 'cache: RHEL 6 note rewritten' ($rhel6.Note -eq 'ELS ends Jun 2029') $rhel6.Note
Check 'cache: change reported' (@($changes | Where-Object { $_.Label -eq 'RHEL 6' -and $_.Field -eq 'ExtendedEnd' }).Count -eq 1)
Check 'cache: RHEL 8 missing in cache -> kept + reported' (@($changes | Where-Object { $_.Label -eq 'RHEL 8' -and $_.Field -eq 'Live' }).Count -eq 1)
(Get-Item (Join-Path $tmp 'eol_rhel.json')).LastWriteTime = (Get-Date).AddDays(-30)
$cat = Get-LifecycleCatalog -Settings (New-Settings) -Offline -CacheDir $tmp
Check 'cache: 30-day-old cache -> StaleCache' ($cat.Source['rhel'] -eq 'StaleCache') $cat.Source['rhel']
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
Check 'online-fail: failure messages recorded' (@($cat.Messages | Where-Object { $_ -like 'Could not read*' }).Count -ge 2)
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

Remove-Item -Recurse -Force $tmp
Write-Host ("{0} passed, {1} failed" -f $script:pass, $script:fail) -ForegroundColor $(if ($script:fail) { 'Red' } else { 'Green' })
exit $script:fail
