# v3.0 design contract (live data + exact OS version chooser)

This file is the contract between the two build streams (A = live data, B = version chooser).
Stick to the names and shapes below so the two halves merge without edits.

## Goals
1. Make everything that can come from a live source come from it, with a safe fallback chain and a
   visible record of where each value came from (new workbook sheet `Data_Sources`).
2. Let the client choose the exact Windows version for VMs whose RVTools label is ambiguous
   (Windows 10, Windows 11, Windows Server 2012 vs 2012 R2, Windows Server "2016 or later") and see
   support dates, OS status AND the strategy/target/readiness update live in the workbook.
3. Never break: no network, no AWS credentials, or an API change must still produce a workbook
   (fallback to cache -> shipped snapshot -> settings-file dates), with warnings, never a crash.
4. Keep Windows PowerShell 5.1 + PowerShell 7 compatibility, UTF-8 with BOM, ImportExcel/EPPlus 4.5.
5. Default settings must reproduce v2.2 results except where a live date or a client choice changes them.

## Live sources (stream A)
| What | Source | Credentials | Fallback |
|---|---|---|---|
| OS support dates (`OsLifecycle` EndOfSupport / ExtendedEnd / Note) | endoflife.date API v1 `https://endoflife.date/api/v1/products/<product>` (`result.releases[]`: `name`, `label`, `releaseDate`, `eoasFrom`, `eolFrom`, `eoesFrom`, `isLts`). Also accept the legacy shape `https://endoflife.date/api/<product>.json` (array of `cycle`, `eol`, `extendedSupport`, `support`) | none | cache file (<= CacheDays, else stale cache with warning) -> `Config\LifecycleSnapshot.json` -> settings-file dates |
| Version chooser options (Windows 10/11 releases, Windows Server versions) | same API, products `windows` and `windows-server` | none | same |
| EC2 prices, instance vCPU/RAM (already live) | AWS bulk pricing CSV | none | unchanged (7-day cache) |
| CPU architecture (Graviton fix) | pricing CSV column `Physical Processor` (contains `Graviton` for ARM) + `Processor Architecture`; name regex `^[a-z]+\d+g[a-z]*\.` and `^a1\.` as backup | none | regex |
| Max NICs per instance type | `aws ec2 describe-instance-types` -> `NetworkInfo.MaximumNetworkInterfaces`, `ProcessorInfo.SupportedArchitectures` | AWS CLI + read-only creds, only with `-UseAwsApis` | existing `$MaxEniBySize` table |
| Region service availability (WorkSpaces, RDS, DocumentDB, EVS) | SSM public parameters `/aws/service/global-infrastructure/regions/<region>/services` | creds, only with `-UseAwsApis` | not checked (reported as "not checked") |
| Windows AMI availability | SSM public parameters path `/aws/service/ami-windows-latest` | creds, only with `-UseAwsApis` | not checked |
| VM Import rules page (supported OS list, 16 TiB / 8 TiB / 21 volumes, 32-bit) | `https://docs.aws.amazon.com/vm-import/latest/userguide/prerequisites.html` - NO API: page text is hashed and compared with `LiveData.VmImportDocKnownHash`; a change raises a warning "review the VM Import settings" | none | "not checked" |

endoflife.date date semantics: `eolFrom` = end of support (the value the tool calls EndOfSupport),
`eoesFrom` = end of paid extended support (ExtendedEnd), `eoasFrom` = end of mainstream/active support.
Values can be a date string, `true`/`false` or null -> treat non-dates as '' (unknown).

## Settings (v3.0)
New keys (all OPTIONAL in the loader; missing = feature off / defaults):
```powershell
LiveData = @{
    Enabled              = $true      # $false = never call the internet; use snapshot + settings dates
    CacheDays            = 7
    LifecycleApiBase     = 'https://endoflife.date/api/v1/products'
    LifecycleLegacyBase  = 'https://endoflife.date/api'
    VmImportDocUrl       = 'https://docs.aws.amazon.com/vm-import/latest/userguide/prerequisites.html'
    VmImportDocKnownHash = ''         # SHA-256 of the page text when the VM Import settings were last reviewed ('' = not yet recorded)
    TimeoutSec           = 30
}
SupportCutoffDate = '2027-12-31'      # unchanged default; ALSO accepts 'today+<n>m' (e.g. 'today+12m') = run date + n months
```
`OsLifecycle` rows get an optional `Live = '<product>:<cycle>'` key, e.g. `Live = 'rhel:7'`,
`Live = 'windows-server:2012-r2'`. Rows without it keep their static dates.
Stream A owns `LiveData`, `Live` keys and the `SupportCutoffDate` parser.

Stream B owns `VersionChoices`:
```powershell
VersionChoices = @(
    @{ Match = 'Windows Server\D*2016 or later'; Family = 'Windows Server 2016 or later'; Product = 'windows-server'; CyclePattern = '^(2016|2019|2022|2025)$';      OsTemplate = 'Microsoft Windows Server {0} (64-bit)'; Help = 'Inside the VM run winver, or PowerShell (Get-ComputerInfo).OsName' }
    @{ Match = 'Windows Server\D*2012';          Family = 'Windows Server 2012 / 2012 R2'; Product = 'windows-server'; CyclePattern = '^2012(-r2)?$';                OsTemplate = 'Microsoft Windows Server {0} (64-bit)'; Help = '...' }
    @{ Match = 'Windows 10\b';                   Family = 'Windows 10';                    Product = 'windows';        CyclePattern = '^10-';                       OsTemplate = 'Microsoft Windows 10 (64-bit)'; Help = 'Inside the VM run winver: it shows the version (e.g. 22H2) and edition' }
    @{ Match = 'Windows 11\b';                   Family = 'Windows 11';                    Product = 'windows';        CyclePattern = '^11-';                       OsTemplate = 'Microsoft Windows 11 (64-bit)'; Help = '...' }
)
```
Order matters (first match). "2016 or later" must be above any generic Windows Server row.

## Interfaces
### Stream A - `Modules\LiveDataProviders.ps1` (new) and changes to pricing/sizing/OS mapping
```powershell
Get-LifecycleCatalog -Settings $s [-Offline] [-CacheDir <dir>]
  -> [pscustomobject]@{
       Products = @{ '<product>' = @( [pscustomobject]@{ Product; Cycle; Label; ReleaseDate; MainstreamEnd; EndOfSupport; ExtendedEnd; IsLts } ) }   # dates 'yyyy-MM-dd' or ''
       Source   = @{ '<product>' = 'Live' | 'Cache' | 'StaleCache' | 'Snapshot' }
       RetrievedAt = @{ '<product>' = 'yyyy-MM-dd HH:mm' }
       Messages = [System.Collections.Generic.List[string]]
     }
  Products fetched = every product named in OsLifecycle Live keys + VersionChoices.Product.
Update-OsLifecycleFromCatalog -Settings $s -Catalog $c
  -> mutates each OsLifecycle row with a Live key: EndOfSupport/ExtendedEnd from the catalog when found,
     adds 'DateSource' ('endoflife.date (live yyyy-MM-dd)' / 'endoflife.date (cache ...)' / 'snapshot' / 'settings file');
     Note text rebuilt only when it was a pure date sentence; returns list of changes (Label, Field, Old, New) for Data_Sources.
Resolve-SupportCutoffDate -Value <string> -AsOf <datetime> -> 'yyyy-MM-dd'
Test-VmImportDocChanged -Settings $s [-Offline] -> [pscustomobject]@{ Status = 'Unchanged'|'Changed'|'NotRecorded'|'NotChecked'; Hash; Detail }
Get-AwsLiveFacts -Region <r> [-Enabled] -> [pscustomobject]@{ Enabled; InstanceTypes = @{ '<type>' = @{ MaxENI; Architectures } }; RegionServices = @(...); WindowsAmis = @(...); Messages }
New-DataSourceRows ... -> array of [pscustomobject]@{ Area; Source; Status; RetrievedAt; Detail }   # rows for the Data_Sources sheet
```
Graviton fix: instance specs get `Architecture` ('x86_64' | 'arm64'); sizing ALWAYS excludes arm64 for
VMware VMs (they are x86), for Linux and Windows. `Get-MaxENI` uses `$AwsFacts` when available.
`OSMapping.ps1` must stop hard-coding EOL dates: EndOfLife text comes from the (live-updated) OsLifecycle
via `Get-AWSOSLifecycle` (engine function) so there is one source of truth.

### Stream B - engine, exporter, summary
```powershell
Get-VersionChoiceCatalog -Settings $s -LifecycleCatalog $c
  -> @( [pscustomobject]@{ Family; Match; Help; Options = @( [pscustomobject]@{ Option; OsString; EndOfSupport; ExtendedEnd; MainstreamEnd; DateSource } ) } )
  Option = client-facing text, e.g. 'Windows 11 24H2 - Enterprise / Education', 'Windows 10 22H2', 'Windows Server 2012 R2'.
Get-VersionScenarioRows -StrategyInventory $inv -Settings $s -VersionCatalog $vc -PreferRelocateForLegacyOS:$x -AsOf $d
  -> rows with the SAME columns as Get-MigrationStrategyRows plus 'ScenarioKey' = '<VMName>|<OS field label>|<Option>'
     (OS field label = 'VMware Tools' | 'Configuration file', the workbook text), one row per VM x field x option where that
     field's OS matches a VersionChoices family. The engine evaluates the rules with OS = OsString and lifecycle dates
     forced to the option's dates (lifecycle override), OS_Source = 'Client choice (workbook)'.
Import-WorkbookChoices -Path <previous output .xlsx> -> @{ '<vmname lower>' = @{ OSField = '...'; ExactVersion = '...' } }
Export-MigrationReport ... -VersionCatalog $vc -VersionScenarios $rows -PriorChoices $pc -DataSources $ds   # all new params optional
```
Workbook (B):
- `Migration_Strategy` gets column `Exact_OS_Version` right after `OS_Used`: dropdown (yellow) ONLY on rows whose OS
  matches a family in EITHER field; other rows show 'n/a' (grey, no dropdown). Blank/'Not checked yet' = use the
  RVTools label (today's result). Plus `Version_Note` column: what to do / what was applied.
- Every result column formula: `IFERROR(INDEX(<scenario col>, MATCH($A&"|"&<OS_Field_Applied>&"|"&<Exact_OS_Version>, <scenario key col>, 0))&"", <existing formula>)` with bounded ranges.
- Hidden sheet `Strategy_Versions` holds the scenario rows.
- Visible sheet `OS_Versions` (placed right after `OS_Choice`): plain-language 3-step instructions, then one block per family
  listing every option with End of support / Extended support end / Source; these blocks are the dropdown source ranges.
- Visible sheet `Data_Sources`: rows from `-DataSources`.
- Strategy_Rules / Summary_Metrics counts already use COUNTIF on Migration_Strategy -> they update automatically; keep it so.
- Formulas must recalc with 0 errors in LibreOffice and Excel.

### Main script (stream A owns `Map_On-Prem_to_AWS.ps1`)
New params: `-Offline`, `-UseAwsApis`, `-AwsRegion` (default 'eu-central-1'), `-ReuseChoicesFrom <xlsx>`.
Order: load settings -> `Get-LifecycleCatalog` -> `Update-OsLifecycleFromCatalog` -> resolve cutoff -> VM Import doc check ->
pricing -> `Get-AwsLiveFacts` -> sizing -> strategy rows (as today) -> `Get-VersionChoiceCatalog` -> `Get-VersionScenarioRows`
-> `Import-WorkbookChoices` (if given) -> `New-DataSourceRows` -> `Export-MigrationReport` with the new params.

## Test harness (stream B builds it, both use it): scratchpad `harness\`
- `harness\ImportExcel\` minimal shim module (EPPlus 4.5.3.3 from NuGet) implementing exactly the cmdlets/params the code uses.
- `harness\inputs\<site>_RVTools_export.xlsx` rebuilt from the four sample outputs (vInfo + vDisk columns the code reads).
- `harness\pricing\` small pricing CSVs with the REAL header (from the live file) and representative rows, placed as the cache files.
- `harness\run_baseline.ps1` (v2.2 from git commit 076871a) and `harness\run_v3.ps1`, plus a Python comparer (openpyxl) and a
  LibreOffice headless recalc check (0 formula errors).

## Offline snapshot dates (Config\LifecycleSnapshot.json) - from vendor lifecycle pages, as known 2026-11-04; verify live
Windows client (`windows`): 10-22h2 (all editions) EOS 2025-10-14, ESU (commercial, paid) to 2028-10-10;
10-21h2 Ent/Edu 2024-06-11, Home/Pro 2023-06-13; 10 Enterprise LTSB 2016 2026-10-13; 10 Enterprise LTSC 2019 2029-01-09;
10 Enterprise LTSC 2021 2027-01-12; 10 IoT Enterprise LTSC 2021 2032-01-13;
11-21h2 Home/Pro 2023-10-10, Ent/Edu 2024-10-08; 11-22h2 H/P 2024-10-08, E/E 2025-10-14; 11-23h2 H/P 2025-11-11, E/E 2026-11-10;
11-24h2 H/P 2026-10-13, E/E 2027-10-12; 11-25h2 H/P 2027-10-12, E/E 2028-10-10; 11 Enterprise LTSC 2024 2029-10-09; 11 IoT Enterprise LTSC 2024 2034-10-10.
Windows Server (`windows-server`): 2012 and 2012-r2 EOS 2023-10-10, ESU 2026-10-13; 2016 2027-01-12; 2019 2029-01-09; 2022 2031-10-14; 2025 2034-11-14.
Other products: take the dates already in the v2.2 OsLifecycle table.
Every snapshot entry carries `"source": "offline snapshot 2026-11-04 - verify"`.
