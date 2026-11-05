# Handoff – AWS Migration Strategy v3.0

Repo `soghosh-f97/Capitex-AI`, branch `claude/festive-cerf-j33av2`, folder `AWS_Migration_Strategy_Code/`. Latest code commit `2d944e5` (`ae88df2` adds Word copies in `Word_documents_v3.0/`). v2.2 baseline = `076871a`. PowerShell 5.1 + 7, UTF-8 with BOM, ImportExcel ≥ 7.

## 1. Files created / changed
- **New `Modules/LiveDataProviders.ps1`** – all outside calls: OS support dates (endoflife.date → cache → snapshot → settings file), `today±Nm/y` dates, VM Import page hash check, optional AWS CLI facts, Data_Sources rows, live-date impact comparison.
- **New `Config/LifecycleSnapshot.json`** – offline fallback (API v1 shape). Written from knowledge, not downloaded; entries marked "verify".
- **New `Tests/Test-LiveData.ps1`** – 126 offline checks; exit code = failures.
- **`Map_On-Prem_to_AWS.ps1`** – params `-Offline`, `-UseAwsApis`, `-AwsRegion`, `-ReuseChoicesFrom`; wiring; ImportExcel version check.
- **`MigrationStrategyGenerator.ps1`** – exact-version chooser (`Get-VersionChoiceCatalog`, `Get-VersionScenarioRows`, `Import-WorkbookChoices`), lifecycle override, new optional keys.
- **`ExcelExporter.ps1`** – Exact_OS_Version dropdown; sheets OS_Versions, Strategy_Versions (hidden), Data_Sources, Live_Date_Changes; hidden helper columns. **`SummaryBuilder.ps1`** – "not checked yet" count.
- **`PricingIndex.ps1` / `InstanceSizing.ps1`** – ARM/Graviton excluded; exact NIC limits with `-UseAwsApis`. **`OSMapping.ps1`** – EOL text from the same `OsLifecycle` table.
- **Settings** – `SettingsVersion 3.0`, `LiveData`, `VersionChoices`, `Live`/`LiveAlt`/`LiveEndField` keys. **README.md** updated; **DESIGN_v3.md** = build contract.

## 2. Decisions
- Never crash: no network / credentials still gives a workbook; Data_Sources shows each source.
- Offline = v2.2 (snapshot dates equal settings dates). VM Import rules have no API → page hash vs `LiveData.VmImportDocKnownHash`, settings never auto-edited. AWS APIs optional (`-UseAwsApis`).
- Version chooser = pre-computed scenarios (VM × OS field × option) looked up by formula: picks update dates, status, strategy, target, readiness and counts live. EC2 sizing/costs never follow picks.
- Rows keyed by hidden `VM_Index`, not VM name; result cells `INDEX(...,VM_Index)` → sort-safe.
- Graviton excluded for all VMs (imported VMware VMs are x86).
- Live_Date_Changes: each run recomputes with settings-file dates and lists VMs whose status/strategy/target/readiness differs.

## 3. Bugs found and fixed
- EPPlus returns 0 for a failed MATCH → cached `#VALUE!`: helper `Version_Lookup_Row` + `IF(row>0,INDEX(...),base)`.
- Name-keyed lookup broke on `* ? ~`, numeric-looking, `=`-prefixed and duplicate names: `VM_Index` key; names written as text.
- VMs with different families in the two OS fields got the wrong list / false note: per-field family columns + one `INDIRECT("VerList_"&…)` validation.
- Blank/junk picks: treated as "Not checked yet" / "not an option" note. 32-bit lost on pick: `{1}` bitness in `OsTemplate`.
- Also fixed: field-unaware highlight/counts; Summary built before scenarios; relative cut-off misreported; ~9 min network hang (host now skipped after first no-answer); 2012 vs 2012 R2 AMI patterns; VM Import page text extraction; stale cache beating newer snapshot; PS7 one-item JSON array; time-zone date shift; NIC upgrade mutating shared objects.

## 4. Unfinished / fragile
- **Live sources never reached** (endoflife.date, docs.aws.amazon.com blocked). Windows release names (`11-24h2-e`, `10-22h2`…) assumed; `LiveAlt` covers variants. First run on a connected machine and check Data_Sources.
- Least certain snapshot dates: Ubuntu ESM (30 Apr), RHEL 6 ELS end. `-UseAwsApis` tested only with a fake `aws` CLI; dropdown never clicked in real Excel (XML + LibreOffice only); test inputs lack vSphere notes/tags → R04 under-tested.
- `AWS_Migration_Rules_At_A_Glance.docx` still describes v2.1 (14 rules, USB, R05 Replatform). Pre-existing, unchanged: buffer default 15% (docs say 10%); EBS totals include powered-off VMs. Word copies: paste back via Notepad as UTF-8 with BOM.

## 5. Environment / run
- Tests: `pwsh -NoProfile -File AWS_Migration_Strategy_Code/Tests/Test-LiveData.ps1` → 126/126.
- Run: `.\Map_On-Prem_to_AWS.ps1 -InputExcelPath <site>_RVTools_export.xlsx [-Offline] [-UseAwsApis -AwsRegion eu-central-1] [-ReuseChoicesFrom <reviewed.xlsx>]`. Output `AWS_Mapping_Output_<site>.xlsx` next to the input.
- Test harness was in the session scratchpad (ephemeral, not in repo, client-derived): ImportExcel shim on EPPlus 4.5.3.3 + `Microsoft.Extensions.*` 3.1.0 (NuGet), 4 rebuilt RVTools inputs (976 VMs), synthetic pricing CSVs (real header), per-VM diff and LibreOffice recalc scripts.
- Last verified: no picks → identical to v2.2 on 4 sites, 0 errors in 21,718 formulas; altered dates → Live_Date_Changes matched an independent recount.

## 6. Differences from the reference rule set
Reference: R01 Template; R02–R10 first-match (R05 Database = Manual Review); R11 disk / R12 DMZ readiness; R13 Powered off; no USB rule.

- **R01–R13 unchanged**; no rule-setting value changed vs v2.2 (`DatabaseStrategy 'Manual Review'`, `PoweredOffAction 'ClientReview'`, cut-off `2027-12-31`). **No USB rule or setting** (vUSB not read).
- Added, not a rule: "OS version" pre-step row on Strategy_Rules; with a pick, R06–R09 use the picked version's OS text and dates. R06 dates can change at run time from live data. Sizing (not a rule): ARM never chosen, so some EC2 types/costs differ from v2.2.
- Settings vs v2.2: Ubuntu 18.04–24.04 `ExtendedEnd` 30 May → **30 Apr**; Windows 10 and RHEL 6 notes reworded; no `Live` key on Server 2008.