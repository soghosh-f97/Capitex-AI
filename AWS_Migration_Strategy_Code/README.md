# AWS Recommendation Engine – Migration Strategy add-on

This follows the Azure team's `Migration_Strategy_Generator.ps1` pattern:
- settings in one place;
- a first-match decision chain;
- review flags and design notes that never change the strategy;
- a rule catalogue written into the workbook.

It is wired into the existing AWS tool (`Map_On-Prem_to_AWS.ps1`).

## Version 3.0

| Change | Detail |
|---|---|
| Exact OS version chooser | RVTools shows only a family label for some Windows VMs: "Windows 10", "Windows 11", "Windows Server 2012" (used for both 2012 and 2012 R2) and "Windows Server 2016 or later". These VMs now get a yellow dropdown in the new column **Exact_OS_Version** on Migration_Strategy (e.g. *Windows 10 22H2 - all editions*, *Windows 11 24H2 - Enterprise / Education*, *Windows Server 2012 R2*, *Windows Server 2019*). Picking a version updates that VM's **support dates, OS status, strategy, rule, AWS target, readiness, actions and notes** live, and the counts on Strategy_Rules / Summary_Metrics follow. *Not checked yet* = the RVTools label is used (today's result). Example: "2016 or later" → *Windows Server 2016* changes Rehost (R09) to Replatform (R06), because 2016 support ends 12 Jan 2027, before the cut-off |
| How it works | The engine pre-computes every option (`Get-VersionScenarioRows`: one row per VM × OS field × option, hidden sheet **Strategy_Versions**); each result column of Migration_Strategy looks up `<VM>|<OS field>|<version>` there and falls back to the normal result. The options and their dates come from the lifecycle data (`Get-VersionChoiceCatalog`), so new Windows releases appear automatically |
| New sheet OS_Versions | Plain-language *What to do* box (3 steps: find the yellow cells, run `winver` / `(Get-ComputerInfo).OsName` inside the VM, pick the version), a live count of VMs still waiting for a choice, and one table per family with every version, its support end, paid extended support end, status today and the source of the dates. These tables are the dropdown lists |
| New columns | **Exact_OS_Version** (dropdown), **OS_Extended_Support_End** (ESU / ELS / LTSS end) and **Version_Note** (what to check, or what the chosen version changed and why) |
| Choices survive a re-run | `-ReuseChoicesFrom <previous output.xlsx>` (`Import-WorkbookChoices`) pre-fills Exact_OS_Version and OS_Field_To_Use; a choice that is no longer an option is reset with a warning |
| Settings | New optional blocks `VersionChoices` (section 6b: which RVTools labels get a dropdown, which lifecycle product / releases feed it, the OS text the rules use) and `LiveData`; OsLifecycle rows can carry `Live = '<product>:<cycle>'`. `SupportCutoffDate` also accepts `today+<n>m`. All optional: an old settings file still loads (the chooser is then off) |
| Data_Sources sheet | Where each live value came from (live API, cache, shipped snapshot or settings file) and any warning |
| Summary_Metrics | New live count "VMs - Exact OS version not checked yet"; strategy / readiness values that only a version choice can produce also get a row |

## Version 2.2

| Change | Detail |
|---|---|
| Database rule | **R05 now gives Manual Review** for every database server (setting `DatabaseStrategy = 'Manual Review'`). AWS_Target shows the managed service to consider ("To be decided after review: Amazon RDS for SQL Server or Amazon EC2"); the action and Client_Confirmation list what to review. `'Replatform'` or `'Rehost'` can still be set |
| USB check removed | The USB readiness check is deleted from the engine, the main script, the exporter and the settings file (`UsbReadinessCheck` is gone). The **vUSB tab is no longer read** |
| Renumbered | Rules stay in run order without gaps: DMZ server is now **R12**, Powered off is now **R13** (13 rules in total) |
| Settings | 59 → 58 (`SettingsVersion = '2.2'`) |
| Results | 70 database VMs change (68 decided by R05 become Manual Review; 2 powered-off databases now show "if still needed: Manual Review (R05)"); IIS-PL-GOS-1 (the only VM with a USB device) goes from Rehost, Blocked to Rehost, Ready. No other VM changes |

## Version 2.1

| Change | Detail |
|---|---|
| Workbook tabs | The **Parameter_Map** and **Strategy_Settings** tabs are no longer written, so the settings values are no longer exposed to the client in the workbook (they stay in the `.psd1`) |
| Columns | **Parameter_Map_Ref** removed from Migration_Strategy (and the hidden sheets); **Register row** removed from Strategy_Rules |
| Settings file | Every setting has a one-line comment saying what the value is and what it does. Values unchanged. The three internal register settings (`ParameterMapDecisionSupport`, `...Default`, `ParameterMapSiteCountDecidedRows`) are now optional |
| Fixes | Strategy_Rules VM column: a precedence bug (`$i - 1` in an indexer) printed a red "op_Subtraction" error on every v2.0 run and left the column uncentred - fixed. Summary COUNTIF labels are escaped. With `DatabaseStrategy = 'Rehost'`, R05 now adds the VM Import blocker for an OS AWS cannot import |
| Results | Identical to v2.0 for every VM |

## Version 2.0 (condensed rule set)

| Change | Detail |
|---|---|
| Rules renumbered | **R01–R14 in the order they run**, so the numbers read 1, 2, 3 … without gaps |
| Rules removed | Decommission notes, plant floor, domain controllers (already off), infrastructure services (row 16), ISO connected (row 24), OS-conflict Manual Review (old R98) |
| Rules merged | Vendor product + discontinued vendor + appliance tag → **R04**; 32-bit OS + old OS → **R06**; disk size + raw/shared disk → **R11** |
| USB | No longer a strategy rule: **R12** is a readiness check that blocks the move until the device is replaced |
| New rule | **R05 Database server** → Replatform to Amazon RDS / Aurora / DocumentDB (AWS version of Azure R07) |
| OS field switch | The strategy is worked out for **both** RVTools OS fields; the client chooses in the workbook (sheet OS_Choice, column OS_Field_To_Use) and every result updates |
| Rule table | Strategy_Rules is now a compact, plain-language table with live VM counts |
| vCD tab | No longer read |
| Settings | 71 → 59 (`SettingsVersion = '2.0'`) |

## Files

Status is against the original AWS team code.

| File | Status | What it does |
|---|---|---|
| `Map_On-Prem_to_AWS.ps1` | modified (v2.2) | Orchestration. Parameters `-StrategySettingsPath`, `-OSOverridePath`, `-PreferRelocateForLegacyOS`, `-SiteName` (used in the output file name), **`-OSFieldDefault`**. Reads vInfo and vDisk. |
| `Modules\MigrationStrategyGenerator.ps1` | new (v1), modified (v3.0) | Decision engine (R01–R13), the rule catalogue and (v3.0) the exact-version chooser: `Get-VersionChoiceCatalog`, `Get-VersionScenarioRows`, `Import-WorkbookChoices`. `Get-ParameterMapRows` and the catalogue's settings table are kept for other tooling but are not written to the workbook. |
| `Config\AWSMigrationStrategySettings.psd1` | new (v1), modified (v3.0) | Every pattern, threshold, date, vendor, database engine and target label; (v3.0) `VersionChoices`, `LiveData`. No values are hard-coded in the engine. |
| `Modules\SummaryBuilder.ps1` | modified (v3.0) | Counts by strategy and readiness (live formulas in the workbook), EC2 cost for Rehost / Replatform VMs, VMs waiting for an exact-version choice. |
| `Modules\ExcelExporter.ps1` | modified (v3.0) | Writes the sheets (see *Workbook tabs*), the OS_Choice switch, the Exact_OS_Version dropdowns, OS_Versions / Data_Sources and the live formulas. |
| `Modules\InstanceSizing.ps1` | bug fix (v1) | `$isWindows` renamed to `$isWindowsPricing`: under PowerShell 7 the old name clashes with the read-only `$IsWindows` |
| `Modules\OSMapping.ps1`, `PricingIndex.ps1`, `VMProcessor.ps1`, `DiskProcessor.ps1` | unchanged | EC2 / EBS sizing and pricing |

Files are saved as UTF-8 with BOM, so they work on Windows PowerShell 5.1 and on PowerShell 7. Requires the ImportExcel module.

## Usage

```powershell
.\Map_On-Prem_to_AWS.ps1 -InputExcelPath .\DE-HUB_RVTools_export.xlsx
# start the workbook on the configuration-file OS instead of VMware Tools (the client can still switch)
.\Map_On-Prem_to_AWS.ps1 -InputExcelPath ... -OSFieldDefault ConfigFile
# guest-checked OS strings (CSV/XLSX, columns VMName or VM, and OS) - always win over both RVTools fields
.\Map_On-Prem_to_AWS.ps1 -InputExcelPath ... -OSOverridePath .\os_checks.csv
# send VMs whose OS is past ALL support to Relocate (VMware on AWS / EVS) instead of Replatform
.\Map_On-Prem_to_AWS.ps1 -InputExcelPath ... -PreferRelocateForLegacyOS
# v3.0: keep the client's choices (Exact_OS_Version, OS_Field_To_Use) from the previous workbook
.\Map_On-Prem_to_AWS.ps1 -InputExcelPath ... -ReuseChoicesFrom .\AWS_Mapping_Output_DE-HUB_reviewed.xlsx
# v3.0: no internet calls (lifecycle dates from the cache / the snapshot shipped with the tool)
.\Map_On-Prem_to_AWS.ps1 -InputExcelPath ... -Offline
```

The output is saved next to the input as `AWS_Mapping_Output_<site>.xlsx` unless `-OutputExcelPath` is given (`<site>` = `-SiteName`, else taken from the input file name).

## Workbook tabs

| Tab | Content |
|---|---|
| VM_Inventory | EC2 sizing and pricing per VM |
| Disk_Inventory | EBS sizing and pricing per VM |
| Summary_Metrics | Totals; strategy / readiness counts are live formulas |
| Legend | Meaning of the columns and values |
| OS_Choice | Dropdown (cell C4): OS field for all VMs |
| OS_Versions | v3.0: how to pick the exact Windows version (3 steps), live count of VMs still to check, every version with its dates (the dropdown lists) |
| Migration_Strategy | One row per VM; column OS_Field_To_Use switches one VM; column Exact_OS_Version (yellow) sets the exact OS version; every result column is a live formula |
| Strategy_VMwareTools, Strategy_ConfigFile | Hidden: the results for each OS field, read by Migration_Strategy |
| Strategy_Versions | Hidden (v3.0): the result of every exact-version option per VM and OS field, read by Migration_Strategy |
| Strategy_Rules | The rules in the order they run, in plain words, with live VM counts |
| Data_Sources | v3.0: where every live value came from (live API / cache / snapshot / settings file) and any warning |

The **Parameter_Map** and **Strategy_Settings** tabs are no longer written, and the **Parameter_Map_Ref** column has been removed from Migration_Strategy (and the hidden sheets); Rule_Applied shows which rule decided each VM. Because Strategy_Settings is gone, the settings values (patterns, thresholds, dates, targets) are no longer exposed to the client in the workbook: they stay in `Config\AWSMigrationStrategySettings.psd1`.

## Rules, in the order they run

| Stage | Rule | Name | Strategy | Register row (internal) |
|---|---|---|---|---|
| Exclude | R01 | Template | Excluded | #1 |
| Decide (first match wins) | R02 | Platform component | Retire | (Azure R01) |
| | R03 | Windows desktop | Replace (WorkSpaces) | #15 |
| | R04 | Vendor product or appliance | Replace; Retire if the vendor ended it; Low if only the client's tag | #14, #13 |
| | R05 | Database server | Manual Review (reviewed before choosing Amazon RDS / Aurora / DocumentDB or EC2) - setting `DatabaseStrategy` | new (Azure R07) |
| | R06 | Old or 32-bit OS | Replatform (Relocate with the switch) | #18, #32, #11 |
| | R07 | OS not accepted by AWS VM Import | Rehost, Blocked | #19 |
| | R08 | OS version not known | Rehost, guest check | #20 |
| | R09 | Supported OS | Rehost | #17 |
| | R10 | No rule matched | Manual Review | - |
| Readiness (never change the strategy) | R11 | Disk limits: ≥ 8 TiB (manifest route), ≥ 16 TiB, > 21 disks, raw / shared (blocker) | - | #22, #23 |
| | R12 | DMZ server: final wave | - | #25 |
| Final check | R13 | Powered off | Client Review (never retired automatically) | #2 |

Anti-affinity (#30), large storage, high memory, multi-NIC, cluster names, stale names and OS mismatches are notes and flags only.

## OS field switch

RVTools reports the OS twice: *OS according to the VMware Tools* (what the running guest reports) and *OS according to the configuration file* (set when the VM was created, often stale). The engine runs the rules for **both**:

- **OS_Choice** sheet, cell C4: the OS field for all VMs (*VMware Tools* or *Configuration file*).
- **Migration_Strategy**, column **OS_Field_To_Use**: *Default* (follow C4), *VMware Tools* or *Configuration file* for one VM. **OS_Field_Applied** shows the field used.
- Every result column of Migration_Strategy, the VM counts on Strategy_Rules and the strategy / readiness counts on Summary_Metrics update when the choice changes. The two hidden sheets Strategy_VMwareTools and Strategy_ConfigFile hold the two results.
- The EC2 cost totals on Summary_Metrics are for the field chosen when the tool ran (`-OSFieldDefault`, else `DefaultOSField`).
- An override file always wins, so both choices give the same result for those VMs.

## Database rule (R05)

Matched on the VM name, first match wins (setting `DatabaseProducts`): PostgreSQL (`postgres`, `pgsql`, `psql`), MySQL, MariaDB, MongoDB, Oracle, then Microsoft SQL Server (`sql`). The strategy is **Manual Review** (manager decision, the same as Azure R07): every database is reviewed before choosing **Replatform** to the managed service named in AWS_Target or **Rehost to EC2**. Confidence is Low because the engine is inferred from the name; the client confirms engine, version, edition and features. Set `DatabaseStrategy = 'Replatform'` or `'Rehost'` to decide databases automatically instead.

## Test results (4 RVTools exports, 976 VMs, default settings, VMware Tools)

All test runs used a fake pricing file, so the dollar figures are **not real**.

- **Version 2.2:** Rehost 635, Replace 138, Manual Review 68, Replatform 50, Retire 47, Client Review 38. Every cell of Migration_Strategy and both hidden sheets was compared with v2.1: the only differences are the 70 database VMs, IIS-PL-GOS-1 (USB check removed) and the renumbered DMZ / Powered-off rules. The four workbooks recalculate in LibreOffice with 0 errors (18,641 formulas), stored values match, and the OS switch reproduces the configuration-file results row for row.

Earlier versions:

- Rehost 635, Replace 138, Replatform 118, Retire 47, Client Review 38.
- Compared with version 1, 94 VMs change strategy: 28 because of the removed and merged rules (exactly as tested in the Rule Condensation Review) and 66 because of the new database rule (Rehost → Replatform). The database rule decides 68 VMs; 2 of them were already Replatform for their old OS.
- With the configuration file instead of VMware Tools, 323 VMs get a different strategy (mostly Rehost → Replatform, because the configuration file often shows an older OS).
- The four workbooks recalculate in LibreOffice with 0 formula errors (19,618 formulas), and switching the OS field reproduces the engine's configuration-file results row for row.
- After removing the Parameter_Map / Strategy_Settings tabs and the Parameter_Map_Ref column: strategies, readiness and all other Migration_Strategy values, the Strategy_Rules VM counts and the Summary_Metrics counts are unchanged in all four workbooks; they recalculate in LibreOffice with 0 errors (18,642 formulas: one per VM fewer, because the Parameter_Map_Ref column is gone), and the OS switch still reproduces the configuration-file results row for row.

**Upgrading a customised settings file:** version 2.0 removes `OsSourcePrecedence`, `MismatchAction`, `EnableDecomNoteRule`, `DecomNotePatterns`, `RetainIfUsbDevice`, `EnablePlantFloorRule`, `PlantFloorNamePatterns`, `PlantFloorTextPatterns`, `Arch32BitStrategy`, `EnableDomainControllerRule`, `DomainControllerNamePatterns`, `DomainControllerTextPatterns`, `DomainControllerTarget`, `ServiceCatalogue`, `ServiceStrategy`, `CdConnectedPattern`, and adds `DefaultOSField`, `UsbReadinessCheck`, `DatabaseProducts`, `DatabaseStrategy`. Start from the shipped file.

## Pre-existing issues found in the original AWS code (not changed)

- **Graviton sizing for x86 VMs:** the sizing can pick Graviton (ARM) instances for x86 VMs. These would not boot an imported x86 image.
- **Buffer default:** the default is 15% in the code, but 10% is documented.
- **EBS totals:** they include powered-off VMs, while the EC2 totals exclude them.
- **No-break spaces:** the Word copies of the code contain no-break spaces. Always share the `.ps1` files themselves.
