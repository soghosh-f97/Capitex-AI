<#
.SYNOPSIS
    On-Prem to AWS Mapping Generator (from RVTools export)
    Orchestration script only - all logic lives in Modules\
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$InputExcelPath,

    # Default: AWS_Mapping_Output_<site>.xlsx in the same folder as the input file, e.g. AWS_Mapping_Output_DE-HUB.xlsx,
    # so the runs for the four vCenters never overwrite each other. Give a path to choose your own name.
    [Parameter(Mandatory=$false)]
    [string]$OutputExcelPath,

    [Parameter(Mandatory=$false)]
    [double]$BufferPercent = 15,

    # --- Migration strategy (Parameter_Map automation) ---
    # Settings file with every keyword list / threshold / lookup used by the strategy rules
    [Parameter(Mandatory=$false)]
    [string]$StrategySettingsPath = (Join-Path (Join-Path $PSScriptRoot 'Config') 'AWSMigrationStrategySettings.psd1'),

    # Optional CSV/XLSX with columns VMName, OS - guest-check results that override the RVTools OS strings
    [Parameter(Mandatory=$false)]
    [string]$OSOverridePath,

    # Parameter_Map #32: send VMs whose OS is past ALL support to Relocate (Amazon EVS / VMware Cloud on AWS) instead of Replatform
    [Parameter(Mandatory=$false)]
    [switch]$PreferRelocateForLegacyOS,

    [Parameter(Mandatory = $false, HelpMessage = "Site name used in the output file name, AWS_Mapping_Output_<site>.xlsx (default: taken from the input file name, e.g. DE-HUB)")]
    [string]$SiteName,

    # Which RVTools OS field decides the strategy by default: VMwareTools (recommended) or ConfigFile.
    # Empty = setting DefaultOSField. Both versions are always worked out, so the client can switch in the workbook.
    [Parameter(Mandatory = $false)]
    [ValidateSet('', 'VMwareTools', 'ConfigFile')]
    [string]$OSFieldDefault = '',

    # --- Live data (v3.0) ---
    # No internet at all: OS support dates from the local cache / Config\LifecycleSnapshot.json, pricing from the cached file
    [Parameter(Mandatory = $false)]
    [switch]$Offline,

    # Also ask AWS (needs the AWS CLI + read-only credentials): exact max NICs and CPU architecture per instance type,
    # services offered in the region, standard Windows AMIs. Without it, built-in tables are used.
    [Parameter(Mandatory = $false)]
    [switch]$UseAwsApis,

    # Region for the -UseAwsApis checks
    [Parameter(Mandatory = $false)]
    [string]$AwsRegion = 'eu-central-1',

    # A previous output workbook: the client's choices in it (OS_Field_To_Use, Exact_OS_Version) are carried over
    [Parameter(Mandatory = $false)]
    [string]$ReuseChoicesFrom
)

# Pre-checks
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel -ErrorAction Stop
Add-Type -AssemblyName Microsoft.VisualBasic

if (-not (Test-Path $InputExcelPath)) { Write-Error "Input file not found: $InputExcelPath"; exit 1 }

$Buffer = 1 + ($BufferPercent / 100)

# Load all function modules (Join-Path so the script also runs under PowerShell 7 on Linux / macOS)
$moduleDir = Join-Path $PSScriptRoot 'Modules'
foreach ($moduleFile in @('LiveDataProviders.ps1', 'OSMapping.ps1', 'InstanceSizing.ps1', 'PricingIndex.ps1', 'VMProcessor.ps1',
                          'DiskProcessor.ps1', 'SummaryBuilder.ps1', 'ExcelExporter.ps1', 'MigrationStrategyGenerator.ps1')) {
    . (Join-Path $moduleDir $moduleFile)
}

# Strategy settings are validated before the (slow) pricing download, so a bad settings file fails fast
$StrategySettings = Get-MigrationStrategyDefaults -SettingsPath $StrategySettingsPath
$StrategyAsOf     = Get-Date   # one date for every lifecycle check in this run (support dates are compared with it)

#---------------------------------------------------------------------
# LIVE DATA (v3.0): OS support dates, support cut-off, AWS VM Import page check
# Each source falls back (cache -> offline snapshot -> settings file); the sheet Data_Sources shows what was used.
#---------------------------------------------------------------------
Write-Host "=== Live data ===" -ForegroundColor Cyan
$LiveCacheDir     = Get-LiveCacheRoot
$LifecycleCatalog = Get-LifecycleCatalog -Settings $StrategySettings -Offline:$Offline -CacheDir $LiveCacheDir
$LifecycleChanges = @(Update-OsLifecycleFromCatalog -Settings $StrategySettings -Catalog $LifecycleCatalog)
foreach ($p in ($LifecycleCatalog.Products.Keys | Sort-Object)) {
    Write-Host ("  OS support dates {0,-15}: {1} ({2})" -f $p, $LifecycleCatalog.Source[$p], $LifecycleCatalog.RetrievedAt[$p])
}
$dateChanges = @($LifecycleChanges | Where-Object { $_.Field -ne 'Live' })
if ($dateChanges.Count -gt 0) { Write-Host ("  {0} OS support date(s) differ from the settings file and were updated (see sheet Data_Sources)" -f $dateChanges.Count) -ForegroundColor Yellow }
foreach ($m in @($LifecycleCatalog.Messages)) { Write-Host "  note: $m" -ForegroundColor DarkGray }

# SupportCutoffDate may be relative ('today+12m'): resolved once, so every rule and sheet uses the same date
$CutoffSetting = [string]$StrategySettings.SupportCutoffDate
try { $StrategySettings['SupportCutoffDate'] = Resolve-SupportCutoffDate -Value $CutoffSetting -AsOf $StrategyAsOf }
catch { Write-Error $_.Exception.Message; exit 1 }
Write-Host ("  Support cut-off date (R06): {0}{1}" -f $StrategySettings.SupportCutoffDate, $(if ($CutoffSetting -ne $StrategySettings.SupportCutoffDate) { " (from '$CutoffSetting')" } else { '' }))

$VmImportCheck = Test-VmImportDocChanged -Settings $StrategySettings -Offline:$Offline
if ($VmImportCheck.Status -eq 'Changed') { Write-Warning ("AWS VM Import page has changed since the settings were reviewed - " + $VmImportCheck.Detail) }
else { Write-Host ("  AWS VM Import page: {0}" -f $VmImportCheck.Status) }

# Client choices from a previous workbook - read now, before an output file with the same name is replaced
$PriorChoices = @{}
if ($ReuseChoicesFrom) {
    if (-not (Test-Path -LiteralPath $ReuseChoicesFrom)) { Write-Error "Previous workbook not found: $ReuseChoicesFrom"; exit 1 }
    if (Get-Command Import-WorkbookChoices -ErrorAction SilentlyContinue) {
        $PriorChoices = Import-WorkbookChoices -Path $ReuseChoicesFrom
        if ($null -eq $PriorChoices) { $PriorChoices = @{} }
        Write-Host ("  Client choices carried over from {0}: {1} VM(s)" -f $ReuseChoicesFrom, $PriorChoices.Count) -ForegroundColor Green
    }
    else { Write-Warning "-ReuseChoicesFrom: this version of the engine cannot read previous choices - ignored." }
}

# Site / vCenter name: -SiteName, else taken from the input file name ('DE-HUB_RVTools_export_all_....xlsx' -> 'DE-HUB')
if ([string]::IsNullOrWhiteSpace($SiteName)) {
    $inputLeaf = [System.IO.Path]::GetFileNameWithoutExtension($InputExcelPath)
    $SiteName  = if ($inputLeaf -match $StrategySettings.SiteNameFromFilePattern) { $Matches[1] } else { $inputLeaf }
}

# Output file: one per site, so runs for different vCenters never overwrite each other
if ([string]::IsNullOrWhiteSpace($OutputExcelPath)) {
    $safeSite = ($SiteName -replace '[\\/:*?"<>|]', '_').Trim()
    $inputDir = Split-Path -Path (Resolve-Path -LiteralPath $InputExcelPath).Path -Parent
    $OutputExcelPath = Join-Path $inputDir ("AWS_Mapping_Output_{0}.xlsx" -f $safeSite)
}
if (Test-Path -LiteralPath $OutputExcelPath) {
    try { Remove-Item -LiteralPath $OutputExcelPath -Force -ErrorAction Stop }
    catch { Write-Error "Cannot replace '$OutputExcelPath' - close it in Excel and run again."; exit 1 }
}
Write-Host "Site: $SiteName  ->  output: $OutputExcelPath" -ForegroundColor Cyan

# Pipeline
$csvPathUS = Get-AWSBulkPricingFile -RegionCode "us-east-1" -Offline:$Offline
$csvPathDE = Get-AWSBulkPricingFile -RegionCode "eu-central-1" -Offline:$Offline
if (-not $csvPathUS -or -not $csvPathDE) { Write-Error "Could not obtain pricing files."; exit 1 }

$USIndex = Build-AWSPricingIndex -CsvPath $csvPathUS
$DEIndex = Build-AWSPricingIndex -CsvPath $csvPathDE
$InstanceSpecs = $USIndex.InstanceSpecMap.Values
if (-not $InstanceSpecs -or $InstanceSpecs.Count -eq 0) { Write-Error "Instance spec table is empty."; exit 1 }

# Optional AWS API facts (exact NIC limits / CPU architectures / region services / Windows AMIs). Read by the sizing
# and OS-mapping modules through the variable name $AwsLiveFacts - set it before sizing.
$AwsLiveFacts = Get-AwsLiveFacts -Region $AwsRegion -Enabled:$UseAwsApis
foreach ($m in @($AwsLiveFacts.Messages)) { Write-Host "  note: $m" -ForegroundColor DarkGray }

$vInfo = @(Import-Excel -Path $InputExcelPath -WorksheetName "vInfo")
if ($vInfo.Count -eq 0) { Write-Error "The vInfo tab in $InputExcelPath has no rows - nothing to size or classify."; exit 1 }
# NOTE: vDisk tab is no longer needed - EBS sizing now derives from vInfo's
# per-VM aggregate columns (Total disk capacity MiB / Provisioned MiB / In Use MiB).

# --- Identify Template VM names to exclude from disk calculations ---
$templateVMNames = $vInfo | Where-Object { $_.Template -eq $true } | Select-Object -ExpandProperty VM

$EC2Results = Build-EC2Mapping -VInfoData $vInfo -USIndex $USIndex -DEIndex $DEIndex
$EBSResults = Build-EBSMapping -VInfoData $vInfo -TemplateVMNames $templateVMNames -USIndex $USIndex -DEIndex $DEIndex

#---------------------------------------------------------------------
# MIGRATION STRATEGY - automates the Parameter_Map rules (see Modules\MigrationStrategyGenerator.ps1)
#---------------------------------------------------------------------
Write-Host "=== Building Migration Strategy (rules R01-R13) ===" -ForegroundColor Cyan

# Extra RVTools tabs used by the strategy rules (optional: a missing tab just means no findings)
$sheetNames = @((Get-ExcelSheetInfo -Path $InputExcelPath).Name)
function Import-RVToolsTab {
    param([string]$Name)
    if ($sheetNames -contains $Name) { return @(Import-Excel -Path $InputExcelPath -WorksheetName $Name) }
    Write-Warning "RVTools tab '$Name' not found - related strategy checks are skipped."
    return @()
}
$vDisk = Import-RVToolsTab -Name "vDisk"   # R11 Disk limits (Parameter_Map #22, #23)

# Optional guest-check results that override the RVTools OS strings (columns: VMName or VM, and OS)
$osOverrides = @{}
if ($OSOverridePath) {
    if (-not (Test-Path $OSOverridePath)) { Write-Error "OS override file not found: $OSOverridePath"; exit 1 }
    $ovRows = if ($OSOverridePath -match '\.xlsx$') { @(Import-Excel -Path $OSOverridePath) } else { @(Import-Csv -Path $OSOverridePath) }
    foreach ($r in $ovRows) {
        $ovName = if ($r.PSObject.Properties['VMName']) { $r.VMName } elseif ($r.PSObject.Properties['VM']) { $r.VM } else { $null }
        if ($ovName -and $r.OS) { $osOverrides[([string]$ovName).Trim().ToLower()] = ([string]$r.OS).Trim() }
    }
    if ($ovRows.Count -gt 0 -and $osOverrides.Count -eq 0) { Write-Warning "OS override file has no usable rows - expected columns 'VMName' (or 'VM') and 'OS'." }
    $vmNamesLower = @{}; foreach ($v in $vInfo) { if ($v.VM) { $vmNamesLower[([string]$v.VM).Trim().ToLower()] = $true } }
    $unmatched = @($osOverrides.Keys | Where-Object { -not $vmNamesLower.ContainsKey($_) })
    if ($unmatched.Count -gt 0) { Write-Warning ("{0} OS override name(s) match no VM in this export: {1}" -f $unmatched.Count, (($unmatched | Select-Object -First 10) -join ', ')) }
    Write-Host "Loaded $($osOverrides.Count) OS override(s) from $OSOverridePath" -ForegroundColor Green
}

$StrategyInventory = Build-StrategyInventory -VInfoData $vInfo -VDiskData $vDisk -OSOverrides $osOverrides -Settings $StrategySettings
Write-Host ("Strategy lifecycle checks as of {0:yyyy-MM-dd}" -f $StrategyAsOf)

# The strategy is worked out for BOTH OS fields, so the client can switch in the workbook (sheet OS_Choice)
if ([string]::IsNullOrWhiteSpace($OSFieldDefault)) { $OSFieldDefault = $StrategySettings.DefaultOSField }
$StrategyRowsByField = [ordered]@{}
foreach ($field in @('VMwareTools', 'ConfigFile')) {
    $StrategyRowsByField[$field] = @(Get-MigrationStrategyRows -StrategyInventory $StrategyInventory -Settings $StrategySettings -OSField $field -PreferRelocateForLegacyOS:$PreferRelocateForLegacyOS -AsOf $StrategyAsOf)
}
$StrategyRows      = $StrategyRowsByField[$OSFieldDefault]
Write-Host "Default OS field: $OSFieldDefault (the client can switch it in the workbook)"
# Template count: shown as the static VM count of rule R01 on the Strategy_Rules sheet
$templateCount     = @($vInfo | Where-Object { Test-RVToolsTrue (Get-RowValue $_ 'Template') $StrategySettings.TemplateValuePattern }).Count
# The Parameter_Map and Strategy_Settings sheets are no longer written to the workbook
# (Get-ParameterMapRows and the rule catalogue's .Settings stay available in the module for other tooling)
$RuleCatalog       = Get-MigrationStrategyRuleCatalog -Settings $StrategySettings -PreferRelocateForLegacyOS:$PreferRelocateForLegacyOS -TemplateCount $templateCount

$Summary    = Build-SummaryReport -EC2Results $EC2Results -EBSResults $EBSResults -BufferPercent $BufferPercent -InstanceSpecCount $InstanceSpecs.Count -StrategyRows $StrategyRows -StrategyRowsAll @($StrategyRowsByField.Values | ForEach-Object { $_ })

#---------------------------------------------------------------------
# EXACT OS VERSION CHOOSER (v3.0): the client picks e.g. 'Windows 11 24H2' for a VM that RVTools only calls
# 'Windows 11'. The engine works out every choice in advance, so the workbook updates the moment one is picked.
#---------------------------------------------------------------------
$VersionCatalog   = @()
$VersionScenarios = @()
if (Get-Command Get-VersionChoiceCatalog -ErrorAction SilentlyContinue) {
    $VersionCatalog = @(Get-VersionChoiceCatalog -Settings $StrategySettings -LifecycleCatalog $LifecycleCatalog)
    if ($VersionCatalog.Count -gt 0 -and (Get-Command Get-VersionScenarioRows -ErrorAction SilentlyContinue)) {
        $VersionScenarios = @(Get-VersionScenarioRows -StrategyInventory $StrategyInventory -Settings $StrategySettings -VersionCatalog $VersionCatalog -PreferRelocateForLegacyOS:$PreferRelocateForLegacyOS -AsOf $StrategyAsOf)
    }
    $optionCount = 0; foreach ($fam in $VersionCatalog) { $optionCount += @($fam.Options).Count }
    Write-Host ("Exact OS version chooser: {0} famil(ies), {1} option(s), {2} pre-calculated result row(s)" -f $VersionCatalog.Count, $optionCount, $VersionScenarios.Count)
}

#---------------------------------------------------------------------
# DATA SOURCES sheet: where every value of this run came from
#---------------------------------------------------------------------
$pricingInfo = @()
foreach ($pf in @(@{ Region = 'us-east-1'; Path = $csvPathUS }, @{ Region = 'eu-central-1'; Path = $csvPathDE })) {
    $i = Get-PricingFileInfo -CsvPath $pf.Path
    $i | Add-Member -NotePropertyName Region -NotePropertyValue $pf.Region -Force
    $pricingInfo += $i
}
$DataSources = @(New-DataSourceRows -LifecycleCatalog $LifecycleCatalog -LifecycleChanges $LifecycleChanges -VmImportCheck $VmImportCheck `
    -AwsFacts $AwsLiveFacts -PricingFiles $pricingInfo -CutoffSetting $CutoffSetting -CutoffResolved ([string]$StrategySettings.SupportCutoffDate) `
    -ArmTypesExcluded ([int]$USIndex.ArmInstanceTypeCount) -Settings $StrategySettings)

# The new workbook parameters are passed only when the exporter supports them (keeps this script usable with older modules)
$exportArgs = @{
    EC2Results          = $EC2Results
    EBSResults          = $EBSResults
    Summary             = $Summary
    OutputPath          = $OutputExcelPath
    StrategyRows        = $StrategyRows
    StrategyRowsByField = $StrategyRowsByField
    DefaultOSField      = $OSFieldDefault
    RuleCatalog         = $RuleCatalog
}
$exportParams = (Get-Command Export-MigrationReport).Parameters
$optionalExport = @{ VersionCatalog = $VersionCatalog; VersionScenarios = $VersionScenarios; PriorChoices = $PriorChoices; DataSources = $DataSources }
foreach ($k in $optionalExport.Keys) {
    if ($exportParams.ContainsKey($k)) { $exportArgs[$k] = $optionalExport[$k] }
    elseif ($k -eq 'DataSources') { Write-Warning "This exporter has no Data_Sources sheet - the data sources are listed below instead." }
}
Export-MigrationReport @exportArgs
if (-not $exportParams.ContainsKey('DataSources')) {
    $DataSources | Format-Table Area, Status, RetrievedAt -AutoSize | Out-String -Width 200 | Write-Host
}

Write-Host "Migration strategy distribution (OS field: $OSFieldDefault):" -ForegroundColor Cyan
$StrategyRows | Group-Object Migration_Strategy | Sort-Object Count -Descending |
    ForEach-Object { Write-Host ("  {0,-14}: {1}" -f $_.Name, $_.Count) }

Write-Host "Output written to: $OutputExcelPath" -ForegroundColor Green
