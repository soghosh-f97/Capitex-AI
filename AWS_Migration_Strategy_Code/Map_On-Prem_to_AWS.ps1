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
    [string]$StrategySettingsPath = (Join-Path $PSScriptRoot "Config\AWSMigrationStrategySettings.psd1"),

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
    [string]$OSFieldDefault = ''
)

# Pre-checks
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel -ErrorAction Stop
Add-Type -AssemblyName Microsoft.VisualBasic

if (-not (Test-Path $InputExcelPath)) { Write-Error "Input file not found: $InputExcelPath"; exit 1 }

$Buffer = 1 + ($BufferPercent / 100)

# Load all function modules
. "$PSScriptRoot\Modules\OSMapping.ps1"
. "$PSScriptRoot\Modules\InstanceSizing.ps1"
. "$PSScriptRoot\Modules\PricingIndex.ps1"
. "$PSScriptRoot\Modules\VMProcessor.ps1"
. "$PSScriptRoot\Modules\DiskProcessor.ps1"
. "$PSScriptRoot\Modules\SummaryBuilder.ps1"
. "$PSScriptRoot\Modules\ExcelExporter.ps1"
. "$PSScriptRoot\Modules\MigrationStrategyGenerator.ps1"

# Strategy settings are validated before the (slow) pricing download, so a bad settings file fails fast
$StrategySettings = Get-MigrationStrategyDefaults -SettingsPath $StrategySettingsPath

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
$csvPathUS = Get-AWSBulkPricingFile -RegionCode "us-east-1"
$csvPathDE = Get-AWSBulkPricingFile -RegionCode "eu-central-1"
if (-not $csvPathUS -or -not $csvPathDE) { Write-Error "Could not obtain pricing files."; exit 1 }

$USIndex = Build-AWSPricingIndex -CsvPath $csvPathUS
$DEIndex = Build-AWSPricingIndex -CsvPath $csvPathDE
$InstanceSpecs = $USIndex.InstanceSpecMap.Values
if (-not $InstanceSpecs -or $InstanceSpecs.Count -eq 0) { Write-Error "Instance spec table is empty."; exit 1 }

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
$StrategyAsOf      = Get-Date   # one date for every lifecycle check in this run (support dates are compared with it)
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

Export-MigrationReport -EC2Results $EC2Results -EBSResults $EBSResults -Summary $Summary -OutputPath $OutputExcelPath `
    -StrategyRows $StrategyRows -StrategyRowsByField $StrategyRowsByField -DefaultOSField $OSFieldDefault `
    -RuleCatalog $RuleCatalog

Write-Host "Migration strategy distribution (OS field: $OSFieldDefault):" -ForegroundColor Cyan
$StrategyRows | Group-Object Migration_Strategy | Sort-Object Count -Descending |
    ForEach-Object { Write-Host ("  {0,-14}: {1}" -f $_.Name, $_.Count) }

Write-Host "Output written to: $OutputExcelPath" -ForegroundColor Green
