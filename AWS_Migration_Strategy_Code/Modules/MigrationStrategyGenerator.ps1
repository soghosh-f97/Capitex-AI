<#
.SYNOPSIS
    MigrationStrategyGenerator.ps1 (AWS) - builds the "Migration_Strategy" sheet: one primary migration
    strategy per VM (Retire / Retain / Replace / Replatform / Relocate / Rehost) with confidence, the rule
    that fired, the Parameter_Map row behind it, rationale, AWS target, readiness, pre-migration actions,
    client confirmations, design considerations and review flags. Also generates the "Strategy_Rules" sheet
    from the same settings, so the documentation can never drift from the logic.
    The Parameter_Map rows (Get-ParameterMapRows), the settings table (.Settings of
    Get-MigrationStrategyRuleCatalog) and the Parameter_Map_Ref column are still produced for other tooling,
    but the exporter no longer writes them to the workbook (no Parameter_Map / Strategy_Settings sheets).
.DESCRIPTION
    Follows the pattern of the Azure MigrationStrategyGenerator.ps1, adapted to the AWS Parameter_Map.
    Dot-source from Map-OnPrem-to-AWS.ps1 like the other modules. No other module is required.

    ALL keyword lists, thresholds, lookup tables and targets are read from
    Config\AWSMigrationStrategySettings.psd1 (Get-MigrationStrategyDefaults). The code holds logic only.

    INPUTS (built by Build-StrategyInventory from the RVTools export):
      vInfo  - VM, Template, Powerstate, Cluster, Host, CPUs, Memory, NICs, Disks, Annotation,
               OS according to the configuration file, OS according to the VMware Tools,
               custom attributes (machinetype, managedby, application - if present),
               Cluster rule(s), Cluster rule name(s), storage column (setting StorageColumn)
      vDisk  - VM, Capacity MiB, Raw, Sharing mode, Shared Bus         (optional tab)
    Missing tabs or columns are tolerated: the related checks simply find nothing.

    EXACT OS VERSION (v3.0): for VMs whose RVTools label covers several versions (setting VersionChoices: Windows 10 / 11,
    Windows Server 2012 / 2012 R2, Windows Server 2016 or later) Get-VersionScenarioRows pre-computes every option, so the
    client picks the version in the workbook (column Exact_OS_Version) and every result column updates live.
    Support dates can come from live lifecycle data (Modules\LiveDataProviders.ps1 updates OsLifecycle before the rules run).

    RULES - numbered in the order they run (settings version 3.0):
      R01 Template              excluded (not a running server)                         [PM #1]
      --- Decide the strategy: first matching rule wins ---
      R02 Retire       Platform component (vCenter, 3PAR SP, SSMC / OneView, SimpliVity, vSAN witness)  [Azure R01]
      R03 Replace      Windows desktop OS -> Amazon WorkSpaces                          [PM #15]
      R04 Replace      Vendor product or appliance -> vendor AWS edition / SaaS         [PM #14, #13]
                       (Retire when the vendor has ended the product; Low when only the client's tag says 'appliance')
      R05 Manual Review  Database server -> reviewed before choosing RDS / Aurora / DocumentDB or EC2  [new, Azure R07]
      R06 Replatform   Old or 32-bit OS -> upgrade / rebuild on a supported 64-bit OS   [PM #18, #32, #11]
                       (Relocate when -PreferRelocateForLegacyOS and the OS is past ALL support)
      R07 Rehost       OS not accepted by AWS VM Import -> BLOCKED until upgraded       [PM #19]
      R08 Rehost       OS version not known -> guest check first                        [PM #20]
      R09 Rehost       OS on the AWS VM Import list, nothing above matched              [PM #17]
      R10 Manual Review  No rule matched (safety net)
      --- Readiness checks: every check that applies; never change the strategy ---
      R11 Disk limits  >= 16 TiB, >= 8 TiB (manifest route), > 21 disks, raw / shared disks    [PM #22, #23]
      R12 DMZ server   moves in the DMZ wave, after the AWS perimeter is live           [PM #25]
      --- Final check ---
      R13 Client Review  Powered off / suspended - never retired automatically; the strategy the rules give if the
                         VM is still needed is kept in Rationale / Alternative_Strategy   [PM #2]

    OS FIELD: the OS that decides comes from the override file, else the chosen field (DefaultOSField or
    -OSField: 'VMwareTools' or 'ConfigFile'), else the other field. Get-MigrationStrategyRows can run once per
    field, so the workbook can let the client switch between them.
    NOTES (never change the strategy): anti-affinity (PM #30), stale names, OS mismatch, large storage, disk count,
    multi-NIC, high RAM, cluster names.

    Refactor is never auto-assigned: it needs application-owner input, not infrastructure data.
#>

#region ------------------------- Settings -------------------------

function Get-MigrationStrategyDefaults {
    <#
        SINGLE place the strategy settings are loaded from. Get-MigrationStrategyRows, Get-ParameterMapRows
        and Get-MigrationStrategyRuleCatalog all receive the hashtable returned here, so the published
        "Strategy_Rules" sheet always shows the values actually used.
    #>
    param(
        [string]$SettingsPath
    )
    if ([string]::IsNullOrWhiteSpace($SettingsPath)) {
        $SettingsPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Config\AWSMigrationStrategySettings.psd1'
    }
    if (-not (Test-Path $SettingsPath)) {
        throw "Migration strategy settings file not found: $SettingsPath"
    }
    $settings = Import-PowerShellDataFile -Path $SettingsPath

    $required = @(
        'DefaultOSField', 'InfraNamePatterns', 'InfraTextPatterns', 'PoweredOffStates', 'PoweredOffAction',
        'VendorProducts', 'TagColumnsSearched', 'AnnotationPurposeRegex', 'ApplianceTags', 'ApplianceTagTarget',
        'Arch32BitPatterns', 'Allow32BitImport', 'DesktopOsPatterns', 'DesktopTarget',
        'DatabaseProducts', 'DatabaseStrategy', 'SupportCutoffDate', 'LegacyOsStrategy',
        'PreferRelocateForLegacyOS', 'RelocateTarget', 'OsLifecycle', 'VmImportSupportedOs', 'VmImportUnsupportedOs',
        'UnversionedOsPatterns', 'UnknownOsPatterns', 'MaxEbsVolumeGiB', 'MaxVolumesPerImport', 'RawDiskValuePattern',
        'SharedModePattern', 'SharedBusPattern', 'DmzPatterns', 'AntiAffinityPattern',
        'StaleNamePatterns', 'ClusterNamePatterns', 'LargeDiskGB', 'HighRamGB', 'MultiNicThreshold', 'HighDiskCount',
        'StorageColumn', 'DefaultRehostConfidence', 'AllowHighRehostConfidence', 'OsFamilies', 'VagueOsFamilies', 'EnableInfraRule',
        'Arch32BitNote', 'DesktopExcludePattern', 'ManifestImportGiB', 'Targets', 'TemplateValuePattern',
        'PoweredOffStrategy', 'SiteNameFromFilePattern',
        'GroupedOsLabels', 'GroupedOsLabelAction', 'VmImportVerifyOs'
        # ParameterMapDecisionSupport / ...Default / ParameterMapSiteCountDecidedRows are optional (internal register only)
    )
    $missing = @($required | Where-Object { -not $settings.ContainsKey($_) })
    if ($missing.Count -gt 0) {
        throw ("Migration strategy settings file is missing: " + ($missing -join ', '))
    }
    if ($settings.DefaultOSField -notin @('VMwareTools', 'ConfigFile')) { throw "Setting DefaultOSField must be 'VMwareTools' or 'ConfigFile'" }
    if ($settings.PoweredOffAction -notin @('ClientReview', 'Flag')) { throw "Setting PoweredOffAction must be 'ClientReview' or 'Flag'" }
    if ([string]::IsNullOrWhiteSpace([string]$settings.PoweredOffStrategy)) { throw "Setting PoweredOffStrategy must not be empty" }
    if ($settings.LegacyOsStrategy -notin @('Replatform', 'Rehost')) { throw "Setting LegacyOsStrategy must be 'Replatform' or 'Rehost'" }
    if ($settings.DatabaseStrategy -notin @('Replatform', 'Rehost', 'Manual Review')) { throw "Setting DatabaseStrategy must be 'Replatform', 'Rehost' or 'Manual Review'" }
    $missingTargets = @(@('Decommission', 'OnPremises', 'Ec2', 'Ec2SupportedOs', 'Ec2Os64Bit') | Where-Object { -not $settings.Targets.ContainsKey($_) })
    if ($missingTargets.Count -gt 0) { throw ("Setting Targets is missing: " + ($missingTargets -join ', ')) }

    # ---------------- Optional v3.0 keys (missing = feature off) ----------------
    # LiveData: owned by Modules\LiveDataProviders.ps1; only its shape is checked here
    if (-not $settings.ContainsKey('LiveData')) { $settings['LiveData'] = @{ Enabled = $false } }
    elseif ($settings.LiveData -isnot [hashtable]) { throw "Setting LiveData must be a hashtable (@{ ... })" }

    # OsLifecycle rows: optional Live = '<product>:<cycle>' (endoflife.date) and DateSource (set at run time)
    foreach ($r in $settings.OsLifecycle) {
        if ($r.ContainsKey('Live') -and -not [string]::IsNullOrWhiteSpace([string]$r.Live) -and ([string]$r.Live -notmatch '^[a-z0-9][a-z0-9-]*:[A-Za-z0-9._-]+$')) {
            throw "OsLifecycle row '$($r.Label)': Live must look like '<product>:<cycle>' (e.g. 'rhel:7'), not '$($r.Live)'"
        }
    }

    # VersionChoices: families the client can refine to an exact version in the workbook
    if (-not $settings.ContainsKey('VersionChoices') -or $null -eq $settings.VersionChoices) { $settings['VersionChoices'] = @() }
    $settings['VersionChoices'] = @($settings.VersionChoices)
    foreach ($vc in $settings.VersionChoices) {
        if ($vc -isnot [hashtable]) { throw 'Setting VersionChoices must be a list of @{ Match; Family; Product; CyclePattern; OsTemplate; Help }' }
        $vcMissing = @(@('Match', 'Family', 'Product', 'CyclePattern', 'OsTemplate') | Where-Object { [string]::IsNullOrWhiteSpace([string]$vc[$_]) })
        if ($vcMissing.Count -gt 0) { throw ("VersionChoices entry '$($vc.Family)' is missing: " + ($vcMissing -join ', ')) }
        foreach ($p in @('Match', 'CyclePattern')) {
            try { [void][regex]::new([string]$vc[$p]) } catch { throw "VersionChoices entry '$($vc.Family)': $p is not a valid regular expression: $($vc[$p])" }
        }
        if (-not $vc.ContainsKey('Help')) { $vc['Help'] = 'Check the exact version inside the VM (winver)' }
    }

    # SupportCutoffDate: 'yyyy-MM-dd', 'today+<n>m' or 'today+<n>y' (relative to the run date; resolved once, here).
    # The text as typed is kept in SupportCutoffSetting (shown on Data_Sources); SupportCutoffDate becomes the date used.
    $cutoffRaw = ([string]$settings.SupportCutoffDate).Trim()
    $settings['SupportCutoffSetting'] = $cutoffRaw
    $resolved = $null
    if (Get-Command Resolve-SupportCutoffDate -ErrorAction SilentlyContinue) {
        try { $resolved = Resolve-SupportCutoffDate -Value $cutoffRaw -AsOf (Get-Date) }
        catch { $resolved = $null }
    }
    else {
        $inv = [System.Globalization.CultureInfo]::InvariantCulture
        if ($null -ne (ConvertTo-StrategyDate $cutoffRaw)) { $resolved = (ConvertTo-StrategyDate $cutoffRaw).ToString('yyyy-MM-dd', $inv) }
        elseif ($cutoffRaw -match '^(?i)today\s*\+\s*(\d+)\s*m$') { $resolved = (Get-Date).Date.AddMonths([int]$Matches[1]).ToString('yyyy-MM-dd', $inv) }
        elseif ($cutoffRaw -match '^(?i)today\s*\+\s*(\d+)\s*y$') { $resolved = (Get-Date).Date.AddYears([int]$Matches[1]).ToString('yyyy-MM-dd', $inv) }
        elseif ($cutoffRaw -match '^(?i)today$') { $resolved = (Get-Date).Date.ToString('yyyy-MM-dd', $inv) }
    }
    if ([string]::IsNullOrWhiteSpace([string]$resolved)) {
        Write-Warning "Setting SupportCutoffDate '$cutoffRaw' is not valid (use yyyy-MM-dd, today+<n>m or today+<n>y) - the 'ends before cut-off' check of R06 is skipped."
    }
    else { $settings['SupportCutoffDate'] = [string]$resolved }

    $settings['SettingsPath'] = $SettingsPath
    return $settings
}

function Get-SettingsForOSField {
    <#
        Returns the settings with the OS precedence for one OS field: the override file first, then the chosen
        field, then the other field (used only when the chosen one is blank). Empty -OSField = DefaultOSField.
    #>
    param(
        [Parameter(Mandatory)] [hashtable]$Settings,
        [ValidateSet('', 'VMwareTools', 'ConfigFile')] [string]$OSField = ''
    )
    if ([string]::IsNullOrWhiteSpace($OSField)) { $OSField = $Settings.DefaultOSField }
    $copy  = $Settings.Clone()
    $other = if ($OSField -eq 'VMwareTools') { 'ConfigFile' } else { 'VMwareTools' }
    $copy['OsSourcePrecedence'] = @('Override', $OSField, $other)
    $copy['OSField'] = $OSField
    return $copy
}

#endregion

#region ------------------------- Small helpers -------------------------

function Test-AnyPattern {
    param([string]$Text, [string[]]$Patterns)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    foreach ($p in $Patterns) {
        if ($Text -match $p) { return $true }
    }
    return $false
}

function Get-FirstPatternHit {
    param([string]$Text, [string[]]$Patterns)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    foreach ($p in $Patterns) {
        if ($Text -match $p) { return $p }
    }
    return $null
}

function Get-FirstNonEmpty {
    # Windows PowerShell 5.1 has no ?? operator.
    param($First, $Second)
    if ($null -ne $First -and [string]$First -ne '') { return $First }
    return $Second
}

function Get-RowValue {
    # Reads a property that may not exist in every RVTools export (custom attributes, optional columns).
    param($Row, [string]$Name)
    if ($null -eq $Row) { return '' }
    $prop = $Row.PSObject.Properties[$Name]
    if ($null -eq $prop -or $null -eq $prop.Value) { return '' }
    return ([string]$prop.Value).Trim()
}

function Test-RVToolsTrue {
    # RVTools / ImportExcel deliver booleans as $true / 'True' / 'TRUE' / 1 depending on the export.
    param($Value, [string]$Pattern = '^(true|yes|1)$')
    if ($null -eq $Value) { return $false }
    return ([string]$Value).Trim() -match $Pattern
}

function ConvertTo-StrategyDate {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact($Text.Trim(), 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    return $null
}

function Get-OSStateGroup {
    # Collapses an OS_Status string into the three buckets that matter for strategy (same as Azure).
    param([string]$Status)
    if ($Status -like 'Supported*')                                 { return 'OK' }
    if ($Status -like 'Extended*' -or $Status -like 'Unsupported*') { return 'ACTION' }
    return 'VERIFY'
}

function Get-OSDistro {
    # Coarse OS family (setting OsFamilies) used to tell a benign Config-vs-Tools mismatch from a material one.
    param([string]$OSString, [Parameter(Mandatory)] [hashtable]$Settings)
    if ([string]::IsNullOrWhiteSpace($OSString)) { return 'Unknown' }
    foreach ($fam in $Settings.OsFamilies) {
        if ($OSString -match $fam.Match) { return $fam.Family }
    }
    return 'Unknown'
}

function ConvertTo-InvariantNumber {
    # RVTools numbers parsed with the invariant culture (a de-DE Windows would otherwise misread decimals).
    param([string]$Text)
    $n = 0.0
    if ([double]::TryParse($Text, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { return $n }
    return 0.0
}

#endregion

#region ------------------------- OS lifecycle and VM Import -------------------------

function Get-AWSOSLifecycle {
    <#
        Returns Status / Label / EndOfSupport / ExtendedEnd / Note for one OS string, using the OsLifecycle
        table of the settings file. Status values (Azure-compatible wording):
          Supported | Supported - ends before cut-off | Extended support only (ESU/ELS/LTSS) |
          Unsupported - End of life | Unsupported - 32-bit OS | Verify - version not identified | Verify - OS not identified
    #>
    param(
        [string]$OSString,
        [Parameter(Mandatory)] [hashtable]$Settings,
        [datetime]$AsOf = (Get-Date),
        # Optional: @{ Label; EndOfSupport; ExtendedEnd; Note; DateSource } used instead of the OsLifecycle row
        # (the exact version the client picked in the workbook)
        [hashtable]$LifecycleOverride = $null
    )
    $out = [ordered]@{ Status = ''; Label = ''; EndOfSupport = ''; ExtendedEnd = ''; Note = ''; DateSource = '' }

    if ([string]::IsNullOrWhiteSpace($OSString) -or (Test-AnyPattern $OSString $Settings.UnknownOsPatterns)) {
        $out.Status = 'Verify - OS not identified'; $out.Label = 'Unknown'
        return [PSCustomObject]$out
    }
    if (-not $Settings.Allow32BitImport -and (Test-AnyPattern $OSString $Settings.Arch32BitPatterns)) {
        $out.Status = 'Unsupported - 32-bit OS'; $out.Label = $OSString
        $out.Note   = $Settings.Arch32BitNote
        return [PSCustomObject]$out
    }

    $rule = $null
    if ($null -ne $LifecycleOverride) { $rule = $LifecycleOverride }
    else {
        foreach ($r in $Settings.OsLifecycle) {
            if ($OSString -match $r.Match) { $rule = $r; break }
        }
    }
    if ($null -eq $rule) {
        $out.Status = 'Verify - version not identified'; $out.Label = $OSString
        $out.Note   = 'Not in the OsLifecycle table - confirm distro/version inside the guest'
        return [PSCustomObject]$out
    }

    $out.Label = $rule.Label; $out.EndOfSupport = [string]$rule.EndOfSupport; $out.ExtendedEnd = [string]$rule.ExtendedEnd; $out.Note = [string]$rule.Note
    $out.DateSource = if ([string]::IsNullOrWhiteSpace([string]$rule.DateSource)) { 'settings file' } else { [string]$rule.DateSource }
    $end    = ConvertTo-StrategyDate $rule.EndOfSupport
    $ext    = ConvertTo-StrategyDate $rule.ExtendedEnd
    $cutoff = ConvertTo-StrategyDate $Settings.SupportCutoffDate

    if ($null -eq $end) {
        $out.Status = 'Verify - version not identified'
    }
    elseif ($end -lt $AsOf) {
        if ($null -ne $ext -and $ext -ge $AsOf) { $out.Status = 'Extended support only (ESU/ELS/LTSS)' }
        else                                      { $out.Status = 'Unsupported - End of life' }
    }
    elseif ($null -ne $cutoff -and $end -lt $cutoff) {
        $out.Status = 'Supported - ends before cut-off'
    }
    else {
        $out.Status = 'Supported'
    }
    return [PSCustomObject]$out
}

function Get-AWSVmImportStatus {
    # 'Supported' | 'Not supported' | 'Version not identified' | 'Unknown'  (unsupported list checked first)
    param([string]$OSString, [Parameter(Mandatory)] [hashtable]$Settings)
    if ([string]::IsNullOrWhiteSpace($OSString) -or (Test-AnyPattern $OSString $Settings.UnknownOsPatterns)) { return 'Unknown' }
    if (-not $Settings.Allow32BitImport -and (Test-AnyPattern $OSString $Settings.Arch32BitPatterns))       { return 'Not supported' }
    if (Test-AnyPattern $OSString $Settings.VmImportUnsupportedOs)                                         { return 'Not supported' }
    if (Test-AnyPattern $OSString $Settings.UnversionedOsPatterns)                                         { return 'Version not identified' }
    if (Test-AnyPattern $OSString $Settings.VmImportVerifyOs)                                              { return 'Version not identified' }
    if (Test-AnyPattern $OSString $Settings.VmImportSupportedOs)                                           { return 'Supported' }
    return 'Version not identified'
}

#endregion

#region ------------------------- Inventory (facts per VM) -------------------------

function Build-StrategyInventory {
    <#
        Joins vInfo (templates excluded) with the vDisk rows of the same export and returns one
        fact object per VM. Every later step reads only these facts.
    #>
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array]$VInfoData,
        [array]$VDiskData = @(),
        [hashtable]$OSOverrides = @{},
        [Parameter(Mandatory)] [hashtable]$Settings
    )

    # Index the optional tabs by VM name (templates are dropped later together with their vInfo row)
    $diskByVm = @{}; foreach ($d in $VDiskData) { $k = (Get-RowValue $d 'VM').ToLower(); if ($k) { if (-not $diskByVm.ContainsKey($k)) { $diskByVm[$k] = New-Object System.Collections.Generic.List[object] }; $diskByVm[$k].Add($d) } }

    $inventory = foreach ($vm in $VInfoData) {
        if (Test-RVToolsTrue (Get-RowValue $vm 'Template') $Settings.TemplateValuePattern) { continue }

        $name = Get-RowValue $vm 'VM'
        if (-not $name) { continue }
        $key  = $name.ToLower()

        $annotation = Get-RowValue $vm 'Annotation'
        $purpose = ''
        if ($annotation -match $Settings.AnnotationPurposeRegex) { $purpose = $Matches[1].Trim() }

        # Custom attributes: only the columns present in this export
        $tagValues = foreach ($col in $Settings.TagColumnsSearched) { $v = Get-RowValue $vm $col; if ($v) { $v } }
        $tagText   = (@($tagValues) -join ' ')
        $applianceTagHits = @(foreach ($col in ($Settings.ApplianceTags.Keys | Sort-Object)) {
            $v = Get-RowValue $vm $col
            if ($v -and $v -match $Settings.ApplianceTags[$col]) { "$col = $v" }
        })

        # Disks: vDisk rows when available, else vInfo aggregate
        $vmDisks = if ($diskByVm.ContainsKey($key)) { $diskByVm[$key].ToArray() } else { @() }
        $diskCount = if ($vmDisks.Count -gt 0) { $vmDisks.Count } else { [int](ConvertTo-InvariantNumber (Get-RowValue $vm 'Disks')) }
        $largestGiB = 0
        $rawShared  = New-Object System.Collections.Generic.List[string]
        foreach ($d in $vmDisks) {
            $capMiB = ConvertTo-InvariantNumber (Get-RowValue $d 'Capacity MiB')
            if (($capMiB / 1024) -gt $largestGiB) { $largestGiB = $capMiB / 1024 }
            $label = Get-RowValue $d 'Disk'
            if (Test-RVToolsTrue (Get-RowValue $d 'Raw') $Settings.RawDiskValuePattern)   { $rawShared.Add("$label raw (RDM)") }
            $mode = Get-RowValue $d 'Sharing mode'
            if ($mode -and $mode -match $Settings.SharedModePattern)                       { $rawShared.Add("$label sharing mode '$mode'") }
            $bus  = Get-RowValue $d 'Shared Bus'
            if ($bus -and $bus -match $Settings.SharedBusPattern)                          { $rawShared.Add("$label shared bus '$bus'") }
        }


        $storageMiB = ConvertTo-InvariantNumber (Get-RowValue $vm $Settings.StorageColumn)
        $memMiB     = ConvertTo-InvariantNumber (Get-RowValue $vm 'Memory')
        $cpus       = [int](ConvertTo-InvariantNumber (Get-RowValue $vm 'CPUs'))
        $nics       = [int](ConvertTo-InvariantNumber (Get-RowValue $vm 'NICs'))

        $cluster = Get-RowValue $vm 'Cluster'
        $ovr     = if ($OSOverrides.ContainsKey($key)) { [string]$OSOverrides[$key] } else { '' }

        [PSCustomObject][ordered]@{
            VMName             = $name
            VM_ID              = Get-RowValue $vm 'VM ID'
            Powerstate         = Get-RowValue $vm 'Powerstate'
            Cluster            = $cluster
            Host               = Get-RowValue $vm 'Host'
            OSConfig           = Get-RowValue $vm 'OS according to the configuration file'
            OSTools            = Get-RowValue $vm 'OS according to the VMware Tools'
            OSOverride         = $ovr
            Annotation         = $annotation
            Purpose            = $purpose
            TagText            = $tagText
            ApplianceTagHits   = $applianceTagHits
            CurrentVCPU        = $cpus
            CurrentRamGB       = [math]::Round($memMiB / 1024, 2)
            StorageGB          = [math]::Round($storageMiB / 1024, 2)
            DiskCount          = $diskCount
            LargestDiskGiB     = [math]::Round($largestGiB, 2)
            RawOrSharedDisks   = @($rawShared)
            NicCount           = $nics
            ClusterRules       = Get-RowValue $vm 'Cluster rule(s)'
            ClusterRuleNames   = Get-RowValue $vm 'Cluster rule name(s)'
            IsDmz              = (Test-AnyPattern $name $Settings.DmzPatterns) -or (Test-AnyPattern $cluster $Settings.DmzPatterns)
        }
    }
    $inventory = @($inventory | Sort-Object VMName)
    $dupes = @($inventory | Group-Object { $_.VMName.ToLower() } | Where-Object Count -gt 1)
    foreach ($g in $dupes) {
        Write-Warning "VM name '$($g.Name)' appears $($g.Count) times in vInfo - all rows are kept; vDisk rows are joined by name and are shared between them."
    }
    return ,$inventory
}

#endregion

#region ------------------------- Facts used by both the decision chain and the Parameter_Map -------------------------

function Get-StrategyFacts {
    # Derives every signal for one VM once, so the decision chain and the Parameter_Map counts agree.
    param(
        [Parameter(Mandatory)] $Vm,
        [Parameter(Mandatory)] [hashtable]$Settings,
        [datetime]$AsOf = (Get-Date)
    )

    # ---------------- OS used for the decision (override file, then the chosen OS field, then the other) ----------------
    $precedence = if ($Settings.ContainsKey('OsSourcePrecedence')) { $Settings.OsSourcePrecedence } else { (Get-SettingsForOSField -Settings $Settings).OsSourcePrecedence }
    $osUsed = ''; $osSource = ''
    foreach ($src in $precedence) {
        $t = switch ($src) { 'Override' { $Vm.OSOverride } 'VMwareTools' { $Vm.OSTools } 'ConfigFile' { $Vm.OSConfig } default { '' } }
        if (-not [string]::IsNullOrWhiteSpace($t)) { $osUsed = $t; $osSource = $src; break }
    }

    # Exact version picked by the client in the workbook (scenario rows of Get-VersionScenarioRows only).
    # A guest-check override file always wins, so the choice never replaces an override.
    $choice = if ($Vm.PSObject.Properties['VersionChoice']) { $Vm.VersionChoice } else { $null }
    if ($null -ne $choice -and $osSource -ne 'Override') {
        $osUsed = [string]$choice.OsString; $osSource = 'Client choice (workbook)'
        $lc = Get-AWSOSLifecycle -OSString $osUsed -Settings $Settings -AsOf $AsOf -LifecycleOverride $choice.Lifecycle
    }
    else {
        $lc = Get-AWSOSLifecycle -OSString $osUsed -Settings $Settings -AsOf $AsOf
    }

    # ---------------- Config-vs-Tools mismatch (only when no override) ----------------
    $mismatch = 'None'; $mismatchNote = ''
    if ($osSource -ne 'Override' -and $Vm.OSConfig -and $Vm.OSTools -and ($Vm.OSConfig -ne $Vm.OSTools)) {
        $dCfg = Get-OSDistro $Vm.OSConfig $Settings; $dTool = Get-OSDistro $Vm.OSTools $Settings
        $gCfg = Get-OSStateGroup (Get-AWSOSLifecycle -OSString $Vm.OSConfig -Settings $Settings -AsOf $AsOf).Status
        $gTool = Get-OSStateGroup (Get-AWSOSLifecycle -OSString $Vm.OSTools -Settings $Settings -AsOf $AsOf).Status
        if ($dCfg -in $Settings.VagueOsFamilies -or $dTool -in $Settings.VagueOsFamilies) { $mismatch = 'Benign'; $mismatchNote = 'one side is a generic OS label' }
        elseif ($dCfg -ne $dTool)                   { $mismatch = 'Material'; $mismatchNote = "OS family differs ($dCfg vs $dTool)" }
        elseif ($gCfg -ne $gTool -and $gCfg -ne 'VERIFY' -and $gTool -ne 'VERIFY') { $mismatch = 'Material'; $mismatchNote = "support state differs ($gCfg vs $gTool)" }
        elseif ($gCfg -ne $gTool) { $mismatch = 'Benign'; $mismatchNote = "same OS family; one side's version is not identifiable" }
        else                                        { $mismatch = 'Benign'; $mismatchNote = 'same OS family and support state' }
    }

    $allOS     = "$($Vm.OSOverride) | $($Vm.OSConfig) | $($Vm.OSTools)"
    $fullText  = "$($Vm.VMName) | $($Vm.Annotation) | $allOS | $($Vm.TagText)"

    # Vendor product: first catalogue entry that matches (In = 'Name' -> VM name only)
    $vendor = $null
    foreach ($vp in $Settings.VendorProducts) {
        $subject = if ($vp.In -eq 'Name') { $Vm.VMName } else { $fullText }
        if ($subject -match $vp.Match) { $vendor = $vp; break }
    }
    # Database server (R05): first DatabaseProducts entry that matches (In = 'Name' -> VM name only)
    $database = $null
    foreach ($db in $Settings.DatabaseProducts) {
        $subject = if ($db.In -eq 'Name') { $Vm.VMName } else { $fullText }
        if ($subject -match $db.Match) { $database = $db; break }
    }

    $importStatus = Get-AWSVmImportStatus -OSString $osUsed -Settings $Settings

    [PSCustomObject][ordered]@{
        OSUsed            = $osUsed
        OSSource          = $osSource
        Lifecycle         = $lc
        OSStatus          = [string]$lc.Status
        OSGroup           = Get-OSStateGroup $lc.Status
        Mismatch          = $mismatch
        MismatchNote      = $mismatchNote
        ImportStatus      = $importStatus
        InfraHit          = if (-not $Settings.EnableInfraRule) { $null } else { Get-FirstNonEmpty (Get-FirstPatternHit $Vm.VMName $Settings.InfraNamePatterns) (Get-FirstPatternHit "$($Vm.Annotation) | $allOS" $Settings.InfraTextPatterns) }
        IsPoweredOff      = ($Vm.Powerstate -in $Settings.PoweredOffStates)
        Vendor            = $vendor
        Is32BitUsed       = (-not $Settings.Allow32BitImport) -and (Test-AnyPattern $osUsed $Settings.Arch32BitPatterns)
        Is32BitAnyField   = (Test-AnyPattern $Vm.OSConfig $Settings.Arch32BitPatterns) -or (Test-AnyPattern $Vm.OSTools $Settings.Arch32BitPatterns)
        IsDesktop         = (Test-AnyPattern $osUsed $Settings.DesktopOsPatterns) -and ($osUsed -notmatch $Settings.DesktopExcludePattern)
        IsApplianceTagged = (@($Vm.ApplianceTagHits).Count -gt 0)
        Database          = $database
        IsLegacyOS        = ($lc.Status -like 'Extended*' -or $lc.Status -eq 'Unsupported - End of life' -or $lc.Status -eq 'Supported - ends before cut-off')
        IsPastAllSupport  = ($lc.Status -eq 'Unsupported - End of life')
        IsGroupedLabel    = (Test-AnyPattern $osUsed $Settings.GroupedOsLabels)
        IsUnversioned     = (Test-AnyPattern $osUsed $Settings.UnversionedOsPatterns) -or (($lc.Status -like 'Verify*') -and -not (Test-AnyPattern $osUsed $Settings.GroupedOsLabels))
        IsStale           = Test-AnyPattern $Vm.VMName $Settings.StaleNamePatterns
        IsCluster         = Test-AnyPattern $Vm.VMName $Settings.ClusterNamePatterns
        DiskTooLarge      = ($Vm.LargestDiskGiB -ge $Settings.MaxEbsVolumeGiB)
        DiskNeedsManifest = ($Vm.LargestDiskGiB -ge $Settings.ManifestImportGiB -and $Vm.LargestDiskGiB -lt $Settings.MaxEbsVolumeGiB)
        TooManyVolumes    = ($Vm.DiskCount -gt $Settings.MaxVolumesPerImport)
        HasRawShared      = (@($Vm.RawOrSharedDisks).Count -gt 0)
        HasAntiAffinity   = ($Vm.ClusterRules -match $Settings.AntiAffinityPattern)
    }
}

#endregion

#region ------------------------- Decision chain -------------------------

function Get-MigrationStrategyRows {
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array]$StrategyInventory,
        [Parameter(Mandatory)] [hashtable]$Settings,
        # Which RVTools OS column decides: 'VMwareTools' or 'ConfigFile' (empty = setting DefaultOSField)
        [ValidateSet('', 'VMwareTools', 'ConfigFile')] [string]$OSField = '',
        # Past-all-support OS: Relocate (Amazon EVS / VMware Cloud on AWS) instead of Replatform
        [switch]$PreferRelocateForLegacyOS,
        [datetime]$AsOf = (Get-Date)
    )

    $Settings       = Get-SettingsForOSField -Settings $Settings -OSField $OSField
    $preferRelocate = [bool]$PreferRelocateForLegacyOS -or [bool]$Settings.PreferRelocateForLegacyOS
    $migrating      = @('Rehost', 'Replatform', 'Relocate')

    $rows = foreach ($vm in $StrategyInventory) {
        $f = Get-StrategyFacts -Vm $vm -Settings $Settings -AsOf $AsOf

        $actions  = New-Object System.Collections.Generic.List[string]
        $blockers = New-Object System.Collections.Generic.List[string]
        $confirm  = New-Object System.Collections.Generic.List[string]
        $design   = New-Object System.Collections.Generic.List[string]
        $flags    = New-Object System.Collections.Generic.List[string]

        $osName = if ($f.Lifecycle.Label) { $f.Lifecycle.Label } else { $f.OSUsed }
        $osEnd  = [string]$f.Lifecycle.EndOfSupport

        # Exact OS version: the RVTools label covers several versions with different support dates (setting VersionChoices).
        # Scenario rows (client choice) get their note from Get-VersionScenarioRows; the exporter clears this note when the
        # workbook has no dropdown for the VM (no lifecycle data for the family).
        $versionNote = ''
        if ($f.OSSource -notin @('Override', 'Client choice (workbook)')) {
            $vcFam = Get-VersionChoiceFamily -OSString $f.OSUsed -Settings $Settings
            if ($null -ne $vcFam) {
                $versionNote = "Not checked yet: '$($vcFam.Family)' has versions with different support dates. $($vcFam.Help), then pick the version in column Exact_OS_Version (sheet OS_Versions explains how). Until then the RVTools label is used."
            }
        }

        # ---------------- Review flags (never change the strategy) ----------------
        if ($f.IsPoweredOff -and $Settings.PoweredOffAction -eq 'Flag') { $flags.Add("$($vm.Powerstate) - confirm the VM is still required before migrating (possible Retire)") }
        if ($f.IsStale) { $flags.Add('Name suggests a restore / backup / old copy / test VM - possible Retire') }
        if ($f.Mismatch -eq 'Benign') { $flags.Add("OS fields differ (config file '$($vm.OSConfig)' vs VMware Tools '$($vm.OSTools)') - $($f.MismatchNote); $($f.OSSource) value used") }
        if ($f.Mismatch -eq 'Material') {
            $flags.Add("OS fields disagree materially (config file '$($vm.OSConfig)' vs VMware Tools '$($vm.OSTools)': $($f.MismatchNote)) - $($f.OSSource) value used; confirm in the guest or switch the OS field")
        }

        # ---------------- Design notes (never change the strategy) ----------------
        if ($vm.StorageGB -ge $Settings.LargeDiskGB) { $design.Add(("Large storage: {0} TiB across {1} disk(s) - plan EBS volume layout and IOPS (gp3/io2)" -f [math]::Round($vm.StorageGB / 1024, 1), $vm.DiskCount)) }
        if ($vm.DiskCount -ge $Settings.HighDiskCount) { $design.Add("High disk count ($($vm.DiskCount)): check the instance type's EBS volume attachment limit") }
        if ($vm.NicCount -ge $Settings.MultiNicThreshold) { $design.Add("Multi-NIC ($($vm.NicCount)): VM Import creates one NIC - record each NIC's network/IP and add ENIs after import") }
        if ($vm.CurrentRamGB -gt $Settings.HighRamGB) { $design.Add(("High memory ({0} GB): confirm role (database / in-memory app); consider a memory-optimised instance" -f $vm.CurrentRamGB)) }
        if ($f.IsCluster) { $design.Add('Name suggests clustering (Always On / FCI / NLB): migrate the cluster as a unit; plan shared storage or AZ placement') }
        if ($f.HasAntiAffinity) { $design.Add("Anti-affinity rule '$($vm.ClusterRuleNames)': place the members in different Availability Zones or an EC2 spread placement group") }

        $strategy = ''; $rule = ''; $pm = ''; $conf = ''; $why = ''; $alt = ''; $target = ''

        # ======================= DECIDE THE STRATEGY: R02-R10, first match wins =======================
        # (keep in sync with Get-MigrationStrategyRuleCatalog; R01 Template is applied in Build-StrategyInventory)
        if ($f.InfraHit) {
            $strategy = 'Retire'; $rule = 'R02'; $pm = 'Azure R01'; $conf = 'High'
            $why = "VMware / storage platform component (matches '$($f.InfraHit)') - not needed in AWS"
            $actions.Clear(); $design.Clear()
            $actions.Add('Keep until the platform it serves is decommissioned; take a final backup')
            $alt = 'None expected'; $target = $Settings.Targets.Decommission
        }
        elseif ($f.IsDesktop) {
            $strategy = 'Replace'; $rule = 'R03'; $pm = 'PM #15'; $conf = 'Medium'
            $why = "Windows desktop OS ('$($f.OSUsed)') - Windows client licences on EC2 are BYOL only and need Dedicated Hosts"
            $actions.Add('Confirm users, GPU needs, Windows edition and licence eligibility')
            if ($f.OSGroup -eq 'ACTION') { $actions.Add("OS status is '$($f.OSStatus)' - upgrade to Windows 11 or buy ESU") }
            $confirm.Add('Microsoft licence position (Enterprise E3/E5 / VDA, hosted rights)')
            $alt = 'EC2 with BYOL on Dedicated Hosts; Retire if no longer used'; $target = $Settings.DesktopTarget
        }
        elseif ($f.Vendor -and $f.Vendor.Action -eq 'Retire') {
            $strategy = 'Retire'; $rule = 'R04'; $pm = 'PM #14'; $conf = 'High'
            $why = "Vendor product ended by the vendor: $($f.Vendor.Product) (matches '$($f.Vendor.Match)')"
            $actions.Clear(); $design.Clear()
            $actions.Add('Confirm nothing still consumes it; use the vendor successor service if the function is still needed')
            $alt = 'Vendor successor (cloud) service'; $target = $f.Vendor.Target
        }
        elseif ($f.Vendor) {
            $strategy = 'Replace'; $rule = 'R04'; $pm = 'PM #14'; $conf = 'Medium'
            $why = "Vendor product: $($f.Vendor.Product) (matches '$($f.Vendor.Match)') - deploy the vendor's AWS edition, do not lift the VM"
            $actions.Add('Export the configuration; deploy the vendor AWS edition / SaaS; transfer the licence with the vendor')
            $confirm.Add('Vendor AWS support and licence portability')
            $alt = 'Retain if the vendor has no AWS edition'; $target = $f.Vendor.Target
        }
        elseif ($f.IsApplianceTagged) {
            $strategy = 'Replace'; $rule = 'R04'; $pm = 'PM #13'; $conf = 'Low'
            $why = 'Tagged as an appliance by the client (' + (@($vm.ApplianceTagHits) -join ', ') + ') - product not identifiable from RVTools'
            $actions.Add('Owner names the product; vendor confirms AWS support')
            $confirm.Add('Product name and vendor AWS support')
            $alt = 'Rehost if the vendor supports VM import; Retain if no AWS edition'; $target = $Settings.ApplianceTagTarget
        }
        elseif ($f.Database) {
            $strategy = $Settings.DatabaseStrategy; $rule = 'R05'; $pm = 'new (Azure R07)'; $conf = 'Low'
            $why = if ($strategy -eq 'Manual Review') { "Database server: name suggests $($f.Database.Engine) (matches '$($f.Database.Match)') - every database is reviewed before choosing a managed AWS database or EC2" }
                   else { "Database server: name suggests $($f.Database.Engine) (matches '$($f.Database.Match)') - move the database to a managed AWS service instead of copying the server" }
            $confirm.Add('Database engine, version and edition; size; features used (SSIS / SSRS / SSAS, linked servers, agents)')
            if ($strategy -eq 'Replatform') {
                $actions.Add('Migrate the database with AWS DMS or native backup / restore; repoint the applications; then retire the VM')
                $alt = 'Rehost to EC2 (self-managed database) if the engine, version or features are not supported by the managed service'
                $target = $f.Database.Target
            }
            elseif ($strategy -eq 'Rehost') {
                $alt = "Replatform to $($f.Database.Target)"; $target = $Settings.Targets.Ec2
                if ($f.ImportStatus -eq 'Not supported') { $blockers.Add("OS not accepted by AWS VM Import ('$($f.OSUsed)'): upgrade it in place before export, or rebuild") }
            }
            else {
                $actions.Add("Review the database (engine, version, edition, size, features) and choose Replatform to $($f.Database.Target) or Rehost to EC2")
                $alt = "Replatform to $($f.Database.Target), or Rehost to EC2 (self-managed database)"
                $target = "To be decided after review: $($f.Database.Target) or Amazon EC2"
            }
        }
        elseif ($f.IsLegacyOS -or $f.Is32BitUsed) {
            $rule = 'R06'; $pm = if ($f.Is32BitUsed) { 'PM #11' } else { 'PM #18, #32' }
            $conf = if ($f.OSSource -in @('Override', 'VMwareTools', 'Client choice (workbook)') -and $f.Mismatch -eq 'None') { 'High' } else { 'Medium' }
            if ($f.Is32BitUsed) {
                # 32-bit cannot be imported and cannot be upgraded in place to 64-bit: always rebuild
                $strategy = 'Replatform'
                $why = "32-bit OS ('$($f.OSUsed)') - $($Settings.Arch32BitNote); a 32-bit OS cannot be upgraded in place to 64-bit"
                $actions.Add('Rebuild on a current 64-bit AMI and move the application / data; validate the application on 64-bit')
                $alt = 'Retain on-premises if the application cannot run on 64-bit; Retire if no longer used'; $target = $Settings.Targets.Ec2Os64Bit
            }
            else {
                $why = "Guest OS '$osName' is '$($f.OSStatus)'" + $(if ($osEnd) { " (end of support $osEnd)" } else { '' })
                if ($preferRelocate -and $f.IsPastAllSupport) {
                    $strategy = 'Relocate'
                    $actions.Add('Move as-is to a VMware platform in AWS (no OS change); plan OS modernisation later')
                    $alt = 'Replatform (upgrade / rebuild on a supported OS)'; $target = $Settings.RelocateTarget
                }
                elseif ($Settings.LegacyOsStrategy -eq 'Rehost') {
                    $strategy = 'Rehost'
                    $actions.Add("Out-of-support OS (support ended/ends $osEnd): rehost as-is, then upgrade or buy extended support; keep it in an isolated subnet")
                    $alt = 'Replatform (upgrade during migration)'; $target = $Settings.Targets.Ec2
                }
                else {
                    $strategy = 'Replatform'
                    $actions.Add('Upgrade in place before export, or rebuild on a current AMI and move the application/data; validate application compatibility')
                    $alt = if ($f.IsPastAllSupport) { "Relocate ($($Settings.RelocateTarget)) if the OS cannot be upgraded; Retire if no longer needed" } else { 'Rehost first and upgrade after cutover, covered by extended support (ESU / ELS / LTSS)' }
                    $target = $Settings.Targets.Ec2SupportedOs
                }
            }
            # Relocate moves the VM with HCX (no image import), so the VM Import blocker does not apply there
            if ($f.ImportStatus -eq 'Not supported' -and $strategy -ne 'Relocate') { $blockers.Add("OS not accepted by AWS VM Import ('$($f.OSUsed)') - rebuild instead of importing") }
        }
        elseif ($f.ImportStatus -eq 'Not supported') {
            $strategy = 'Rehost'; $rule = 'R07'; $pm = 'PM #19'; $conf = 'Medium'
            $why = "OS not accepted by AWS VM Import ('$($f.OSUsed)')"
            $blockers.Add('OS not accepted by AWS VM Import: upgrade it in place to a supported version before export (or rebuild)')
            $alt = 'Replatform (rebuild on a supported OS)'; $target = $Settings.Targets.Ec2
        }
        elseif ($f.ImportStatus -in @('Version not identified', 'Unknown') -or $f.IsUnversioned) {
            $strategy = 'Rehost'; $rule = 'R08'; $pm = 'PM #20'; $conf = 'Medium'
            $why = "OS family / version not identifiable from RVTools ('$($f.OSUsed)'); no other blocker found"
            $actions.Add('Confirm the exact OS version inside the guest and compare with the AWS VM Import list (re-run with -OSOverridePath)')
            $confirm.Add('Exact OS version (guest check)')
            $alt = 'Replatform if the OS turns out to be end-of-life or not importable'; $target = $Settings.Targets.Ec2
        }
        elseif ($f.ImportStatus -eq 'Supported') {
            $strategy = 'Rehost'; $rule = 'R09'; $pm = 'PM #17'
            $conf = if ($Settings.AllowHighRehostConfidence -and $f.OSSource -in @('Override', 'VMwareTools', 'Client choice (workbook)') -and $f.Mismatch -eq 'None' -and $f.OSStatus -eq 'Supported') { 'High' } else { $Settings.DefaultRehostConfidence }
            $why = "OS '$osName' is on the AWS VM Import list and '$($f.OSStatus)'; agentless image import"
            $alt = 'Refactor / Replatform only if the application owner wants modernisation'; $target = $Settings.Targets.Ec2
            if ($f.IsGroupedLabel) {
                $why += " - grouped VMware label: exact version not shown"
                $actions.Add($Settings.GroupedOsLabelAction)
                $confirm.Add('Exact OS version (guest check)')
                $alt = 'Replatform if the guest check shows a version whose support ends before the cut-off'
            }
        }
        else {
            $strategy = 'Manual Review'; $rule = 'R10'; $pm = ''; $conf = 'Low'
            $why = 'No rule matched'
            $actions.Add('Review manually')
        }

        # ======================= READINESS CHECKS R11-R12: every check that applies; never change the strategy =======================
        $checks = New-Object System.Collections.Generic.List[string]
        if ($strategy -in @('Rehost', 'Replatform') -and ($f.DiskTooLarge -or $f.DiskNeedsManifest -or $f.TooManyVolumes -or $f.HasRawShared)) { $checks.Add('R11') }
        if ($strategy -in @('Rehost', 'Replatform')) {
            if ($f.DiskTooLarge)      { $actions.Add(("R11 Disk limits: disk of {0} GiB - AWS VM Import accepts disk images below {1} GiB; split it or move the data to FSx/S3 before export" -f $vm.LargestDiskGiB, $Settings.MaxEbsVolumeGiB)) }
            if ($f.DiskNeedsManifest) { $actions.Add(("R11 Disk limits: disk of {0} GiB is above {1} GiB - use the ImportInstance / manifest import route" -f $vm.LargestDiskGiB, $Settings.ManifestImportGiB)) }
            if ($f.TooManyVolumes)    { $actions.Add("R11 Disk limits: $($vm.DiskCount) disks - VM Import accepts up to $($Settings.MaxVolumesPerImport) volumes per task; import the extra disks separately") }
            if ($f.HasRawShared)      { $blockers.Add('R11 Disk limits: raw (RDM) or shared disks cannot be image-exported: ' + (@($vm.RawOrSharedDisks) -join ', ')) }
        }
        if ($strategy -in $migrating -and $vm.IsDmz) {
            $actions.Add('R12 DMZ server: moves in the DMZ wave, after the AWS perimeter (inspection VPC) is live')
            $checks.Add('R12')
        }

        # ======================= FINAL CHECK R13: powered-off VMs are never decommissioned automatically =======================
        # The rules above have already decided what the VM gets if it is still needed; that result is kept in
        # Rationale, Alternative_Strategy and AWS_Target, and the strategy is set to Client Review.
        $clientReview = $false
        if ($f.IsPoweredOff -and $Settings.PoweredOffAction -eq 'ClientReview') {
            $clientReview = $true
            $baseStrategy = $strategy; $baseRule = $rule; $baseWhy = $why
            $flags.Insert(0, "Powered off ($($vm.Powerstate)) - migration strategy to be reviewed with the client")
            $why = "Powered off ($($vm.Powerstate)) in RVTools; RVTools does not show since when or why (standby, DR or seasonal use is possible). If the VM is still needed the rules give: $baseStrategy ($baseRule - $baseWhy)"
            $alt = if ($baseStrategy -eq 'Retire') { "Retire ($baseRule) - after owner confirmation and a final backup" }
                   else { "If still needed: $baseStrategy ($baseRule); if not needed: Retire after owner confirmation and a final backup" }
            $strategy = $Settings.PoweredOffStrategy; $rule = 'R13'; $pm = 'PM #2'; $conf = 'Low'
            $confirm.Insert(0, 'Is the VM still needed (standby / DR / seasonal use)?')
            $confirm.Insert(0, 'Date powered off')
            # $target stays: the AWS target if the VM is kept (see Legend)
        }

        # ---------------- Readiness value ----------------
        $readiness = if ($clientReview) { 'Awaiting client input' }
        elseif ($strategy -in @('Retire', 'Retain')) {
            if ($confirm.Count -gt 0) { 'Awaiting client input' } else { 'Not migrating' }
        }
        elseif ($blockers.Count -gt 0) { 'Blocked' }
        elseif ($confirm.Count -gt 0)  { 'Awaiting client input' }
        elseif ($actions.Count -gt 0)  { 'Ready after prep' }
        else                           { 'Ready' }

        $allActions = @($blockers | ForEach-Object { "BLOCKER: $_" }) + @($actions)

        [PSCustomObject][ordered]@{
            VMName                = $vm.VMName
            Powerstate            = $vm.Powerstate
            Cluster               = $vm.Cluster
            Host                  = $vm.Host
            OS_ConfigFile         = $vm.OSConfig
            OS_VMwareTools        = $vm.OSTools
            OS_Used               = $f.OSUsed
            OS_Source             = $f.OSSource
            OS_Status             = $f.OSStatus
            OS_EndOfSupport       = $osEnd
            OS_Extended_Support_End = [string]$f.Lifecycle.ExtendedEnd
            Version_Note          = $versionNote
            OS_Mismatch           = $f.Mismatch
            VM_Import_OS          = $f.ImportStatus
            Current_vCPU          = $vm.CurrentVCPU
            Current_RAM_GB        = $vm.CurrentRamGB
            Current_Disk_GB       = $vm.StorageGB
            Largest_Disk_GiB      = $vm.LargestDiskGiB
            Current_Disk_Count    = $vm.DiskCount
            Current_NIC_Count     = $vm.NicCount
            Migration_Strategy    = $strategy
            Strategy_Confidence   = $conf
            Rule_Applied          = $rule
            Parameter_Map_Ref     = $pm
            Rationale             = $why
            AWS_Target            = $target
            Readiness             = $readiness
            Readiness_Checks      = ($checks -join ', ')
            Pre_Migration_Actions = ($allActions -join '; ')
            Client_Confirmation   = ($confirm -join '; ')
            Alternative_Strategy  = $alt
            Design_Considerations = ($design -join '; ')
            Review_Flags          = ($flags -join '; ')
        }
    }
    return @($rows)
}

#endregion

#region ------------------------- Exact OS version chooser (setting VersionChoices) -------------------------

function Get-VersionChoiceFamily {
    # First VersionChoices entry whose Match fits the OS string (or $null).
    param([string]$OSString, [Parameter(Mandatory)] [hashtable]$Settings)
    if ([string]::IsNullOrWhiteSpace($OSString) -or -not $Settings.ContainsKey('VersionChoices')) { return $null }
    foreach ($vc in @($Settings.VersionChoices)) {
        if ($OSString -match $vc.Match) { return $vc }
    }
    return $null
}

function Get-CatalogValue {
    # Reads a member of a hashtable or an object (the lifecycle catalog may come from JSON or from PowerShell objects).
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { if ($Object.Contains($Name)) { return $Object[$Name] } else { return $null } }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p) { return $null }
    return $p.Value
}

function ConvertTo-LifecycleDateText {
    # endoflife.date values can be a date, true/false or null: keep 'yyyy-MM-dd' dates only.
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture) }
    $t = ([string]$Value).Trim()
    if ($t -match '^(\d{4}-\d{2}-\d{2})') { return $Matches[1] }
    return ''
}

function Format-StrategyDateLong {
    # 'yyyy-MM-dd' -> '12 Jan 2027' (culture-independent, for client-facing notes)
    param([string]$Text)
    $d = ConvertTo-StrategyDate $Text
    if ($null -eq $d) { return $Text }
    return $d.ToString('d MMM yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-VersionOptionInfo {
    <#
        Client-facing option text and a sort key for one lifecycle release.
          windows-server '2012-r2'     -> 'Windows Server 2012 R2'             (OS text for the rules: OsTemplate -f '2012 R2')
          windows '11-24h2-e'          -> 'Windows 11 24H2 - Enterprise / Education'
          windows '11-24h2-w'          -> 'Windows 11 24H2 - Home / Pro'
          windows '10-22h2'            -> 'Windows 10 22H2 - all editions'
          windows '10-21h2-e-lts'      -> 'Windows 10 Enterprise LTSC 2021 (21H2)'
          windows '10-21h2-iot-lts'    -> 'Windows 10 IoT Enterprise LTSC 2021 (21H2)'
        Unknown cycle shapes fall back to the catalog label (or the cycle name).
    #>
    param([string]$Product, [string]$Cycle, [string]$Label)
    $c = ([string]$Cycle).Trim().ToLower()
    $info = [ordered]@{ Text = ''; Version = ''; Sort = 0.0 }
    if ($Product -eq 'windows-server') {
        $ver = (($c -split '-') | ForEach-Object { if ($_ -match '^r\d$' -or $_ -match '^sp\d$') { $_.ToUpper() } else { $_ } }) -join ' '
        $info.Version = $ver
        $info.Text = "Windows Server $ver"
        $n = 0.0; if ($c -match '^(\d{4})') { $n = [double]$Matches[1] }; if ($c -match '-r2') { $n += 0.5 }
        $info.Sort = $n
        return [pscustomobject]$info
    }
    if ($Product -eq 'windows' -and $c -match '^(10|11)-(\d{2}h\d|\d{4})(.*)$') {
        $major = $Matches[1]; $rel = $Matches[2]; $rest = $Matches[3]
        $relText = $rel.ToUpper()
        $ltscYear = @{ '10-1507' = 'LTSB 2015'; '10-1607' = 'LTSB 2016'; '10-1809' = 'LTSC 2019'; '10-21h2' = 'LTSC 2021'; '11-24h2' = 'LTSC 2024' }
        $relNum = if ($rel -match '^(\d{2})h(\d)$') { [double]("{0}.{1}" -f $Matches[1], $Matches[2]) } else { [double]$rel.Substring(0, 2) + ([double]$rel.Substring(2, 2) / 100) }
        $sort = [double]$major * 100 + $relNum
        if ($rest -match 'lts') {
            $l = if ($ltscYear.ContainsKey("$major-$rel")) { $ltscYear["$major-$rel"] } else { "LTSC ($relText)" }
            $kind = if ($rest -match 'iot') { 'IoT Enterprise' } else { 'Enterprise' }
            $info.Text = "Windows $major $kind $l ($relText)"
            $sort -= 0.002; if ($rest -match 'iot') { $sort -= 0.001 }
        }
        elseif ($rest -match '^-e\b|^-e$') { $info.Text = "Windows $major $relText - Enterprise / Education" }
        elseif ($rest -match '^-w\b|^-w$') { $info.Text = "Windows $major $relText - Home / Pro"; $sort -= 0.0005 }
        else { $info.Text = "Windows $major $relText - all editions" }
        $info.Version = $relText; $info.Sort = $sort
        return [pscustomobject]$info
    }
    $info.Text = if (-not [string]::IsNullOrWhiteSpace($Label)) { $(if ($Label -match '^Windows') { $Label } elseif ($Product -eq 'windows') { "Windows $Label" } else { $Label }) } else { $Cycle }
    $info.Version = $Cycle
    return [pscustomobject]$info
}

function Get-LifecycleSourceText {
    # 'Live' / 'Cache' / 'StaleCache' / 'Snapshot' + retrieval time -> client-facing source text
    param($Catalog, [string]$Product)
    $src = [string](Get-CatalogValue (Get-CatalogValue $Catalog 'Source') $Product)
    $at  = [string](Get-CatalogValue (Get-CatalogValue $Catalog 'RetrievedAt') $Product)
    $when = if ($at) { " ($at)" } else { '' }
    switch ($src) {
        'Live'       { return "endoflife.date - live$when" }
        'Cache'      { return "endoflife.date - cached copy$when" }
        'StaleCache' { return "endoflife.date - cached copy, out of date$when - verify" }
        'Snapshot'   { return 'offline snapshot shipped with the tool - verify' }
        default      { if ($src) { return $src } else { return 'lifecycle data' } }
    }
}

function Get-VersionChoiceCatalog {
    <#
        Options of the exact-version dropdown, per VersionChoices family, from the lifecycle catalog
        (Get-LifecycleCatalog in Modules\LiveDataProviders.ps1). Families without lifecycle data are left out (warning),
        so their VMs keep today's result and get no dropdown.
        Returns @( [pscustomobject]@{ Family; Match; Help; Product; Options = @( [pscustomobject]@{ Option; OsString; Version;
                   EndOfSupport; ExtendedEnd; MainstreamEnd; DateSource; Status; Lifecycle } ) } )
    #>
    param(
        [Parameter(Mandatory)] [hashtable]$Settings,
        $LifecycleCatalog = $null,
        [datetime]$AsOf = (Get-Date)
    )
    $out = New-Object System.Collections.Generic.List[object]
    if (@($Settings.VersionChoices).Count -eq 0) { return @() }
    $products = Get-CatalogValue $LifecycleCatalog 'Products'
    foreach ($vc in @($Settings.VersionChoices)) {
        $releases = @(Get-CatalogValue $products $vc.Product)
        $releases = @($releases | Where-Object { $null -ne $_ -and ([string](Get-CatalogValue $_ 'Cycle')) -match $vc.CyclePattern })
        if ($releases.Count -eq 0) {
            Write-Warning "Exact-version chooser: no lifecycle data for '$($vc.Family)' (product '$($vc.Product)') - those VMs keep the RVTools label and get no dropdown."
            continue
        }
        $srcText = Get-LifecycleSourceText -Catalog $LifecycleCatalog -Product $vc.Product
        # Optional HideEndedBefore ('yyyy-MM-dd' or 'today-<n>m' / 'today-<n>y'; '' = keep all): drop releases whose support
        # AND paid extended support both ended before that date (releases with no published end date are kept)
        $hideBefore = $null
        $hideText = if ($vc.ContainsKey('HideEndedBefore')) { ([string]$vc.HideEndedBefore).Trim() } else { '' }
        if ($hideText) {
            if ($null -ne (ConvertTo-StrategyDate $hideText)) { $hideBefore = ConvertTo-StrategyDate $hideText }
            elseif ($hideText -match '^(?i)today\s*-\s*(\d+)\s*m$') { $hideBefore = $AsOf.Date.AddMonths(-[int]$Matches[1]) }
            elseif ($hideText -match '^(?i)today\s*-\s*(\d+)\s*y$') { $hideBefore = $AsOf.Date.AddYears(-[int]$Matches[1]) }
            else { Write-Warning "VersionChoices '$($vc.Family)': HideEndedBefore '$hideText' is not valid (use yyyy-MM-dd, today-<n>m or today-<n>y) - all versions are listed." }
        }
        $seen = @{}
        $opts = foreach ($rel in $releases) {
            $cycle = [string](Get-CatalogValue $rel 'Cycle')
            $info  = Get-VersionOptionInfo -Product $vc.Product -Cycle $cycle -Label ([string](Get-CatalogValue $rel 'Label'))
            if ($seen.ContainsKey($info.Text)) { continue }
            if ($null -ne $hideBefore) {
                $eosD = ConvertTo-StrategyDate (ConvertTo-LifecycleDateText (Get-CatalogValue $rel 'EndOfSupport'))
                $extD = ConvertTo-StrategyDate (ConvertTo-LifecycleDateText (Get-CatalogValue $rel 'ExtendedEnd'))
                $lastD = $eosD; if ($null -ne $extD -and ($null -eq $lastD -or $extD -gt $lastD)) { $lastD = $extD }
                if ($null -ne $eosD -and $lastD -lt $hideBefore) { continue }
            }
            $seen[$info.Text] = $true
            # OsTemplate: {0} = version, optional {1} = '(32-bit)' / '(64-bit)' taken from each VM's own OS string. '{1}' is kept
            # in OsString here and filled per VM by Get-VersionScenarioRows; '(64-bit)' is used for the option's status below.
            $osString = if ($vc.OsTemplate -match '\{[01]\}') { ([string]$vc.OsTemplate).Replace('{0}', [string]$info.Version) } else { [string]$vc.OsTemplate }
            $osString64 = $osString.Replace('{1}', '(64-bit)')
            $lc = @{
                Label        = $info.Text
                EndOfSupport = ConvertTo-LifecycleDateText (Get-CatalogValue $rel 'EndOfSupport')
                ExtendedEnd  = ConvertTo-LifecycleDateText (Get-CatalogValue $rel 'ExtendedEnd')
                Note         = ''
                DateSource   = $srcText
            }
            $status = (Get-AWSOSLifecycle -OSString $osString64 -Settings $Settings -AsOf $AsOf -LifecycleOverride $lc).Status
            [pscustomobject][ordered]@{
                Option        = $info.Text
                OsString      = $osString
                Version       = $info.Version
                EndOfSupport  = $lc.EndOfSupport
                ExtendedEnd   = $lc.ExtendedEnd
                MainstreamEnd = ConvertTo-LifecycleDateText (Get-CatalogValue $rel 'MainstreamEnd')
                DateSource    = $srcText
                Status        = $status
                Lifecycle     = $lc
                SortKey       = $info.Sort
            }
        }
        $opts = @($opts | Sort-Object @{ Expression = { $_.SortKey }; Descending = $true }, Option)
        if ($opts.Count -eq 0) {
            Write-Warning "Exact-version chooser: every '$($vc.Family)' version is hidden by HideEndedBefore - those VMs keep the RVTools label and get no dropdown."
            continue
        }
        $maxOptions = 25
        if ($opts.Count -gt $maxOptions) {
            Write-Warning ("Exact-version chooser: '{0}' has {1} versions - only the {2} newest are listed (set HideEndedBefore to choose which)." -f $vc.Family, $opts.Count, $maxOptions)
            $opts = @($opts | Select-Object -First $maxOptions)
        }
        $out.Add([pscustomobject][ordered]@{ Family = $vc.Family; Match = $vc.Match; Help = $vc.Help; Product = $vc.Product; Source = $srcText; Options = $opts })
    }
    return @($out.ToArray())
}

function Get-VersionChangeNote {
    # Plain-language note for one scenario row: what was applied and what changed against the RVTools-label result.
    param([Parameter(Mandatory)] $Base, [Parameter(Mandatory)] $Row, [Parameter(Mandatory)] $Option, [Parameter(Mandatory)] [hashtable]$Settings, [datetime]$AsOf = (Get-Date))
    $end = ConvertTo-StrategyDate $Option.EndOfSupport
    $support = if ($null -eq $end) { 'support end date not published' }
               elseif ($end -lt $AsOf) { "support ended $(Format-StrategyDateLong $Option.EndOfSupport)" }
               else { "support ends $(Format-StrategyDateLong $Option.EndOfSupport)" }
    if ($Option.ExtendedEnd) { $support += " (paid extended support to $(Format-StrategyDateLong $Option.ExtendedEnd))" }
    $why = "$($Option.Option) $support"
    if ($Row.OS_Status -eq 'Supported - ends before cut-off') { $why += ", before the cut-off $(Format-StrategyDateLong $Settings.SupportCutoffDate)" }
    if ($Base.Migration_Strategy -ne $Row.Migration_Strategy) {
        return "Strategy changes from $($Base.Migration_Strategy) ($($Base.Rule_Applied)) to $($Row.Migration_Strategy) ($($Row.Rule_Applied)) because $why."
    }
    $changes = @()
    if ($Base.Rule_Applied -ne $Row.Rule_Applied) { $changes += "rule $($Base.Rule_Applied) -> $($Row.Rule_Applied)" }
    if ($Base.Readiness -ne $Row.Readiness)       { $changes += "readiness $($Base.Readiness) -> $($Row.Readiness)" }
    if ($Base.OS_Status -ne $Row.OS_Status)       { $changes += "OS status '$($Base.OS_Status)' -> '$($Row.OS_Status)'" }
    if ($changes.Count -gt 0) { return "Applied: $why. Strategy stays $($Row.Migration_Strategy); " + ($changes -join ', ') + '.' }
    return "Applied: $why. Strategy unchanged ($($Row.Migration_Strategy), $($Row.Rule_Applied))."
}

function Get-VersionScenarioRows {
    <#
        Pre-computes the result of every exact-version option, so the workbook can switch live:
        one row per VM x OS field x option, for VMs whose OS in that field (no override) matches a VersionChoices family
        that has options in -VersionCatalog. The rules run with OS = the option's OsString and the option's support dates
        (OS_Source = 'Client choice (workbook)'). Rows have the Get-MigrationStrategyRows columns plus ScenarioKey
        ('<VM_Index>|<VMware Tools|Configuration file>|<Option>'), VM_Index (1-based position of the VM in -StrategyInventory,
        = its row on the Strategy_* sheets minus 1 - VM names can repeat or contain wildcard characters, the index cannot),
        Scenario_Field, Scenario_Option and the plain-language Version_Note.
    #>
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array]$StrategyInventory,
        [Parameter(Mandatory)] [hashtable]$Settings,
        [AllowEmptyCollection()] [array]$VersionCatalog = @(),
        [switch]$PreferRelocateForLegacyOS,
        [datetime]$AsOf = (Get-Date)
    )
    $byFamily = @{}; foreach ($fam in @($VersionCatalog)) { if ($null -ne $fam -and @($fam.Options).Count -gt 0) { $byFamily[$fam.Family] = $fam } }
    if ($byFamily.Count -eq 0 -or $StrategyInventory.Count -eq 0) { return @() }
    $fieldLabel = @{ VMwareTools = 'VMware Tools'; ConfigFile = 'Configuration file' }
    $all = New-Object System.Collections.Generic.List[object]
    foreach ($field in @('VMwareTools', 'ConfigFile')) {
        $fs = Get-SettingsForOSField -Settings $Settings -OSField $field
        $eligible = New-Object System.Collections.Generic.List[object]   # original VMs
        $copies   = New-Object System.Collections.Generic.List[object]
        $meta     = New-Object System.Collections.Generic.List[object]   # @{ BaseIndex; Option; VmIndex }
        for ($vi = 0; $vi -lt $StrategyInventory.Count; $vi++) {
            $vm = $StrategyInventory[$vi]
            $osUsed = ''; $osSource = ''
            foreach ($src in $fs.OsSourcePrecedence) {
                $t = switch ($src) { 'Override' { $vm.OSOverride } 'VMwareTools' { $vm.OSTools } 'ConfigFile' { $vm.OSConfig } default { '' } }
                if (-not [string]::IsNullOrWhiteSpace($t)) { $osUsed = $t; $osSource = $src; break }
            }
            if ($osSource -eq 'Override' -or [string]::IsNullOrWhiteSpace($osUsed)) { continue }
            $vc = Get-VersionChoiceFamily -OSString $osUsed -Settings $Settings
            if ($null -eq $vc -or -not $byFamily.ContainsKey($vc.Family)) { continue }
            $baseIndex = $eligible.Count
            $eligible.Add($vm)
            # keep the VM's own bitness: a 32-bit 'Windows 10' stays 32-bit whatever release is picked
            $is32 = Test-AnyPattern $osUsed $Settings.Arch32BitPatterns
            $bits = if ($is32) { '(32-bit)' } else { '(64-bit)' }
            foreach ($opt in @($byFamily[$vc.Family].Options)) {
                $os = [string]$opt.OsString
                if ($os.Contains('{1}')) { $os = $os.Replace('{1}', $bits) }
                elseif ($is32 -and -not (Test-AnyPattern $os $Settings.Arch32BitPatterns)) {
                    $os = if ($os -match '\(64-bit\)') { $os -replace '\(64-bit\)', '(32-bit)' } else { "$os (32-bit)" }
                }
                $copy = $vm.PSObject.Copy()
                $copy | Add-Member -NotePropertyName VersionChoice -NotePropertyValue ([pscustomobject]@{ OsString = $os; Lifecycle = $opt.Lifecycle; Option = $opt.Option }) -Force
                $copies.Add($copy)
                $meta.Add(@{ BaseIndex = $baseIndex; Option = $opt; VmIndex = ($vi + 1) })
            }
        }
        if ($copies.Count -eq 0) { continue }
        $baseRows = @(Get-MigrationStrategyRows -StrategyInventory $eligible.ToArray() -Settings $Settings -OSField $field -PreferRelocateForLegacyOS:$PreferRelocateForLegacyOS -AsOf $AsOf)
        $rows     = @(Get-MigrationStrategyRows -StrategyInventory $copies.ToArray() -Settings $Settings -OSField $field -PreferRelocateForLegacyOS:$PreferRelocateForLegacyOS -AsOf $AsOf)
        for ($i = 0; $i -lt $rows.Count; $i++) {
            $row = $rows[$i]; $m = $meta[$i]; $base = $baseRows[$m.BaseIndex]
            $o = [ordered]@{
                ScenarioKey     = "$($m.VmIndex)|$($fieldLabel[$field])|$($m.Option.Option)"
                VM_Index        = [int]$m.VmIndex
                Scenario_Field  = $fieldLabel[$field]
                Scenario_Option = [string]$m.Option.Option
            }
            foreach ($p in $row.PSObject.Properties) { $o[$p.Name] = $p.Value }
            $o['Version_Note'] = Get-VersionChangeNote -Base $base -Row $row -Option $m.Option -Settings $Settings -AsOf $AsOf
            $all.Add([pscustomobject]$o)
        }
    }
    return @($all.ToArray())
}

function Import-WorkbookChoices {
    <#
        Reads the client's choices from a previous output workbook (sheet Migration_Strategy), so a re-run keeps them:
        returns @{ '<vm name, exact text>' = @{ OSField = <OS_Field_To_Use>; ExactVersion = <Exact_OS_Version> } } (keys are
        case-sensitive). A name that appears more than once in the previous workbook is skipped with a warning - the tool
        cannot tell which row was which VM. Values are validated by the exporter (an option that no longer exists is dropped
        with a warning).
    #>
    param([Parameter(Mandatory)] [string]$Path)
    $out = New-Object System.Collections.Hashtable ([System.StringComparer]::Ordinal)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Previous workbook not found: $Path" }
    $rows = @()
    try { $rows = @(Import-Excel -Path $Path -WorksheetName 'Migration_Strategy') }
    catch { Write-Warning "Could not read sheet Migration_Strategy from '$Path' - previous choices are not reused: $($_.Exception.Message)"; return $out }
    $seen = New-Object System.Collections.Hashtable ([System.StringComparer]::Ordinal)
    foreach ($r in $rows) {
        $p = $r.PSObject.Properties['VMName']
        if ($null -eq $p -or $null -eq $p.Value) { continue }
        $name = [string]$p.Value
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ($seen.ContainsKey($name)) { $seen[$name]++ } else { $seen[$name] = 1 }
        $field = [string](Get-RowValue $r 'OS_Field_To_Use'); $exact = [string](Get-RowValue $r 'Exact_OS_Version')
        $out[$name] = @{ OSField = $field.Trim(); ExactVersion = $exact.Trim() }
    }
    $dupes = @($seen.Keys | Where-Object { $seen[$_] -gt 1 })
    foreach ($d in $dupes) { [void]$out.Remove($d) }
    if ($dupes.Count -gt 0) {
        Write-Warning ("{0} VM name(s) appear more than once in '{1}' - their previous choices are not reused: {2}" -f $dupes.Count, (Split-Path $Path -Leaf), (($dupes | Select-Object -First 10) -join ', '))
    }
    Write-Host ("Reused choices from {0}: {1} VM row(s)" -f (Split-Path $Path -Leaf), $out.Count) -ForegroundColor Green
    return $out
}

#endregion

#region ------------------------- Parameter_Map (generated) -------------------------

function Get-ParameterMapRows {
    <#
        Not written to the workbook any more (kept for other tooling).
        Rebuilds the client Parameter_Map for this export: one row per Parameter_Map entry with the RVTools
        input, the condition (built from the settings in force), what it means, the strategy, the manager's columns
        (Highlighted Comments, Decision support, <site> count, Client decision, Client comments) and the counts:
          <SiteName>             - same basis as the manager's Parameter_Map: VMs decided by the row's rule for the rows in
                                   ParameterMapSiteCountDecidedRows, VMs meeting the condition for the others
          Rule                   - the rule (R01-R13) that implements the row
          VMs_Meeting_Condition  - VMs (templates excluded; #1 counts templates) that meet the condition
          VMs_Decided_By_Rule    - VMs whose strategy was set by this row's rule (decision rows)
          Migrating_VMs_Affected - Rehost / Replatform / Relocate VMs that carry this readiness item (readiness rows);
                                   row 32: VMs decided by R06 whose OS is past ALL support (they Relocate when the switch is ON)
        Counts are for one OS field (-OSField, default DefaultOSField).
    #>
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array]$VInfoData,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array]$StrategyInventory,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array]$StrategyRows,
        [Parameter(Mandatory)] [hashtable]$Settings,
        [ValidateSet('', 'VMwareTools', 'ConfigFile')] [string]$OSField = '',
        [string]$SourceLabel = '',
        # Header of the per-site count column (manager's layout: DE-HUB / US-HUB / VVC-USFY / DE-DMZ)
        [string]$SiteName = '',
        [datetime]$AsOf = (Get-Date)
    )

    $Settings = Get-SettingsForOSField -Settings $Settings -OSField $OSField
    # internal register settings are optional in the settings file
    if (-not $Settings.ContainsKey('ParameterMapDecisionSupport'))        { $Settings['ParameterMapDecisionSupport'] = @{} }
    if (-not $Settings.ContainsKey('ParameterMapDecisionSupportDefault')) { $Settings['ParameterMapDecisionSupportDefault'] = 'Yes/No' }
    if (-not $Settings.ContainsKey('ParameterMapSiteCountDecidedRows'))   { $Settings['ParameterMapSiteCountDecidedRows'] = @() }
    $facts = @{}
    foreach ($vm in $StrategyInventory) { $facts[$vm.VMName] = [PSCustomObject]@{ Vm = $vm; F = (Get-StrategyFacts -Vm $vm -Settings $Settings -AsOf $AsOf) } }
    $byName = @{}; foreach ($r in $StrategyRows) { $byName[$r.VMName] = $r }
    $migrating = @('Rehost', 'Replatform', 'Relocate')

    function Measure-FactWhere([scriptblock]$Test) { @($facts.Values | Where-Object $Test).Count }
    function Measure-RuleCount([string[]]$Rules, [string]$PmLike = '*')  { @($StrategyRows | Where-Object { $_.Rule_Applied -in $Rules -and $_.Parameter_Map_Ref -like $PmLike }).Count }
    function Measure-MigratingWhere([scriptblock]$Test) { @($facts.Values | Where-Object { ($byName[$_.Vm.VMName].Migration_Strategy -in $migrating) -and (& $Test $_) }).Count }

    $templateCount = @($VInfoData | Where-Object { Test-RVToolsTrue (Get-RowValue $_ 'Template') }).Count
    $vendorNames = ($Settings.VendorProducts | ForEach-Object { $_.Product }) -join ', '
    $dbNames     = ($Settings.DatabaseProducts | ForEach-Object { "$($_.Engine) ~ $($_.Match)" }) -join '; '
    $legacyStrategy = if ($Settings.LegacyOsStrategy -eq 'Rehost') { 'Rehost (then upgrade)' } else { 'Replatform (Relocate when -PreferRelocateForLegacyOS is ON)' }

    # Row 2 texts follow PoweredOffAction
    if ($Settings.PoweredOffAction -eq 'ClientReview') {
        $pwrMeans = 'Flag as powered off and review the strategy with the client: Retire only after owner confirmation and a final backup, otherwise the strategy the other rules give'
        $pwrStrategy = "$($Settings.PoweredOffStrategy) (Retire or migrate after owner confirmation)"
    } else { $pwrMeans = 'Review_Flag only - the strategy comes from the other rules'; $pwrStrategy = '(flag only)' }

    # ---------------- Highlighted Comments (manager's column), built from this export's data ----------------
    $comments = @{}
    $comments['1'] = "Templates ($templateCount, vInfo Template = True) are excluded: they are master images used to create VMs, not running servers, so nothing runs on them to migrate; any still needed are rebuilt as AWS AMIs."
    $offRows = @($StrategyRows | Where-Object { $_.Rule_Applied -eq 'R13' })
    $comments['2'] = if ($Settings.PoweredOffAction -eq 'ClientReview') { "Not decommissioned automatically: RVTools shows the VM is off but not since when or why (standby, DR or seasonal use is possible). Each of the $($offRows.Count) VM(s) is flagged 'Powered off' with Migration_Strategy = '$($Settings.PoweredOffStrategy)' and Readiness = 'Awaiting client input'; the strategy it gets if still needed is shown in Rationale / Alternative_Strategy. Retire only after owner confirmation and a final backup." } else { 'Flag only: the strategy comes from the other rules; confirm with the owner whether the VM is still needed.' }
    $comments['11'] = 'Part of R06: a 32-bit OS cannot be imported and cannot be upgraded in place to 64-bit, so it is rebuilt on a 64-bit AMI (Replatform).'
    $comments['13'] = 'Part of R04: the client''s appliance tag is used when no known product name matches (confidence Low - the owner names the product).'
    $comments['DB'] = 'New row (AWS version of Azure R07): database servers found by name get the strategy in setting DatabaseStrategy (default Manual Review: reviewed before choosing Amazon RDS / Aurora / DocumentDB or EC2). Confidence Low - the engine is inferred from the name; the client confirms engine, version and features.'
    $r06 = @($StrategyRows | Where-Object { $_.Rule_Applied -eq 'R06' })
    $r06Past = @($facts.Values | Where-Object { $_.F.IsPastAllSupport -and $byName[$_.Vm.VMName].Rule_Applied -eq 'R06' })
    $pastBreakdown = (@($r06Past | Group-Object { if ($_.F.Lifecycle.Label) { $_.F.Lifecycle.Label } else { $_.F.OSUsed } } | Sort-Object Count -Descending | ForEach-Object { "$($_.Count) x $($_.Name)" }) -join ', ')
    $comments['32'] = ("Decision (Confidence: High when the OS comes from VMware Tools; Medium when from the configuration file or a grouped label such as '2016 or later'). " +
        "Of the $(Measure-FactWhere { $_.F.IsLegacyOS }) VM(s) meeting the condition, $(Measure-RuleCount 'R06' 'PM #18*') are decided by R06; $($r06Past.Count) are past all support" +
        $(if ($pastBreakdown) { " ($pastBreakdown)" } else { '' }) +
        " and would go to Relocate with the switch ON. Relocate needs a VMware licence in AWS (Broadcom) - confirm cost before switching on.")
    $countHeader = if ([string]::IsNullOrWhiteSpace($SiteName)) { 'Site count' } else { $SiteName }

    # #, rule, tab, column, condition, means, strategy, why, meeting, decided, migrating affected
    $defs = @(
        @(1,  'R01', 'vInfo', 'Template', 'True', 'Not a server: excluded; rebuild the image as an AMI if still needed', 'Excluded', 'Templates are golden images, not running workloads.', $templateCount, '', ''),
        @(2,  'R13', 'vInfo', 'Powerstate', ("In: " + ($Settings.PoweredOffStates -join ', ') + "  [PoweredOffAction = $($Settings.PoweredOffAction)]"), $pwrMeans, $pwrStrategy, 'Powered-off VMs may be unused, but RVTools cannot say since when or why (standby, DR or seasonal use is possible).', (Measure-FactWhere { $_.F.IsPoweredOff }), (Measure-RuleCount 'R13'), ''),
        @(11, 'R06', 'vInfo', 'OS according to the VMware Tools / configuration file', ("32-bit OS (" + ($Settings.Arch32BitPatterns -join ', ') + ")"), ("Cannot be imported ($($Settings.Arch32BitNote)) or upgraded in place: rebuild on 64-bit"), 'Replatform', 'Decided on the OS field used; the count includes a 32-bit value in either field.', (Measure-FactWhere { $_.F.Is32BitAnyField }), (Measure-RuleCount 'R06' 'PM #11'), ''),
        @(13, 'R04', 'vInfo', ('Custom attributes ' + (($Settings.ApplianceTags.Keys | Sort-Object) -join ' / ')), ('Value matches ' + (($Settings.ApplianceTags.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name) ~ $($_.Value)" }) -join '; ')), 'Replace: vendor AWS edition or SaaS', 'Replace', 'The client tags appliances explicitly; a known product name (row 14) takes precedence.', (Measure-FactWhere { $_.F.IsApplianceTagged }), (Measure-RuleCount 'R04' 'PM #13'), ''),
        @(14, 'R04', 'vInfo', 'Annotation / OS / VM (name) / custom attributes', ("Vendor products: $vendorNames"), 'Replace: vendor AWS edition or SaaS (Retire when the vendor has ended the product)', 'Replace / Retire', 'Appliances are deployed from the vendor image, not lifted.', (Measure-FactWhere { $null -ne $_.F.Vendor }), (Measure-RuleCount 'R04' 'PM #14'), ''),
        @(15, 'R03', 'vInfo', 'OS (field chosen)', ("Windows desktop OS (" + ($Settings.DesktopOsPatterns -join ', ') + ")"), "Replace: $($Settings.DesktopTarget)", 'Replace', 'Desktop OS licensing on EC2 is restricted (BYOL on Dedicated Hosts only).', (Measure-FactWhere { $_.F.IsDesktop }), (Measure-RuleCount 'R03'), ''),
        @('DB', 'R05', 'vInfo', 'VM (name)', ("Database names: $dbNames"), 'Reviewed first; then Replatform to a managed AWS database (Amazon RDS / Aurora / DocumentDB) or Rehost to EC2', $Settings.DatabaseStrategy, 'A managed database removes OS and database patching; AWS version of Azure R07.', (Measure-FactWhere { $null -ne $_.F.Database }), (Measure-RuleCount 'R05'), ''),
        @(17, 'R09', 'vInfo', 'OS (field chosen)', 'OS on the AWS VM Import list (VmImportSupportedOs)', 'Rehost possible', 'Rehost', 'VMware Tools reports the OS from inside the guest and is more reliable than the stale .vmx setting.', (Measure-FactWhere { $_.F.ImportStatus -eq 'Supported' }), (Measure-RuleCount 'R09'), ''),
        @(18, 'R06', 'vInfo', 'OS (field chosen)', "Support ended or ends before $($Settings.SupportCutoffDate)", 'Rehost, then upgrade or buy extended support / Replatform', $legacyStrategy, 'AWS still imports these; the risk is security, not feasibility.', (Measure-FactWhere { $_.F.IsLegacyOS }), (Measure-RuleCount 'R06' 'PM #18*'), ''),
        @(19, 'R07', 'vInfo', 'OS (field chosen)', ('Not accepted by VM Import: ' + ($Settings.VmImportUnsupportedOs -join ', ')), 'Rehost blocked until the OS is upgraded', 'Rehost', 'Not accepted by AWS VM Import.', (Measure-FactWhere { $_.F.ImportStatus -eq 'Not supported' }), (Measure-RuleCount 'R07'), (Measure-MigratingWhere { param($x) $x.F.ImportStatus -eq 'Not supported' })),
        @(20, 'R08', 'vInfo', 'OS (field chosen)', ('Version not identifiable: ' + ($Settings.UnversionedOsPatterns -join ', ') + '; blank / ' + ($Settings.UnknownOsPatterns -join ', ')), 'Confirm OS version with a guest check', 'Rehost', 'RVTools does not carry the version.', (Measure-FactWhere { $_.F.IsUnversioned }), (Measure-RuleCount 'R08'), (Measure-MigratingWhere { param($x) $x.F.IsUnversioned })),
        @(22, 'R11', 'vDisk', 'Capacity MiB', ">= $($Settings.MaxEbsVolumeGiB) GiB on one disk (>= $($Settings.ManifestImportGiB) GiB: manifest import route); > $($Settings.MaxVolumesPerImport) disks", 'Split the volume or move data to FSx/S3', '(no change)', "AWS VM Import accepts disk images below $($Settings.MaxEbsVolumeGiB) GiB and $($Settings.MaxVolumesPerImport) volumes per task.", (Measure-FactWhere { $_.F.DiskTooLarge }), '', (Measure-MigratingWhere { param($x) $x.F.DiskTooLarge })),
        @(23, 'R11', 'vDisk', 'Raw / Sharing mode / Shared Bus', "Raw ~ $($Settings.RawDiskValuePattern); Sharing mode ~ $($Settings.SharedModePattern); Shared Bus ~ $($Settings.SharedBusPattern)", 'Blocker for image export', '(no change) - Blocked', 'RDM and shared disks cannot be exported.', (Measure-FactWhere { $_.F.HasRawShared }), '', (Measure-MigratingWhere { param($x) $x.F.HasRawShared })),
        @(25, 'R12', 'vInfo', 'Cluster / VM (name)', ('Contains: ' + ($Settings.DmzPatterns -join ', ')), 'Final (DMZ) wave after the inspection VPC is live', '(no change)', 'DMZ servers need the AWS perimeter first.', (Measure-FactWhere { $_.Vm.IsDmz }), '', (Measure-MigratingWhere { param($x) $x.Vm.IsDmz })),
        @(30, 'Note', 'vInfo', 'Cluster rule(s)', "Matches $($Settings.AntiAffinityPattern)", 'Spread across Availability Zones', '(no change)', 'Intent: keep these VMs apart.', (Measure-FactWhere { $_.F.HasAntiAffinity }), '', (Measure-MigratingWhere { param($x) $x.F.HasAntiAffinity })),
        @(32, 'R06', 'vInfo', 'Replatform (switch -PreferRelocateForLegacyOS sends past-all-support VMs to Relocate instead) - OS + OsLifecycle table', "Extended support only, end of life, or support ends before $($Settings.SupportCutoffDate); switch ON: past all support -> Relocate", 'Replatform on a supported OS; switch ON: Relocate to a VMware platform in AWS', $legacyStrategy, 'Rehosting an out-of-support OS carries its security risk into AWS.', (Measure-FactWhere { $_.F.IsLegacyOS }), (Measure-RuleCount 'R06' 'PM #18*'), (Measure-FactWhere { $_.F.IsPastAllSupport -and $byName[$_.Vm.VMName].Rule_Applied -eq 'R06' }))
    )

    $out = foreach ($d in $defs) {
        $key = [string]$d[0]
        $ds  = if ($Settings.ParameterMapDecisionSupport.ContainsKey($key)) { $Settings.ParameterMapDecisionSupport[$key] } else { $Settings.ParameterMapDecisionSupportDefault }
        $row = [ordered]@{
            '#'                                = $d[0]
            'RVTools tab'                      = $d[2]
            'Column (input parameter)'         = $d[3]
            'Condition in the data'            = $d[4]
            'What it means for the strategy'   = $d[5]
            'Strategy it points to'            = $d[6]
            'Why (reason)'                     = $d[7]
            'Highlighted Comments'             = $(if ($comments.ContainsKey($key)) { $comments[$key] } else { '' })
            'Decision support'                 = $ds
        }
        $row[$countHeader]                     = $(if ($key -in @($Settings.ParameterMapSiteCountDecidedRows) -and [string]$d[9] -ne '') { $d[9] } else { $d[8] })
        $row['Client decision']                = ''
        $row['Client comments']                = ''
        $row['Rule']                           = $d[1]
        $row['Source export']                  = $SourceLabel
        $row['OS field used for the counts']   = $(if ($Settings.OSField -eq 'ConfigFile') { 'Configuration file' } else { 'VMware Tools' })
        $row['VMs_Meeting_Condition']          = $d[8]
        $row['VMs_Decided_By_Rule']            = $d[9]
        $row['Migrating_VMs_Affected (row 32: past all support = Relocate when switch ON)'] = $d[10]
        [PSCustomObject]$row
    }
    return @($out)
}

#endregion

#region ------------------------- Rule catalogue (Strategy_Rules sheet; .Settings for other tooling) -------------------------

function Get-MigrationStrategyRuleCatalog {
    <#
        Returns .Rules (Strategy_Rules sheet: a compact, plain-language table in the order the rules run) and
        .Settings (a settings table; no longer written to the workbook). CountRule / CountText tell the exporter how to count the VMs live from
        the Migration_Strategy sheet (they follow the client's OS-field choice); they are not written as columns.
        IMPORTANT: if you add, remove or reorder branches in Get-MigrationStrategyRows, update $defs to match.
    #>
    param(
        [Parameter(Mandatory)] [hashtable]$Settings,
        [switch]$PreferRelocateForLegacyOS,
        [int]$TemplateCount = 0
    )
    $preferRelocate = [bool]$PreferRelocateForLegacyOS -or [bool]$Settings.PreferRelocateForLegacyOS
    $oldOs = if ($preferRelocate) { 'Replatform (Relocate if past all support)' } elseif ($Settings.LegacyOsStrategy -eq 'Rehost') { 'Rehost, then upgrade' } else { 'Replatform' }
    $dbTargets = 'Amazon RDS / Aurora (DocumentDB for MongoDB)'

    # Stage, Rule, Rule name, Strategy, In plain words, What happens next, How sure, Register row, CountRule, CountText
    $defs = @(
        @('Before the rules', 'OS field', 'Which OS field is used', '-', "The OS can be read from two RVTools fields that sometimes disagree. Default: $(if ($Settings.DefaultOSField -eq 'VMwareTools') { 'VMware Tools (what the running server reports)' } else { 'Configuration file' }).", 'You can switch to the other field for all VMs (sheet OS_Choice) or per VM (column OS_Field_To_Use); all results update.', '-', '#17', '', ''),
        @('Before the rules', 'OS version', 'Exact Windows version', '-', 'RVTools shows only "Windows 10", "Windows 11", "Windows Server 2012" (2012 or 2012 R2) or "Windows Server 2016 or later". Each covers versions with different support dates.', 'Check the version inside the VM (winver) and pick it in column Exact_OS_Version (sheet OS_Versions explains how); the dates, status and strategy update. Until then the RVTools label is used.', '-', '-', '', ''),
        @('1. Exclude', 'R01', 'Template', 'Excluded', 'A template is a master copy used to create VMs, not a running server.', 'Not counted or costed; rebuild it as an AWS image (AMI) if still needed.', 'High', '#1', 'static', ''),
        @('2. Decide the strategy (first match wins)', 'R02', 'Platform component', 'Retire', 'Part of the VMware or storage platform (vCenter, 3PAR, SimpliVity, vSAN witness). Not needed once workloads are on AWS.', 'Keep until the platform is switched off; take a final backup.', 'High', '-', 'R02', ''),
        @('2. Decide the strategy (first match wins)', 'R03', 'Windows desktop', 'Replace', 'Windows 7 / 8 / 10 / 11 (not Server). Desktop licences on EC2 are very restricted.', 'Move users to Amazon WorkSpaces; confirm your Microsoft licences.', 'Medium', '#15', 'R03', ''),
        @('2. Decide the strategy (first match wins)', 'R04', 'Vendor product or appliance', 'Replace (Retire if discontinued)', 'A known vendor product, or a VM you tagged as "appliance". Vendors supply their own AWS version.', 'Deploy the vendor''s AWS edition or SaaS; Retire products the vendor has ended.', 'Medium (Low if tag only)', '#14, #13', 'R04', ''),
        @('2. Decide the strategy (first match wins)', 'R05', 'Database server', $Settings.DatabaseStrategy, 'The name suggests a database (SQL, PostgreSQL, MySQL, MariaDB, MongoDB, Oracle).', $(if ($Settings.DatabaseStrategy -eq 'Manual Review') { "Every database is reviewed first; the team then chooses $dbTargets or EC2. Confirm the engine, version and features." } else { "Move the database to $dbTargets; keep it on EC2 if a feature is not supported. Confirm the engine." }), 'Low', 'new (Azure R07)', 'R05', ''),
        @('2. Decide the strategy (first match wins)', 'R06', 'Old or 32-bit OS', $oldOs, "The OS is out of support, ends support before $($Settings.SupportCutoffDate), or is 32-bit. Support dates come from the live lifecycle data (sheet Data_Sources) and follow the exact version you pick.", 'Upgrade or rebuild on a supported 64-bit OS as part of the move.', 'High (Medium if from the configuration file)', '#18, #32, #11', 'R06', ''),
        @('2. Decide the strategy (first match wins)', 'R07', 'OS not accepted by AWS', 'Rehost (blocked)', 'A supported OS that the AWS import tool does not accept (e.g. AlmaLinux, Rocky 8, Debian 8/9).', 'Upgrade the OS to an accepted version first, or rebuild.', 'Medium', '#19', 'R07', ''),
        @('2. Decide the strategy (first match wins)', 'R08', 'OS version not known', 'Rehost (check first)', 'RVTools shows the OS family but not the version (e.g. "Other Linux").', 'Check the version inside the server; re-run with the result.', 'Medium', '#20', 'R08', ''),
        @('2. Decide the strategy (first match wins)', 'R09', 'Supported OS', 'Rehost', 'Nothing above applies and AWS accepts the OS as it is.', 'Copy the server to Amazon EC2 as it is.', $Settings.DefaultRehostConfidence, '#17', 'R09', ''),
        @('2. Decide the strategy (first match wins)', 'R10', 'No rule matched', 'Manual Review', 'Safety net: the rules could not decide.', 'Reviewed by the migration team.', 'Low', '-', 'R10', ''),
        @('3. Readiness checks (never change the strategy)', 'R11', 'Disk limits', '(no change)', "A disk of $([math]::Round($Settings.ManifestImportGiB / 1024)) TiB or more, more than $($Settings.MaxVolumesPerImport) disks, or a raw / shared disk.", "From $([math]::Round($Settings.ManifestImportGiB / 1024)) TiB use the larger-disk import route; from $([math]::Round($Settings.MaxEbsVolumeGiB / 1024)) TiB split the disk or move the data; raw / shared disks block the move.", 'High', '#22, #23', '', 'R11'),
        @('3. Readiness checks (never change the strategy)', 'R12', 'DMZ server', '(no change)', 'The server is in the DMZ (internet-facing zone).', 'Moves in the last wave, after the AWS security perimeter is ready.', 'High', '#25', '', 'R12'),
        @('4. Final check', 'R13', 'Powered off', $Settings.PoweredOffStrategy, 'The VM is switched off. RVTools cannot say since when or why.', 'You confirm: still needed? If yes, the strategy shown in Alternative_Strategy; if not, Retire after a final backup.', 'Low', '#2', 'R13', '')
    )
    $hasVersionChoices = $Settings.ContainsKey('VersionChoices') -and @($Settings.VersionChoices).Count -gt 0
    $rules = foreach ($d in $defs) {
        if ($d[1] -eq 'OS version' -and -not $hasVersionChoices) { continue }
        [PSCustomObject][ordered]@{
            Stage = $d[0]; Rule = $d[1]; 'Rule name' = $d[2]; Strategy = $d[3]; 'In plain words' = $d[4]
            'What happens next' = $d[5]; 'How sure' = $d[6]; 'Register row' = $d[7]
            'VMs' = $(if ($d[8] -eq 'static') { $TemplateCount } else { '' })
            CountRule = $d[8]; CountText = $d[9]
        }
    }

    $settingRows = foreach ($k in ($Settings.Keys | Sort-Object)) {
        $v = $Settings[$k]
        $text = if ($v -is [array]) {
            (@($v | ForEach-Object { if ($_ -is [hashtable]) { (($_.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join '; ') } else { [string]$_ } }) -join ' | ')
        } elseif ($v -is [hashtable]) {
            (($v.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join '; ')
        } else { [string]$v }
        if ($k -eq 'PreferRelocateForLegacyOS') { $text = "$text (effective this run: $preferRelocate)" }
        [PSCustomObject][ordered]@{ Setting = $k; Current_Value = $text; Source = (Split-Path $Settings.SettingsPath -Leaf) }
    }

    return [PSCustomObject]@{
        Rules    = @($rules)
        Settings = @($settingRows)
    }
}

#endregion
