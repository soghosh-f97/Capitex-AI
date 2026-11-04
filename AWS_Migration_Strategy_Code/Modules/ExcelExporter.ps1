#---------------------------------------------------------------------
# WRAPPED SHEET - fixed widths, wrapped text, top/left aligned (long rule / condition text)
#---------------------------------------------------------------------
function Write-WrappedSheet {
    param(
        [array]$Rows,
        [string]$OutputPath,
        [string]$WorksheetName,
        [int[]]$Widths,
        [int]$CentreColumns = 0
    )
    $pkg = $Rows | Export-Excel -Path $OutputPath -WorksheetName $WorksheetName -TableStyle Medium9 -FreezeTopRow -BoldTopRow -PassThru
    $ws  = $pkg.Workbook.Worksheets[$WorksheetName]
    $lastRow = $ws.Dimension.End.Row
    $lastCol = $ws.Dimension.End.Column
    for ($c = 1; $c -le [Math]::Min($lastCol, $Widths.Count); $c++) { $ws.Column($c).Width = $Widths[$c - 1] }
    $all = $ws.Cells[1, 1, $lastRow, $lastCol]
    $all.Style.WrapText            = $true
    $all.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
    $all.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
    if ($CentreColumns -gt 0) {
        $ws.Cells[1, 1, $lastRow, $CentreColumns].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    }
    Close-ExcelPackage $pkg
}

#---------------------------------------------------------------------
# EXPORT TO EXCEL - writes the sizing tabs and, when supplied, the migration strategy tabs
#---------------------------------------------------------------------
function Export-MigrationReport {
    param(
        [array]$EC2Results,
        [array]$EBSResults,
        [array]$Summary,
        [string]$OutputPath,
        # Optional migration strategy outputs (Modules\MigrationStrategyGenerator.ps1)
        [array]$StrategyRows = @(),
        # Optional: the strategy rows for each OS field (keys VMwareTools, ConfigFile). When both are given the
        # workbook gets the OS_Choice switch and Migration_Strategy follows the client's choice live.
        [System.Collections.IDictionary]$StrategyRowsByField = $null,
        [string]$DefaultOSField = 'VMwareTools',
        # Optional: Get-MigrationStrategyRuleCatalog output; only .Rules is written (sheet Strategy_Rules)
        $RuleCatalog = $null
    )

    Write-Host "Writing output to $OutputPath ..." -ForegroundColor Cyan

    $EC2Export = $EC2Results | Select-Object `
        VMName, Powerstate, OS_ConfigFile, OS_VMwareTools, OS_Note, OS_Flag, OS_EndOfSupport, OS_Confidence, `
        Recommended_AMI, Cluster, Host, Current_vCPU, Current_RAM_GB, Current_NIC_Count, `
        CPU_Util_Pct_P95, CPU_Util_Pct_Avg, Mem_Util_Pct_P95, Mem_Util_Pct_Avg, Sizing_Note, `
        Max_ENI_Supported, NIC_Note, `
        Recommended_Instance, Recommended_Instance_Family, Recommended_Instance_vCPU, Recommended_Instance_RAM_GiB, `
        US_Region, US_Monthly_USD, US_3yr_RI_Monthly_USD, US_3yr_RI_Savings_Pct, US_3yr_RI_Recommendation, `
        DE_Region, DE_Monthly_USD, DE_3yr_RI_Monthly_USD, DE_3yr_RI_Savings_Pct, DE_3yr_RI_Recommendation, `
        DE_Availability_Note

    # --- Disk_Inventory now reflects one aggregated row PER VM (Option A) ---
    $EBSExport = $EBSResults | Select-Object `
        VM_Name, VM_ID, Allocated_GiB, Sizing_Basis, `
        Recommended_Volume, US_Monthly_USD, DE_Monthly_USD

    # --- Legend data explaining OS_Flag, OS_Note, NIC_Note, Recommended_Volume, and Disk_Inventory values ---
    $Legend = @(
        [PSCustomObject]@{ Column = "OS_Flag"; 'Value / Pattern' = "GREEN"; Meaning = "OS is supported/current. No OS lifecycle action is required." }
        [PSCustomObject]@{ Column = "OS_Flag"; 'Value / Pattern' = "YELLOW"; Meaning = "OS is in extended support, not explicitly recognized, or its version could not be parsed. Review and plan remediation where applicable." }
        [PSCustomObject]@{ Column = "OS_Flag"; 'Value / Pattern' = "ORANGE-EOL"; Meaning = "OS is near or at end-of-life. Upgrade is recommended before migration." }
        [PSCustomObject]@{ Column = "OS_Flag"; 'Value / Pattern' = "RED-EOL"; Meaning = "OS is end-of-life and is not available as a standard AWS AMI. Upgrade the OS or use VM Import with a custom AMI." }
        [PSCustomObject]@{ Column = "OS_Flag"; 'Value / Pattern' = "ORANGE-VENDOR"; Meaning = "Vendor appliance detected. Use the vendor's official AWS Marketplace image rather than a generic OS AMI." }
        [PSCustomObject]@{ Column = "OS_Flag"; 'Value / Pattern' = "UNKNOWN"; Meaning = "Source OS value is blank or unavailable. Manual review is required." }

        [PSCustomObject]@{ Column = "OS_Note"; 'Value / Pattern' = "Supported."; Meaning = "The detected OS release is supported." }
        [PSCustomObject]@{ Column = "OS_Note"; 'Value / Pattern' = "Supported (Current)."; Meaning = "The detected OS release is current and supported." }
        [PSCustomObject]@{ Column = "OS_Note"; 'Value / Pattern' = "Extended support only. Plan OS upgrade."; Meaning = "The OS is no longer in mainstream support; plan an upgrade." }
        [PSCustomObject]@{ Column = "OS_Note"; 'Value / Pattern' = "Mainstream support ended <date>. Upgrade recommended."; Meaning = "Mainstream support has ended on the stated date; upgrade is recommended." }
        [PSCustomObject]@{ Column = "OS_Note"; 'Value / Pattern' = "Near/at EOL - upgrade recommended."; Meaning = "The OS is approaching or has reached end-of-life; upgrade before migration where possible." }
        [PSCustomObject]@{ Column = "OS_Note"; 'Value / Pattern' = "OS not explicitly recognized - verify manually."; Meaning = "The OS description could not be matched confidently to a supported lifecycle record." }
        [PSCustomObject]@{ Column = "OS_Note"; 'Value / Pattern' = "Version not explicitly parsed - verify manually."; Meaning = "The OS family was identified, but the exact version could not be parsed." }
        [PSCustomObject]@{ Column = "OS_Note"; 'Value / Pattern' = "2008/2008 R2 not available as standard AWS AMI. ..."; Meaning = "Windows Server 2008/2008 R2 requires VM Import with a custom AMI or an OS upgrade before migration." }
        [PSCustomObject]@{ Column = "OS_Note"; 'Value / Pattern' = "<base note> | OS mismatch between config file and VMware Tools - verify manually."; Meaning = "The lifecycle note before the separator still applies, and the two source OS descriptions disagree; validate the actual guest OS." }

        [PSCustomObject]@{ Column = "OS_EndOfSupport"; 'Value / Pattern' = "<date or lifecycle text>"; Meaning = "End-of-life/end-of-support date or review status for the detected OS. Use it to prioritize upgrades." }

        [PSCustomObject]@{ Column = "OS_Confidence"; 'Value / Pattern' = "High"; Meaning = "OS was identified directly and confidently from the VM configuration file, with no mismatch versus VMware Tools." }
        [PSCustomObject]@{ Column = "OS_Confidence"; 'Value / Pattern' = "Medium"; Meaning = "The configuration file did not provide a clear/confident match, so VMware Tools was used as a fallback to identify the OS." }
        [PSCustomObject]@{ Column = "OS_Confidence"; 'Value / Pattern' = "Low"; Meaning = "Either the OS version could not be clearly identified (family/version unclear), or the configuration file and VMware Tools values disagree - verify manually." }

        [PSCustomObject]@{ Column = "CPU_Util_Pct_P95"; 'Value / Pattern' = "<percent> or N/A"; Meaning = "P95 CPU utilization percent sourced from vInfo; this is the primary driver for right-sizing vCPU. 'N/A' means this data was not available for the VM." }
        [PSCustomObject]@{ Column = "CPU_Util_Pct_Avg"; 'Value / Pattern' = "<percent> or N/A"; Meaning = "Average CPU utilization percent; secondary/sanity-check reference only - does not drive sizing." }
        [PSCustomObject]@{ Column = "Mem_Util_Pct_P95"; 'Value / Pattern' = "<percent> or N/A"; Meaning = "P95 memory utilization percent sourced from vInfo; this is the primary driver for right-sizing RAM. 'N/A' means this data was not available for the VM." }
        [PSCustomObject]@{ Column = "Mem_Util_Pct_Avg"; 'Value / Pattern' = "<percent> or N/A"; Meaning = "Average memory utilization percent; secondary/sanity-check reference only - does not drive sizing." }
        [PSCustomObject]@{ Column = "Sizing_Note"; 'Value / Pattern' = "Utilization data unavailable - sized on allocated capacity"; Meaning = "P95 CPU/Memory utilization was missing for this VM, so sizing fell back to the VM's currently allocated vCPU/RAM (plus buffer) instead of utilization-based sizing." }

        [PSCustomObject]@{ Column = "US_3yr_RI_Recommendation"; 'Value / Pattern' = "RI recommended"; Meaning = "A priced 3-year Reserved Instance option is available and is recommended based on the report's pricing comparison." }
        [PSCustomObject]@{ Column = "US_3yr_RI_Recommendation"; 'Value / Pattern' = "No RI price available"; Meaning = "No matching 3-year Reserved Instance price was available; review pricing or use the on-demand estimate." }

        [PSCustomObject]@{ Column = "Allocated_GiB (Disk_Inventory)"; 'Value / Pattern' = "<number>"; Meaning = "Per-VM EBS sizing requirement = MAX(Total disk capacity MiB, Provisioned MiB, In Use MiB) / 1024, sourced from vInfo. One row per VM (not per disk)." }
        [PSCustomObject]@{ Column = "Sizing_Basis (Disk_Inventory)"; 'Value / Pattern' = "Total disk capacity / Provisioned / In Use / No disk data available"; Meaning = "Indicates which of the three source fields produced the maximum value, and therefore drove the Allocated_GiB figure." }
        [PSCustomObject]@{ Column = "Migration_Strategy (Migration_Strategy)"; 'Value / Pattern' = "Retire / Replace / Replatform / Relocate / Rehost / Manual Review / Client Review"; Meaning = "One primary strategy per VM. Rules R02-R10 are checked in order and the first that matches decides (see Strategy_Rules). Refactor is never assigned automatically. Client Review = the VM is powered off (rule R13): it is never decommissioned automatically and the strategy is agreed with the client." }
        [PSCustomObject]@{ Column = "Client Review rows (Migration_Strategy)"; 'Value / Pattern' = "Rule_Applied = R13"; Meaning = "Rationale and Alternative_Strategy show the strategy the rules give if the VM is still needed; AWS_Target and Pre_Migration_Actions are for that case. If the owner confirms it is not needed: Retire after a final backup." }
        [PSCustomObject]@{ Column = "OS_Field_To_Use (Migration_Strategy)"; 'Value / Pattern' = "Default / VMware Tools / Configuration file"; Meaning = "Which RVTools OS field decides this VM's strategy. Default = follow the setting on sheet OS_Choice. Pick a value from the dropdown and every result column of the row updates. VMware Tools (what the running server reports) is recommended; the configuration file is set when the VM is created and is often out of date." }
        [PSCustomObject]@{ Column = "OS_Field_Applied (Migration_Strategy)"; 'Value / Pattern' = "VMware Tools / Configuration file"; Meaning = "The OS field actually used for the row. If a guest-check override file was used, OS_Source shows 'Override' and both choices give the same result." }
        [PSCustomObject]@{ Column = "Pre_Migration_Actions (Migration_Strategy)"; 'Value / Pattern' = "BLOCKER: ... ; R11 Disk limits: ... ; R12 DMZ server: ..."; Meaning = "Readiness checks R11-R12 start with their rule number. BLOCKER items must be fixed before the VM can move (Readiness = Blocked)." }
        [PSCustomObject]@{ Column = "Readiness_Checks (Migration_Strategy)"; 'Value / Pattern' = "R11, R12 or blank"; Meaning = "Which readiness checks apply to the VM: R11 Disk limits, R12 DMZ server. They never change the strategy; details are in Pre_Migration_Actions." }
        [PSCustomObject]@{ Column = "VMs (Strategy_Rules)"; 'Value / Pattern' = "<number>"; Meaning = "R01: templates excluded. R02-R10 and R13: VMs whose strategy the rule decided. R11-R12: VMs that carry the check. Counted live from Migration_Strategy, so they follow the OS field choice." }
        [PSCustomObject]@{ Column = "Strategy_Confidence (Migration_Strategy)"; 'Value / Pattern' = "High / Medium / Low"; Meaning = "Strength of the RVTools evidence. Low = the decision depends on a fact RVTools does not contain (see Client_Confirmation)." }
        [PSCustomObject]@{ Column = "Rule_Applied (Migration_Strategy)"; 'Value / Pattern' = "R02-R10, R13"; Meaning = "The rule that decided the strategy (numbered in the order the rules run, see Strategy_Rules)." }
        [PSCustomObject]@{ Column = "Readiness (Migration_Strategy)"; 'Value / Pattern' = "Ready / Ready after prep / Awaiting client input / Blocked / Not migrating"; Meaning = "Blocked = a blocker must be fixed before export (e.g. raw/shared disk, OS not accepted by VM Import). Not migrating = Retire / Retain." }
        [PSCustomObject]@{ Column = "OS_Status (Migration_Strategy)"; 'Value / Pattern' = "Supported / Supported - ends before cut-off / Extended support only (ESU/ELS/LTSS) / Unsupported - End of life / Unsupported - 32-bit OS / Verify - ..."; Meaning = "From the OS lifecycle table and the OS used (OS_Source: Override, VMwareTools or ConfigFile)." }
        [PSCustomObject]@{ Column = "Recommended_Volume (Disk_Inventory)"; 'Value / Pattern' = "gp3 (VM-aggregated - verify thin/thick & IOPS needs per volume manually)"; Meaning = "Sizing is now a per-VM aggregate, so individual disk thin/thick-provisioning status is no longer available; a generic gp3 recommendation is given and per-volume detail must be verified manually if needed." }
    )

    $EC2Export | Export-Excel -Path $OutputPath -WorksheetName "VM_Inventory"    -AutoSize -FreezeTopRow -TableStyle Medium2 -Append
    $EBSExport | Export-Excel -Path $OutputPath -WorksheetName "Disk_Inventory"  -AutoSize -FreezeTopRow -TableStyle Medium6 -Append
    $Summary   | Export-Excel -Path $OutputPath -WorksheetName "Summary_Metrics" -AutoSize -FreezeTopRow -TableStyle Medium9 -Append
    $Legend    | Export-Excel -Path $OutputPath -WorksheetName "Legend"          -AutoSize -FreezeTopRow -TableStyle Medium4 -Append

    # --- Migration strategy sheets (same formatting approach as the Azure tool) ---
    # Internal columns kept on the in-memory rows (Get-ParameterMapRows uses Parameter_Map_Ref) but never written to the workbook
    $excludedColumns = @('Parameter_Map_Ref')
    $fieldLabel = @{ VMwareTools = 'VMware Tools'; ConfigFile = 'Configuration file' }
    $haveBoth   = ($null -ne $StrategyRowsByField) -and $StrategyRowsByField.Contains('VMwareTools') -and $StrategyRowsByField.Contains('ConfigFile')
    if ($StrategyRows.Count -gt 0) {
        # Rows stay in inventory order (sorted by VM name), so row N is the same VM on every strategy sheet
        $visible = foreach ($r in $StrategyRows) {
            $o = [ordered]@{}
            foreach ($prop in $r.PSObject.Properties) {
                if ($prop.Name -in $excludedColumns) { continue }
                $o[$prop.Name] = $prop.Value
                if ($haveBoth -and $prop.Name -eq 'OS_VMwareTools') { $o['OS_Field_To_Use'] = 'Default'; $o['OS_Field_Applied'] = $fieldLabel[$DefaultOSField] }
            }
            [PSCustomObject]$o
        }
        $visible | Export-Excel -Path $OutputPath -WorksheetName "Migration_Strategy" -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium9 -Append
        if ($haveBoth) {
            foreach ($field in @('VMwareTools', 'ConfigFile')) {
                $StrategyRowsByField[$field] | Select-Object -Property * -ExcludeProperty $excludedColumns |
                    Export-Excel -Path $OutputPath -WorksheetName "Strategy_$field" -FreezeTopRow -BoldTopRow -Append
            }
        }
    }
    if ($null -ne $RuleCatalog) {
        # Compact rule table: Stage, Rule, Rule name, Strategy, In plain words, What happens next, How sure, VMs
        # ('Register row' refers to the internal Parameter_Map register, which is not in the client workbook)
        $ruleSheet = @($RuleCatalog.Rules | Select-Object * -ExcludeProperty CountRule, CountText, 'Register row')
        Write-WrappedSheet -Rows $ruleSheet -OutputPath $OutputPath -WorksheetName "Strategy_Rules" -Widths @(24, 9, 24, 22, 60, 60, 18, 8) -CentreColumns 0
        # $RuleCatalog.Settings (the Strategy_Settings sheet) is no longer written: the settings values are not shown to the client
    }

    if ($StrategyRows.Count -gt 0) {
        Set-StrategyWorkbookFormulas -OutputPath $OutputPath -HaveBoth:$haveBoth -DefaultLabel $fieldLabel[$DefaultOSField] -RuleCatalog $RuleCatalog `
            -DifferentStrategyCount $(if ($haveBoth) { $n = 0; for ($i = 0; $i -lt $StrategyRowsByField['VMwareTools'].Count; $i++) { if ($StrategyRowsByField['VMwareTools'][$i].Migration_Strategy -ne $StrategyRowsByField['ConfigFile'][$i].Migration_Strategy) { $n++ } }; $n } else { 0 })
    }

    Write-Host "`n✅ Done! Output saved to: $OutputPath" -ForegroundColor Green
}


#---------------------------------------------------------------------
# LIVE FORMULAS - OS field switch, compact rule counts and summary counts
#---------------------------------------------------------------------
function Set-StrategyWorkbookFormulas {
    <#
        Turns the strategy part of the workbook into a live view:
          - OS_Choice sheet: one dropdown (cell C4) choosing the OS field for all VMs
          - Migration_Strategy: per-VM dropdown OS_Field_To_Use (Default / VMware Tools / Configuration file); every
            result column reads the matching hidden sheet (Strategy_VMwareTools / Strategy_ConfigFile)
          - Strategy_Rules 'VMs' and the Summary_Metrics strategy / readiness counts are COUNTIF formulas on Migration_Strategy
        The workbook is calculated before saving, so the values are visible even in viewers that do not recalculate.
    #>
    param(
        [Parameter(Mandatory)] [string]$OutputPath,
        [switch]$HaveBoth,
        [string]$DefaultLabel = 'VMware Tools',
        $RuleCatalog = $null,
        [int]$DifferentStrategyCount = 0
    )
    $pkg = Open-ExcelPackage -Path $OutputPath
    $wb  = $pkg.Workbook
    $ms  = $wb.Worksheets['Migration_Strategy']
    $last = $ms.Dimension.End.Row
    $col = @{}; for ($c = 1; $c -le $ms.Dimension.End.Column; $c++) { $col[$ms.Cells[1, $c].Text] = $c }
    $L = { param($n) [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($n) }
    $range = { param($name) "Migration_Strategy!`$$(& $L $col[$name])`$2:`$$(& $L $col[$name])`$$last" }

    if ($HaveBoth) {
        # ---------- OS_Choice control sheet ----------
        $oc = $wb.Worksheets.Add('OS_Choice')
        $wb.Worksheets.MoveBefore('OS_Choice', 'Migration_Strategy')
        $oc.Cells['A1'].Value = 'Which OS field should decide the migration strategy?'
        $oc.Cells['A1'].Style.Font.Bold = $true; $oc.Cells['A1'].Style.Font.Size = 14
        $oc.Cells['A2'].Value = 'RVTools reports the operating system in two fields, and they sometimes disagree. Choose which one the strategy should be based on.'
        $oc.Cells['B4'].Value = 'OS field for all VMs:'; $oc.Cells['B4'].Style.Font.Bold = $true
        $oc.Cells['C4'].Value = $DefaultLabel
        $oc.Cells['C4'].Style.Font.Bold = $true
        $oc.Cells['C4'].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $oc.Cells['C4'].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 242, 204))
        $oc.Cells['C4'].Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Medium)
        $dv = $oc.DataValidations.AddListValidation('C4')
        [void]$dv.Formula.Values.Add('VMware Tools'); [void]$dv.Formula.Values.Add('Configuration file')
        $dv.ShowErrorMessage = $true; $dv.ErrorTitle = 'OS field'; $dv.Error = 'Choose VMware Tools or Configuration file.'
        $notes = @(
            @('VMware Tools (recommended)', 'The OS reported by the running server itself. Up to date as long as VMware Tools runs in the VM.'),
            @('Configuration file', 'The OS chosen in vSphere when the VM was created. It is not updated when the OS is upgraded, so it is often out of date.'),
            @('One VM only', 'On sheet Migration_Strategy, column OS_Field_To_Use: pick VMware Tools or Configuration file for that row. Default = follow cell C4.'),
            @('What updates', 'Migration_Strategy (every result column), the VM counts on Strategy_Rules and the strategy / readiness counts on Summary_Metrics.'),
            @('What does not update', "The EC2 cost totals on Summary_Metrics: they are for the OS field chosen when the tool ran ($DefaultLabel)."),
            @('Guest-check override', 'If the tool was run with an OS override file, the override always wins and both choices give the same result.'),
            @('VMs affected', "$DifferentStrategyCount VM(s) in this export get a different strategy depending on the OS field.")
        )
        $r = 6
        foreach ($n in $notes) { $oc.Cells[$r, 2].Value = $n[0]; $oc.Cells[$r, 2].Style.Font.Bold = $true; $oc.Cells[$r, 3].Value = $n[1]; $r++ }
        $oc.Column(1).Width = 3; $oc.Column(2).Width = 30; $oc.Column(3).Width = 110
        $oc.Cells[6, 3, $r, 3].Style.WrapText = $true
        $oc.View.ShowGridLines = $false

        # ---------- Migration_Strategy: per-VM dropdown and live result columns ----------
        $cChoice = $col['OS_Field_To_Use']; $cApplied = $col['OS_Field_Applied']
        $dv2 = $ms.DataValidations.AddListValidation("$(& $L $cChoice)2:$(& $L $cChoice)$last")
        foreach ($v in @('Default', 'VMware Tools', 'Configuration file')) { [void]$dv2.Formula.Values.Add($v) }
        $dv2.ShowErrorMessage = $true; $dv2.ErrorTitle = 'OS field'; $dv2.Error = 'Choose Default, VMware Tools or Configuration file.'
        $ms.Cells[2, $cChoice, $last, $cChoice].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $ms.Cells[2, $cChoice, $last, $cChoice].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 242, 204))

        $hidden = @{}
        foreach ($sheet in @('Strategy_VMwareTools', 'Strategy_ConfigFile')) {
            $ws = $wb.Worksheets[$sheet]
            $h = @{}; for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) { $h[$ws.Cells[1, $c].Text] = $c }
            $hidden[$sheet] = $h
            $ws.Hidden = [OfficeOpenXml.eWorkSheetHidden]::Hidden
        }
        $static = @('VMName', 'Powerstate', 'Cluster', 'Host', 'OS_ConfigFile', 'OS_VMwareTools', 'OS_Field_To_Use', 'OS_Field_Applied',
                    'Current_vCPU', 'Current_RAM_GB', 'Current_Disk_GB', 'Largest_Disk_GiB', 'Current_Disk_Count', 'Current_NIC_Count')
        $dynamic = @($col.Keys | Where-Object { $_ -notin $static -and $hidden['Strategy_VMwareTools'].ContainsKey($_) })
        $ca = & $L $cApplied; $cc = & $L $cChoice
        for ($row = 2; $row -le $last; $row++) {
            $ms.Cells[$row, $cApplied].Formula = "IF(`$$cc$row=""Default"",OS_Choice!`$C`$4,`$$cc$row)"
            foreach ($name in $dynamic) {
                $t = & $L $hidden['Strategy_VMwareTools'][$name]; $cf = & $L $hidden['Strategy_ConfigFile'][$name]
                $ms.Cells[$row, $col[$name]].Formula = "IF(`$$ca$row=""Configuration file"",Strategy_ConfigFile!$cf$row,Strategy_VMwareTools!$t$row)&"""""
            }
        }
    }

    # ---------- Strategy_Rules: live VM counts ----------
    if ($null -ne $RuleCatalog) {
        $sr = $wb.Worksheets['Strategy_Rules']
        $vmCol = 0; for ($c = 1; $c -le $sr.Dimension.End.Column; $c++) { if ($sr.Cells[1, $c].Text -eq 'VMs') { $vmCol = $c } }
        $i = 2
        foreach ($rule in $RuleCatalog.Rules) {
            if ($vmCol -gt 0) {
                if ($rule.CountRule -and $rule.CountRule -ne 'static') { $sr.Cells[$i, $vmCol].Formula = "COUNTIF($(& $range 'Rule_Applied'),""$($rule.CountRule)"")" }
                elseif ($rule.CountText) {
                    # Readiness_Checks holds R11 / R12 in this order, joined with ', ' - count every combination that
                    # contains the check (exact matches: no wildcards, so every spreadsheet engine gives the same result)
                    $all = @('R11', 'R12'); $combos = @()
                    for ($m = 1; $m -lt (1 -shl $all.Count); $m++) { $set = @(for ($k = 0; $k -lt $all.Count; $k++) { if ($m -band (1 -shl $k)) { $all[$k] } }); if ($set -contains $rule.CountText) { $combos += ($set -join ', ') } }
                    $sr.Cells[$i, $vmCol].Formula = (@($combos | ForEach-Object { "COUNTIF($(& $range 'Readiness_Checks'),""$_"")" }) -join '+')
                }
            }
            $i++
        }
        if ($vmCol -gt 0) { $sr.Cells[2, $vmCol, ($i - 1), $vmCol].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center }
    }

    # ---------- Summary_Metrics: live strategy / readiness counts ----------
    # COUNTIF criteria: escape " (doubled) and the wildcards * ? ~ (prefixed with ~) so any label counts exactly
    $crit = { param($t) (([string]$t) -replace '~', '~~' -replace '\*', '~*' -replace '\?', '~?' -replace '"', '""') }
    $sm = $wb.Worksheets['Summary_Metrics']
    for ($row = 2; $row -le $sm.Dimension.End.Row; $row++) {
        $m = $sm.Cells[$row, 1].Text
        if ($m -match '^VMs - Migration strategy: (.+)$') { $sm.Cells[$row, 2].Formula = "COUNTIF($(& $range 'Migration_Strategy'),""$(& $crit $Matches[1])"")" }
        elseif ($m -match '^VMs - Readiness: (.+)$')      { $sm.Cells[$row, 2].Formula = "COUNTIF($(& $range 'Readiness'),""$(& $crit $Matches[1])"")" }
        elseif ($m -eq 'VMs - Client confirmation needed') { $sm.Cells[$row, 2].Formula = "COUNTIF($(& $range 'Client_Confirmation'),""?*"")" }
        elseif ($m -like 'VMs - Powered off*(rule R13)')   { $sm.Cells[$row, 2].Formula = "COUNTIF($(& $range 'Rule_Applied'),""R13"")" }
        elseif ($m -like 'OS field deciding the strategy*' -and $HaveBoth) { $sm.Cells[$row, 2].Formula = 'OS_Choice!$C$4' }
    }

    $wb.CalcMode = [OfficeOpenXml.ExcelCalcMode]::Automatic
    $wb.FullCalcOnLoad = $true
    # Store calculated values too (Excel recalculates on open; this helps viewers that do not)
    try { [OfficeOpenXml.CalculationExtension]::Calculate($wb) } catch { Write-Warning "Formula pre-calculation skipped: $($_.Exception.Message) - Excel calculates on open." }
    Close-ExcelPackage $pkg
}
