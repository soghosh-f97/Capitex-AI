#---------------------------------------------------------------------
# BUILD SUMMARY - aggregates all totals/counts into Summary rows
#---------------------------------------------------------------------
function Build-SummaryReport {
    param(
        [array]$EC2Results,
        [array]$EBSResults,
        [double]$BufferPercent,
        [int]$InstanceSpecCount,
        # Optional: output of Get-MigrationStrategyRows (adds strategy / readiness metrics)
        [array]$StrategyRows = @(),
        # Optional: the rows for every OS field, so every strategy / readiness value that the client's
        # OS-field choice can produce gets a metric row (the exporter turns these counts into live formulas)
        [array]$StrategyRowsAll = @()
    )

    $totalVMs = $EC2Results.Count
    $flaggedOS = (@($EC2Results | Where-Object { $_.OS_Flag -like "*EOL*" })).Count
    $flaggedDEUnavailable = (@($EC2Results | Where-Object { $_.DE_Availability_Note -ne "" })).Count
    $flaggedNICUpgrade = (@($EC2Results | Where-Object { $_.NIC_Note -like "Upgraded*" })).Count
    $flaggedNICWarning = (@($EC2Results | Where-Object { $_.NIC_Note -like "WARNING*" })).Count

    $totalUSMonthly      = [math]::Round(($EC2Results | Where-Object { $_.VM_Status_Note -eq "" } | Measure-Object -Property US_Monthly_USD -Sum).Sum, 2)
    $totalUSRIMonthly    = [math]::Round(($EC2Results | Where-Object { $_.VM_Status_Note -eq "" } | Measure-Object -Property US_3yr_RI_Monthly_USD -Sum).Sum, 2)
    $totalDEMonthlyUSD   = [math]::Round(($EC2Results | Where-Object { $_.VM_Status_Note -eq "" } | Measure-Object -Property DE_Monthly_USD -Sum).Sum, 2)
    $totalDERIMonthlyUSD = [math]::Round(($EC2Results | Where-Object { $_.VM_Status_Note -eq "" } | Measure-Object -Property DE_3yr_RI_Monthly_USD -Sum).Sum, 2)

    $totalEBSUS      = [math]::Round(($EBSResults | Measure-Object -Property US_Monthly_USD -Sum).Sum, 2)
    $totalEBSDEUsd   = [math]::Round(($EBSResults | Measure-Object -Property DE_Monthly_USD -Sum).Sum, 2)

    # --- Migration strategy metrics (only when the strategy step ran) ---
    $strategyMetrics = @()
    if ($StrategyRows.Count -gt 0) {
        $allRows = if ($StrategyRowsAll.Count -gt 0) { $StrategyRowsAll } else { $StrategyRows }
        # the OS field most rows used (a row falls back to the other field only when the chosen one is blank)
        $usedField = @($StrategyRows | Where-Object { $_.OS_Source -in @('VMwareTools', 'ConfigFile') } | Group-Object OS_Source | Sort-Object Count -Descending | Select-Object -First 1 -ExpandProperty Name)
        $strategyMetrics += [PSCustomObject]@{ Metric = "OS field deciding the strategy (change it on sheet OS_Choice)"; Value = $(if ($usedField -eq 'ConfigFile') { 'Configuration file' } else { 'VMware Tools' }) }
        foreach ($name in (@($allRows | ForEach-Object { $_.Migration_Strategy }) | Sort-Object -Unique)) {
            $strategyMetrics += [PSCustomObject]@{ Metric = "VMs - Migration strategy: $name"; Value = (@($StrategyRows | Where-Object { $_.Migration_Strategy -eq $name })).Count }
        }
        foreach ($name in (@($allRows | ForEach-Object { $_.Readiness }) | Sort-Object -Unique)) {
            $strategyMetrics += [PSCustomObject]@{ Metric = "VMs - Readiness: $name"; Value = (@($StrategyRows | Where-Object { $_.Readiness -eq $name })).Count }
        }
        $strategyMetrics += [PSCustomObject]@{ Metric = "VMs - Client confirmation needed"; Value = (@($StrategyRows | Where-Object { $_.Client_Confirmation -ne '' })).Count }
        $strategyMetrics += [PSCustomObject]@{ Metric = "VMs - Powered off, strategy to be reviewed with the client (rule R13)"; Value = (@($StrategyRows | Where-Object { $_.Rule_Applied -eq 'R13' })).Count }

        # EC2 cost of the VMs that actually move to EC2 (Rehost / Replatform); Relocate goes to a VMware platform instead
        $toEc2 = @{}
        foreach ($s in ($StrategyRows | Where-Object { $_.Migration_Strategy -in @('Rehost', 'Replatform') })) { $toEc2[$s.VMName] = $true }
        # same basis as the headline totals: powered-off VMs (VM_Status_Note set) are excluded
        $ec2Moving = @($EC2Results | Where-Object { $toEc2.ContainsKey([string]$_.VMName) -and $_.VM_Status_Note -eq "" })
        $strategyMetrics += [PSCustomObject]@{ Metric = "Est. EC2 Monthly Cost - Rehost/Replatform VMs only - US OnDemand (USD) - OS field chosen when the tool ran"; Value = [math]::Round(($ec2Moving | Measure-Object -Property US_Monthly_USD -Sum).Sum, 2) }
        $strategyMetrics += [PSCustomObject]@{ Metric = "Est. EC2 Monthly Cost - Rehost/Replatform VMs only - DE OnDemand (USD) - OS field chosen when the tool ran"; Value = [math]::Round(($ec2Moving | Measure-Object -Property DE_Monthly_USD -Sum).Sum, 2) }
    }

    return @(
        [PSCustomObject]@{ Metric = "Total VMs";                                          Value = $totalVMs }
        [PSCustomObject]@{ Metric = "VMs Flagged EOL/Unsupported OS";                     Value = $flaggedOS }
        [PSCustomObject]@{ Metric = "VMs Flagged - Instance possibly unavailable in Germany"; Value = $flaggedDEUnavailable }
        [PSCustomObject]@{ Metric = "VMs Flagged - Instance upgraded due to NIC/ENI limit"; Value = $flaggedNICUpgrade }
        [PSCustomObject]@{ Metric = "VMs Flagged - NIC/ENI requirement exceeds largest available instance"; Value = $flaggedNICWarning }
        [PSCustomObject]@{ Metric = "Est. Total EC2 Monthly Cost - US OnDemand (USD)";     Value = $totalUSMonthly }
        [PSCustomObject]@{ Metric = "Est. Total EC2 Monthly Cost - US 3yr-RI (USD)";       Value = $totalUSRIMonthly }
        [PSCustomObject]@{ Metric = "Est. Total EC2 Monthly Cost - DE OnDemand (USD)";     Value = $totalDEMonthlyUSD }
        [PSCustomObject]@{ Metric = "Est. Total EC2 Monthly Cost - DE 3yr-RI (USD)";       Value = $totalDERIMonthlyUSD }
        [PSCustomObject]@{ Metric = "Est. Total EBS Monthly Cost - US (USD)";              Value = $totalEBSUS }
        [PSCustomObject]@{ Metric = "Est. Total EBS Monthly Cost - DE (USD)";              Value = $totalEBSDEUsd }
        [PSCustomObject]@{ Metric = "Buffer applied to sizing";                           Value = "$BufferPercent%" }
        [PSCustomObject]@{ Metric = "Unique instance types in dynamic spec table";        Value = $InstanceSpecCount }
    ) + $strategyMetrics + @(
        [PSCustomObject]@{ Metric = "Report generated on";                                Value = (Get-Date) }
    )
}
