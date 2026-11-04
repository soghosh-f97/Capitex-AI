#---------------------------------------------------------------------
# BUILD EC2 MAPPING - processes all VMs from vInfo into recommendations
#---------------------------------------------------------------------
function Build-EC2Mapping {
    param(
        [array]$VInfoData,
        $USIndex,
        $DEIndex
    )
 
    $results = @()
    $counter = 0
    $total = $VInfoData.Count
 
    foreach ($vm in $VInfoData) {
        $counter++
        Write-Progress -Activity "Mapping VMs to AWS" -Status "$($vm.VM)" -PercentComplete (($counter / $total) * 100)
 
        # --- Skip Templates entirely - not included in output or calculations ---
        if ($vm.Template -eq $true) {
            continue
        }
 
        $statusNote = ""
        if ($vm.Powerstate -ne "poweredOn") {
            $statusNote = "Powered Off - validate if still in use before migrating"
        }
 
        $vmCluster = $vm.Cluster
        $vmHost    = $vm.Host
 
        $osConfig = $vm.'OS according to the configuration file'
        $osTools  = $vm.'OS according to the VMware Tools'
        $osMap = Get-OSMapping -OSString $osConfig
 
        $usedToolsFallback = $false
        if ($osMap.Flag -eq "YELLOW" -and -not [string]::IsNullOrWhiteSpace($osTools)) {
            $toolsMap = Get-OSMapping -OSString $osTools
            if ($toolsMap.Flag -ne "YELLOW" -and $toolsMap.Flag -ne "UNKNOWN") {
                $osMap = $toolsMap
                $usedToolsFallback = $true
            }
        }
 
        $osNote = $osMap.Note
        $hasMismatch = ($osConfig -ne $osTools) -and (-not [string]::IsNullOrWhiteSpace($osTools))
        if ($hasMismatch) {
            $osNote += " | OS mismatch between config file and VMware Tools - verify manually."
        }
 
        $osEndOfLife = $osMap.EndOfLife
 
        # --- OS_Confidence determination ---
        $osConfidence = if ($usedToolsFallback) {
            "Medium"
        } elseif ($osMap.Flag -eq "YELLOW" -or $osMap.Flag -eq "UNKNOWN") {
            "Low"
        } elseif ($hasMismatch) {
            "Low"
        } else {
            "High"
        }
 
        # --- Allocated (current) CPU/RAM from vInfo ---
        $allocatedCPU   = [double]$vm.CPUs
        $allocatedRAM   = [double]$vm.Memory / 1024
 
        $nicCount = 1
        if (-not [string]::IsNullOrWhiteSpace([string]$vm.NICs)) {
            $parsedNic = 0
            if ([int]::TryParse([string]$vm.NICs, [ref]$parsedNic) -and $parsedNic -gt 0) {
                $nicCount = $parsedNic
            }
        }
 
        # --- Parse utilization columns (P95 drives sizing, Average is display-only) ---
        $cpuP95Raw = $vm.'CPU Utilization % - P95'
        $cpuAvgRaw = $vm.'CPU Utilization % - Average'
        $memP95Raw = $vm.'Memory Utilization % - P95'
        $memAvgRaw = $vm.'Memory Utilization % - Average'
 
        $cpuP95Val = $null
        if (-not [string]::IsNullOrWhiteSpace([string]$cpuP95Raw)) {
            $parsedVal = 0.0
            if ([double]::TryParse([string]$cpuP95Raw, [ref]$parsedVal)) { $cpuP95Val = $parsedVal }
        }
 
        $memP95Val = $null
        if (-not [string]::IsNullOrWhiteSpace([string]$memP95Raw)) {
            $parsedVal = 0.0
            if ([double]::TryParse([string]$memP95Raw, [ref]$parsedVal)) { $memP95Val = $parsedVal }
        }
 
        $cpuAvgVal = $null
        if (-not [string]::IsNullOrWhiteSpace([string]$cpuAvgRaw)) {
            $parsedVal = 0.0
            if ([double]::TryParse([string]$cpuAvgRaw, [ref]$parsedVal)) { $cpuAvgVal = $parsedVal }
        }
 
        $memAvgVal = $null
        if (-not [string]::IsNullOrWhiteSpace([string]$memAvgRaw)) {
            $parsedVal = 0.0
            if ([double]::TryParse([string]$memAvgRaw, [ref]$parsedVal)) { $memAvgVal = $parsedVal }
        }
 
        # --- Determine sizing basis: Utilization-based (preferred) vs Allocated-capacity fallback ---
        $sizingNote = ""
        if ($null -ne $cpuP95Val -and $null -ne $memP95Val) {
            $requiredCPU    = $allocatedCPU * ($cpuP95Val / 100)
            $requiredRAM_GiB = $allocatedRAM * ($memP95Val / 100)
        }
        else {
            $requiredCPU    = $allocatedCPU
            $requiredRAM_GiB = $allocatedRAM
            $sizingNote = "Utilization data unavailable - sized on allocated capacity"
        }
 
        # --- Best-fit instance now based on utilization-adjusted requirement (10% buffer auto-applied inside InstanceSizing.ps1) ---
        $bestFit = Get-BestFitInstanceWithNIC -RequiredCPU $requiredCPU -RequiredRAM_GiB $requiredRAM_GiB -PricingOS $osMap.PricingOS -RequiredNICs $nicCount
 
        $usHourly = Get-EC2Price -Index $USIndex -InstanceType $bestFit.Type -OperatingSystem $osMap.PricingOS -PriceType "OnDemand"
        $deHourly = Get-EC2Price -Index $DEIndex -InstanceType $bestFit.Type -OperatingSystem $osMap.PricingOS -PriceType "OnDemand"
        
        $usRIHourly = Get-EC2Price -Index $USIndex -InstanceType $bestFit.Type -OperatingSystem $osMap.PricingOS -PriceType "Reserved"
        $deRIHourly = Get-EC2Price -Index $DEIndex -InstanceType $bestFit.Type -OperatingSystem $osMap.PricingOS -PriceType "Reserved"
 
        $deAvailabilityNote = ""
        if ($deHourly -eq 0 -and $usHourly -gt 0) {
            $deAvailabilityNote = "Instance type '$($bestFit.Type)' may not be available/priced in eu-central-1 for OS '$($osMap.PricingOS)' - verify manually."
        }
 
        $usMonthlyUSD = [math]::Round($usHourly * 730, 2)
        $deMonthlyUSD = [math]::Round($deHourly * 730, 2)
 
        $usRIMonthlyUSD = [math]::Round($usRIHourly * 730, 2)
        $deRIMonthlyUSD = [math]::Round($deRIHourly * 730, 2)
 
        $usRISavingsPct = if ($usMonthlyUSD -gt 0 -and $usRIMonthlyUSD -gt 0) { [math]::Round((1 - ($usRIMonthlyUSD / $usMonthlyUSD)) * 100, 1) } else { 0 }
        $deRISavingsPct = if ($deMonthlyUSD -gt 0 -and $deRIMonthlyUSD -gt 0) { [math]::Round((1 - ($deRIMonthlyUSD / $deMonthlyUSD)) * 100, 1) } else { 0 }
 
        $usRIRecommendation = if ($usRIMonthlyUSD -eq 0) { "No RI price available" }
                              elseif ($usRISavingsPct -lt 0) { "On-Demand cheaper - RI not recommended" }
                              else { "RI recommended" }
 
        $deRIRecommendation = if ($deRIMonthlyUSD -eq 0) { "No RI price available" }
                              elseif ($deRISavingsPct -lt 0) { "On-Demand cheaper - RI not recommended" }
                              else { "RI recommended" }
 
        $results += [PSCustomObject]@{
            VMName                        = $vm.VM
            VM_ID                         = $vm.'VM ID'
            Cluster                       = $vmCluster
            Host                          = $vmHost
            Powerstate                    = $vm.Powerstate
            VM_Status_Note                = $statusNote
            Current_vCPU                  = $allocatedCPU
            Current_RAM_GB                = [math]::Round($allocatedRAM,2)
            Current_NIC_Count             = $nicCount
            CPU_Util_Pct_P95              = if ($null -ne $cpuP95Val) { [math]::Round($cpuP95Val,1) } else { "N/A" }
            CPU_Util_Pct_Avg              = if ($null -ne $cpuAvgVal) { [math]::Round($cpuAvgVal,1) } else { "N/A" }
            Mem_Util_Pct_P95              = if ($null -ne $memP95Val) { [math]::Round($memP95Val,1) } else { "N/A" }
            Mem_Util_Pct_Avg              = if ($null -ne $memAvgVal) { [math]::Round($memAvgVal,1) } else { "N/A" }
            Sizing_Note                   = $sizingNote
            OS_ConfigFile                 = $osConfig
            OS_VMwareTools                = $osTools
            OS_Flag                       = $osMap.Flag
            OS_EndOfSupport               = $osEndOfLife
            Recommended_AMI               = $osMap.RecommendedAMI
            OS_Note                       = $osNote
            OS_Confidence                 = $osConfidence
            Recommended_Instance          = $bestFit.Type
            Recommended_Instance_Family   = $bestFit.Family
            Recommended_Instance_vCPU     = $bestFit.vCPU
            Recommended_Instance_RAM_GiB  = $bestFit.RAM
            Max_ENI_Supported             = $bestFit.MaxENI
            NIC_Note                      = $bestFit.NIC_Note
            US_Region                     = "us-east-1"
            US_Monthly_USD                = $usMonthlyUSD
            US_3yr_RI_Monthly_USD         = $usRIMonthlyUSD
            US_3yr_RI_Savings_Pct         = $usRISavingsPct
            US_3yr_RI_Recommendation      = $usRIRecommendation
            DE_Region                     = "eu-central-1"
            DE_Monthly_USD                = $deMonthlyUSD
            DE_3yr_RI_Monthly_USD         = $deRIMonthlyUSD
            DE_3yr_RI_Savings_Pct         = $deRISavingsPct
            DE_3yr_RI_Recommendation      = $deRIRecommendation
            DE_Availability_Note          = $deAvailabilityNote
        }
    }
    Write-Progress -Activity "Mapping VMs to AWS" -Completed
 
    return $results
}
