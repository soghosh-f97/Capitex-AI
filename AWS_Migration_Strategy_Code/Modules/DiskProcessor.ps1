#---------------------------------------------------------------------
# BUILD EBS MAPPING - processes one aggregated row PER VM (not per disk)
# Allocated_GiB = MAX(Total disk capacity MiB, Provisioned MiB, In Use MiB) / 1024
#---------------------------------------------------------------------
function Build-EBSMapping {
    param(
        [array]$VInfoData,
        [array]$TemplateVMNames,
        $USIndex,
        $DEIndex
    )

    $results = @()
    $usGp3Price = $USIndex.EBSGp3Price
    $deGp3Price = $DEIndex.EBSGp3Price

    foreach ($vm in $VInfoData) {

        # --- Skip Templates entirely - not included in output or calculations ---
        if ($TemplateVMNames -contains $vm.VM) {
            continue
        }

        # --- Safely parse the three candidate MiB values (fields may be blank/missing) ---
        $totalCapMiB    = 0.0
        $provisionedMiB = 0.0
        $inUseMiB       = 0.0

        $raw = $vm.'Total disk capacity MiB'
        if (-not [string]::IsNullOrWhiteSpace([string]$raw)) {
            $parsed = 0.0
            if ([double]::TryParse([string]$raw, [ref]$parsed)) { $totalCapMiB = $parsed }
        }

        $raw = $vm.'Provisioned MiB'
        if (-not [string]::IsNullOrWhiteSpace([string]$raw)) {
            $parsed = 0.0
            if ([double]::TryParse([string]$raw, [ref]$parsed)) { $provisionedMiB = $parsed }
        }

        $raw = $vm.'In Use MiB'
        if (-not [string]::IsNullOrWhiteSpace([string]$raw)) {
            $parsed = 0.0
            if ([double]::TryParse([string]$raw, [ref]$parsed)) { $inUseMiB = $parsed }
        }

        $maxMiB  = ($totalCapMiB, $provisionedMiB, $inUseMiB | Measure-Object -Maximum).Maximum
        $sizeGiB = [math]::Round($maxMiB / 1024, 2)

        # --- Note which source field drove the sizing (useful for manual review) ---
        $basis = if ($maxMiB -eq $totalCapMiB -and $maxMiB -gt 0) { "Total disk capacity" }
                 elseif ($maxMiB -eq $provisionedMiB -and $maxMiB -gt 0) { "Provisioned" }
                 elseif ($maxMiB -eq $inUseMiB -and $maxMiB -gt 0) { "In Use" }
                 else { "No disk data available" }

        $usMonthly    = [math]::Round($sizeGiB * $usGp3Price, 2)
        $deMonthlyUSD = [math]::Round($sizeGiB * $deGp3Price, 2)

        $results += [PSCustomObject]@{
            VM_Name            = $vm.VM
            VM_ID              = $vm.'VM ID'
            Allocated_GiB      = $sizeGiB
            Sizing_Basis       = $basis
            Recommended_Volume = "gp3 (VM-aggregated - verify thin/thick & IOPS needs per volume manually)"
            US_Monthly_USD     = $usMonthly
            DE_Monthly_USD     = $deMonthlyUSD
        }
    }

    return $results
}
