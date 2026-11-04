#---------------------------------------------------------------------
# SIZING FUNCTION (closest-fit with buffer)
#---------------------------------------------------------------------
function Get-BestFitInstance {
    param(
        [double]$RequiredCPU,
        [double]$RequiredRAM_GiB,
        [string]$PricingOS = "Linux"
    )

    $cpuTarget = [math]::Ceiling($RequiredCPU * $Buffer)
    $ramTarget = [math]::Ceiling($RequiredRAM_GiB * $Buffer)

    $ratio = if ($RequiredCPU -gt 0) { $RequiredRAM_GiB / $RequiredCPU } else { 4 }

    $preferredFamily = switch ($true) {
        { $RequiredCPU -le 2 -and $RequiredRAM_GiB -le 8 } { "Burstable"; break }
        { $ratio -le 3 }   { "ComputeOptimized"; break }
        { $ratio -ge 7 }   { "MemoryOptimized"; break }
        default            { "GeneralPurpose" }
    }

    $isWindowsPricing = $PricingOS -eq "Windows"

    $candidates = $InstanceSpecs | Where-Object {
        $_.Family -eq $preferredFamily -and $_.vCPU -ge $cpuTarget -and $_.RAM -ge $ramTarget -and
        (-not $isWindowsPricing -or $_.Type -notmatch '^[a-z]\d+g[a-z]*\.')
    }

    if (-not $candidates) {
        $candidates = $InstanceSpecs | Where-Object {
            $_.vCPU -ge $cpuTarget -and $_.RAM -ge $ramTarget -and
            (-not $isWindowsPricing -or $_.Type -notmatch '^[a-z]\d+g[a-z]*\.')
        }
    }

    if (-not $candidates) {
        return [PSCustomObject]@{ Type = "CUSTOM-OVERSIZED"; vCPU = $cpuTarget; RAM = $ramTarget; Family = "N/A" }
    }

    return ($candidates | Sort-Object @{Expression={$_.vCPU * $_.RAM}} | Select-Object -First 1)
}

#---------------------------------------------------------------------
# MAX ENI (NIC) LOOKUP + NIC-AWARE SIZING WRAPPER
#---------------------------------------------------------------------
# NOTE: AWS does not publish max-ENI counts in the pricing CSV.
# This is a best-effort approximation based on AWS's published limits.
# For critical/production sizing, verify exact limits via:
#   aws ec2 describe-instance-types --instance-types <type> `
#     --query "InstanceTypes[].NetworkInfo.MaximumNetworkInterfaces"
$MaxEniBySize = @{
    "nano"=2; "micro"=2; "small"=3; "medium"=3; "large"=3
    "xlarge"=4; "2xlarge"=4; "4xlarge"=8; "8xlarge"=8; "9xlarge"=8
    "10xlarge"=8; "12xlarge"=8; "16xlarge"=15; "18xlarge"=15
    "24xlarge"=15; "32xlarge"=15; "48xlarge"=15; "metal"=15
}

function Get-MaxENI {
    param([string]$InstanceType)
    if ($InstanceType -match '\.([a-z0-9]+)$') {
        $size = $Matches[1]
        if ($MaxEniBySize.ContainsKey($size)) { return $MaxEniBySize[$size] }
    }
    return 3
}

function Get-BestFitInstanceWithNIC {
    param(
        [double]$RequiredCPU,
        [double]$RequiredRAM_GiB,
        [string]$PricingOS = "Linux",
        [int]$RequiredNICs = 1
    )

    $bestFit = Get-BestFitInstance -RequiredCPU $RequiredCPU -RequiredRAM_GiB $RequiredRAM_GiB -PricingOS $PricingOS

    if ($bestFit.Type -eq "CUSTOM-OVERSIZED") {
        return [PSCustomObject]@{
            Type=$bestFit.Type; vCPU=$bestFit.vCPU; RAM=$bestFit.RAM; Family=$bestFit.Family
            MaxENI=0; NIC_Note="Custom oversized instance - verify NIC/ENI support manually."
        }
    }

    $maxEni = Get-MaxENI -InstanceType $bestFit.Type
    $isWindowsPricing = $PricingOS -eq "Windows"

    if ($RequiredNICs -gt $maxEni) {
        $upgradeCandidates = $InstanceSpecs |
            Where-Object {
                $_.Family -eq $bestFit.Family -and $_.vCPU -ge $bestFit.vCPU -and $_.RAM -ge $bestFit.RAM -and
                (-not $isWindowsPricing -or $_.Type -notmatch '^[a-z]\d+g[a-z]*\.')
            } |
            ForEach-Object { $_ | Add-Member -NotePropertyName MaxENI -NotePropertyValue (Get-MaxENI -InstanceType $_.Type) -PassThru } |
            Where-Object { $_.MaxENI -ge $RequiredNICs } |
            Sort-Object @{Expression={$_.vCPU * $_.RAM}}

        if ($upgradeCandidates) {
            $u = $upgradeCandidates | Select-Object -First 1
            return [PSCustomObject]@{
                Type=$u.Type; vCPU=$u.vCPU; RAM=$u.RAM; Family=$u.Family
                MaxENI=$u.MaxENI
                NIC_Note="Upgraded from $($bestFit.Type) to $($u.Type) to support $RequiredNICs NIC(s) (original max ENI was $maxEni)."
            }
        }
        else {
            return [PSCustomObject]@{
                Type=$bestFit.Type; vCPU=$bestFit.vCPU; RAM=$bestFit.RAM; Family=$bestFit.Family
                MaxENI=$maxEni
                NIC_Note="WARNING: VM requires $RequiredNICs NIC(s) but '$($bestFit.Type)' supports max $maxEni ENIs, and no larger instance found in family '$($bestFit.Family)'. Manual review required."
            }
        }
    }
    else {
        return [PSCustomObject]@{
            Type=$bestFit.Type; vCPU=$bestFit.vCPU; RAM=$bestFit.RAM; Family=$bestFit.Family
            MaxENI=$maxEni
            NIC_Note="OK - $RequiredNICs NIC(s) required, $maxEni ENIs supported."
        }
    }
}
