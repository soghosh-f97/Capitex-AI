#---------------------------------------------------------------------
# CPU ARCHITECTURE FILTER
#---------------------------------------------------------------------
# VMware VMs are x86_64: an imported x86 image cannot boot on an ARM (AWS Graviton) instance, so ARM types are
# never recommended - for Linux and Windows alike. The architecture comes from the live pricing file
# ('Physical Processor'); with -UseAwsApis the exact architectures from ec2:DescribeInstanceTypes win.
function Test-X86InstanceSpec {
    param($Spec)
    if ($null -eq $Spec) { return $false }
    $type = [string]$Spec.Type
    $facts = Get-Variable -Name AwsLiveFacts -ValueOnly -ErrorAction SilentlyContinue
    if ($null -ne $facts -and $null -ne $facts.InstanceTypes -and $facts.InstanceTypes.ContainsKey($type)) {
        $archs = @($facts.InstanceTypes[$type].Architectures)
        if ($archs.Count -gt 0) { return ($archs -contains 'x86_64') }
    }
    $archProp = $Spec.PSObject.Properties['Architecture']
    if ($null -ne $archProp -and -not [string]::IsNullOrWhiteSpace([string]$archProp.Value)) { return ([string]$archProp.Value -ne 'arm64') }
    # Older index without the Architecture field: fall back to the instance name (a1, t4g, m7g, c8gn, x2gd, im4gn, ...)
    return -not ($type -match '^a1\.' -or $type -match '^[a-z]+\d+g[a-z]*(-[a-z]+)?\.')
}

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

    # x86 only, for every OS. Before v3.0 ARM was excluded for Windows only, so Linux VMs could be given a
    # Graviton type that cannot boot their imported x86 image.
    $x86Specs = Get-X86InstanceSpecs

    $candidates = $x86Specs | Where-Object {
        $_.Family -eq $preferredFamily -and $_.vCPU -ge $cpuTarget -and $_.RAM -ge $ramTarget
    }

    if (-not $candidates) {
        $candidates = $x86Specs | Where-Object {
            $_.vCPU -ge $cpuTarget -and $_.RAM -ge $ramTarget
        }
    }

    if (-not $candidates) {
        return [PSCustomObject]@{ Type = "CUSTOM-OVERSIZED"; vCPU = $cpuTarget; RAM = $ramTarget; Family = "N/A" }
    }

    return ($candidates | Sort-Object @{Expression={$_.vCPU * $_.RAM}} | Select-Object -First 1)
}

function Get-X86InstanceSpecs {
    # The x86 subset of $InstanceSpecs, worked out once per run (sizing is called for every VM).
    $cache = Get-Variable -Name X86InstanceSpecCache -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    $source = Get-Variable -Name X86InstanceSpecCacheSource -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    if ($null -eq $cache -or -not [object]::ReferenceEquals($source, $InstanceSpecs)) {
        $cache = @($InstanceSpecs | Where-Object { Test-X86InstanceSpec $_ })
        Set-Variable -Name X86InstanceSpecCache -Scope Script -Value $cache
        Set-Variable -Name X86InstanceSpecCacheSource -Scope Script -Value $InstanceSpecs
    }
    return $cache
}

#---------------------------------------------------------------------
# MAX ENI (NIC) LOOKUP + NIC-AWARE SIZING WRAPPER
#---------------------------------------------------------------------
# Exact limits come from AWS (ec2:DescribeInstanceTypes -> NetworkInfo.MaximumNetworkInterfaces) when the main
# script runs with -UseAwsApis. Otherwise this best-effort table by instance size is used (AWS does not publish
# max-ENI counts in the pricing file). Verify one type manually with:
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
    $facts = Get-Variable -Name AwsLiveFacts -ValueOnly -ErrorAction SilentlyContinue
    if ($null -ne $facts -and $null -ne $facts.InstanceTypes -and $facts.InstanceTypes.ContainsKey($InstanceType)) {
        $live = [int]$facts.InstanceTypes[$InstanceType].MaxENI
        if ($live -gt 0) { return $live }
    }
    if ($InstanceType -match '\.([a-z0-9]+)$') {
        $size = $Matches[1]
        if ($MaxEniBySize.ContainsKey($size)) { return $MaxEniBySize[$size] }
    }
    return 3
}

function Get-MaxENISource {
    # 'AWS API' when the limit for this type came from ec2:DescribeInstanceTypes, else 'approximate table'.
    param([string]$InstanceType)
    $facts = Get-Variable -Name AwsLiveFacts -ValueOnly -ErrorAction SilentlyContinue
    if ($null -ne $facts -and $null -ne $facts.InstanceTypes -and $facts.InstanceTypes.ContainsKey($InstanceType) -and [int]$facts.InstanceTypes[$InstanceType].MaxENI -gt 0) { return 'AWS API' }
    return 'approximate table'
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
    $eniSrc = Get-MaxENISource -InstanceType $bestFit.Type

    if ($RequiredNICs -gt $maxEni) {
        # New objects instead of Add-Member on the shared spec objects (adding MaxENI a second time to the same
        # object raises "member already exists" when a later VM also needs an upgrade)
        $upgradeCandidates = Get-X86InstanceSpecs |
            Where-Object {
                $_.Family -eq $bestFit.Family -and $_.vCPU -ge $bestFit.vCPU -and $_.RAM -ge $bestFit.RAM
            } |
            ForEach-Object { [PSCustomObject]@{ Type = $_.Type; vCPU = $_.vCPU; RAM = $_.RAM; Family = $_.Family; MaxENI = (Get-MaxENI -InstanceType $_.Type) } } |
            Where-Object { $_.MaxENI -ge $RequiredNICs } |
            Sort-Object @{Expression={$_.vCPU * $_.RAM}}

        if ($upgradeCandidates) {
            $u = $upgradeCandidates | Select-Object -First 1
            return [PSCustomObject]@{
                Type=$u.Type; vCPU=$u.vCPU; RAM=$u.RAM; Family=$u.Family
                MaxENI=$u.MaxENI
                NIC_Note="Upgraded from $($bestFit.Type) to $($u.Type) to support $RequiredNICs NIC(s) (original max ENI was $maxEni; limits: $(Get-MaxENISource -InstanceType $u.Type))."
            }
        }
        else {
            return [PSCustomObject]@{
                Type=$bestFit.Type; vCPU=$bestFit.vCPU; RAM=$bestFit.RAM; Family=$bestFit.Family
                MaxENI=$maxEni
                NIC_Note="WARNING: VM requires $RequiredNICs NIC(s) but '$($bestFit.Type)' supports max $maxEni ENIs ($eniSrc), and no larger instance found in family '$($bestFit.Family)'. Manual review required."
            }
        }
    }
    else {
        return [PSCustomObject]@{
            Type=$bestFit.Type; vCPU=$bestFit.vCPU; RAM=$bestFit.RAM; Family=$bestFit.Family
            MaxENI=$maxEni
            NIC_Note="OK - $RequiredNICs NIC(s) required, $maxEni ENIs supported ($eniSrc)."
        }
    }
}
