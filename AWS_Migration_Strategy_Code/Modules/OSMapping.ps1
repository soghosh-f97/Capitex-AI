#---------------------------------------------------------------------
# OS MAPPING LOGIC
#   PricingOS / AMI / flag per OS string. Support dates are NOT typed here any more (v3.0): they come from the
#   OsLifecycle table of the strategy settings, which the run refreshes from endoflife.date - one source of
#   truth for VM_Inventory and Migration_Strategy.
#---------------------------------------------------------------------
function Format-OSMappingDate {
    param([string]$Date)
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact($Date, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) {
        return $d.ToString('MMM d, yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return $Date
}

function Get-OSMappingEndOfLife {
    <#
        End-of-support text for VM_Inventory, from the (live-updated) OsLifecycle table in $StrategySettings
        (first matching row, same matching as the strategy rules). Examples:
          'Oct 14, 2031'   'Jan 14, 2020 (EOL)'   'Oct 10, 2023 (EOL; paid extended support to Oct 13, 2026)'
    #>
    param([string]$OSString)
    $settings = Get-Variable -Name StrategySettings -ValueOnly -ErrorAction SilentlyContinue
    if ($null -eq $settings -or -not ($settings -is [hashtable]) -or -not $settings.ContainsKey('OsLifecycle')) { return 'Unknown - manual review' }
    $asOf = Get-Variable -Name StrategyAsOf -ValueOnly -ErrorAction SilentlyContinue
    if (-not ($asOf -is [datetime])) { $asOf = Get-Date }

    $row = $null
    foreach ($r in @($settings.OsLifecycle)) { if ($OSString -match $r.Match) { $row = $r; break } }
    if ($null -eq $row) { return 'Unknown - version not in the OS lifecycle table' }

    $endText = [string]$row.EndOfSupport
    $extText = [string]$row.ExtendedEnd
    $end = [datetime]::MinValue; $ext = [datetime]::MinValue
    $inv = [System.Globalization.CultureInfo]::InvariantCulture; $none = [System.Globalization.DateTimeStyles]::None
    $hasEnd = [datetime]::TryParseExact($endText, 'yyyy-MM-dd', $inv, $none, [ref]$end)
    $hasExt = [datetime]::TryParseExact($extText, 'yyyy-MM-dd', $inv, $none, [ref]$ext)
    if (-not $hasEnd) {
        $n = [string]$row.Note
        if ($n) { return ('Unknown - {0}' -f $n) }
        return 'Unknown - manual review'
    }
    if ($end -ge $asOf) { return (Format-OSMappingDate $endText) }
    if ($hasExt -and $ext -ge $asOf) { return ('{0} (EOL; paid extended support to {1})' -f (Format-OSMappingDate $endText), (Format-OSMappingDate $extText)) }
    return ('{0} (EOL)' -f (Format-OSMappingDate $endText))
}

function Get-OSMappingAmiNote {
    # With -UseAwsApis: whether AWS offers a standard Windows Server AMI for this version in the region.
    param([string]$OSString)
    $facts = Get-Variable -Name AwsLiveFacts -ValueOnly -ErrorAction SilentlyContinue
    if ($null -eq $facts -or @($facts.WindowsAmis).Count -eq 0) { return '' }
    $ver = $null
    if ($OSString -match 'Windows Server\D*(2016 or later)') { return '' }
    elseif ($OSString -match 'Windows Server\D*(2008|2012)\s*R2') { $ver = $Matches[1] + ' R2' }
    elseif ($OSString -match 'Windows Server\D*(2008|2012|2016|2019|2022|2025)') { $ver = $Matches[1] }
    if (-not $ver) { return '' }
    if (Get-Command Test-WindowsAmiOffered -ErrorAction SilentlyContinue) {
        $offered = Test-WindowsAmiOffered -AmiNames @($facts.WindowsAmis) -Version $ver
    }
    else {
        # fallback without LiveDataProviders.ps1: AWS names 2012 R2 'Windows_Server-2012-R2_...' and 2012 'Windows_Server-2012-RTM-...'
        $pat = switch ($ver) { '2012 R2' { 'Windows_Server-2012-R2_' } '2012' { 'Windows_Server-2012-RTM-' } '2008 R2' { 'Windows_Server-2008-R2_' } '2008' { 'Windows_Server-2008-SP2-' } default { "Windows_Server-$ver-" } }
        $offered = @($facts.WindowsAmis | Where-Object { [string]$_ -match [regex]::Escape($pat) }).Count -gt 0
    }
    if ($offered) { return (' [standard AMI offered in {0}]' -f $facts.Region) }
    return (' [no standard AMI in {0}: custom AMI via VM Import]' -f $facts.Region)
}

function Get-OSMapping {
    param([string]$OSString)

    if ([string]::IsNullOrWhiteSpace($OSString)) {
        return [PSCustomObject]@{ PricingOS="Linux"; RecommendedAMI="Unknown - manual review"; Flag="UNKNOWN"; EndOfLife="Unknown"; Note="OS string empty - verify manually" }
    }

    $m = switch -Regex ($OSString) {
        "2008"                                                              { [PSCustomObject]@{ PricingOS="Windows"; RecommendedAMI="Windows Server 2012 R2 (custom AMI/upgrade required)"; Flag="RED-EOL";    Note="2008/2008 R2 not available as standard AWS AMI. Requires VM Import + custom AMI or OS upgrade before migration." }; break }
        "2012"                                                              { [PSCustomObject]@{ PricingOS="Windows"; RecommendedAMI="Windows Server 2012 R2 (deprecated AMI)";               Flag="ORANGE-EOL"; Note="Support has ended (see OS_EndOfSupport). Upgrade recommended." }; break }
        "2016"                                                              { [PSCustomObject]@{ PricingOS="Windows"; RecommendedAMI="Windows Server 2016";                                    Flag="YELLOW";      Note="Extended support only. Plan OS upgrade." }; break }
        "2019"                                                              { [PSCustomObject]@{ PricingOS="Windows"; RecommendedAMI="Windows Server 2019";                                    Flag="GREEN";       Note="Supported." }; break }
        "2022"                                                              { [PSCustomObject]@{ PricingOS="Windows"; RecommendedAMI="Windows Server 2022";                                    Flag="GREEN";       Note="Supported (Current)." }; break }
        "CentOS\s*6|RHEL\s*6|Red\s*Hat\s*Enterprise\s*Linux\s*6"            { [PSCustomObject]@{ PricingOS="RHEL";    RecommendedAMI="RHEL 8 (custom AMI/upgrade required)";                   Flag="RED-EOL";     Note="EOL - not available as standard AMI." }; break }
        "CentOS\s*7|RHEL\s*7|Red\s*Hat\s*Enterprise\s*Linux\s*7"            { [PSCustomObject]@{ PricingOS="RHEL";    RecommendedAMI="RHEL 8/9";                                                Flag="ORANGE-EOL";  Note="Near/at EOL - upgrade recommended." }; break }
        "RHEL\s*8|RHEL\s*9|CentOS\s*8|Red\s*Hat\s*Enterprise\s*Linux\s*8|Red\s*Hat\s*Enterprise\s*Linux\s*9" { [PSCustomObject]@{ PricingOS="RHEL";    RecommendedAMI="RHEL 8/9 (direct match)";                                 Flag="GREEN";       Note="Supported." }; break }
        "SUSE"                                                              { [PSCustomObject]@{ PricingOS="SUSE";    RecommendedAMI="SUSE Linux Enterprise 15 (direct match)";                Flag="GREEN";       Note="Supported." }; break }
        "Ubuntu"                                                            { [PSCustomObject]@{ PricingOS="Linux";   RecommendedAMI="Ubuntu (direct match)";                                   Flag="GREEN";       Note="Supported." }; break }
        "Windows"                                                           { [PSCustomObject]@{ PricingOS="Windows"; RecommendedAMI="Closest matching Windows Server AMI";                    Flag="YELLOW";      Note="Version not explicitly parsed - verify manually." }; break }
        "FortiManager|Fortinet|FortiGate"                                   { [PSCustomObject]@{ PricingOS="Linux"; RecommendedAMI="Fortinet AWS Marketplace AMI (BYOL or hourly)"; Flag="ORANGE-VENDOR"; Note="Detected Fortinet appliance - use official Fortinet AWS Marketplace AMI, not a generic Linux AMI. Verify licensing (BYOL vs hourly)." }; break }
        default                                                             { [PSCustomObject]@{ PricingOS="Linux";   RecommendedAMI="Closest matching Linux AMI";                             Flag="YELLOW";      Note="OS not explicitly recognized - verify manually." } }
    }

    $eol = if ($m.Flag -eq 'ORANGE-VENDOR') { 'N/A - Vendor appliance' } else { Get-OSMappingEndOfLife -OSString $OSString }
    $ami = [string]$m.RecommendedAMI
    if ($m.PricingOS -eq 'Windows') { $ami += (Get-OSMappingAmiNote -OSString $OSString) }

    return [PSCustomObject]@{ PricingOS = $m.PricingOS; RecommendedAMI = $ami; Flag = $m.Flag; EndOfLife = $eol; Note = $m.Note }
}
