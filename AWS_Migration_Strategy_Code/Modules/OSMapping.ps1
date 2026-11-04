#---------------------------------------------------------------------
# OS MAPPING LOGIC
#---------------------------------------------------------------------
function Get-OSMapping {
    param([string]$OSString)
 
    if ([string]::IsNullOrWhiteSpace($OSString)) {
        return [PSCustomObject]@{ PricingOS="Linux"; RecommendedAMI="Unknown - manual review"; Flag="UNKNOWN"; EndOfLife="Unknown"; Note="OS string empty - verify manually" }
    }
 
    switch -Regex ($OSString) {
        "2008"                                                              { return [PSCustomObject]@{ PricingOS="Windows"; RecommendedAMI="Windows Server 2012 R2 (custom AMI/upgrade required)"; Flag="RED-EOL";    EndOfLife="Jan 14, 2020 (EOL)";       Note="2008/2008 R2 not available as standard AWS AMI. Requires VM Import + custom AMI or OS upgrade before migration." } }
        "2012"                                                              { return [PSCustomObject]@{ PricingOS="Windows"; RecommendedAMI="Windows Server 2012 R2 (deprecated AMI)";               Flag="ORANGE-EOL"; EndOfLife="Oct 10, 2023 (EOL)";       Note="Mainstream support ended Oct 2023. Upgrade recommended." } }
        "2016"                                                              { return [PSCustomObject]@{ PricingOS="Windows"; RecommendedAMI="Windows Server 2016";                                    Flag="YELLOW";      EndOfLife="Jan 12, 2027";              Note="Extended support only. Plan OS upgrade." } }
        "2019"                                                              { return [PSCustomObject]@{ PricingOS="Windows"; RecommendedAMI="Windows Server 2019";                                    Flag="GREEN";       EndOfLife="Jan 9, 2029";               Note="Supported." } }
        "2022"                                                              { return [PSCustomObject]@{ PricingOS="Windows"; RecommendedAMI="Windows Server 2022";                                    Flag="GREEN";       EndOfLife="Oct 14, 2031";              Note="Supported (Current)." } }
        "CentOS\s*6|RHEL\s*6|Red\s*Hat\s*Enterprise\s*Linux\s*6"            { return [PSCustomObject]@{ PricingOS="RHEL";    RecommendedAMI="RHEL 8 (custom AMI/upgrade required)";                   Flag="RED-EOL";     EndOfLife="Nov 30, 2020 (EOL)";        Note="EOL - not available as standard AMI." } }
        "CentOS\s*7|RHEL\s*7|Red\s*Hat\s*Enterprise\s*Linux\s*7"            { return [PSCustomObject]@{ PricingOS="RHEL";    RecommendedAMI="RHEL 8/9";                                                Flag="ORANGE-EOL";  EndOfLife="Jun 30, 2024 (EOL)";        Note="Near/at EOL - upgrade recommended." } }
        "RHEL\s*8|RHEL\s*9|CentOS\s*8|Red\s*Hat\s*Enterprise\s*Linux\s*8|Red\s*Hat\s*Enterprise\s*Linux\s*9" { return [PSCustomObject]@{ PricingOS="RHEL";    RecommendedAMI="RHEL 8/9 (direct match)";                                 Flag="GREEN";       EndOfLife="RHEL 8: May 31, 2029 / RHEL 9: May 31, 2032"; Note="Supported." } }
        "SUSE"                                                              { return [PSCustomObject]@{ PricingOS="SUSE";    RecommendedAMI="SUSE Linux Enterprise 15 (direct match)";                Flag="GREEN";       EndOfLife="Varies by SP - see SUSE lifecycle"; Note="Supported." } }
        "Ubuntu"                                                            { return [PSCustomObject]@{ PricingOS="Linux";   RecommendedAMI="Ubuntu (direct match)";                                   Flag="GREEN";       EndOfLife="Varies by version - see Ubuntu release cycle"; Note="Supported." } }
        "Windows"                                                           { return [PSCustomObject]@{ PricingOS="Windows"; RecommendedAMI="Closest matching Windows Server AMI";                    Flag="YELLOW";      EndOfLife="Unknown - version not parsed"; Note="Version not explicitly parsed - verify manually." } }
        "FortiManager|Fortinet|FortiGate"                                   { return [PSCustomObject]@{ PricingOS="Linux"; RecommendedAMI="Fortinet AWS Marketplace AMI (BYOL or hourly)"; Flag="ORANGE-VENDOR"; EndOfLife="N/A - Vendor appliance"; Note="Detected Fortinet appliance - use official Fortinet AWS Marketplace AMI, not a generic Linux AMI. Verify licensing (BYOL vs hourly)." } }
        default                                                             { return [PSCustomObject]@{ PricingOS="Linux";   RecommendedAMI="Closest matching Linux AMI";                             Flag="YELLOW";      EndOfLife="Unknown - manual review"; Note="OS not explicitly recognized - verify manually." } }
    }
}
