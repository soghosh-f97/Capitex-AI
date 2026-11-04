<#
    AWSMigrationStrategySettings.psd1 - every keyword list, threshold, lookup table and target used by
    MigrationStrategyGenerator.ps1 (AWS). Nothing in the strategy code is a literal list or threshold:
    change behaviour here, not in the code. All patterns are regular expressions, matched case-insensitively
    with PowerShell -match. Loaded with Import-PowerShellDataFile (data only - no script blocks).

    RULES (numbered in the order they run - version 2.2):
      R01 Template (excluded)                                   Parameter_Map #1
      --- Decide the strategy: first matching rule wins ---
      R02 Platform component        -> Retire                   (Azure R01)
      R03 Windows desktop           -> Replace (WorkSpaces)     #15
      R04 Vendor product/appliance  -> Replace / Retire         #14, #13
      R05 Database server           -> Manual Review            new (Azure R07)
      R06 Old or 32-bit OS          -> Replatform               #18, #32, #11
      R07 OS not accepted by VM Import -> Rehost (blocked)      #19
      R08 OS version not known      -> Rehost (check first)     #20
      R09 Supported OS              -> Rehost                   #17
      R10 No rule matched           -> Manual Review
      --- Readiness checks: every check that applies; never change the strategy ---
      R11 Disk limits (size, count, raw / shared)               #22, #23
      R12 DMZ server                                            #25
      --- Final check ---
      R13 Powered off               -> Client Review            #2
    EXACT OS VERSION (3.0): VersionChoices (section 6b) lets the client pick the exact Windows version per VM in the
      workbook (column Exact_OS_Version, sheet OS_Versions); the rules above are then re-evaluated for that version.
    Removed in 2.0: decommission notes, plant floor, domain controllers (rows 6, 7, 8, 12 removed earlier),
    infrastructure services (#16) and ISO connected (#24). Anti-affinity (#30) is a design note.
#>
@{
    SettingsVersion = '3.0'                    # Version of this settings file (for reference only; not used by the rules).

    #---------------------------------------------------------------------------------------------
    # 0. LIVE DATA (v3.0) - values read from live sources at run time; every source has a fallback and the
    #    workbook sheet Data_Sources shows where each value came from (Live / cached copy / offline snapshot / settings file).
    #---------------------------------------------------------------------------------------------
    LiveData = @{
        Enabled              = $true      # $true = read OS support dates and Windows versions from endoflife.date and check the AWS VM Import page; $false = never use the internet (offline snapshot + dates below).
        CacheDays            = 7          # Days a downloaded copy is reused before it is downloaded again (an older copy is still used, with a warning, when the internet is not reachable).
        LifecycleApiBase     = 'https://endoflife.date/api/v1/products'   # OS support dates API (no login). Products used: the 'Live' keys of OsLifecycle and VersionChoices.
        LifecycleLegacyBase  = 'https://endoflife.date/api'               # Older version of the same API, tried only when the first one answers 404 or an unexpected format.
        VmImportDocUrl       = 'https://docs.aws.amazon.com/vm-import/latest/userguide/prerequisites.html'   # AWS VM Import rules page (no API exists): its text is hashed to detect changes.
        VmImportDocKnownHash = ''         # SHA-256 of the VM Import page text when sections 7-8 were last reviewed ('' = not recorded yet; the run prints the current hash on Data_Sources).
        TimeoutSec           = 20         # Seconds to wait for each live source. After the first 'no answer' (timeout, DNS, refused, proxy) the host is skipped for the rest of the run.
    }

    #---------------------------------------------------------------------------------------------
    # 1. OS FIELD (which RVTools OS column decides)
    #---------------------------------------------------------------------------------------------
    DefaultOSField         = 'VMwareTools'     # Starting OS column: 'VMwareTools' (live, recommended) or 'ConfigFile' (.vmx, can be stale). Other column used if blank; override file wins; client can switch in workbook.
    # Review flag: OS name -> OS family, used to judge if a difference between the two OS columns is material. First match wins.
    OsFamilies = @(
        @{ Match = 'Windows Server'; Family = 'Windows Server' }
        @{ Match = 'Windows';        Family = 'Windows client' }
        @{ Match = 'Red Hat|RHEL';   Family = 'RHEL' }
        @{ Match = 'CentOS';         Family = 'CentOS' }
        @{ Match = 'Oracle Linux';   Family = 'Oracle Linux' }
        @{ Match = 'Rocky';          Family = 'Rocky' }
        @{ Match = 'Alma';           Family = 'AlmaLinux' }
        @{ Match = 'openSUSE';       Family = 'openSUSE' }
        @{ Match = 'SUSE|SLES';      Family = 'SUSE' }
        @{ Match = 'Ubuntu';         Family = 'Ubuntu' }
        @{ Match = 'Debian';         Family = 'Debian' }
        @{ Match = 'Photon';         Family = 'Photon' }
        @{ Match = 'ESXi';           Family = 'ESXi' }
        @{ Match = 'Forti';          Family = 'Fortinet' }
        @{ Match = 'FreeBSD';        Family = 'FreeBSD' }
        @{ Match = 'Linux';          Family = 'Other Linux' }
    )
    VagueOsFamilies        = @('Other Linux', 'Unknown')   # Review flag: families too vague to call an OS-column difference material.

    #---------------------------------------------------------------------------------------------
    # 2. R02 PLATFORM COMPONENT -> Retire
    #---------------------------------------------------------------------------------------------
    EnableInfraRule        = $true             # [R02] Platform component -> Retire rule: $true = on, $false = off.
    InfraNamePatterns      = @('^vCLS', '^vcsa', 'vcenter', 'serviceprocessor', '^3parsp', '^ssmc[-_]', '^oneview[-_]', '^omnistackvc', '^vvw', 'TMPL-')   # [R02] VM name patterns of VMware / storage platform components.
    InfraTextPatterns      = @('vCenter Server Appliance', 'vSAN .*Witness', 'VMware ESXi', 'OmniStack', '3PAR', 'SSMC', 'OneView', 'vSphere Cluster Service')   # [R02] Same, searched in VM notes and both OS columns.

    PoweredOffStates       = @('poweredOff', 'suspended')   # [R13] RVTools power states that count as "powered off".
    PoweredOffAction       = 'ClientReview'    # [R13] 'ClientReview' = strategy set to Client Review (default); 'Flag' = strategy kept, review flag only.
    PoweredOffStrategy     = 'Client Review'   # [R13] Strategy text written for powered-off VMs when PoweredOffAction = 'ClientReview'.

    #---------------------------------------------------------------------------------------------
    # 3. R04 VENDOR PRODUCT OR APPLIANCE (Parameter_Map #14, #13) - first match wins
    #    In     = 'All'  -> matched on VM name | Annotation | both OS strings | tag columns
    #           = 'Name' -> matched on the VM name only
    #    Action = 'Retire'  -> the vendor has ended the product
    #           = 'Replace' -> vendor AWS edition / SaaS
    #    A VM that matches no product but carries the client's appliance tag (ApplianceTags) is also R04 (Low).
    #---------------------------------------------------------------------------------------------
    # [R04] Known vendor products/appliances: Retire (product ended) or Replace with the AWS target shown. First match wins.
    VendorProducts = @(
        @{ Match = 'MineMeld';                              In = 'All';  Action = 'Retire';  Product = 'Palo Alto Networks MineMeld';            Target = 'None - product discontinued by the vendor' }
        @{ Match = 'Forti(Gate|Manager|Analyzer)|Fortinet'; In = 'All';  Action = 'Replace'; Product = 'Fortinet FortiGate / FortiManager / FortiAnalyzer'; Target = 'Fortinet image from AWS Marketplace' }
        @{ Match = 'Panorama';                              In = 'All';  Action = 'Replace'; Product = 'Palo Alto Networks Panorama';           Target = 'Palo Alto Panorama from AWS Marketplace' }
        @{ Match = 'BeyondTrust|Bomgar';                    In = 'All';  Action = 'Replace'; Product = 'BeyondTrust (Bomgar) Secure Remote Access'; Target = 'BeyondTrust SaaS or AWS Marketplace image' }
        @{ Match = 'Tenable|Nessus';                        In = 'All';  Action = 'Replace'; Product = 'Tenable Core + Nessus';                 Target = 'Tenable cloud / AWS Marketplace scanner' }
        @{ Match = 'Unified Access Gateway';                In = 'All';  Action = 'Replace'; Product = 'VMware (Omnissa) Unified Access Gateway'; Target = 'Workspace ONE UAG vendor image on AWS' }
        @{ Match = '^airwatch';                             In = 'Name'; Action = 'Replace'; Product = 'Workspace ONE UEM (AirWatch)';          Target = 'Workspace ONE UEM SaaS' }
        @{ Match = 'netsight|Extreme Management';           In = 'All';  Action = 'Replace'; Product = 'Extreme Networks NetSight';             Target = 'ExtremeCloud IQ (SaaS)' }
        @{ Match = 'OpenScape';                             In = 'All';  Action = 'Replace'; Product = 'Unify OpenScape Voice';                 Target = 'Vendor cloud voice service - confirm with vendor' }
        @{ Match = 'Nvidia (GRID )?License';                In = 'All';  Action = 'Replace'; Product = 'NVIDIA License Server';                 Target = 'NVIDIA Cloud License Service' }
        @{ Match = 'Bluecat';                               In = 'All';  Action = 'Replace'; Product = 'BlueCat DDI';                           Target = 'BlueCat on AWS / Amazon Route 53' }
        @{ Match = 'Photon';                                In = 'All';  Action = 'Replace'; Product = 'VMware appliance on Photon OS';         Target = 'VMware appliance: vendor AWS image or SaaS' }
    )
    TagColumnsSearched     = @('machinetype', 'managedby', 'application')   # [R04] vInfo custom-attribute columns also searched for vendor names (missing ones skipped).
    AnnotationPurposeRegex = 'Purpose:\s*(.*?)\s*(POC:|$)'   # Reads the "Purpose:" text from the vSphere notes (currently unused: no rule or sheet uses it).

    ApplianceTags          = @{ machinetype = '^appliance$'; managedby = '^appliance$' }   # [R04] Client tags marking an appliance (column = value); used when no product matches (Low).
    ApplianceTagTarget     = 'Vendor AWS edition or SaaS - owner to identify the product'   # [R04] AWS target text for VMs found only by the appliance tag.

    #---------------------------------------------------------------------------------------------
    # 4. R03 WINDOWS DESKTOP -> Replace (Parameter_Map #15). Never matches "Windows Server".
    #---------------------------------------------------------------------------------------------
    DesktopOsPatterns      = @('Windows\s+(XP|Vista|7|8(\.x|\.1)?|10|11)\b')   # [R03] OS names that count as a Windows desktop (XP, Vista, 7, 8, 10, 11).
    DesktopExcludePattern  = 'Server'          # [R03] An OS name containing this is never treated as a desktop.
    DesktopTarget          = 'Amazon WorkSpaces (WorkSpaces Graphics if the user needs a GPU)'   # [R03] AWS target text for Windows desktops.

    #---------------------------------------------------------------------------------------------
    # 5. R05 DATABASE SERVER -> Manual Review (AWS version of Azure R07)
    #    Matched on the VM name (In = 'Name') or name | notes | tags (In = 'All'). First match wins, so
    #    keep the specific engines (PostgreSQL 'PSQL', MySQL) above the generic 'SQL' (Microsoft SQL Server).
    #---------------------------------------------------------------------------------------------
    # [R05] Database engines recognised from the VM name, with the managed AWS service to consider for each. First match wins.
    DatabaseProducts = @(
        @{ Match = 'postgres|pgsql|psql';  In = 'Name'; Engine = 'PostgreSQL';            Target = 'Amazon RDS for PostgreSQL or Aurora PostgreSQL' }
        @{ Match = 'mysql';                In = 'Name'; Engine = 'MySQL';                 Target = 'Amazon RDS for MySQL or Aurora MySQL' }
        @{ Match = 'maria';                In = 'Name'; Engine = 'MariaDB';               Target = 'Amazon RDS for MariaDB' }
        @{ Match = 'mongo';                In = 'Name'; Engine = 'MongoDB';               Target = 'Amazon DocumentDB (MongoDB compatible) or MongoDB Atlas on AWS' }
        @{ Match = 'oracle|^ora[-_]|oradb'; In = 'Name'; Engine = 'Oracle Database';      Target = 'Amazon RDS for Oracle (licence included or BYOL)' }
        @{ Match = 'sql';                  In = 'Name'; Engine = 'Microsoft SQL Server';  Target = 'Amazon RDS for SQL Server' }
    )
    DatabaseStrategy       = 'Manual Review'   # [R05] Strategy for database servers: 'Manual Review' (default: every database reviewed), 'Replatform' (managed AWS DB) or 'Rehost' (EC2).

    #---------------------------------------------------------------------------------------------
    # 6. R06 OLD OR 32-BIT OS / OS LIFECYCLE (Parameter_Map #18, #32, #11)
    #    EndOfSupport = vendor standard support end; ExtendedEnd = end of the paid extension available
    #    OUTSIDE Azure (Microsoft ESU / Red Hat ELS / SUSE LTSS / Ubuntu Pro); '' = none / not tracked.
    #    First match wins - keep specific patterns above generic ones. Review dates before each run.
    #    Live = '<product>:<release>' (endoflife.date) -> EndOfSupport / ExtendedEnd are refreshed from the API at run time;
    #    the dates typed here are the fallback. LiveAlt = alternatives separated by '|' (tried in order when Live is not found) cover
    #    release names that may differ (e.g. LiveAlt = 'windows:10-22h2-e|windows:10-22h2-w'). No Live key = the dates here are always used. An empty live value never erases a date typed here. LiveEndField = 'eoes'
    #    takes the API's extended (LTS) end as EndOfSupport, for rows whose support end here is the free LTS end (Debian).
    #---------------------------------------------------------------------------------------------
    SupportCutoffDate      = '2027-12-31'      # [R06] Cut-off date: an OS whose support ends before it counts as old. A fixed date 'yyyy-MM-dd', or relative to the run date: 'today+12m' (months) / 'today+2y' (years) / 'today-6m'.
    LegacyOsStrategy       = 'Replatform'      # [R06] Strategy for an old OS: 'Replatform' (upgrade/rebuild in the move) or 'Rehost' (move as-is, upgrade later).
    PreferRelocateForLegacyOS = $false         # [R06] $true = OS past ALL support goes to Relocate (VMware in AWS) instead; script switch can force it on.
    RelocateTarget         = 'Amazon Elastic VMware Service (Amazon EVS) or VMware Cloud on AWS, via VMware HCX'   # [R06] AWS target text for Relocate.
    Arch32BitPatterns      = @('32-bit', '\bi[3-6]86\b', '\bx86\)')   # [R06] OS name patterns that mark a 32-bit OS.
    Allow32BitImport       = $false            # [R06] $false = 32-bit OS is not importable and is rebuilt on 64-bit; $true = ignore 32-bit.
    Arch32BitNote          = 'AWS VM Import stopped accepting 32-bit (i386) images on 1 Apr 2026'   # [R06] Reason text shown for 32-bit VMs.
    # [R06] Support end dates per OS version (yyyy-MM-dd; '' = none/unknown); sets OS_Status. First match wins.
    OsLifecycle = @(
        @{ Match = 'Windows Server\D*2003|Windows (NT|2000|XP|Vista)\b'; Label = 'Windows Server 2003 / XP or older'; EndOfSupport = '2015-07-14'; ExtendedEnd = '';           Live = 'windows-server:2003-r2-sp2'; LiveAlt = 'windows-server:2003-sp2|windows-server:2003'; Note = 'No extended support available' }
        @{ Match = 'Windows Server\D*2008';                  Label = 'Windows Server 2008 / 2008 R2'; EndOfSupport = '2020-01-14'; ExtendedEnd = '2023-01-10'; Note = 'ESU ended Jan 2023 outside Azure' }   # no Live key: the API may give the Azure-only 4th ESU year (Jan 2024)
        @{ Match = 'Windows Server\D*2012';                  Label = 'Windows Server 2012 / 2012 R2'; EndOfSupport = '2023-10-10'; ExtendedEnd = '2026-10-13'; Live = 'windows-server:2012-r2'; LiveAlt = 'windows-server:2012'; Note = 'Paid ESU ended Oct 2026' }
        @{ Match = 'Windows Server\D*2016 or later';         Label = 'Windows Server 2016 or later (exact version unknown)'; EndOfSupport = ''; ExtendedEnd = ''; Note = 'Grouped VMware label - confirm the exact version' }
        @{ Match = 'Windows Server\D*2016';                  Label = 'Windows Server 2016';           EndOfSupport = '2027-01-12'; ExtendedEnd = '';           Live = 'windows-server:2016'; Note = '' }
        @{ Match = 'Windows Server\D*2019';                  Label = 'Windows Server 2019';           EndOfSupport = '2029-01-09'; ExtendedEnd = '';           Live = 'windows-server:2019'; Note = '' }
        @{ Match = 'Windows Server\D*2022';                  Label = 'Windows Server 2022';           EndOfSupport = '2031-10-14'; ExtendedEnd = '';           Live = 'windows-server:2022'; Note = '' }
        @{ Match = 'Windows Server\D*2025';                  Label = 'Windows Server 2025';           EndOfSupport = '2034-11-14'; ExtendedEnd = '';           Live = 'windows-server:2025'; Note = '' }
        @{ Match = 'Windows 7\b';                            Label = 'Windows 7';                     EndOfSupport = '2020-01-14'; ExtendedEnd = '2023-01-10'; Live = 'windows:7-sp1'; LiveAlt = 'windows:7'; Note = 'ESU ended Jan 2023' }
        @{ Match = 'Windows 8';                              Label = 'Windows 8 / 8.1';               EndOfSupport = '2023-01-10'; ExtendedEnd = '';           Live = 'windows:8.1'; LiveAlt = 'windows:8-1|windows:8'; Note = '' }
        @{ Match = 'Windows 10\b';                           Label = 'Windows 10';                    EndOfSupport = '2025-10-14'; ExtendedEnd = '2028-10-10'; Live = 'windows:10-22h2'; LiveAlt = 'windows:10-22h2-e|windows:10-22h2-w'; Note = 'Commercial ESU (paid, organisations) to Oct 2028; consumer ESU ended 13 Oct 2026' }
        @{ Match = 'Windows 11\b';                           Label = 'Windows 11';                    EndOfSupport = '';           ExtendedEnd = '';           Note = 'Support depends on the feature-update version' }
        @{ Match = '(Red Hat Enterprise Linux|RHEL)\D*\b[2-5](\.\d+)?\b'; Label = 'RHEL 5 or older'; EndOfSupport = '2017-03-31'; ExtendedEnd = '2020-11-30'; Live = 'rhel:5'; Note = '' }
        @{ Match = '(Red Hat Enterprise Linux|RHEL)\D*\b6(\.\d+)?\b';    Label = 'RHEL 6';          EndOfSupport = '2020-11-30'; ExtendedEnd = '2024-06-30'; Live = 'rhel:6'; Note = 'ELS ended Jun 2024 (verify: Red Hat may offer a longer ELS)' }
        @{ Match = '(Red Hat Enterprise Linux|RHEL)\D*\b7(\.\d+)?\b';    Label = 'RHEL 7';          EndOfSupport = '2024-06-30'; ExtendedEnd = '2028-06-30'; Live = 'rhel:7'; Note = 'ELS add-on required' }
        @{ Match = '(Red Hat Enterprise Linux|RHEL)\D*\b8(\.\d+)?\b';    Label = 'RHEL 8';          EndOfSupport = '2029-05-31'; ExtendedEnd = '2032-05-31'; Live = 'rhel:8'; Note = '' }
        @{ Match = '(Red Hat Enterprise Linux|RHEL)\D*\b9(\.\d+)?\b';    Label = 'RHEL 9';          EndOfSupport = '2032-05-31'; ExtendedEnd = '2035-05-31'; Live = 'rhel:9'; Note = '' }
        @{ Match = '(Red Hat Enterprise Linux|RHEL)\D*\b10(\.\d+)?\b';   Label = 'RHEL 10';         EndOfSupport = '2035-05-31'; ExtendedEnd = '2038-05-31'; Live = 'rhel:10'; Note = '' }
        @{ Match = 'Oracle Linux\D*\b[4-6](\.\d+)?\b';       Label = 'Oracle Linux 6 or older';      EndOfSupport = '2021-03-01'; ExtendedEnd = '2024-12-31'; Live = 'oracle-linux:6'; Note = '' }
        @{ Match = 'Oracle Linux\D*\b7(\.\d+)?\b';           Label = 'Oracle Linux 7';               EndOfSupport = '2024-12-31'; ExtendedEnd = '2028-06-30'; Live = 'oracle-linux:7'; Note = 'Extended support required' }
        @{ Match = 'Oracle Linux\D*\b8(\.\d+)?\b';           Label = 'Oracle Linux 8';               EndOfSupport = '2029-07-31'; ExtendedEnd = '2032-07-31'; Live = 'oracle-linux:8'; Note = '' }
        @{ Match = 'Oracle Linux\D*\b9(\.\d+)?\b';           Label = 'Oracle Linux 9';               EndOfSupport = '2032-06-30'; ExtendedEnd = '2035-06-30'; Live = 'oracle-linux:9'; Note = '' }
        @{ Match = 'CentOS\D*\b[4-6](\.\d+)?\b';             Label = 'CentOS 6 or older';            EndOfSupport = '2020-11-30'; ExtendedEnd = '';           Live = 'centos:6'; Note = '' }
        @{ Match = 'CentOS\D*\b7(\.\d+)?\b';                 Label = 'CentOS 7';                     EndOfSupport = '2024-06-30'; ExtendedEnd = '';           Live = 'centos:7'; Note = '' }
        @{ Match = 'CentOS\D*\b8(\.\d+)?\b';                 Label = 'CentOS 8';                     EndOfSupport = '2021-12-31'; ExtendedEnd = '';           Live = 'centos:8'; Note = '' }
        @{ Match = 'Debian\D*\b([4-9]|10)(\.\d+)?\b';        Label = 'Debian 10 or older';           EndOfSupport = '2024-06-30'; ExtendedEnd = '';           Live = 'debian:10'; LiveEndField = 'eoes'; Note = '' }
        @{ Match = 'Debian\D*\b11(\.\d+)?\b';                Label = 'Debian 11';                    EndOfSupport = '2026-08-31'; ExtendedEnd = '';           Live = 'debian:11'; LiveEndField = 'eoes'; Note = 'LTS ended Aug 2026' }
        @{ Match = 'Debian\D*\b12(\.\d+)?\b';                Label = 'Debian 12';                    EndOfSupport = '2028-06-30'; ExtendedEnd = '';           Live = 'debian:12'; LiveEndField = 'eoes'; Note = '' }
        @{ Match = 'Debian\D*\b13(\.\d+)?\b';                Label = 'Debian 13';                    EndOfSupport = '2030-06-30'; ExtendedEnd = '';           Live = 'debian:13'; LiveEndField = 'eoes'; Note = '' }
        @{ Match = 'Ubuntu\D*(1[0-6]\.\d+)';                 Label = 'Ubuntu 16.04 or older';        EndOfSupport = '2021-04-30'; ExtendedEnd = '';           Live = 'ubuntu:16.04'; Note = '' }
        @{ Match = 'Ubuntu\D*18\.04';                        Label = 'Ubuntu 18.04 LTS';             EndOfSupport = '2023-05-31'; ExtendedEnd = '2028-04-30'; Live = 'ubuntu:18.04'; Note = 'Ubuntu Pro ESM required' }
        @{ Match = 'Ubuntu\D*20\.04';                        Label = 'Ubuntu 20.04 LTS';             EndOfSupport = '2025-05-31'; ExtendedEnd = '2030-04-30'; Live = 'ubuntu:20.04'; Note = 'Ubuntu Pro ESM required' }
        @{ Match = 'Ubuntu\D*22\.04';                        Label = 'Ubuntu 22.04 LTS';             EndOfSupport = '2027-04-30'; ExtendedEnd = '2032-04-30'; Live = 'ubuntu:22.04'; Note = '' }
        @{ Match = 'Ubuntu\D*24\.04';                        Label = 'Ubuntu 24.04 LTS';             EndOfSupport = '2029-04-30'; ExtendedEnd = '2034-04-30'; Live = 'ubuntu:24.04'; Note = '' }
        @{ Match = '(SUSE Linux Enterprise|SLES)\D*\b(9|10|11)\b'; Label = 'SLES 11 or older';       EndOfSupport = '2022-03-31'; ExtendedEnd = '';           Note = '' }
        @{ Match = '(SUSE Linux Enterprise|SLES)\D*\b12\b';  Label = 'SLES 12';                      EndOfSupport = '2024-10-31'; ExtendedEnd = '2027-10-31'; Live = 'sles:12.5'; Note = 'LTSS required' }
        @{ Match = '(SUSE Linux Enterprise|SLES)\D*\b15\b';  Label = 'SLES 15';                      EndOfSupport = '2031-07-31'; ExtendedEnd = '';           Note = '' }
        @{ Match = 'Rocky\D*\b8\b';                          Label = 'Rocky Linux 8';                EndOfSupport = '2029-05-31'; ExtendedEnd = '';           Live = 'rocky-linux:8'; Note = '' }
        @{ Match = 'Rocky\D*\b9\b';                          Label = 'Rocky Linux 9';                EndOfSupport = '2032-05-31'; ExtendedEnd = '';           Live = 'rocky-linux:9'; Note = '' }
    )

    #---------------------------------------------------------------------------------------------
    # 6b. EXACT OS VERSION CHOOSER (workbook column Exact_OS_Version, sheet OS_Versions)
    #     RVTools shows only a family label for some Windows VMs ('Windows 10', 'Windows 11', 'Windows Server 2012' for both
    #     2012 and 2012 R2, 'Windows Server 2016 or later'). For a VM whose OS (field in use, no override) matches Match, the
    #     client picks the exact version from a dropdown and the dates, OS status, strategy, target and readiness update live.
    #     The options and their dates come from the lifecycle data (Product = endoflife.date product; CyclePattern selects
    #     its releases), so new versions appear automatically. OsTemplate = the OS text the rules are evaluated with for an
    #     option ({0} = the version, e.g. '2012 R2', '2019'; {1} = the VM's bitness text '(64-bit)' / '(32-bit)', kept from RVTools).
    #     HideEndedBefore = options whose support AND paid extended support both ended before this date are not offered
    #     ('today-36m' = 3 years before the run; '' = show every release). First match wins: keep '2016 or later' above other Server rows.
    #---------------------------------------------------------------------------------------------
    # [R03/R06/R09] Families the client can refine to an exact version (first match wins).
    VersionChoices = @(
        @{ Match = 'Windows Server\D*2016 or later'; Family = 'Windows Server 2016 or later'; Product = 'windows-server'; CyclePattern = '^20(1[6-9]|[2-9]\d)$'; HideEndedBefore = ''; OsTemplate = 'Microsoft Windows Server {0} {1}'; Help = 'Inside the VM run winver, or in PowerShell: (Get-ComputerInfo).OsName - it shows e.g. "Windows Server 2019 Standard"' }
        @{ Match = 'Windows Server\D*2012';          Family = 'Windows Server 2012 / 2012 R2'; Product = 'windows-server'; CyclePattern = '^2012(-r2)?$';         HideEndedBefore = '';          OsTemplate = 'Microsoft Windows Server {0} {1}'; Help = 'Inside the VM run winver, or in PowerShell: (Get-ComputerInfo).OsName - it shows "Windows Server 2012" or "Windows Server 2012 R2"' }
        @{ Match = 'Windows 10\b';                   Family = 'Windows 10';                    Product = 'windows';        CyclePattern = '^10-';                 HideEndedBefore = 'today-36m'; OsTemplate = 'Microsoft Windows 10 {1}'; Help = 'Inside the VM press Win+R and run winver: it shows the version (e.g. "Version 22H2") and the edition (Pro, Enterprise, LTSC ...)' }
        @{ Match = 'Windows 11\b';                   Family = 'Windows 11';                    Product = 'windows';        CyclePattern = '^11-';                 HideEndedBefore = 'today-36m'; OsTemplate = 'Microsoft Windows 11 {1}'; Help = 'Inside the VM press Win+R and run winver: it shows the version (e.g. "Version 24H2") and the edition (Pro, Enterprise, LTSC ...)' }
    )

    #---------------------------------------------------------------------------------------------
    # 7. AWS VM IMPORT COMPATIBILITY (Parameter_Map #17, #19, #20)
    #    Source: docs.aws.amazon.com/vm-import/latest/userguide/prerequisites.html (re-check before each wave)
    #---------------------------------------------------------------------------------------------
    # [R09] OS names that AWS VM Import accepts (Rehost possible). Checked after the not-accepted list below.
    VmImportSupportedOs    = @(
        'Windows Server\D*(2003|2008|2012|2016|2019|2022|2025)', 'Windows (7|8|10|11)\b',
        'Red Hat Enterprise Linux|RHEL', 'CentOS', 'Oracle Linux', 'SUSE Linux Enterprise|SLES',
        'Ubuntu', 'Debian\D*\b(6|7|10|11|12|13)\b', 'Rocky\D*\b(9|10)\b', 'Fedora', 'Amazon Linux'
    )
    VmImportUnsupportedOs  = @('Debian\D*\b(8|9)\b', 'openSUSE', 'AlmaLinux', 'Rocky\D*\b8\b', 'VMware ESXi', 'FreeBSD', 'Solaris')   # [R07] OS names AWS VM Import does NOT accept (Rehost blocked). Checked first.
    UnversionedOsPatterns  = @('^Ubuntu Linux \((32|64)-bit\)$', '^Rocky Linux \((32|64)-bit\)$', '^Other\b.*\bLinux\b')   # [R08] OS labels with a family but no version: guest check needed.
    GroupedOsLabels        = @('Windows Server\D*2016 or later')   # [R09] Grouped labels on the import list that hide the exact version: Rehost + guest-check action.
    VmImportVerifyOs       = @('^CentOS 4/5')  # [R08] Grouped labels only partly importable (CentOS 4/5/6/7): VM_Import_OS = 'Version not identified' -> guest check.
    GroupedOsLabelAction   = 'Guest check: confirm the exact Windows Server version (winver / (Get-ComputerInfo).OsName) - 2016 support ends 12 Jan 2027'   # [R09] Action text for grouped labels.
    UnknownOsPatterns      = @('^Other \((32|64)-bit\)$')   # [R08] OS labels with no usable OS at all (a blank OS is treated the same).

    #---------------------------------------------------------------------------------------------
    # 8. READINESS CHECKS R11-R12 (never change the strategy)
    #---------------------------------------------------------------------------------------------
    MaxEbsVolumeGiB        = 16384             # [R11] Disk size limit of AWS VM Import in GiB (16384 = 16 TiB): a disk this big or bigger must be split/moved.
    ManifestImportGiB      = 8192              # [R11] Disk size in GiB (8192 = 8 TiB) from which the larger-disk (manifest) import route is needed.
    TemplateValuePattern   = '^(true|yes|1)$'  # [R01] vInfo 'Template' values meaning "is a template": excluded from the strategy and the R01 count.
    MaxVolumesPerImport    = 21                # [R11] Max disks per VM in one VM Import task; extra disks must be imported separately.
    RawDiskValuePattern    = '^(true|yes|1)$'  # [R11] vDisk 'Raw' values that mark a raw (RDM) disk - blocks image export.
    SharedModePattern      = 'multi'           # [R11] vDisk 'Sharing mode' value that marks a shared disk - blocks image export.
    SharedBusPattern       = 'physical|virtual'   # [R11] vDisk 'Shared Bus' values that mark a shared disk - blocks image export.
    DmzPatterns            = @('dmz')          # [R12] Text in the VM name or cluster name that marks a DMZ server (moves in the DMZ wave).
    AntiAffinityPattern    = 'Anti.?Affinity'  # Design note: vInfo 'Cluster rule(s)' text marking an anti-affinity rule (spread across AZs).

    #---------------------------------------------------------------------------------------------
    # 9. REVIEW FLAGS AND DESIGN CONSIDERATIONS (never change the strategy)
    #---------------------------------------------------------------------------------------------
    StaleNamePatterns      = @('[_\-](restore|restored|old|backup|bak|tbd|copy)([_\-\d]|$)', 'clonemaster', '_te?mp$', '^copy[-_ ]', '[-_]test$', '^test[-_]')   # Review flag: names of restore / backup / old copy / test VMs (possible Retire).
    ClusterNamePatterns    = @('cluster', 'SQLCL(?![A-Za-z])', 'SQLAO', '[-_]FCI', '[-_]CL\d')   # Design note: names suggesting a cluster (SQL Always On, FCI): migrate as a unit.
    LargeDiskGB            = 1024              # Design note: total storage in GB from which EBS layout / IOPS planning is noted.
    HighRamGB              = 32                # Design note: RAM in GB above which a memory-optimised instance is suggested.
    MultiNicThreshold      = 2                 # Design note: NIC count from which a note is added (VM Import gives the VM one NIC).
    HighDiskCount          = 16                # Design note: disk count from which the instance's EBS volume limit must be checked.
    StorageColumn          = 'Total disk capacity MiB'   # vInfo column for VM storage size: 'Total disk capacity MiB', 'Provisioned MiB' or 'In Use MiB'.

    #---------------------------------------------------------------------------------------------
    # 10. CONFIDENCE
    #---------------------------------------------------------------------------------------------
    DefaultRehostConfidence   = 'Medium'       # [R09] Confidence given to a standard Rehost ('High', 'Medium' or 'Low').
    AllowHighRehostConfidence = $false         # [R09] $true = 'High' when the OS is from the override file or VMware Tools, fully supported and the OS columns agree.

    #---------------------------------------------------------------------------------------------
    # 10b. INTERNAL REGISTER SUPPORT (not written to the client workbook) - columns added by the manager (Highlighted Comments, Decision support,
    #      per-site count, Client decision, Client comments)
    #---------------------------------------------------------------------------------------------
    ParameterMapDecisionSupport        = @{ '11' = 'Yes'; 'DB' = 'Yes/No' }   # Internal register only (optional; not in the client workbook): 'Decision support' per row.
    ParameterMapDecisionSupportDefault = 'Yes/No'   # Internal register only (optional; not in the client workbook): 'Decision support' for unlisted rows.
    SiteNameFromFilePattern            = '^(.+?)_RVTools'   # Site name from the input file name ('DE-HUB_RVTools_...' -> 'DE-HUB') unless -SiteName given; used in the default output file name.
    ParameterMapSiteCountDecidedRows   = @('2', '14', '15', 'DB', '17', '18', '19', '20', '32')   # Internal register only (optional; not in the client workbook): rows counting VMs decided by the rule.

    #---------------------------------------------------------------------------------------------
    # 11. TARGET LABELS written to the AWS_Target column
    #---------------------------------------------------------------------------------------------
    # Standard texts written to the AWS_Target column (vendor, database, desktop and relocate targets are set above).
    Targets = @{
        Decommission    = 'None - decommission'            # [R02] Platform components that are retired.
        OnPremises      = 'Stays on-premises'              # Not used by the current rules (reserved for Retain); required by the loader.
        Ec2             = 'Amazon EC2'                     # [R05 Rehost, R06 Rehost, R07, R08, R09] Plain Rehost to EC2.
        Ec2SupportedOs  = 'Amazon EC2 on a supported OS'   # [R06] Old OS upgraded or rebuilt (Replatform).
        Ec2Os64Bit      = 'Amazon EC2 on a 64-bit OS'      # [R06] 32-bit OS rebuilt on 64-bit (Replatform).
    }
}
