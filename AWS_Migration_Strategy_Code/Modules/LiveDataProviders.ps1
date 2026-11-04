#---------------------------------------------------------------------
# LIVE DATA PROVIDERS (v3.0)
#   - OS support dates and Windows version lists from the endoflife.date API (no login)
#   - AWS VM Import rules page change check (the page has no API: its text is hashed)
#   - Optional AWS API facts (needs the AWS CLI and read-only credentials): exact max NICs and CPU
#     architecture per instance type, services offered in the region, Windows AMIs offered
#   - Rows for the workbook sheet Data_Sources (where every value came from)
# Every network call is wrapped: a failure falls back to cache -> shipped snapshot -> settings file
# and is reported as a message, never as an error that stops the run.
# Compatible with Windows PowerShell 5.1 and PowerShell 7.
#---------------------------------------------------------------------

function Get-LiveCacheRoot {
    # Folder for downloaded live data. $env:TEMP on Windows; the system temp folder elsewhere (TEMP is unset on Linux/macOS).
    param([string]$SubFolder = 'AWSMigrationLiveCache')
    $base = $env:TEMP
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetTempPath() }
    $dir = Join-Path $base $SubFolder
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    return $dir
}

function Get-LiveDataSetting {
    # One LiveData setting with its default (the whole LiveData block is optional in the settings file).
    param([hashtable]$Settings, [string]$Name)
    $defaults = @{
        Enabled              = $true
        CacheDays            = 7
        LifecycleApiBase     = 'https://endoflife.date/api/v1/products'
        LifecycleLegacyBase  = 'https://endoflife.date/api'
        VmImportDocUrl       = 'https://docs.aws.amazon.com/vm-import/latest/userguide/prerequisites.html'
        VmImportDocKnownHash = ''
        TimeoutSec           = 20
    }
    if ($Settings -and $Settings.ContainsKey('LiveData') -and $Settings.LiveData -is [hashtable] -and $Settings.LiveData.ContainsKey($Name)) {
        return $Settings.LiveData[$Name]
    }
    return $defaults[$Name]
}

function Get-LiveHttpStatusCode {
    # HTTP status code carried by a web exception (PowerShell 7 HttpResponseException / 5.1 WebException), or 0
    # when the target never answered (timeout, DNS, connection refused, proxy refusing the tunnel).
    param($Exception)
    $e = $Exception
    while ($null -ne $e) {
        $resp = $null
        $p = $e.PSObject.Properties['Response']
        if ($null -ne $p) { try { $resp = $p.Value } catch { $resp = $null } }
        if ($null -ne $resp) {
            $sc = $resp.PSObject.Properties['StatusCode']
            if ($null -ne $sc) { try { $code = [int]$sc.Value; if ($code -gt 0) { return $code } } catch { } }
        }
        $e = $e.InnerException
    }
    return 0
}

function Invoke-LiveHttpRequest {
    <#
        GET a URL. Returns [pscustomobject]@{ Text; StatusCode; TransportFailure; Error }:
          Text             = body ($null on failure)
          StatusCode       = HTTP status of a failed request (0 = the target never answered)
          TransportFailure = $true when no HTTP response came back (timeout / DNS / refused / proxy tunnel refused):
                             the caller then treats the whole host as unreachable for this run
    #>
    param(
        [Parameter(Mandatory)] [string]$Url,
        [int]$TimeoutSec = 20
    )
    $res = [pscustomobject]@{ Text = $null; StatusCode = 0; TransportFailure = $false; Error = '' }
    try {
        # TLS 1.2 for Windows PowerShell 5.1 (PowerShell 7 negotiates it by default)
        if ($PSVersionTable.PSVersion.Major -lt 6) {
            try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
        }
        $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec $TimeoutSec -ErrorAction Stop
        $content = $resp.Content
        if ($content -is [byte[]]) { $content = [System.Text.Encoding]::UTF8.GetString($content) }
        $res.Text = [string]$content
    }
    catch {
        $res.Error = [string]$_.Exception.Message
        $res.StatusCode = Get-LiveHttpStatusCode $_.Exception
        $res.TransportFailure = ($res.StatusCode -eq 0)
    }
    return $res
}

function Invoke-LiveHttpGet {
    # GET a URL and return its text, or $null on any failure (the reason is added to $Messages).
    param(
        [Parameter(Mandatory)] [string]$Url,
        [int]$TimeoutSec = 20,
        [System.Collections.Generic.List[string]]$Messages
    )
    $r = Invoke-LiveHttpRequest -Url $Url -TimeoutSec $TimeoutSec
    if ($null -eq $r.Text -and $null -ne $Messages) { $Messages.Add(("Could not read {0}: {1}" -f $Url, $r.Error)) }
    return $r.Text
}

function ConvertFrom-LiveJsonText {
    <#
        ConvertFrom-Json for endoflife.date / cache / snapshot text. Timestamps are cut to their date part in the RAW
        text first ("2029-01-09T23:30:00-05:00" -> "2029-01-09"), because PowerShell 7 turns ISO timestamps into
        local DateTime values and a time-zone offset could otherwise move the date by a day.
    #>
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $t = [regex]::Replace($Text, '"(\d{4}-\d{2}-\d{2})T[0-9:.]*(Z|[+-]\d{2}:?\d{2})?"', '"$1"')
    $obj = $t | ConvertFrom-Json
    # PowerShell 7 unrolls a one-element JSON array into a single object: keep a JSON array an array (legacy API shape)
    if ($t.TrimStart().StartsWith('[')) { return , @($obj) }
    return $obj
}

function ConvertTo-LiveDate {
    # endoflife.date values can be a date string, a DateTime (PowerShell 7 converts ISO strings), true/false or null.
    # Only real dates are kept ('yyyy-MM-dd'); everything else means "unknown" ('').
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [bool]) { return '' }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [System.DateTimeOffset]) { return $Value.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture) }
    $s = ([string]$Value).Trim()
    if ($s -match '^(\d{4}-\d{2}-\d{2})') { return $Matches[1] }
    return ''
}

function ConvertTo-LiveBool {
    param($Value)
    if ($null -eq $Value) { return $false }
    if ($Value -is [bool]) { return $Value }
    if ($Value -is [datetime]) { return $true }      # 'lts: <date>' in the legacy API means LTS
    $s = ([string]$Value).Trim()
    if ($s -match '^(true|yes|1)$') { return $true }
    if ($s -match '^\d{4}-\d{2}-\d{2}') { return $true }
    return $false
}

function Get-LiveProperty {
    # Property of a PSCustomObject from ConvertFrom-Json, or $null when absent (works on 5.1 and 7).
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p) { return $null }
    return $p.Value
}

function Get-LifecycleDisplayName {
    <#
        Client-facing name for one endoflife.date release, e.g.
          windows        11-24h2-e      -> 'Windows 11 24H2 - Enterprise / Education'
          windows        11-24h2-w      -> 'Windows 11 24H2 - Home / Pro'
          windows        10-22h2        -> 'Windows 10 22H2'
          windows        10-21h2-e-lts  -> 'Windows 10 Enterprise LTSC 2021'
          windows        10-21h2-iot-lts-> 'Windows 10 IoT Enterprise LTSC 2021'
          windows-server 2012-r2        -> 'Windows Server 2012 R2'
        Release names are parsed generically; anything unrecognised falls back to the API label.
    #>
    param([string]$Product, [string]$Cycle, [string]$Label)
    $c = ([string]$Cycle).Trim().ToLowerInvariant()
    $ltscYear = @{ '1507' = '2015'; '1607' = '2016'; '1809' = '2019'; '21h2' = '2021'; '24h2' = '2024' }
    switch ($Product) {
        'windows' {
            if ($c -match '^(?<maj>\d+(\.\d+)?)(-(?<ver>[0-9a-z]+?))?(?<rest>(-[a-z]+)*)$') {
                $maj  = $Matches['maj']; $ver = [string]$Matches['ver']; $rest = [string]$Matches['rest']
                $isLts = $rest -match '-ltsc?\b|-lts$'
                $isIot = $rest -match '-iot\b'
                $verUp = $ver.ToUpperInvariant()
                if ($isLts) {
                    $year = if ($ltscYear.ContainsKey($ver)) { $ltscYear[$ver] } else { $verUp }
                    $kind = if ($year -eq '2015' -or $year -eq '2016') { 'LTSB' } else { 'LTSC' }
                    if ($isIot) { return ('Windows {0} IoT Enterprise {1} {2}' -f $maj, $kind, $year) }
                    return ('Windows {0} Enterprise {1} {2}' -f $maj, $kind, $year)
                }
                $edition = ''
                if ($rest -match '-e$')       { $edition = ' - Enterprise / Education' }
                elseif ($rest -match '-w$')   { $edition = ' - Home / Pro' }
                elseif ($rest -match '-iot$') { $edition = ' - IoT Enterprise' }
                elseif ($rest -ne '')         { $edition = '' }
                if ($ver -ne '' -and ($rest -eq '' -or $edition -ne '')) {
                    return ('Windows {0} {1}{2}' -f $maj, $verUp, $edition)
                }
                if ($ver -eq '' -and $rest -eq '') { return ('Windows {0}' -f $maj) }
            }
            if ($Label) { return ('Windows {0}' -f $Label) }
            return ('Windows {0}' -f $Cycle)
        }
        'windows-server' {
            if ($c -match '^(?<y>\d{4})(?<r2>-r2)?(?<rest>(-[a-z0-9]+)*)$') {
                $name = 'Windows Server ' + $Matches['y']
                if ($Matches['r2']) { $name += ' R2' }
                $rest = [string]$Matches['rest']
                if ($rest -match '-sp(\d+)') { $name += (' SP' + $Matches[1]) }
                return $name
            }
            if ($c -match '^\d{2}h\d$') { return ('Windows Server {0} (Annual Channel)' -f $c.ToUpperInvariant()) }
            if ($Label) { return ('Windows Server {0}' -f $Label) }
            return ('Windows Server {0}' -f $Cycle)
        }
        default {
            $names = @{ 'rhel' = 'Red Hat Enterprise Linux'; 'oracle-linux' = 'Oracle Linux'; 'centos' = 'CentOS'; 'debian' = 'Debian'; 'ubuntu' = 'Ubuntu'; 'sles' = 'SUSE Linux Enterprise Server'; 'rocky-linux' = 'Rocky Linux'; 'almalinux' = 'AlmaLinux' }
            $pn = if ($names.ContainsKey($Product)) { $names[$Product] } else { $Product }
            $lb = if ($Label) { $Label } else { $Cycle }
            return ('{0} {1}' -f $pn, $lb)
        }
    }
}

function ConvertFrom-LifecycleJson {
    <#
        Normalises one product's JSON into release objects. Accepts:
          - endoflife.date API v1:  { result: { releases: [ { name, label, releaseDate, eoasFrom, eolFrom, eoesFrom, isLts } ] } }
          - the same object without the 'result' wrapper (used in the snapshot file)
          - legacy API:            [ { cycle, releaseDate, support, eol, extendedSupport, lts } ]
        Returns @() when the shape is not recognised.
    #>
    param([Parameter(Mandatory)] [string]$Product, $Json, [string]$Source = '')
    $out = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Json) { return @() }

    $releases = $null; $legacy = $false
    $result = Get-LiveProperty $Json 'result'
    if ($null -ne $result -and $null -ne (Get-LiveProperty $result 'releases')) { $releases = @(Get-LiveProperty $result 'releases') }
    elseif ($null -ne (Get-LiveProperty $Json 'releases')) { $releases = @(Get-LiveProperty $Json 'releases') }
    elseif ($Json -is [System.Array] -or $Json -is [System.Collections.IEnumerable] -and -not ($Json -is [string])) {
        $arr = @($Json)
        if ($arr.Count -gt 0 -and $null -ne (Get-LiveProperty $arr[0] 'cycle')) { $releases = $arr; $legacy = $true }
    }
    if ($null -eq $releases) { return @() }

    foreach ($r in $releases) {
        if ($null -eq $r) { continue }
        if ($legacy) {
            $cycle = [string](Get-LiveProperty $r 'cycle')
            $label = $cycle
            $rel = [pscustomobject]@{
                Product       = $Product
                Cycle         = $cycle
                Label         = $label
                DisplayName   = Get-LifecycleDisplayName -Product $Product -Cycle $cycle -Label $label
                ReleaseDate   = ConvertTo-LiveDate (Get-LiveProperty $r 'releaseDate')
                MainstreamEnd = ConvertTo-LiveDate (Get-LiveProperty $r 'support')
                EndOfSupport  = ConvertTo-LiveDate (Get-LiveProperty $r 'eol')
                ExtendedEnd   = ConvertTo-LiveDate (Get-LiveProperty $r 'extendedSupport')
                IsLts         = ConvertTo-LiveBool (Get-LiveProperty $r 'lts')
                Note          = ''
                Source        = $Source
            }
        }
        else {
            $cycle = [string](Get-LiveProperty $r 'name')
            $label = [string](Get-LiveProperty $r 'label')
            if ([string]::IsNullOrWhiteSpace($label)) { $label = $cycle }
            $rel = [pscustomobject]@{
                Product       = $Product
                Cycle         = $cycle
                Label         = $label
                DisplayName   = Get-LifecycleDisplayName -Product $Product -Cycle $cycle -Label $label
                ReleaseDate   = ConvertTo-LiveDate (Get-LiveProperty $r 'releaseDate')
                MainstreamEnd = ConvertTo-LiveDate (Get-LiveProperty $r 'eoasFrom')
                EndOfSupport  = ConvertTo-LiveDate (Get-LiveProperty $r 'eolFrom')
                ExtendedEnd   = ConvertTo-LiveDate (Get-LiveProperty $r 'eoesFrom')
                IsLts         = ConvertTo-LiveBool (Get-LiveProperty $r 'isLts')
                Note          = [string](Get-LiveProperty $r 'note')   # optional, snapshot only (e.g. commercial vs consumer ESU)
                Source        = $Source
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($rel.Cycle)) { $out.Add($rel) }
    }
    return $out.ToArray()
}

function Get-RowLiveKey {
    # Live key of an OsLifecycle row plus its optional LiveAlt alternatives ('product:cycle|product:cycle'), as one string.
    param($Row)
    if (-not ($Row -is [hashtable])) { return '' }
    $k = if ($Row.ContainsKey('Live')) { [string]$Row.Live } else { '' }
    if ($Row.ContainsKey('LiveAlt') -and -not [string]::IsNullOrWhiteSpace([string]$Row.LiveAlt)) {
        $k = if ([string]::IsNullOrWhiteSpace($k)) { [string]$Row.LiveAlt } else { $k + '|' + [string]$Row.LiveAlt }
    }
    return $k.Trim()
}

function Get-LiveKeyAlternatives {
    <#
        Parses an OsLifecycle Live key into its alternatives, first = preferred:
          'rhel:7'                                            -> rhel/7
          'windows:10-22h2|windows:10-22h2-e|windows:10-22h2-w' -> three alternatives (first one found wins)
          'windows:10-22h2|10-22h2-e'                          -> an alternative without a product reuses the previous product
        Returns @( [pscustomobject]@{ Product; Cycle } ); invalid parts are skipped.
    #>
    param([string]$Live)
    $out = New-Object System.Collections.Generic.List[object]
    $lastProduct = ''
    foreach ($part in ([string]$Live).Split('|')) {
        $t = $part.Trim()
        if (-not $t) { continue }
        $i = $t.IndexOf(':')
        if ($i -gt 0) { $product = $t.Substring(0, $i).Trim().ToLowerInvariant(); $cycle = $t.Substring($i + 1).Trim() }
        elseif ($lastProduct) { $product = $lastProduct; $cycle = $t }
        else { continue }
        if (-not $product -or -not $cycle) { continue }
        $lastProduct = $product
        $out.Add([pscustomobject]@{ Product = $product; Cycle = $cycle })
    }
    return $out.ToArray()
}

function Get-LifecycleProductNames {
    # Every endoflife.date product the settings need: OsLifecycle 'Live' keys + VersionChoices products.
    param([Parameter(Mandatory)] [hashtable]$Settings)
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($row in @($Settings.OsLifecycle)) {
        $key = Get-RowLiveKey $row
        if ($key) {
            foreach ($alt in (Get-LiveKeyAlternatives $key)) {
                if (-not $names.Contains($alt.Product)) { $names.Add($alt.Product) }
            }
        }
    }
    if ($Settings.ContainsKey('VersionChoices')) {
        foreach ($vc in @($Settings.VersionChoices)) {
            if ($vc -is [hashtable] -and $vc.ContainsKey('Product')) {
                $p = ([string]$vc.Product).Trim().ToLowerInvariant()
                if ($p -and -not $names.Contains($p)) { $names.Add($p) }
            }
        }
    }
    return $names.ToArray()
}

function Get-LifecycleCatalog {
    <#
        Builds the OS lifecycle catalogue used by Update-OsLifecycleFromCatalog and the version chooser.
        Per product, first that works: live API (v1, then legacy) -> cache (<= CacheDays) -> older cache
        ('StaleCache', with a warning) -> Config\LifecycleSnapshot.json ('Snapshot'). A product found
        nowhere is simply absent: the settings-file dates are then used.
    #>
    param(
        [Parameter(Mandatory)] [hashtable]$Settings,
        [switch]$Offline,
        [string]$CacheDir,
        [string]$SnapshotPath
    )
    $messages = New-Object System.Collections.Generic.List[string]   # full detail (console)
    $notes    = New-Object System.Collections.Generic.List[string]   # short general notes (sheet Data_Sources)
    $catalog = [pscustomobject]@{
        Products    = @{}
        Source      = @{}      # '<product>' = Live | Cache | StaleCache | Snapshot | None
        RetrievedAt = @{}
        Reason      = @{}      # '<product>' = short reason why the live API was not used ('' when live)
        Messages    = $messages
        Notes       = $notes
    }
    if ([string]::IsNullOrWhiteSpace($CacheDir)) { $CacheDir = Get-LiveCacheRoot }
    if (-not (Test-Path -LiteralPath $CacheDir)) { New-Item -ItemType Directory -Path $CacheDir -Force | Out-Null }
    if ([string]::IsNullOrWhiteSpace($SnapshotPath)) {
        $cfgDir = $null
        if ($Settings.ContainsKey('SettingsPath') -and $Settings.SettingsPath) { $cfgDir = Split-Path -Path $Settings.SettingsPath -Parent }
        if (-not $cfgDir) { $cfgDir = Join-Path (Split-Path $PSScriptRoot -Parent) 'Config' }
        $SnapshotPath = Join-Path $cfgDir 'LifecycleSnapshot.json'
    }

    $enabled   = [bool](Get-LiveDataSetting $Settings 'Enabled')
    $cacheDays = [double](Get-LiveDataSetting $Settings 'CacheDays')
    $timeout   = [int](Get-LiveDataSetting $Settings 'TimeoutSec')
    if ($timeout -le 0) { $timeout = 20 }
    $v1Base    = ([string](Get-LiveDataSetting $Settings 'LifecycleApiBase')).TrimEnd('/')
    $oldBase   = ([string](Get-LiveDataSetting $Settings 'LifecycleLegacyBase')).TrimEnd('/')
    $online    = $enabled -and -not $Offline
    $offlineReason = ''
    if (-not $enabled) { $offlineReason = 'LiveData.Enabled = $false'; $messages.Add('LiveData.Enabled = $false: the internet is not used for OS support dates.'); $notes.Add('LiveData.Enabled = $false: the internet was not used for OS support dates.') }
    elseif ($Offline)  { $offlineReason = 'run with -Offline'; $messages.Add('-Offline: the internet is not used for OS support dates.'); $notes.Add('Run with -Offline: the internet was not used for OS support dates.') }

    # Snapshot (read once)
    $snapshot = $null
    if (Test-Path -LiteralPath $SnapshotPath) {
        try { $snapshot = ConvertFrom-LiveJsonText (Get-Content -LiteralPath $SnapshotPath -Raw -Encoding UTF8) }
        catch { $messages.Add(("Snapshot file {0} could not be read: {1}" -f $SnapshotPath, $_.Exception.Message)); $notes.Add('The shipped snapshot Config\LifecycleSnapshot.json could not be read.') }
    }
    else { $messages.Add(("Snapshot file not found: {0}" -f $SnapshotPath)); $notes.Add('The shipped snapshot Config\LifecycleSnapshot.json is missing.') }
    $snapshotDate = [string](Get-LiveProperty $snapshot 'generated_at')
    $snapshotProducts = Get-LiveProperty $snapshot 'products'
    $snapDay = $null
    if ($snapshotDate -match '^(\d{4}-\d{2}-\d{2})') {
        $tmp = [datetime]::MinValue
        if ([datetime]::TryParseExact($Matches[1], 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$tmp)) { $snapDay = $tmp.Date }
    }

    $shortErr = { param($t) $x = (([string]$t) -split "`n")[0].Trim(); if ($x.Length -gt 120) { $x = $x.Substring(0, 117) + '...' }; $x }
    $hostOf = { param($u) try { ([uri]$u).Host } catch { [string]$u } }
    $unreachable = @{}   # host -> short reason: after the first transport failure the host is not tried again in this run

    foreach ($product in (Get-LifecycleProductNames -Settings $Settings)) {
        $cacheFile = Join-Path $CacheDir ("eol_{0}.json" -f $product)
        $releases = @(); $source = ''; $when = ''
        $why = New-Object System.Collections.Generic.List[string]
        if (-not $online) { $why.Add($offlineReason) }

        # 1. Live: v1 API; the legacy API only after an HTTP 404 or an unexpected response (never after a timeout)
        if ($online) {
            $tryLegacy = $false
            $liveFail = ''
            foreach ($attempt in @(@{ Kind = 'v1'; Url = ("{0}/{1}" -f $v1Base, $product) }, @{ Kind = 'legacy'; Url = ("{0}/{1}.json" -f $oldBase, $product) })) {
                if ($attempt.Kind -eq 'legacy' -and -not $tryLegacy) { break }
                $url = $attempt.Url
                $h = & $hostOf $url
                if ($unreachable.ContainsKey($h)) { $liveFail = ('{0} not reachable ({1})' -f $h, $unreachable[$h]); break }
                Write-Host ("  Contacting {0} for {1} ..." -f $h, $product) -ForegroundColor DarkGray
                $r = Invoke-LiveHttpRequest -Url $url -TimeoutSec $timeout
                if ($null -eq $r.Text) {
                    $messages.Add(("Could not read {0}: {1}" -f $url, $r.Error))
                    if ($r.TransportFailure) {
                        $unreachable[$h] = (& $shortErr $r.Error)
                        $messages.Add(("{0} is not reachable - no further live calls to it in this run" -f $h))
                        $liveFail = ('{0} not reachable ({1})' -f $h, $unreachable[$h])
                        break
                    }
                    $liveFail = ('HTTP {0} from {1}' -f $r.StatusCode, $h)
                    $tryLegacy = ($r.StatusCode -eq 404)
                    continue
                }
                $rels = @()
                try { $rels = @(ConvertFrom-LifecycleJson -Product $product -Json (ConvertFrom-LiveJsonText $r.Text) -Source 'endoflife.date (live)') }
                catch { $messages.Add(("{0}: response is not valid JSON ({1})" -f $url, $_.Exception.Message)); $rels = @() }
                if ($rels.Count -gt 0) {
                    try { [System.IO.File]::WriteAllText($cacheFile, $r.Text, (New-Object System.Text.UTF8Encoding($false))) } catch { $messages.Add(("Cache write failed for {0}: {1}" -f $product, $_.Exception.Message)) }
                    $releases = $rels; $source = 'Live'; $when = (Get-Date).ToString('yyyy-MM-dd HH:mm')
                    break
                }
                $messages.Add(("{0}: no releases found in the response" -f $url))
                $liveFail = 'unexpected response format'
                $tryLegacy = $true
            }
            if ($source -ne 'Live' -and $liveFail) { $why.Add($liveFail) }
        }

        # Snapshot releases of this product (parsed lazily)
        $snapRels = $null
        $getSnap = {
            if ($null -eq $snapRels) {
                $snapRels = @()
                if ($null -ne $snapshotProducts) {
                    $snapJson = Get-LiveProperty $snapshotProducts $product
                    if ($null -ne $snapJson) { $snapRels = @(ConvertFrom-LifecycleJson -Product $product -Json $snapJson -Source 'offline snapshot') }
                }
            }
            $snapRels   # the caller collects it with @(& $getSnap)
        }

        # 2. Cache (fresh, then stale). A stale cache gives way to the shipped snapshot when the snapshot is newer,
        #    or when the cache has no release for a version-choice family that the snapshot has.
        if ($releases.Count -eq 0 -and (Test-Path -LiteralPath $cacheFile)) {
            try {
                $cacheTime = (Get-Item -LiteralPath $cacheFile).LastWriteTime
                $age = (Get-Date) - $cacheTime
                $json = ConvertFrom-LiveJsonText (Get-Content -LiteralPath $cacheFile -Raw -Encoding UTF8)
                $label = if ($age.TotalDays -le $cacheDays) { 'Cache' } else { 'StaleCache' }
                $rels = @(ConvertFrom-LifecycleJson -Product $product -Json $json -Source ("endoflife.date ({0})" -f $label.ToLowerInvariant()))
                $cacheWhen = $cacheTime.ToString('yyyy-MM-dd HH:mm')
                $useCache = $rels.Count -gt 0
                if ($useCache -and $label -eq 'StaleCache') {
                    $snapNow = @(& $getSnap)
                    $snapRels = $snapNow
                    if ($snapNow.Count -gt 0) {
                        if ($null -ne $snapDay -and $snapDay -gt $cacheTime.Date) {
                            $useCache = $false
                            $why.Add(('old cache from {0} replaced by the newer shipped snapshot ({1})' -f $cacheWhen, $snapshotDate))
                        }
                        elseif ((Get-LifecycleChoiceMatchGap -Settings $Settings -Product $product -Releases $rels -Reference $snapNow)) {
                            $useCache = $false
                            $why.Add(('old cache from {0} has no release for a version-choice family - shipped snapshot used' -f $cacheWhen))
                        }
                    }
                }
                if ($useCache) {
                    $releases = $rels; $source = $label; $when = $cacheWhen
                    if ($label -eq 'StaleCache') {
                        $messages.Add(("{0}: using cached dates from {1} ({2:N0} days old) - check them" -f $product, $when, $age.TotalDays))
                        $why.Add(('old cached copy from {0} ({1:N0} days) used - check' -f $when, $age.TotalDays))
                    }
                }
            }
            catch { $messages.Add(("Cache file for {0} could not be read: {1}" -f $product, $_.Exception.Message)); $why.Add('cache file unreadable'); $releases = @() }
        }

        # 3. Snapshot shipped with the tool
        if ($releases.Count -eq 0) {
            $snapNow = @(& $getSnap)
            if ($snapNow.Count -gt 0) { $releases = $snapNow; $source = 'Snapshot'; $when = $snapshotDate }
        }

        if ($releases.Count -gt 0) {
            $catalog.Products[$product]    = $releases
            $catalog.Source[$product]      = $source
            $catalog.RetrievedAt[$product] = $when
        }
        else {
            $catalog.Source[$product] = 'None'
            $why.Add('not in the cache or the shipped snapshot - settings-file dates used')
            $messages.Add(("{0}: no live, cached or snapshot dates - the settings-file dates are used" -f $product))
        }
        $catalog.Reason[$product] = (@($why | Where-Object { $_ }) -join '; ')
    }
    return $catalog
}

function Get-LifecycleChoiceMatchGap {
    # $true when a VersionChoices family of this product has no matching release in $Releases but has one in $Reference.
    param([hashtable]$Settings, [string]$Product, [object[]]$Releases, [object[]]$Reference)
    if (-not $Settings.ContainsKey('VersionChoices')) { return $false }
    foreach ($vc in @($Settings.VersionChoices)) {
        if (-not ($vc -is [hashtable]) -or ([string]$vc.Product).Trim().ToLowerInvariant() -ne $Product -or -not $vc.CyclePattern) { continue }
        $pat = [string]$vc.CyclePattern
        $inRel = @($Releases | Where-Object { [string]$_.Cycle -match $pat }).Count
        $inRef = @($Reference | Where-Object { [string]$_.Cycle -match $pat }).Count
        if ($inRel -eq 0 -and $inRef -gt 0) { return $true }
    }
    return $false
}

function Find-LifecycleRelease {
    # One release of a product by cycle name (case-insensitive), or $null.
    param($Catalog, [string]$Product, [string]$Cycle)
    if ($null -eq $Catalog -or -not $Catalog.Products.ContainsKey($Product)) { return $null }
    foreach ($r in $Catalog.Products[$Product]) {
        if ([string]::Equals([string]$r.Cycle, $Cycle, [System.StringComparison]::OrdinalIgnoreCase)) { return $r }
    }
    return $null
}

function Format-LifecycleMonthYear {
    param([string]$Date)
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact($Date, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) {
        return $d.ToString('MMM yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return $Date
}

function Update-OsLifecycleFromCatalog {
    <#
        Updates the OsLifecycle rows that carry a Live = '<product>:<cycle>' key with the catalogue dates.
          - EndOfSupport <- the release's end of support (eolFrom), or its extended end (eoesFrom) when the row
            sets LiveEndField = 'eoes' (used where the settings table counts free LTS as the support end, e.g. Debian)
          - ExtendedEnd  <- eoesFrom (not touched for LiveEndField = 'eoes' rows)
          - an empty live value never erases a settings date (endoflife.date does not track every paid extension)
          - Note: a pure date sentence ('ESU ended Jan 2023', 'Paid ESU ended Oct 2026') is rewritten when its date changed
          - every row gets DateSource ('endoflife.date (live yyyy-MM-dd HH:mm)', '... (cache ...)', 'offline snapshot ...', 'settings file')
        Returns the list of changes: [pscustomobject]@{ Label; Field; Old; New; Source }.
    #>
    param(
        [Parameter(Mandatory)] [hashtable]$Settings,
        $Catalog
    )
    $changes = New-Object System.Collections.Generic.List[object]
    $sourceText = @{ 'Live' = 'endoflife.date (live {0})'; 'Cache' = 'endoflife.date (cache {0})'; 'StaleCache' = 'endoflife.date (old cache {0})'; 'Snapshot' = 'offline snapshot {0} - verify' }
    $asOf = Get-Date
    foreach ($row in @($Settings.OsLifecycle)) {
        if (-not ($row -is [hashtable])) { continue }
        $row['DateSource'] = 'settings file'
        $row['LiveProduct'] = ''; $row['LiveFound'] = $false; $row['LiveRelease'] = ''
        $liveKey = Get-RowLiveKey $row
        if (-not $liveKey) { continue }
        $alts = @(Get-LiveKeyAlternatives $liveKey)
        if ($alts.Count -eq 0) { $changes.Add([pscustomobject]@{ Label = $row.Label; Field = 'Live'; Old = $row.Live; New = ''; Source = "Invalid Live key (expected '<product>:<cycle>', alternatives separated by '|')" }); continue }
        # first alternative found wins (endoflife.date may name a release e.g. '10-22h2' or '10-22h2-e')
        $rel = $null; $product = $alts[0].Product; $cycle = $alts[0].Cycle
        foreach ($a in $alts) {
            $hit = Find-LifecycleRelease -Catalog $Catalog -Product $a.Product -Cycle $a.Cycle
            if ($null -ne $hit) { $rel = $hit; $product = $a.Product; $cycle = $a.Cycle; break }
        }
        $row['LiveProduct'] = $product
        if ($null -eq $rel) {
            $inCatalog = @($alts | Where-Object { $null -ne $Catalog -and $Catalog.Products.ContainsKey($_.Product) }).Count -gt 0
            if ($inCatalog) {
                $tried = (@($alts | ForEach-Object { $_.Cycle }) -join "' / '")
                $changes.Add([pscustomobject]@{ Label = $row.Label; Field = 'Live'; Old = $row.Live; New = ''; Source = ("Release '{0}' not found in {1} - settings dates kept" -f $tried, $product) })
            }
            continue
        }
        $row['LiveFound'] = $true; $row['LiveRelease'] = [string]$rel.Cycle
        $src = $Catalog.Source[$product]; $when = $Catalog.RetrievedAt[$product]
        $fmt = if ($sourceText.ContainsKey($src)) { $sourceText[$src] } else { '{0}' }
        $srcLabel = $fmt -f $when

        $endField = 'eol'
        if ($row.ContainsKey('LiveEndField') -and ([string]$row.LiveEndField).Trim().ToLowerInvariant() -eq 'eoes') { $endField = 'eoes' }
        $newEnd = if ($endField -eq 'eoes') { $rel.ExtendedEnd } else { $rel.EndOfSupport }
        $newExt = if ($endField -eq 'eoes') { '' } else { $rel.ExtendedEnd }

        $changed = $false
        $oldExt = [string]$row.ExtendedEnd
        if ($newEnd -and $newEnd -ne [string]$row.EndOfSupport) {
            $changes.Add([pscustomobject]@{ Label = $row.Label; Field = 'EndOfSupport'; Old = [string]$row.EndOfSupport; New = $newEnd; Source = $srcLabel })
            $row['EndOfSupport'] = $newEnd; $changed = $true
        }
        if ($newExt -and $newExt -ne $oldExt) {
            $changes.Add([pscustomobject]@{ Label = $row.Label; Field = 'ExtendedEnd'; Old = $oldExt; New = $newExt; Source = $srcLabel })
            $row['ExtendedEnd'] = $newExt; $changed = $true
        }
        # Rewrite a note that is only a date sentence, so it cannot contradict the new date
        $note = [string]$row.Note
        $dateField = if ($endField -eq 'eoes') { [string]$row.EndOfSupport } else { [string]$row.ExtendedEnd }
        $d = [datetime]::MinValue
        $dateOk = [datetime]::TryParseExact($dateField, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)
        if ($changed -and $dateOk -and $note -match '^(?<what>(Paid )?(ESU|ELS|LTSS|LTS|Extended support))( \w+)* (ended|ends) (Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]* \d{4}(?<tail>.*)$') {
            $what = $Matches['what']; $tail = $Matches['tail']
            $verb = if ($d -lt $asOf) { 'ended' } else { 'ends' }
            $newNote = ('{0} {1} {2}{3}' -f $what, $verb, (Format-LifecycleMonthYear $dateField), $tail)
            if ($newNote -ne $note) {
                $changes.Add([pscustomobject]@{ Label = $row.Label; Field = 'Note'; Old = $note; New = $newNote; Source = $srcLabel })
                $row['Note'] = $newNote
            }
        }
        $row['DateSource'] = $srcLabel
    }
    return $changes.ToArray()
}

function Resolve-SupportCutoffDate {
    <#
        A date setting: a fixed date 'yyyy-MM-dd', or relative to the run date: 'today+<n>m' / 'today-<n>m' (months),
        'today+<n>y' / 'today-<n>y' (years), or 'today'. Used for SupportCutoffDate and VersionChoices.HideEndedBefore.
        Returns 'yyyy-MM-dd'. Throws on any other text, so a typo cannot silently switch a rule off.
    #>
    param([string]$Value, [datetime]$AsOf = (Get-Date))
    $v = ([string]$Value).Trim()
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact($v, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) {
        return $d.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    if ($v -match '^(?i)today\s*([+-])\s*(\d+)\s*([my])$') {
        $n = [int]$Matches[2]; if ($Matches[1] -eq '-') { $n = -$n }
        $d = if ($Matches[3].ToLowerInvariant() -eq 'm') { $AsOf.Date.AddMonths($n) } else { $AsOf.Date.AddYears($n) }
        return $d.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    if ($v -match '^(?i)today$') { return $AsOf.Date.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture) }
    throw "Setting SupportCutoffDate '$Value' is not valid: use 'yyyy-MM-dd' (e.g. '2027-12-31') or 'today+<n>m' / 'today-<n>m' (e.g. 'today+12m')"
}

function Get-VmImportDocText {
    # Visible text of the AWS docs page, reduced to the main article so menus/footers do not change the hash.
    param([string]$Html)
    $h = $Html
    foreach ($marker in @('id="main-col-body"', '<main', 'id="main-content"')) {
        $i = $h.IndexOf($marker, [System.StringComparison]::OrdinalIgnoreCase)
        if ($i -ge 0) {
            # Start at the '<' of the tag that holds the marker, so the rest of that tag is stripped as markup
            $lt = $h.LastIndexOf('<', $i)
            if ($lt -ge 0) { $i = $lt }
            $h = $h.Substring($i); break
        }
    }
    # ... and stop at the end of the article, so footer / feedback / copyright changes do not change the hash
    $end = -1
    foreach ($marker in @('id="main-col-footer"', 'Document Conventions', 'Did this page help you?', '</main>', 'id="js_error_message"')) {
        $j = $h.IndexOf($marker, [System.StringComparison]::OrdinalIgnoreCase)
        if ($j -gt 0 -and ($end -lt 0 -or $j -lt $end)) { $end = $j }
    }
    if ($end -gt 0) {
        # cut at the '<' of the tag holding the marker when the marker is inside a tag
        $lt = $h.LastIndexOf('<', $end); $gt = $h.LastIndexOf('>', $end)
        if ($lt -gt $gt) { $end = $lt }
        $h = $h.Substring(0, $end)
    }
    $h = [regex]::Replace($h, '(?is)<(script|style|noscript)[^>]*>.*?</\1>', ' ')
    $h = [regex]::Replace($h, '(?s)<[^>]+>', ' ')
    $h = [System.Net.WebUtility]::HtmlDecode($h)
    $h = [regex]::Replace($h, '\s+', ' ').Trim()
    return $h
}

function Test-VmImportDocChanged {
    <#
        The AWS VM Import rules (supported OS list, 16 TiB / 8 TiB disk limits, 21 volumes, no 32-bit) are only
        published on a docs page - there is no API. This hashes the page text (SHA-256) and compares it with
        LiveData.VmImportDocKnownHash. Status: Unchanged | Changed | NotRecorded | NotChecked.
        A change never alters the settings: it tells you to review the VM Import settings.
    #>
    param([Parameter(Mandatory)] [hashtable]$Settings, [switch]$Offline)
    $url   = [string](Get-LiveDataSetting $Settings 'VmImportDocUrl')
    $known = ([string](Get-LiveDataSetting $Settings 'VmImportDocKnownHash')).Trim().ToLowerInvariant()
    $res = [pscustomobject]@{ Status = 'NotChecked'; Hash = ''; Url = $url; Detail = ''; CheckedAt = '' }
    if (-not [bool](Get-LiveDataSetting $Settings 'Enabled') -or $Offline -or [string]::IsNullOrWhiteSpace($url)) {
        $res.Detail = 'Not checked (offline). Review the VM Import settings against the AWS page before each wave.'
        return $res
    }
    $msgs = New-Object System.Collections.Generic.List[string]
    $html = Invoke-LiveHttpGet -Url $url -TimeoutSec ([int](Get-LiveDataSetting $Settings 'TimeoutSec')) -Messages $msgs
    if ($null -eq $html) {
        $res.Detail = ('Not checked: {0}' -f ($msgs -join '; '))
        return $res
    }
    $text = Get-VmImportDocText -Html $html
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($text)) } finally { $sha.Dispose() }
    $hash = (($bytes | ForEach-Object { $_.ToString('x2') }) -join '')
    $res.Hash = $hash
    $res.CheckedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm')

    # Facts the settings depend on - reported so a reviewer can see at a glance what the page still says
    $facts = New-Object System.Collections.Generic.List[string]
    foreach ($f in @(@('16 TiB', '16 TiB'), @('8 TiB', '8 TiB'), @('21 volumes', '\b21\b'), @('32-bit', '32-bit|i386'))) {
        $found = if ($text -match $f[1]) { 'found' } else { 'NOT found' }
        $facts.Add(("'{0}' {1}" -f $f[0], $found))
    }
    $factText = 'Page mentions: ' + ($facts -join ', ')
    if ([string]::IsNullOrWhiteSpace($known)) {
        $res.Status = 'NotRecorded'
        $res.Detail = ("No reviewed hash recorded yet. After checking the VM Import settings against the page, set LiveData.VmImportDocKnownHash = '{0}'. {1}" -f $hash, $factText)
    }
    elseif ($known -eq $hash) {
        $res.Status = 'Unchanged'
        $res.Detail = ('Page unchanged since the VM Import settings were last reviewed. {0}' -f $factText)
    }
    else {
        $res.Status = 'Changed'
        $res.Detail = ("The AWS VM Import page has changed since the settings were reviewed: check VmImportSupportedOs, VmImportUnsupportedOs, MaxEbsVolumeGiB, ManifestImportGiB, MaxVolumesPerImport and Allow32BitImport, then set LiveData.VmImportDocKnownHash = '{0}'. {1}" -f $hash, $factText)
    }
    return $res
}

function Invoke-AwsCliJson {
    # Runs the AWS CLI and returns parsed JSON, or $null (reason added to $Messages).
    param([string[]]$Arguments, [System.Collections.Generic.List[string]]$Messages)
    try {
        $out = & aws @Arguments --output json 2>&1
        if ($LASTEXITCODE -ne 0) {
            $Messages.Add(("aws {0} failed: {1}" -f ($Arguments[0..1] -join ' '), (($out | Out-String).Trim() -split "`n")[0]))
            return $null
        }
        $text = ($out | Where-Object { $_ -is [string] }) -join "`n"
        if ([string]::IsNullOrWhiteSpace($text)) { return $null }
        return ($text | ConvertFrom-Json)
    }
    catch {
        $Messages.Add(("aws {0} failed: {1}" -f ($Arguments[0..1] -join ' '), $_.Exception.Message))
        return $null
    }
}

function Get-AwsLiveFacts {
    <#
        Optional facts from AWS APIs (switch -UseAwsApis on the main script). Needs the AWS CLI and read-only credentials:
          ec2:DescribeInstanceTypes, ssm:GetParametersByPath on the public /aws/service/* parameters.
          - InstanceTypes  : @{ '<type>' = @{ MaxENI = <int>; Architectures = @('x86_64', ...) } }
          - RegionServices : service ids AWS lists for the region (SSM /aws/service/global-infrastructure/regions/<r>/services)
          - WindowsAmis    : public Windows AMI parameter names (SSM /aws/service/ami-windows-latest)
        Without -Enabled, or without the CLI / credentials, everything stays empty and the built-in tables are used.
    #>
    param([string]$Region = 'eu-central-1', [switch]$Enabled)
    $msgs = New-Object System.Collections.Generic.List[string]
    $facts = [pscustomobject]@{
        Enabled        = [bool]$Enabled
        Region         = $Region
        InstanceTypes  = @{}
        RegionServices = @()
        WindowsAmis    = @()
        Status         = @{ InstanceTypes = 'Not checked'; RegionServices = 'Not checked'; WindowsAmis = 'Not checked' }
        Messages       = $msgs
    }
    if (-not $Enabled) { $msgs.Add('AWS APIs not used (run with -UseAwsApis and AWS CLI credentials to use them).'); return $facts }
    if (-not (Get-Command aws -ErrorAction SilentlyContinue)) { $msgs.Add('-UseAwsApis: the AWS CLI (aws) is not installed - built-in tables used.'); return $facts }

    $it = Invoke-AwsCliJson -Arguments @('ec2', 'describe-instance-types', '--region', $Region, '--query', 'InstanceTypes[].{T:InstanceType,E:NetworkInfo.MaximumNetworkInterfaces,A:ProcessorInfo.SupportedArchitectures}') -Messages $msgs
    if ($null -ne $it) {
        foreach ($t in @($it)) {
            $name = [string](Get-LiveProperty $t 'T')
            if (-not $name) { continue }
            $eni = 0; [void][int]::TryParse([string](Get-LiveProperty $t 'E'), [ref]$eni)
            $facts.InstanceTypes[$name] = @{ MaxENI = $eni; Architectures = @(Get-LiveProperty $t 'A' | ForEach-Object { [string]$_ }) }
        }
        $facts.Status.InstanceTypes = ('Live ({0} instance types)' -f $facts.InstanceTypes.Count)
    }
    else { $facts.Status.InstanceTypes = 'Failed - built-in table used' }

    $svc = Invoke-AwsCliJson -Arguments @('ssm', 'get-parameters-by-path', '--region', $Region, '--path', ("/aws/service/global-infrastructure/regions/{0}/services" -f $Region), '--query', 'Parameters[].Value') -Messages $msgs
    if ($null -ne $svc) { $facts.RegionServices = @($svc | ForEach-Object { [string]$_ }); $facts.Status.RegionServices = ('Live ({0} services)' -f $facts.RegionServices.Count) }
    else { $facts.Status.RegionServices = 'Failed - not checked' }

    $ami = Invoke-AwsCliJson -Arguments @('ssm', 'get-parameters-by-path', '--region', $Region, '--path', '/aws/service/ami-windows-latest', '--query', 'Parameters[].Name') -Messages $msgs
    if ($null -ne $ami) { $facts.WindowsAmis = @($ami | ForEach-Object { [string]$_ }); $facts.Status.WindowsAmis = ('Live ({0} AMI names)' -f $facts.WindowsAmis.Count) }
    else { $facts.Status.WindowsAmis = 'Failed - not checked' }

    return $facts
}

function Get-WindowsAmiPattern {
    <#
        Regex that finds the public Windows Server AMI parameter names of one version, e.g. names under
        /aws/service/ami-windows-latest/ : 'Windows_Server-2012-R2_RTM-English-64Bit-Base', 'Windows_Server-2012-RTM-English-64Bit-Base',
        'Windows_Server-2016-English-Full-Base', 'Windows_Server-2022-English-Full-Base', 'Windows_Server-2008-R2_SP1-English-64Bit-Base'.
        2012 must not match the 2012 R2 names (and 2008 not 2008 R2). Version text: '2008', '2008 R2', '2012', '2012 R2', '2016' ...
    #>
    param([string]$Version)
    $v = ([string]$Version).Trim() -replace '_', ' '
    switch -Regex ($v) {
        '^2012 R2$' { return 'Windows_Server-2012-R2_' }
        '^2012$'    { return 'Windows_Server-2012-RTM-' }
        '^2008 R2$' { return 'Windows_Server-2008-R2_' }
        '^2008$'    { return 'Windows_Server-2008-SP2-' }
        '^(\d{4})$' { return ('Windows_Server-{0}-' -f $Matches[1]) }
        default     { return $null }
    }
}

function Test-WindowsAmiOffered {
    # $true when a standard Windows Server AMI of this version is in the SSM name list (see Get-WindowsAmiPattern).
    param([object[]]$AmiNames, [string]$Version)
    $pat = Get-WindowsAmiPattern -Version $Version
    if (-not $pat) { return $false }
    $rx = [regex]::Escape($pat)
    return (@($AmiNames | Where-Object { [string]$_ -match $rx }).Count -gt 0)
}

function Get-AwsTargetServiceCheck {
    # Which AWS target services the region offers (needs Get-AwsLiveFacts with -Enabled). Returns text for Data_Sources.
    param($AwsFacts)
    if ($null -eq $AwsFacts -or @($AwsFacts.RegionServices).Count -eq 0) { return 'Not checked (needs -UseAwsApis)' }
    $wanted = [ordered]@{ 'workspaces' = 'Amazon WorkSpaces'; 'rds' = 'Amazon RDS'; 'docdb' = 'Amazon DocumentDB'; 'evs' = 'Amazon EVS'; 'mgn' = 'AWS Application Migration Service' }
    $parts = foreach ($k in $wanted.Keys) {
        $ok = if (@($AwsFacts.RegionServices) -contains $k) { 'offered' } else { 'NOT listed' }
        '{0}: {1}' -f $wanted[$k], $ok
    }
    return ($parts -join '; ')
}

function Get-PricingFileInfo {
    # Publication date and local age of an AWS bulk pricing CSV (first lines of the file).
    param([string]$CsvPath)
    $info = [pscustomobject]@{ Path = $CsvPath; PublicationDate = ''; FileDate = ''; AgeDays = $null }
    if ([string]::IsNullOrWhiteSpace($CsvPath) -or -not (Test-Path -LiteralPath $CsvPath)) { return $info }
    $item = Get-Item -LiteralPath $CsvPath
    $info.FileDate = $item.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
    $info.AgeDays = [math]::Round(((Get-Date) - $item.LastWriteTime).TotalDays, 1)
    try {
        $reader = New-Object System.IO.StreamReader($CsvPath)
        try {
            for ($i = 0; $i -lt 10; $i++) {
                $line = $reader.ReadLine(); if ($null -eq $line) { break }
                if ($line -match '^"?Publication Date"?\s*,\s*"?([^"]+)"?') { $info.PublicationDate = $Matches[1]; break }
            }
        } finally { $reader.Close() }
    } catch { }
    return $info
}

function New-DataSourceRows {
    <#
        Rows for the workbook sheet Data_Sources: where every value of this run came from.
        Shape: [pscustomobject]@{ Area; Source; Status; RetrievedAt; Detail }
    #>
    param(
        $LifecycleCatalog,
        [object[]]$LifecycleChanges = @(),
        $VmImportCheck,
        $AwsFacts,
        [object[]]$PricingFiles = @(),
        [string]$CutoffSetting,
        [string]$CutoffResolved,
        [int]$ArmTypesExcluded = -1,
        [hashtable]$Settings
    )
    $rows = New-Object System.Collections.Generic.List[object]
    $add = { param($a, $s, $st, $t, $d) $rows.Add([pscustomobject]@{ Area = $a; Source = $s; Status = $st; RetrievedAt = $t; Detail = $d }) }
    $statusText = @{ 'Live' = 'Live'; 'Cache' = 'Cached copy (recent)'; 'StaleCache' = 'Cached copy (OLD - check)'; 'Snapshot' = 'Offline snapshot (verify)'; 'None' = 'Not available - settings dates used' }

    # OS support dates: ONE row per product (status + short reason); full error texts are printed on the console only
    if ($null -ne $LifecycleCatalog) {
        $reasons = @{}
        if ($LifecycleCatalog.PSObject.Properties['Reason'] -and $LifecycleCatalog.Reason -is [hashtable]) { $reasons = $LifecycleCatalog.Reason }
        $products = @(@($LifecycleCatalog.Source.Keys) + @($LifecycleCatalog.Products.Keys) | Sort-Object -Unique)
        foreach ($p in $products) {
            $src = [string]$LifecycleCatalog.Source[$p]
            $st = if ($statusText.ContainsKey($src)) { $statusText[$src] } else { $src }
            $srcName = switch ($src) { 'Snapshot' { 'Config\LifecycleSnapshot.json' } 'None' { 'settings file' } default { 'endoflife.date API' } }
            $parts = New-Object System.Collections.Generic.List[string]
            if ($LifecycleCatalog.Products.ContainsKey($p)) { $parts.Add(('{0} release(s).' -f @($LifecycleCatalog.Products[$p]).Count)) }
            if ($reasons.ContainsKey($p) -and $reasons[$p]) { $parts.Add(('Why not live: {0}.' -f $reasons[$p])) }
            if ($Settings) {
                $rowsP = @($Settings.OsLifecycle | Where-Object { $_ -is [hashtable] -and $_['LiveProduct'] -eq $p })
                $got  = @($rowsP | Where-Object { $_['LiveFound'] } | ForEach-Object { $_.Label })
                $kept = @($rowsP | Where-Object { -not $_['LiveFound'] } | ForEach-Object { $_.Label })
                if ($got.Count -gt 0)  { $parts.Add(('Dates taken from this source for: {0}.' -f ($got -join ', '))) }
                if ($kept.Count -gt 0) { $parts.Add(('Release not found here, settings-file dates kept for: {0}.' -f ($kept -join ', '))) }
            }
            if ($p -eq 'windows' -or $p -eq 'windows-server') { $parts.Add('Also the source of the Exact_OS_Version choices.') }
            & $add ('OS support dates - ' + $p) $srcName $st $LifecycleCatalog.RetrievedAt[$p] ($parts -join ' ')
        }
        $notes = @()
        if ($LifecycleCatalog.PSObject.Properties['Notes']) { $notes = @($LifecycleCatalog.Notes) }
        foreach ($m in $notes) { & $add 'OS support dates - note' 'endoflife.date API' 'Note' '' $m }
    }
    foreach ($c in @($LifecycleChanges)) {
        if ($null -eq $c -or $c.Field -eq 'Live') { continue }   # 'not found' is part of the product row above
        & $add 'OS support dates - updated' $c.Source 'Changed' '' ('{0} {1}: {2} -> {3}' -f $c.Label, $c.Field, $(if ($c.Old) { $c.Old } else { '(none)' }), $c.New)
    }
    if ($Settings) {
        $static = @($Settings.OsLifecycle | Where-Object { $_ -is [hashtable] -and $_['DateSource'] -eq 'settings file' } | ForEach-Object { $_.Label })
        if ($static.Count -gt 0) { & $add 'OS support dates - settings file' 'Config\AWSMigrationStrategySettings.psd1' 'Static (no live source)' '' ('Dates typed in the settings file for: ' + ($static -join ', ')) }
    }

    if ($CutoffSetting) {
        & $add 'Support cut-off date (rule R06)' 'settings file (SupportCutoffDate)' $(if ($CutoffSetting -eq $CutoffResolved) { 'Fixed date' } else { 'Calculated from the run date' }) (Get-Date).ToString('yyyy-MM-dd') ('Setting: {0} -> used: {1}' -f $CutoffSetting, $CutoffResolved)
    }

    foreach ($pf in @($PricingFiles)) {
        if ($null -eq $pf) { continue }
        $st = if ($null -eq $pf.AgeDays) { 'Missing' } elseif ($pf.AgeDays -lt 0.05) { 'Live (downloaded now)' } else { ('Cached copy ({0} days old)' -f $pf.AgeDays) }
        & $add ('EC2 / EBS prices and instance sizes - ' + $pf.Region) 'AWS bulk pricing file (pricing.us-east-1.amazonaws.com)' $st $pf.FileDate ('AWS publication date: {0}' -f $(if ($pf.PublicationDate) { $pf.PublicationDate } else { 'unknown' }))
    }
    if ($ArmTypesExcluded -ge 0) {
        & $add 'CPU architecture (instance choice)' 'AWS pricing file column Physical Processor' 'Live' '' ('{0} ARM (Graviton) instance types excluded: VMware VMs are x86 and an imported x86 image cannot boot on ARM.' -f $ArmTypesExcluded)
    }

    if ($null -ne $AwsFacts) {
        $t = if ($AwsFacts.Enabled) { (Get-Date).ToString('yyyy-MM-dd HH:mm') } else { '' }
        $eniSrc = if ($AwsFacts.InstanceTypes.Count -gt 0) { 'AWS API ec2:DescribeInstanceTypes' } else { 'built-in table in InstanceSizing.ps1 (approximate)' }
        & $add 'Max network interfaces per instance' $eniSrc $AwsFacts.Status.InstanceTypes $t $(if ($AwsFacts.InstanceTypes.Count -gt 0) { ('Exact limits for {0} instance types in {1}.' -f $AwsFacts.InstanceTypes.Count, $AwsFacts.Region) } else { 'Approximate limits by instance size - run with -UseAwsApis for exact values.' })
        & $add ('AWS services offered in ' + $AwsFacts.Region) 'AWS SSM public parameters' $AwsFacts.Status.RegionServices $t (Get-AwsTargetServiceCheck -AwsFacts $AwsFacts)
        $amiDetail = 'Not checked (needs -UseAwsApis)'
        if (@($AwsFacts.WindowsAmis).Count -gt 0) {
            $vers = foreach ($v in @('2012', '2012 R2', '2016', '2019', '2022', '2025')) { '{0}: {1}' -f $v, $(if (Test-WindowsAmiOffered -AmiNames @($AwsFacts.WindowsAmis) -Version $v) { 'offered' } else { 'NOT offered' }) }
            $amiDetail = 'Standard Windows Server AMIs - ' + ($vers -join '; ')
        }
        & $add 'Windows Server AMIs' 'AWS SSM public parameters' $AwsFacts.Status.WindowsAmis $t $amiDetail
        foreach ($m in $AwsFacts.Messages) { & $add 'AWS APIs - note' 'AWS CLI' 'Note' '' $m }
    }

    if ($null -ne $VmImportCheck) {
        $st = switch ($VmImportCheck.Status) { 'Unchanged' { 'Live - unchanged' } 'Changed' { 'Live - CHANGED: review settings' } 'NotRecorded' { 'Live - no reviewed hash yet' } default { 'Not checked' } }
        & $add 'AWS VM Import rules (OS list, disk limits, 32-bit)' $(if ($VmImportCheck.Url) { $VmImportCheck.Url } else { 'AWS docs page' }) $st $VmImportCheck.CheckedAt $VmImportCheck.Detail
    }
    & $add 'VM Import OS list and disk limits used by the rules' 'settings file (AWS publishes no API for these)' 'Static' '' 'VmImportSupportedOs, VmImportUnsupportedOs, MaxEbsVolumeGiB, ManifestImportGiB, MaxVolumesPerImport, Allow32BitImport - review when the VM Import page changes.'
    return $rows.ToArray()
}

#---------------------------------------------------------------------
# LIVE DATE IMPACT - which VMs got a different result because a support date differs from the settings file
#---------------------------------------------------------------------
function Copy-OsLifecycleRows {
    # Copy of the OsLifecycle rows as typed in the settings file, taken BEFORE live dates are applied
    param([Parameter(Mandatory)] [hashtable]$Settings)
    $copy = @()
    foreach ($r in @($Settings.OsLifecycle)) {
        if ($r -is [hashtable]) { $copy += , ($r.Clone()) }
        else { $h = @{}; foreach ($p in $r.PSObject.Properties) { $h[$p.Name] = $p.Value }; $copy += , $h }
    }
    return , $copy
}

function Format-ImpactDate {
    # 'yyyy-MM-dd' -> '31 Dec 2027'; '' -> 'none'
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return 'none' }
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact($Value.Trim(), 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) {
        return $d.ToString('d MMM yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return $Value
}

function Get-ImpactDatePosition {
    # Where a support end date sits against the two lines the rules use: the run date and the cut-off
    param([string]$Value, [datetime]$AsOf, [string]$Cutoff)
    if ([string]::IsNullOrWhiteSpace($Value)) { return 'no date' }
    $d = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($Value.Trim(), 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) { return 'no date' }
    if ($d -lt $AsOf.Date) { return 'already passed' }
    $c = [datetime]::MinValue
    if (-not [string]::IsNullOrWhiteSpace($Cutoff) -and
        [datetime]::TryParseExact($Cutoff.Trim(), 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$c) -and $d -lt $c) {
        return 'before the cut-off'
    }
    return 'after the cut-off'
}

function Compare-LiveDateImpact {
    <#
        Compares the strategy rows of this run (support dates from live data / cache / snapshot) with the same rows worked
        out with the dates typed in the settings file. Returns one row per VM and OS field whose OS status, strategy, rule,
        AWS target or readiness differs - i.e. the VMs whose result was changed by a newer support date.
        Rows are matched by position (both runs use the same inventory in the same order), never by name, so duplicate
        VM names cannot be mixed up. Exact-version picks in the workbook use the picked version's own dates and are not covered.
    #>
    param(
        [Parameter(Mandatory)] $UsedRowsByField,        # ordered hashtable: 'VMwareTools' / 'ConfigFile' -> rows of this run
        [Parameter(Mandatory)] $SettingsRowsByField,    # same, worked out with the settings-file dates
        [string]$DefaultOSField = 'VMwareTools',
        [datetime]$AsOf = (Get-Date),
        [string]$Cutoff = ''
    )
    $label = @{ VMwareTools = 'VMware Tools'; ConfigFile = 'Configuration file' }
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($field in @('VMwareTools', 'ConfigFile')) {
        if (-not $UsedRowsByField.Contains($field) -or -not $SettingsRowsByField.Contains($field)) { continue }
        $used = @($UsedRowsByField[$field]); $base = @($SettingsRowsByField[$field])
        if ($used.Count -ne $base.Count) {
            Write-Warning "Live date impact check skipped for $($label[$field]): the two runs have a different number of VMs."
            continue
        }
        for ($i = 0; $i -lt $used.Count; $i++) {
            $u = $used[$i]; $b = $base[$i]
            if ([string]$u.VMName -ne [string]$b.VMName) {
                Write-Warning "Live date impact check skipped for $($label[$field]): VM order differs between the two runs."
                break
            }
            $differs = $false
            foreach ($col in @('OS_Status', 'Migration_Strategy', 'Rule_Applied', 'AWS_Target', 'Readiness')) {
                if ([string]$u.$col -ne [string]$b.$col) { $differs = $true; break }
            }
            if (-not $differs) { continue }

            $endB = [string]$b.OS_EndOfSupport; $endU = [string]$u.OS_EndOfSupport
            $extB = if ($b.PSObject.Properties['OS_Extended_Support_End']) { [string]$b.OS_Extended_Support_End } else { '' }
            $extU = if ($u.PSObject.Properties['OS_Extended_Support_End']) { [string]$u.OS_Extended_Support_End } else { '' }
            $why = @()
            if ($endB -ne $endU) {
                $pb = Get-ImpactDatePosition -Value $endB -AsOf $AsOf -Cutoff $Cutoff
                $pu = Get-ImpactDatePosition -Value $endU -AsOf $AsOf -Cutoff $Cutoff
                $t = "Support end changed from $(Format-ImpactDate $endB) to $(Format-ImpactDate $endU)"
                if ($pb -ne $pu) { $t += " (was $pb, now $pu" + $(if ($Cutoff) { "; cut-off $(Format-ImpactDate $Cutoff))" } else { ')' }) }
                $why += $t
            }
            if ($extB -ne $extU) { $why += "Paid extended support end changed from $(Format-ImpactDate $extB) to $(Format-ImpactDate $extU)" }
            if ($why.Count -eq 0) { $why += 'The OS lifecycle row changed (label or note) - compare OS_Status' }
            $strategyChanged = ([string]$u.Migration_Strategy -ne [string]$b.Migration_Strategy)

            $out.Add([pscustomobject][ordered]@{
                VMName                       = [string]$u.VMName
                OS_Field                     = $label[$field]
                Default_OS_Field             = $(if ($field -eq $DefaultOSField) { 'Yes' } else { 'No' })
                OS                           = [string]$u.OS_Used
                Strategy_Changed             = $(if ($strategyChanged) { 'Yes' } else { 'No - status / target / readiness only' })
                Strategy_With_Settings_Dates = ("{0} ({1})" -f $b.Migration_Strategy, $b.Rule_Applied)
                Strategy_Now                 = ("{0} ({1})" -f $u.Migration_Strategy, $u.Rule_Applied)
                Support_End                  = ("{0} -> {1}" -f (Format-ImpactDate $endB), (Format-ImpactDate $endU))
                Extended_Support_End         = ("{0} -> {1}" -f (Format-ImpactDate $extB), (Format-ImpactDate $extU))
                OS_Status                    = ("{0} -> {1}" -f $b.OS_Status, $u.OS_Status)
                AWS_Target                   = $(if ([string]$u.AWS_Target -ne [string]$b.AWS_Target) { "{0} -> {1}" -f $b.AWS_Target, $u.AWS_Target } else { "unchanged: $($u.AWS_Target)" })
                Readiness                    = $(if ([string]$u.Readiness -ne [string]$b.Readiness) { "{0} -> {1}" -f $b.Readiness, $u.Readiness } else { "unchanged: $($u.Readiness)" })
                Why                          = ($why -join '; ')
            })
        }
    }
    return $out.ToArray()
}

function New-LiveDateImpactSourceRow {
    # One Data_Sources row summarising the live date impact check
    param([object[]]$ImpactRows = @(), [int]$DatesChanged = 0, [string]$DefaultOSField = 'VMwareTools', [switch]$Skipped)
    $label = @{ VMwareTools = 'VMware Tools'; ConfigFile = 'Configuration file' }
    $rows = @($ImpactRows | Where-Object { $null -ne $_ })
    $def = @($rows | Where-Object { $_.Default_OS_Field -eq 'Yes' })
    $defStrat = @($def | Where-Object { $_.Strategy_Changed -eq 'Yes' })
    if ($Skipped) { $status = 'Not checked'; $detail = 'The comparison could not be made - see the console warnings.' }
    elseif ($DatesChanged -eq 0) { $status = 'No change'; $detail = 'No support date differs from the settings file, so no VM result can differ.' }
    elseif ($rows.Count -eq 0) { $status = 'No VM changed'; $detail = "$DatesChanged support date(s) differ from the settings file, but no VM's OS status, strategy, target or readiness changed." }
    else {
        $status = "$($def.Count) VM(s) changed"
        $detail = ("With the default OS field ({0}): {1} VM(s) changed, {2} of them changed strategy. Including the other OS field: {3} row(s). Every VM is listed on sheet Live_Date_Changes." -f $label[$DefaultOSField], $def.Count, $defStrat.Count, $rows.Count)
    }
    return [pscustomobject]@{
        Area        = 'Live dates - effect on the results'
        Source      = 'this run compared with the dates typed in the settings file'
        Status      = $status
        RetrievedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
        Detail      = $detail
    }
}
