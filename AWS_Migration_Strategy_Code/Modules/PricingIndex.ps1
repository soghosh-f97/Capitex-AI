#---------------------------------------------------------------------
# AWS PRICING FILE DOWNLOAD
#---------------------------------------------------------------------
function Get-AWSBulkPricingFile {
    param(
        [string]$ServiceCode = "AmazonEC2",
        [string]$RegionCode,
        [string]$CacheDir = "$env:TEMP\AWSPricingCache"
    )
 
    if (-not (Test-Path $CacheDir)) { New-Item -ItemType Directory -Path $CacheDir | Out-Null }
    $cacheFile = Join-Path $CacheDir "$ServiceCode`_$RegionCode.csv"
 
    if (Test-Path $cacheFile) {
        $age = (Get-Date) - (Get-Item $cacheFile).LastWriteTime
        if ($age.TotalDays -lt 7) {
            Write-Host "Using cached pricing file for $RegionCode (age: $([math]::Round($age.TotalHours,1))h)" -ForegroundColor DarkGray
            return $cacheFile
        }
    }
 
    $url = "https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/$ServiceCode/current/$RegionCode/index.csv"
    Write-Host "Downloading pricing data for $RegionCode (one-time, may take a few minutes)..." -ForegroundColor Yellow
 
    try {
        Invoke-WebRequest -Uri $url -OutFile $cacheFile -UseBasicParsing
        Write-Host "Download complete: $cacheFile" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to download pricing file for $RegionCode : $_"
        return $null
    }
 
    return $cacheFile
}
 
#---------------------------------------------------------------------
# STREAMING INDEX BUILDER (specs + OnDemand + 3yr Reserved pricing)
#---------------------------------------------------------------------
function Build-AWSPricingIndex {
    param([string]$CsvPath)
 
    Write-Host "Building pricing index from $(Split-Path $CsvPath -Leaf) (single streaming pass, this takes a few minutes)..." -ForegroundColor Yellow
 
    $streamReader = New-Object System.IO.StreamReader($CsvPath)
    $headerLine = $null
    while (($line = $streamReader.ReadLine()) -ne $null) {
        if ($line -match '^"?SKU"?,') { $headerLine = $line; break }
    }
 
    if (-not $headerLine) {
        Write-Error "Could not locate header row (line starting with 'SKU') in $CsvPath"
        $streamReader.Close()
        return $null
    }
 
    $headers = $headerLine -split '","' | ForEach-Object { $_.Trim('"') }
 
    $colIndex = @{}
    for ($i = 0; $i -lt $headers.Count; $i++) { $colIndex[$headers[$i]] = $i }
 
    $requiredCols = @('Product Family','Operating System','Tenancy','CapacityStatus','Pre Installed S/W','Current Generation','TermType','Instance Type','vCPU','Memory','Instance Family','PricePerUnit','Unit','Volume API Name','Group','LeaseContractLength','PurchaseOption','OfferingClass','PriceDescription')
    foreach ($c in $requiredCols) {
        if (-not $colIndex.ContainsKey($c)) { Write-Warning "Column '$c' not found in CSV header!" }
    }
 
    $maxColIndex = ($colIndex.Values | Measure-Object -Maximum).Maximum
 
    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($streamReader)
    $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $parser.SetDelimiters(",")
    $parser.HasFieldsEnclosedInQuotes = $true
 
    $instanceSpecMap = @{}
    $ec2PriceMap = @{}
    $ec2RIPriceMap = @{}
    $ebsGp3Price = $null
    $ebsGp3IOPSPrice = $null
    $ebsGp3ThroughputPrice = $null
    $ebsIo2Price = $null
    $ebsIo2IOPSPrice = $null
    $rowCount = 0
 
    while (-not $parser.EndOfData) {
        $fields = $parser.ReadFields()
        $rowCount++
        if ($rowCount % 300000 -eq 0) { Write-Host "  ...processed $rowCount rows" -ForegroundColor DarkGray }
 
        if ($fields.Count -le $maxColIndex) { continue }
 
        $productFamily = $fields[$colIndex['Product Family']]
 
        if ($productFamily -eq 'Compute Instance') {
            $termType   = $fields[$colIndex['TermType']]
            $tenancy    = $fields[$colIndex['Tenancy']]
            $capStatus  = $fields[$colIndex['CapacityStatus']]
            $preInstall = $fields[$colIndex['Pre Installed S/W']]
            $os         = $fields[$colIndex['Operating System']]
            $instType   = $fields[$colIndex['Instance Type']]
            $currentGen = $fields[$colIndex['Current Generation']]
 
            if ($tenancy -ne 'Shared' -or $capStatus -ne 'Used' -or $preInstall -ne 'NA') { continue }
            if ([string]::IsNullOrWhiteSpace($instType)) { continue }
 
            if ($termType -eq 'OnDemand') {
    $priceDesc = $fields[$colIndex['PriceDescription']]
    $isExcludedVariant = $priceDesc -match 'BYOL' -or $priceDesc -match 'without licenses' -or $priceDesc -match 'No License required' -or $priceDesc -match 'SQL Server'
 
    if (-not $isExcludedVariant) {
        $priceKey = "$instType|$os"
        $priceVal = $fields[$colIndex['PricePerUnit']]
        if ($priceVal) {
            $priceValDouble = [double]$priceVal
            if (-not $ec2PriceMap.ContainsKey($priceKey) -or $priceValDouble -lt $ec2PriceMap[$priceKey]) {
                $ec2PriceMap[$priceKey] = $priceValDouble
            }
        }
    }
 
                if ($os -eq 'Linux' -and $currentGen -eq 'Yes' -and -not $instanceSpecMap.ContainsKey($instType)) {
                    $vcpu = $fields[$colIndex['vCPU']]
                    $mem  = $fields[$colIndex['Memory']]
                    $fam  = $fields[$colIndex['Instance Family']]
 
                    $ramGiB = 0
                    if ($mem -match '([\d\.]+)\s*GiB') { $ramGiB = [double]$Matches[1] }
 
                    $familyNorm = switch -Regex ($fam) {
                        'General purpose'   { 'GeneralPurpose'; break }
                        'Compute optimized' { 'ComputeOptimized'; break }
                        'Memory optimized'  { 'MemoryOptimized'; break }
                        'Storage optimized' { 'StorageOptimized'; break }
                        default              { 'Other' }
                    }
                    if ($instType -match '^t[2-4][a-z]?\.') { $familyNorm = 'Burstable' }
 
                    if ([double]$vcpu -gt 0 -and $ramGiB -gt 0) {
                        $instanceSpecMap[$instType] = [PSCustomObject]@{
                            Type   = $instType
                            vCPU   = [double]$vcpu
                            RAM    = $ramGiB
                            Family = $familyNorm
                        }
                    }
                }
            }
            elseif ($termType -eq 'Reserved') {
    $leaseLength    = $fields[$colIndex['LeaseContractLength']]
    $purchaseOption = $fields[$colIndex['PurchaseOption']]
    $offeringClass  = $fields[$colIndex['OfferingClass']]
    $unit           = $fields[$colIndex['Unit']]
    $riPriceDesc    = $fields[$colIndex['PriceDescription']]
 
    $riIsExcludedVariant = $riPriceDesc -match 'BYOL' -or $riPriceDesc -match 'without licenses' -or $riPriceDesc -match 'No License required' -or $riPriceDesc -match 'SQL Server'
 
    if ($leaseLength -eq '3yr' -and $purchaseOption -eq 'No Upfront' -and $offeringClass -eq 'standard' -and $unit -eq 'Hrs' -and -not $riIsExcludedVariant) {
        $riKey = "$instType|$os"
        $riPriceVal = $fields[$colIndex['PricePerUnit']]
        if ($riPriceVal) {
            $riPriceValDouble = [double]$riPriceVal
            if (-not $ec2RIPriceMap.ContainsKey($riKey) -or $riPriceValDouble -lt $ec2RIPriceMap[$riKey]) {
                $ec2RIPriceMap[$riKey] = $riPriceValDouble
            }
        }
    }
}
        }
        elseif ($productFamily -eq 'Storage') {
            $volApiName = $fields[$colIndex['Volume API Name']]
            $termType   = $fields[$colIndex['TermType']]
            $unit       = $fields[$colIndex['Unit']]

            if ($termType -eq 'OnDemand' -and $unit -eq 'GB-Mo') {
                if ($volApiName -eq 'gp3' -and $null -eq $ebsGp3Price) {
                    $ebsGp3Price = [double]$fields[$colIndex['PricePerUnit']]
                }
                elseif ($volApiName -eq 'io2' -and $null -eq $ebsIo2Price) {
                    $ebsIo2Price = [double]$fields[$colIndex['PricePerUnit']]
                }
            }
        }
        elseif ($productFamily -eq 'System Operation' -and $colIndex.ContainsKey('Group')) {
            $group      = $fields[$colIndex['Group']]
            $volApiName = $fields[$colIndex['Volume API Name']]
            $termType   = $fields[$colIndex['TermType']]

            if ($termType -eq 'OnDemand') {
                if ($group -like '*IOPS*' -and $volApiName -eq 'gp3' -and $null -eq $ebsGp3IOPSPrice) {
                    $ebsGp3IOPSPrice = [double]$fields[$colIndex['PricePerUnit']]
                }
                elseif ($group -like '*Throughput*' -and $volApiName -eq 'gp3' -and $null -eq $ebsGp3ThroughputPrice) {
                    $ebsGp3ThroughputPrice = [double]$fields[$colIndex['PricePerUnit']]
                }
                elseif ($group -like '*IOPS*' -and $volApiName -eq 'io2' -and $null -eq $ebsIo2IOPSPrice) {
                    $ebsIo2IOPSPrice = [double]$fields[$colIndex['PricePerUnit']]
                }
            }
        }
    }
    $parser.Close()
 
    if (-not $ebsGp3Price) { $ebsGp3Price = 0.08 }
    # --- No hardcoded fallback for new SKUs - if not found in CSV, stays 0 so DiskProcessor.ps1 can flag "verify manually" instead of silently using a guessed rate ---
    if (-not $ebsGp3IOPSPrice)       { $ebsGp3IOPSPrice = 0 }
    if (-not $ebsGp3ThroughputPrice) { $ebsGp3ThroughputPrice = 0 }
    if (-not $ebsIo2Price)           { $ebsIo2Price = 0 }
    if (-not $ebsIo2IOPSPrice)       { $ebsIo2IOPSPrice = 0 }
 
    Write-Host "Index built. Rows scanned: $rowCount | Unique instance specs: $($instanceSpecMap.Count) | OnDemand entries: $($ec2PriceMap.Count) | 3yr-RI entries: $($ec2RIPriceMap.Count) | gp3 price: `$$ebsGp3Price/GB-Mo | io2 price: `$$ebsIo2Price/GB-Mo" -ForegroundColor Green
 
    return [PSCustomObject]@{
        InstanceSpecMap        = $instanceSpecMap
        EC2PriceMap            = $ec2PriceMap
        EC2RIPriceMap          = $ec2RIPriceMap
        EBSGp3Price            = $ebsGp3Price
        EBSGp3IOPSPrice        = $ebsGp3IOPSPrice
        EBSGp3ThroughputPrice  = $ebsGp3ThroughputPrice
        EBSIo2Price            = $ebsIo2Price
        EBSIo2IOPSPrice        = $ebsIo2IOPSPrice
    }
}
 
function Get-EC2Price {
    param(
        $Index,
        [string]$InstanceType,
        [string]$OperatingSystem,
        [string]$PriceType = "OnDemand"
    )
    $key = "$InstanceType|$OperatingSystem"
    $map = if ($PriceType -eq "Reserved") { $Index.EC2RIPriceMap } else { $Index.EC2PriceMap }
    if ($map.ContainsKey($key)) { return $map[$key] }
    return 0
}
