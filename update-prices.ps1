$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$workbookPath = Join-Path $root 'package-prices.xlsx'
$htmlPath = Join-Path $root 'index.html'
$expectedPackages = @('home-4', 'office-8', 'enterprise-16')
$rupee = [string][char]0x20B9
$archive = $null
$temporaryHtml = $null

function Get-CellText {
    param(
        [System.Xml.XmlNode]$Cell,
        [object[]]$SharedStrings
    )

    if ($null -eq $Cell) {
        return $null
    }

    $valueNode = $Cell.SelectSingleNode("*[local-name()='is']//*[local-name()='t']")
    if ($null -eq $valueNode) {
        $valueNode = $Cell.SelectSingleNode("*[local-name()='v']")
    }
    if ($null -eq $valueNode) {
        return $null
    }

    if ($Cell.GetAttribute('t') -eq 's') {
        $sharedIndex = [int]$valueNode.InnerText
        if ($sharedIndex -lt 0 -or $sharedIndex -ge $SharedStrings.Count) {
            throw "Invalid shared string index in cell $($Cell.GetAttribute('r'))."
        }
        return $SharedStrings[$sharedIndex]
    }

    return $valueNode.InnerText
}

try {
    if (-not (Test-Path -LiteralPath $workbookPath -PathType Leaf)) {
        throw "Workbook not found: $workbookPath"
    }
    if (-not (Test-Path -LiteralPath $htmlPath -PathType Leaf)) {
        throw "Website file not found: $htmlPath"
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($workbookPath)
    $worksheetEntry = $archive.GetEntry('xl/worksheets/sheet1.xml')
    if ($null -eq $worksheetEntry) {
        throw 'The workbook must contain a first worksheet named Packages.'
    }

    $sharedStrings = @()
    $sharedStringsEntry = $archive.GetEntry('xl/sharedStrings.xml')
    if ($null -ne $sharedStringsEntry) {
        $sharedStream = $sharedStringsEntry.Open()
        try {
            $settings = New-Object System.Xml.XmlReaderSettings
            $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
            $settings.XmlResolver = $null
            $xmlReader = [System.Xml.XmlReader]::Create($sharedStream, $settings)
            try {
                $sharedXml = New-Object System.Xml.XmlDocument
                $sharedXml.XmlResolver = $null
                $sharedXml.Load($xmlReader)
                foreach ($item in $sharedXml.SelectNodes("//*[local-name()='si']")) {
                    $sharedStrings += (($item.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText }) -join '')
                }
            }
            finally {
                $xmlReader.Dispose()
            }
        }
        finally {
            $sharedStream.Dispose()
        }
    }

    $worksheetStream = $worksheetEntry.Open()
    try {
        $settings = New-Object System.Xml.XmlReaderSettings
        $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
        $settings.XmlResolver = $null
        $xmlReader = [System.Xml.XmlReader]::Create($worksheetStream, $settings)
        try {
            $worksheetXml = New-Object System.Xml.XmlDocument
            $worksheetXml.XmlResolver = $null
            $worksheetXml.Load($xmlReader)
        }
        finally {
            $xmlReader.Dispose()
        }
    }
    finally {
        $worksheetStream.Dispose()
    }

    $rows = $worksheetXml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")
    $headerRow = $null
    foreach ($row in $rows) {
        $idCell = $row.SelectSingleNode("*[local-name()='c' and starts-with(@r, 'A')]")
        if ($null -ne $idCell) {
            if ((Get-CellText -Cell $idCell -SharedStrings $sharedStrings) -eq 'Package ID') {
                $headerRow = $row
                break
            }
        }
    }
    if ($null -eq $headerRow) {
        throw "Could not find the 'Package ID' header in the Packages worksheet."
    }

    $columns = @{}
    foreach ($cell in $headerRow.SelectNodes("*[local-name()='c']")) {
        $column = [regex]::Match($cell.GetAttribute('r'), '^[A-Z]+').Value
        $headerText = Get-CellText -Cell $cell -SharedStrings $sharedStrings
        if (-not [string]::IsNullOrWhiteSpace($headerText)) {
            $columns[$headerText] = $column
        }
    }
    foreach ($requiredColumn in @('Package ID', 'Price INR')) {
        if (-not $columns.ContainsKey($requiredColumn)) {
            throw "Could not find the '$requiredColumn' column in the Packages worksheet."
        }
    }

    $priceColumn = $columns['Price INR']
    $idColumn = $columns['Package ID']
    $prices = @{}
    foreach ($row in $rows) {
        if ($row -eq $headerRow) {
            continue
        }
        $idCell = $row.SelectSingleNode("*[local-name()='c' and starts-with(@r, '$idColumn')]")
        if ($null -eq $idCell) {
            continue
        }

        $idText = Get-CellText -Cell $idCell -SharedStrings $sharedStrings

        if ([string]::IsNullOrWhiteSpace($idText)) {
            continue
        }
        $priceCell = $row.SelectSingleNode("*[local-name()='c' and starts-with(@r, '$priceColumn')]")
        if ($idText -notin $expectedPackages) {
            if ($null -eq $priceCell -or $null -eq $priceCell.SelectSingleNode("*[local-name()='v']")) {
                continue
            }
            throw "Unexpected package ID '$idText'. Do not change the Package ID values."
        }
        if ($prices.ContainsKey($idText)) {
            throw "Duplicate package ID '$idText' in the workbook."
        }

        if ($null -eq $priceCell -or $priceCell.GetAttribute('t') -eq 's') {
            throw "Enter a numeric price in the Price INR column for package '$idText'."
        }
        $priceValue = $priceCell.SelectSingleNode("*[local-name()='v']")
        $price = [decimal]0
        if ($null -eq $priceValue -or -not [decimal]::TryParse(
            $priceValue.InnerText,
            [System.Globalization.NumberStyles]::Number,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [ref]$price
        )) {
            throw "Invalid numeric price for package '$idText'."
        }
        if ($price -le 0 -or $price -ne [decimal]::Truncate($price)) {
            throw "Price for package '$idText' must be a positive whole number of rupees."
        }

        $prices[$idText] = [long]$price
    }

    foreach ($packageId in $expectedPackages) {
        if (-not $prices.ContainsKey($packageId)) {
            throw "Package '$packageId' is missing a price in the workbook."
        }
    }

    $html = [System.IO.File]::ReadAllText($htmlPath)
    foreach ($packageId in $expectedPackages) {
        $number = $prices[$packageId].ToString([System.Globalization.CultureInfo]::InvariantCulture)
        $lastGroup = $number.Substring([Math]::Max(0, $number.Length - 3))
        $leadingDigits = $number.Substring(0, $number.Length - $lastGroup.Length)
        $groups = @()
        while ($leadingDigits.Length -gt 2) {
            $groups = @($leadingDigits.Substring($leadingDigits.Length - 2)) + $groups
            $leadingDigits = $leadingDigits.Substring(0, $leadingDigits.Length - 2)
        }
        if ($leadingDigits.Length -gt 0) {
            $groups = @($leadingDigits) + $groups
        }
        $formattedPrice = if ($groups.Count -gt 0) { ($groups -join ',') + ',' + $lastGroup } else { $lastGroup }
        $pattern = '(<div class="price" data-price-id="' + [regex]::Escape($packageId) + '">\s*)' + [regex]::Escape($rupee) + '[\d,]+(\s*<small>[^<]*</small>\s*</div>)'
        $priceRegex = [regex]::new($pattern)
        if ($priceRegex.Matches($html).Count -ne 1) {
            throw "Expected exactly one website price element for '$packageId'."
        }
        $html = $priceRegex.Replace($html, ('$1' + $rupee + $formattedPrice + '$2'), 1)
    }

    $temporaryHtml = Join-Path $root ('index.html.' + [guid]::NewGuid().ToString('N') + '.tmp')
    [System.IO.File]::WriteAllText($temporaryHtml, $html, [System.Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $temporaryHtml -Destination $htmlPath -Force
    $temporaryHtml = $null

    Write-Output 'Website package prices updated successfully:'
    foreach ($packageId in $expectedPackages) {
        Write-Output ('  {0}: {1}{2}' -f $packageId, $rupee, $prices[$packageId])
    }
}
catch {
    Write-Error ("Price update failed: " + $_.Exception.Message)
    exit 1
}
finally {
    if ($null -ne $archive) {
        $archive.Dispose()
    }
    if ($null -ne $temporaryHtml -and (Test-Path -LiteralPath $temporaryHtml)) {
        Remove-Item -LiteralPath $temporaryHtml -Force
    }
}
