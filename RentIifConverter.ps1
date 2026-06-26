#Requires -Version 7.0
[CmdletBinding()]
param(
    [switch]$NoGui,

    [string]$InputPath,

    [string]$OutputDirectory,

    [ValidatePattern('^(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\d{4}$')]
    [string]$ProcessingDate,

    [ValidateSet('Payment', 'Invoice')]
    [string]$ProcessType = 'Payment',

    [string]$ReceivableAccount = 'A11000 - Accounts Receivable',

    [string]$DepositAccount = 'A12000 - Undeposited Funds',

    [string]$IncomeAccount = 'A47600 - ARB Rental Income'
)

$ErrorActionPreference = 'Stop'

if (-not $IsWindows) {
    throw 'Rent IIF Converter requires Windows.'
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

[System.Windows.Forms.Application]::EnableVisualStyles()

function Convert-ToIifText {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [double] -or $Value -is [decimal] -or $Value -is [single] -or $Value -is [int]) {
        return [string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0}', $Value)
    }

    return ([string]$Value) -replace "`t|`r|`n", ' '
}

function Convert-ToIifAmount {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }

    if ($Value -is [double] -or $Value -is [decimal] -or $Value -is [single] -or $Value -is [int]) {
        return [decimal]$Value
    }

    $text = ([string]$Value).Trim()
    $isParenthesizedNegative = $text.StartsWith('(') -and $text.EndsWith(')')
    $text = $text.Trim('(', ')').Replace('$', '').Replace(',', '')

    $amount = [decimal]::Zero
    if (-not [decimal]::TryParse(
            $text,
            [Globalization.NumberStyles]::Number -bor [Globalization.NumberStyles]::AllowLeadingSign,
            [Globalization.CultureInfo]::InvariantCulture,
            [ref]$amount
        )) {
        return $null
    }

    if ($isParenthesizedNegative) {
        $amount = -1 * $amount
    }

    return $amount
}

function New-IifLine {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Fields
    )

    return (($Fields | ForEach-Object { Convert-ToIifText $_ }) -join "`t")
}

function Format-IifAmount {
    param(
        [Parameter(Mandatory)]
        [decimal]$Amount
    )

    return $Amount.ToString('0.00', [Globalization.CultureInfo]::InvariantCulture)
}

function Normalize-IifName {
    param(
        [AllowNull()]
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ''
    }

    return ($Name -replace '\s+', ' ').Trim()
}

function Convert-XlsxColumnNameToNumber {
    param(
        [Parameter(Mandatory)]
        [string]$ColumnName
    )

    $number = 0
    foreach ($character in $ColumnName.ToUpperInvariant().ToCharArray()) {
        $number = ($number * 26) + ([int][char]$character - [int][char]'A' + 1)
    }

    return $number
}

function Get-XlsxFirstSheetRows {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $stream = [IO.FileStream]::new($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $zip = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Read)
    try {
        $sharedStrings = @()
        $sharedEntry = $zip.GetEntry('xl/sharedStrings.xml')
        if ($sharedEntry) {
            $reader = [IO.StreamReader]::new($sharedEntry.Open())
            try {
                [xml]$sharedXml = $reader.ReadToEnd()
            }
            finally {
                $reader.Close()
            }

            $sharedNamespace = [Xml.XmlNamespaceManager]::new($sharedXml.NameTable)
            $sharedNamespace.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
            foreach ($item in $sharedXml.SelectNodes('//x:si', $sharedNamespace)) {
                $sharedStrings += (($item.SelectNodes('.//x:t', $sharedNamespace) | ForEach-Object { $_.InnerText }) -join '')
            }
        }

        $workbookRelsEntry = $zip.GetEntry('xl/_rels/workbook.xml.rels')
        $workbookEntry = $zip.GetEntry('xl/workbook.xml')
        if (-not $workbookRelsEntry -or -not $workbookEntry) {
            throw 'The selected workbook is missing required workbook metadata.'
        }

        $reader = [IO.StreamReader]::new($workbookEntry.Open())
        try {
            [xml]$workbookXml = $reader.ReadToEnd()
        }
        finally {
            $reader.Close()
        }

        $reader = [IO.StreamReader]::new($workbookRelsEntry.Open())
        try {
            [xml]$relsXml = $reader.ReadToEnd()
        }
        finally {
            $reader.Close()
        }

        $workbookNamespace = [Xml.XmlNamespaceManager]::new($workbookXml.NameTable)
        $workbookNamespace.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        $workbookNamespace.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $firstSheet = $workbookXml.SelectSingleNode('//x:sheets/x:sheet[1]', $workbookNamespace)
        if (-not $firstSheet) {
            throw 'The selected workbook does not contain a worksheet.'
        }

        $relationshipId = $firstSheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $relsNamespace = [Xml.XmlNamespaceManager]::new($relsXml.NameTable)
        $relsNamespace.AddNamespace('r', 'http://schemas.openxmlformats.org/package/2006/relationships')
        $relationship = $relsXml.SelectSingleNode("//r:Relationship[@Id='$relationshipId']", $relsNamespace)
        if (-not $relationship) {
            throw 'The selected workbook is missing first-sheet relationship metadata.'
        }

        $target = ([string]$relationship.Target).TrimStart('/')
        if (-not $target.StartsWith('xl/')) {
            $target = "xl/$target"
        }

        $sheetEntry = $zip.GetEntry($target)
        if (-not $sheetEntry) {
            throw "The selected workbook is missing worksheet data: $target"
        }

        $reader = [IO.StreamReader]::new($sheetEntry.Open())
        try {
            [xml]$sheetXml = $reader.ReadToEnd()
        }
        finally {
            $reader.Close()
        }

        $sheetNamespace = [Xml.XmlNamespaceManager]::new($sheetXml.NameTable)
        $sheetNamespace.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

        foreach ($row in $sheetXml.SelectNodes('//x:sheetData/x:row', $sheetNamespace)) {
            $values = @{}
            foreach ($cell in $row.SelectNodes('x:c', $sheetNamespace)) {
                $cellReference = [string]$cell.r
                $columnName = $cellReference -replace '\d', ''
                $columnIndex = Convert-XlsxColumnNameToNumber -ColumnName $columnName
                $valueNode = $cell.SelectSingleNode('x:v', $sheetNamespace)
                $value = if ($valueNode) { [string]$valueNode.InnerText } else { '' }

                if ($cell.t -eq 's' -and $value -ne '') {
                    $value = $sharedStrings[[int]$value]
                }

                $values[$columnIndex] = $value
            }

            [pscustomobject]@{
                RowNumber = [int]$row.r
                Values = $values
            }
        }
    }
    finally {
        $zip.Dispose()
        $stream.Dispose()
    }
}

function Get-XlsxCellText {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Values,

        [Parameter(Mandatory)]
        [int]$Column
    )

    if (-not $Values.ContainsKey($Column)) {
        return ''
    }

    return Convert-ToIifText $Values[$Column]
}

function Assert-RentReportHeaders {
    param(
        [Parameter(Mandatory)]
        [object[]]$Rows
    )

    if ($Rows.Count -lt 2) {
        throw 'The selected workbook does not contain rent transaction rows.'
    }

    $header = $Rows[0].Values
    $expected = @{
        3 = 'Tenant'
        7 = 'Invoiced'
        6 = 'Datetime'
        9 = 'Payment'
    }

    $missing = foreach ($column in $expected.Keys) {
        $actual = Get-XlsxCellText -Values $header -Column $column
        if ($actual -ne $expected[$column]) {
            "Column $column expected '$($expected[$column])' but found '$actual'"
        }
    }

    if ($missing) {
        throw "The selected workbook does not look like the expected rent transaction report.`r`n$($missing -join "`r`n")"
    }
}

function New-RentIifData {
    param(
        [Parameter(Mandatory)]
        [string]$InputPath,

        [Parameter(Mandatory)]
        [datetime]$ProcessingDate,

        [Parameter(Mandatory)]
        [ValidateSet('Payment', 'Invoice')]
        [string]$ProcessType,

        [Parameter(Mandatory)]
        [string]$ReceivableAccount,

        [Parameter(Mandatory)]
        [string]$DepositAccount,

        [Parameter(Mandatory)]
        [string]$IncomeAccount
    )

    if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) {
        throw "Input file was not found: $InputPath"
    }

    if ([string]::IsNullOrWhiteSpace($ReceivableAccount)) {
        throw 'Enter the QuickBooks accounts receivable account name.'
    }

    if ([string]::IsNullOrWhiteSpace($DepositAccount)) {
        throw 'Enter the QuickBooks deposit account name.'
    }

    if ([string]::IsNullOrWhiteSpace($IncomeAccount)) {
        throw 'Enter the QuickBooks income account name.'
    }

    $rows = @(Get-XlsxFirstSheetRows -Path $InputPath)
    Assert-RentReportHeaders -Rows $rows

    $iifDate = $ProcessingDate.ToString('M/d/yyyy', [Globalization.CultureInfo]::InvariantCulture)
    $transactionType = if ($ProcessType -eq 'Invoice') { 'INVOICE' } else { 'PAYMENT' }
    $amountColumn = if ($ProcessType -eq 'Invoice') { 7 } else { 9 }
    $trnsAccount = if ($ProcessType -eq 'Invoice') { $ReceivableAccount } else { $DepositAccount }
    $splAccount = if ($ProcessType -eq 'Invoice') { $IncomeAccount } else { $ReceivableAccount }
    $modeLabel = if ($ProcessType -eq 'Invoice') { 'invoiced amounts in column G' } else { 'payment amounts in column I' }
    $iifLines = [System.Collections.Generic.List[string]]::new()
    $previewRows = [System.Collections.Generic.List[object]]::new()
    $iifLines.Add((New-IifLine @('!TRNS', 'TRNSID', 'TRNSTYPE', 'DATE', 'ACCNT', 'NAME', 'AMOUNT', 'DOCNUM')))
    $iifLines.Add((New-IifLine @('!SPL', 'SPLID', 'TRNSTYPE', 'DATE', 'ACCNT', 'NAME', 'AMOUNT', 'DOCNUM')))
    $iifLines.Add((New-IifLine @('!ENDTRNS', '', '', '', '', '', '', '')))

    $processedRows = 0
    $skippedRows = 0

    foreach ($row in $rows | Where-Object { $_.RowNumber -gt 1 }) {
        $tenantValue = Normalize-IifName (Get-XlsxCellText -Values $row.Values -Column 3)
        if ([string]::IsNullOrWhiteSpace($tenantValue)) {
            $skippedRows++
            continue
        }

        $currentAmount = Convert-ToIifAmount (Get-XlsxCellText -Values $row.Values -Column $amountColumn)
        if ($null -eq $currentAmount -or $currentAmount -eq 0) {
            $skippedRows++
            continue
        }

        $dateTimeValue = Get-XlsxCellText -Values $row.Values -Column 6
        $docNum = Get-XlsxCellText -Values $row.Values -Column 14
        $amount = [Math]::Abs($currentAmount)
        $splitAmount = -1 * $amount
        $amountText = Format-IifAmount -Amount $amount
        $splitAmountText = Format-IifAmount -Amount $splitAmount

        $previewRows.Add([pscustomobject]@{
            Type = $ProcessType
            SourceRow = $row.RowNumber
            Tenant = $tenantValue
            Date = $iifDate
            Amount = $amountText
            TrnsAccount = $trnsAccount
            SplAccount = $splAccount
            DocNum = $docNum
            SourceDateTime = $dateTimeValue
        })

        $iifLines.Add((New-IifLine @('TRNS', ' ', $transactionType, $iifDate, $trnsAccount, $tenantValue, $amountText, $docNum)))
        $iifLines.Add((New-IifLine @('SPL', ' ', $transactionType, $iifDate, $splAccount, $tenantValue, $splitAmountText, $docNum)))
        $processedRows++
    }

    if ($processedRows -eq 0) {
        throw "No $($ProcessType.ToLowerInvariant()) rows were found. Check that the selected workbook has $modeLabel."
    }

    $iifLines.Add((New-IifLine @('ENDTRNS', '', '', '', '', '', '', '')))

    [pscustomobject]@{
        IifLines = $iifLines
        PreviewRows = $previewRows
        ProcessedRows = $processedRows
        SkippedRows = $skippedRows
    }
}

function Convert-RentReportToIif {
    param(
        [Parameter(Mandatory)]
        [string]$InputPath,

        [Parameter(Mandatory)]
        [string]$OutputDirectory,

        [Parameter(Mandatory)]
        [datetime]$ProcessingDate,

        [Parameter(Mandatory)]
        [ValidateSet('Payment', 'Invoice')]
        [string]$ProcessType,

        [Parameter(Mandatory)]
        [string]$ReceivableAccount,

        [Parameter(Mandatory)]
        [string]$DepositAccount,

        [Parameter(Mandatory)]
        [string]$IncomeAccount
    )

    if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
        New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
    }

    $processingDateText = $ProcessingDate.ToString('MMddyyyy', [Globalization.CultureInfo]::InvariantCulture)
    $outputPrefix = if ($ProcessType -eq 'Invoice') { 'RentInvoice' } else { 'RentPayment' }
    $outputPath = Join-Path -Path $OutputDirectory -ChildPath "$outputPrefix$processingDateText.iif"
    $data = New-RentIifData `
        -InputPath $InputPath `
        -ProcessingDate $ProcessingDate `
        -ProcessType $ProcessType `
        -ReceivableAccount $ReceivableAccount `
        -DepositAccount $DepositAccount `
        -IncomeAccount $IncomeAccount
    [System.IO.File]::WriteAllLines($outputPath, $data.IifLines, [System.Text.UTF8Encoding]::new($false))

    [pscustomobject]@{
        OutputPath = $outputPath
        ProcessedRows = $data.ProcessedRows
        SkippedRows = $data.SkippedRows
        PreviewRows = $data.PreviewRows
    }
}

function Get-DateFromFileName {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $name = [IO.Path]::GetFileNameWithoutExtension($Path)
    if ($name -match '(\d{8})') {
        $date = [datetime]::MinValue
        if ([datetime]::TryParseExact(
                $Matches[1],
                'MMddyyyy',
                [Globalization.CultureInfo]::InvariantCulture,
                [Globalization.DateTimeStyles]::None,
                [ref]$date
            )) {
            return $date.ToString('MMddyyyy', [Globalization.CultureInfo]::InvariantCulture)
        }
    }

    return (Get-Date).ToString('MMddyyyy', [Globalization.CultureInfo]::InvariantCulture)
}

if ($NoGui) {
    if ([string]::IsNullOrWhiteSpace($InputPath)) {
        throw 'NoGui mode requires -InputPath.'
    }

    if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
        throw 'NoGui mode requires -OutputDirectory.'
    }

    if ([string]::IsNullOrWhiteSpace($ProcessingDate)) {
        $ProcessingDate = Get-DateFromFileName -Path $InputPath
    }

    $date = [datetime]::MinValue
    if (-not [datetime]::TryParseExact(
            $ProcessingDate,
            'MMddyyyy',
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::None,
            [ref]$date
        )) {
        throw 'Enter the date in mmddyyyy format, for example 06122026.'
    }

    $result = Convert-RentReportToIif `
        -InputPath $InputPath `
        -OutputDirectory $OutputDirectory `
        -ProcessingDate $date `
        -ProcessType $ProcessType `
        -ReceivableAccount $ReceivableAccount `
        -DepositAccount $DepositAccount `
        -IncomeAccount $IncomeAccount
    Write-Host "Output: $($result.OutputPath)"
    Write-Host "Processed $($ProcessType.ToLowerInvariant()) rows: $($result.ProcessedRows)"
    Write-Host "Skipped rows: $($result.SkippedRows)"
    return
}

$settingsPath = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'RentIifConverter\settings.json'

function Get-AppSettings {
    $defaults = [pscustomobject]@{
        LastInputDirectory = [Environment]::GetFolderPath('MyDocuments')
        LastOutputDirectory = [Environment]::GetFolderPath('MyDocuments')
        ProcessType = 'Payment'
        ReceivableAccount = 'A11000 - Accounts Receivable'
        DepositAccount = 'A12000 - Undeposited Funds'
        IncomeAccount = 'A47600 - ARB Rental Income'
    }

    if (-not (Test-Path -LiteralPath $settingsPath -PathType Leaf)) {
        return $defaults
    }

    try {
        $saved = Get-Content -Raw -LiteralPath $settingsPath | ConvertFrom-Json
        if ([string]::IsNullOrWhiteSpace([string]$saved.DepositAccount) -and -not [string]::IsNullOrWhiteSpace([string]$saved.IncomeAccount)) {
            $saved | Add-Member -NotePropertyName DepositAccount -NotePropertyValue 'A12000 - Undeposited Funds' -Force
        }

        if ($saved.ReceivableAccount -eq 'Accounts Receivable') {
            $saved.ReceivableAccount = 'A11000 - Accounts Receivable'
        }

        if ($saved.DepositAccount -eq 'Undeposited Funds') {
            $saved.DepositAccount = 'A12000 - Undeposited Funds'
        }

        foreach ($name in $defaults.PSObject.Properties.Name) {
            if ([string]::IsNullOrWhiteSpace([string]$saved.$name)) {
                $saved | Add-Member -NotePropertyName $name -NotePropertyValue $defaults.$name -Force
            }
        }

        return $saved
    }
    catch {
        return $defaults
    }
}

function Save-AppSettings {
    param(
        [string]$InputFile,
        [string]$OutputDirectory,
        [string]$ProcessType,
        [string]$ReceivableAccount,
        [string]$DepositAccount,
        [string]$IncomeAccount
    )

    $settingsDirectory = Split-Path -Parent $settingsPath
    if (-not (Test-Path -LiteralPath $settingsDirectory -PathType Container)) {
        New-Item -Path $settingsDirectory -ItemType Directory -Force | Out-Null
    }

    $inputDirectory = if (-not [string]::IsNullOrWhiteSpace($InputFile) -and (Test-Path -LiteralPath $InputFile -PathType Leaf)) {
        Split-Path -Parent $InputFile
    }
    else {
        $appSettings.LastInputDirectory
    }

    [pscustomobject]@{
        LastInputDirectory = $inputDirectory
        LastOutputDirectory = $OutputDirectory
        ProcessType = $ProcessType
        ReceivableAccount = $ReceivableAccount
        DepositAccount = $DepositAccount
        IncomeAccount = $IncomeAccount
    } | ConvertTo-Json | Set-Content -LiteralPath $settingsPath -Encoding UTF8
}

function Get-ValidatedDate {
    param(
        [Parameter(Mandatory)]
        [string]$Text
    )

    $date = [datetime]::MinValue
    if (-not [datetime]::TryParseExact(
            $Text,
            'MMddyyyy',
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::None,
            [ref]$date
        )) {
        throw 'Enter the date in mmddyyyy format, for example 06122026.'
    }

    return $date
}

$appSettings = Get-AppSettings

$form = [System.Windows.Forms.Form]::new()
$form.Text = 'Rent IIF Converter'
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = [System.Drawing.Size]::new(860, 680)
$form.Size = [System.Drawing.Size]::new(1020, 760)

$margin = 18
$labelWidth = 110
$fieldLeft = 136
$buttonWidth = 96
$gap = 10

$title = [System.Windows.Forms.Label]::new()
$title.Text = 'Rent IIF Converter'
$title.Font = [System.Drawing.Font]::new('Segoe UI', 16, [System.Drawing.FontStyle]::Bold)
$title.Location = [System.Drawing.Point]::new($margin, 16)
$title.Size = [System.Drawing.Size]::new(760, 32)

$statusLabel = [System.Windows.Forms.Label]::new()
$statusLabel.Text = 'Choose a rent transaction report and create a QuickBooks IIF file.'
$statusLabel.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$statusLabel.Location = [System.Drawing.Point]::new($margin, 56)
$statusLabel.Size = [System.Drawing.Size]::new(940, 24)

$inputLabel = [System.Windows.Forms.Label]::new()
$inputLabel.Text = 'Rent report'
$inputLabel.Location = [System.Drawing.Point]::new($margin, 104)
$inputLabel.Size = [System.Drawing.Size]::new($labelWidth, 24)

$inputPathBox = [System.Windows.Forms.TextBox]::new()
$inputPathBox.Location = [System.Drawing.Point]::new($fieldLeft, 101)
$inputPathBox.Size = [System.Drawing.Size]::new(540, 26)

$browseInput = [System.Windows.Forms.Button]::new()
$browseInput.Text = 'Browse...'
$browseInput.Location = [System.Drawing.Point]::new(700, 99)
$browseInput.Size = [System.Drawing.Size]::new($buttonWidth, 30)

$dateLabel = [System.Windows.Forms.Label]::new()
$dateLabel.Text = 'Date'
$dateLabel.Location = [System.Drawing.Point]::new($margin, 146)
$dateLabel.Size = [System.Drawing.Size]::new($labelWidth, 24)

$processingDateBox = [System.Windows.Forms.TextBox]::new()
$processingDateBox.Location = [System.Drawing.Point]::new($fieldLeft, 143)
$processingDateBox.Size = [System.Drawing.Size]::new(120, 26)
$processingDateBox.Text = (Get-Date).ToString('MMddyyyy', [Globalization.CultureInfo]::InvariantCulture)

$dateHint = [System.Windows.Forms.Label]::new()
$dateHint.Text = 'mmddyyyy'
$dateHint.Location = [System.Drawing.Point]::new(266, 146)
$dateHint.Size = [System.Drawing.Size]::new(120, 24)

$processLabel = [System.Windows.Forms.Label]::new()
$processLabel.Text = 'Process'
$processLabel.Location = [System.Drawing.Point]::new(380, 146)
$processLabel.Size = [System.Drawing.Size]::new(70, 24)

$processTypeBox = [System.Windows.Forms.ComboBox]::new()
$processTypeBox.DropDownStyle = 'DropDownList'
[void]$processTypeBox.Items.Add('Payment')
[void]$processTypeBox.Items.Add('Invoice')
$processTypeBox.Location = [System.Drawing.Point]::new(456, 143)
$processTypeBox.Size = [System.Drawing.Size]::new(140, 26)
$processTypeBox.SelectedItem = if ($appSettings.ProcessType -eq 'Invoice') { 'Invoice' } else { 'Payment' }

$outputLabel = [System.Windows.Forms.Label]::new()
$outputLabel.Text = 'Output folder'
$outputLabel.Location = [System.Drawing.Point]::new($margin, 188)
$outputLabel.Size = [System.Drawing.Size]::new($labelWidth, 24)

$outputPathBox = [System.Windows.Forms.TextBox]::new()
$outputPathBox.Location = [System.Drawing.Point]::new($fieldLeft, 185)
$outputPathBox.Size = [System.Drawing.Size]::new(540, 26)
$outputPathBox.Text = $appSettings.LastOutputDirectory

$browseOutput = [System.Windows.Forms.Button]::new()
$browseOutput.Text = 'Browse...'
$browseOutput.Location = [System.Drawing.Point]::new(700, 183)
$browseOutput.Size = [System.Drawing.Size]::new($buttonWidth, 30)

$receivableLabel = [System.Windows.Forms.Label]::new()
$receivableLabel.Text = 'A/R account'
$receivableLabel.Location = [System.Drawing.Point]::new($margin, 230)
$receivableLabel.Size = [System.Drawing.Size]::new($labelWidth, 24)

$receivableAccountBox = [System.Windows.Forms.TextBox]::new()
$receivableAccountBox.Location = [System.Drawing.Point]::new($fieldLeft, 227)
$receivableAccountBox.Size = [System.Drawing.Size]::new(320, 26)
$receivableAccountBox.Text = $appSettings.ReceivableAccount

$depositLabel = [System.Windows.Forms.Label]::new()
$depositLabel.Text = 'Deposit account'
$depositLabel.Location = [System.Drawing.Point]::new($margin, 272)
$depositLabel.Size = [System.Drawing.Size]::new($labelWidth, 24)

$depositAccountBox = [System.Windows.Forms.TextBox]::new()
$depositAccountBox.Location = [System.Drawing.Point]::new($fieldLeft, 269)
$depositAccountBox.Size = [System.Drawing.Size]::new(320, 26)
$depositAccountBox.Text = $appSettings.DepositAccount

$incomeLabel = [System.Windows.Forms.Label]::new()
$incomeLabel.Text = 'Income account'
$incomeLabel.Location = [System.Drawing.Point]::new($margin, 314)
$incomeLabel.Size = [System.Drawing.Size]::new($labelWidth, 24)

$incomeAccountBox = [System.Windows.Forms.TextBox]::new()
$incomeAccountBox.Location = [System.Drawing.Point]::new($fieldLeft, 311)
$incomeAccountBox.Size = [System.Drawing.Size]::new(320, 26)
$incomeAccountBox.Text = $appSettings.IncomeAccount

$previewButton = [System.Windows.Forms.Button]::new()
$previewButton.Text = 'Preview'
$previewButton.Location = [System.Drawing.Point]::new($fieldLeft, 356)
$previewButton.Size = [System.Drawing.Size]::new(100, 34)

$createButton = [System.Windows.Forms.Button]::new()
$createButton.Text = 'Create IIF'
$createButton.Location = [System.Drawing.Point]::new(246, 356)
$createButton.Size = [System.Drawing.Size]::new(110, 34)

$openFolderButton = [System.Windows.Forms.Button]::new()
$openFolderButton.Text = 'Open Folder'
$openFolderButton.Location = [System.Drawing.Point]::new(366, 356)
$openFolderButton.Size = [System.Drawing.Size]::new(110, 34)
$openFolderButton.Enabled = $false

$openIifButton = [System.Windows.Forms.Button]::new()
$openIifButton.Text = 'Open IIF'
$openIifButton.Location = [System.Drawing.Point]::new(486, 356)
$openIifButton.Size = [System.Drawing.Size]::new(100, 34)
$openIifButton.Enabled = $false

$copyPathButton = [System.Windows.Forms.Button]::new()
$copyPathButton.Text = 'Copy Path'
$copyPathButton.Location = [System.Drawing.Point]::new(596, 356)
$copyPathButton.Size = [System.Drawing.Size]::new(100, 34)
$copyPathButton.Enabled = $false

$previewGrid = [System.Windows.Forms.DataGridView]::new()
$previewGrid.Location = [System.Drawing.Point]::new($margin, 410)
$previewGrid.Size = [System.Drawing.Size]::new(960, 230)
$previewGrid.ReadOnly = $true
$previewGrid.AllowUserToAddRows = $false
$previewGrid.AllowUserToDeleteRows = $false
$previewGrid.AutoSizeColumnsMode = 'Fill'
$previewGrid.RowHeadersVisible = $false
$previewGrid.SelectionMode = 'FullRowSelect'

$log = [System.Windows.Forms.TextBox]::new()
$log.Location = [System.Drawing.Point]::new($margin, 658)
$log.Size = [System.Drawing.Size]::new(960, 90)
$log.Multiline = $true
$log.ScrollBars = 'Vertical'
$log.ReadOnly = $true
$log.Font = [System.Drawing.Font]::new('Consolas', 10)

$lastOutputDirectory = $null
$lastOutputPath = $null

function Update-Layout {
    $clientWidth = $form.ClientSize.Width
    $clientHeight = $form.ClientSize.Height
    $right = $clientWidth - $margin

    $browseInput.Location = [System.Drawing.Point]::new($right - $buttonWidth, 99)
    $browseOutput.Location = [System.Drawing.Point]::new($right - $buttonWidth, 183)

    $inputPathBox.Size = [System.Drawing.Size]::new([Math]::Max(260, $browseInput.Left - $gap - $fieldLeft), 26)
    $outputPathBox.Size = [System.Drawing.Size]::new([Math]::Max(260, $browseOutput.Left - $gap - $fieldLeft), 26)
    $receivableAccountBox.Size = [System.Drawing.Size]::new([Math]::Max(320, $right - $fieldLeft), 26)
    $depositAccountBox.Size = [System.Drawing.Size]::new([Math]::Max(320, $right - $fieldLeft), 26)
    $incomeAccountBox.Size = [System.Drawing.Size]::new([Math]::Max(320, $right - $fieldLeft), 26)

    $title.Size = [System.Drawing.Size]::new([Math]::Max(360, $clientWidth - (2 * $margin)), 32)
    $statusLabel.Size = [System.Drawing.Size]::new([Math]::Max(360, $clientWidth - (2 * $margin)), 24)
    $previewGrid.Size = [System.Drawing.Size]::new([Math]::Max(360, $clientWidth - (2 * $margin)), [Math]::Max(150, $clientHeight - 572))
    $log.Location = [System.Drawing.Point]::new($margin, $previewGrid.Bottom + 16)
    $log.Size = [System.Drawing.Size]::new([Math]::Max(360, $clientWidth - (2 * $margin)), [Math]::Max(70, $clientHeight - $log.Top - 18))
}

$form.Controls.AddRange(@(
    $title,
    $statusLabel,
    $inputLabel,
    $inputPathBox,
    $browseInput,
    $dateLabel,
    $processingDateBox,
    $dateHint,
    $processLabel,
    $processTypeBox,
    $outputLabel,
    $outputPathBox,
    $browseOutput,
    $receivableLabel,
    $receivableAccountBox,
    $depositLabel,
    $depositAccountBox,
    $incomeLabel,
    $incomeAccountBox,
    $previewButton,
    $createButton,
    $openFolderButton,
    $openIifButton,
    $copyPathButton,
    $previewGrid,
    $log
))

$form.Add_Shown({ Update-Layout })
$form.Add_Resize({ Update-Layout })

function Update-ProcessAccountFields {
    $isInvoice = $processTypeBox.SelectedItem -eq 'Invoice'

    $depositLabel.Enabled = -not $isInvoice
    $depositAccountBox.Enabled = -not $isInvoice
    $incomeLabel.Enabled = $isInvoice
    $incomeAccountBox.Enabled = $isInvoice

    if ($isInvoice) {
        if ([string]::IsNullOrWhiteSpace($receivableAccountBox.Text)) {
            $receivableAccountBox.Text = 'A11000 - Accounts Receivable'
        }
        if ([string]::IsNullOrWhiteSpace($incomeAccountBox.Text)) {
            $incomeAccountBox.Text = 'A47600 - ARB Rental Income'
        }
    }
    else {
        if ([string]::IsNullOrWhiteSpace($receivableAccountBox.Text)) {
            $receivableAccountBox.Text = 'A11000 - Accounts Receivable'
        }
        if ([string]::IsNullOrWhiteSpace($depositAccountBox.Text)) {
            $depositAccountBox.Text = 'A12000 - Undeposited Funds'
        }
    }

    $previewGrid.DataSource = $null
    $openIifButton.Enabled = $false
    $copyPathButton.Enabled = $false
}

$processTypeBox.Add_SelectedIndexChanged({
    Update-ProcessAccountFields
})

Update-ProcessAccountFields

$browseInput.Add_Click({
    $dialog = [System.Windows.Forms.OpenFileDialog]::new()
    $dialog.Title = 'Choose rent transaction report'
    $dialog.Filter = 'Excel workbooks (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    $dialog.CheckFileExists = $true
    if (-not [string]::IsNullOrWhiteSpace($appSettings.LastInputDirectory) -and (Test-Path -LiteralPath $appSettings.LastInputDirectory -PathType Container)) {
        $dialog.InitialDirectory = $appSettings.LastInputDirectory
    }

    if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $inputPathBox.Text = $dialog.FileName
        $processingDateBox.Text = Get-DateFromFileName -Path $dialog.FileName
        $outputPathBox.Text = Split-Path -Parent $dialog.FileName
        $previewGrid.DataSource = $null
        $openIifButton.Enabled = $false
        $copyPathButton.Enabled = $false
        Save-AppSettings -InputFile $inputPathBox.Text -OutputDirectory $outputPathBox.Text -ProcessType $processTypeBox.SelectedItem -ReceivableAccount $receivableAccountBox.Text -DepositAccount $depositAccountBox.Text -IncomeAccount $incomeAccountBox.Text
    }
})

$browseOutput.Add_Click({
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description = 'Choose where to save the QuickBooks IIF file'
    if (-not [string]::IsNullOrWhiteSpace($outputPathBox.Text) -and (Test-Path -LiteralPath $outputPathBox.Text -PathType Container)) {
        $dialog.SelectedPath = $outputPathBox.Text
    }

    if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $outputPathBox.Text = $dialog.SelectedPath
        Save-AppSettings -InputFile $inputPathBox.Text -OutputDirectory $outputPathBox.Text -ProcessType $processTypeBox.SelectedItem -ReceivableAccount $receivableAccountBox.Text -DepositAccount $depositAccountBox.Text -IncomeAccount $incomeAccountBox.Text
    }
})

function Get-FormInput {
    if ([string]::IsNullOrWhiteSpace($inputPathBox.Text)) {
        throw 'Choose a rent transaction report first.'
    }

    if ([string]::IsNullOrWhiteSpace($outputPathBox.Text)) {
        throw 'Choose an output folder first.'
    }

    [pscustomobject]@{
        InputPath = $inputPathBox.Text
        OutputDirectory = $outputPathBox.Text
        ProcessType = [string]$processTypeBox.SelectedItem
        ProcessingDate = Get-ValidatedDate -Text $processingDateBox.Text
        ReceivableAccount = $receivableAccountBox.Text
        DepositAccount = $depositAccountBox.Text
        IncomeAccount = $incomeAccountBox.Text
    }
}

$previewButton.Add_Click({
    $previewButton.Enabled = $false
    $statusLabel.Text = 'Building preview...'
    $log.Text = ''
    $form.Refresh()

    try {
        $input = Get-FormInput
        $data = New-RentIifData `
            -InputPath $input.InputPath `
            -ProcessingDate $input.ProcessingDate `
            -ProcessType $input.ProcessType `
            -ReceivableAccount $input.ReceivableAccount `
            -DepositAccount $input.DepositAccount `
            -IncomeAccount $input.IncomeAccount

        $previewGrid.DataSource = [System.Collections.ArrayList]::new($data.PreviewRows)
        $statusLabel.Text = 'Preview ready.'
        $log.Text = @(
            "Preview rows: $($data.ProcessedRows)"
            "Skipped rows: $($data.SkippedRows)"
            'Review the rows below before creating the IIF file.'
        ) -join [Environment]::NewLine
        Save-AppSettings -InputFile $input.InputPath -OutputDirectory $input.OutputDirectory -ProcessType $input.ProcessType -ReceivableAccount $input.ReceivableAccount -DepositAccount $input.DepositAccount -IncomeAccount $input.IncomeAccount
    }
    catch {
        $statusLabel.Text = 'Could not build preview.'
        $previewGrid.DataSource = $null
        $log.Text = $_ | Out-String
    }
    finally {
        $previewButton.Enabled = $true
    }
})

$createButton.Add_Click({
    $createButton.Enabled = $false
    $previewButton.Enabled = $false
    $openFolderButton.Enabled = $false
    $openIifButton.Enabled = $false
    $copyPathButton.Enabled = $false
    $statusLabel.Text = 'Creating IIF file...'
    $log.Text = ''
    $form.Refresh()

    try {
        $input = Get-FormInput
        $result = Convert-RentReportToIif `
            -InputPath $input.InputPath `
            -OutputDirectory $input.OutputDirectory `
            -ProcessingDate $input.ProcessingDate `
            -ProcessType $input.ProcessType `
            -ReceivableAccount $input.ReceivableAccount `
            -DepositAccount $input.DepositAccount `
            -IncomeAccount $input.IncomeAccount

        $script:lastOutputDirectory = Split-Path -Parent $result.OutputPath
        $script:lastOutputPath = $result.OutputPath
        $previewGrid.DataSource = [System.Collections.ArrayList]::new($result.PreviewRows)

        $statusLabel.Text = 'IIF file created.'
        $log.Text = @(
            'Created QuickBooks IIF file.'
            "Output: $($result.OutputPath)"
            "Processed $($input.ProcessType.ToLowerInvariant()) rows: $($result.ProcessedRows)"
            "Skipped rows: $($result.SkippedRows)"
            ''
            'The output is plain tab-delimited text with an .iif extension.'
        ) -join [Environment]::NewLine
        $openFolderButton.Enabled = $true
        $openIifButton.Enabled = $true
        $copyPathButton.Enabled = $true
        Save-AppSettings -InputFile $input.InputPath -OutputDirectory $input.OutputDirectory -ProcessType $input.ProcessType -ReceivableAccount $input.ReceivableAccount -DepositAccount $input.DepositAccount -IncomeAccount $input.IncomeAccount
    }
    catch {
        $statusLabel.Text = 'Could not create IIF file.'
        $log.Text = $_ | Out-String
    }
    finally {
        $createButton.Enabled = $true
        $previewButton.Enabled = $true
    }
})

$openFolderButton.Add_Click({
    if ($script:lastOutputDirectory -and (Test-Path -LiteralPath $script:lastOutputDirectory -PathType Container)) {
        Start-Process explorer.exe -ArgumentList $script:lastOutputDirectory
    }
})

$openIifButton.Add_Click({
    if ($script:lastOutputPath -and (Test-Path -LiteralPath $script:lastOutputPath -PathType Leaf)) {
        Start-Process notepad.exe -ArgumentList $script:lastOutputPath
    }
})

$copyPathButton.Add_Click({
    if ($script:lastOutputPath) {
        [System.Windows.Forms.Clipboard]::SetText($script:lastOutputPath)
        $statusLabel.Text = 'Output path copied to clipboard.'
    }
})

[void]$form.ShowDialog()
