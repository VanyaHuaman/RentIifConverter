#Requires -Version 7.0
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$converter = Join-Path $repoRoot 'RentIifConverter.ps1'
$inputDir = Join-Path $PSScriptRoot 'fixtures\input'
$expectedDir = Join-Path $PSScriptRoot 'fixtures\expected'
$outputDir = Join-Path $PSScriptRoot 'output'

if (Test-Path -LiteralPath $outputDir -PathType Container) {
    Remove-Item -LiteralPath $outputDir -Recurse -Force
}
New-Item -Path $outputDir -ItemType Directory -Force | Out-Null

function Invoke-ConverterCase {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [ValidateSet('Payment', 'Invoice', 'Both')]
        [string]$ProcessType,

        [Parameter(Mandatory)]
        [string]$InputFile,

        [Parameter(Mandatory)]
        [string]$ProcessingDate,

        [Parameter(Mandatory)]
        [string[]]$ExpectedFiles
    )

    Write-Host "Running $Name..."

    & pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File $converter `
        -NoGui `
        -ProcessType $ProcessType `
        -InputPath (Join-Path $inputDir $InputFile) `
        -OutputDirectory $outputDir `
        -ProcessingDate $ProcessingDate | Out-Host

    if ($LASTEXITCODE -ne 0) {
        throw "$Name failed with exit code $LASTEXITCODE."
    }

    foreach ($file in $ExpectedFiles) {
        $expectedPath = Join-Path $expectedDir $file
        $actualPath = Join-Path $outputDir $file

        if (-not (Test-Path -LiteralPath $actualPath -PathType Leaf)) {
            throw "$Name did not create expected file: $file"
        }

        $expectedBytes = [IO.File]::ReadAllBytes($expectedPath)
        $actualBytes = [IO.File]::ReadAllBytes($actualPath)

        if ($expectedBytes.Length -ne $actualBytes.Length) {
            throw "$Name output length mismatch for $file. Expected $($expectedBytes.Length), got $($actualBytes.Length)."
        }

        for ($i = 0; $i -lt $expectedBytes.Length; $i++) {
            if ($expectedBytes[$i] -ne $actualBytes[$i]) {
                throw "$Name output mismatch for $file at byte $i. Expected $($expectedBytes[$i]), got $($actualBytes[$i])."
            }
        }
    }
}

Invoke-ConverterCase `
    -Name 'Payment report' `
    -ProcessType Payment `
    -InputFile 'payment-06122026.xlsx' `
    -ProcessingDate 06122026 `
    -ExpectedFiles @('RentPayment06122026.iif')

Invoke-ConverterCase `
    -Name 'Invoice report' `
    -ProcessType Invoice `
    -InputFile 'invoice-06112026.xlsx' `
    -ProcessingDate 06112026 `
    -ExpectedFiles @('RentInvoice06112026.iif')

Invoke-ConverterCase `
    -Name 'Mixed report' `
    -ProcessType Both `
    -InputFile 'mixed-06262026.xlsx' `
    -ProcessingDate 06262026 `
    -ExpectedFiles @('RentInvoice06262026.iif', 'RentPayment06262026.iif')

Write-Host 'All regression tests passed.'
