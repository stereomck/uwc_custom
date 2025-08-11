# Run-Tests.ps1
# Script to run Pester tests for OCRFunctions.ps1

param(
    [string]$OutputFormat = "NUnitXml",
    [string]$OutputFile = "TestResults.xml",
    [switch]$PassThru,
    [string[]]$Tag,
    [string[]]$ExcludeTag,
    [switch]$CodeCoverage
)

# Check if Pester is installed
if (-not (Get-Module -ListAvailable -Name Pester)) {
    Write-Host "Pester module not found. Installing Pester..." -ForegroundColor Yellow
    Install-Module -Name Pester -Force -SkipPublisherCheck
}

# Import Pester
Import-Module Pester -Force

# Set up test configuration
$PesterConfiguration = New-PesterConfiguration

# Basic configuration - updated path to point to tests directory
$PesterConfiguration.Run.Path = "$PSScriptRoot\OCRFunctions.Tests.ps1"
$PesterConfiguration.Run.PassThru = $PassThru

# Output configuration
if ($OutputFormat -and $OutputFile) {
    $PesterConfiguration.TestResult.Enabled = $true
    $PesterConfiguration.TestResult.OutputFormat = $OutputFormat
    $PesterConfiguration.TestResult.OutputPath = Join-Path $PSScriptRoot $OutputFile
}

# Filter configuration
if ($Tag) {
    $PesterConfiguration.Filter.Tag = $Tag
}
if ($ExcludeTag) {
    $PesterConfiguration.Filter.ExcludeTag = $ExcludeTag
}

# Code coverage configuration - updated path to point to parent directory
if ($CodeCoverage) {
    $PesterConfiguration.CodeCoverage.Enabled = $true
    $PesterConfiguration.CodeCoverage.Path = "$PSScriptRoot\..\OCRFunctions.ps1"
    $PesterConfiguration.CodeCoverage.OutputFormat = "JaCoCo"
    $PesterConfiguration.CodeCoverage.OutputPath = Join-Path $PSScriptRoot "CodeCoverage.xml"
}

# Output configuration for better visibility
$PesterConfiguration.Output.Verbosity = "Detailed"

Write-Host "Running Pester tests..." -ForegroundColor Green
Write-Host "Test file: $PSScriptRoot\OCRFunctions.Tests.ps1" -ForegroundColor Cyan
Write-Host "Target module: $PSScriptRoot\..\OCRFunctions.ps1" -ForegroundColor Cyan
Write-Host "Output format: $OutputFormat" -ForegroundColor Cyan
if ($OutputFile) {
    Write-Host "Output file: $OutputFile" -ForegroundColor Cyan
}

# Run the tests
$TestResults = Invoke-Pester -Configuration $PesterConfiguration

# Display summary
Write-Host "`n" -NoNewline
Write-Host "=" * 50 -ForegroundColor Green
Write-Host "TEST SUMMARY" -ForegroundColor Green
Write-Host "=" * 50 -ForegroundColor Green
Write-Host "Total Tests: $($TestResults.TotalCount)" -ForegroundColor White
Write-Host "Passed: $($TestResults.PassedCount)" -ForegroundColor Green
Write-Host "Failed: $($TestResults.FailedCount)" -ForegroundColor Red
Write-Host "Skipped: $($TestResults.SkippedCount)" -ForegroundColor Yellow
Write-Host "Duration: $($TestResults.Duration)" -ForegroundColor White

if ($TestResults.FailedCount -gt 0) {
    Write-Host "`nFailed Tests:" -ForegroundColor Red
    foreach ($test in $TestResults.Failed) {
        Write-Host "  - $($test.FullName)" -ForegroundColor Red
        if ($test.ErrorRecord) {
            Write-Host "    Error: $($test.ErrorRecord.Exception.Message)" -ForegroundColor DarkRed
        }
    }
}

# Return appropriate exit code
if ($TestResults.FailedCount -eq 0) {
    Write-Host "`nAll tests passed! ✅" -ForegroundColor Green
    exit 0
} else {
    Write-Host "`nSome tests failed! ❌" -ForegroundColor Red
    exit 1
}