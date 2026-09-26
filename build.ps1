#Requires -Version 7.2
<#
.SYNOPSIS
    Local and CI entry point for linting and testing the OpenDataSync module.

.DESCRIPTION
    CI runs exactly the same commands a developer runs locally, so a green workstation run
    means a green pipeline.

.PARAMETER Task
    Analyze runs PSScriptAnalyzer, Test runs Pester, All runs both.

.PARAMETER CI
    Emit NUnit and JaCoCo result files for the workflow to publish, fail on any analyzer
    finding, and enforce the code coverage target.

.PARAMETER CoverageTarget
    Minimum percentage of module commands that must be covered when -CI is used.

.EXAMPLE
    ./build.ps1 -Task All
#>
[CmdletBinding()]
param(
    [ValidateSet('Analyze', 'Test', 'All')]
    [string]$Task = 'All',

    [switch]$CI,

    [ValidateRange(0, 100)]
    [int]$CoverageTarget = 80
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
$sourcePath = Join-Path -Path $root -ChildPath 'src'
$testPath = Join-Path -Path $root -ChildPath 'tests'
$outputPath = Join-Path -Path $root -ChildPath 'out'
$settingsPath = Join-Path -Path $root -ChildPath 'PSScriptAnalyzerSettings.psd1'

$failed = $false

if ($Task -in 'Analyze', 'All') {
    Write-Host '==> PSScriptAnalyzer' -ForegroundColor Cyan
    Import-Module -Name PSScriptAnalyzer -ErrorAction Stop

    $findings = @(
        Invoke-ScriptAnalyzer -Path $sourcePath -Recurse -Settings $settingsPath
        Invoke-ScriptAnalyzer -Path $testPath -Recurse -Settings $settingsPath
        Invoke-ScriptAnalyzer -Path (Join-Path -Path $root -ChildPath 'examples') -Recurse -Settings $settingsPath
    )

    if ($findings.Count -gt 0) {
        $findings | Format-Table -AutoSize -Property Severity, RuleName, ScriptName, Line, Message |
            Out-String -Width 200 | Write-Host
        $failed = $true
        Write-Host "PSScriptAnalyzer reported $($findings.Count) finding(s)." -ForegroundColor Red
    }
    else {
        Write-Host 'No analyzer findings.' -ForegroundColor Green
    }
}

if ($Task -in 'Test', 'All') {
    Write-Host '==> Pester' -ForegroundColor Cyan
    Import-Module -Name Pester -MinimumVersion 5.0 -ErrorAction Stop

    $configuration = New-PesterConfiguration
    $configuration.Run.Path = $testPath
    $configuration.Run.PassThru = $true
    $configuration.Output.Verbosity = if ($CI) { 'Detailed' } else { 'Normal' }

    if ($CI) {
        if (-not (Test-Path -LiteralPath $outputPath)) {
            $null = New-Item -Path $outputPath -ItemType Directory -Force
        }

        $configuration.TestResult.Enabled = $true
        $configuration.TestResult.OutputPath = Join-Path -Path $outputPath -ChildPath 'testResults.xml'
        $configuration.CodeCoverage.Enabled = $true
        $configuration.CodeCoverage.Path = $sourcePath
        $configuration.CodeCoverage.OutputPath = Join-Path -Path $outputPath -ChildPath 'coverage.xml'
        $configuration.CodeCoverage.CoveragePercentTarget = $CoverageTarget
    }

    $result = Invoke-Pester -Configuration $configuration

    if ($result.FailedCount -gt 0) {
        $failed = $true
        Write-Host "$($result.FailedCount) test(s) failed." -ForegroundColor Red
    }
    else {
        Write-Host "$($result.PassedCount) test(s) passed." -ForegroundColor Green
    }

    if ($CI -and $result.CodeCoverage) {
        $analyzed = $result.CodeCoverage.CommandsAnalyzedCount
        $covered = $result.CodeCoverage.CommandsExecutedCount
        $percent = if ($analyzed -gt 0) { [Math]::Round(($covered / $analyzed) * 100, 2) } else { 0 }

        if ($percent -lt $CoverageTarget) {
            $failed = $true
            Write-Host "Code coverage $percent% is below the $CoverageTarget% target." -ForegroundColor Red
        }
        else {
            Write-Host "Code coverage $percent% meets the $CoverageTarget% target." -ForegroundColor Green
        }
    }
}

if ($failed) {
    exit 1
}

Write-Host '==> Build succeeded' -ForegroundColor Green
