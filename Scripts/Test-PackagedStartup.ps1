#requires -Version 5.1
[CmdletBinding()]param([string]$TestOutputFolder=(Join-Path $env:TEMP ('TalabatPackageTest-'+[guid]::NewGuid().ToString('N'))))
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
$engine=(Get-Process -Id $PID).Path
$packageFolder=Join-Path $TestOutputFolder 'Package'
$extracted=Join-Path $TestOutputFolder 'Extracted'
if(Test-Path -LiteralPath $TestOutputFolder){throw 'Use a new isolated test folder.'}
function Invoke-TestScript([string]$Name,[string[]]$Arguments=@()) {
    $output=& $engine -NoProfile -File (Join-Path $extracted ('Scripts\'+$Name)) @Arguments 2>&1
    if($LASTEXITCODE -ne 0){throw "$Name failed: $($output -join ' ')"}
}
& (Join-Path $PSScriptRoot 'Build-Release.ps1') -Version smoke -DestinationFolder $packageFolder 6>$null
Add-Type -AssemblyName System.IO.Compression.FileSystem
[IO.Compression.ZipFile]::ExtractToDirectory((Join-Path $packageFolder 'TalabatRefundAutomation-smoke.zip'),$extracted)
Invoke-TestScript 'Initialize-Workspace.ps1'
if(-not(Test-Path -LiteralPath (Join-Path $extracted 'Register\workspace.json'))){throw 'Default workspace initialization failed.'}
Copy-Item -LiteralPath (Join-Path $extracted 'Examples\Tracker_Fake.csv') -Destination (Join-Path $extracted 'Input\Tracker.csv')
Invoke-TestScript 'Prepare-Refunds.ps1' @('-ReviewOnly')
Invoke-TestScript 'Generate-Reports.ps1'
$run=@(Get-ChildItem -LiteralPath (Join-Path $extracted 'Reports') -Directory -Filter 'RUN_*')[0]
$inputCount=@(Read-CsvTable (Join-Path $extracted 'Input\Tracker.csv') @('Order ID','Refund Status') 4).Count
if(@(Import-Csv -LiteralPath (Join-Path $run.FullName 'Row_Dispositions.csv')).Count -ne $inputCount){throw 'Packaged preparation lost input rows.'}
if(@(Import-Csv -LiteralPath (Join-Path $run.FullName 'Batches.csv')).Count -ne 0){throw 'Unconfirmed package produced submission batches.'}
$dashboard=@(Get-ChildItem -LiteralPath (Join-Path $extracted 'Reports') -Directory -Filter 'Dashboard_*')[0]
$manifest=Join-Path $dashboard.FullName 'Tracker_Update_Manifest.csv'
if(@(Import-Csv -LiteralPath $manifest).Count -ne 0){throw 'Review-only package reported actual attempts.'}
Invoke-TestScript 'Match-TrackerUpdates.ps1' @('-TrackerCsv',(Join-Path $extracted 'Input\Tracker.csv'),'-ManifestCsv',$manifest)
if(-not(Test-Path -LiteralPath (Join-Path $extracted 'Reports\TrackerMatching'))){throw 'Default matching destination failed.'}
Write-CsvFile (Join-Path $TestOutputFolder 'Test_Results.csv') @([pscustomobject]@{Test='Extracted ZIP starts with default workspace and completes review, reports and matching';Result='PASS';Detail='Synthetic data; no external integration'}) @('Test','Result','Detail')
Write-Host "Packaged startup: PASS. Files: $TestOutputFolder"
