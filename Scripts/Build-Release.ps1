#requires -Version 5.1
[CmdletBinding()]param([string]$Version='1.0.0',[string]$DestinationFolder)
$ErrorActionPreference='Stop'
if(-not $DestinationFolder){$DestinationFolder=(Join-Path (Split-Path $PSScriptRoot -Parent) 'dist')}
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
$root=[IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$allow=[Collections.Generic.List[object]]::new()
foreach($name in @('README.md','Run-Prepare.bat','Run-Reports.bat','.gitignore')){$allow.Add((Get-Item -LiteralPath (Join-Path $root $name)))}
foreach($pair in @(@('Scripts','*.ps1'),@('Scripts','*.psm1'),@('Config','*.json'),@('Templates','*.csv'),@('docs','*.md'),@('Examples','*.csv'),@('Examples','*.html'),@('Examples','*.txt'))){foreach($f in @(Get-ChildItem -LiteralPath (Join-Path $root $pair[0]) -File -Recurse -Filter $pair[1])){if($f.Name -notlike '*.local.json'){$allow.Add($f)}}}
foreach($f in $allow){if($f.FullName -match '[\\/](Input|Register|Output|Reports|PrivateReview)[\\/]' -or $f.Extension -in @('.exe','.env','.xlsx','.pdf','.mp4','.vtt')){throw 'Release allowlist contains a forbidden operational file.'}}
[IO.Directory]::CreateDirectory($DestinationFolder)|Out-Null
$zip=Join-Path $DestinationFolder ('TalabatRefundAutomation-'+$Version+'.zip')
if(Test-Path -LiteralPath $zip){throw 'Release ZIP already exists. Use a new version or a new destination; never overwrite a reviewed release.'}
Add-Type -AssemblyName System.IO.Compression;Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive=[IO.Compression.ZipFile]::Open($zip,[IO.Compression.ZipArchiveMode]::Create)
try{foreach($f in $allow){$relative=$f.FullName.Substring($root.Length+1).Replace('\','/');[IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive,$f.FullName,$relative)|Out-Null};foreach($dir in @('Input/','Register/','Output/','Reports/')){$archive.CreateEntry($dir)|Out-Null}}finally{$archive.Dispose()}
Write-AtomicText ($zip+'.sha256.txt') ((Get-FileSha $zip)+'  '+[IO.Path]::GetFileName($zip))
Write-Host "Created source-only release: $zip ($($allow.Count) files). Operational records and binaries excluded."
