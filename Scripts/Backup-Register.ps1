#requires -Version 5.1
[CmdletBinding()]param([string]$Workspace,[Parameter(Mandatory)][string]$DestinationFolder)
$ErrorActionPreference='Stop'
if(-not $Workspace){$Workspace=Split-Path $PSScriptRoot -Parent}
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
$s=Open-Workspace $Workspace
try{
 $dest=[IO.Path]::GetFullPath($DestinationFolder);$reg=[IO.Path]::GetFullPath((Join-Path $s.Root 'Register'))
 if($dest.StartsWith($reg,[StringComparison]::OrdinalIgnoreCase)){throw 'Choose a backup location outside Register.'}
 [IO.Directory]::CreateDirectory($dest)|Out-Null
 $name='Register_'+[datetime]::UtcNow.ToString('yyyyMMdd_HHmmss')+'_'+[guid]::NewGuid().ToString('N').Substring(0,6)+'.zip';$zip=Join-Path $dest $name
 Add-Type -AssemblyName System.IO.Compression;Add-Type -AssemblyName System.IO.Compression.FileSystem
 $archive=[IO.Compression.ZipFile]::Open($zip,[IO.Compression.ZipArchiveMode]::Create)
 try{foreach($f in @(Get-ChildItem -LiteralPath $reg -File -Recurse|Where-Object Name -ne '.lock')){[IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive,$f.FullName,$f.FullName.Substring($reg.Length+1).Replace('\','/'))|Out-Null}}finally{$archive.Dispose()}
 Write-Host "Register backup created while holding the writer lock: $zip"
}finally{$s.Lock.Dispose()}
