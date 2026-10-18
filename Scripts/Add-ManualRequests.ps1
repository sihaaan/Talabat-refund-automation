#requires -Version 5.1
[CmdletBinding()]param([Parameter(Mandatory)][string]$InputCsv,[string]$Workspace,[string]$AsOfDate,[switch]$ReviewOnly,[string]$EnrichmentCsv)
$ErrorActionPreference='Stop'
if(-not $Workspace){$Workspace=Split-Path $PSScriptRoot -Parent}
& (Join-Path $PSScriptRoot 'Prepare-Refunds.ps1') -Workspace $Workspace -InputCsv $InputCsv -Source manual -AsOfDate $AsOfDate -ReviewOnly:$ReviewOnly -EnrichmentCsv $EnrichmentCsv
