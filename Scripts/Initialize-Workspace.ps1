#requires -Version 5.1
[CmdletBinding()]param([string]$Workspace)
$ErrorActionPreference='Stop'
if(-not $Workspace){$Workspace=Split-Path $PSScriptRoot -Parent}
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
$Workspace=[IO.Path]::GetFullPath($Workspace)
$reg=Join-Path $Workspace 'Register'
if(Test-Path -LiteralPath (Join-Path $reg 'workspace.json')){Write-Host 'Workspace already initialized. Existing history preserved.';return}
if(Test-Path -LiteralPath $reg){$existing=@(Get-ChildItem -LiteralPath $reg -Force|Where-Object Name -ne '.gitkeep');if($existing.Count){throw 'Register contains files but has no workspace marker. Restore the original register; do not reinitialize.'}}
foreach($d in @('Input','Register\Journal','Output','Reports')){[IO.Directory]::CreateDirectory((Join-Path $Workspace $d))|Out-Null}
if(-not(Test-Path -LiteralPath (Join-Path $Workspace 'Config\config.json'))){throw 'Config/config.json missing. Extract the complete application first.'}
Write-AtomicText (Join-Path $reg 'head.json') (@{Sequence=0;Hash='GENESIS'}|ConvertTo-Json)
Write-AtomicText (Join-Path $reg 'workspace.json') (@{DatasetId=[guid]::NewGuid().ToString();CreatedAt=[datetime]::UtcNow.ToString('o');Warning='Preserve this folder across all weeks. Do not replace with an empty register.'}|ConvertTo-Json)
Write-Host "Initialized $Workspace. Import the full tracker to preserve historical attempt signals."
