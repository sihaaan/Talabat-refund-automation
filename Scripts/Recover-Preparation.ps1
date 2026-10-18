#requires -Version 5.1
[CmdletBinding()]param([string]$Workspace)
$ErrorActionPreference='Stop'
if(-not $Workspace){$Workspace=Split-Path $PSScriptRoot -Parent}
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
$s=Open-Workspace $Workspace -AllowPending
try{
 $path=Join-Path $s.Root 'Register\pending-preparation.json'
 if(-not(Test-Path -LiteralPath $path)){Write-Host 'No interrupted preparation.';return}
 $p=Get-Content -LiteralPath $path -Raw|ConvertFrom-Json
 if($s.Sequence -eq $p.SequenceBefore){
  $dirs=@(Get-ChildItem -LiteralPath (Join-Path $s.Root 'Output') -Directory -Recurse|Where-Object Name -eq $p.RunId)
  foreach($dir in $dirs){Write-AtomicText (Join-Path $dir.FullName 'UNCOMMITTED_DO_NOT_SUBMIT.txt') 'This preparation did not commit. Files are invalid. Prepare a new run. Start-Submission will reject unknown batches.'}
  Save-Transaction $s 'RecoverAbandonedPreparation' @() @() @() @([pscustomobject]@{At=[datetime]::UtcNow.ToString('o');Type='AbandonedPreparation';RecordId='';BatchId='';Reference=$p.SourceHash;Detail=$p.RunId})
  Write-Host 'Uncommitted run invalidated. No attempt inferred. Prepare again; never use files from the abandoned run.'
 }else{
  if(@($s.Events|Where-Object{$_.Type -eq 'Preparation' -and $_.Detail -like ('Run='+$p.RunId+';*')}).Count -ne 1){throw 'Recovery could not identify the committed run. Stop and investigate.'}
  Write-Host 'Committed preparation recovered. Existing active batches remain reserved for their records; use the original verified files.'
 }
 Export-Register $s
 Remove-Item -LiteralPath $path
}finally{$s.Lock.Dispose()}
