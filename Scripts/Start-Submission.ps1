#requires -Version 5.1
[CmdletBinding()]param([string]$Workspace,[Parameter(Mandatory)][string]$BatchId,[Parameter(Mandatory)][string]$ApprovalReference,[switch]$ConfirmBusinessAuthorization)
$ErrorActionPreference='Stop'
if(-not $Workspace){$Workspace=Split-Path $PSScriptRoot -Parent}
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
if(-not $ConfirmBusinessAuthorization -or [string]::IsNullOrWhiteSpace($ApprovalReference)){throw 'Explicit business authorization and an approval reference are required.'}
$s=Open-Workspace $Workspace
try{
 if(-not $s.Batches.ContainsKey($BatchId)){throw 'Unknown batch.'};$b=$s.Batches[$BatchId]
 if($b.Status -ne 'Prepared'){throw 'Batch is not Prepared. A reserved or attempted batch cannot be started again.'}
 if(-not(Test-Path -LiteralPath $b.FilePath) -or (Get-FileSha $b.FilePath) -cne $b.FileHash){throw 'Submission file is missing or changed. Stop; do not submit.'}
 if((Get-FileSha (Join-Path $s.Root 'Config\config.json')) -cne $b.ConfigHash){throw 'Configuration changed since preparation. Cancel this unsubmitted batch and prepare again.'}
 $now=Get-LocalNow $s.Config;$changed=@()
 foreach($id in $b.RecordIds){$r=$s.Records[$id];$days=[int](Get-Setting $s.Config.psps $r.PSP).delayedDays;if(($now.Date-(Get-DateValue $r.TransactionDate)).Days -ge $days){throw 'Batch now includes delayed transactions. Cancel it and prepare again.'};if($r.SubmissionStatus -ne 'NotSubmitted' -or $r.FinalStatus -ne 'Unknown' -or $r.ActiveBatchId -ne $BatchId -or (Get-Text $r 'HoldReason')){throw 'A batch record is blocked/on hold or no longer belongs to this batch.'};$changed+=$r}
 $b.Status='Reserved';$b.ReservedAt=[datetime]::UtcNow.ToString('o');$b.ApprovalReference=$ApprovalReference
 foreach($r in $changed){$r.SubmissionStatus='Reserved';$r.ApprovalReference=$ApprovalReference}
 Save-Transaction $s 'ReserveBeforeManualSubmission' $changed @($b) @() @([pscustomobject]@{At=$b.ReservedAt;Type='Reserved';RecordId='';BatchId=$BatchId;Reference=$ApprovalReference;Detail='Not an attempt. If submission outcome is uncertain, do not release or rerun.'})
 Export-Register $s
 Write-Host "Reserved $BatchId. This script did not submit anything. Manually use the approved file once, then record evidence. If interrupted, investigate; do not rerun."
}finally{$s.Lock.Dispose()}
