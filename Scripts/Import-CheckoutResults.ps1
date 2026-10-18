#requires -Version 5.1
[CmdletBinding()]param([string]$Workspace,[Parameter(Mandatory)][string]$BatchId,[Parameter(Mandatory)][string]$ResultCsv,[Parameter(Mandatory)][string]$Evidence,[switch]$ConfirmProductionAttempt)
$ErrorActionPreference='Stop'
if(-not $Workspace){$Workspace=Split-Path $PSScriptRoot -Parent}
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
if(-not $ConfirmProductionAttempt -or [string]::IsNullOrWhiteSpace($Evidence)){throw 'Only real production attempt evidence may enter this register. Dry-run/sandbox outputs belong in a separate test workspace.'}
$s=Open-Workspace $Workspace
try{
 if(-not $s.Batches.ContainsKey($BatchId)){throw 'Unknown batch.'};$b=$s.Batches[$BatchId]
 if($b.PSP -ne 'Checkout' -or $b.Status -notin @('Reserved','Uncertain','Imported')){throw 'Checkout results require a reserved or previously imported/uncertain Checkout batch.'}
 $hash=Get-FileSha $ResultCsv;$importKey=$BatchId+'|'+$hash
 if($s.Imports.ContainsKey($importKey)){Write-Host 'This result file was already imported. No changes.';return}
 $rows=@(Read-CsvTable $ResultCsv @('id','amount','refund_initiated','refund_action_id','refund_error') 1)
 if(-not $rows.Count){throw 'Empty result file. Record the batch as Uncertain instead.'}
 $expected=New-Map;foreach($id in $b.RecordIds){$r=$s.Records[$id];$expected[$r.PaymentId]=$r}
 $found=New-Map;$next=New-Map
 foreach($row in $rows){
  $id=Get-Text $row 'id'
  if(-not $expected.ContainsKey($id) -or $found.ContainsKey($id)){throw 'Unexpected or repeated payment ID in results. No results imported.'}
  $found[$id]=$true;$r=$expected[$id]
  if((Get-Text $row 'amount') -cne $r.AmountMinor){throw 'Result amount differs from the approved batch.'}
  $flag=(Get-Text $row 'refund_initiated').Trim().ToLowerInvariant();$action=(Get-Text $row 'refund_action_id').Trim();$errorText=Get-Text $row 'refund_error'
  if($flag -notin @('true','false')){throw 'Invalid refund_initiated value; expected true or false.'}
  if($flag -eq 'true' -and (-not $action -or $errorText)){throw 'Inconsistent initiated result. Investigate before importing.'}
  $status=if($flag -eq 'true'){'Submitted'}elseif($errorText){'RequestFailed'}else{'Uncertain'}
  $aid='AT-'+(Get-Sha ($BatchId+'|'+$r.RecordId)).Substring(0,24)
  if($s.Attempts.ContainsKey($aid)){
   $prior=$s.Attempts[$aid]
   if($prior.SubmissionStatus -ne 'Uncertain' -and ($prior.SubmissionStatus -cne $status -or $prior.ActionId -cne $action -or $prior.FailureReason -cne $errorText)){throw 'Result conflicts with recorded evidence. Investigate; do not replace an established attempt.'}
  }
  $next[$r.RecordId]=[pscustomobject]@{Status=$status;Action=$action;Error=$errorText;TrackerEligible=($flag -eq 'true')}
 }
 $records=@();$attempts=@();$at=[datetime]::UtcNow.ToString('o')
 foreach($rid in $b.RecordIds){
  $r=$s.Records[$rid];$aid='AT-'+(Get-Sha ($BatchId+'|'+$rid)).Substring(0,24)
  if(-not $next.ContainsKey($rid)){
   if($s.Attempts.ContainsKey($aid) -and $s.Attempts[$aid].SubmissionStatus -ne 'Uncertain'){continue}
   $n=[pscustomobject]@{Status='Uncertain';Action='';Error='Row missing from result file; do not retry.';TrackerEligible=$false}
  }else{$n=$next[$rid]}
  $attemptAt=if($s.Attempts.ContainsKey($aid)){$s.Attempts[$aid].AttemptedAt}else{$at}
  $attempts+=[pscustomobject]@{AttemptId=$aid;RecordId=$rid;BatchId=$BatchId;RunId=$b.RunId;AttemptedAt=$attemptAt;Method='CLI';SubmissionStatus=$n.Status;ActionId=$n.Action;FailureReason=$n.Error;Evidence=$Evidence;TrackerEligible=$n.TrackerEligible}
  $r.SubmissionStatus=$n.Status;$r.SubmissionMethod='CLI';$r.SubmittedAt=$attemptAt;$r.ActionId=$n.Action;$r.FailureReason=$n.Error;$r.Evidence=$Evidence;$r.UpdatedAt=$at;$records+=$r
 }
 $evidenceDir=Join-Path $s.Root ('Register\Evidence\'+$BatchId);[IO.Directory]::CreateDirectory($evidenceDir)|Out-Null
 $copy=Join-Path $evidenceDir ($hash+'.csv');if(-not(Test-Path -LiteralPath $copy)){[IO.File]::Copy([IO.Path]::GetFullPath($ResultCsv),$copy)}
 $b.Status='Imported';$b.Evidence=$Evidence
 Save-Transaction $s 'ImportCheckout' $records @($b) $attempts @([pscustomobject]@{At=$at;Type='CheckoutResult';RecordId='';BatchId=$BatchId;Reference=$Evidence;Detail="SHA256=$hash; Imported=$($rows.Count); Missing=$($b.RecordIds.Count-$rows.Count). Initiation is not final success."}) @($importKey)
 Export-Register $s
 Write-Host 'Results imported without resending. Submitted = accepted request, not final success. False/missing rows remain blocked; false alone does not prove a customer refund failed.'
}finally{$s.Lock.Dispose()}
