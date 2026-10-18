#requires -Version 5.1
[CmdletBinding()]param(
 [string]$Workspace,
 [Parameter(Mandatory)][string[]]$RecordId,
 [Parameter(Mandatory)][ValidateSet('Email','Portal','CLI','Acquirer')][string]$Method,
 [Parameter(Mandatory)][string]$Evidence,
 [Parameter(Mandatory)][string]$ApprovalReference,
 [Parameter(Mandatory)][string]$AttemptedAt,
 [switch]$ConfirmActualSubmission
)
$ErrorActionPreference='Stop'
if(-not $Workspace){$Workspace=Split-Path $PSScriptRoot -Parent}
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
if(-not $ConfirmActualSubmission -or [string]::IsNullOrWhiteSpace($Evidence) -or [string]::IsNullOrWhiteSpace($ApprovalReference)){throw 'Record only an already-occurred, authorized submission with evidence. This command never sends a request.'}
if($AttemptedAt -notmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?(Z|[+-]\d{2}:\d{2})$'){throw 'AttemptedAt must be an ISO timestamp with timezone, for example 2026-10-19T09:30:00+04:00.'}
$occurred=[datetimeoffset]::Parse($AttemptedAt,[Globalization.CultureInfo]::InvariantCulture)
if($occurred.UtcDateTime -gt [datetime]::UtcNow.AddMinutes(5)){throw 'An actual submission cannot be in the future.'}
$s=Open-Workspace $Workspace
try{
 $seen=New-Map;$records=@();$attempts=@();$events=@();$recorded=[datetime]::UtcNow.ToString('o')
 foreach($id in $RecordId){
  if(-not $s.Records.ContainsKey($id) -or $seen.ContainsKey($id)){throw 'Unknown or duplicate RecordId.'};$seen[$id]=$true;$r=$s.Records[$id]
  $aid='EXT-'+(Get-Sha $id).Substring(0,24)
  if($s.Attempts.ContainsKey($aid)){
   if($s.Attempts[$aid].Evidence -ceq $Evidence -and $s.Attempts[$aid].Method -ceq $Method -and $s.Attempts[$aid].AttemptedAt -ceq $occurred.UtcDateTime.ToString('o')){continue}
   throw 'External attempt evidence already exists. Record a communication or final outcome instead of another attempt.'
  }
  if($r.ActiveBatchId -or $r.SubmissionStatus -notin @('NotSubmitted','LegacyAttempt') -or $r.FinalStatus -ne 'Unknown'){throw 'Record already has batch/attempt/outcome history. Use its batch result or outcome workflow; never record a second refund request.'}
  $r.SubmissionStatus='Submitted';$r.SubmittedAt=$occurred.UtcDateTime.ToString('o');$r.SubmissionMethod=$Method;$r.ApprovalReference=$ApprovalReference;$r.Evidence=$Evidence;$r.UpdatedAt=$recorded
  if($Method -in @('Email','Acquirer')){$r.EmailSentAt=$r.SubmittedAt}
  $records+=$r
  $attempts+=[pscustomobject]@{AttemptId=$aid;RecordId=$id;BatchId='EXTERNAL-EVIDENCE';RunId=$r.RunId;AttemptedAt=$r.SubmittedAt;Method=$Method;SubmissionStatus='Submitted';ActionId='';FailureReason='';Evidence=$Evidence;TrackerEligible=$true}
  $events+=[pscustomobject]@{At=$recorded;Type='ExternalAttemptEvidence';RecordId=$id;BatchId='';Reference=$Evidence;Detail="Actual time=$($r.SubmittedAt); Method=$Method; Approval=$ApprovalReference. No request sent by automation."}
 }
 if($attempts.Count){Save-Transaction $s 'ExternalAttemptEvidence' $records @() $attempts $events;Export-Register $s}
 Write-Host "Recorded $($attempts.Count) already-occurred external attempts. No refund request was sent. Repeat evidence does not create a new attempt."
}finally{$s.Lock.Dispose()}
