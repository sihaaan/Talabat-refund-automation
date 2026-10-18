#requires -Version 5.1
[CmdletBinding()]param([string]$Workspace,[Parameter(Mandatory)][string]$BatchId,[Parameter(Mandatory)][string]$Evidence,[ValidateSet('Submitted','Uncertain')][string]$State='Submitted',[switch]$ConfirmActualSubmission)
$ErrorActionPreference='Stop'
if(-not $Workspace){$Workspace=Split-Path $PSScriptRoot -Parent}
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
if([string]::IsNullOrWhiteSpace($Evidence)){throw 'Evidence/reference is required.'}
if($State -eq 'Submitted' -and -not $ConfirmActualSubmission){throw 'Confirm that the exact batch was actually submitted. Prepared files do not count.'}
$s=Open-Workspace $Workspace
try{
 if(-not $s.Batches.ContainsKey($BatchId)){throw 'Unknown batch.'};$b=$s.Batches[$BatchId]
 if($b.Status -ne 'Reserved'){throw 'Only a reserved batch can be recorded. Repeating this command cannot create a second attempt.'}
 if($b.PSP -eq 'Checkout' -and $State -eq 'Submitted'){throw 'Import Checkout result rows instead. For a missing result use -State Uncertain.'}
 $at=[datetime]::UtcNow.ToString('o');$records=@();$attempts=@()
 foreach($id in $b.RecordIds){
  $r=$s.Records[$id];$r.SubmissionStatus=$State;$r.SubmissionMethod=$b.Method;$r.SubmittedAt=$at;$r.Evidence=$Evidence;$r.UpdatedAt=$at
  if($b.Method -eq 'Email' -and $State -eq 'Submitted'){$r.EmailSentAt=$at}
  $attempts+=[pscustomobject]@{AttemptId='AT-'+(Get-Sha ($BatchId+'|'+$id)).Substring(0,24);RecordId=$id;BatchId=$BatchId;RunId=$b.RunId;AttemptedAt=$at;Method=$b.Method;SubmissionStatus=$State;ActionId='';FailureReason='';Evidence=$Evidence;TrackerEligible=($State -eq 'Submitted')}
  $records+=$r
 }
 $b.Status=$State;$b.Evidence=$Evidence
 Save-Transaction $s 'RecordSubmission' $records @($b) $attempts @([pscustomobject]@{At=$at;Type='Submission';RecordId='';BatchId=$BatchId;Reference=$Evidence;Detail=$State})
 Export-Register $s
 Write-Host "Recorded $State. No payment was sent. Run Generate-Reports to refresh the tracker manifest."
}finally{$s.Lock.Dispose()}
