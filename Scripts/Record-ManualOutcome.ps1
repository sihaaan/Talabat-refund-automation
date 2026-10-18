#requires -Version 5.1
[CmdletBinding()]param([string]$Workspace,[Parameter(Mandatory)][string]$RecordId,[Parameter(Mandatory)][ValidateSet('ConfirmedSuccessful','ConfirmedFailed','Voided')][string]$Status,[Parameter(Mandatory)][string]$Evidence,[string]$FailureReason,[switch]$Correction)
$ErrorActionPreference='Stop'
if(-not $Workspace){$Workspace=Split-Path $PSScriptRoot -Parent}
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
if([string]::IsNullOrWhiteSpace($Evidence)){throw 'PSP confirmation evidence is required.'}
$s=Open-Workspace $Workspace
try{
 if(-not $s.Records.ContainsKey($RecordId)){throw 'Unknown record.'};$r=$s.Records[$RecordId]
 if($r.SubmissionStatus -in @('NotSubmitted','Reserved') -and $Status -ne 'Voided'){throw 'No recorded submission. Record actual submission evidence first.'}
 if($r.FinalStatus -ne 'Unknown' -and $r.FinalStatus -ne $Status -and -not $Correction){throw 'Changing a final outcome requires -Correction and new evidence.'}
 $at=[datetime]::UtcNow.ToString('o');$old=$r.FinalStatus
 $r.FinalStatus=$Status;$r.FinalConfirmationAt=$at;$r.PSPResponseAt=$at;$r.FailureReason=$FailureReason;$r.Evidence=$Evidence;$r.UpdatedAt=$at
 Save-Transaction $s 'FinalOutcome' @($r) @() @() @([pscustomobject]@{At=$at;Type='FinalOutcome';RecordId=$RecordId;BatchId=$r.ActiveBatchId;Reference=$Evidence;Detail="From=$old; To=$Status; Correction=$Correction; Reason=$FailureReason"})
 Export-Register $s
 Write-Host 'Outcome recorded. Voids are reported separately and never counted as successful refunds.'
}finally{$s.Lock.Dispose()}
