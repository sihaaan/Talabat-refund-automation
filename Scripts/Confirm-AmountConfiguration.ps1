#requires -Version 5.1
[CmdletBinding()]param([string]$Workspace,[Parameter(Mandatory)][ValidateSet('major','minor')][string]$TrackerAmountUnit,[Parameter(Mandatory)][string]$ApprovalReference,[switch]$ConfirmCurrencyMapping,[switch]$ConfirmRefundAmountPolicy)
$ErrorActionPreference='Stop'
if(-not $Workspace){$Workspace=Split-Path $PSScriptRoot -Parent}
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
if(-not $ConfirmCurrencyMapping -or -not $ConfirmRefundAmountPolicy -or [string]::IsNullOrWhiteSpace($ApprovalReference)){throw 'Obtain and explicitly confirm the currency mapping, source amount units and approved refund amount policy first.'}
$s=Open-Workspace $Workspace
try{
 $s.Config.trackerAmountUnit=$TrackerAmountUnit;$s.Config.amountUnitsConfirmed=$true;$s.Config.countryCurrencyConfirmed=$true;$s.Config.refundAmountPolicyConfirmed=$true;$s.Config.amountApprovalReference=$ApprovalReference
 Write-AtomicText (Join-Path $s.Root 'Config\config.json') ($s.Config|ConvertTo-Json -Depth 12)
 Save-Transaction $s 'ConfigurationApproval' @() @() @() @([pscustomobject]@{At=[datetime]::UtcNow.ToString('o');Type='ConfigurationApproval';RecordId='';BatchId='';Reference=$ApprovalReference;Detail="Source=$TrackerAmountUnit; country currency and refund policy confirmed"})
 Write-Host 'Assumptions recorded. Existing prepared batches must keep their original configuration hash or be cancelled and prepared again.'
}finally{$s.Lock.Dispose()}
