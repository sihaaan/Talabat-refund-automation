#requires -Version 5.1
[CmdletBinding()]param([string]$Workspace,[Parameter(Mandatory)][string]$RecordId,[Parameter(Mandatory)][ValidateSet('OrderId','Country','Currency','TransactionDate','OriginalAmount','AmountUnit','ChargeId','OperationId','Reason','ApprovalReference','RequestReference')][string]$Field,[Parameter(Mandatory)][string]$Value,[Parameter(Mandatory)][string]$Evidence)
$ErrorActionPreference='Stop'
if(-not $Workspace){$Workspace=Split-Path $PSScriptRoot -Parent}
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
if([string]::IsNullOrWhiteSpace($Evidence)){throw 'Correction evidence is required.'}
$s=Open-Workspace $Workspace
try{
 if(-not $s.Records.ContainsKey($RecordId)){throw 'Unknown record.'};$r=$s.Records[$RecordId]
 if($r.SubmissionStatus -ne 'NotSubmitted' -or $r.ActiveBatchId -or $r.FinalStatus -ne 'Unknown'){throw 'Cannot change payment details of a prepared/attempted/resolved record.'}
 $old=$r.$Field;$r.$Field=$Value
 $dt=Get-DateValue $r.TransactionDate;$r.TransactionDate=$dt.ToString('yyyy-MM-dd')
 if($r.Country -notmatch '^[A-Z]{2}$' -or -not(Get-Setting $s.Config.countryCurrencies $r.Country)){throw 'Country is not configured.'}
 $a=Convert-Amount $r.OriginalAmount $r.AmountUnit $r.Currency $s.Config;$r.AmountMajor=$a.Major;$r.AmountMinor=$a.Minor;$r.MissingFields=Get-MissingFields $r $s.Config;$r.UpdatedAt=[datetime]::UtcNow.ToString('o')
 Save-Transaction $s 'Correction' @($r) @() @() @([pscustomobject]@{At=$r.UpdatedAt;Type='Correction';RecordId=$RecordId;BatchId='';Reference=$Evidence;Detail="$Field changed from '$old' to '$Value'. Update the local input to match before re-preparing."})
 Export-Register $s
}finally{$s.Lock.Dispose()}
