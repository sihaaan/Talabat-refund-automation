#requires -Version 5.1
[CmdletBinding()]param([string]$Workspace,[Parameter(Mandatory)][string[]]$RecordId,[Parameter(Mandatory)][ValidateSet('Escalation','EmailSent','PSPAcknowledged')][string]$Kind,[Parameter(Mandatory)][string]$Evidence)
$ErrorActionPreference='Stop'
if(-not $Workspace){$Workspace=Split-Path $PSScriptRoot -Parent}
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
if([string]::IsNullOrWhiteSpace($Evidence)){throw 'Communication reference is required.'}
$s=Open-Workspace $Workspace
try{
 $records=@();$events=@();$at=[datetime]::UtcNow.ToString('o')
 foreach($id in $RecordId){if(-not $s.Records.ContainsKey($id)){throw "Unknown record $id"};$r=$s.Records[$id];if($Kind -eq 'PSPAcknowledged'){$r.PSPResponseAt=$at}else{$r.EmailSentAt=$at};$r.UpdatedAt=$at;$records+=$r;$events+=[pscustomobject]@{At=$at;Type=$Kind;RecordId=$id;BatchId='';Reference=$Evidence;Detail='Communication only; no new refund attempt or final success inferred.'}}
 Save-Transaction $s 'Communication' $records @() @() $events
 Export-Register $s
 Write-Host 'Communication recorded. It does not add a refund attempt or tracker status update.'
}finally{$s.Lock.Dispose()}
