#requires -Version 5.1
[CmdletBinding()]param([string]$Workspace,[Parameter(Mandatory)][string]$BatchId,[Parameter(Mandatory)][string]$Evidence,[switch]$ConfirmNeverSubmitted)
$ErrorActionPreference='Stop'
if(-not $Workspace){$Workspace=Split-Path $PSScriptRoot -Parent}
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
if(-not $ConfirmNeverSubmitted -or [string]::IsNullOrWhiteSpace($Evidence)){throw 'Confirm this batch was never submitted and give a reason.'}
$s=Open-Workspace $Workspace
try{
 if(-not $s.Batches.ContainsKey($BatchId)){throw 'Unknown batch.'};$b=$s.Batches[$BatchId]
 if($b.Status -ne 'Prepared'){throw 'Only unreserved Prepared batches can be cancelled. Reserved/uncertain batches require investigation and remain blocked.'}
 $records=@();foreach($id in $b.RecordIds){$r=$s.Records[$id];$r.ActiveBatchId='';$records+=$r}
 $b.Status='Cancelled';$b.Evidence=$Evidence
 Save-Transaction $s 'CancelPrepared' $records @($b) @() @([pscustomobject]@{At=[datetime]::UtcNow.ToString('o');Type='CancelPrepared';RecordId='';BatchId=$BatchId;Reference=$Evidence;Detail='Never reserved/submitted. Old file must not be used.'})
 Write-AtomicText ($b.FilePath+'.CANCELLED.txt') 'CANCELLED. Never upload or execute the associated file.'
 Export-Register $s
}finally{$s.Lock.Dispose()}
