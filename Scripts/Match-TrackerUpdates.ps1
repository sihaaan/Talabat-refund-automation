#requires -Version 5.1
[CmdletBinding()]param([Parameter(Mandatory)][string]$TrackerCsv,[Parameter(Mandatory)][string]$ManifestCsv,[string]$OutputFolder)
$ErrorActionPreference='Stop'
if(-not $OutputFolder){$OutputFolder=(Join-Path (Split-Path $PSScriptRoot -Parent) 'Reports\TrackerMatching')}
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
$tracker=@(Read-CsvTable $TrackerCsv @('PSP Payment IDs','Order ID','Country Code','Refund Status') 4)
$manifest=@(Read-CsvTable $ManifestCsv @('RecordId','Payment ID','Original Payment IDs','Order ID','Country','New status','Evidence/reference') 1)
$lookup=New-Map
foreach($r in $tracker){
 $raw=Get-Text $r 'PSP Payment IDs';$parts=@($raw.Split(@('||'),[StringSplitOptions]::None)|ForEach-Object{$_.Trim()});$pays=@($parts|Where-Object{$_ -cmatch '^pay_[a-z0-9]{26}$'}|Sort-Object -Unique)
 $id=if($pays.Count -eq 1){$pays[0]}else{$raw}
 $key=$id+'|'+(Get-Text $r 'Order ID')+'|'+(Get-Text $r 'Country Code')
 if(-not $lookup.ContainsKey($key)){$lookup[$key]=[Collections.Generic.List[object]]::new()};$lookup[$key].Add($r)
}
$seen=New-Map;$results=@();$snapshot=Get-FileSha $TrackerCsv
foreach($m in $manifest){
 $key=(Get-Text $m 'Payment ID')+'|'+(Get-Text $m 'Order ID')+'|'+(Get-Text $m 'Country');$matches=if($lookup.ContainsKey($key)){@($lookup[$key])}else{@()}
 $status='';$rowNumber='';$original=''
 if($seen.ContainsKey($key)){$status='DuplicateManifestKey'}elseif(-not(Get-Text $m 'Evidence/reference')){$status='MissingEvidence'}elseif($matches.Count -eq 0){$status='Unmatched'}elseif($matches.Count -gt 1){$status='MultipleMatches'}else{
  $rowNumber=$matches[0].InputRow;$original=Get-Text $matches[0] 'Refund Status'
  if($original -ceq 'Refund Processed'){$status='AlreadyProcessed'}elseif($original -cin @('Pending','Refund Failed')){$status='ReadyForHumanReview'}else{$status='StatusChangedInvestigate'}
 };$seen[$key]=$true
 $results+=[pscustomobject]@{RecordId=$m.RecordId;PaymentId=(Get-Text $m 'Payment ID');OrderId=(Get-Text $m 'Order ID');Country=(Get-Text $m 'Country');Matches=$matches.Count;Result=$status;SnapshotRow=$rowNumber;CurrentStatus=$original;ProposedStatus='Refund Processed';TrackerSnapshotSHA256=$snapshot;Evidence=(Get-Text $m 'Evidence/reference')}
}
$folder=Join-Path $OutputFolder ([guid]::NewGuid().ToString('N').Substring(0,10))
Write-CsvFile (Join-Path $folder 'Reviewed_Tracker_Matches.csv') $results @('RecordId','PaymentId','OrderId','Country','Matches','Result','SnapshotRow','CurrentStatus','ProposedStatus','TrackerSnapshotSHA256','Evidence')
Write-AtomicText (Join-Path $folder 'Read_First.txt') 'This is a local review list. SnapshotRow is not a live Google Sheets address. Sync and export again if rows changed. Require exactly one match per manifest entry; reconcile unmatched/multiple/status-changed rows. With the controller, keep the live tracker stable and manually change only individually verified Refund Status cells. Never paste over all filtered rows or replace the whole status column from a stale export. The tracker supplies timestamps. If stable rows cannot be guaranteed and helper columns are unauthorized, individual verified updates are the safe fallback.'
Write-Host "Matching report: $folder"
[pscustomobject]@{ReportPath=$folder;Entries=$manifest.Count;Ready=@($results|Where-Object Result -eq 'ReadyForHumanReview').Count;AlreadyProcessed=@($results|Where-Object Result -eq 'AlreadyProcessed').Count;NeedsInvestigation=@($results|Where-Object Result -notin @('ReadyForHumanReview','AlreadyProcessed')).Count}
