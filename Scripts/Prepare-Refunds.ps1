#requires -Version 5.1
[CmdletBinding()]param(
 [string]$Workspace,[string]$InputCsv,
 [ValidateSet('tracker','manual')][string]$Source='tracker',
 [string]$AsOfDate,[switch]$ReviewOnly,[string]$EnrichmentCsv
)
$ErrorActionPreference='Stop'
if(-not $Workspace){$Workspace=Split-Path $PSScriptRoot -Parent}
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
$s=Open-Workspace $Workspace
try{
 $now=Get-LocalNow $s.Config;$asof=if($AsOfDate){Get-DateValue $AsOfDate}else{$now.Date}
 if(-not $InputCsv){$files=@(Get-ChildItem -LiteralPath (Join-Path $s.Root 'Input') -Filter '*.csv'|Where-Object Name -notlike 'Manual*');if($files.Count -ne 1){throw 'Put exactly one tracker CSV in Input, or supply -InputCsv explicitly.'};$InputCsv=$files[0].FullName}
 $InputCsv=[IO.Path]::GetFullPath($InputCsv);$hash=Get-FileSha $InputCsv
 $run='RUN_'+$now.ToString('yyyyMMdd_HHmmss')+'_'+[guid]::NewGuid().ToString('N').Substring(0,8)
 $required=if($Source -eq 'tracker'){@('Transaction Date','Country Code','Reconciliation Source','PSP Payment IDs','Order ID','PSP Capture Amount LC','Refund Status')}else{@('Payment ID','Order ID','Country','PSP/source','Transaction Date','Amount','Currency','Request reference','Reason','Approval reference','AmountUnit')}
 $rows=@(Read-CsvTable $InputCsv $required $(if($Source -eq 'tracker'){[int]$s.Config.trackerHeaderRow}else{1}))
 $normalized=[Collections.Generic.List[object]]::new();$counts=New-Map;$enrich=New-Map;$usedEnrich=New-Map
 if($EnrichmentCsv){
  foreach($e in @(Read-CsvTable $EnrichmentCsv @('Payment ID','Order ID','Country') 1)){
   $key=(Get-Text $e 'Payment ID')+'|'+(Get-Text $e 'Order ID')+'|'+(Get-Text $e 'Country')
   if($enrich.ContainsKey($key)){throw 'Duplicate enrichment keys. Resolve them before preparation.'};$enrich[$key]=$e
  }
 }
 $validated=0
 foreach($row in $rows){
  $n=Convert-InputRow $row $Source $s.Config $hash $run $asof
  $r=$n.Record;$key=$r.PaymentId+'|'+$r.OrderId+'|'+$r.Country
  if($enrich.ContainsKey($key)){
   $e=$enrich[$key];$usedEnrich[$key]=$true
   foreach($mapping in @(@('ChargeId','Charge ID'),@('OperationId','Operation ID'),@('Reason','Reason'))){$v=Get-Text $e $mapping[1];if($v){$r.($mapping[0])=$v}}
   $r.MissingFields=Get-MissingFields $r $s.Config
  }
  $normalized.Add($n)
  $validated++;if($validated % 10000 -eq 0){Write-Host "Validated $validated of $($rows.Count) rows..."}
  if(-not $counts.ContainsKey($r.RecordId)){$counts[$r.RecordId]=0};$counts[$r.RecordId]++
 }
 $changed=New-Map;$audit=[Collections.Generic.List[object]]::new();$decisions=[Collections.Generic.List[object]]::new();$ready=[Collections.Generic.List[object]]::new()
 $otherSources=New-Map
 foreach($known in $s.Records.Values){$key=$known.OrderId+'|'+$known.Country;if(-not $otherSources.ContainsKey($key)){$otherSources[$key]=[Collections.Generic.List[object]]::new()};$otherSources[$key].Add($known)}
 $classified=0
 foreach($n in $normalized){
  $new=$n.Record;$r=$new;$conflict=''
  if($s.Records.ContainsKey($new.RecordId) -or $changed.ContainsKey($new.RecordId)){
   $r=if($changed.ContainsKey($new.RecordId)){$changed[$new.RecordId]}else{$s.Records[$new.RecordId]}
   foreach($field in @('OrderId','Country','Currency','AmountMinor','TransactionDate')){if((Get-Text $r $field) -cne (Get-Text $new $field)){$conflict+='Changed '+$field+'; '}}
   if($r.Source -notmatch $Source){$r.Source='tracker+manual'}
   # Preserve all attempt/final evidence, while recording a new observation.
   if($Source -eq 'tracker'){$r.TrackerStatus=$new.TrackerStatus;if($new.ProcessedAt){$r.ProcessedAt=$new.ProcessedAt}}
   if($new.SubmissionStatus -eq 'LegacyAttempt' -and $r.SubmissionStatus -eq 'NotSubmitted'){$r.SubmissionStatus='LegacyAttempt'}
   if(-not $conflict -and -not $r.ActiveBatchId -and $r.SubmissionStatus -eq 'NotSubmitted'){
    foreach($field in @('ChargeId','OperationId','Reason','ApprovalReference','RequestReference','MissingFields')){if(Get-Text $new $field){$r.$field=$new.$field}}
    $r.MissingFields=Get-MissingFields $r $s.Config
   }
   $r.UpdatedAt=[datetime]::UtcNow.ToString('o')
  }
  if($null -eq $r.PSObject.Properties['HoldReason']){$r|Add-Member -NotePropertyName HoldReason -NotePropertyValue ''}
  $orderKey=$new.OrderId+'|'+$new.Country
  if($otherSources.ContainsKey($orderKey)){
   foreach($other in $otherSources[$orderKey]){if($other.RecordId -cne $new.RecordId -and (($Source -eq 'manual' -and $other.Source -match 'tracker') -or ($Source -eq 'tracker' -and $other.Source -match 'manual'))){$conflict+='Same order/country exists through another payment ID or PSP; investigate aliases/partial payments; ';break}}
  }
  if($n.Error -or $conflict -or $counts[$new.RecordId] -gt 1){$r.HoldReason=($n.Error+'; '+$conflict+$(if($counts[$new.RecordId] -gt 1){'; Duplicate input occurrences'}))}
  elseif(-not $r.ActiveBatchId -and $r.SubmissionStatus -eq 'NotSubmitted'){$r.HoldReason=''}
  $changed[$r.RecordId]=$r
  $category='';$why='';$eligible=$new.TrackerStatus -cin @('Pending','Refund Failed')
  $pspcfg=Get-Setting $s.Config.psps $new.PSP;$days=[int]$s.Config.delayedDays;if($pspcfg -and (Get-Setting $pspcfg 'delayedDays')){$days=[int]$pspcfg.delayedDays}
  $blocked=$r.SubmissionStatus -ne 'NotSubmitted' -or $r.FinalStatus -ne 'Unknown'
  if(-not $eligible){$category='Ineligible';$why='Tracker status excluded; historical evidence retained.'}
  elseif($n.Age -ne '' -and [int]$n.Age -ge $days){
   $category='Delayed';$why="Age is at least $days days. Investigation/escalation only."
   if($n.Error){$why+=' Validation: '+$n.Error}
   if($conflict){$why+=' Conflicting stored data: '+$conflict}
   if($counts[$new.RecordId] -gt 1){$why+=' Duplicate input occurrences; do not request multiple refunds.'}
   if($blocked){$why+=' Prior attempt/status exists; do not request a duplicate refund.'}
   if($r.PSP -eq 'QPay' -and [int]$n.Age -gt 90){$why+=' TESS limit exceeded: controller must identify acquiring bank.'}
   if($r.PSP -eq 'Knet'){$why+=' Failed or older-than-six-month refunds require authorized Kuwait Finance handling.'}
  }
  elseif($n.Error){$category='Exceptions';$why=$n.Error}
  elseif($counts[$new.RecordId] -gt 1){$category='Duplicates';$why='Payment appears multiple times in this input. All occurrences blocked; compare amounts/order IDs.'}
  elseif($conflict){$category='Exceptions';$why='Conflicting stored transaction: '+$conflict+'Use Record-Correction only for an unsubmitted, unprepared record.'}
  elseif($r.HoldReason){$category='ReviewOnly';$why='Record is on hold: '+$r.HoldReason+' Cancel any unsubmitted prepared batch and reconcile the source.'}
  elseif($blocked){$category='Duplicates';$why='Previously attempted, legacy processed, uncertain or resolved. Investigate; never automatically retry.'}
  elseif($r.ActiveBatchId){$category='Duplicates';$why='Already belongs to active batch '+$r.ActiveBatchId+'. Use that batch; do not prepare another.'}
  elseif($Source -eq 'manual' -and -not $r.ApprovalReference){$category='ReviewOnly';$why='Manual request needs business approval reference.'}
  elseif($ReviewOnly){$category='ReviewOnly';$why='Review-only run; no submission files authorized.'}
  elseif(-not $s.Config.amountUnitsConfirmed -or -not $s.Config.countryCurrencyConfirmed -or -not $s.Config.refundAmountPolicyConfirmed -or -not $s.Config.amountApprovalReference){$category='ReviewOnly';$why='Confirm source units, currency mapping and refund amount policy with an approval reference.'}
  elseif($r.MissingFields){$category='ReviewOnly';$why=$r.MissingFields}
  elseif($null -eq $pspcfg -or ($r.PSP -ne 'Tap' -and -not $pspcfg.templateConfirmed)){$category='ReviewOnly';$why='Provider format not approved.'}
  else{$category='Prepared';$why='Prepared only. Human approval and reservation required before manual submission.';$ready.Add($r)}
  $d=[ordered]@{};foreach($col in $ReviewColumns){$d[$col]=Get-Text $new $col};$d.RecordId=$r.RecordId;$d.AgeDays=$n.Age;$d.Category=$category;$d.ReasonForReview=$why;$d.SubmissionStatus=$r.SubmissionStatus;$d.FinalStatus=$r.FinalStatus;$d.MissingFields=$r.MissingFields
  $decisions.Add([pscustomobject]$d)
  $classified++;if($classified % 10000 -eq 0){Write-Host "Classified $classified of $($rows.Count) rows..."}
 }
 $batches=[Collections.Generic.List[object]]::new();$report=Join-Path $s.Root ('Reports\'+$run)
 $pending=Join-Path $s.Root 'Register\pending-preparation.json'
 Write-AtomicText $pending (@{RunId=$run;SequenceBefore=$s.Sequence;SourceHash=$hash;At=[datetime]::UtcNow.ToString('o')}|ConvertTo-Json)
 $sourceDir=Join-Path $s.Root 'Register\Sources';[IO.Directory]::CreateDirectory($sourceDir)|Out-Null
 $sourceCopy=Join-Path $sourceDir ($hash+'.csv');if(-not(Test-Path -LiteralPath $sourceCopy)){[IO.File]::Copy($InputCsv,$sourceCopy)}
 [IO.Directory]::CreateDirectory($report)|Out-Null
 foreach($group in @($ready|Group-Object PSP,Country,Currency)){
  $rr=@($group.Group);$first=$rr[0];$folder=Get-RunFolder $s $first $now $run
  $limit=if($first.PSP -eq 'Tap'){[int]$s.Config.tap.batchSize}else{10000};$index=0
  for($offset=0;$offset -lt $rr.Count;$offset+=$limit){
   $index++;$last=[Math]::Min($offset+$limit-1,$rr.Count-1);$part=@($rr[$offset..$last]);$bid='B-'+[guid]::NewGuid().ToString('N')
   $name=if($first.PSP -eq 'Checkout'){"refunds_$($first.Country)_$index.csv"}elseif($first.PSP -eq 'Tap'){"tap_$($first.Country)_$index.csv"}else{"$($first.PSP)_$($first.Country)_$index.csv"}
   if($first.PSP -eq 'Tap' -and $name.Length -gt [int]$s.Config.tap.maxFilenameLength){throw 'Tap filename exceeds configured length.'}
   $provider=@(Get-ProviderRows $first.PSP $part);$cols=@($provider[0].PSObject.Properties.Name)
   $path=Join-Path $folder $name;Write-CsvFile $path $provider $cols -Machine
   $batch=[pscustomobject]@{BatchId=$bid;RunId=$run;PSP=$first.PSP;Country=$first.Country;Currency=$first.Currency;Method=(Get-Setting $s.Config.psps $first.PSP).method;FilePath=$path;FileHash=(Get-FileSha $path);RecordIds=@($part|ForEach-Object RecordId);Status='Prepared';PreparedAt=[datetime]::UtcNow.ToString('o');ReservedAt='';ApprovalReference='';Evidence='';ConfigHash=(Get-FileSha (Join-Path $s.Root 'Config\config.json'))}
   $batches.Add($batch)
   foreach($r in $part){$r.ActiveBatchId=$bid;$r.PreparedAt=$batch.PreparedAt;$r.RunId=$run}
   Write-CsvFile (Join-Path $folder ($name+'.manifest.csv')) $part $RegisterColumns
   Write-AtomicText (Join-Path $folder ($name+'.batch.txt')) "Batch ID: $bid`r`nRows: $($part.Count)`r`nCurrency: $($first.Currency)`r`nFile SHA256: $($batch.FileHash)`r`nPREPARED ONLY. Reserve this batch before any manual submission. Never rerun a Checkout batch."
   if($batch.Method -eq 'Email'){Write-EmailDraft $s $folder $first.PSP $first.Country}
  }
 }
 # Every delayed/review/invalid/blocked candidate remains available in country-specific investigation files.
 Write-Host 'Writing PSP/country investigation files...'
 foreach($g in @($decisions|Where-Object Category -in @('Delayed','ReviewOnly','Exceptions','Duplicates')|Group-Object PSP,Country,Category)){
  $rr=@($g.Group);$first=$rr[0];$folder=Get-RunFolder $s $first $now $run -Delayed
  if($first.Category -ne 'Delayed'){$folder=Join-Path (Split-Path $folder -Parent) 'Review_Only'}
  $file=if($first.Category -eq 'Delayed'){'Delayed_30plus_days.csv'}else{$first.Category+'.csv'}
  Write-CsvFile (Join-Path $folder $file) $rr $ReviewColumns
  Write-EmailDraft $s $folder $first.PSP $first.Country -Review
 }
 Write-Host 'Writing full row reconciliation and summary reports...'
 Write-CsvFile (Join-Path $report 'Row_Dispositions.csv') $decisions $ReviewColumns
 foreach($cat in @('Delayed','ReviewOnly','Exceptions','Duplicates','Ineligible')){Write-CsvFile (Join-Path $report ($cat+'.csv')) @($decisions|Where-Object Category -eq $cat) $ReviewColumns}
 $summary=[Collections.Generic.List[object]]::new()
 foreach($g in @($decisions|Group-Object Category,PSP,Country)){$first=$g.Group[0];$summary.Add([pscustomobject]@{Category=$first.Category;PSP=$first.PSP;Country=$first.Country;Rows=$g.Count})}
 Write-CsvFile (Join-Path $report 'Preparation_Summary.csv') $summary @('Category','PSP','Country','Rows')
 Write-CsvFile (Join-Path $report 'Batches.csv') $batches @('BatchId','PSP','Country','Currency','Method','FilePath','FileHash','Status')
 $unused=@($enrich.Keys|Where-Object{-not $usedEnrich.ContainsKey($_)}|ForEach-Object{[pscustomobject]@{Key=$_;Reason='No exact payment/order/country input match'}})
 Write-CsvFile (Join-Path $report 'Unmatched_Enrichment.csv') $unused @('Key','Reason')
 $eligibleCount=@($decisions|Where-Object Category -ne 'Ineligible').Count
 $countsText=@($decisions|Group-Object Category|ForEach-Object{"$($_.Name): $($_.Count)"}) -join "`r`n"
 $text="Run: $run`r`nAs-of date: $($asof.ToString('yyyy-MM-dd'))`r`nRows read: $($rows.Count)`r`nEligible-status rows: $eligibleCount`r`n$countsText`r`n`r`n1. Review Row_Dispositions and the currency/amount assumptions.`r`n2. Delayed_30plus_days and Review_Only files are investigation attachments, not upload files. Review prior-attempt warnings before emailing.`r`n3. Batches.csv lists prepared submission files. Record approval and reserve a batch with Start-Submission before manually using it.`r`n4. Import Checkout results or record actual Tap/email submission evidence. Missing results mean Uncertain, never rerun.`r`n5. Generate reports. Only recorded real attempts enter Tracker_Update_Manifest.`r`n6. Sync the controller tracker through its existing manual process, match a fresh export, then change only approved Refund Status cells. The tracker stamps its own timestamp.`r`n7. Record final PSP outcomes separately. A failed signal or email is not final success.`r`n`r`nNo refund, email, Drive upload or tracker update was performed. CSVs preserve text IDs on disk; import into Excel with ID columns as Text.`r`n"
 Write-AtomicText (Join-Path $report 'Run_Sheet.txt') $text
 $audit.Add([pscustomobject]@{At=[datetime]::UtcNow.ToString('o');Type='Preparation';RecordId='';BatchId='';Reference=$hash;Detail="Run=$run; Source=$Source; Read=$($rows.Count); Eligible=$eligibleCount; ReviewOnly=$ReviewOnly"})
 Write-Host 'Committing the register and refreshing its CSV views...'
 Save-Transaction $s 'Prepare' @($changed.Values) @($batches) @() @($audit)
 Export-Register $s
 Remove-Item -LiteralPath $pending
 Write-Host $text
 Write-Host "Reports: $report"
 [pscustomobject]@{RunId=$run;ReportPath=$report;Read=$rows.Count;Prepared=$ready.Count;Batches=@($batches|ForEach-Object BatchId)}
}finally{$s.Lock.Dispose()}
