#requires -Version 5.1
[CmdletBinding()]param([string]$TestOutputFolder=(Join-Path $env:TEMP ('TalabatEdgeTests_'+[guid]::NewGuid().ToString('N'))))
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
$app=Split-Path $PSScriptRoot -Parent;$work=Join-Path $TestOutputFolder 'Workspace'
[IO.Directory]::CreateDirectory((Join-Path $work 'Config'))|Out-Null
Copy-Item -LiteralPath (Join-Path $app 'Config\config.json') -Destination (Join-Path $work 'Config\config.json');Copy-Item -LiteralPath (Join-Path $app 'Config\contacts.json') -Destination (Join-Path $work 'Config\contacts.json')
& (Join-Path $PSScriptRoot 'Initialize-Workspace.ps1') -Workspace $work 6>$null
$cfg=Get-Content -LiteralPath (Join-Path $work 'Config\config.json') -Raw|ConvertFrom-Json;$cfg.amountUnitsConfirmed=$true;$cfg.countryCurrencyConfirmed=$true;$cfg.refundAmountPolicyConfirmed=$true;$cfg.amountApprovalReference='FAKE';$cfg.tap.templateConfirmed=$true;$cfg.tap.batchSize=1000
Write-AtomicText (Join-Path $work 'Config\config.json') ($cfg|ConvertTo-Json -Depth 12)
$today=Get-LocalNow $cfg;$recent=$today.AddDays(-1).ToString('yyyy-MM-dd');$checks=[Collections.Generic.List[object]]::new()
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message}}
function Check([string]$Name,[scriptblock]$Action){try{& $Action;$checks.Add([pscustomobject]@{Test=$Name;Result='PASS';Detail=''})}catch{$checks.Add([pscustomobject]@{Test=$Name;Result='FAIL';Detail=$_.Exception.Message});Write-Host "FAIL: $Name - $($_.Exception.Message)"}}
function MustFail([scriptblock]$Action){$failed=$false;try{& $Action|Out-Null}catch{$failed=$true};Assert $failed 'Expected rejection did not occur.'}
function Pay([int]$N){'pay_'+$N.ToString('00000000000000000000000000')}
function Row([int]$N,[string]$ID,[string]$PSP='Credit Card',[string]$Status='Pending',[string]$Date=$recent){[pscustomobject]@{'Transaction Date'=$Date;'Country Code'='AE';'Reconciliation Source'=$PSP;'PSP Payment IDs'=$ID;'Order ID'=('FAKE-'+$N);'PSP Capture Amount LC'='10.50';'Refund Status'=$Status;'Refund Processed At'='';'Charge ID'='';Reason='requested_by_customer'} }
$rows=@((Row 101 (Pay 101)),(Row 102 (Pay 102)),(Row 103 (Pay 103)),(Row 104 'auth_enrich' Tap),(Row 105 (Pay 105) 'Credit Card' 'Refund Failed'),(Row 106 'NA' 'QPAY' Pending ($today.AddDays(-30).ToString('yyyy-MM-dd'))))
$tracker=Join-Path $work 'Input\Edge.csv';Write-CsvFile $tracker $rows @($rows[0].PSObject.Properties.Name) -Machine
$enrich=Join-Path $work 'Input\Enrichment.csv';Write-CsvFile $enrich @([pscustomobject]@{'Payment ID'='auth_enrich';'Order ID'='FAKE-104';Country='AE';'Charge ID'='chg_verified_104';'Operation ID'='';Reason='requested_by_customer'}) @('Payment ID','Order ID','Country','Charge ID','Operation ID','Reason') -Machine
$run=& (Join-Path $PSScriptRoot 'Prepare-Refunds.ps1') -Workspace $work -InputCsv $tracker -EnrichmentCsv $enrich 6>$null
Check 'Enrichment enables future Tap charge IDs with exact matching' {$s=Open-Workspace $work;try{Assert (@($s.Batches.Values|Where-Object PSP -eq 'Tap').Count -eq 1) 'Enrichment did not prepare Tap'}finally{$s.Lock.Dispose()}}
Check 'Failed status without local history or timestamp remains blocked' {$d=Import-Csv -LiteralPath (Join-Path $run.ReportPath 'Duplicates.csv');Assert (@($d|Where-Object OrderId -eq 'FAKE-105').Count -eq 1) 'Failed row retried'}
Check 'Exactly 30 days is delayed even with missing ID' {$d=Import-Csv -LiteralPath (Join-Path $run.ReportPath 'Delayed.csv');Assert (@($d|Where-Object{$_.OrderId -eq 'FAKE-106' -and $_.ReasonForReview -like '*placeholder*'}).Count -eq 1) 'Delayed invalid row disappeared'}
$s=Open-Workspace $work;try{$batch=@($s.Batches.Values|Where-Object PSP -eq 'Checkout')[0];$tap=@($s.Batches.Values|Where-Object PSP -eq 'Tap')[0]}finally{$s.Lock.Dispose()}
Check 'Editing prepared file prevents reservation' {$original=Get-Content -LiteralPath $tap.FilePath -Raw;try{Write-AtomicText $tap.FilePath ($original.Replace('10.5','999'));MustFail {& (Join-Path $PSScriptRoot 'Start-Submission.ps1') -Workspace $work -BatchId $tap.BatchId -ApprovalReference FAKE -ConfirmBusinessAuthorization 6>$null}}finally{Write-AtomicText $tap.FilePath $original}}
& (Join-Path $PSScriptRoot 'Start-Submission.ps1') -Workspace $work -BatchId $batch.BatchId -ApprovalReference FAKE -ConfirmBusinessAuthorization 6>$null
$result=Join-Path $TestOutputFolder 'PartialResult.csv';$resultRows=@([pscustomobject]@{id=(Pay 101);amount='1050';refund_initiated='true';refund_action_id='act_FAKE_101';refund_error=''},[pscustomobject]@{id=(Pay 102);amount='1050';refund_initiated='false';refund_action_id='';refund_error='FAKE request rejected'})
Write-CsvFile $result $resultRows @('id','amount','refund_initiated','refund_action_id','refund_error') -Machine
& (Join-Path $PSScriptRoot 'Import-CheckoutResults.ps1') -Workspace $work -BatchId $batch.BatchId -ResultCsv $result -Evidence 'FAKE ONLY' -ConfirmProductionAttempt 6>$null
Check 'Partial result preserves uncertain missing row and blocks all attempts' {$s=Open-Workspace $work;try{Assert ($s.Attempts.Count -eq 3) 'Missing attempt boundary';Assert (@($s.Attempts.Values|Where-Object SubmissionStatus -eq 'Uncertain').Count -eq 1) 'Missing row not uncertain';Assert (@($s.Attempts.Values|Where-Object SubmissionStatus -eq 'RequestFailed').Count -eq 1) 'False row not request failed'}finally{$s.Lock.Dispose()}}
Check 'CLI false and missing rows do not falsely update tracker' {$rep=& (Join-Path $PSScriptRoot 'Generate-Reports.ps1') -Workspace $work 6>$null;Assert ($rep.TrackerRows -eq 1 -and $rep.VerifiedFailed -eq 0) 'False result treated as final failure or verified attempt'}
Check 'Conflicting result action is rejected atomically' {$bad=Join-Path $TestOutputFolder 'Conflicting.csv';$badRows=@([pscustomobject]@{id=(Pay 101);amount='1050';refund_initiated='true';refund_action_id='act_DIFFERENT';refund_error=''});Write-CsvFile $bad $badRows @('id','amount','refund_initiated','refund_action_id','refund_error') -Machine;MustFail {& (Join-Path $PSScriptRoot 'Import-CheckoutResults.ps1') -Workspace $work -BatchId $batch.BatchId -ResultCsv $bad -Evidence FAKE -ConfirmProductionAttempt 6>$null}}
Check 'Unexpected result IDs rejected' {$bad=Join-Path $TestOutputFolder 'WrongId.csv';Write-CsvFile $bad @([pscustomobject]@{id=(Pay 999);amount='1050';refund_initiated='false';refund_action_id='';refund_error='fake'}) @('id','amount','refund_initiated','refund_action_id','refund_error') -Machine;MustFail {& (Join-Path $PSScriptRoot 'Import-CheckoutResults.ps1') -Workspace $work -BatchId $batch.BatchId -ResultCsv $bad -Evidence FAKE -ConfirmProductionAttempt 6>$null}}
$manual=Join-Path $TestOutputFolder 'Manual.csv';$manualRows=@([pscustomobject]@{'Payment ID'=(Pay 101);'Order ID'='FAKE-101';Country='AE';'PSP/source'='Credit Card';'Transaction Date'=$recent;Amount='10.50';Currency='AED';'Request reference'='FAKE-MAIL-1';Reason='requested_by_customer';'Approval reference'='FAKE-APPROVAL';AmountUnit='major'},[pscustomobject]@{'Payment ID'=(Pay 201);'Order ID'='FAKE-201';Country='AE';'PSP/source'='Credit Card';'Transaction Date'=$recent;Amount='10.50';Currency='AED';'Request reference'='FAKE-MAIL-2';Reason='requested_by_customer';'Approval reference'='';AmountUnit='major'})
Write-CsvFile $manual $manualRows @($manualRows[0].PSObject.Properties.Name) -Machine
$man=& (Join-Path $PSScriptRoot 'Add-ManualRequests.ps1') -Workspace $work -InputCsv $manual 6>$null
Check 'Manual and tracker requests deduplicate together' {Assert (@(Import-Csv -LiteralPath (Join-Path $man.ReportPath 'Duplicates.csv')|Where-Object OrderId -eq 'FAKE-101').Count -eq 1) 'Cross-source duplicate'}
Check 'Manual requests need approval even with configured amount rules' {Assert (@(Import-Csv -LiteralPath (Join-Path $man.ReportPath 'ReviewOnly.csv')|Where-Object OrderId -eq 'FAKE-201').Count -eq 1) 'Manual unapproved prepared'}
Check 'Configured Tap 1000 limit splits 1001 rows into 1000 + 1' {
 $large=Join-Path $TestOutputFolder 'Tap1001.csv';$many=@(1..1001|ForEach-Object{Row (10000+$_) ('chg_FAKE_'+$_) Tap});Write-CsvFile $large $many @($many[0].PSObject.Properties.Name) -Machine
 $largeRun=& (Join-Path $PSScriptRoot 'Prepare-Refunds.ps1') -Workspace $work -InputCsv $large 6>$null
 $bs=@(Import-Csv -LiteralPath (Join-Path $largeRun.ReportPath 'Batches.csv'));Assert ($bs.Count -eq 2) 'Wrong batch split';$sizes=@($bs|ForEach-Object{@(Import-Csv -LiteralPath $_.FilePath).Count}|Sort-Object);Assert ($sizes[0] -eq 1 -and $sizes[1] -eq 1000) 'Wrong per-file counts'
}
Check 'Final outcomes need evidence and explicit correction' {$rid=$batch.RecordIds[0];MustFail {& (Join-Path $PSScriptRoot 'Record-ManualOutcome.ps1') -Workspace $work -RecordId $rid -Status ConfirmedFailed -Evidence ' ' 6>$null}; & (Join-Path $PSScriptRoot 'Record-ManualOutcome.ps1') -Workspace $work -RecordId $rid -Status ConfirmedFailed -Evidence 'FAKE final failure' 6>$null;MustFail {& (Join-Path $PSScriptRoot 'Record-ManualOutcome.ps1') -Workspace $work -RecordId $rid -Status ConfirmedSuccessful -Evidence FAKE 6>$null}}
Check 'Missing committed journal tail cannot silently reset history' {$f=@(Get-ChildItem -LiteralPath (Join-Path $work 'Register\Journal') -Filter '*.json'|Sort-Object Name)[-1];$hidden=$f.FullName+'.held';[IO.File]::Move($f.FullName,$hidden);try{MustFail {$x=Open-Workspace $work;$x.Lock.Dispose()}}finally{[IO.File]::Move($hidden,$f.FullName)}}
Check 'CSV report neutralizes spreadsheet formulas' {$p=Join-Path $TestOutputFolder 'Injection.csv';Write-CsvFile $p @([pscustomobject]@{Text='=1+1'}) @('Text');Assert ((Import-Csv -LiteralPath $p).Text.StartsWith("'=")) 'Formula not neutralized'}
Check 'CSV parser handles embedded comma and newline without dropping records' {$p=Join-Path $TestOutputFolder 'Quoted.csv';Write-CsvFile $p @([pscustomobject]@{ID='a';Note="comma, and`nnewline"}) @('ID','Note') -Machine;$parsed=@(Read-CsvTable $p @('ID','Note') 1);Assert ($parsed.Count -eq 1 -and $parsed[0].Note.Contains("`n")) 'Quoted multiline parse failed'}
Check 'A newly conflicting observation puts an existing batch on hold' {
 $changed=@(Row 104 'auth_enrich' Tap);$changed[0].'PSP Capture Amount LC'='999';$path=Join-Path $TestOutputFolder 'Changed.csv';Write-CsvFile $path $changed @($changed[0].PSObject.Properties.Name) -Machine
 & (Join-Path $PSScriptRoot 'Prepare-Refunds.ps1') -Workspace $work -InputCsv $path 6>$null|Out-Null
 MustFail {& (Join-Path $PSScriptRoot 'Start-Submission.ps1') -Workspace $work -BatchId $tap.BatchId -ApprovalReference FAKE -ConfirmBusinessAuthorization 6>$null}
}
Check 'Different payment ID for a manual/tracker order is held for alias review' {
 $alias=$manualRows[0].PSObject.Copy();$alias.'Payment ID'=Pay 9999;$path=Join-Path $TestOutputFolder 'Alias.csv';Write-CsvFile $path @($alias) @($alias.PSObject.Properties.Name) -Machine
 $ar=& (Join-Path $PSScriptRoot 'Add-ManualRequests.ps1') -Workspace $work -InputCsv $path 6>$null
 Assert ($ar.Prepared -eq 0 -and @(Import-Csv -LiteralPath (Join-Path $ar.ReportPath 'Exceptions.csv')).Count -eq 1) 'Alias could prepare another refund'
}
Check 'Unreserved batch can be cancelled without creating an attempt' {
 $s=Open-Workspace $work;try{$n=$s.Attempts.Count}finally{$s.Lock.Dispose()}
 & (Join-Path $PSScriptRoot 'Cancel-PreparedBatch.ps1') -Workspace $work -BatchId $tap.BatchId -Evidence 'FAKE reconciliation needed' -ConfirmNeverSubmitted 6>$null
 $s=Open-Workspace $work;try{Assert ($s.Attempts.Count -eq $n -and $s.Batches[$tap.BatchId].Status -eq 'Cancelled') 'Cancellation changed attempts'}finally{$s.Lock.Dispose()}
}
Check 'Register backup excludes the writer lock and includes canonical journal' {
 $folder=Join-Path $TestOutputFolder 'Backups'; & (Join-Path $PSScriptRoot 'Backup-Register.ps1') -Workspace $work -DestinationFolder $folder 6>$null
 $zip=@(Get-ChildItem -LiteralPath $folder -Filter '*.zip')[0];$z=[IO.Compression.ZipFile]::OpenRead($zip.FullName)
 try{Assert (@($z.Entries|Where-Object FullName -eq '.lock').Count -eq 0 -and @($z.Entries|Where-Object FullName -like 'Journal/*').Count -gt 0) 'Backup incomplete or lock copied'}finally{$z.Dispose()}
}
Check 'External evidence records an already-occurred delayed submission once' {
 $s=Open-Workspace $work;try{$r=@($s.Records.Values|Where-Object OrderId -eq 'FAKE-106')[0];$before=$s.Attempts.Count}finally{$s.Lock.Dispose()}
 $at=[datetimeoffset]::UtcNow.AddMinutes(-1).ToString('o')
 & (Join-Path $PSScriptRoot 'Record-ExternalAttempt.ps1') -Workspace $work -RecordId $r.RecordId -Method Acquirer -Evidence 'FAKE actual external submission' -ApprovalReference FAKE -AttemptedAt $at -ConfirmActualSubmission 6>$null
 & (Join-Path $PSScriptRoot 'Record-ExternalAttempt.ps1') -Workspace $work -RecordId $r.RecordId -Method Acquirer -Evidence 'FAKE actual external submission' -ApprovalReference FAKE -AttemptedAt $at -ConfirmActualSubmission 6>$null
 $s=Open-Workspace $work;try{Assert ($s.Attempts.Count -eq ($before+1) -and $s.Records[$r.RecordId].FinalStatus -eq 'Unknown') 'External evidence was duplicated or treated successful'}finally{$s.Lock.Dispose()}
}
Write-CsvFile (Join-Path $TestOutputFolder 'Test_Results.csv') $checks @('Test','Result','Detail')
Write-Host "Edge tests: $($checks.Count); passed: $(@($checks|Where-Object Result -eq 'PASS').Count); failed: $(@($checks|Where-Object Result -eq 'FAIL').Count). Files: $TestOutputFolder"
if(@($checks|Where-Object Result -eq 'FAIL').Count){throw 'Edge-case tests failed.'}
