#requires -Version 5.1
[CmdletBinding()]param([string]$TestOutputFolder=(Join-Path $env:TEMP ('TalabatRefundTests_'+[guid]::NewGuid().ToString('N'))),[switch]$KeepExampleOutputs)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
$app=Split-Path $PSScriptRoot -Parent
$work=Join-Path $TestOutputFolder 'Workspace';[IO.Directory]::CreateDirectory((Join-Path $work 'Config'))|Out-Null
Copy-Item -LiteralPath (Join-Path $app 'Config\config.json') -Destination (Join-Path $work 'Config\config.json')
Copy-Item -LiteralPath (Join-Path $app 'Config\contacts.json') -Destination (Join-Path $work 'Config\contacts.json')
& (Join-Path $PSScriptRoot 'Initialize-Workspace.ps1') -Workspace $work 6>$null
$checks=[Collections.Generic.List[object]]::new()
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message}}
function Check([string]$Name,[scriptblock]$Action){try{& $Action;$checks.Add([pscustomobject]@{Test=$Name;Result='PASS';Detail=''})}catch{$checks.Add([pscustomobject]@{Test=$Name;Result='FAIL';Detail=$_.Exception.Message});Write-Host "FAIL: $Name - $($_.Exception.Message)"}}
function MustFail([scriptblock]$Action){$failed=$false;try{& $Action|Out-Null}catch{$failed=$true};Assert $failed 'Expected rejection did not occur.'}
function Pay([int]$N){return 'pay_'+$N.ToString('00000000000000000000000000')}
$cfg=Get-Content -LiteralPath (Join-Path $work 'Config\config.json') -Raw|ConvertFrom-Json
$today=Get-LocalNow $cfg;$recent=$today.AddDays(-1).ToString('yyyy-MM-dd');$old=$today.AddDays(-40).ToString('yyyy-MM-dd')
function Row([int]$N,[string]$Source='Credit Card',[string]$Country='AE',[string]$Amount='25.50',[string]$Status='Pending',[string]$ID='',[string]$Date=$recent){if(-not $ID){$ID=Pay $N};[pscustomobject]@{'Transaction Date'=$Date;'Transaction Month'=$Date.Substring(0,7);'Country Code'=$Country;'Reconciliation Source'=$Source;'PSP Payment IDs'=$ID;'Order ID'=('FAKE-ORDER-'+$N);'PSP Capture Amount LC'=$Amount;'Refund Status'=$Status;'Refund Processed At'=$(if($Status -ne 'Pending'){$recent}else{''});Notes='SYNTHETIC TEST ONLY';'Currency'='';'AmountUnit'='major';'Charge ID'='';'Operation ID'='';Reason=''} }
$rows=@(
 (Row 1), (Row 2 'CKO-Knet' 'KW' '1.555'), (Row 3 'Credit Card & talabat Pay' 'QA' '2.50' 'Pending' ((Pay 3)+'||auth_fake')),
 (Row 4 'Tap' 'EG' '10.00' 'Pending' ((Pay 4)+'||tr_fake')), (Row 5 'CKO-QPay & talabat Pay' 'QA'),
 (Row 6 'Credit Card' 'JO' '2.125'), (Row 7 'Credit Card' 'BH' '1.555'), (Row 8 'Credit Card' 'OM' '2.125'),
 (Row 9 'Credit Card' 'IQ' '1250'), (Row 10 'Credit Card' 'AE' '5' 'Pending' 'pay_short'),
 (Row 11 'Credit Card' 'AE' ''), (Row 12 'UnknownPSP' 'AE' '5' 'Pending' 'unknown123'),
 (Row 13), (Row 13 'Credit Card' 'AE' '40'), (Row 14 'Credit Card' 'AE' '5' 'Refund Failed'),
 (Row 15 'Credit Card' 'AE' '5' 'Pending' '' $old), (Row 16 'Tap' 'KW' '5.450' 'Pending' 'auth_fake16'),
 (Row 17 'Tap' 'AE' '12.50' 'Pending' 'chg_fake17'), (Row 18 'HyperPay' 'IQ' '9000' 'Pending' 'fakehyper18'),
 (Row 19 'BenefitPay' 'BH' '10.555' 'Pending' 'fakebenefit19'), (Row 20 'Credit Card' 'AE' '5' 'Refunded'),
 (Row 21 'Credit Card' 'AE' '5' 'Pending' ((Pay 21)+'||'+(Pay 22))), (Row 23 'QPAY' 'QA' '10' 'Pending' 'NA'),
 (Row 24 'Zaincash' 'IQ' '12000' 'Pending' 'fakezain24'), (Row 25 'KNET' 'KW' '1.25' 'Pending' '000000000001234'),
 (Row 26 'Credit Card' 'AE' '1.555'), (Row 27 'Credit Card' 'AE' '0'), (Row 28 'Credit Card' 'AE' '3' 'Pending' '' ($today.AddDays(1).ToString('yyyy-MM-dd'))),
 (Row 29 'Credit Card' 'AE' '5' 'Refund Failed' '' $old), (Row 30 'Credit Card' 'AE' '5' 'Refunded'), (Row 30 'Credit Card' 'AE' '5')
)
$rows[17].Reason='requested_by_customer'
$input=Join-Path $work 'Input\Tracker.csv';$cols=@($rows[0].PSObject.Properties.Name);Write-CsvFile $input $rows $cols -Machine
$csv=Get-Content -LiteralPath $input -Raw;Write-AtomicText $input ("Synthetic tracker`r`nRead-only heading`r`nSummary`r`n"+$csv)
Check 'Major AED -> integer minor units' {Assert ((Convert-Amount '25.50' major AED $cfg).Minor -eq '2550') 'Wrong multiplier'}
Check 'Minor units are not multiplied twice' {Assert ((Convert-Amount '2550' minor AED $cfg).Minor -eq '2550') 'Double conversion'}
Check 'Exact three decimal KWD' {Assert ((Convert-Amount '1.555' major KWD $cfg).Minor -eq '1555') 'Three decimals lost'}
Check 'Fractional minor units rejected' {MustFail {Convert-Amount '1.555' major AED $cfg}}
Check 'Locale/exponent amounts rejected' {MustFail {Convert-Amount '1,234.50' major AED $cfg};MustFail {Convert-Amount '1e3' major AED $cfg}}
Check 'ISO week boundary uses ISO year' {Assert ((Get-IsoWeek ([datetime]'2021-01-01')) -eq '2020-W53') 'Wrong ISO year';Assert ((Get-IsoWeek ([datetime]'2018-12-31')) -eq '2019-W01') 'Wrong ISO first week'}
$first=& (Join-Path $PSScriptRoot 'Prepare-Refunds.ps1') -Workspace $work -InputCsv $input 6>$null
Check 'Unconfirmed assumptions produce no executable batch' {Assert ($first.Prepared -eq 0) 'Unconfirmed preparation'}
Check 'Every CSV row is classified' {$d=@(Import-Csv -LiteralPath (Join-Path $first.ReportPath 'Row_Dispositions.csv'));Assert ($d.Count -eq $rows.Count) 'Row dropped'}
Check 'Failed tracker rows seed historical blocking' {$s=Open-Workspace $work;try{Assert (@($s.Records.Values|Where-Object{$_.PaymentId -eq (Pay 14) -and $_.SubmissionStatus -eq 'LegacyAttempt'}).Count -eq 1) 'Legacy lost'}finally{$s.Lock.Dispose()}}
Check 'Duplicate historical and pending rows cannot erase prior evidence' {$s=Open-Workspace $work;try{Assert (@($s.Records.Values|Where-Object{$_.PaymentId -eq (Pay 30) -and $_.SubmissionStatus -eq 'LegacyAttempt'}).Count -eq 1) 'Historical evidence overwritten'}finally{$s.Lock.Dispose()}}
Check 'Delayed files include prior-attempt warnings and country folders' {$d=@(Import-Csv -LiteralPath (Join-Path $first.ReportPath 'Delayed.csv'));Assert ($d.Count -eq 2) 'Expected two delayed rows';Assert (@($d|Where-Object ReasonForReview -like '*Prior attempt*').Count -eq 1) 'Prior attempt warning missing';Assert (@(Get-ChildItem -LiteralPath (Join-Path $work 'Output\Checkout\AE') -Recurse -Filter 'Delayed_30plus_days.csv').Count -eq 1) 'Missing country file'}
& (Join-Path $PSScriptRoot 'Confirm-AmountConfiguration.ps1') -Workspace $work -TrackerAmountUnit major -ApprovalReference 'FAKE-APPROVAL-ONLY' -ConfirmCurrencyMapping -ConfirmRefundAmountPolicy 6>$null
$cfg=Get-Content -LiteralPath (Join-Path $work 'Config\config.json') -Raw|ConvertFrom-Json;$cfg.tap.templateConfirmed=$true;$cfg.psps.Switch.templateConfirmed=$true
Write-AtomicText (Join-Path $work 'Config\config.json') ($cfg|ConvertTo-Json -Depth 12)
$second=& (Join-Path $PSScriptRoot 'Prepare-Refunds.ps1') -Workspace $work -InputCsv $input 6>$null
$s=Open-Workspace $work;try{$allBatches=@($s.Batches.Values);$ae=@($allBatches|Where-Object{$_.PSP -eq 'Checkout' -and $_.Country -eq 'AE'})[0];$tap=@($allBatches|Where-Object PSP -eq 'Tap')[0];$benefit=@($allBatches|Where-Object PSP -eq 'BenefitPay')[0];$kw=@($allBatches|Where-Object{$_.PSP -eq 'Checkout' -and $_.Country -eq 'KW'})[0]}finally{$s.Lock.Dispose()}
Check 'Checkout overrides Tap label and strips companion ID' {$eg=@($allBatches|Where-Object{$_.PSP -eq 'Checkout' -and $_.Country -eq 'EG'});Assert ($eg.Count -eq 1) 'Incorrect routing';$r=Import-Csv -LiteralPath $eg[0].FilePath;Assert ($r.id -eq (Pay 4)) 'Companion ID retained'}
Check 'Checkout file has exact amount and only id amount columns' {$r=Import-Csv -LiteralPath $kw.FilePath;Assert ($r.amount -eq '1555') 'Incorrect minor value';Assert ((@($r.PSObject.Properties.Name)-join ',') -eq 'id,amount') 'Wrong headers'}
Check 'Tap uses major units and verified charge ID' {$r=Import-Csv -LiteralPath $tap.FilePath;Assert ($r.amount -eq '12.5' -and $r.charge_id -eq 'chg_fake17') 'Incorrect Tap data'}
Check 'Missing Tap charge ID excluded from upload' {$d=@(Import-Csv -LiteralPath (Join-Path $second.ReportPath 'ReviewOnly.csv'));Assert (@($d|Where-Object OrderId -eq 'FAKE-ORDER-16').Count -eq 1) 'Missing charge allowed'}
Check 'HyperPay uses Switch route' {Assert (@($allBatches|Where-Object PSP -eq 'Switch').Count -eq 1) 'Wrong shared route'}
Check 'BenefitPay preserves three decimals' {$r=Import-Csv -LiteralPath $benefit.FilePath;Assert ($r.Value -eq '10.555') 'Benefit precision lost'}
$third=& (Join-Path $PSScriptRoot 'Prepare-Refunds.ps1') -Workspace $work -InputCsv $input 6>$null
Check 'Repeated preparation cannot create a second active batch' {Assert ($third.Prepared -eq 0) 'Duplicate preparation'}
$before=& (Join-Path $PSScriptRoot 'Generate-Reports.ps1') -Workspace $work 6>$null
Check 'Prepared-only does not enter tracker manifest' {Assert ($before.TrackerRows -eq 0) 'Prepared treated attempted'}
Check 'Business approval required before reservation' {MustFail {& (Join-Path $PSScriptRoot 'Start-Submission.ps1') -Workspace $work -BatchId $ae.BatchId -ApprovalReference FAKE 6>$null}}
& (Join-Path $PSScriptRoot 'Start-Submission.ps1') -Workspace $work -BatchId $ae.BatchId -ApprovalReference FAKE -ConfirmBusinessAuthorization 6>$null
Check 'Reservation itself is not a tracker update' {$x=& (Join-Path $PSScriptRoot 'Generate-Reports.ps1') -Workspace $work 6>$null;Assert ($x.TrackerRows -eq 0) 'Reservation treated attempted'}
Check 'Reserved batch cannot start twice' {MustFail {& (Join-Path $PSScriptRoot 'Start-Submission.ps1') -Workspace $work -BatchId $ae.BatchId -ApprovalReference FAKE -ConfirmBusinessAuthorization 6>$null}}
$result=Join-Path $TestOutputFolder 'fake_checkout_result.csv';$resultRows=@(Import-Csv -LiteralPath $ae.FilePath|ForEach-Object{[pscustomobject]@{id=$_.id;amount=$_.amount;refund_initiated='true';refund_action_id='act_FAKE_TEST_ONLY';refund_error=''}});Write-CsvFile $result $resultRows @('id','amount','refund_initiated','refund_action_id','refund_error') -Machine
Check 'Result import requires explicit actual-attempt attestation' {MustFail {& (Join-Path $PSScriptRoot 'Import-CheckoutResults.ps1') -Workspace $work -BatchId $ae.BatchId -ResultCsv $result -Evidence FAKE 6>$null}}
& (Join-Path $PSScriptRoot 'Import-CheckoutResults.ps1') -Workspace $work -BatchId $ae.BatchId -ResultCsv $result -Evidence 'SYNTHETIC TEST; no actual execution' -ConfirmProductionAttempt 6>$null
& (Join-Path $PSScriptRoot 'Import-CheckoutResults.ps1') -Workspace $work -BatchId $ae.BatchId -ResultCsv $result -Evidence FAKE -ConfirmProductionAttempt 6>$null
Check 'Identical result import is idempotent' {$s=Open-Workspace $work;try{Assert ($s.Attempts.Count -eq 1) 'Duplicate attempt'}finally{$s.Lock.Dispose()}}
Check 'Accepted request is not final success' {$x=& (Join-Path $PSScriptRoot 'Generate-Reports.ps1') -Workspace $work 6>$null;Assert ($x.VerifiedSuccessful -eq 0 -and $x.TrackerRows -eq 1) 'Submission/final confused'}
& (Join-Path $PSScriptRoot 'Record-ManualOutcome.ps1') -Workspace $work -RecordId $ae.RecordIds[0] -Status ConfirmedSuccessful -Evidence 'FAKE PSP FINAL CONFIRMATION' 6>$null
Check 'Confirmed PSP outcome changes verified rate' {$x=& (Join-Path $PSScriptRoot 'Generate-Reports.ps1') -Workspace $work 6>$null;Assert ($x.VerifiedSuccessful -eq 1) 'Final outcome absent';$m=Import-Csv -LiteralPath (Join-Path $x.ReportPath 'Refund_Success_Rate_Report.csv');Assert (($m|Where-Object Metric -eq 'Verified success rate').Value -eq '1.0000') 'Wrong verified rate'}
& (Join-Path $PSScriptRoot 'Start-Submission.ps1') -Workspace $work -BatchId $kw.BatchId -ApprovalReference FAKE -ConfirmBusinessAuthorization 6>$null
& (Join-Path $PSScriptRoot 'Record-Refunds.ps1') -Workspace $work -BatchId $kw.BatchId -State Uncertain -Evidence 'FAKE missing result' 6>$null
Check 'Missing result blocks cancellation/resubmission' {MustFail {& (Join-Path $PSScriptRoot 'Cancel-PreparedBatch.ps1') -Workspace $work -BatchId $kw.BatchId -Evidence FAKE -ConfirmNeverSubmitted 6>$null}}
& (Join-Path $PSScriptRoot 'Start-Submission.ps1') -Workspace $work -BatchId $benefit.BatchId -ApprovalReference FAKE -ConfirmBusinessAuthorization 6>$null
& (Join-Path $PSScriptRoot 'Record-Refunds.ps1') -Workspace $work -BatchId $benefit.BatchId -Evidence 'FAKE sent email reference' -ConfirmActualSubmission 6>$null
Check 'Manual email submission does not prove success' {$s=Open-Workspace $work;try{Assert ($s.Records[$benefit.RecordIds[0]].FinalStatus -eq 'Unknown') 'Email treated success'}finally{$s.Lock.Dispose()}}
Check 'Communication does not create another attempt' {$s=Open-Workspace $work;try{$n=$s.Attempts.Count}finally{$s.Lock.Dispose()}; & (Join-Path $PSScriptRoot 'Record-Communication.ps1') -Workspace $work -RecordId $benefit.RecordIds[0] -Kind PSPAcknowledged -Evidence 'FAKE acknowledgement' 6>$null;$s=Open-Workspace $work;try{Assert ($s.Attempts.Count -eq $n) 'Communication counted attempt'}finally{$s.Lock.Dispose()}}
Check 'Concurrent writer is rejected' {$s=Open-Workspace $work;try{MustFail {$other=Open-Workspace $work;$other.Lock.Dispose()}}finally{$s.Lock.Dispose()}}
$final=& (Join-Path $PSScriptRoot 'Generate-Reports.ps1') -Workspace $work 6>$null
$match=& (Join-Path $PSScriptRoot 'Match-TrackerUpdates.ps1') -TrackerCsv $input -ManifestCsv (Join-Path $final.ReportPath 'Tracker_Update_Manifest.csv') -OutputFolder (Join-Path $TestOutputFolder 'Match') 6>$null
Check 'Tracker matching uses exact payment order country and accounts for rows' {Assert ($match.Entries -eq 2 -and $match.Ready -eq 2) 'Unexpected matching counts'}
Check 'Journal tampering is detected' {$f=@(Get-ChildItem -LiteralPath (Join-Path $work 'Register\Journal') -Filter '*.json'|Sort-Object Name)[0];$original=Get-Content -LiteralPath $f.FullName -Raw;try{Write-AtomicText $f.FullName ($original.Replace('GENESIS','ALTERED'));MustFail {$bad=Open-Workspace $work;$bad.Lock.Dispose()}}finally{Write-AtomicText $f.FullName $original}}
Check 'Pending preparation blocks new work and can recover safely' {Write-AtomicText (Join-Path $work 'Register\pending-preparation.json') (@{RunId='RUN_FAKE_ABANDONED';SequenceBefore=(Get-Content -LiteralPath (Join-Path $work 'Register\head.json') -Raw|ConvertFrom-Json).Sequence;SourceHash='FAKE'}|ConvertTo-Json);MustFail {$bad=Open-Workspace $work;$bad.Lock.Dispose()}; & (Join-Path $PSScriptRoot 'Recover-Preparation.ps1') -Workspace $work 6>$null;Assert (-not(Test-Path -LiteralPath (Join-Path $work 'Register\pending-preparation.json'))) 'Pending marker not cleared'}
Write-CsvFile (Join-Path $TestOutputFolder 'Test_Results.csv') $checks @('Test','Result','Detail')
Write-Host "Tests: $($checks.Count); passed: $(@($checks|Where-Object Result -eq 'PASS').Count); failed: $(@($checks|Where-Object Result -eq 'FAIL').Count). Files: $TestOutputFolder"
if($KeepExampleOutputs){$target=Join-Path $app 'Examples\Generated';[IO.Directory]::CreateDirectory($target)|Out-Null;Copy-Item -LiteralPath (Join-Path $TestOutputFolder 'Test_Results.csv') -Destination (Join-Path $target 'Test_Results.csv');Copy-Item -LiteralPath $input -Destination (Join-Path $app 'Examples\Tracker_Fake.csv');foreach($name in @('Dashboard.html','Refund_Success_Rate_Report.csv','Breakdowns.csv','Tracker_Update_Manifest.csv')){Copy-Item -LiteralPath (Join-Path $final.ReportPath $name) -Destination (Join-Path $target $name)};Copy-Item -LiteralPath (Join-Path $second.ReportPath 'Run_Sheet.txt') -Destination (Join-Path $target 'Run_Sheet.txt')}
if(@($checks|Where-Object Result -eq 'FAIL').Count){throw 'Refund tests failed. Read Test_Results.csv.'}
