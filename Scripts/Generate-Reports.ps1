#requires -Version 5.1
[CmdletBinding()]param([string]$Workspace)
$ErrorActionPreference='Stop'
if(-not $Workspace){$Workspace=Split-Path $PSScriptRoot -Parent}
Import-Module (Join-Path $PSScriptRoot 'RefundCore.psm1') -Force -DisableNameChecking
$s=Open-Workspace $Workspace
try{
 $rows=@($s.Records.Values);$attempts=@($s.Attempts.Values);$valid=New-Map
 foreach($a in $attempts){if($a.TrackerEligible -eq $true){$valid[$a.RecordId]=$true}}
 $success=@($rows|Where-Object FinalStatus -eq 'ConfirmedSuccessful').Count;$failed=@($rows|Where-Object FinalStatus -eq 'ConfirmedFailed').Count
 $verifiedDenom=$success+$failed;$verified=if($verifiedDenom){([decimal]$success/[decimal]$verifiedDenom).ToString('0.0000',[Globalization.CultureInfo]::InvariantCulture)}else{''}
 $submissionRate=if($rows.Count){([decimal]$valid.Count/[decimal]$rows.Count).ToString('0.0000',[Globalization.CultureInfo]::InvariantCulture)}else{''}
 $reserved=@($rows|Where-Object SubmissionStatus -eq 'Reserved').Count;$unknown=@($rows|Where-Object{($_.SubmissionStatus -ne 'NotSubmitted') -and $_.FinalStatus -eq 'Unknown'}).Count
 $metrics=@(
 [pscustomobject]@{Metric='Total registered refund requests';Value=$rows.Count;Numerator=$rows.Count;Denominator='';Formula='Unique register RecordIds, including legacy records';Note='Repeated input rows do not increase unique requests.'},
 [pscustomobject]@{Metric='Ever prepared';Value=@($rows|Where-Object PreparedAt).Count;Numerator='';Denominator='';Formula='Records with PreparedAt';Note='Not an attempt. Includes subsequently cancelled preparations.'},
 [pscustomobject]@{Metric='Recorded attempted/submitted with evidence';Value=$valid.Count;Numerator=$valid.Count;Denominator='';Formula='Unique records with affirmative submission evidence';Note='CLI false/missing results are excluded because they may not establish an API attempt.'},
 [pscustomobject]@{Metric='CLI request errors requiring investigation';Value=@($attempts|Where-Object SubmissionStatus -eq 'RequestFailed').Count;Numerator='';Denominator='';Formula='CLI refund_initiated=false with error';Note='Blocked permanently from automatic resubmission; not a verified final failure.'},
 [pscustomobject]@{Metric='Confirmed successful';Value=$success;Numerator=$success;Denominator='';Formula='FinalStatus = ConfirmedSuccessful';Note='Requires explicit PSP evidence.'},
 [pscustomobject]@{Metric='Confirmed failed';Value=$failed;Numerator=$failed;Denominator='';Formula='FinalStatus = ConfirmedFailed';Note='Tracker failure signals and CLI request errors are excluded.'},
 [pscustomobject]@{Metric='Pending verification';Value=$unknown;Numerator=$unknown;Denominator='';Formula='Non-new records with unknown final outcome';Note='Includes legacy, reserved and uncertain records; not evidence of failure.'},
 [pscustomobject]@{Metric='Reserved, outcome not recorded';Value=$reserved;Numerator='';Denominator='';Formula='SubmissionStatus = Reserved';Note='Investigate before any further action.'},
 [pscustomobject]@{Metric='Legacy attempt signals';Value=@($rows|Where-Object SubmissionStatus -eq 'LegacyAttempt').Count;Numerator='';Denominator='';Formula='Imported historical processed/refunded/failed evidence';Note='Blocked from resubmission; not verified success.'},
 [pscustomobject]@{Metric='Verified success rate';Value=$verified;Numerator=$success;Denominator=$verifiedDenom;Formula='ConfirmedSuccessful / (ConfirmedSuccessful + ConfirmedFailed)';Note='Blank if no verified outcomes. Read beside pending verification count.'},
 [pscustomobject]@{Metric='Recorded submission rate';Value=$submissionRate;Numerator=$valid.Count;Denominator=$rows.Count;Formula='Recorded attempted/submitted unique records / all registered requests';Note='Legacy history is not assumed complete; denominator includes legacy requests.'},
 [pscustomobject]@{Metric='Confirmed voided';Value=@($rows|Where-Object FinalStatus -eq 'Voided').Count;Numerator='';Denominator='';Formula='FinalStatus = Voided';Note='Excluded from refund success denominator.'}
 )
 $folder=Join-Path $s.Root ('Reports\Dashboard_'+(Get-LocalNow $s.Config).ToString('yyyyMMdd_HHmmss')+'_'+[guid]::NewGuid().ToString('N').Substring(0,6))
 [IO.Directory]::CreateDirectory($folder)|Out-Null
 Write-CsvFile (Join-Path $folder 'Refund_Success_Rate_Report.csv') $metrics @('Metric','Value','Numerator','Denominator','Formula','Note')
 $breakdown=[Collections.Generic.List[object]]::new()
 foreach($dimension in @('PSP','Country','Source','TransactionMonth','TransactionISOWeek','SubmissionMonth','SubmissionISOWeek')){
  $groups=New-Map
  foreach($r in $rows){
   $key=if($dimension -eq 'TransactionMonth'){$r.TransactionDate.Substring(0,[Math]::Min(7,$r.TransactionDate.Length))}elseif($dimension -eq 'TransactionISOWeek'){try{Get-IsoWeek (Get-DateValue $r.TransactionDate)}catch{'Invalid date'}}elseif($dimension -like 'Submission*'){if($r.SubmittedAt){$d=[datetime]::Parse($r.SubmittedAt,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::RoundtripKind);if($dimension -eq 'SubmissionMonth'){$d.ToString('yyyy-MM')}else{Get-IsoWeek $d}}else{'Not recorded'}}else{Get-Text $r $dimension}
   $key=$key+'|'+$r.Currency
   if(-not $groups.ContainsKey($key)){$groups[$key]=[Collections.Generic.List[object]]::new()};$groups[$key].Add($r)
  }
  foreach($key in $groups.Keys){$group=@($groups[$key]);$sum=[decimal]0;$amountMissing=0;foreach($r in $group){if($r.AmountMajor){$sum+=[decimal]::Parse($r.AmountMajor,[Globalization.CultureInfo]::InvariantCulture)}else{$amountMissing++}};$ok=@($group|Where-Object FinalStatus -eq 'ConfirmedSuccessful').Count;$bad=@($group|Where-Object FinalStatus -eq 'ConfirmedFailed').Count
   $breakdown.Add([pscustomobject]@{Dimension=$dimension;Group=$key.Split('|')[0];Currency=$group[0].Currency;Requested=$group.Count;KnownAmountMajor=$sum.ToString([Globalization.CultureInfo]::InvariantCulture);MissingAmounts=$amountMissing;Attempted=@($group|Where-Object{$valid.ContainsKey($_.RecordId)}).Count;ConfirmedSuccessful=$ok;ConfirmedFailed=$bad;UnknownFinal=@($group|Where-Object FinalStatus -eq 'Unknown').Count;VerifiedSuccessRate=$(if(($ok+$bad)-gt 0){([decimal]$ok/($ok+$bad)).ToString('0.0000',[Globalization.CultureInfo]::InvariantCulture)}else{''})})
  }
 }
 Write-CsvFile (Join-Path $folder 'Breakdowns.csv') $breakdown @('Dimension','Group','Currency','Requested','KnownAmountMajor','MissingAmounts','Attempted','ConfirmedSuccessful','ConfirmedFailed','UnknownFinal','VerifiedSuccessRate')
 $manifest=@(foreach($a in $attempts){if($a.TrackerEligible -eq $true){$r=$s.Records[$a.RecordId];if($r.Source -match 'tracker'){[pscustomobject]@{RecordId=$r.RecordId;'Payment ID'=$r.PaymentId;'Original Payment IDs'=$r.OriginalPaymentId;'Order ID'=$r.OrderId;Country=$r.Country;'Original status'=$r.TrackerStatus;'New status'='Refund Processed';'Attempt date'=$a.AttemptedAt;'Run ID'=$a.RunId;'Evidence/reference'=$a.Evidence}}}})
 Write-CsvFile (Join-Path $folder 'Tracker_Update_Manifest.csv') $manifest @('RecordId','Payment ID','Original Payment IDs','Order ID','Country','Original status','New status','Attempt date','Run ID','Evidence/reference')
 # Latest preparation counts are explicitly row counts, not cumulative unique requests.
 $runs=@(Get-ChildItem -LiteralPath (Join-Path $s.Root 'Reports') -Directory -Filter 'RUN_*'|Sort-Object Name -Descending)
 $latest=@();if($runs.Count){$path=Join-Path $runs[0].FullName 'Preparation_Summary.csv';if(Test-Path -LiteralPath $path){$latest=@(Import-Csv -LiteralPath $path);Write-CsvFile (Join-Path $folder 'Latest_Preparation_Counts.csv') $latest @('Category','PSP','Country','Rows')}}
 $html='<html><head><meta charset="utf-8"><title>Refund dashboard</title><style>body{font:16px Segoe UI,Arial;margin:36px;color:#172b4d}table{border-collapse:collapse;margin:24px 0}th,td{padding:8px 12px;border-bottom:1px solid #ddd;text-align:left}th{background:#183b56;color:white}p{max-width:1000px}</style></head><body><h1>Refund dashboard</h1><p>Snapshot '+[Net.WebUtility]::HtmlEncode((Get-LocalNow $s.Config).ToString('yyyy-MM-dd HH:mm'))+' (Dubai). Amount totals are separated by currency. Legacy tracker statuses are not verified outcomes. Prepared, reserved, requested and confirmed are separate measures.</p>'
 $html+=($metrics|ConvertTo-Html -Fragment|Out-String)
 $html+='<h2>Latest preparation: input row counts</h2><p>Delayed, duplicate-blocked and exception counts below belong only to the latest preparation. One row has one primary category; a delayed row may also contain a prior-attempt warning.</p>'
 $html+=($latest|ConvertTo-Html -Fragment|Out-String)
 $html+='<h2>Breakdown by PSP, with separate currencies</h2>'+(@($breakdown|Where-Object Dimension -eq 'PSP')|ConvertTo-Html -Fragment|Out-String)+'</body></html>'
 Write-AtomicText (Join-Path $folder 'Dashboard.html') $html
 Export-Register $s
 Write-Host "Dashboard and Excel-compatible CSVs: $folder"
 [pscustomobject]@{ReportPath=$folder;TrackerRows=$manifest.Count;VerifiedSuccessful=$success;VerifiedFailed=$failed}
}finally{$s.Lock.Dispose()}
