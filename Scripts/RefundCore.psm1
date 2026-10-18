#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-JsonText([string]$Text) {
    # PowerShell 7.5+ otherwise converts ISO strings into DateTime, changing evidence
    # comparisons and culture-sensitive CSV output. Windows PowerShell keeps strings.
    if ((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')) {
        return ConvertFrom-Json -InputObject $Text -DateKind String
    }
    if ($PSVersionTable.PSVersion.Major -gt 5) {
        throw 'Use Windows PowerShell 5.1 or PowerShell 7.5+ to preserve journal timestamp strings.'
    }
    return ConvertFrom-Json -InputObject $Text
}

function Get-Text($Object, [string]$Name, [string]$Default = '') {
    if ($null -eq $Object) { return $Default }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    return [string]$p.Value
}
function Get-Setting($Object, [string]$Name) {
    if ($null -eq $Object) { return $null }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -ne $p) { return $p.Value }
    return $null
}
function New-Map { return ,([System.Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)) }
function Get-Sha([string]$Text) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-','').ToLowerInvariant() } finally { $sha.Dispose() }
}
function Get-FileSha([string]$Path) { return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Write-AtomicText([string]$Path, [string]$Text) {
    $dir = Split-Path -Parent $Path
    [IO.Directory]::CreateDirectory($dir) | Out-Null
    $tmp = Join-Path $dir ('.tmp-' + [guid]::NewGuid().ToString('N'))
    try {
        $bytes = [Text.UTF8Encoding]::new($false).GetBytes($Text)
        $fs = [IO.FileStream]::new($tmp,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
        try { $fs.Write($bytes,0,$bytes.Length); $fs.Flush($true) } finally { $fs.Dispose() }
        if (Test-Path -LiteralPath $Path) { [IO.File]::Replace($tmp,$Path,($Path+'.bak')) }
        else { [IO.File]::Move($tmp,$Path) }
    } finally { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp } }
}
function Write-CsvFile([string]$Path, $Rows, [string[]]$Columns, [switch]$Machine) {
    # Stream a quoted CSV into a temporary file, then atomically publish it.
    # Large trackers must not require another in-memory copy of every cell.
    $dir = Split-Path -Parent $Path
    [IO.Directory]::CreateDirectory($dir) | Out-Null
    $tmp = Join-Path $dir ('.tmp-' + [guid]::NewGuid().ToString('N'))
    try {
        $fs = [IO.FileStream]::new($tmp,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
        $writer = [IO.StreamWriter]::new($fs,[Text.UTF8Encoding]::new($false))
        try {
            $writer.NewLine = "`r`n"
            $writer.WriteLine((($Columns | ForEach-Object { '"' + $_.Replace('"','""') + '"' }) -join ','))
            foreach ($row in $Rows) {
                $cells = [string[]]::new($Columns.Length)
                for($i=0;$i -lt $Columns.Length;$i++) {
                    $p = $row.PSObject.Properties[$Columns[$i]]
                    $v = if($null -eq $p -or $null -eq $p.Value){''}else{[string]$p.Value}
                    if (-not $Machine -and $v -match '^[\s]*[=+@-]') { $v = "'" + $v }
                    $cells[$i] = '"' + $v.Replace('"','""') + '"'
                }
                $writer.WriteLine([string]::Join(',',$cells))
            }
            $writer.Flush(); $fs.Flush($true)
        } finally { $writer.Dispose(); $fs.Dispose() }
        if(Test-Path -LiteralPath $Path){[IO.File]::Replace($tmp,$Path,($Path+'.bak'))}
        else{[IO.File]::Move($tmp,$Path)}
    } finally { if(Test-Path -LiteralPath $tmp){Remove-Item -LiteralPath $tmp} }
}
function Read-CsvTable([string]$Path, [string[]]$Required, [int]$HeaderRow = 1) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Input file not found: $Path" }
    Add-Type -AssemblyName Microsoft.VisualBasic
    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path,[Text.Encoding]::UTF8,$true)
    $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $parser.SetDelimiters(','); $parser.HasFieldsEnclosedInQuotes = $true; $parser.TrimWhiteSpace = $false
    try {
        $logical = 0; $headers = $null
        while (-not $parser.EndOfData) {
            $physicalLine = $parser.LineNumber; $fields = $parser.ReadFields(); $logical++
            if ($null -eq $headers) {
                # Accept a clean CSV or the row-4 Google export; verify labels, never trust position alone.
                if ($logical -eq 1 -or $logical -eq $HeaderRow) {
                    $trim = @($fields | ForEach-Object { $_.Trim().Trim([char]0xfeff) })
                    $missing = @($Required | Where-Object { $_ -notin $trim })
                    if ($missing.Count -eq 0) {
                        $nonblank = @($trim | Where-Object { $_ })
                        if (@($nonblank | Group-Object | Where-Object Count -gt 1).Count) { throw 'Duplicate CSV headers.' }
                        $headers = $trim; continue
                    }
                }
                if ($logical -ge $HeaderRow) { throw 'Required CSV headers were not found. Export the Bulk Refunds tab, not a view.' }
                continue
            }
            if ($fields.Length -gt $headers.Length) { throw "Extra CSV fields at line $physicalLine. No register changes were made." }
            $row = [ordered]@{ InputRow = [string]$physicalLine }
            for ($i=0;$i -lt $headers.Length;$i++) {
                if ($headers[$i]) { $row[$headers[$i]] = if ($i -lt $fields.Length) { $fields[$i] } else { '' } }
            }
            [pscustomobject]$row
        }
        if ($null -eq $headers) { throw 'CSV has no valid header.' }
    } finally { $parser.Close(); $parser.Dispose() }
}
function Get-LocalNow($Config) { return [TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow,[TimeZoneInfo]::FindSystemTimeZoneById($Config.timeZoneId)) }
function Get-IsoWeek([datetime]$Date) {
    $day = [int]$Date.DayOfWeek; if ($day -eq 0) { $day = 7 }
    $thu = $Date.Date.AddDays(4-$day)
    $week = [Globalization.CultureInfo]::InvariantCulture.Calendar.GetWeekOfYear($thu,[Globalization.CalendarWeekRule]::FirstFourDayWeek,[DayOfWeek]::Monday)
    return ('{0}-W{1:00}' -f $thu.Year,$week)
}
function Get-DateValue([string]$Text) {
    $dt = [datetime]::MinValue
    $formats = @('yyyy-MM-dd','yyyy-MM-dd HH:mm:ss','yyyy-MM-dd HH:mm:ss.fff','yyyy-MM-ddTHH:mm:ss','yyyy-MM-ddTHH:mm:ss.fff','dd MMM yyyy','d MMM yyyy','dd-MMM-yyyy')
    if (-not [datetime]::TryParseExact($Text.Trim(),[string[]]$formats,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$dt)) { throw "Invalid/ambiguous date '$Text'. Use yyyy-MM-dd." }
    return $dt.Date
}
function Convert-Amount([string]$Text,[string]$Unit,[string]$Currency,$Config) {
    if ($Text -notmatch '^\d+(\.\d+)?$') { throw 'Amount must be positive decimal text with a dot, no separators or exponent.' }
    $v = [decimal]::Parse($Text,[Globalization.CultureInfo]::InvariantCulture)
    if ($v -le 0) { throw 'Amount must be greater than zero.' }
    $factor = Get-Setting $Config.currencies $Currency
    if ($null -eq $factor -or [decimal]$factor -notin @([decimal]1,[decimal]100,[decimal]1000)) { throw 'Currency multiplier is not configured.' }
    if ($Unit -eq 'major') { $minor = $v * [decimal]$factor; $major = $v }
    elseif ($Unit -eq 'minor') { $minor = $v; $major = $v / [decimal]$factor }
    else { throw 'AmountUnit must be major or minor.' }
    if ($minor -ne [decimal]::Truncate($minor)) { throw 'Amount produces fractional minor units. Do not round.' }
    if ($minor -gt [long]::MaxValue) { throw 'Amount exceeds supported integer range.' }
    [pscustomobject]@{ Major=$major.ToString('0.############################',[Globalization.CultureInfo]::InvariantCulture); Minor=$minor.ToString('0',[Globalization.CultureInfo]::InvariantCulture) }
}
function Open-Workspace([string]$Workspace,[switch]$AllowPending) {
    $root = [IO.Path]::GetFullPath($Workspace)
    $marker = Join-Path $root 'Register\workspace.json'
    if (-not (Test-Path -LiteralPath $marker)) { throw 'Register is not initialized. Run Initialize-Workspace.ps1 once. Never initialize a replacement register to bypass missing history.' }
    $lockPath = Join-Path $root 'Register\.lock'
    try { $lock = [IO.FileStream]::new($lockPath,[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None) }
    catch { throw 'Another operation is using this register. Close it or wait; do not delete the lock file.' }
    try {
        if (-not $AllowPending -and (Test-Path -LiteralPath (Join-Path $root 'Register\pending-preparation.json'))) { throw 'An interrupted preparation needs recovery. Run Recover-Preparation.ps1; do not submit files from the interrupted run.' }
        $config = Read-JsonText (Get-Content -LiteralPath (Join-Path $root 'Config\config.json') -Raw)
        if ($config.schemaVersion -ne 1) { throw 'Unsupported configuration version.' }
        if ([int]$config.tap.batchSize -lt 1 -or [int]$config.tap.batchSize -gt 1000) { throw 'Tap batchSize must be between 1 and 1000.' }
        $s = [pscustomobject]@{Root=$root;Lock=$lock;Config=$config;Records=(New-Map);Batches=(New-Map);Attempts=(New-Map);Imports=(New-Map);Events=[Collections.Generic.List[object]]::new();Sequence=0;LastHash='GENESIS'}
        $files = @(Get-ChildItem -LiteralPath (Join-Path $root 'Register\Journal') -Filter '*.json' | Sort-Object Name)
        foreach ($f in $files) {
            $envelope = Read-JsonText (Get-Content -LiteralPath $f.FullName -Raw)
            if ((Get-Sha $envelope.Payload) -cne $envelope.Hash) { throw "Register integrity failure: $($f.Name). Stop and restore verified history." }
            $tx = Read-JsonText $envelope.Payload
            if ($tx.Schema -ne 1) { throw 'Unsupported journal schema. Do not downgrade or rewrite the register.' }
            if ($tx.Sequence -ne ($s.Sequence+1) -or $tx.PreviousHash -cne $s.LastHash) { throw 'Register journal chain is incomplete or reordered.' }
            Apply-Transaction $s $tx
            $s.Sequence = $tx.Sequence; $s.LastHash = $envelope.Hash
        }
        $headPath = Join-Path $root 'Register\head.json'
        if (Test-Path -LiteralPath $headPath) {
            $head = Get-Content -LiteralPath $headPath -Raw | ConvertFrom-Json
            if ($head.Sequence -gt $s.Sequence -or ($head.Sequence -eq $s.Sequence -and $head.Hash -cne $s.LastHash)) { throw 'Committed history is missing or altered. Restore a complete backup.' }
            # A journal commit is durable before the advisory head. Recover a lagging head after a crash.
            if ($head.Sequence -lt $s.Sequence) { Write-AtomicText $headPath (@{Sequence=$s.Sequence;Hash=$s.LastHash}|ConvertTo-Json) }
        } elseif ($s.Sequence -gt 0) { throw 'Register head is missing. Restore a complete backup.' }
        return $s
    } catch { $lock.Dispose(); throw }
}
function Apply-Transaction($State,$Tx) {
    foreach ($r in @($Tx.Records)) { $State.Records[$r.RecordId] = $r }
    foreach ($b in @($Tx.Batches)) { $State.Batches[$b.BatchId] = $b }
    foreach ($a in @($Tx.Attempts)) { $State.Attempts[$a.AttemptId] = $a }
    foreach ($h in @($Tx.Imports)) { $State.Imports[[string]$h] = $true }
    foreach ($e in @($Tx.Events)) { $State.Events.Add($e) }
}
function Save-Transaction($State,[string]$Kind,$Records=@(),$Batches=@(),$Attempts=@(),$Events=@(),$Imports=@()) {
    $tx = [ordered]@{Schema=1;Sequence=($State.Sequence+1);PreviousHash=$State.LastHash;Kind=$Kind;At=[datetime]::UtcNow.ToString('o');Actor=[Environment]::UserName;Records=@($Records);Batches=@($Batches);Attempts=@($Attempts);Events=@($Events);Imports=@($Imports)}
    $payload = $tx | ConvertTo-Json -Depth 25 -Compress
    $hash = Get-Sha $payload
    $path = Join-Path $State.Root ('Register\Journal\{0:00000000}-{1}.json' -f $tx.Sequence,[guid]::NewGuid().ToString('N'))
    Write-AtomicText $path (@{Payload=$payload;Hash=$hash}|ConvertTo-Json -Depth 3 -Compress)
    $State.Sequence = $tx.Sequence; $State.LastHash = $hash
    Apply-Transaction $State ([pscustomobject]$tx)
    Write-AtomicText (Join-Path $State.Root 'Register\head.json') (@{Sequence=$State.Sequence;Hash=$hash}|ConvertTo-Json)
}
function Get-Route([string]$Raw,[string]$Source,$Config) {
    $parts = @($Raw.Split(@('||'),[StringSplitOptions]::None) | ForEach-Object { $_.Trim() })
    $pays = @($parts | Where-Object { $_ -cmatch '^pay_' })
    $unique = @($pays | Sort-Object -Unique)
    if ($pays.Count) {
        if (@($pays | Where-Object { $_ -cnotmatch '^pay_[a-z0-9]{26}$' }).Count) { return @{PSP='Checkout';ID=$Raw;Error='Invalid Checkout payment ID; require pay_ plus 26 lowercase letters/digits.'} }
        if ($unique.Count -ne 1) { return @{PSP='Checkout';ID=$Raw;Error='Multiple distinct Checkout payment IDs. Amount allocation requires investigation.'} }
        return @{PSP='Checkout';ID=$unique[0];Error=''}
    }
    $mapped = Get-Setting $Config.pspMappings $Source
    if (-not $mapped) { $mapped = 'Unknown' }
    $errorText = ''
    if ($mapped -eq 'Unknown') { $errorText = 'Unrecognized reconciliation source.' }
    elseif ($mapped -eq 'Checkout') { $errorText = 'Checkout source without a valid pay_ payment ID.' }
    if (-not $Raw -or @($parts | Where-Object { $_ -match '^(NA|N/A|null|none)?$' }).Count -gt 0) { $errorText = 'Missing or placeholder payment identifier.' }
    elseif ($parts.Count -gt 1) { $errorText = 'Multiple non-Checkout identifiers need manual interpretation.' }
    return @{PSP=[string]$mapped;ID=$Raw;Error=$errorText}
}
function Get-MissingFields($Record,$Config) {
    $list=[Collections.Generic.List[string]]::new()
    if (-not $Record.Currency) { $list.Add('Currency') }
    switch ($Record.PSP) {
        'Tap' {
            if ($Record.ChargeId -cnotmatch '^chg_[A-Za-z0-9_]+$') { $list.Add('Tap charge_id (not auth_ or tr_)') }
            if ($Record.Reason -cnotin @($Config.tap.allowedReasons)) { $list.Add('Approved Tap reason') }
            if (-not $Config.tap.templateConfirmed) { $list.Add('Current Tap template/reasons confirmation') }
        }
        'Switch' { if (-not (Get-Setting $Config.psps 'Switch').templateConfirmed) { $list.Add('Confirm shared HyperPay/Switch format and blank Transaction ID acceptance') } }
        'ZainCash' { if (-not $Record.OperationId) { $list.Add('Operation ID') } }
        'Knet' { $list.Add('Verified Knet transaction ID, current accepted template and authorized sender') }
        'Deema' { $list.Add('Deema contact, accepted bulk format and identifiers') }
        'TapAcquirer' { $list.Add('Acquirer identifier mapping and approved attachment format') }
        'Unknown' { $list.Add('Approved PSP mapping and contact') }
    }
    return ($list -join '; ')
}
$script:RegisterColumns = @('RecordId','Source','PaymentId','OriginalPaymentId','OrderId','Country','PSP','ReconciliationSource','TransactionDate','OriginalAmount','AmountUnit','AmountMajor','AmountMinor','Currency','TrackerStatus','ProcessedAt','RunId','PreparedAt','SubmittedAt','SubmissionMethod','SubmissionStatus','FinalStatus','FinalConfirmationAt','FailureReason','ActionId','EmailSentAt','PSPResponseAt','Evidence','RequestReference','ApprovalReference','ChargeId','OperationId','Reason','MissingFields','HoldReason','ActiveBatchId','InputFileHash','InputRow','CreatedAt','UpdatedAt')
$script:AttemptColumns = @('AttemptId','RecordId','BatchId','RunId','AttemptedAt','Method','SubmissionStatus','ActionId','FailureReason','Evidence','TrackerEligible')
$script:ReviewColumns = @('RecordId','PSP','Country','PaymentId','OriginalPaymentId','OrderId','TransactionDate','AgeDays','OriginalAmount','AmountUnit','AmountMajor','Currency','TrackerStatus','ProcessedAt','SubmissionStatus','FinalStatus','ChargeId','OperationId','MissingFields','Category','ReasonForReview','InputRow','InputFileHash','RunId')
function Convert-InputRow($Row,[string]$Source,$Config,[string]$Hash,[string]$RunId,[datetime]$AsOf) {
    $isManual = $Source -eq 'manual'
    $raw = (Get-Text $Row $(if($isManual){'Payment ID'}else{'PSP Payment IDs'})).Trim()
    $sourceName=(Get-Text $Row $(if($isManual){'PSP/source'}else{'Reconciliation Source'})).Trim()
    $route=Get-Route $raw $sourceName $Config
    $country=(Get-Text $Row 'Country Code' (Get-Text $Row 'Country')).Trim().ToUpperInvariant()
    $currency=(Get-Text $Row 'Currency').Trim().ToUpperInvariant()
    if (-not $currency) { $currency=[string](Get-Setting $Config.countryCurrencies $country) }
    $unit=Get-Text $Row 'AmountUnit' $Config.trackerAmountUnit
    $amount=(Get-Text $Row $(if($isManual){'Amount'}else{'PSP Capture Amount LC'})).Trim()
    $status=if($isManual){'Pending'}else{(Get-Text $Row 'Refund Status').Trim()}
    $identity=$route.PSP+'|'+$route.ID
    if (-not $raw) { $identity='MISSING|'+$Hash+'|'+$Row.InputRow }
    $now=[datetime]::UtcNow.ToString('o')
    $r=[ordered]@{}
    foreach($c in $script:RegisterColumns){$r[$c]=''}
    $r.RecordId='RF-'+(Get-Sha $identity).Substring(0,24);$r.Source=$Source;$r.PaymentId=$route.ID;$r.OriginalPaymentId=$raw
    $r.OrderId=(Get-Text $Row 'Order ID').Trim();$r.Country=$country;$r.PSP=$route.PSP;$r.ReconciliationSource=$sourceName
    $r.TransactionDate=(Get-Text $Row 'Transaction Date').Trim();$r.OriginalAmount=$amount;$r.AmountUnit=$unit;$r.Currency=$currency
    $r.TrackerStatus=$status;$r.ProcessedAt=Get-Text $Row 'Refund Processed At';$r.RunId=$RunId;$r.InputFileHash=$Hash;$r.InputRow=$Row.InputRow
    $r.RequestReference=Get-Text $Row 'Request reference';$r.ApprovalReference=Get-Text $Row 'Approval reference';$r.Reason=Get-Text $Row 'Reason' ([string]$Config.tap.defaultReason)
    $r.ChargeId=Get-Text $Row 'Charge ID';if(-not $r.ChargeId -and $raw -cmatch '^chg_'){$r.ChargeId=$raw}
    $r.OperationId=Get-Text $Row 'Operation ID';$r.CreatedAt=$now;$r.UpdatedAt=$now;$r.FinalStatus='Unknown';$r.SubmissionStatus='NotSubmitted'
    if($status -in @('Refund Failed','Refund Processed','Refunded','Refunded via Acquirer') -or $r.ProcessedAt){$r.SubmissionStatus='LegacyAttempt'}
    $errors=[Collections.Generic.List[string]]::new();if($route.Error){$errors.Add($route.Error)}
    if(-not $r.OrderId){$errors.Add('Order ID missing.')}
    foreach($identifier in @($r.OrderId,$raw,$r.ChargeId,$r.OperationId)){if($identifier -match '^[\s]*[=+@-]' -or $identifier -match '[\r\n\t]'){$errors.Add('Identifier contains a formula prefix or control character. Review exact source value.')}}
    if($country -notmatch '^[A-Z]{2}$' -or -not (Get-Setting $Config.countryCurrencies $country)){$errors.Add('Unsupported/missing country.')}
    $age=''
    try{$dt=Get-DateValue $r.TransactionDate;$r.TransactionDate=$dt.ToString('yyyy-MM-dd');$age=[string]($AsOf.Date-$dt).Days;if([int]$age -lt 0){$errors.Add('Transaction date is in the future.')}}catch{$errors.Add($_.Exception.Message)}
    try{$converted=Convert-Amount $amount $unit $currency $Config;$r.AmountMajor=$converted.Major;$r.AmountMinor=$converted.Minor}catch{$errors.Add($_.Exception.Message)}
    if($isManual -and -not $r.RequestReference){$errors.Add('Manual request reference missing.')}
    $record=[pscustomobject]$r;$record.MissingFields=Get-MissingFields $record $Config
    [pscustomobject]@{Record=$record;Age=$age;Error=($errors -join '; ')}
}
function Get-RunFolder($State,$Record,[datetime]$Date,[string]$RunId,[switch]$Delayed) {
    $pspCfg=Get-Setting $State.Config.psps $Record.PSP
    $psp=if($Record.PSP -match '^[A-Za-z0-9]+$'){$Record.PSP}else{'Unknown'}
    $country=if($Record.Country -match '^[A-Z]{2}$'){$Record.Country}else{'Unknown'}
    $parts=@('Output',$psp)
    if($Delayed -or $null -eq $pspCfg -or $pspCfg.normalCountryFolders){$parts+=$country}
    $parts+=@($Date.ToString('yyyy'),$Date.ToString('MM'),(Get-IsoWeek $Date),$RunId)
    if($Delayed){$parts+='Delayed_30plus_days'}
    return Join-Path $State.Root ($parts -join [IO.Path]::DirectorySeparatorChar)
}
function Get-ProviderRows([string]$PSP,$Records) {
    foreach($r in $Records){
        switch($PSP){
            'Checkout' {[pscustomobject]@{id=$r.PaymentId;amount=$r.AmountMinor}}
            'Tap' {[pscustomobject]@{charge_id=$r.ChargeId;amount=$r.AmountMajor;currency=$r.Currency;reason=$r.Reason;reference=$r.OrderId}}
            'BenefitPay' {[pscustomobject]@{Order_ID=$r.OrderId;Value=$r.AmountMajor;Currency=$r.Currency;OrderDate=$r.TransactionDate}}
            'Switch' {[pscustomobject]@{'Order Date'=$r.TransactionDate;Amount=$r.AmountMajor;'Transaction ID'='';'Order ID'=$r.OrderId;'PSP Payment IDs'=$r.PaymentId;'Reconciliation Source'=$r.ReconciliationSource}}
            'ZainCash' {[pscustomobject]@{'Transaction Date'=$r.TransactionDate;'Transaction Month'=$r.TransactionDate.Substring(0,7);'PSP Payment IDs'=$r.PaymentId;'Order ID'=$r.OrderId;Amount=$r.AmountMajor;Currency=$r.Currency;'Operation ID'=$r.OperationId}}
            default {[pscustomobject]@{'Transaction Date'=$r.TransactionDate;'Transaction Month'=$r.TransactionDate.Substring(0,7);'PSP Payment IDs'=$r.PaymentId;'Order ID'=$r.OrderId;Amount=$r.AmountMajor;Currency=$r.Currency}}
        }
    }
}
function Write-EmailDraft($State,[string]$Folder,[string]$PSP,[string]$Country,[switch]$Review) {
    $contacts=Get-Content -LiteralPath (Join-Path $State.Root 'Config\contacts.json') -Raw|ConvertFrom-Json
    $c=Get-Setting $contacts $PSP;if($null -eq $c){$c=$contacts.Unknown}
    $purpose=if($Review){'INVESTIGATION / ESCALATION ONLY - not an instruction to issue duplicate refunds'}else{'Refund request - review and approve before manually sending'}
    $body=@("To: $($c.to -join '; ')","Cc: $($c.cc -join '; ')","External escalation options (confirm before use): $($c.external -join '; ')","Subject: $PSP $Country refund review",'', $purpose,'', 'Please review the attached transaction list, confirm prior refund/void status and remaining refundable amount, and advise next steps. Missing fields and prior attempts are explicitly listed. Do not process duplicate or uncertain payments without reconciliation.','', $c.note,'', 'Sending this review email is a communication event, not proof of a refund attempt or success. Record actual refund requests separately with evidence.') -join "`r`n"
    Write-AtomicText (Join-Path $Folder 'Email_Draft.txt') $body
}
function Export-Register($State) {
    Write-CsvFile (Join-Path $State.Root 'Register\RefundRegister.csv') @($State.Records.Values) $script:RegisterColumns
    Write-CsvFile (Join-Path $State.Root 'Register\RefundAttempts.csv') @($State.Attempts.Values) $script:AttemptColumns
    Write-CsvFile (Join-Path $State.Root 'Register\AuditEvents.csv') @($State.Events) @('At','Type','RecordId','BatchId','Reference','Detail')
}
Export-ModuleMember -Function * -Variable RegisterColumns,AttemptColumns,ReviewColumns
