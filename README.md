# Talabat refund preparation and tracking

A Windows PowerShell 5.1 application for preparing refund files, separating delayed cases, preserving attempt history and reporting verified outcomes. It runs with built-in Windows components. No Python, Git, Excel automation, extra PowerShell modules or administrator access is required on the work laptop.

**Phase 1 never executes a payment, sends an email, reads a secret, edits Google Sheets or uploads to Drive.** You review files and perform those actions manually. The application cannot prevent someone from independently rerunning an old file in a PSP tool; follow the reservation workflow and use one authoritative register.

## What is included

- Checkout CSVs with `id,amount` and exact integer minor units.
- Tap CSV drafts with `charge_id,amount,currency,reason,reference`, major-unit amounts and configurable batch size (100 by default; editable up to 1,000).
- BenefitPay, shared HyperPay/Switch, QPay and ZainCash email files when required data and format confirmations are present.
- Review-only attachments for every PSP with missing information, including Knet, Deema and Tap as acquirer.
- Delayed files split by **PSP and country**, including rows with missing identifiers or prior attempts. These are investigation lists, never automatic refund queues.
- A central integrity-checked transaction journal, generated register/attempt CSVs, manual request intake and corrections.
- Result import, final outcome recording, local tracker matching, an HTML dashboard and Excel-compatible CSV reports with explicit rate definitions.
- Synthetic examples and automated tests. No company transaction files, Checkout executable or credentials are distributed.

## 1. Install on the work laptop

Use your company's permitted method to transfer the ZIP. Email gateways sometimes block ZIPs containing scripts; if yours does, use an IT-approved transfer method instead of disguising the extension. Download the code ZIP/release from this private repository on the home laptop; work-laptop GitHub access is not required.

1. Create `%LOCALAPPDATA%\TalabatRefundAutomation` (normally `C:\Users\YOUR_WORK_USERNAME\AppData\Local\TalabatRefundAutomation`). It is outside the usual OneDrive Desktop/Documents folders. Check that your company has not redirected it.
2. Extract the ZIP contents there. You should see `Scripts`, `Config`, `Templates`, `README.md` and `Run-Prepare.bat` directly inside that folder. Avoid an accidental extra nested folder.
3. Open Windows PowerShell from Start, then paste:

```powershell
$refundRoot = Join-Path $env:LOCALAPPDATA 'TalabatRefundAutomation'
Set-Location -LiteralPath $refundRoot
$PSVersionTable.PSVersion
Get-ExecutionPolicy -List
.\Scripts\Initialize-Workspace.ps1
```

Use built-in Windows PowerShell **5.1**, or PowerShell **7.5 or later**. Older PowerShell 7 versions are rejected because their JSON date conversion can change stored timestamps. The initializer creates `Input`, `Output`, `Register` and `Reports`; it preserves existing history. Do not initialize an empty replacement when history is missing.

If scripts are blocked, follow company IT policy. This project deliberately does not set `Bypass`, change machine policy or use administrator privileges. If IT permits unblocking a verified downloaded ZIP, unblock it through Windows Properties **before extracting**. If signing/allowlisting is required, ask IT to sign/approve the scripts. A `.bat` launcher does not bypass policy.

Create a **separate unsynced** folder such as `%LOCALAPPDATA%\TalabatCheckoutTool`. Obtain the approved executable and `.env` directly through Engineering's approved process. Keep them together there. Never put them in this repository, the transfer ZIP, Google Drive, OneDrive or chat. These scripts never open `.env` or invoke the executable.

## 2. Confirm amounts once with Finance

There are two separate questions:

1. What unit does the input amount use?
2. What format does the destination PSP require?

| Input | Source unit | Checkout output | Tap output |
|---|---|---:|---:|
| AED 25.50 | major | 2550 | 25.5 |
| 2550 AED minor units | minor | 2550 | 25.5 |
| KWD 1.555 | major | 1555 | 1.555 |
| 1555 KWD minor units | minor | 1555 | 1.555 |

The supplied configuration proposes AED/QAR/EGP x100 and BHD/KWD/OMR/JOD/IQD x1000. **A multiplier is applied only to major-unit input.** Country/currency mappings and PSP requirements still need confirmation. The absence of a decimal point does not prove that a value is in minor units: `9000` can mean 9,000 IQD.

Finance must also confirm that `PSP Capture Amount LC` is the approved refund amount for these rows, taking existing refunds, partial captures and wallet components into account. This application cannot retrieve a remaining refundable balance.

Until confirmed, preparation creates review files and no executable batches. After receiving approval, use the following command, replacing the approval reference and choosing `major` or `minor` correctly:

```powershell
.\Scripts\Confirm-AmountConfiguration.ps1 -TrackerAmountUnit major `
  -ApprovalReference 'REPLACE_WITH_FINANCE_APPROVAL_REFERENCE' `
  -ConfirmCurrencyMapping -ConfirmRefundAmountPolicy
```

You can edit `Config\config.json` in Notepad. Keep valid JSON. Optional `Currency` and `AmountUnit` columns in a tracker CSV override configured defaults; manual requests require those headers. Values producing fractional minor units are rejected, never rounded. Dates must be unambiguous, preferably `yyyy-MM-dd`; numeric identifiers must remain text.

## 3. Prepare the weekly tracker

1. Open the controller's Google Sheet and perform its existing manual **Sync** step so the latest refunds appear. The automation does not run Sync.
2. Download **Bulk Refunds**, not Pending Refunds or Refund Failed, as CSV. The importer accepts the original row-4 header or a clean first-row header.
3. Keep the full tracker, including historical statuses, so the first import can block prior processed/failed/refunded payments. `Refunded` is retained as a legacy signal, not independently verified success.
4. Place exactly one tracker CSV in `Input`. Move last week's input to your approved archive, retaining the permanent `Register`.
5. Double-click `Run-Prepare.bat`, or run:

```powershell
.\Scripts\Prepare-Refunds.ps1 -InputCsv '.\Input\BulkRefunds.csv'
```

Replace `BulkRefunds.csv` with the filename you downloaded. Explicit `-InputCsv` is useful if there are multiple files. To force investigation-only output:

```powershell
.\Scripts\Prepare-Refunds.ps1 -InputCsv '.\Input\BulkRefunds.csv' -ReviewOnly
```

Read the newest `Reports\RUN_...\Run_Sheet.txt`, then `Row_Dispositions.csv`, `Preparation_Summary.csv` and `Batches.csv`. Every data row receives one primary category: Ineligible, Exceptions, Duplicates, Delayed, ReviewOnly or Prepared. Delayed rows take priority and retain secondary validation/duplicate/prior-attempt warnings. Blank physical lines are ignored by the CSV parser; malformed CSV stops the import.

Routing always prioritizes a valid `pay_` component. A combined label does not create a second refund. Multiple distinct `pay_` components, placeholder IDs and unknown PSP mappings require investigation. HyperPay maps into the shared Switch workflow. Repeating preparation does not release an existing active batch.

## 4. Understand the folders

Folders use calendar year/month plus an explicitly labelled ISO week-year, so year boundaries remain unambiguous. Each run has a unique suffix.

```text
Output/
  Checkout/AE/2026/10/2026-W43/RUN_.../
    refunds_AE_1.csv
    refunds_AE_1.csv.manifest.csv
    refunds_AE_1.csv.batch.txt
    Delayed_30plus_days/Delayed_30plus_days.csv
    Review_Only/ReviewOnly.csv
  Checkout/QA/2026/10/2026-W43/RUN_.../
  Tap/AE/2026/10/2026-W43/RUN_.../
    tap_AE_1.csv
  Switch/2026/10/2026-W43/RUN_.../
    Switch_IQ_1.csv
  Switch/IQ/2026/10/2026-W43/RUN_.../
    Delayed_30plus_days/Delayed_30plus_days.csv
    Delayed_30plus_days/Email_Draft.txt
Reports/RUN_.../
  Preparation_Summary.csv
  Row_Dispositions.csv
  Delayed.csv
  Exceptions.csv
  Duplicates.csv
  ReviewOnly.csv
  Ineligible.csv
  Batches.csv
  Unmatched_Enrichment.csv
  Run_Sheet.txt
Register/
  workspace.json
  head.json
  Journal/
  Sources/
  Evidence/
  RefundRegister.csv
  RefundAttempts.csv
  AuditEvents.csv
```

Normal email PSPs omit country folders but their filenames/batches still separate countries and currencies. **All delayed and review files include PSP and country folders**, as requested. The delay threshold defaults to 30 days per PSP and can be edited in `psps.<name>.delayedDays`.

For an illustrative input with 50 Checkout AE, 30 Checkout QA, 20 Tap AE, 15 Switch IQ and 10 BenefitPay BH normal rows, those five groups produce separate batches if fully validated. An additional 8 delayed, 5 duplicate and 3 unknown rows produce 8/5/3 review rows, not extra executable refunds. Your real counts are always in the run summary.

## 5. Tap and missing identifiers

The screenshot confirms `charge_id`, `amount`, `currency`, `reason`; optional `reference` carries Order ID. The screenshot's batch limit is **100**. Change `tap.batchSize` to **1000** when confirmed; the code supports either and keeps filenames within 30 characters. Keep `tap.templateConfirmed=false` until Finance/PSP confirms the current headers and reason strings. The screenshot's reason wording is preserved in `allowedReasons` and is editable; no reason is guessed by default.

When charge IDs become available, either add `Charge ID` to the local input CSV or fill a copy of `Templates\IdentifierEnrichment.csv`. Match on **normalized payment ID + Order ID + Country**:

```powershell
.\Scripts\Prepare-Refunds.ps1 -InputCsv '.\Input\BulkRefunds.csv' `
  -EnrichmentCsv '.\Input\IdentifierEnrichment.csv'
```

Use the single extracted `pay_` ID when a tracker cell contains a combined Checkout identifier. For Tap rows use the original `auth_`/`tr_`/other tracker ID as the matching key and put the real charge ID in `Charge ID`. An `auth_` or `tr_` is never transformed into a guessed `chg_`. Duplicate enrichment keys stop preparation, and unused keys are reported. `Operation ID` supports ZainCash in the same file.

## 6. Review, reserve, then submit manually once

Before any submission, check the PSP/account/country, original and converted amounts, row counts, prior attempts and business approval. Use a current PSP report/portal to check captures, prior refunds and remaining refundable balance. An old `Pending` label is not proof of refundability.

Copy a BatchId from `Batches.csv`, then reserve it:

```powershell
.\Scripts\Start-Submission.ps1 -BatchId 'REPLACE_WITH_BATCH_ID' `
  -ApprovalReference 'REPLACE_WITH_APPROVAL_REFERENCE' -ConfirmBusinessAuthorization
```

This **does not execute anything**. It checks the file hash, unchanged configuration, current transaction ages and record holds, then durably blocks the batch from another start. It is not counted as an attempt. Once reserved, if anything is uncertain, investigate instead of rerunning.

### Checkout

Copy only the approved `refunds_COUNTRY_N.csv` into the separate Engineering-approved tool folder. Verify the copy matches the original SHA256 with `Get-FileHash`. A dry run can be performed before reservation/production approval; do not import dry-run output into the operational register.

```powershell
# Replace BOTH paths with your approved file and tool locations.
Get-FileHash -LiteralPath 'C:\REPLACE\original\refunds_AE_1.csv' -Algorithm SHA256
Get-FileHash -LiteralPath 'C:\REPLACE\TalabatCheckoutTool\refunds_AE_1.csv' -Algorithm SHA256
Set-Location -LiteralPath 'C:\REPLACE\TalabatCheckoutTool'
.\tlb-refund-cli-windows.exe --file '.\refunds_AE_1.csv' --dry-run
```

Engineering must confirm the build, environment/key scope, dry-run behavior and retry semantics. The supplied PDF does not document `--environment`, `--timeout` or `--verbose`, so this guide does not invent those flags. Do not assume a dry run checks refundable balances or tests a network request.

After business authorization, reservation and confirmation of the intended production account, the documented manual command is:

```powershell
# REAL MONEY ACTION. Run once only after the checks above.
.\tlb-refund-cli-windows.exe --file '.\refunds_AE_1.csv' --concurrency 1
```

**Never follow the old PDF's advice to rerun a failed/interrupted batch.** This project cannot control retries inside the supplied binary. Retain all result files and investigate uncertain rows. A late refund can fail, but age alone does not establish the failure cause or acquirer.

Return to the automation folder and import the actual result file:

```powershell
Set-Location -LiteralPath (Join-Path $env:LOCALAPPDATA 'TalabatRefundAutomation')
.\Scripts\Import-CheckoutResults.ps1 -BatchId 'REPLACE_WITH_BATCH_ID' `
  -ResultCsv 'C:\REPLACE\TalabatCheckoutTool\refunds_AE_1_result.csv' `
  -Evidence 'REPLACE_WITH_EXECUTION_OR_RESULT_REFERENCE' -ConfirmProductionAttempt
```

The importer checks IDs and amounts against the batch. Reimporting the same file does not create another attempt. Partial results leave missing rows Uncertain. `refund_initiated=true` with an action ID means submission, not confirmed success. False rows remain blocked and do not enter the tracker-update list unless separate affirmative attempt evidence is recorded through an authorized process; this version intentionally does not infer that a network call occurred.

If a result file is missing after possible execution:

```powershell
.\Scripts\Record-Refunds.ps1 -BatchId 'REPLACE_WITH_BATCH_ID' `
  -State Uncertain -Evidence 'REPLACE_WITH_INCIDENT_REFERENCE'
```

### Tap and normal email batches

After reserving the batch, manually upload to Tap or send the approved refund request to the configured recipients. Record actual submission evidence only after that action:

```powershell
.\Scripts\Record-Refunds.ps1 -BatchId 'REPLACE_WITH_BATCH_ID' `
  -Evidence 'REPLACE_WITH_PORTAL_BATCH_ID_OR_SENT_EMAIL_REFERENCE' -ConfirmActualSubmission
```

Tap upload acceptance and sent email are submissions, not final successes. Check the PSP's result/status view later.

### Delayed, blocked and incomplete files

Open the PSP/country review file and `Email_Draft.txt`. Confirm the recipients, then manually email for investigation. Do not submit these attachments as portal batches or silently ask for another refund on previously attempted payments. `Config\contacts.json` includes internal recipients, external escalation options and missing-contact notes. Knet manual requests go through authorized Kuwait Finance senders; QPay older than 90 days needs acquiring-bank routing.

Record a review email without creating a refund attempt:

```powershell
.\Scripts\Record-Communication.ps1 -RecordId 'REPLACE_WITH_RECORD_ID' `
  -Kind Escalation -Evidence 'REPLACE_WITH_SENT_EMAIL_REFERENCE'
```

If a PSP confirms an existing refund, record the final outcome below. If an old/unresolved case truly needs a **new** refund attempt, keep it outside this Phase 1 queue under the controller's exceptional process. There is deliberately no general unlock/retry command.

For a delayed/review record with no existing batch or recorded attempt, you can record evidence of an authorized submission that **already occurred** through the controller/acquirer process. This does not send a request, release a block or authorize a retry:

```powershell
.\Scripts\Record-ExternalAttempt.ps1 -RecordId 'REPLACE_WITH_RECORD_ID' `
  -Method Acquirer -Evidence 'REPLACE_WITH_ACTUAL_REQUEST_REFERENCE' `
  -ApprovalReference 'REPLACE_WITH_CONTROLLER_APPROVAL' `
  -AttemptedAt '2026-10-19T09:30:00+04:00' -ConfirmActualSubmission
```

Replace the timestamp with when the actual request occurred. Review emails alone do not qualify. The same evidence is idempotent; a record with existing batch/attempt history cannot receive another attempt through this command. Its final outcome remains Unknown until independently confirmed.

## 7. Manual refund requests

Copy `Templates\ManualRefundRequests.csv` to `Input`, fill the columns, and obtain business approval. Use an explicit currency and AmountUnit (`major` or `minor`), original transaction date, request reference and approved amount.

```powershell
.\Scripts\Add-ManualRequests.ps1 -InputCsv '.\Input\ManualRefundRequests.csv'
```

These requests use the same register, amount rules, delay logic and duplicate checks. Payment IDs deduplicate across tracker/manual sources; different IDs for the same cross-source order/country are held for investigation. Missing approval creates a review file, never a ready batch.

## 8. Update the tracker safely

```powershell
.\Scripts\Generate-Reports.ps1
```

Open the newest `Reports\Dashboard_...\Tracker_Update_Manifest.csv`. It includes only tracker-linked records with affirmative submission evidence. Prepared/reserved/uncertain, review-email-only and legacy-only records do not appear. Previously recorded attempts remain in the cumulative manifest; already processed rows will be recognized during matching.

Download a fresh full tracker CSV after its manual Sync and run:

```powershell
.\Scripts\Match-TrackerUpdates.ps1 -TrackerCsv '.\Input\FreshBulkRefunds.csv' `
  -ManifestCsv 'C:\REPLACE\Reports\Dashboard_RUN\Tracker_Update_Manifest.csv'
```

Import `Reviewed_Tracker_Matches.csv` into a local Excel workbook. Reconcile:

`manifest entries = ready for human review + already processed + entries requiring investigation`.

Require exactly one match by payment ID, order ID and country. Investigate unmatched, duplicate and changed-status rows before editing. The row number is from that exact downloaded snapshot, **not a guaranteed live Google Sheets row**.

Coordinate a stable editing window with the controller; verify the keys of each intended live row, then manually change only its `Refund Status` cell to `Refund Processed`. The sheet automatically stamps the time. Do not write timestamp cells yourself. Do not paste over an entire filtered selection or replace the whole status column from an old export. If the controller authorizes a temporary helper/matching list, it can speed this up; the application never adds one. Without stable rows or an authorized helper, individual verified updates remain the safe fallback. Automated row-level tracker writing is Phase 2, not implemented.

Export again afterward and rerun matching to confirm the expected count is `AlreadyProcessed`. Keep the reviewed match report as evidence.

## 9. Record final outcomes and show your manager

```powershell
.\Scripts\Record-ManualOutcome.ps1 -RecordId 'REPLACE_WITH_RECORD_ID' `
  -Status ConfirmedSuccessful -Evidence 'REPLACE_WITH_PSP_CONFIRMATION_REFERENCE'

# For a confirmed failure:
.\Scripts\Record-ManualOutcome.ps1 -RecordId 'REPLACE_WITH_RECORD_ID' `
  -Status ConfirmedFailed -Evidence 'REPLACE_WITH_PSP_CONFIRMATION_REFERENCE' `
  -FailureReason 'REPLACE_WITH_CONFIRMED_REASON'

.\Scripts\Generate-Reports.ps1
```

Use `-Status Voided` for a confirmed authorization void; it is not counted as a successful refund. Changing a prior final outcome requires `-Correction` and new evidence. The journal retains prior versions.

Open `Dashboard.html` in your browser or the CSVs in Excel. Verified success rate is `ConfirmedSuccessful / (ConfirmedSuccessful + ConfirmedFailed)`, blank when the denominator is zero. Pending verification is shown alongside it. The submission rate counts affirmative recorded submissions over unique registered requests, including legacy records in the denominator; historical coverage is therefore explicitly limited. CLI false errors and uncertain rows are separate. Amounts are grouped by currency, never summed across currencies. Breakdown periods distinguish transaction cohorts and recorded submission cohorts. Preparation category totals are the **latest run's row counts**, not cumulative unique refunds.

The tracker reappearing after three days is only a failure signal. Disappearing from it is not proof of success. A PSP-confirmed refund does not prove the customer's bank has credited their account; keep any ARN/reference in the evidence.

## 10. Excel, backup and next week

Do not double-click PSP CSVs and resave them from Excel. Use **Data > From Text/CSV**, set payment/order/operation identifiers to **Text**, use a dot-decimal locale, and preserve dates. A CSV cannot force Excel to preserve long numeric identifiers or leading zeros. Machine files use UTF-8 without BOM and CRLF with CSV quoting. Human review CSVs neutralize formula prefixes. Provider acceptance of encoding/headers still needs a controlled approved validation.

Keep one register on one work laptop, with one operator at a time. The application uses an exclusive writer lock and atomic journal commits. Register CSVs are regenerated views: never edit them. To correct an unprepared/unattempted record, use `Record-Correction.ps1` with evidence and update the local source to match. To discard an unreserved, never-submitted batch, use `Cancel-PreparedBatch.ps1 -ConfirmNeverSubmitted -Evidence ...`; the old file is invalidated and must never be used. Reserved/uncertain batches cannot be released through that command.

Back up the complete register to your approved protected location:

```powershell
.\Scripts\Backup-Register.ps1 -DestinationFolder 'C:\REPLACE\ApprovedRefundBackups'
```

Journal hashes detect accidental corruption, not a malicious editor who can rewrite all files. Protect access to the register and backups. If preparation is interrupted, run `Recover-Preparation.ps1`; it either recovers the committed batch or invalidates uncommitted files. If integrity checks fail or the register is missing, stop and restore a complete verified backup. Restoring an older backup can omit real attempts, so reconcile all activity since that backup before preparing again. Do not delete locks or reset history as a workaround.

Manually copy only the approved reports/output files to the team's approved Drive folder. Keep the operational register local and back it up separately. Next week, sync, download the full tracker and repeat preparation using **the same Register**. Do not replace it when installing a code update.

## Tests and limitations

```powershell
.\Scripts\Test-Refunds.ps1
.\Scripts\Test-EdgeCases.ps1
.\Scripts\Test-PackagedStartup.ps1
```

Tests use isolated temporary workspaces and fake data, never the real executable, network APIs or PSP portals. See [testing](docs/TESTING.md), [controller questions](docs/CONTROLLER-QUESTIONS.md), [weekly checklist](docs/WEEKLY-CHECKLIST.md), and [workflow](docs/WORKFLOW.md).

The downloadable application has CSV/HTML output, not a native interactive XLSX dashboard. CSVs open/import in Excel without requiring Excel on the machine running the scripts. Portal formats without sufficient evidence remain review-only. Final status checks, live integration validation, corporate script approval and account permissions remain operational responsibilities.

## Repository and release hygiene

This repository is private. `.gitignore` excludes operational Input/Register/Output/Reports, local overrides, binaries, documents, videos and ZIPs. The release builder uses an explicit source allowlist; real transaction artifacts are delivered separately and must never be committed. Fake data is explicitly labelled. There is no telemetry or payment-network client in the application.
