# Validation and integration limits

`Scripts/Test-Refunds.ps1` and `Scripts/Test-EdgeCases.ps1` are executable integration tests using only isolated synthetic workspaces. They exercise CSV parsing, routing, amounts, history commits, preparation, reservation, result imports, manual outcomes, reporting and matching through the actual scripts.

`Scripts/Test-PackagedStartup.ps1` builds and extracts the application ZIP, then starts separate PowerShell processes to check the beginner commands with default workspace paths. It checks that review-only preparation accounts for every input row and creates neither submission batches nor attempted-refund tracker entries.

They cover two- and three-decimal currencies, existing minor units, malformed/zero amounts, combined Checkout identifiers, source overrides, missing identifiers, multiple Checkout IDs, duplicate/conflicting rows, historical failed rows with and without timestamps, cross-source duplicate requests, missing manual approvals, delayed review exports, the 30-day boundary, ISO year boundaries, prepared-only/uncertain exclusions from tracker updates, exact matching, exclusive locks, modified files, repeated results, partial results, conflicting result evidence, corrupted/missing history, interrupted preparation recovery, CSV formula safety, quoted multiline CSV and Tap's editable 1,000-row boundary.

Run both scripts on the target laptop after IT approval. Each prints the pass/fail count and location of Test_Results.csv. Test runs deliberately label their result imports as synthetic and never invoke Checkout; their workspaces must not be merged into operational history.

Validated locally on Windows PowerShell 5.1 and PowerShell 7.6.5 during development. Supported runtimes are Windows PowerShell 5.1 or PowerShell 7.5+. Final test counts are recorded in the release notes/test CSVs. Journal timestamps are retained as strings across both engines; repeated external submission evidence is tested for idempotency.

Not tested: real Checkout executable behavior; the executable's internal retries or idempotency; production/sandbox account selection; live PSP CSV/portal acceptance; actual email delivery; Google Sheets Sync/status triggers; Google Drive; customer settlement; corporate execution policy/signing; controller/acquirer permissions. The program has no code paths that call these integrations.

The real supplied tracker can be processed in `-ReviewOnly` mode to validate classification and export counts. That is a data-processing check, not a payment integration test. Real output stays outside the repository and application release.

Limitations: a local tool cannot enforce exactly-once refunds performed outside its workflow. A stale restored backup can miss later attempts. Hashes detect accidental alteration but are not signatures against an administrator rewriting the entire journal. Manual tracker updates require a stable view and exact key checks. Dashboard files are CSV/HTML snapshots, not a native XLSX model; formula definitions and denominators are explicit in the reports.
