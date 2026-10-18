# Version 1.0.0

Phase 1 prepares and tracks refunds locally. It never executes the Checkout binary, calls a payment API, sends email, updates Google Sheets, uploads to Drive or reads credentials.

Included: configurable provider/country routing, exact amount conversion, Tap batching from 1 to 1,000 rows, delayed and incomplete-data investigation attachments, editable contacts, permanent journal and attempt tracking, manual-request intake, submission reservation, result import, external submission evidence, final outcomes, corrections, backup/recovery, tracker matching, CSV/HTML reports, fake examples, setup instructions and a weekly checklist.

## Validation

- Windows PowerShell 5.1.26100.9711: 33 core tests and 20 edge-case tests passed.
- PowerShell 7.6.5: the same 53 tests passed.
- Both engines also passed the extracted-ZIP startup test, covering initialization, review-only preparation, reporting and matching without a supplied workspace path (54 checks per engine).
- Test details are in `Examples/TestEvidence`. All tests use synthetic data in isolated workspaces.
- Repeated evidence imports remain idempotent across both engines. PowerShell 7.5+ explicitly preserves JSON timestamp strings; Windows PowerShell 5.1 already does so.
- CSVs stream into temporary files before atomic publication to limit additional memory use on large trackers.

The release ZIP contains source, configuration, templates, documentation and synthetic examples only. Operational inputs/registers/outputs, supplied documents/video, the executable and credentials are excluded.

## Before operational use

Confirm amount units, currency mapping and the refundable amount policy. Obtain the missing provider fields/formats in `CONTROLLER-QUESTIONS.md`. Tap defaults to 100 rows until the current limit is confirmed; its template remains gated. Historical failures remain blocked from automatic resubmission. External review emails are distinct from actual refund requests.

Use the existing manual weekly tracker Sync. Mark only verified matched rows as Refund Processed after recording submission evidence; the tracker supplies its own timestamp.

Live PSP acceptance, the Engineering executable, account/environment selection, corporate script policy, Google Sheet behavior and settlement were not integration-tested. See `TESTING.md` for the full limits.
