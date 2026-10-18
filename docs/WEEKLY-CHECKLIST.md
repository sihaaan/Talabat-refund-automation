# Weekly operating checklist

1. Back up the existing register. Confirm you are using the authoritative work-laptop workspace.
2. Run the controller's existing Google Sheets Sync. Download the full Bulk Refunds tab as CSV.
3. Put the CSV in Input. Confirm source units/currency/refund policy are approved; add approved identifier enrichment if available.
4. Run Prepare. Reconcile input rows against the complete Row_Dispositions report.
5. Review PSP/country delayed and incomplete files. Email investigation lists manually to the correct POCs; retain references. Do not request duplicate refunds for previously attempted rows.
6. Review each prepared batch's IDs, country/account, amounts, currency, counts and business approval. Check current refundable balances/statuses.
7. Reserve the exact batch with Start-Submission. Manually submit the approved file once. Never rerun a Checkout batch after an uncertain result.
8. Import Checkout results or record actual Tap/email submission evidence. Missing/ambiguous results remain blocked as Uncertain.
9. Generate reports. Match the tracker-update manifest to a fresh synced export using payment ID + order + country. Resolve unmatched/multiple/status-changed rows.
10. Coordinate a stable live tracker view and manually change only verified Refund Status cells to Refund Processed. Let the tracker stamp the time. Re-export to verify counts.
11. Record PSP final outcomes and confirmation references. Refresh the manager dashboard; show unknown outcomes beside verified success rate.
12. Manually copy approved reports/files to the approved team Drive and back up the register. Keep credentials and executable in their separate unsynced folder.

If preparation was interrupted: Recover-Preparation. If history is missing/corrupt or submission uncertain: investigate and restore verified evidence. Never reset the register to continue.
