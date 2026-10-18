# Information to request from the financial controller

The application already produces investigation CSVs using the available data. Missing fields do not prevent those lists from being generated. They do prevent an incomplete record becoming a submission batch.

| PSP/workflow | Already available | Still needed before a submission-ready file |
|---|---|---|
| All PSPs | Order, country, date, source label, payment reference, captured amount, tracker status | Explicit currency/amount units; confirmation of the approved refund amount after prior refunds/partial captures/wallet adjustments; business approval and final-result evidence |
| Checkout | Valid single `pay_` ID, amount conversion configuration, documented CLI input/result headers | Approved Engineering build and account/environment scope; safe dry-run/retry behavior; current result example; acquirer and country-specific escalation contact; confirmation of refundability and remaining amount |
| Tap PSP | Template screenshot, major-unit amounts, required fields, configurable limit | Actual `chg_` charge ID linked to tracker Payment ID + Order ID + Country; currency; captured/refundable status; valid reason; current template/reason wording; confirmation whether the limit is 100 or 1,000 |
| HyperPay / Switch | Shared workflow and internal recipients confirmed by operator; six-column sample | Whether blank `Transaction ID` is accepted; otherwise the correct Transaction ID and its definition. Confirm the shared template by setting `psps.Switch.templateConfirmed=true` only after approval |
| BenefitPay | `Order_ID,Value,Currency,OrderDate` sample and internal recipients | Confirm currency BHD and exact amount (including third decimal), status and any required provider reference not present in the sample |
| QPay / TESS | Five core transaction fields and `accounts@tesspayments.com` | Replace placeholder/missing payment IDs; confirm currency/amount; identify acquiring-bank recipient for transactions older than 90 days |
| ZainCash | Email recipients and ordinary transaction fields | **Operation ID** and any approved identifier/amount requirements |
| Knet | Portal/manual procedure references and authorized Kuwait Finance route | Verified Knet transaction identifier mapping; current accepted batch sample; required merchant/terminal context; approved sender. Older/failed cases need the authorized bank process. Do not collect PAN/CVV in this tool |
| Tap as acquirer | Documented Tap contacts | Which acquirer/account owns each transaction; approved identifiers and attachment format; relationship to the Checkout payment to avoid issuing two refunds |
| Deema | Dashboard/manual bulk route | Named contact, accepted bulk file fields, required payment identifier and acknowledgement/result format |
| Unknown/combined sources without a valid Checkout ID | Original source text and payment/order data | Approved PSP mapping and interpretation of multiple identifiers. Never split amounts or invent IDs |

Useful optional local CSV columns: `Currency`, `AmountUnit`, `Charge ID`, `Operation ID`, `Reason`. You can supply Charge ID/Operation ID/Reason through `Templates/IdentifierEnrichment.csv` without editing the controller's original tracker. Do not change the controller's sheet structure without permission.

Your tracker still controls its own Sync and automatic processed timestamp. This tool records those facts but never writes them back. Legacy `Refunded`, `Refund Processed`, `Refund Failed` and processed timestamps all block resubmission, without inventing PSP-confirmed outcomes.

Contacts in `Config/contacts.json` remain editable. External addresses from the supplied process are preserved as operational contacts or escalation options. Confirm current ownership before manually addressing a message; no email is sent by the application.
