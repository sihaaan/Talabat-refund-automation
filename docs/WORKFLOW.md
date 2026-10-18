# Phase 1 workflow

```mermaid
flowchart TD
    A[Human: Sync Google Sheet and download full Bulk Refunds CSV] --> B[Automatic: validate and preserve source snapshot]
    M[Human: manual email request with approval] --> B
    B --> C[Automatic: read permanent history and route PSP]
    C --> X[Duplicate / unknown PSP / invalid amount: review file]
    C --> D{30 or more days old?}
    D -->|Yes| E[Automatic: PSP + country delayed CSV with prior-attempt warnings]
    E --> F[Human: email POC for investigation]
    F --> G[Record communication; no refund success inferred]
    D -->|No| H{Required data and assumptions confirmed?}
    H -->|No| X
    H -->|Yes| I[Automatic: prepare Checkout / Tap / email batch]
    I --> J[Human: review and obtain approval]
    J --> K[Automatic: reserve exact batch and block duplicate preparation]
    K --> L[Human: execute Checkout once / upload Tap / send refund request]
    L --> N[Import results or record actual submission evidence]
    L --> U[Missing result: Uncertain and blocked]
    N --> O[Automatic: generate exact tracker-update manifest]
    O --> P[Human: match fresh tracker and update only verified status cells]
    P --> Q[Tracker automatically stamps timestamp]
    N --> R{Final PSP confirmation?}
    R -->|Pending| S[Unknown final outcome; no retry]
    R -->|Confirmed failure| T[Record failed outcome and escalate; no retry]
    R -->|Confirmed success| V[Record verified success]
    S --> W[Automatic: manager report with explicit denominators]
    T --> W
    V --> W
```

Prepared, Reserved, Submitted, RequestFailed and Uncertain are submission states. Final status is separately Unknown, ConfirmedSuccessful, ConfirmedFailed or Voided. A prepared file, a reservation and an investigation email never enter the tracker-update manifest.
