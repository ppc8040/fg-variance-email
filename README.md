# Indus Physical FG vs SAP — Variance Email Reporter

Automated GitHub Actions workflow that reads the **Physical FG Variance Report** worksheet from `Indus Ultimate Stock Recon.xlsx` on SharePoint, filters rows where the physical count differs from SAP stock (MB52_1000), and sends a formatted HTML email with an Excel attachment via Microsoft Graph API.

## What It Does

- Authenticates to Microsoft 365 using two separate MSAL app registrations (SharePoint read + Mail send)
- Opens a read-only workbook session on the SharePoint Excel file via Graph API
- Reads columns A–K: Indus Part, Customer, M4_STORE, M4, M5, M6, M16_STORE, Physical Total, SAP MB52_1000, SAP MB52_2000, Variance
- Filters to only rows where the variance column (K) is a numeric value — skips rows marked `OK` or blank
- Builds a colour-coded HTML email: 🔴 red rows = SAP has excess (negative), 🟠 amber rows = physical exceeds SAP (positive)
- Attaches a formatted Excel file with the same variance data (borders, fills, auto-filter, freeze pane)
- Sends to the production team distribution list with CC to operations

## Columns Reported (A–K)

| Column | Field |
|--------|-------|
| A | Indus Part |
| B | Customer |
| C | M4_STORE |
| D | M4 |
| E | M5 |
| F | M6 |
| G | M16_STORE |
| H | Physical Total |
| I | SAP Quantities (MB52_1000) |
| J | SAP Quantities (MB52_2000) |
| K | Physical vs MB52_1000 ← variance filter |

## Tech Stack

- **Python 3.11**
- `msal` — Microsoft identity / OAuth2 token acquisition
- `requests` — Graph API calls
- `pandas` — DataFrame filtering and processing
- `openpyxl` — Formatted Excel attachment generation
- **GitHub Actions** — scheduled execution

## GitHub Actions

The workflow (`.github/workflows/`) runs on a schedule and executes `fg_variance_email.py`. Secrets for `CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID`, and related config must be set in repository secrets.

## Setup

```bash
pip install -r requirements.txt
```

Create a `.env` file (not committed) or set GitHub Secrets:

```
SHAREPOINT_CLIENT_ID=...
SHAREPOINT_CLIENT_SECRET=...
MAIL_CLIENT_ID=...
MAIL_CLIENT_SECRET=...
TENANT_ID=...
```

## Notes

- Row 3 (index 2) in the worksheet holds the Grand Total Variance in column K
- Row 4 (index 3) holds the column headers A4:K4
- Data begins from Row 5 onwards
- The workbook session is always closed in a `finally` block to avoid leaving open sessions

---
*Indus International FZC — Internal Automation Tool*
