# FG Variance Email Reporter

Automated Python tool that reads the **Physical FG Variance Report** from a SharePoint Excel file, filters out variance exceptions, and sends a formatted HTML email with an Excel attachment — all via the **Microsoft Graph API**.

---

## Overview

This tool compares **physical finished goods (FG) stock** against **SAP quantities (MB52)** and reports all discrepancies to the relevant stakeholders daily. It handles two scenarios:

- 🔴 **Negative variance** — SAP shows excess over physical count
- 🟠 **Positive variance** — Physical stock exceeds SAP

---

## Repository Structure

```
fg-variance-email/
├── fg_variance_email.py          # Main script — reads SharePoint, filters, emails
├── fg_variance_admin_email.py    # Admin variant of the email script
├── requirements.txt              # Python dependencies
└── .github/                      # GitHub Actions workflows
```

---

## Scripts

### `fg_variance_email.py`
Core automation script. Executes the full pipeline:

1. Authenticates to SharePoint and Microsoft Graph using MSAL (two separate app registrations)
2. Opens a workbook session on `Indus Ultimate Stock Recon.xlsx` via Graph API
3. Reads the `Physical FG Variance Report` worksheet's used range
4. Filters rows where column K (`Physical vs MB52_1000`) is a **numeric variance** (excludes "OK" and blank entries)
5. Builds a styled HTML email body with a color-coded variance table
6. Builds a formatted `.xlsx` attachment with borders, fills, and auto-filter
7. Sends the email via `POST /users/{sender}/sendMail` with the Excel as a base64 attachment

### `fg_variance_admin_email.py`
Admin-targeted variant of the same pipeline — likely uses a different recipient list or includes additional detail for internal review.

---

## Data Source

| Property | Value |
|---|---|
| File | `Indus Ultimate Stock Recon.xlsx` |
| Worksheet | `Physical FG Variance Report` |
| Header Row | Row 4 (index 3) |
| Grand Total Variance Cell | K3 (row index 2, col index 10) |
| Variance Filter Column | K — `Physical vs MB52_1000` |

**Columns A–K included in the report:**

| Col | Field |
|-----|-------|
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

---

## Email Output

- **To:** Brijendu Pandey, Aanand SR, Meet Shah, Abhir Bhomavat, Shabbir Kachwala
- **CC:** Production DBX, Onkar Shahane
- **Subject:** `Indus Physical FG vs SAP — Variance Report [DD-Mon-YYYY]`
- **Body:** Summary banner + color-coded variance table (HTML)
- **Attachment:** `Indus_Physical_FG_vs_SAP_Variance_Report_DD-Mon-YYYY.xlsx`

---

## Authentication

Two separate Azure AD app registrations are used:

| Purpose | Config Class |
|---|---|
| Read SharePoint file via Graph | `SharePointAuthConfig` |
| Send email via Graph | `MailAuthConfig` |

Both use the **client credentials flow** (`acquire_token_for_client`) with `https://graph.microsoft.com/.default` scope.

> ⚠️ **Note:** Client secrets are currently hardcoded. Move these to environment variables or Azure Key Vault before any production deployment.

---

## Dependencies

```
msal
requests
pandas
openpyxl
```

Install with:

```bash
pip install -r requirements.txt
```

---

## Running the Script

```bash
python fg_variance_email.py
```

Logs are printed to stdout with timestamps and emoji indicators for each pipeline stage.

---

## Author

**Siddhesh Mandhana**  
SAP Functional Consultant — Indus International FZC, Sharjah
