# FG Variance Email Reporter

Automated Python tool that performs a **3-way reconciliation** of Indus warehouse scanner stock, SAP plant stock (MB52), and Open Order Report (OOR) customer data — then emails a formatted HTML variance report with an Excel attachment via the **Microsoft Graph API**.

Two scripts exist in this repo with different scopes:

- **`fg_variance_admin_email.py`** — the primary script, run by GitHub Actions daily. Full 3-source recon: Postgres (scanner) + SAP OData + SharePoint (OOR).
- **`fg_variance_email.py`** — reads the pre-computed `Physical FG Variance Report` worksheet directly from SharePoint and emails it. No SAP/Postgres integration.

---

## Architecture Overview

```
┌─────────────────────┐    ┌────────────────────────┐    ┌────────────────────────────────┐
│  PostgreSQL (Heroku) │    │  SAP S/4HANA OData API │    │  SharePoint (Graph API)        │
│  movements table     │    │  API_MATERIAL_STOCK_SRV│    │  Indus Open Order Report.xlsx  │
│  STOCK_IN / OUT net  │    │  Plant 1000 + 2000     │    │  Col H (part) → Col R (code)   │
└─────────┬───────────┘    └──────────┬─────────────┘    └────────────────┬───────────────┘
          │                           │                                    │
          └──────────────┬────────────┘                                   │
                         ▼                                                 │
               build_recon_rows()  ◄──────────────────────────────────────┘
               (full outer join on indus_part)
                         │
                         ▼
              build_excel()  +  build_html_body()
                         │
                         ▼
              Microsoft Graph API → Send Mail
```

---

## Scripts

### `fg_variance_admin_email.py` ← **Production / GitHub Actions**

7-step pipeline:

| Step | Function | What it does |
|------|----------|--------------|
| 1 | `fetch_scanner_data()` | Queries Postgres `movements` table — computes `net_qty = STOCK_IN - STOCK_OUT` per `indus_part` |
| 2 | `fetch_sap_stock()` | Paginates SAP OData (`A_MatlStkInAcctMod`) for Plants 1000 & 2000; filters to Indus materials (starts with A/C/I) |
| 3 | `fetch_oor_customer_codes()` | Reads `Indus Open Order Report New.xlsx` from SharePoint; maps col H → col R (part → customer code) |
| 4 | `build_recon_rows()` | Full outer join of all three sources on `indus_part`; computes `variance_vs_1000 = scanner_qty - sap_qty_1000` |
| 5 | `build_excel()` | Styled `.xlsx` with dark navy header, red row fills for negative variances |
| 6 | `build_html_body()` | HTML email body — summary banner + discrepant-only table (dark theme) |
| 7 | `send_email()` | Posts to `POST /users/{sender}/sendMail` with Excel as base64 attachment |

**Email output:**
- **Subject:** `Indus Warehouse-SAP FG Variance [DD-Mon-YYYY]`
- **Attachment:** `Indus Warehouse-SAP FG Variance DD-Mon-YYYY HH.MM GST.xlsx`
- **To:** Abhir Bhomavat, Aanand SR, Brijendu Pandey, Shabbir Kachwala
- **CC:** Meet Shah, prodndbx3, ppc2, Sachin Nanekar, prodndbx11
- **Timezone:** Dubai (UTC+4) for all timestamps

**Excel columns:**

| Column | Description |
|--------|-------------|
| `indus_part` | Material code |
| `customer` | OOR customer code (from SharePoint) |
| `customer_part` | Customer's part number (from scanner DB) |
| `scanner_qty` | Net physical qty (STOCK_IN − STOCK_OUT) |
| `sap_qty_1000` | SAP Plant 1000 stock |
| `sap_qty_2000` | SAP Plant 2000 stock |
| `variance_vs_1000` | `scanner_qty − sap_qty_1000` |

---

### `fg_variance_email.py` ← **SharePoint-only variant**

Reads the pre-computed `Physical FG Variance Report` worksheet from `Indus Ultimate Stock Recon.xlsx` on SharePoint, filters rows where column K (`Physical vs MB52_1000`) has a **numeric variance** (non-"OK", non-blank), and emails a full 11-column report.

**Sheet layout:**

| Row | Content |
|-----|---------|
| Row 3 (index 2) | Grand Total Variance in column K |
| Row 4 (index 3) | Headers A4:K4 |
| Row 5+ | Data rows |

**Columns A–K:**

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

**Email output:**
- **Subject:** `Indus Physical FG vs SAP — Variance Report [DD-Mon-YYYY]`
- **To:** Brijendu Pandey, Aanand SR, Meet Shah, Abhir Bhomavat, Shabbir Kachwala
- **CC:** prodndbx3, Onkar Shahane

---

## GitHub Actions — Scheduled Trigger

**File:** `.github/workflows/daily_email.yml`

Runs `fg_variance_admin_email.py` on a cron schedule:

```yaml
on:
  schedule:
    - cron: '0 6,10,14 * * 1-6'   # 10:00, 14:00, 18:00 Dubai time (Mon–Sat)
  workflow_dispatch:                # Manual trigger from Actions tab
```

Runs Monday–Saturday at 10:00 AM, 2:00 PM, and 6:00 PM Dubai time (UTC+4).

---

## Authentication

Two Azure AD app registrations, both using the **client credentials flow**:

| Purpose | Variables |
|---------|-----------|
| SharePoint read (Graph) | `SP_CLIENT_ID`, `SP_CLIENT_SECRET`, `SP_TENANT_ID` |
| Send email (Graph) | `CLIENT_ID`, `CLIENT_SECRET`, `AUTHORITY` |

> ⚠️ **Credentials are currently hardcoded.** Move all secrets to GitHub Actions secrets (`${{ secrets.* }}`) or Azure Key Vault before any wider deployment. The Postgres `DATABASE_URL` and SAP password are also hardcoded.

---

## Dependencies

```
msal
requests
pandas
openpyxl
psycopg2-binary   # required by fg_variance_admin_email.py (not in requirements.txt yet)
```

> **Note:** `requirements.txt` currently only lists `msal requests pandas openpyxl`. Add `psycopg2-binary` if running `fg_variance_admin_email.py` locally.

Install:

```bash
pip install -r requirements.txt
pip install psycopg2-binary   # for admin script
```

---

## Running Locally

```bash
# SharePoint-only variant
python fg_variance_email.py

# Full 3-way recon (admin script)
python fg_variance_admin_email.py
```

Both scripts log to stdout with timestamps. SAP pagination can take **2–5 minutes** — expected behavior.

---

## Author

**Siddhesh Mandhana**  
SAP S/4HANA Consultant & Process Automation Lead — Indus International FZC, Sharjah
