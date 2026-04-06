"""
INDUS PHYSICAL FG vs SAP — VARIANCE EMAIL REPORTER
====================================================
Reads "Physical FG Variance Report" worksheet from the SharePoint file
"Indus Ultimate Stock Recon.xlsx", filters rows where column J
(Final Variance Physical vs SAP) is non-zero, sends a formatted HTML
email via Microsoft Graph API, and attaches the same variance report
as a downloadable Excel file.

Columns included in report:
  A – Indus Part
  B – Customer
  E – Physical Warehouse Total
  H – Production FG Total
  I – SAP Quantities (1000+2000)
  J – Final Variance (Physical vs SAP)

The Excel attachment is formatted with borders, fills, alignment, and
a worksheet name derived from the email subject.

Author: Siddhesh Mandhana

Secrets are read from environment variables (set as GitHub Actions secrets):
  SP_CLIENT_ID, SP_CLIENT_SECRET, SP_TENANT_ID
  MAIL_CLIENT_ID, MAIL_CLIENT_SECRET
"""

import base64
import os
import re
import requests
import pandas as pd
import msal
import logging
import sys
import traceback
from pathlib import Path
from datetime import datetime
from urllib.parse import quote
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# ============================================================
# SYSTEM SETUP
# ============================================================
if sys.platform == "win32":
    os.environ["PYTHONIOENCODING"] = "utf-8"

# ============================================================
# LOGGING
# ============================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# ============================================================
# CONFIGURATION  (secrets from environment variables)
# ============================================================
TENANT_ID = os.environ["SP_TENANT_ID"]

class SharePointAuthConfig:
    CLIENT_ID     = os.environ["SP_CLIENT_ID"]
    CLIENT_SECRET = os.environ["SP_CLIENT_SECRET"]
    AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"

class MailAuthConfig:
    CLIENT_ID     = os.environ["MAIL_CLIENT_ID"]
    CLIENT_SECRET = os.environ["MAIL_CLIENT_SECRET"]
    AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"

class SharePointConfig:
    SITE_ID              = "201c767f-23cc-4eff-b76c-47a275706531,094b21f9-6ed4-47c3-a751-3b217a18b98e"
    FILE_NAME            = "Indus Ultimate Stock Recon.xlsx"
    WORKSHEET_NAME       = "Physical FG Variance Report"
    HEADER_ROW_INDEX     = 3
    TOTAL_VARIANCE_ROW_INDEX = 1
    TOTAL_VARIANCE_COL_INDEX = 9

class EmailConfig:
    SENDER_EMAIL         = "siddhesh.mandhana@indusair.ae"
    TO_EMAILS            = ["onkar.shahane@indusair.ae]
    CC_EMAILS            = []
    SUBJECT              = f"Indus Physical FG vs SAP \u2014 Variance Report [{datetime.now().strftime('%d-%b-%Y')}]"
    ATTACHMENT_FILE_NAME = f"Indus_Physical_FG_vs_SAP_Variance_Report_{datetime.now().strftime('%d-%b-%Y')}.xlsx"

# ============================================================
# AUTHENTICATION
# ============================================================
def get_sharepoint_token():
    logger.info("\U0001f510 Acquiring SharePoint access token...")
    try:
        app = msal.ConfidentialClientApplication(
            SharePointAuthConfig.CLIENT_ID,
            authority=SharePointAuthConfig.AUTHORITY,
            client_credential=SharePointAuthConfig.CLIENT_SECRET
        )
        result = app.acquire_token_silent(
            scopes=["https://graph.microsoft.com/.default"],
            account=None
        )
        if not result:
            result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" in result:
            logger.info("\u2705 SharePoint token acquired")
            return result["access_token"]
        raise Exception(result.get("error_description", result.get("error")))
    except Exception as e:
        logger.error(f"\u274c SharePoint token error: {e}")
        traceback.print_exc()
        raise


def get_mail_token():
    logger.info("\U0001f510 Acquiring Mail send token...")
    try:
        app = msal.ConfidentialClientApplication(
            MailAuthConfig.CLIENT_ID,
            authority=MailAuthConfig.AUTHORITY,
            client_credential=MailAuthConfig.CLIENT_SECRET
        )
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" in result:
            logger.info("\u2705 Mail token acquired")
            return result["access_token"]
        raise Exception(result.get("error_description", result.get("error")))
    except Exception as e:
        logger.error(f"\u274c Mail token error: {e}")
        traceback.print_exc()
        raise

# ============================================================
# SHAREPOINT READ
# ============================================================
def read_worksheet(sp_token):
    logger.info(f"\U0001f4c2 Reading worksheet: '{SharePointConfig.WORKSHEET_NAME}'")
    session_obj = requests.Session()
    api_headers = {
        "Authorization": f"Bearer {sp_token}",
        "Content-Type": "application/json"
    }
    site_id   = SharePointConfig.SITE_ID
    file_name = quote(SharePointConfig.FILE_NAME)
    ws_name   = quote(SharePointConfig.WORKSHEET_NAME)
    session_id = None
    try:
        session_url = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}"
            f"/drive/root:/{file_name}:/workbook/createSession"
        )
        session_resp = session_obj.post(
            session_url,
            headers=api_headers,
            json={"persistChanges": False},
            timeout=30
        )
        if session_resp.status_code not in (200, 201):
            raise Exception(f"Failed to create workbook session: {session_resp.status_code} \u2014 {session_resp.text[:300]}")
        session_id = session_resp.json().get("id")
        api_headers["workbook-session-id"] = session_id
        logger.info(f"\u2705 Workbook session created: {session_id[:30]}...")

        range_url = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}"
            f"/drive/root:/{file_name}:/workbook/worksheets('{ws_name}')/usedRange"
        )
        range_resp = session_obj.get(range_url, headers=api_headers, timeout=120)
        if range_resp.status_code != 200:
            raise Exception(f"Failed to read usedRange: {range_resp.status_code} \u2014 {range_resp.text[:300]}")
        values = range_resp.json().get("values", [])
        logger.info(f"\u2705 Raw rows returned from API: {len(values)}")
    except Exception as e:
        logger.error(f"\u274c Worksheet read error: {e}")
        traceback.print_exc()
        raise
    finally:
        if session_id:
            try:
                close_url = (
                    f"https://graph.microsoft.com/v1.0/sites/{site_id}"
                    f"/drive/root:/{file_name}:/workbook/closeSession"
                )
                session_obj.post(close_url, headers=api_headers, timeout=15)
                logger.info("\u2705 Workbook session closed")
            except Exception as ce:
                logger.warning(f"\u26a0\ufe0f Could not close session: {ce}")

    total_variance = 0
    try:
        total_variance_raw = values[SharePointConfig.TOTAL_VARIANCE_ROW_INDEX][SharePointConfig.TOTAL_VARIANCE_COL_INDEX]
        total_variance = float(total_variance_raw) if total_variance_raw not in ("", None) else 0
        logger.info(f"\U0001f4ca Total variance (J2): {total_variance}")
    except (IndexError, ValueError, TypeError) as e:
        logger.warning(f"\u26a0\ufe0f Could not read total variance cell: {e}. Defaulting to 0.")

    header_idx = SharePointConfig.HEADER_ROW_INDEX
    if len(values) <= header_idx:
        raise Exception(f"Not enough rows returned. Expected headers at index {header_idx}, but only {len(values)} row(s) found.")
    headers   = values[header_idx]
    data_rows = values[header_idx + 1:]
    logger.info(f"\U0001f4cb Headers detected: {headers}")
    logger.info(f"\U0001f4cb Data rows: {len(data_rows)}")
    if not data_rows:
        raise Exception("No data rows found after header row.")
    df = pd.DataFrame(data_rows, columns=headers)
    logger.info(f"\u2705 DataFrame built: {df.shape[0]} rows \u00d7 {df.shape[1]} columns")
    return df, total_variance

# ============================================================
# FILTER DATA
# ============================================================
def filter_variance_rows(df):
    logger.info("\U0001f50d Selecting columns and filtering variance rows...")
    total_cols = df.shape[1]
    logger.info(f"  Total columns available: {total_cols}")
    if total_cols < 10:
        raise Exception(
            f"Expected at least 10 columns (A\u2013J), but got {total_cols}. "
            f"Check HEADER_ROW_INDEX (currently {SharePointConfig.HEADER_ROW_INDEX})."
        )
    df_sel = df.iloc[:, [0, 1, 4, 7, 8, 9]].copy()
    df_sel.columns = [
        "Indus Part",
        "Customer",
        "Physical Warehouse Total",
        "Production FG Total",
        "SAP Quantities (1000+2000)",
        "Final Variance (Physical vs SAP)"
    ]
    before = len(df_sel)
    df_sel = df_sel[
        ~(
            df_sel["Indus Part"].astype(str).str.strip().isin(["", "nan"])
            & df_sel["Customer"].astype(str).str.strip().isin(["", "nan"])
        )
    ].copy()
    logger.info(f"  Dropped {before - len(df_sel)} blank rows. Remaining: {len(df_sel)}")

    numeric_cols = [
        "Physical Warehouse Total",
        "Production FG Total",
        "SAP Quantities (1000+2000)",
        "Final Variance (Physical vs SAP)"
    ]
    for col in numeric_cols:
        df_sel[col] = pd.to_numeric(df_sel[col], errors="coerce")
    df_sel["Final Variance (Physical vs SAP)"] = df_sel["Final Variance (Physical vs SAP)"].fillna(0)
    result = df_sel[df_sel["Final Variance (Physical vs SAP)"] != 0].copy()
    for col in ["Physical Warehouse Total", "Production FG Total", "SAP Quantities (1000+2000)"]:
        result[col] = result[col].fillna(0)
    logger.info(f"\u2705 Variance rows found: {len(result)} out of {len(df_sel)} total")
    if not result.empty:
        logger.info(f"  Sample:\n{result.head(5).to_string()}")
    return result

# ============================================================
# EMAIL HTML
# ============================================================
def build_html_body(variance_df, total_variance):
    run_date = datetime.now().strftime("%d-%b-%Y %H:%M")

    def variance_row_colour(val):
        try:
            v = float(val)
        except (ValueError, TypeError):
            return "#FFFFFF"
        if v < 0: return "#FFD7D7"
        if v > 0: return "#FFF3CD"
        return "#FFFFFF"

    def variance_cell_colour(val):
        try:
            v = float(val)
        except (ValueError, TypeError):
            return "#FFFFFF"
        if v < 0: return "#FF6B6B"
        if v > 0: return "#FFA500"
        return "#FFFFFF"

    def fmt(v):
        try:
            fv = float(v)
            return f"{int(fv):,}" if fv == int(fv) else f"{fv:,.2f}"
        except (ValueError, TypeError):
            return str(v)

    def fmt_var(v):
        try:
            fv = float(v)
            return f"{int(fv):+,}" if fv == int(fv) else f"{fv:+,.2f}"
        except (ValueError, TypeError):
            return str(v)

    total_color = "#C0392B" if total_variance < 0 else ("#E67E22" if total_variance > 0 else "#27AE60")
    total_label = fmt_var(total_variance)

    summary_banner = f"""
    <table style="width:100%;border-collapse:collapse;margin-bottom:18px;">
      <tr style="background:#1F3864;">
        <td style="padding:10px 16px;color:#FFFFFF;font-size:15px;font-weight:bold;">
          \U0001f4ca Physical FG vs SAP \u2014 Variance Summary
        </td>
      </tr>
      <tr style="background:#F0F4FA;">
        <td style="padding:10px 16px;font-size:13px;color:#333;">
          <b>Report Date:</b> {run_date} &nbsp;|&nbsp;
          <b>Variance Rows:</b> {len(variance_df)} &nbsp;|&nbsp;
          <b>Grand Total Variance:</b>
          <span style="color:{total_color};font-weight:bold;">{total_label}</span>
          &nbsp;<i>(Negative = excess stock in SAP)</i>
        </td>
      </tr>
    </table>
    """

    if variance_df.empty:
        body_content = """
        <p style="color:#27AE60;font-weight:bold;font-size:14px;">
          \u2705 No variance exceptions found. All records are currently balanced.
        </p>
        """
    else:
        col_cfg = [
            ("Indus Part",                       "130px", "#1F3864", "#FFFFFF"),
            ("Customer",                          "110px", "#1F3864", "#FFFFFF"),
            ("Physical Warehouse Total",          "150px", "#1F3864", "#FFFFFF"),
            ("Production FG Total",               "145px", "#1F3864", "#FFFFFF"),
            ("SAP Quantities (1000+2000)",         "155px", "#1F3864", "#FFFFFF"),
            ("Final Variance (Physical vs SAP)",  "170px", "#C0392B", "#FFFFFF"),
        ]
        header_row = "".join(
            f'<th style="background:{bg};color:{fg};padding:8px 10px;"
               f'width:{w};border:1px solid #CCC;white-space:nowrap;">{col}</th>'
            for col, w, bg, fg in col_cfg
        )
        data_rows_html = ""
        for _, row in variance_df.iterrows():
            var_val = row["Final Variance (Physical vs SAP)"]
            row_bg  = variance_row_colour(var_val)
            cell_bg = variance_cell_colour(var_val)
            cells   = "".join(
                f'<td style="padding:7px 10px;border:1px solid #DDD;background:{row_bg};text-align:right;">'
                f'{fmt(row[c])}</td>'
                for c in ["Indus Part", "Customer", "Physical Warehouse Total",
                           "Production FG Total", "SAP Quantities (1000+2000)"]
            )
            cells += (
                f'<td style="padding:7px 10px;border:1px solid #DDD;background:{cell_bg};'
                f'color:#000;font-weight:bold;text-align:right;">{fmt_var(var_val)}</td>'
            )
            data_rows_html += f"<tr>{cells}</tr>\n"

        body_content = f"""
        <table style="border-collapse:collapse;width:100%;font-size:13px;">
          <thead><tr>{header_row}</tr></thead>
          <tbody>{data_rows_html}</tbody>
        </table>
        <p style="font-size:12px;color:#555;margin-top:8px;">
          \U0001f534 Red rows = SAP has excess (negative variance) &nbsp;|
          &nbsp;\U0001f7e0 Amber rows = Physical stock exceeds SAP (positive variance)
        </p>
        """

    html = f"""<!DOCTYPE html>
<html><body style="font-family:Calibri,Arial,sans-serif;font-size:13px;color:#333;margin:20px;">
  <h2 style="color:#1F3864;margin-bottom:6px;">Indus Physical FG vs SAP \u2014 Variance Report</h2>
  <hr style="border:1px solid #1F3864;margin-bottom:16px;">
  {summary_banner}
  {body_content}
  <p style="font-size:11px;color:#999;margin-top:20px;">
    This is an automated report generated by the FG Variance Email Reporter.
  </p>
</body></html>"""
    return html

# ============================================================
# BUILD EXCEL ATTACHMENT
# ============================================================
def build_excel_attachment(variance_df, total_variance):
    logger.info("\U0001f4ce Building Excel attachment...")
    wb = Workbook()
    ws = wb.active
    subject_safe = re.sub(r'[\\/*?:\[\]]', '_', EmailConfig.SUBJECT)[:31]
    ws.title = subject_safe

    thin  = Side(style='thin',   color='000000')
    thick = Side(style='medium', color='000000')
    full_border  = Border(left=thin,  right=thin,  top=thin,  bottom=thin)
    thick_border = Border(left=thick, right=thick, top=thick, bottom=thick)

    header_fill    = PatternFill("solid", fgColor="1F3864")
    var_fill       = PatternFill("solid", fgColor="C0392B")
    red_row_fill   = PatternFill("solid", fgColor="FFD7D7")
    amber_row_fill = PatternFill("solid", fgColor="FFF3CD")
    red_cell_fill  = PatternFill("solid", fgColor="FF6B6B")
    amber_cell_fill= PatternFill("solid", fgColor="FFA500")

    columns = [
        ("Indus Part",                      16),
        ("Customer",                         16),
        ("Physical Warehouse Total",         22),
        ("Production FG Total",              20),
        ("SAP Quantities (1000+2000)",        24),
        ("Final Variance (Physical vs SAP)", 26),
    ]
    for col_idx, (col_name, width) in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        fill = var_fill if "Variance" in col_name else header_fill
        cell.fill   = fill
        cell.font   = Font(bold=True, color="FFFFFF", size=11)
        cell.border = full_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 30

    for row_idx, (_, row) in enumerate(variance_df.iterrows(), start=2):
        var_val = row["Final Variance (Physical vs SAP)"]
        try:
            v = float(var_val)
        except (ValueError, TypeError):
            v = 0
        row_fill  = red_row_fill   if v < 0 else (amber_row_fill  if v > 0 else None)
        cell_fill = red_cell_fill  if v < 0 else (amber_cell_fill if v > 0 else None)

        row_values = [
            row["Indus Part"],
            row["Customer"],
            row["Physical Warehouse Total"],
            row["Production FG Total"],
            row["SAP Quantities (1000+2000)"],
            var_val
        ]
        for col_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            if col_idx == 6 and cell_fill:
                cell.fill = cell_fill
                cell.font = Font(bold=True, color="000000")
            elif row_fill:
                cell.fill = row_fill
            cell.border = full_border
            if col_idx > 2:
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="right")

    # Summary row
    summary_row = variance_df.shape[0] + 3
    ws.cell(row=summary_row, column=5, value="Grand Total Variance").font = Font(bold=True)
    tv_cell = ws.cell(row=summary_row, column=6, value=total_variance)
    tv_cell.font   = Font(bold=True, color="C0392B" if total_variance < 0 else ("E67E22" if total_variance > 0 else "27AE60"))
    tv_cell.border = thick_border
    tv_cell.number_format = '+#,##0;-#,##0;0'

    from io import BytesIO
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    logger.info("\u2705 Excel attachment built")
    return buf.read()

# ============================================================
# SEND EMAIL
# ============================================================
def send_email(mail_token, html_body, excel_bytes):
    logger.info("\U0001f4e7 Sending email via Microsoft Graph...")
    attachment_b64 = base64.b64encode(excel_bytes).decode("utf-8")
    payload = {
        "message": {
            "subject": EmailConfig.SUBJECT,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients":  [{"emailAddress": {"address": e}} for e in EmailConfig.TO_EMAILS],
            "ccRecipients":  [{"emailAddress": {"address": e}} for e in EmailConfig.CC_EMAILS],
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": EmailConfig.ATTACHMENT_FILE_NAME,
                    "contentBytes": attachment_b64,
                    "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                }
            ]
        },
        "saveToSentItems": True
    }
    url  = f"https://graph.microsoft.com/v1.0/users/{EmailConfig.SENDER_EMAIL}/sendMail"
    resp = requests.post(
        url,
        headers={"Authorization": f"Bearer {mail_token}", "Content-Type": "application/json"},
        json=payload,
        timeout=60
    )
    if resp.status_code == 202:
        logger.info("\u2705 Email sent successfully (HTTP 202)")
    else:
        raise Exception(f"Email send failed: {resp.status_code} \u2014 {resp.text[:400]}")

# ============================================================
# MAIN
# ============================================================
def main():
    logger.info("=" * 60)
    logger.info("FG VARIANCE EMAIL REPORTER — START")
    logger.info("=" * 60)
    try:
        sp_token   = get_sharepoint_token()
        mail_token = get_mail_token()
        df, total_variance = read_worksheet(sp_token)
        variance_df        = filter_variance_rows(df)
        html_body          = build_html_body(variance_df, total_variance)
        excel_bytes        = build_excel_attachment(variance_df, total_variance)
        send_email(mail_token, html_body, excel_bytes)
        logger.info("=" * 60)
        logger.info("FG VARIANCE EMAIL REPORTER — COMPLETE")
        logger.info("=" * 60)
    except Exception as e:
        logger.error(f"\u274c Fatal error: {e}")
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
