"""
INDUS PHYSICAL FG vs SAP — VARIANCE EMAIL REPORTER
====================================================

Reads "Physical FG Variance Report" worksheet from the SharePoint file
"Indus Ultimate Stock Recon.xlsx", filters rows where column K
(Final Variance Physical vs SAP) is non-zero, sends a formatted HTML
email via Microsoft Graph API, and attaches the same variance report
as a downloadable Excel file.

Columns included in report:
A – Indus Part
B – Customer
E – Physical Warehouse Total
H – Production FG Total
I – SAP Quantities (MB52_1000)
J – SAP Quantities (MB52_2000)
K – Final Variance (Physical vs SAP)

The Excel attachment is formatted with borders, fills, alignment,
and a worksheet name derived from the email subject.
"""

import base64
import requests
import pandas as pd
import msal
import logging
import sys
import traceback
from datetime import datetime
from urllib.parse import quote

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# ============================================================
# SYSTEM SETUP
# ============================================================
if sys.platform == "win32":
    import os
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
# CONFIGURATION
# ============================================================
class SharePointAuthConfig:
    CLIENT_ID = "a2239be7-12cd-442d-983f-ea7e316ef767"
    CLIENT_SECRET = "sho8Q~sanlYEX3SKXjYJb0NSXL.gPRpdi63-tcmp"
    AUTHORITY = "https://login.microsoftonline.com/ed97e9bb-e119-4bd4-ab00-307d64bdf908"


class MailAuthConfig:
    CLIENT_ID = "785b5fc9-1624-458d-90cb-a140dff579c7"
    CLIENT_SECRET = "Zcw8Q~YSEqXPypF4Gn_mNVpp2MVPRMM.95hjVdbP"
    AUTHORITY = "https://login.microsoftonline.com/ed97e9bb-e119-4bd4-ab00-307d64bdf908"


class SharePointConfig:
    SITE_ID = "201c767f-23cc-4eff-b76c-47a275706531,094b21f9-6ed4-47c3-a751-3b217a18b98e"
    FILE_NAME = "Indus Ultimate Stock Recon.xlsx"
    WORKSHEET_NAME = "Physical FG Variance Report"
    HEADER_ROW_INDEX = 3
    TOTAL_VARIANCE_ROW_INDEX = 1
    TOTAL_VARIANCE_COL_INDEX = 10  # K column (0-based index)


class EmailConfig:
    SENDER_EMAIL = "releases@indusair.ae"
    TO_EMAILS = ["siddhesh.mandhana@indusair.ae"]
    CC_EMAILS = []
    SUBJECT = f"Indus Physical FG vs SAP — Variance Report [{datetime.now().strftime('%d-%b-%Y')}]"
    ATTACHMENT_FILE_NAME = f"Indus_Physical_FG_vs_SAP_Variance_Report_{datetime.now().strftime('%d-%b-%Y')}.xlsx"

# ============================================================
# AUTHENTICATION
# ============================================================
def get_sharepoint_token():
    logger.info("🔐 Acquiring SharePoint access token...")
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
            result = app.acquire_token_for_client(
                scopes=["https://graph.microsoft.com/.default"]
            )

        if "access_token" in result:
            logger.info("✅ SharePoint token acquired")
            return result["access_token"]

        raise Exception(result.get("error_description", result.get("error")))
    except Exception as e:
        logger.error(f"❌ SharePoint token error: {e}")
        traceback.print_exc()
        raise


def get_mail_token():
    logger.info("🔐 Acquiring Mail send token...")
    try:
        app = msal.ConfidentialClientApplication(
            MailAuthConfig.CLIENT_ID,
            authority=MailAuthConfig.AUTHORITY,
            client_credential=MailAuthConfig.CLIENT_SECRET
        )
        result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )

        if "access_token" in result:
            logger.info("✅ Mail token acquired")
            return result["access_token"]

        raise Exception(result.get("error_description", result.get("error")))
    except Exception as e:
        logger.error(f"❌ Mail token error: {e}")
        traceback.print_exc()
        raise

# ============================================================
# SHAREPOINT READ
# ============================================================
def read_worksheet(sp_token):
    logger.info(f"📂 Reading worksheet: '{SharePointConfig.WORKSHEET_NAME}'")

    session_obj = requests.Session()
    api_headers = {
        "Authorization": f"Bearer {sp_token}",
        "Content-Type": "application/json"
    }

    site_id = SharePointConfig.SITE_ID
    file_name = quote(SharePointConfig.FILE_NAME)
    ws_name = quote(SharePointConfig.WORKSHEET_NAME)
    session_id = None
    values = []

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
            raise Exception(
                f"Failed to create workbook session: {session_resp.status_code} — {session_resp.text[:300]}"
            )

        session_id = session_resp.json().get("id")
        api_headers["workbook-session-id"] = session_id
        logger.info(f"✅ Workbook session created: {session_id[:30]}...")

        range_url = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}"
            f"/drive/root:/{file_name}:/workbook/worksheets('{ws_name}')/usedRange"
        )
        range_resp = session_obj.get(range_url, headers=api_headers, timeout=120)
        if range_resp.status_code != 200:
            raise Exception(
                f"Failed to read usedRange: {range_resp.status_code} — {range_resp.text[:300]}"
            )

        values = range_resp.json().get("values", [])
        logger.info(f"✅ Raw rows returned from API: {len(values)}")

    except Exception as e:
        logger.error(f"❌ Worksheet read error: {e}")
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
                logger.info("✅ Workbook session closed")
            except Exception as ce:
                logger.warning(f"⚠️ Could not close session: {ce}")

    total_variance = 0
    try:
        total_variance_raw = values[SharePointConfig.TOTAL_VARIANCE_ROW_INDEX][SharePointConfig.TOTAL_VARIANCE_COL_INDEX]
        total_variance = float(total_variance_raw) if total_variance_raw not in ("", None) else 0
        logger.info(f"📊 Total variance (K2): {total_variance}")
    except (IndexError, ValueError, TypeError) as e:
        logger.warning(f"⚠️ Could not read total variance cell: {e}. Defaulting to 0.")

    header_idx = SharePointConfig.HEADER_ROW_INDEX
    if len(values) <= header_idx:
        raise Exception(
            f"Not enough rows returned. Expected headers at index {header_idx}, but only {len(values)} row(s) found."
        )

    headers = values[header_idx]
    data_rows = values[header_idx + 1:]

    logger.info(f"📋 Headers detected: {headers}")
    logger.info(f"📋 Data rows: {len(data_rows)}")

    if not data_rows:
        raise Exception("No data rows found after header row.")

    df = pd.DataFrame(data_rows, columns=headers)
    logger.info(f"✅ DataFrame built: {df.shape[0]} rows × {df.shape[1]} columns")
    return df, total_variance

# ============================================================
# FILTER DATA
# ============================================================
def filter_variance_rows(df):
    logger.info("🔍 Selecting columns and filtering variance rows...")

    total_cols = df.shape[1]
    logger.info(f" Total columns available: {total_cols}")

    if total_cols < 11:
        raise Exception(
            f"Expected at least 11 columns (A–K), but got {total_cols}. "
            f"Check HEADER_ROW_INDEX (currently {SharePointConfig.HEADER_ROW_INDEX})."
        )

    # A, B, E, H, I, J, K
    df_sel = df.iloc[:, [0, 1, 4, 7, 8, 9, 10]].copy()
    df_sel.columns = [
        "Indus Part",
        "Customer",
        "Physical Warehouse Total",
        "Production FG Total",
        "SAP Quantities (MB52_1000)",
        "SAP Quantities (MB52_2000)",
        "Final Variance (Physical vs SAP)"
    ]

    before = len(df_sel)
    df_sel = df_sel[
        ~(
            df_sel["Indus Part"].astype(str).str.strip().isin(["", "nan"])
            & df_sel["Customer"].astype(str).str.strip().isin(["", "nan"])
        )
    ].copy()
    logger.info(f" Dropped {before - len(df_sel)} blank rows. Remaining: {len(df_sel)}")

    numeric_cols = [
        "Physical Warehouse Total",
        "Production FG Total",
        "SAP Quantities (MB52_1000)",
        "SAP Quantities (MB52_2000)",
        "Final Variance (Physical vs SAP)"
    ]
    for col in numeric_cols:
        df_sel[col] = pd.to_numeric(df_sel[col], errors="coerce")

    df_sel["Final Variance (Physical vs SAP)"] = df_sel["Final Variance (Physical vs SAP)"].fillna(0)

    result = df_sel[df_sel["Final Variance (Physical vs SAP)"] != 0].copy()

    for col in [
        "Physical Warehouse Total",
        "Production FG Total",
        "SAP Quantities (MB52_1000)",
        "SAP Quantities (MB52_2000)"
    ]:
        result[col] = result[col].fillna(0)

    logger.info(f"✅ Variance rows found: {len(result)} out of {len(df_sel)} total")
    if not result.empty:
        logger.info(f" Sample:\n{result.head(5).to_string()}")
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
        if v < 0:
            return "#FFD7D7"
        if v > 0:
            return "#FFF3CD"
        return "#FFFFFF"

    def variance_cell_colour(val):
        try:
            v = float(val)
        except (ValueError, TypeError):
            return "#FFFFFF"
        if v < 0:
            return "#FF6B6B"
        if v > 0:
            return "#FFA500"
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
    <div style="
        background:#EEF5FF;
        border:1px solid #BDD7EE;
        border-left:6px solid {total_color};
        padding:12px 16px;
        margin:12px 0 18px 0;
        font-family:Calibri, Arial, sans-serif;
        font-size:14px;
        color:#1F1F1F;
    ">
        <div><b>📊 Physical FG vs SAP — Variance Summary</b></div>
        <div style="margin-top:6px;">
            <b>Report Date:</b> {run_date}
            &nbsp;&nbsp;|&nbsp;&nbsp;
            <b>Variance Rows:</b> {len(variance_df)}
            &nbsp;&nbsp;|&nbsp;&nbsp;
            <b>Grand Total Variance:</b>
            <span style="color:{total_color}; font-weight:bold;">{total_label}</span>
            &nbsp;&nbsp;(Negative = excess stock in SAP)
        </div>
    </div>
    """

    if variance_df.empty:
        table_section = """
        <p style="font-size:14px; color:#008000; font-weight:bold;">
            ✅ No variance exceptions found. All records are currently balanced.
        </p>
        """
    else:
        col_cfg = [
            ("Indus Part", "130px", "#1F3864", "#FFFFFF"),
            ("Customer", "110px", "#1F3864", "#FFFFFF"),
            ("Physical Warehouse Total", "150px", "#1F3864", "#FFFFFF"),
            ("Production FG Total", "145px", "#1F3864", "#FFFFFF"),
            ("SAP Quantities (MB52_1000)", "150px", "#1F3864", "#FFFFFF"),
            ("SAP Quantities (MB52_2000)", "150px", "#1F3864", "#FFFFFF"),
            ("Final Variance (Physical vs SAP)", "170px", "#C0392B", "#FFFFFF"),
        ]

        header_row = "".join(
            f'<th style="background:{bg};color:{fg};border:1px solid #7F7F7F;padding:8px 6px;'
            f'font-size:12px;font-weight:700;text-align:center;width:{width};">{col}</th>'
            for col, width, bg, fg in col_cfg
        )

        body_rows = ""
        for _, row in variance_df.iterrows():
            row_bg = variance_row_colour(row["Final Variance (Physical vs SAP)"])
            var_bg = variance_cell_colour(row["Final Variance (Physical vs SAP)"])

            cells = []
            for col, _, _, _ in col_cfg:
                value = row[col]
                align = "left" if col in ("Indus Part", "Customer") else "right"

                if col == "Final Variance (Physical vs SAP)":
                    display = fmt_var(value)
                    bg = var_bg
                    color = "#000000"
                    weight = "bold"
                else:
                    display = fmt(value)
                    bg = row_bg
                    color = "#000000"
                    weight = "normal"

                cells.append(
                    f'<td style="border:1px solid #A6A6A6;padding:6px 8px;background:{bg};'
                    f'color:{color};font-size:12px;font-weight:{weight};text-align:{align};">{display}</td>'
                )

            body_rows += "<tr>" + "".join(cells) + "</tr>"

        table_section = f"""
        <div style="margin:8px 0 10px 0; font-size:13px;">
            🔴 Red rows = SAP has excess (negative variance) &nbsp;&nbsp;|&nbsp;&nbsp;
            🟠 Amber rows = Physical stock exceeds SAP (positive variance)
        </div>

        <table style="border-collapse:collapse; width:100%; font-family:Calibri, Arial, sans-serif;">
            <thead><tr>{header_row}</tr></thead>
            <tbody>{body_rows}</tbody>
        </table>
        """

    html = f"""
    <html>
    <body style="font-family:Calibri, Arial, sans-serif; color:#1F1F1F; font-size:14px; line-height:1.5;">
        <p>Hi Team,</p>

        <p>
            Please find below the <b>Indus Physical FG vs SAP Variance Report</b>
            as of <b>{run_date}</b>.
        </p>

        {summary_banner}

        {table_section}

        <br>
        <p>
            Regards,<br>
            <b>Sharon Patil</b><br>
            Logistics<br>
            Indus International FZC, Sharjah
        </p>
    </body>
    </html>
    """
    return html

# ============================================================
# EXCEL ATTACHMENT
# ============================================================
def build_excel_attachment(variance_df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Variance Report"

    thin = Side(style="thin", color="7F7F7F")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    header_blue_fill = PatternFill(fill_type="solid", fgColor="1F3864")
    header_red_fill = PatternFill(fill_type="solid", fgColor="C0392B")
    white_bold = Font(color="FFFFFF", bold=True)
    black_font = Font(color="000000", bold=False)
    black_bold = Font(color="000000", bold=True)

    neg_row_fill = PatternFill(fill_type="solid", fgColor="FFD7D7")
    pos_row_fill = PatternFill(fill_type="solid", fgColor="FFF3CD")
    neg_var_fill = PatternFill(fill_type="solid", fgColor="FF6B6B")
    pos_var_fill = PatternFill(fill_type="solid", fgColor="FFA500")

    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    center_align = Alignment(horizontal="center", vertical="center")

    columns = list(variance_df.columns)

    for col_idx, col_name in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.border = border
        cell.alignment = center_align
        if col_name == "Final Variance (Physical vs SAP)":
            cell.fill = header_red_fill
            cell.font = white_bold
        else:
            cell.fill = header_blue_fill
            cell.font = white_bold

    for row_idx, (_, row) in enumerate(variance_df.iterrows(), start=2):
        try:
            variance_val = float(row["Final Variance (Physical vs SAP)"])
        except (ValueError, TypeError):
            variance_val = 0

        if variance_val < 0:
            row_fill = neg_row_fill
            var_fill = neg_var_fill
        elif variance_val > 0:
            row_fill = pos_row_fill
            var_fill = pos_var_fill
        else:
            row_fill = None
            var_fill = None

        for col_idx, col_name in enumerate(columns, start=1):
            value = row[col_name]
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border

            if col_name in ("Indus Part", "Customer"):
                cell.alignment = left_align
            else:
                cell.alignment = right_align

            if col_name == "Final Variance (Physical vs SAP)":
                if var_fill:
                    cell.fill = var_fill
                cell.font = black_bold
            else:
                if row_fill:
                    cell.fill = row_fill
                cell.font = black_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = {
        "Indus Part": 16,
        "Customer": 14,
        "Physical Warehouse Total": 22,
        "Production FG Total": 20,
        "SAP Quantities (MB52_1000)": 21,
        "SAP Quantities (MB52_2000)": 21,
        "Final Variance (Physical vs SAP)": 24,
    }

    for idx, col_name in enumerate(columns, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = widths.get(col_name, 18)

    output = base64.b64encode(b"")  # placeholder to keep structure obvious
    from io import BytesIO
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return EmailConfig.ATTACHMENT_FILE_NAME, buffer.read()

# ============================================================
# SEND EMAIL
# ============================================================
def send_email_with_attachment(mail_token, html_body, attachment_name, attachment_bytes):
    logger.info("📧 Sending email with Excel attachment...")

    url = f"https://graph.microsoft.com/v1.0/users/{EmailConfig.SENDER_EMAIL}/sendMail"
    headers = {
        "Authorization": f"Bearer {mail_token}",
        "Content-Type": "application/json"
    }

    to_recipients = [{"emailAddress": {"address": email}} for email in EmailConfig.TO_EMAILS]
    cc_recipients = [{"emailAddress": {"address": email}} for email in EmailConfig.CC_EMAILS]
    attachment_b64 = base64.b64encode(attachment_bytes).decode("utf-8")

    payload = {
        "message": {
            "subject": EmailConfig.SUBJECT,
            "body": {
                "contentType": "HTML",
                "content": html_body
            },
            "toRecipients": to_recipients,
            "ccRecipients": cc_recipients,
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": attachment_name,
                    "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "contentBytes": attachment_b64
                }
            ]
        },
        "saveToSentItems": True
    }

    response = requests.post(url, headers=headers, json=payload, timeout=120)

    if response.status_code == 202:
        logger.info("✅ Email sent successfully")
    else:
        raise Exception(f"Email send failed: {response.status_code} — {response.text[:1000]}")

# ============================================================
# MAIN
# ============================================================
def main():
    try:
        logger.info("=" * 70)
        logger.info("STARTING INDUS PHYSICAL FG vs SAP VARIANCE EMAIL REPORT")
        logger.info("=" * 70)

        sp_token = get_sharepoint_token()
        mail_token = get_mail_token()

        raw_df, total_variance = read_worksheet(sp_token)
        variance_df = filter_variance_rows(raw_df)

        html_body = build_html_body(variance_df, total_variance)
        attachment_name, attachment_bytes = build_excel_attachment(variance_df)

        send_email_with_attachment(
            mail_token=mail_token,
            html_body=html_body,
            attachment_name=attachment_name,
            attachment_bytes=attachment_bytes
        )

        logger.info("🎉 Process completed successfully")

    except Exception as e:
        logger.error(f"❌ Process failed: {e}")
        traceback.print_exc()
        raise


if __name__ == "__main__":
    main()
