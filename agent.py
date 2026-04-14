#!/usr/bin/env python3
"""
Order Agent — Automated Receipt & Email → Google Sheet Pipeline
================================================================
Scans Gmail and Google Drive every 12 hours for order confirmations and
receipts, extracts structured data via Claude AI, and appends rows to a
Google Sheet.
"""

import os
import sys
import re
import json
import time
import base64
import logging
import mimetypes
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional

import schedule
from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import anthropic

# Optional: Amazon SP-API product lookup (works fine without it)
try:
    from amazon_lookup import (
        get_sp_api_access_token,
        build_inventory_cache,
        ensure_product_map_tab,
        load_product_map,
        lookup_amazon_product,
    )
    SP_API_AVAILABLE = True
except ImportError:
    SP_API_AVAILABLE = False

# ─────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────
load_dotenv()

SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = "token.json"
LAST_RUN_FILE = "last_run.txt"
PROCESSED_FILES_FILE = "processed_files.txt"
LOG_FILE = "agent.log"

DRIVE_FOLDER_NAME = os.getenv("DRIVE_FOLDER_NAME", "receipts")
SHEET_NAME = os.getenv("SHEET_NAME", "ORDER SHEET")
TAB_NAME = os.getenv("TAB_NAME", "Orders")

CLAUDE_MODEL = os.getenv("CLAUDE_MODEL", "claude-sonnet-4-20250514")
FIRST_RUN_LOOKBACK_HOURS = int(os.getenv("LOOKBACK_HOURS", "168"))  # 7 days
RUN_INTERVAL_HOURS = int(os.getenv("RUN_INTERVAL_HOURS", "12"))

ORDER_KEYWORDS = [
    "order", "confirmation", "commande", "invoice",
    "facture", "receipt", "reçu", "réception",
    "purchase", "payment", "achat", "paiement",
]

RECEIPT_EXTENSIONS = {".pdf", ".jpg", ".jpeg", ".png"}

# ─────────────────────────────────────────────
# Load custom store rules if present
# ─────────────────────────────────────────────
CUSTOM_RULES_FILE = "store_rules.txt"

def load_custom_rules() -> str:
    """Load custom store-specific extraction rules from store_rules.txt."""
    if os.path.exists(CUSTOM_RULES_FILE):
        return Path(CUSTOM_RULES_FILE).read_text(encoding="utf-8").strip()
    return ""

# ─────────────────────────────────────────────
# Tax configuration
# ─────────────────────────────────────────────
# Customize these for your province/region.
# Defaults are for Quebec (GST 5% + QST 9.975%).
# Ontario: set TAX1_NAME=HST, TAX1_RATE=0.13, remove TAX2 vars.
# Alberta: set TAX1_NAME=GST, TAX1_RATE=0.05, remove TAX2 vars.

TAX1_NAME = os.getenv("TAX1_NAME", "GST")
TAX1_RATE = os.getenv("TAX1_RATE", "5%")
TAX2_NAME = os.getenv("TAX2_NAME", "QST")
TAX2_RATE = os.getenv("TAX2_RATE", "9.975%")


def build_extraction_prompt() -> str:
    """Build the system prompt for Claude, including any custom store rules."""
    custom_rules = load_custom_rules()
    custom_section = ""
    if custom_rules:
        custom_section = f"""

STORE-SPECIFIC RULES (from your store_rules.txt):
{custom_rules}
"""

    return f"""You are an order processing agent for an Amazon reselling business based in Canada.
Extract purchase data from receipts and emails and return ONLY structured JSON. No explanations.

COLUMNS (exactly 15 fields per item):
1. Date — DD/MM/YYYY format always.
2. Name — descriptive product name from the receipt. NOT SKU numbers.
3. Landed — always empty string.
4. ASIN — only real Amazon ASINs (10 chars starting with "B0"). Store SKUs are NOT ASINs. Leave empty for store receipts.
5. {TAX1_NAME} — {TAX1_RATE} tax amount as a number. If one total for multiple items, split proportionally.
6. {TAX2_NAME} — {TAX2_RATE} tax amount as a number. Same proportional split rule. Leave empty if your province has no second tax.
7. Units — total quantity. Just a number.
8. SubTotal — pre-tax total for ALL units of this item. Just a number.
9. PPU — price per unit = SubTotal / Units. Just a number.
10. Location — store name with city if visible (e.g. "Walmart Toronto", "Costco Laval").
11. OrderID — reference number, invoice number, or receipt number.
12. Notes — anomalies, promo codes, quantity limits. Empty string if nothing.
13. CARD — last 4 digits of the card that SUCCESSFULLY paid. If a payment was refused and another card was used, use the one that actually paid.
14. FinalPrices — for single-item receipts: TOTAL line. For multi-item: SubTotal + proportional taxes. Just a number.
15. MessageId — echo back the filename or email ID provided.
{custom_section}
TAX SPLITTING FOR MULTI-ITEM RECEIPTS:
When receipt has multiple products but one tax total:
- Item Tax1 = (item SubTotal / receipt SubTotal) × total Tax1. Round to 2 decimals.
- Item Tax2 = (item SubTotal / receipt SubTotal) × total Tax2. Round to 2 decimals.
- Item FinalPrices = item SubTotal + item Tax1 + item Tax2

GENERAL RECEIPT READING:
- Look for quantity indicators: "QTY", "x", "@", or a number before a product description
- Date formats vary — always output DD/MM/YYYY regardless of input format
- SubTotal = pre-tax total, FinalPrices = after-tax total per item
- If GST/QST are not explicitly listed but you know the total and subtotal, calculate them
- For multi-item receipts: EVERY distinct product gets its OWN separate row
- Skip lines with $0.00 total
- SubTotal = Units × PPU (verify this)
- FinalPrices ≈ SubTotal + Tax1 + Tax2 (verify this)
- Order confirmations WITH product names, quantities, and prices ARE valid — extract them
- NEVER skip an email just because the date is missing. Leave Date empty and still extract.
- For non-order emails with no extractable product data: return {{"skip": "reason"}}

OUTPUT: Your ENTIRE response must be ONLY valid JSON. No text before or after.
Example:
{{"rows": [
{{"Date":"03/03/2026","Name":"Product Name Here","Landed":"","ASIN":"","{TAX1_NAME}":21.0,"{TAX2_NAME}":41.89,"Units":6,"SubTotal":420.0,"PPU":70.0,"Location":"Store Name City","OrderID":"123456","Notes":"","CARD":"2259","FinalPrices":482.89,"MessageId":"receipt.pdf"}}
]}}"""


# ─────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("order-agent")


# ═════════════════════════════════════════════
#  Google Auth
# ═════════════════════════════════════════════
def get_google_credentials() -> Credentials:
    creds: Optional[Credentials] = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            log.info("Refreshing expired Google token ...")
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                log.error("credentials.json not found.")
                sys.exit(1)
            log.info("Starting OAuth2 consent flow ...")
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())
        log.info("Google token saved.")
    return creds


# ═════════════════════════════════════════════
#  Timestamp helpers
# ═════════════════════════════════════════════
def get_last_run_time() -> datetime:
    if os.path.exists(LAST_RUN_FILE):
        raw = Path(LAST_RUN_FILE).read_text().strip()
        try:
            return datetime.fromisoformat(raw)
        except ValueError:
            pass
    return datetime.now(timezone.utc) - timedelta(hours=FIRST_RUN_LOOKBACK_HOURS)


def save_last_run_time() -> None:
    Path(LAST_RUN_FILE).write_text(datetime.now(timezone.utc).isoformat())


def load_processed_files() -> set[str]:
    if os.path.exists(PROCESSED_FILES_FILE):
        return set(Path(PROCESSED_FILES_FILE).read_text().splitlines())
    return set()


def mark_file_processed(file_id: str) -> None:
    with open(PROCESSED_FILES_FILE, "a") as f:
        f.write(file_id + "\n")


# ═════════════════════════════════════════════
#  Google Sheets helpers
# ═════════════════════════════════════════════
def get_sheet_id(sheets_service, spreadsheet_name: str) -> Optional[str]:
    creds = sheets_service._http.credentials
    drive = build("drive", "v3", credentials=creds)
    query = (
        f"name='{spreadsheet_name}' and "
        "mimeType='application/vnd.google-apps.spreadsheet' and "
        "trashed=false"
    )
    results = drive.files().list(q=query, fields="files(id,name)").execute()
    files = results.get("files", [])
    return files[0]["id"] if files else None


def get_existing_message_ids(sheets_service, spreadsheet_id: str) -> set[str]:
    ids = set()
    for col in ["O", "P"]:
        try:
            result = (
                sheets_service.spreadsheets()
                .values()
                .get(spreadsheetId=spreadsheet_id, range=f"'{TAB_NAME}'!{col}:{col}")
                .execute()
            )
            for row in result.get("values", []):
                if row and row[0]:
                    ids.add(row[0])
        except HttpError:
            pass
    return ids


def get_existing_fingerprints(sheets_service, spreadsheet_id: str) -> set[str]:
    """Load Date (A) + Location (J) + SubTotal (H) from existing rows
    to detect possible duplicates from different sources."""
    fingerprints = set()
    try:
        result = (
            sheets_service.spreadsheets()
            .values()
            .get(spreadsheetId=spreadsheet_id, range=f"'{TAB_NAME}'!A:K")
            .execute()
        )
        for row in result.get("values", [])[1:]:  # skip header
            date_val = row[0].strip() if len(row) > 0 else ""
            subtotal_val = row[7].strip() if len(row) > 7 else ""
            location_val = row[10].strip() if len(row) > 10 else ""
            if date_val and subtotal_val:
                fingerprints.add(f"{date_val}|{location_val}|{subtotal_val}")
    except HttpError as e:
        log.warning("Could not load fingerprints: %s", e)
    return fingerprints


def append_rows(sheets_service, spreadsheet_id: str, rows: list[list[str]]) -> int:
    if not rows:
        return 0
    body = {"values": rows}
    for attempt in range(3):
        try:
            result = (
                sheets_service.spreadsheets()
                .values()
                .append(
                    spreadsheetId=spreadsheet_id,
                    range=f"'{TAB_NAME}'!A1",
                    valueInputOption="USER_ENTERED",
                    insertDataOption="INSERT_ROWS",
                    body=body,
                )
                .execute()
            )
            updated = result.get("updates", {}).get("updatedRows", 0)
            log.info("Appended %d rows to Google Sheet.", updated)
            return updated
        except HttpError as e:
            if e.resp.status >= 500 and attempt < 2:
                log.warning("Google API error (attempt %d): %s. Retrying...", attempt + 1, e)
                time.sleep(5)
            else:
                raise
    return 0


# ═════════════════════════════════════════════
#  Claude API
# ═════════════════════════════════════════════
def call_claude_text(client: anthropic.Anthropic, system_prompt: str, text: str) -> str:
    message = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=8192,
        system=system_prompt,
        messages=[{"role": "user", "content": text}],
    )
    return message.content[0].text


def call_claude_image(client: anthropic.Anthropic, system_prompt: str, b64_data: str, media_type: str, filename: str) -> str:
    message = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=8192,
        system=system_prompt,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64_data}},
                {"type": "text", "text": f"Extract order data from this receipt image. Filename: {filename}"},
            ],
        }],
    )
    return message.content[0].text


def call_claude_pdf(client: anthropic.Anthropic, system_prompt: str, b64_data: str, filename: str) -> str:
    message = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=8192,
        system=system_prompt,
        messages=[{
            "role": "user",
            "content": [
                {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": b64_data}},
                {"type": "text", "text": f"Extract order data from this receipt PDF. Filename: {filename}"},
            ],
        }],
    )
    return message.content[0].text


# ═════════════════════════════════════════════
#  JSON response parser
# ═════════════════════════════════════════════
# Number format: set DECIMAL_COMMA=true in .env if your Google Sheet
# uses French/European number format (comma as decimal separator).
USE_DECIMAL_COMMA = os.getenv("DECIMAL_COMMA", "false").lower() == "true"


def parse_claude_response(raw: str) -> list[list[str]]:
    text = raw.strip()
    log.info("Raw Claude response: %s", text[:800])

    # Strip code fences
    text = re.sub(r"^```(?:json)?\s*\n?", "", text, flags=re.MULTILINE)
    text = re.sub(r"\n?```\s*$", "", text, flags=re.MULTILINE)
    text = text.strip()

    # Try direct parse
    data = None
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        match = re.search(r'(\{.*"rows"\s*:\s*\[.*\]\s*\})', text, re.DOTALL)
        if not match:
            match = re.search(r'(\{.*"skip"\s*:.*\})', text, re.DOTALL)
        if match:
            try:
                data = json.loads(match.group(1))
                log.info("Extracted JSON from mixed response.")
            except json.JSONDecodeError:
                pass

    if data is None:
        log.error("Failed to parse JSON. Raw: %s", text[:500])
        return []

    if "skip" in data:
        log.info("Claude says SKIP: %s", data["skip"])
        return []

    if "rows" not in data:
        log.error("No 'rows' key in response: %s", list(data.keys()))
        return []

    # Number field indices: Tax1(4), Tax2(5), Units(6), SubTotal(7), PPU(8), FinalPrices(9)
    NUMBER_FIELDS = {4, 5, 6, 7, 8, 9}

    rows: list[list[str]] = []
    for item in data["rows"]:
        row = [
            str(item.get("Date", "")),
            str(item.get("Name", "")),
            str(item.get("Landed", "")),
            str(item.get("ASIN", "")),
            str(item.get(TAX1_NAME, item.get("GST", ""))),
            str(item.get(TAX2_NAME, item.get("QST", ""))),
            str(item.get("Units", "")),
            str(item.get("SubTotal", "")),
            str(item.get("PPU", "")),
            str(item.get("FinalPrices", "")),
            str(item.get("Location", "")),
            str(item.get("OrderID", "")),
            str(item.get("Notes", "")),
            str(item.get("CARD", "")),
            str(item.get("MessageId", "")),
        ]
        # Replace dots with commas in number fields if using French locale
        if USE_DECIMAL_COMMA:
            for i in NUMBER_FIELDS:
                if i < len(row) and row[i]:
                    row[i] = row[i].replace(".", ",")
        log.info("Parsed row: %s", row)
        rows.append(row)

    return rows


# ═════════════════════════════════════════════
#  Gmail scanning
# ═════════════════════════════════════════════
VB_LABEL_NAME = "OrderAgent-Processed"


def get_or_create_label(gmail_service) -> str:
    """Find or create the processed Gmail label. Returns the label ID."""
    try:
        results = gmail_service.users().labels().list(userId="me").execute()
        for label in results.get("labels", []):
            if label["name"] == VB_LABEL_NAME:
                return label["id"]

        label_body = {
            "name": VB_LABEL_NAME,
            "labelListVisibility": "labelShow",
            "messageListVisibility": "show",
        }
        created = gmail_service.users().labels().create(userId="me", body=label_body).execute()
        log.info("Created Gmail label: %s (id=%s)", VB_LABEL_NAME, created["id"])
        return created["id"]
    except Exception as e:
        log.error("Failed to get/create label '%s': %s", VB_LABEL_NAME, e)
        return ""


def build_gmail_query() -> str:
    keyword_clause = " OR ".join(f"subject:{kw}" for kw in ORDER_KEYWORDS)
    return f"-label:{VB_LABEL_NAME} ({keyword_clause})"


def get_email_body(payload: dict) -> str:
    body_text = ""
    if payload.get("mimeType", "").startswith("text/plain"):
        data = payload.get("body", {}).get("data", "")
        if data:
            body_text = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
    elif payload.get("mimeType", "").startswith("text/html"):
        data = payload.get("body", {}).get("data", "")
        if data:
            body_text = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
    for part in payload.get("parts", []):
        child = get_email_body(part)
        if child:
            if part.get("mimeType", "").startswith("text/plain"):
                return child
            if not body_text:
                body_text = child
    return body_text


def get_message_id_header(headers: list[dict]) -> str:
    for h in headers:
        if h.get("name", "").lower() == "message-id":
            return h.get("value", "")
    return ""


def get_attachments(gmail_service, msg_id: str, payload: dict) -> list[dict]:
    """Recursively extract PDF and image attachments from an email."""
    attachments = []
    supported_mimes = {
        "application/pdf",
        "image/jpeg", "image/jpg", "image/png",
    }

    def _walk_parts(parts, msg_id):
        for part in parts:
            mime = part.get("mimeType", "")
            filename = part.get("filename", "")
            body = part.get("body", {})
            attachment_id = body.get("attachmentId")

            if mime in supported_mimes and attachment_id and filename:
                try:
                    att = (
                        gmail_service.users().messages()
                        .attachments()
                        .get(userId="me", messageId=msg_id, id=attachment_id)
                        .execute()
                    )
                    data = att.get("data", "")
                    if data:
                        data_bytes = base64.urlsafe_b64decode(data)
                        b64_standard = base64.standard_b64encode(data_bytes).decode("ascii")
                        attachments.append({
                            "filename": filename,
                            "data": b64_standard,
                            "mimeType": mime,
                        })
                        log.info("Found attachment: %s (%s)", filename, mime)
                except Exception as e:
                    log.warning("Failed to download attachment %s: %s", filename, e)

            if "parts" in part:
                _walk_parts(part["parts"], msg_id)

    if "parts" in payload:
        _walk_parts(payload["parts"], msg_id)

    return attachments


def process_emails(gmail_service, claude_client: anthropic.Anthropic, system_prompt: str) -> tuple[list[list[str]], bool]:
    """Returns (rows, had_errors)."""
    label_id = get_or_create_label(gmail_service)
    if not label_id:
        log.error("Could not get processed label. Skipping email scan.")
        return [], True

    query = build_gmail_query()
    log.info("Gmail query: %s", query)
    all_rows: list[list[str]] = []
    had_errors = False
    page_token = None

    while True:
        try:
            result = (
                gmail_service.users().messages()
                .list(userId="me", q=query, pageToken=page_token)
                .execute()
            )
        except HttpError as e:
            log.error("Gmail list error: %s", e)
            had_errors = True
            break

        messages = result.get("messages", [])
        if not messages:
            log.info("No (more) matching emails found.")
            break

        for msg_meta in messages:
            msg_id = msg_meta["id"]
            try:
                msg = gmail_service.users().messages().get(userId="me", id=msg_id, format="full").execute()
                headers = msg.get("payload", {}).get("headers", [])
                message_id_header = get_message_id_header(headers)
                subject = next((h["value"] for h in headers if h["name"].lower() == "subject"), "(no subject)")
                email_date_str = next((h["value"] for h in headers if h["name"].lower() == "date"), "")
                log.info("Processing email: %s [%s]", subject, message_id_header)

                # Parse email date for fallback
                email_date_fallback = ""
                if email_date_str:
                    try:
                        from email.utils import parsedate_to_datetime
                        email_dt = parsedate_to_datetime(email_date_str)
                        email_date_fallback = email_dt.strftime("%d/%m/%Y")
                    except Exception:
                        pass

                attachments = get_attachments(gmail_service, msg_id, msg.get("payload", {}))
                pdf_attachments = [a for a in attachments if a["mimeType"] == "application/pdf"]
                img_attachments = [a for a in attachments if a["mimeType"].startswith("image/")]

                if pdf_attachments or img_attachments:
                    for att in pdf_attachments:
                        log.info("Processing PDF attachment: %s", att["filename"])
                        raw = call_claude_pdf(claude_client, system_prompt, att["data"], att["filename"])
                        rows = parse_claude_response(raw)
                        for row in rows:
                            if len(row) > 14:
                                row[14] = message_id_header
                            if email_date_fallback and len(row) > 0 and not row[0].strip():
                                row[0] = email_date_fallback
                        all_rows.extend(rows)

                    for att in img_attachments:
                        log.info("Processing image attachment: %s", att["filename"])
                        raw = call_claude_image(claude_client, system_prompt, att["data"], att["mimeType"], att["filename"])
                        rows = parse_claude_response(raw)
                        for row in rows:
                            if len(row) > 14:
                                row[14] = message_id_header
                            if email_date_fallback and len(row) > 0 and not row[0].strip():
                                row[0] = email_date_fallback
                        all_rows.extend(rows)
                else:
                    body = get_email_body(msg.get("payload", {}))
                    if not body:
                        log.warning("Empty body for message %s, skipping.", msg_id)
                        continue

                    prompt = f"Email Subject: {subject}\nMessageId: {message_id_header}\n\n{body}"
                    raw = call_claude_text(claude_client, system_prompt, prompt)
                    rows = parse_claude_response(raw)
                    for row in rows:
                        if email_date_fallback and len(row) > 0 and not row[0].strip():
                            row[0] = email_date_fallback
                    all_rows.extend(rows)

                gmail_service.users().messages().modify(
                    userId="me", id=msg_id, body={"addLabelIds": [label_id]}
                ).execute()
                log.info("Labeled email as processed.")

            except Exception as exc:
                log.error("Error processing email %s: %s", msg_id, exc, exc_info=True)
                had_errors = True

        page_token = result.get("nextPageToken")
        if not page_token:
            break

    return all_rows, had_errors


# ═════════════════════════════════════════════
#  Google Drive scanning
# ═════════════════════════════════════════════
def find_drive_folder(drive_service, folder_name: str) -> Optional[str]:
    query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    result = drive_service.files().list(q=query, fields="files(id)").execute()
    files = result.get("files", [])
    return files[0]["id"] if files else None


def list_vendor_folders(drive_service, receipts_folder_id: str) -> dict[str, str]:
    """Return {folder_name_lowercase: folder_id} for all subfolders in receipts."""
    folders = {}
    query = (
        f"'{receipts_folder_id}' in parents and "
        "mimeType='application/vnd.google-apps.folder' and trashed=false"
    )
    page_token = None
    while True:
        result = drive_service.files().list(
            q=query, fields="files(id,name)", pageToken=page_token
        ).execute()
        for f in result.get("files", []):
            folders[f["name"].lower()] = f["id"]
        page_token = result.get("nextPageToken")
        if not page_token:
            break
    return folders


def match_vendor_folder(location: str, vendor_folders: dict[str, str]) -> tuple[Optional[str], Optional[str]]:
    """Match an extracted Location to an existing vendor folder."""
    if not location:
        return None, None

    loc_lower = location.lower().strip()

    if loc_lower in vendor_folders:
        return loc_lower, vendor_folders[loc_lower]

    for folder_name_lower, folder_id in vendor_folders.items():
        if folder_name_lower in loc_lower:
            return folder_name_lower, folder_id

    for folder_name_lower, folder_id in vendor_folders.items():
        if loc_lower in folder_name_lower:
            return folder_name_lower, folder_id

    loc_words = [w for w in loc_lower.split() if len(w) >= 4]
    for word in loc_words:
        for folder_name_lower, folder_id in vendor_folders.items():
            if word in folder_name_lower:
                return folder_name_lower, folder_id

    return None, None


def find_or_create_subfolder(drive_service, parent_id: str, subfolder_name: str) -> str:
    """Find a subfolder by name under parent_id, or create it."""
    query = (
        f"'{parent_id}' in parents and "
        f"name='{subfolder_name}' and "
        "mimeType='application/vnd.google-apps.folder' and trashed=false"
    )
    result = drive_service.files().list(q=query, fields="files(id)").execute()
    files = result.get("files", [])
    if files:
        return files[0]["id"]

    metadata = {
        "name": subfolder_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id],
    }
    folder = drive_service.files().create(body=metadata, fields="id").execute()
    log.info("Created subfolder: %s", subfolder_name)
    return folder["id"]


def classify_and_move_file(
    drive_service,
    file_id: str,
    filename: str,
    rows: list[list[str]],
    receipts_folder_id: str,
    vendor_folders: dict[str, str],
) -> None:
    """Rename and move a receipt file into vendor/yyyy-mm subfolder."""
    if not rows:
        unknown_id = find_or_create_subfolder(drive_service, receipts_folder_id, "Unknown")
        drive_service.files().update(
            fileId=file_id,
            addParents=unknown_id,
            removeParents=receipts_folder_id,
            fields="id,parents",
        ).execute()
        log.info("Moved %s -> Unknown/", filename)
        return

    first_row = rows[0]
    date_str = first_row[0] if len(first_row) > 0 else ""
    location = first_row[10] if len(first_row) > 10 else ""

    file_date = ""
    date_folder = ""
    if date_str:
        try:
            parsed = datetime.strptime(date_str, "%d/%m/%Y")
            file_date = parsed.strftime("%Y-%m-%d")
            date_folder = parsed.strftime("%Y-%m")
        except ValueError:
            pass

    matched_name, vendor_folder_id = match_vendor_folder(location, vendor_folders)

    if not vendor_folder_id:
        vendor_name = location.strip() or "Unknown"
        new_folder = find_or_create_subfolder(drive_service, receipts_folder_id, vendor_name)
        vendor_folder_id = new_folder
        matched_name = vendor_name.lower()
        vendor_folders[matched_name] = vendor_folder_id
        log.info("Auto-created vendor folder: %s", vendor_name)

    if date_folder:
        target_folder_id = find_or_create_subfolder(drive_service, vendor_folder_id, date_folder)
    else:
        target_folder_id = vendor_folder_id

    ext = Path(filename).suffix.lower()
    store_clean = re.sub(r'[^\w\s-]', '', matched_name).strip().replace(' ', '_')
    if file_date:
        new_name = f"{file_date}_receipt_{store_clean}{ext}"
    else:
        new_name = f"receipt_{store_clean}{ext}"

    drive_service.files().update(
        fileId=file_id,
        body={"name": new_name},
        addParents=target_folder_id,
        removeParents=receipts_folder_id,
        fields="id,parents",
    ).execute()
    log.info("Classified: %s -> %s/%s/%s", filename, matched_name, date_folder, new_name)


def process_drive_receipts(drive_service, claude_client: anthropic.Anthropic, system_prompt: str) -> list[list[str]]:
    folder_id = find_drive_folder(drive_service, DRIVE_FOLDER_NAME)
    if not folder_id:
        log.warning("Drive folder '%s' not found.", DRIVE_FOLDER_NAME)
        return []

    processed = load_processed_files()
    vendor_folders = list_vendor_folders(drive_service, folder_id)
    log.info("Found %d vendor folders in receipts.", len(vendor_folders))

    all_rows: list[list[str]] = []
    query = (
        f"'{folder_id}' in parents and trashed=false and "
        "mimeType!='application/vnd.google-apps.folder'"
    )
    page_token = None

    while True:
        result = drive_service.files().list(
            q=query, fields="files(id,name,mimeType)", pageToken=page_token
        ).execute()

        for f in result.get("files", []):
            file_id = f["id"]
            filename = f["name"]
            ext = Path(filename).suffix.lower()

            if ext not in RECEIPT_EXTENSIONS:
                continue
            if file_id in processed:
                log.info("Already processed: %s", filename)
                continue

            log.info("Processing Drive file: %s", filename)

            try:
                content = drive_service.files().get_media(fileId=file_id).execute()
                b64 = base64.standard_b64encode(content).decode("ascii")

                if ext == ".pdf":
                    raw = call_claude_pdf(claude_client, system_prompt, b64, filename)
                else:
                    img_mime = mimetypes.guess_type(filename)[0] or "image/jpeg"
                    raw = call_claude_image(claude_client, system_prompt, b64, img_mime, filename)

                rows = parse_claude_response(raw)
                all_rows.extend(rows)
                mark_file_processed(file_id)

                try:
                    classify_and_move_file(
                        drive_service, file_id, filename, rows,
                        folder_id, vendor_folders,
                    )
                except Exception as cls_exc:
                    log.error("Classification failed for %s: %s", filename, cls_exc, exc_info=True)

            except Exception as exc:
                log.error("Error processing %s: %s", filename, exc, exc_info=True)

        page_token = result.get("nextPageToken")
        if not page_token:
            break

    return all_rows


# ═════════════════════════════════════════════
#  Main cycle
# ═════════════════════════════════════════════
def run_cycle() -> None:
    log.info("=" * 60)
    log.info("Starting processing cycle ...")
    log.info("=" * 60)

    system_prompt = build_extraction_prompt()

    creds = get_google_credentials()
    gmail_service = build("gmail", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    sheets_service = build("sheets", "v4", credentials=creds)

    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        log.error("ANTHROPIC_API_KEY not set. Aborting.")
        return
    claude_client = anthropic.Anthropic(api_key=api_key, max_retries=5)

    spreadsheet_id = get_sheet_id(sheets_service, SHEET_NAME)
    if not spreadsheet_id:
        log.error("Spreadsheet '%s' not found.", SHEET_NAME)
        return

    # ── Optional: SP-API Amazon product lookup ──
    sp_access_token = None
    inventory_cache = []
    product_map = []
    if SP_API_AVAILABLE:
        sp_access_token = get_sp_api_access_token()
        if sp_access_token:
            ensure_product_map_tab(sheets_service, spreadsheet_id)
            product_map = load_product_map(sheets_service, spreadsheet_id)
            log.info("Loaded %d entries from Product Map.", len(product_map))
            inventory_cache = build_inventory_cache(sp_access_token)
        else:
            log.info("SP-API not configured -- Amazon product lookup disabled.")
    else:
        log.info("amazon_lookup.py not found -- Amazon product lookup disabled.")

    existing_ids = get_existing_message_ids(sheets_service, spreadsheet_id)
    log.info("Found %d existing MessageIds in sheet.", len(existing_ids))

    existing_fingerprints = get_existing_fingerprints(sheets_service, spreadsheet_id)
    log.info("Loaded %d existing row fingerprints for duplicate detection.", len(existing_fingerprints))

    since = get_last_run_time()
    log.info("Looking for items since %s", since.isoformat())

    email_rows, had_email_errors = process_emails(gmail_service, claude_client, system_prompt)
    log.info("Extracted %d rows from emails.", len(email_rows))

    drive_rows = process_drive_receipts(drive_service, claude_client, system_prompt)
    log.info("Extracted %d rows from Drive receipts.", len(drive_rows))

    # Deduplication
    all_rows = email_rows + drive_rows
    new_rows = []
    seen_order_keys = set()
    for row in all_rows:
        mid = row[14] if len(row) > 14 else ""
        if mid and mid in existing_ids:
            log.info("Duplicate MessageId (already in sheet), skipping: %s", mid)
            continue

        order_id = row[11] if len(row) > 11 else ""
        date_val = row[0] if len(row) > 0 else ""
        name_val = row[1] if len(row) > 1 else ""
        if order_id:
            order_key = f"{order_id}|{date_val}|{name_val}"
            if order_key in seen_order_keys:
                log.info("Duplicate OrderID+Date+Name in batch, skipping: %s", order_key)
                continue
            seen_order_keys.add(order_key)

        location_val = row[10] if len(row) > 10 else ""
        subtotal_val = row[7] if len(row) > 7 else ""
        fingerprint = f"{date_val}|{location_val}|{subtotal_val}"
        if fingerprint in existing_fingerprints:
            notes_idx = 12
            existing_note = row[notes_idx] if len(row) > notes_idx else ""
            if existing_note:
                row[notes_idx] = existing_note + " | POSSIBLE DUPLICATE"
            else:
                row[notes_idx] = "POSSIBLE DUPLICATE"
            log.warning("Possible duplicate detected: %s", fingerprint)

        new_rows.append(row)

    for row in new_rows:
        mid = row[14] if len(row) > 14 else ""
        if mid:
            existing_ids.add(mid)

    log.info("%d new rows after deduplication.", len(new_rows))

    # ── Optional: SP-API product name/ASIN enrichment ──
    if new_rows and inventory_cache and SP_API_AVAILABLE:
        log.info("Running Amazon product lookup on %d rows...", len(new_rows))
        for row in new_rows:
            vendor_name = row[1] if len(row) > 1 else ""
            if vendor_name:
                amazon_name, asin = lookup_amazon_product(
                    vendor_name, sheets_service, spreadsheet_id,
                    product_map, inventory_cache,
                )
                row[1] = amazon_name
                row[3] = asin
        log.info("Amazon product lookup complete.")

    if new_rows:
        append_rows(sheets_service, spreadsheet_id, new_rows)

    if had_email_errors:
        log.warning("Some emails failed. NOT advancing last_run timestamp so they get retried.")
    else:
        save_last_run_time()

    log.info("Cycle complete.\n")


def main():
    log.info("Order Agent starting up ...")
    run_cycle()
    schedule.every(RUN_INTERVAL_HOURS).hours.do(run_cycle)
    log.info("Scheduled to run every %d hours. Ctrl+C to stop.", RUN_INTERVAL_HOURS)
    try:
        while True:
            schedule.run_pending()
            time.sleep(60)
    except KeyboardInterrupt:
        log.info("Agent stopped by user.")


if __name__ == "__main__":
    main()
