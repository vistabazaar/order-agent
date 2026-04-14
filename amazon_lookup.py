"""
Amazon SP-API Product Lookup (Optional)
=======================================
Fetches your Amazon.ca inventory via SP-API, caches product mappings
in a Google Sheet tab ("Product Map"), and fuzzy-matches receipt product
names to your Amazon listings.

THIS FILE IS OPTIONAL. The agent works fine without it.
To use it, fill in SP-API credentials in your .env file.
"""

import os
import re
import time
import logging
import requests
from typing import Optional
from difflib import SequenceMatcher

from dotenv import load_dotenv

load_dotenv()

log = logging.getLogger("order-agent")

# ─────────────────────────────────────────────
# SP-API Configuration
# ─────────────────────────────────────────────
LWA_TOKEN_URL = "https://api.amazon.com/auth/o2/token"
SP_API_ENDPOINT = "https://sellingpartnerapi-na.amazon.com"
MARKETPLACE_ID_CA = "A2EUQ1WTGCTBG2"   # Amazon.ca

PRODUCT_MAP_TAB = "Product Map"

# Minimum similarity score (0-1) to consider a match "high confidence"
MATCH_THRESHOLD = 0.55


# ═════════════════════════════════════════════
#  LWA (Login with Amazon) Auth
# ═════════════════════════════════════════════
def get_sp_api_access_token() -> Optional[str]:
    """Exchange refresh token for a short-lived access token."""
    client_id = os.getenv("SP_API_CLIENT_ID")
    client_secret = os.getenv("SP_API_CLIENT_SECRET")
    refresh_token = os.getenv("SP_API_REFRESH_TOKEN")

    if not all([client_id, client_secret, refresh_token]):
        log.warning("SP-API credentials not configured. Skipping Amazon lookup.")
        return None

    try:
        resp = requests.post(LWA_TOKEN_URL, data={
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "client_id": client_id,
            "client_secret": client_secret,
        }, timeout=15)
        resp.raise_for_status()
        token = resp.json().get("access_token")
        if token:
            log.info("SP-API access token obtained.")
        return token
    except Exception as e:
        log.error("Failed to get SP-API access token: %s", e)
        return None


# ═════════════════════════════════════════════
#  Fetch YOUR inventory (ASINs + SKUs + titles)
# ═════════════════════════════════════════════
def fetch_my_inventory(access_token: str) -> list[dict]:
    """Pull all your FBA inventory summaries from Amazon.ca."""
    url = f"{SP_API_ENDPOINT}/fba/inventory/v1/summaries"
    items = []
    next_token = None

    while True:
        params = {
            "granularityType": "Marketplace",
            "granularityId": MARKETPLACE_ID_CA,
            "marketplaceIds": MARKETPLACE_ID_CA,
            "details": "true",
        }
        if next_token:
            params["nextToken"] = next_token

        try:
            resp = requests.get(url, params=params, headers={
                "x-amz-access-token": access_token,
                "Content-Type": "application/json",
            }, timeout=30)

            if resp.status_code == 429:
                retry_after = int(resp.headers.get("Retry-After", "5"))
                log.warning("SP-API rate limited. Waiting %ds...", retry_after)
                time.sleep(retry_after)
                continue

            resp.raise_for_status()
            data = resp.json()

            summaries = data.get("payload", {}).get("inventorySummaries", [])
            for s in summaries:
                inv = s.get("inventoryDetails", {})
                items.append({
                    "asin": s.get("asin", ""),
                    "sellerSku": s.get("sellerSku", ""),
                    "productName": s.get("productName", ""),
                    "fnSku": s.get("fnSku", ""),
                    "totalQuantity": s.get("totalQuantity", 0),
                    "fulfillable": inv.get("fulfillableQuantity", 0),
                    "inbound": inv.get("reservedQuantity", {}).get("fcProcessingQuantity", 0),
                })

            next_token = data.get("pagination", {}).get("nextToken")
            if not next_token:
                break

        except requests.exceptions.HTTPError as e:
            log.error("SP-API inventory error: %s -- %s", e, resp.text[:500] if resp else "")
            break
        except Exception as e:
            log.error("SP-API inventory error: %s", e)
            break

    log.info("Fetched %d items from SP-API inventory.", len(items))
    return items


# ═════════════════════════════════════════════
#  Product Map (Google Sheet tab)
# ═════════════════════════════════════════════
def ensure_product_map_tab(sheets_service, spreadsheet_id: str) -> None:
    """Create the 'Product Map' tab if it doesn't exist."""
    try:
        meta = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        tabs = [s["properties"]["title"] for s in meta.get("sheets", [])]
        if PRODUCT_MAP_TAB in tabs:
            return

        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [{"addSheet": {"properties": {"title": PRODUCT_MAP_TAB}}}]},
        ).execute()

        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'{PRODUCT_MAP_TAB}'!A1:F1",
            valueInputOption="RAW",
            body={"values": [["Vendor Name", "Vendor SKU", "UPC", "ASIN", "Amazon Title", "Confirmed"]]},
        ).execute()
        log.info("Created '%s' tab with headers.", PRODUCT_MAP_TAB)

    except Exception as e:
        log.error("Failed to ensure Product Map tab: %s", e)


def load_product_map(sheets_service, spreadsheet_id: str) -> list[dict]:
    """Load all rows from the Product Map tab."""
    try:
        result = (
            sheets_service.spreadsheets()
            .values()
            .get(spreadsheetId=spreadsheet_id, range=f"'{PRODUCT_MAP_TAB}'!A:F")
            .execute()
        )
        rows = result.get("values", [])
        if len(rows) <= 1:
            return []

        mappings = []
        for row in rows[1:]:
            mappings.append({
                "vendor_name": row[0].strip() if len(row) > 0 else "",
                "vendor_sku": row[1].strip() if len(row) > 1 else "",
                "upc": row[2].strip() if len(row) > 2 else "",
                "asin": row[3].strip() if len(row) > 3 else "",
                "amazon_title": row[4].strip() if len(row) > 4 else "",
                "confirmed": row[5].strip().upper() == "TRUE" if len(row) > 5 else False,
            })
        return mappings
    except Exception as e:
        log.warning("Could not load Product Map: %s", e)
        return []


def _update_product_map_title(
    sheets_service, spreadsheet_id: str,
    vendor_name_lower: str, amazon_title: str,
) -> None:
    """Update the Amazon Title column for a matched row."""
    try:
        result = (
            sheets_service.spreadsheets()
            .values()
            .get(spreadsheetId=spreadsheet_id, range=f"'{PRODUCT_MAP_TAB}'!A:E")
            .execute()
        )
        rows = result.get("values", [])
        for i, row in enumerate(rows):
            if i == 0:
                continue
            if len(row) > 0 and row[0].lower().strip() == vendor_name_lower:
                cell = f"'{PRODUCT_MAP_TAB}'!E{i + 1}"
                sheets_service.spreadsheets().values().update(
                    spreadsheetId=spreadsheet_id,
                    range=cell,
                    valueInputOption="RAW",
                    body={"values": [[amazon_title]]},
                ).execute()
                return
    except Exception as e:
        log.warning("Could not update Product Map title: %s", e)


def append_to_product_map(
    sheets_service, spreadsheet_id: str,
    vendor_name: str, vendor_sku: str, upc: str,
    asin: str, amazon_title: str, confirmed: bool = False,
) -> None:
    """Add a new row to the Product Map tab."""
    try:
        sheets_service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f"'{PRODUCT_MAP_TAB}'!A1",
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body={"values": [[vendor_name, vendor_sku, upc, asin, amazon_title, str(confirmed).upper()]]},
        ).execute()
    except Exception as e:
        log.error("Failed to append to Product Map: %s", e)


# ═════════════════════════════════════════════
#  Fuzzy matching
# ═════════════════════════════════════════════
def _normalize(text: str) -> str:
    text = text.lower()
    text = re.sub(r'[^\w\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def _similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, _normalize(a), _normalize(b)).ratio()


def _keyword_overlap_score(vendor_name: str, amazon_title: str) -> float:
    vendor_words = set(_normalize(vendor_name).split())
    amazon_words = set(_normalize(amazon_title).split())

    stop = {"the", "and", "or", "of", "in", "for", "with", "a", "an",
            "de", "du", "le", "la", "les", "et", "au", "aux"}
    vendor_words = {w for w in vendor_words if len(w) >= 2 and w not in stop}
    amazon_words = {w for w in amazon_words if len(w) >= 2 and w not in stop}

    if not vendor_words:
        return 0.0
    return len(vendor_words & amazon_words) / len(vendor_words)


def find_best_match(vendor_name: str, inventory: list[dict]) -> Optional[dict]:
    """Find the best matching Amazon product from your inventory."""
    if not vendor_name or not inventory:
        return None

    best = None
    best_score = 0.0

    for item in inventory:
        candidates = []
        if item.get("productName"):
            candidates.append(item["productName"])
        if item.get("amazon_title") and item["amazon_title"] != item.get("productName"):
            candidates.append(item["amazon_title"])

        for candidate in candidates:
            seq_score = _similarity(vendor_name, candidate)
            kw_score = _keyword_overlap_score(vendor_name, candidate)
            combined = (seq_score * 0.4) + (kw_score * 0.6)

            if combined > best_score:
                best_score = combined
                best = {
                    "asin": item.get("asin", ""),
                    "amazon_title": item.get("amazon_title") or item.get("productName", ""),
                    "score": combined,
                }

    if best and best["score"] >= MATCH_THRESHOLD:
        return best
    return None


# ═════════════════════════════════════════════
#  Main lookup function (called from agent.py)
# ═════════════════════════════════════════════
def lookup_amazon_product(
    vendor_name: str,
    sheets_service,
    spreadsheet_id: str,
    product_map: list[dict],
    inventory_cache: list[dict],
) -> tuple[str, str]:
    """
    Look up an Amazon product for a vendor product name.
    Returns (name_to_use, asin_to_use).
    """
    vendor_lower = vendor_name.lower().strip()
    for mapping in product_map:
        if mapping["vendor_name"].lower().strip() == vendor_lower:
            if mapping["asin"] and mapping["amazon_title"]:
                return mapping["amazon_title"], mapping["asin"]
            elif mapping["asin"] and not mapping["amazon_title"]:
                asin = mapping["asin"]
                title = ""
                for item in inventory_cache:
                    if item.get("asin") == asin:
                        title = item.get("productName", "")
                        break
                if title:
                    mapping["amazon_title"] = title
                    _update_product_map_title(sheets_service, spreadsheet_id, vendor_lower, title)
                    return title, asin
                return vendor_name, asin
            else:
                return vendor_name, ""

    match = find_best_match(vendor_name, inventory_cache)

    if match:
        append_to_product_map(
            sheets_service, spreadsheet_id,
            vendor_name=vendor_name, vendor_sku="", upc="",
            asin=match["asin"], amazon_title=match["amazon_title"],
            confirmed=False,
        )
        product_map.append({
            "vendor_name": vendor_name, "vendor_sku": "", "upc": "",
            "asin": match["asin"], "amazon_title": match["amazon_title"],
            "confirmed": False,
        })
        return match["amazon_title"], match["asin"]
    else:
        append_to_product_map(
            sheets_service, spreadsheet_id,
            vendor_name=vendor_name, vendor_sku="", upc="",
            asin="", amazon_title="", confirmed=False,
        )
        product_map.append({
            "vendor_name": vendor_name, "vendor_sku": "", "upc": "",
            "asin": "", "amazon_title": "", "confirmed": False,
        })
        return vendor_name, ""


def build_inventory_cache(access_token: str) -> list[dict]:
    """Fetch inventory and use productName as the title."""
    items = fetch_my_inventory(access_token)
    for item in items:
        item["amazon_title"] = item.get("productName", "")
    log.info("Built inventory cache with %d products.", len(items))
    return items
