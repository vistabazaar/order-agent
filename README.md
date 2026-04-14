# Order Agent

Automatically extracts purchase data from your Gmail and Google Drive receipts, then logs everything into a Google Sheet. Built for Amazon FBA resellers who buy from retail stores and need to track every purchase.

## What it does

```
Gmail (order confirmation emails)     Google Drive (receipt photos/PDFs)
            |                                      |
            └──────────────┬───────────────────────┘
                           |
                    Claude AI reads it
                           |
                  Extracts structured data
         (date, product, qty, price, tax, store...)
                           |
                  Deduplication check
                           |
                Google Sheet (auto-append)
                           |
              Drive: auto-organize receipts
          (rename + move into Store/YYYY-MM folders)
```

**Every 12 hours**, the agent:
1. Scans Gmail for order/receipt/confirmation emails
2. Scans a Google Drive folder for new receipt files (PDF, JPG, PNG)
3. Sends each to Claude AI which reads the receipt and extracts product data
4. Deduplicates against your existing sheet rows
5. Appends new rows to your Google Sheet
6. Organizes receipt files into vendor/date subfolders in Drive
7. Labels processed emails so they're not re-scanned

**Optional:** If you have Amazon SP-API access, the agent also fuzzy-matches store product names to your Amazon inventory and fills in ASINs automatically.

---

## Setup Guide (Step by Step)

### Step 0 — Prerequisites

- **Python 3.10+** — [python.org/downloads](https://www.python.org/downloads/)
  - During install, check **"Add Python to PATH"**
- A **Google account** (Gmail, Drive, Sheets)
- An **Anthropic API key** — [console.anthropic.com](https://console.anthropic.com/) (sign up, add credits, create key)

### Step 1 — Clone this repo

```bash
git clone https://github.com/YOUR_USERNAME/order-agent.git
cd order-agent
```

### Step 2 — Install Python packages

```bash
pip install -r requirements.txt
```

### Step 3 — Google Cloud Project (one-time, ~10 minutes)

This gives the agent permission to read your Gmail, Drive, and write to Sheets.

1. Go to [console.cloud.google.com](https://console.cloud.google.com/)
2. Click **Select a project** (top bar) > **New Project**
   - Name: `order-agent` (or whatever you want)
   - Click **Create**
3. Make sure your new project is selected in the top bar
4. **Enable 3 APIs** — go to **APIs & Services > Library** and search for each:
   - **Gmail API** — click Enable
   - **Google Drive API** — click Enable
   - **Google Sheets API** — click Enable
5. **Set up OAuth consent screen:**
   - Go to **APIs & Services > OAuth consent screen**
   - Choose **External** > Create
   - Fill in: App name (`order-agent`), User support email (your email), Developer email (your email)
   - Click **Save and Continue**
   - On **Scopes** page, click **Add or Remove Scopes**, then add:
     - `https://www.googleapis.com/auth/gmail.modify`
     - `https://www.googleapis.com/auth/drive`
     - `https://www.googleapis.com/auth/spreadsheets`
   - Click **Update** > **Save and Continue**
   - On **Test users** page, click **Add Users** and add your Gmail address
   - Click **Save and Continue** > **Back to Dashboard**
6. **Create credentials:**
   - Go to **APIs & Services > Credentials**
   - Click **Create Credentials > OAuth 2.0 Client ID**
   - Application type: **Desktop app**
   - Name: `order-agent`
   - Click **Create**
   - Click **Download JSON**
   - Rename the file to **`credentials.json`**
   - Move it into the `order-agent/` folder (same folder as `agent.py`)

### Step 4 — Create your Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com/) and create a new spreadsheet
2. Name it exactly: **`ORDER SHEET`** (or whatever you set in `.env`)
3. Rename the first tab to: **`Orders`**
4. In row 1, add these headers (one per cell, A1 through O1):

| A | B | C | D | E | F | G | H | I | J | K | L | M | N | O |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Date | Name | Landed? | ASIN | GST (5%) | QST (9.975%) | Units | Sub Total | PPU (No tax) | Final Prices | Location | Order ID | Notes | CARD | MessageId |

> **Note:** Adjust columns E and F for your province. Ontario = just "HST (13%)" in E, leave F blank. Alberta = just "GST (5%)" in E, leave F blank.

### Step 5 — Create your Drive folder

1. In [Google Drive](https://drive.google.com/), create a folder named exactly: **`receipts`**
2. This is where you'll upload receipt photos and PDFs
3. The agent auto-detects new files here each cycle

### Step 6 — Configure environment

```bash
cp .env.example .env
```

Edit `.env` and fill in your Anthropic API key:

```
ANTHROPIC_API_KEY=sk-ant-your-actual-key-here
```

Adjust tax settings if you're not in Quebec (see comments in `.env.example`).

### Step 7 — Add your store rules (recommended)

Edit `store_rules.txt` and describe the receipt formats for stores you buy from. The more specific you are, the better Claude parses your receipts. Example:

```
WALMART RECEIPTS:
- Items listed as: product description then price on same line
- "TC#" at bottom = OrderID
- Date at top in MM/DD/YYYY
- SUBTOTAL line = pre-tax total

COSTCO RECEIPTS:
- Item number, description, price per line
- Letters after price (A/E) indicate tax status
- Transaction number at bottom = OrderID
```

### Step 8 — First run

```bash
python agent.py
```

On the first run:
1. A browser window opens — sign in with your Google account
2. You'll see a warning "This app isn't verified" — click **Advanced > Go to order-agent (unsafe)**
   - This is your own app, this is safe
3. Grant all permissions (Gmail read, Drive read, Sheets write)
4. The browser shows "Authentication flow completed" — close it
5. A `token.json` file is created (keep it safe, don't share it)

The agent runs its first cycle immediately, then schedules every 12 hours.

### Step 9 — Keep it running

**Option A: Task Scheduler (Windows, recommended)**
1. Open **Task Scheduler** (search Start menu)
2. **Create Basic Task** > Name: `Order Agent`
3. Trigger: **When the computer starts**
4. Action: **Start a program**
   - Program: `python` (or full path like `C:\Python312\python.exe`)
   - Arguments: `agent.py`
   - Start in: full path to your order-agent folder
5. Under Properties > Settings, uncheck "Stop if runs longer than"

**Option B: Startup shortcut**
1. Press `Win+R`, type `shell:startup`, Enter
2. Create a shortcut with target: `python "C:\path\to\order-agent\agent.py"`

**Option C: Just run it manually** when you want to process receipts.

---

## How to use it

### Emails
Just buy stuff. The agent picks up order confirmation emails automatically by scanning for keywords like "order", "confirmation", "invoice", "receipt", etc. Processed emails get labeled `OrderAgent-Processed` so they're not re-scanned.

### Receipt photos/PDFs
Drop files into your Google Drive `receipts` folder. The agent:
- Reads them with Claude AI (vision for images, document parsing for PDFs)
- Extracts all product rows
- Renames the file (e.g., `2026-03-15_receipt_walmart.pdf`)
- Moves it into a `receipts/Walmart/2026-03/` subfolder

### Deduplication
The agent won't double-count:
- **Hard dedup:** Same email MessageId or same OrderID+Date+Name = skipped
- **Soft dedup:** Same Date+Location+SubTotal = added but flagged "POSSIBLE DUPLICATE"

---

## Optional: Amazon SP-API (ASIN auto-matching)

If you have SP-API access, the agent will look up receipt product names against your Amazon inventory and auto-fill the ASIN column.

Add to your `.env`:
```
SP_API_CLIENT_ID=amzn1.application-oa2-client.xxxxx
SP_API_CLIENT_SECRET=xxxxx
SP_API_REFRESH_TOKEN=xxxxx
```

**Don't have SP-API yet?** No problem — the agent works perfectly without it. You can add ASINs to your sheet manually, or set up SP-API later.

To get SP-API access:
1. Register as a developer at [developer.amazonservices.com](https://developer.amazonservices.com/)
2. Create a new app (self-authorization for your own seller account)
3. Get your LWA client ID, client secret, and refresh token
4. This process can take a few days for Amazon to approve

---

## File reference

| File | Purpose |
|------|---------|
| `agent.py` | Main script — runs the whole pipeline |
| `amazon_lookup.py` | Optional SP-API integration for ASIN matching |
| `credentials.json` | Google OAuth client secret (you create this in Step 3) |
| `token.json` | Auto-generated on first run (Google auth token) |
| `.env` | Your API keys and settings |
| `store_rules.txt` | Custom receipt parsing rules for your stores |
| `requirements.txt` | Python dependencies |
| `last_run.txt` | Timestamp of last successful run |
| `processed_files.txt` | IDs of Drive files already processed |
| `agent.log` | Full execution log |

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `credentials.json not found` | Download from Google Cloud Console (Step 3) |
| `ANTHROPIC_API_KEY not set` | Fill in `.env` file (Step 6) |
| `Spreadsheet 'ORDER SHEET' not found` | Create the sheet with exact name (Step 4) |
| `Drive folder 'receipts' not found` | Create the folder in Drive (Step 5) |
| Token expired | Delete `token.json` and run again |
| `ModuleNotFoundError` | Run `pip install -r requirements.txt` |
| Wrong data extracted | Improve your `store_rules.txt` with more detail |

---

## Security

- **Never share** `credentials.json`, `token.json`, or `.env`
- They're in `.gitignore` so they won't get committed
- The agent only reads Gmail/Drive and writes to your Sheet
- All data goes through Anthropic's API for processing
