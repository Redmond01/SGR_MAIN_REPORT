# ⛽ SGR 60-Station Daily Workbook Automation & Dynamic Ingestion Engine

This project automates the daily preparation, rollover, data entry, and reconciliation of **~60 Retail Fuel Station Excel Workbooks (`.xlsx`)** for SGR.

Previously, a staff member spent **8–9 hours every day** manually opening 60 different Excel files, duplicating the latest daily sheet (`SheetN` $\rightarrow$ `SheetN+1`), copying yesterday's closing figures into today's opening figures, resetting POS and delivery cells, and typing in daily numbers across both the **Left (Outer)** and **Right (Inner Meter)** halves of every sheet.

This tool reduces that 8–9 hour manual process to **minutes** by automating sheet duplication, dynamic coordinate detection, opening-figure rollover, formula preservation, Left-to-Right mirroring, and 1-click `.zip` packaging.

---

## 1. How to Start & Stop the Application Locally

Open your terminal in `/home/im-redmond/Desktop/SGR_MAIN_REPORT`:

### Start the App
```bash
streamlit run app.py
```
*(Or run `./run.sh`)*

Your browser will open automatically at **http://localhost:8501**.

### Stop the App
Click inside the terminal window where Streamlit is running and press:
```text
Ctrl + C
```

---

## 2. Complete Walkthrough of Every Screen, Option & Button in the UI

When you open **http://localhost:8501**, the screen is divided into two areas:
- **Left Sidebar:** Uploading workbooks + Global controls (Date, Prices, Mid-Day Price Changes) that apply across multiple stations at once.
- **Main Workspace (3 Tabs):** Entering station figures, batch Excel import/export, and downloading the final `.zip` archive.

---

### A. Left Sidebar Controls (Step-by-Step)

#### 1. `Batch Upload Station Workbooks (.xlsx)`
- **What it does:** Drag and drop (or browse and select) all ~60 monthly station `.xlsx` workbooks at once.
- **What happens behind the scenes when you upload:**
  1. The engine opens each workbook and locates the **latest daily sheet** (e.g., `Sheet23`).
  2. It dynamically scans the sheet for keyword anchors (`OPENING METER`, `TANK`, `OPENNG DIPPG`, `PRODUCTS RECEIVED`, `EXPENSES`, `ADD UP`, `AVAILABLE CASH`) on both the **Left Half (Outer Meters)** and **Right Half (Inner Meters)**—so it never breaks even when stations have different row/column numbers.
  3. It prepares the **next sheet** (e.g., `Sheet24`) in memory and performs the **Daily Rollover**:
     - Today's `OPENING METER` = Yesterday's final `CLOSING METER`.
     - Today's `OPENING DIPPG` = Yesterday's `CLOSNG DIPPG`.
     - Today's `ADD UP` = Yesterday's calculated `AVAILABLE CASH`.
     - All `POS` terminals, `PRODUCTS RECEIVED` (deliveries), and `EXPENSES` are reset to `0` / blank for the new day.

#### 2. `📊 Reports Progress: Done / Total` & Clickable Station List
Right below the file uploader on the Left Sidebar, the **Station Tracker** appears:
- **Progress Counter & Bar:** Shows **`📊 Reports Progress: X / Total`** (`✅ X Done | ⏳ Y Remaining`) so you always know how many reports are completed and how many are left.
- **Filter Tabs (`All`, `⏳ Remaining`, `✅ Done`) + `🔍 Search station...`:**
  - Click **`⏳ Remaining`** to see only the stations whose reports haven't been saved yet.
  - Click **`✅ Done`** to review stations you have already completed.
  - Type in the search box to find any station immediately.
- **Clickable Station Queue (Asynchronous Entry):**
  - Because station reports arrive at different times of the day, you do **not** have to work in fixed sequential order.
  - Simply click any station's button in the Left Sidebar list (`⏳ Station Name` = Pending, `🟡 Station Name [Unsaved]` = Edited but not yet saved, `✅ Station Name` = Saved & Done) to open and work on that station immediately.

#### 3. `⚙️ Global Date & Price Controls`
Below the station list on the left sidebar, three global control panels appear:

- **`Report Date (B1 / AD1)` + `📅 Apply Date to All Stations`:**
  - **What it does:** Sets the date that will be written into cell `B1` (or `A1`/`C1`, wherever that station stores its date) on the new sheet.
  - By default, the engine automatically detects yesterday's sheet date and adds **+1 day** (e.g., if `Sheet23` was `23/09/2026`, `Sheet24` defaults to `24/09/2026`).
  - If you change the date picker and click **Apply Date to All Stations**, every loaded workbook is updated to that date.

- **`💰 Global Product Price Override`:**
  - **What it does:** Lets you update `Price 1` for `PMS (₦/L)`, `AGO (₦/L)`, or `LPG (₦/kg)` across all 60 stations at once.
  - **Note:** By default, each station automatically keeps yesterday's active price from its own workbook. Only use this box if there is a company-wide price update you want to push to every station simultaneously (leave a product at `0` if you do not want to overwrite it).

- **`⚡ Multi-Tier Price Change Selector (2, 3, or 4 Prices)`:**
  - **What it does:** In the Excel sheets, each pump has 4 native closing meter and price columns (`Closing Meter 1..4` and `Price 1..4`).
  - On a **Normal Day (1 Price)**, a station only sells at `Price 1`.
  - On a **Multi-Tier Price Change Day (2, 3, or 4 Prices)**, the attendant records intermediate meter readings each time the price changes (`Closing Meter 1` at `Price 1`, `Closing Meter 2` at `Price 2`, `Closing Meter 3` at `Price 3`, and up to `Closing Meter 4` at `Price 4`).
  - **How to use it:** Select the specific station(s) that had price changes today, pick the **Number of Active Prices Today (`2`, `3`, or `4`)**, optionally enter batch prices for `Price 2`, `Price 3`, `Price 4`, and click **Apply Tier Count & Prices to Selected Stations**. For those stations, the data-entry table will automatically expand to show the exact number of `Closing Meter`, `Price`, and `Inner Closing` columns needed.

- **`🗑️ Clear All Loaded Workbooks`:**
  - Unloads all currently loaded files from memory so you can start a fresh batch.

---

### B. Top KPI Summary Ribbon
At the top of the main page, four summary cards show your overall batch progress:
1. **Loaded Workbooks:** Total number of station `.xlsx` files currently loaded (e.g., `60 stations`).
2. **Stations with Entries:** How many stations you have entered numbers for so far (`🟢` vs `⚪`).
3. **Mid-Day Price Change Stations:** How many stations currently have the 2-tier price-change mode enabled.
4. **Target Report Date:** The date being written to the new daily sheets.

---

### C. Tab 1: `🏪 1. Station Data-Entry Workspace` (Per-Station Entry)

This is where you view and enter the day's numbers for one station at a time.

#### Save & Navigation Bar (Bottom of Tab 1)
To prevent accidental data loss before moving between stations:
- **Unsaved Changes Banner:** Whenever you edit any figure on the current station, a yellow warning banner appears: *"⚠️ You have unsaved changes on this station. Please click **💾 Save Station Report** before clicking **Next**."*
- **`💾 Save Station Report` Button:**
  - Commits your changes for the current station, marks the station **`✅ Done`** in the Left Sidebar counter (`Done / Total`), and generates that station's updated `.xlsx` sheet immediately.
- **`💾 Save & Next Remaining ➡️` Button:**
  - Saves the current station and automatically jumps to the next **`⏳ Remaining`** station.
- **`⬅️ Prev` / `Next ➡️` Buttons:**
  - **`Next ➡️` is disabled until the current station is saved**, ensuring you never click `Next` without saving your changes first.
- **`⬇️ Download This .xlsx` Button:**
  - Appears as soon as the station is saved so you can download that single station's updated `.xlsx` immediately if needed (or download all saved stations together in Tab 3).

#### Station Live KPI Ribbon
Right below the navigation bar, 6 live calculators update automatically as you type:
- **Outer Sales (₦):** Total Naira sales across all pumps (`Sales Litres × Price`).
- **PMS / AGO Litres:** Total litres dispensed today by product.
- **Total POS (₦):** Sum of all POS terminal and cash/transfer entries for today.
- **Expenses (₦):** Sum of any optional station expenses entered today.
- **Rolled ADD UP (₦):** Today's `ADD UP` figure (automatically rolled over from yesterday's `AVAILABLE CASH`).
- **Available Cash (₦):** Live preview of the Excel formula `= Total Sales - Expenses - Total POS + ADD UP`.

---

#### Section A: `Sales & Inner Meter Readings` Table
This interactive spreadsheet table (`st.data_editor`) combines **Outer Pump Readings (Left Half)** and **Inner Meter Readings (Right Half)** on the same row so you only visit each pump once:

> [!NOTE]
> **Frozen Columns:** The **`Pump`** and **`Product`** columns are **frozen (pinned on the left) by default**. When you slide or scroll horizontally to the right to enter intermediate price tiers (`Closing Meter 2..4`, `Price 2..4`) or Inner Meter readings, the pump nozzle and product stay locked in view so you always know which row you are editing.

| Column Name | Editable? | What It Means & How It Works |
| :--- | :---: | :--- |
| **`Pump`** | 🔒 Read-only (Frozen) | Pump nozzle label from the workbook (`PMS1`, `PMS 2`, `AGO 1`, `LPG PUMP`, etc.). Locked on the left when scrolling right. |
| **`Product`** | 🔒 Read-only (Frozen) | Product family (`PMS`, `AGO`, `LPG`). Locked on the left when scrolling right. |
| **`Opening Meter`** | ✏️ Editable | Automatically pre-filled with **yesterday's final Closing Meter**. You normally don't need to touch this unless correcting a historical error. |
| **`Closing Meter 1`** | ✏️ **Enter Today** | **Standard Day:** Enter today's end-of-day pump closing meter reading here.<br>**Price-Change Day:** Enter the intermediate meter reading taken when the price changed. |
| **`Price 1 (₦)`** | ✏️ Editable | Automatically pre-filled with yesterday's active pump price. automatically mirrored to the Right (Inner Meter) half. |
| **`Closing Meter 2..4`** *(Shown when 2, 3, or 4 Prices are selected)* | ✏️ **Enter Today** | Enter the meter readings for each price tier (`Closing Meter 2`, `Closing Meter 3`, `Closing Meter 4`). The engine automatically updates the Excel subtraction formulas (`Sales 2 = C2 - C1`, `Sales 3 = C3 - C2`, `Sales 4 = C4 - C3`) and locks unused trailing tiers to `0` sales. |
| **`Price 2..4 (₦)`** *(Shown when 2, 3, or 4 Prices are selected)* | ✏️ **Enter Today** | Enter the fuel price for each tier (`Price 2`, `Price 3`, `Price 4`). Automatically mirrored to the Right (Inner Meter) half. On the next day's rollover, the last active price automatically becomes tomorrow's `Price 1`. |
| **`Sales Ltrs` / `Amount (₦)`** | 🔒 Read-only | Live calculated total litres and Naira amount across all active price tiers. |
| **`Inner Opening`** | ✏️ Editable | Automatically pre-filled with **yesterday's final Inner Closing Meter** from the Right half of the sheet. |
| **`Inner Closing 1..4`** | ✏️ **Enter Today** | Enter today's physical Inner Closing Meter reading(s) matching the active price tiers. |

---

#### Section B: `Tank Dippings, Deliveries & Manifolding` Table
Located on the lower-left side of Tab 1. **Everything you enter here is automatically written to BOTH the Left Half and the Right (Inner Meter) Half of the Excel sheet** so you never enter tank numbers twice.

| Column Name | Editable? | What It Means & How It Works |
| :--- | :---: | :--- |
| **`Tank`** | 🔒 Read-only | Tank name from the workbook (`PMS1`, `PMS2`, ` PMS 7`, `AGO 1`, `LPG`, etc.). |
| **`Product`** | 🔒 Read-only | Product type (`PMS`, `AGO`, `LPG`). |
| **`Opening Dip`** | ✏️ Editable | Pre-filled automatically from **yesterday's `CLOSNG DIPPG`**. |
| **`Closing Dip`** | ✏️ **Enter Today** | Enter today's physical tank closing dipping (`CLOSNG DIPPG`). |
| **`Products Received`** | ✏️ **Enter Today** | Defaults to `0.0`. If a truck discharged fuel into this tank today, enter the volume here—the engine writes it into the `PRODUCTS RECEIVED` table on both halves of the sheet. |
| **`Tank Sales`** | 🔒 Read-only | Live preview of `Opening Dip + Products Received - Closing Dip`. |
| **`Under Maintenance`** | ☑️ Checkbox | Check this box if a tank was **under maintenance / inactive today**. When checked, the engine automatically locks `Closing Dip = Opening Dip` and `Products Received = 0` so Tank Sales = `0`. |
| **`Linked Into (Host Tank)`** | 🔽 Dropdown | **Dynamic Manifolding Control:**<br>• Automatically pre-filled by reading yesterday's Excel formulas! For example, on **Lekki 3**, `PMS 7` is manifolded into `PMS1` (`=D29+D35`), so `PMS 7` automatically shows `PMS1` here.<br>• If a tank is manifolded into another host tank today, select the host tank from the dropdown. If a previously manifolded tank is operating independently today, change its dropdown to blank (`""`). The engine dynamically rewrites the Excel sum formulas (`=D29+D35`, `=E29+E35`, etc.) on both halves of the worksheet! |

---

#### Section C: `POS Terminals & Cash Balancing`
Located on the lower-right side of Tab 1.

- **`Today's ADD UP (₦)`:**
  - Automatically pre-filled with **yesterday's evaluated `AVAILABLE CASH`** (per the rollover rule: `Today's ADD UP = Yesterday's AVAILABLE CASH`). You can adjust it if needed.
- **`Mirror POS to Inner Half` Checkbox:**
  - Checked by default. When checked, any POS amounts you enter on the Left half (`cash/transfer`, `SGR ...`, `OPAY ...`, `OBX ...`) are automatically copied to the matching POS columns on the Right (Inner Meter) half. Uncheck if you want the Right-half POS cells left untouched.
- **POS Channels Table (`POS / Channel` | `Settlement Amount (₦)`):**
  - Lists every POS machine/channel detected on that station's sheet (e.g., `OPAY A`, `OPAY B`, `cash/transfer`, `SGR 273` ... `SGR 279`).
  - All channels are automatically reset to `0.0` for the new day. Enter today's settlement amounts for the active terminals.
  - The native Excel formula in `AVAILABLE CASH` (`=SUM(Total_Sales - Expenses - POS_Channels + ADD_UP)`) is preserved 100% intact.

---

#### Section D: `🧾 Optional Expenses, Imprest & Bank Lodgment` (Expandable Drawer)
Click to expand if the station incurred expenses or recorded a bank deposit today:
- **Station Expenses (`Description` | `Amount (₦)`):**
  - Add one or more expense lines (e.g., Generator maintenance, RTT, etc.). You can click the `+` row at the bottom of the grid to add multiple rows.
  - The engine writes them into the `EXPENSES` block above the native `=SUM(...)` total row.
- **Bank Lodgment (Optional):**
  - Fill in `Bank Name`, `Amount Deposited (₦)`, and `Payment Status` if you want the `BANK LODGMENT` block updated on the sheet.

---

### D. Tab 2: `📋 2. Multi-Station Batch Grid & Excel Import`

If you prefer entering numbers for all 60 stations in a single sheet rather than clicking station-by-station in Tab 1, Tab 2 gives you two batch workflows:

1. **Master Batch Excel Template Workflow (Top of Tab 2):**
   - **Step 1 — `⬇️ Download Master Batch Entry Template (.xlsx)`:** Downloads a single Excel file containing 4 tabs (`Sales_Pumps`, `Tanks_Deliveries`, `POS_Channels`, `Expenses`) pre-filled with all 60 stations' rolled-over Opening Meters, Prices, Opening Tank Dips, and POS terminals.
   - **Step 2 — Fill in Excel:** Type or paste the day's Closing Meters, Closing Dips, Deliveries, and POS amounts in Excel and save.
   - **Step 3 — `📤 Upload Filled Master Batch Template (.xlsx)` + `⚡ Apply Batch Template to All Stations`:** Instantly populates all 60 stations inside the app in one click!
2. **Fast All-Stations Pump Closing Meter Grid (Bottom of Tab 2):**
   - Displays all pumps across all loaded stations in one scrollable table so you can rapidly tab down the `Closing Meter 1` and `Inner Closing 1` columns across stations and click **`💾 Save All-Stations Pump Grid Changes`**.

---

### E. Tab 3: `📦 3. Executive Summary & 1-Click .ZIP Download`

1. **Executive Summary Table:**
   - Shows a single unified view across all 60 stations:
     - `Station`, `Workbook`, `Rollover (Sheet23 ➔ Sheet24)`, `Date`, `PMS (L)`, `AGO (L)`, `LPG (kg)`, `Outer Sales (₦)`, `Inner Sales (₦)`, `Total POS (₦)`, `Expenses (₦)`, `ADD UP (₦)`, `Available Cash (₦)`, and `Status (🟢 Entered / ⚪ Rolled Over)`.
2. **`🚀 Process All Workbooks & Build .ZIP Archive` Button:**
   - Click this button once you are ready to generate your updated Excel files.
   - For every loaded station workbook, the engine:
     1. Clones the latest sheet to `Sheet{N+1}` (preserving all cell formatting, borders, column widths, merged cells, and native Excel formulas).
     2. Writes the rolled-over openings and your entered closing figures, prices, tank dips, deliveries, POS amounts, and expenses.
     3. Mirrors Tank Dips, Deliveries, Prices, and POS amounts to the Right (Inner Meter) half.
     4. Packages all 60 updated `.xlsx` workbooks **plus** a master `SGR_Daily_Executive_Summary.xlsx` report into a single `.zip` archive.
3. **`📦 Download All Updated Workbooks (.ZIP)` Button:**
   - Appears immediately after processing completes. Click to download the `.zip` archive containing all 60 updated workbooks.

---

## 3. Project Folder & File Structure

```text
/home/im-redmond/Desktop/SGR_MAIN_REPORT/
├── app.py                      # Main Streamlit web application (UI, tabs, grids, .zip export)
├── run.sh                      # Local startup script (runs streamlit run app.py)
├── requirements.txt            # Python dependencies (openpyxl, streamlit, pandas)
├── README.md                   # This documentation & user guide
├── .streamlit/
│   └── config.toml             # Disables Streamlit onboarding prompts / telemetry locally
├── sgr_engine/                 # Core automation package
│   ├── __init__.py             # Package exports
│   ├── parser.py               # Dynamic keyword-anchor layout detector (Left & Right halves)
│   ├── evaluator.py            # Pure-Python Excel formula evaluator (for AVAILABLE CASH rollover)
│   ├── engine.py               # Sheet duplication, rollover, dynamic tank linking & write-back
│   └── batch_io.py             # Multi-station Master Excel Template export & import
└── tools/                      # CLI inspection & verification utilities
    ├── inspect_layout.py       # Prints detected coordinates, pumps, tanks & POS for any workbook
    └── verify_all.py           # Automated test suite across all 60 station workbooks (60/60 pass)
```

---

## 4. Command-Line Diagnostic Tools (Optional)

You can also inspect any station workbook or run the 60-workbook test suite from the terminal:

### Inspect the Detected Layout of Any Workbook
```bash
python3 tools/inspect_layout.py "/home/im-redmond/Desktop/Desktop/record - Copy (3)/SEPTEMBER Lekki 3 sgr.xlsx"
```

### Run the 60-Workbook Automated Verification Suite
```bash
python3 -u tools/verify_all.py
```
