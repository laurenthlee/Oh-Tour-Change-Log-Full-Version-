# Sent Task – Google Sheets × LINE OA
<img width="1174" height="632" alt="image" src="https://github.com/user-attachments/assets/1b7bd9bc-f9ab-462f-b459-c1341af8250a" />
A Google Apps Script project that sends **task updates** from a Google Sheet to a **LINE Official Account (OA)** group.

It supports:

- ✅ Manual sending of selected rows with status **“On request”**
- ⏰ Automatic sending every day at **07:00** for rows with status **“Pending”** (controlled by a toggle in the UI)
- 👀 Preview dialogs so you can see the exact message before it goes out
- 🧩 Automatic splitting of long messages to respect LINE’s character limit
- 🧠 Flexible header detection (supports multiple English/Thai header variants)

> Default timezone: **Asia/Bangkok**

---

## Table of Contents

1. [Overview](#overview)  
2. [Features](#features)  
3. [Requirements](#requirements)  
4. [Sheet Layout & Headers](#sheet-layout--headers)  
5. [Installation](#installation)  
6. [Configuration](#configuration)  
7. [LINE OA Setup & Secrets](#line-oa-setup--secrets)  
8. [Usage](#usage)  
9. [Message Format](#message-format)  
10. [How It Works (Under the Hood)](#how-it-works-under-the-hood)  
11. [Troubleshooting](#troubleshooting)  
12. [Recommended Repository Structure](#recommended-repository-structure)  
13. [License](#license)

---

## Overview

**Sent Task** is built for teams who maintain a **task / change log** in Google Sheets and need to send clear, formatted updates to their LINE group.

Typical workflow:

1. You maintain a Sheet with change requests or tasks.
2. Each task has a **Status** (e.g. “Pending”, “On request”).
3. The script:
   - Lets you **manually send** selected “On request” tasks.
   - Can **automatically send** all “Pending” tasks at 07:00 each day (if the toggle is ON).
   - Provides a **preview** so you can review text before sending.

---

## Features

### Manual “On request” sending

- Dialog: **Sent Task → ▶ Sent Task - On Request**
- Shows all rows whose Status matches **“On request”** (case-insensitive, also matches `on   request`).
- Card for each row:
  - Checkbox
  - Product Name
  - Action details (monospaced box)
  - Direct link back to that row in the sheet
- Preview area at the bottom shows the **exact combined LINE message** that will be sent.
- Buttons:
  - **Select all**
  - **Clear**
  - **Send selected**

---

### Auto @ 07:00 “Pending” sending

- Controlled by a toggle button in **Sent Task - On Request** dialog:

  > Auto daily Pending @7:00: ON / OFF

- When **ON**:
  - A time-based trigger for `autoSendDaily()` (configured by you) will send a message each day at **07:00**.
  - `autoSendDaily()`:
    - Finds all rows whose status contains **“pending”**.
    - Builds a nicely formatted message with a **Pending Task** header.
    - Splits the message into chunks if it exceeds LINE length limits.
    - Sends all chunks to the configured LINE OA group.

- When **OFF**:
  - `autoSendDaily()` returns immediately and sends **nothing**, even if the trigger runs.

> ⚠️ The toggle only controls a **script property** (`AUTO_SEND_PENDING`).  
> You still need to create one **time-based trigger** for `autoSendDaily()` in Apps Script.

---

### Pending Preview dialog

Menu: **Sent Task → ▶ Preview Task - Pending**

- Displays how many Pending rows were found.
- Shows the **full message text** (including header) that would be sent by `autoSendDaily()`.
- Buttons:
  - **Copy All**
  - **Download .txt**
  - **Close**

Preview does **not** send anything to LINE.

---

### LINE-safe message splitting

- Uses `CONFIG.LINE_CHAR_LIMIT` (default `4900`) as a safe limit.
- Splits long text by line (`\n`) when possible so each LINE message is readable.
- Each chunk is sent via the LINE Messaging API using **push** to a group or **broadcast**.

---

### Robust header detection

- Supports multiple English/Thai header variants.
- Uses:
  - `CONFIG.HEADERS` – primary expected header names
  - `CONFIG.ALT_HEADERS` – alternative header names
- If required headers are missing, the script throws a clear error instead of silently failing.

---

## Requirements

- Google account with:
  - Access to **Google Sheets**
  - Access to **Google Apps Script**
- **LINE Official Account** (OA) with:
  - Messaging API enabled
  - Channel access token
  - A LINE group where the OA bot is added
- A **task sheet** with at least these logical columns:
  - Date
  - Request By
  - Product ID
  - Product Name
  - Platform
  - Changed Type
  - Impact
  - Action
  - Status

---

## Sheet Layout & Headers

By default, the script assumes:

- Headers are on **row 1** (`HEADER_ROW: 1`).
- Columns include something like:

| Logical Field | Primary Header | Some Alt Headers (from `ALT_HEADERS`)                |
|---------------|----------------|------------------------------------------------------|
| `date`        | `Date`         | `Acted Date`, `Effective Date`, `วันที่`            |
| `requestBy`   | `Request By`   | `Request by`, `Requested By`, `Requester`, `ผู้ร้องขอ` |
| `productId`   | `Product ID`   | `Product Id`, `ProductID`, `รหัสสินค้า`            |
| `productName` | `Product Name` | `Name`, `Tour Name`, `ชื่อสินค้า`, `ชื่อทัวร์`      |
| `platform`    | `Platform`     | *(no alternates by default)*                        |
| `changedType` | `Changed Type` | `Change Type`, `Type of Change`, `ประเภทการเปลี่ยนแปลง` |
| `impact`      | `Impact`       | `Impacts`                                           |
| `action`      | `Action`       | `Actions`, `รายละเอียดการเปลี่ยนแปลง`             |
| `status`      | `Status`       | `Progress`, `สถานะ`                                |

Mandatory logical fields: **`status`**, **`date`**, **`productName`**.  
If any of these cannot be found, the script throws an error with the full header row printed for debugging.

---

## Installation

### 1. Prepare your Sheet

1. Create or open the Google Sheet that contains your task / change log.
2. Ensure your header row contains the required columns (see above).
3. Place headers on **row 1**, or change `CONFIG.HEADER_ROW` if you use another row.

---

### 2. Add the Apps Script

1. In the Sheet, go to **Extensions → Apps Script**.
2. Delete any default code in `Code.gs`.
3. Create a file (e.g. `main.gs`) and paste the full script (the file you see in this repository).
4. Click **Save**.

---

### 3. Configure `CONFIG`

At the top of the script you’ll see:

```js
const CONFIG = {
  HEADER_ROW: 1,
  TIMEZONE: 'Asia/Bangkok',
  SHEET_NAME: '',

  // LINE OA
  OA_CHANNEL_ACCESS_TOKEN: '',
  OA_GROUP_ID: '',

  // Status values
  STATUS_FILTER_VALUE: 'On request', // manual (Sent Task - On Request)
  PENDING_STATUS_VALUE: 'Pending',   // auto daily at 07:00

  // Display
  TITLE: 'On request',
  SEPARATOR_LINE: '------------------',
  LINE_CHAR_LIMIT: 4900,

  HEADERS: { ... },
  ALT_HEADERS: { ... }
};
```

Update the fields:

* `HEADER_ROW` – row number where your headers live (usually `1`).

* `TIMEZONE` – e.g. `Asia/Bangkok`.

* `SHEET_NAME`:

  * `''` → use the **active sheet**.
  * Or set to a specific tab, e.g. `"Tour Change Log"`.

* `STATUS_FILTER_VALUE` – text that means **On request** in your sheet.

* `PENDING_STATUS_VALUE` – text that means **Pending**.

* `SEPARATOR_LINE` – text line used between tasks in messages.

* `LINE_CHAR_LIMIT` – max chars per LINE message (keep below 5000; 4900 recommended).

For a simple setup you can paste your real LINE OA values directly into:

```js
  OA_CHANNEL_ACCESS_TOKEN: 'YOUR_ACCESS_TOKEN',
  OA_GROUP_ID: 'Cxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx',
```

> 💡 If you plan to put this code on GitHub, see [LINE OA Setup & Secrets](#line-oa-setup--secrets) for a safer Script Properties approach.

---

### 4. Authorize the script

1. In Apps Script, choose a function like `onOpen` or `openPreviewPending` from the dropdown.
2. Click **Run**.
3. Grant the requested permissions.

---

### 5. Reload the Sheet

After saving & authorizing:

* Reload the spreadsheet.
* You should see a new **Sent Task** menu with:

  * `▶ Preview Task - Pending`
  * `▶ Sent Task - On Request`
  * `▶ Test push to group`

---

### 6. Create the auto-send trigger

To have auto sending at **07:00**:

1. In Apps Script, go to **Triggers** (clock icon).

2. Click **+ Add Trigger**.

3. Select:

   * Function: `autoSendDaily`
   * Event source: **Time-driven**
   * Type: **Day timer**
   * Time: **7:00 – 8:00 AM**

4. Save.

Now `autoSendDaily()` will run every morning.
Whether it actually **sends messages** depends on the **AUTO toggle** in the UI.

---

## LINE OA Setup & Secrets

### Basic (quick) approach – hard-code in `CONFIG`

For internal use or private scripts, you can simply:

1. Get your **Channel Access Token** from LINE Developers.
2. Get your **groupId** (e.g. `Cbc3fe9748cbbf290bcf82de607c57627`).
3. Paste them into `CONFIG`:

```js
OA_CHANNEL_ACCESS_TOKEN: 'YOUR_TOKEN_HERE',
OA_GROUP_ID: 'Cxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx',
```

### Safer approach – Script Properties

If this code lives in a repo, don’t commit real tokens. Instead:

1. In Apps Script editor, go to **Project Settings → Script properties → Add property**.

2. Add:

   | Key                       | Value (example)                     |
   | ------------------------- | ----------------------------------- |
   | `OA_CHANNEL_ACCESS_TOKEN` | `your-long-channel-access-token`    |
   | `OA_GROUP_ID`             | `Cbc3fe9748cbbf290bcf82de607c57627` |

3. Modify `sendLineOA_()` to read from properties instead of `CONFIG` (optional enhancement).

---

## Usage

### Menu: `▶ Preview Task - Pending`

* Shows how many **Pending** tasks exist.
* Builds and displays the exact message that would be sent:

  * Includes a **Pending Task** header.
  * Includes all blocks separated by `SEPARATOR_LINE`.
* You can:

  * **Copy All**
  * **Download .txt**
  * **Close**

Nothing is sent to LINE from this dialog.

---

### Menu: `▶ Sent Task - On Request`

* Opens a picker dialog listing all **On request** tasks.

* Top bar:

  * `All: X`
  * `Selected: Y`
  * **Auto daily Pending @7:00: ON/OFF** pill
  * Toggle button **ON/OFF**:

    * Calls `setAutoDailyEnabled(true/false)` to store the flag in Script Properties (`AUTO_SEND_PENDING`).

* Middle:

  * Cards for each matching row:

    * Checkbox
    * Product Name
    * Link `#row` back to the spreadsheet
    * Action field in a monospaced box

* Bottom:

  * Preview of the combined message **including** the **On request Task** header, separator line, and all selected blocks.

* Buttons:

  * **Select all** – checks all rows.
  * **Clear** – unchecks all rows.
  * **Send selected** – calls `sendSelectedOnRequest(selectedRows)`.
  * **Close** – closes the dialog.

---

### Menu: `▶ Test push to group`

* Sends a simple test message:

```text
[TEST] Kanban Alert connectivity check: <timestamp>
```

* Useful to verify:

  * OA token is valid.
  * Group ID is correct.
  * Bot is in the group.

---

## Message Format

### Pending (auto @07:00 + preview)

`buildPendingHead_()` produces:

```text
Pending Task
[dd/MM/yy]
```

So a full message looks like:

```text
Pending Task
[04/11/25]
------------------
Date: 1/11/25
Product ID: T-001
Product Name: Example Tour
Platform: Viator
Changed Type: Price
Impact: High
Action: Increase price from 1,000 → 1,200
Requested By: Alice
------------------
Date: 2/11/25
...
```

### On request (manual send)

`buildOnRequestHead_()` produces:

```text
On request Task
[dd/MM/yy hh:mm am/pm]
```

Example:

```text
On request Task
[04/11/25 10:30 am]
------------------
Date: 4/11/25
Product ID: T-045
Product Name: Pattaya Columbia Pictures Aquaverse
Platform: GYG
Changed Type: Route
Impact: Medium
Action: Change pickup time from 09:00 → 08:30
Requested By: Nan
```

---

## How It Works (Under the Hood)

### Status matching

`hasStatus_(cell, expectedRaw)`:

* Normalizes text (lowercase + trim).
* For **On request**, allows patterns like `onrequest`, `on    request`.
* For **Pending**, exact normalized match to `PENDING_STATUS_VALUE` (default `pending`).

---

### Data collection

`getRowsByStatusForUi_(statusText)`:

1. Chooses the sheet:

   * `SHEET_NAME` if set, otherwise the active sheet.
2. Reads header row → builds index map via `robustIndexMap_()`.
3. Reads data rows.
4. For each row that matches the status:

   * Normalizes fields.
   * Builds `obj.block` (the text block shown in LINE).
   * Builds `obj.openUrl` (link back to the row).
5. Returns an array of row objects.

`getOnRequestRowsForUi_()` and `getPendingRowsForUi_()` are just wrappers that call this with different statuses.

---

### Auto-send

`autoSendDaily()`:

1. Reads `AUTO_SEND_PENDING` from Script Properties via `getAutoDailyEnabled()`.
2. If `false` → returns immediately, sending nothing.
3. If `true`:

   * Collects all Pending rows.
   * Builds header via `buildPendingHead_()`.
   * Joins blocks with `SEPARATOR_LINE`.
   * Uses `splitAndSendToLine_()` to send the message safely in chunks.

---

### Manual send

`sendSelectedOnRequest(pickedRows)`:

1. Builds header via `buildOnRequestHead_()`.
2. Joins each `row.block` with `SEPARATOR_LINE`.
3. Prepends header and separator.
4. Sends via `splitAndSendToLine_()`.

---

### LINE API

`sendLineOA_(message)`:

* Uses `CONFIG.OA_CHANNEL_ACCESS_TOKEN` and `CONFIG.OA_GROUP_ID`.
* If `OA_GROUP_ID` starts with `C`, sends a **push** to that group.
* Otherwise, uses **broadcast**.
* On error, parses LINE’s JSON and throws a detailed error message with hints for:

  * 401 – invalid/expired token
  * 403 – plan/permission issue
  * 404 – wrong group or bot not in the group

---

## Troubleshooting

### “Header ... not found at row X”

* Check `CONFIG.HEADER_ROW`.
* Verify your header texts match one of:

  * `CONFIG.HEADERS.*`
  * `CONFIG.ALT_HEADERS.*`
* Watch out for extra spaces or hidden characters (non-breaking space).

---

### Dialog shows “All: 0” or cards missing

* Are your Status values exactly “Pending” or “On request” (or close)?
* For Pending:

  * `PENDING_STATUS_VALUE` must match your text (“Pending”, “pending”, etc.).
* For On request:

  * Text must contain “on request” (any spacing).

---

### Auto send doesn’t fire at 07:00

* Check **Apps Script → Triggers**:

  * Verify there is a time-based trigger for `autoSendDaily`.
* Open **Sent Task - On Request** and verify:

  * The pill shows **Auto daily Pending @7:00: ON**.
* Check executions / logs for errors.

---

### LINE error codes

* **401 Unauthorized**

  * Invalid or expired token.
* **403 Forbidden**

  * Plan or permissions do not allow push/broadcast.
* **404 Not Found**

  * Bot is not in the group or `OA_GROUP_ID` is wrong.

The script surfaces these in the error message.

---
