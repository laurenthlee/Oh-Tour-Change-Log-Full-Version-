# Sent Task – Google Sheets × LINE OA
<img width="1174" height="632" alt="image" src="https://github.com/user-attachments/assets/1b7bd9bc-f9ab-462f-b459-c1341af8250a" />

A Google Apps Script project that sends **task updates** from a Google Sheet to a **LINE Official Account (OA)** group.

It supports:

- ✅ Manual sending of selected rows with status **“On request”**
- ⏰ Automatic sending every day at **07:00** for rows with status **“Pending”** (optional toggle)
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
9. [How It Works (Under the Hood)](#how-it-works-under-the-hood)  
10. [Troubleshooting](#troubleshooting)  
11. [Recommended Repository Structure](#recommended-repository-structure)  
12. [License](#license)

---

## Overview

**Sent Task** is built for teams who manage change requests / product updates in Google Sheets and need to send clear, formatted updates to their LINE group.

Typical workflow:

1. You maintain a Sheet of change requests or tasks.
2. Each task has a **Status** (e.g. “Pending”, “On request”).
3. The script:
   - Lets you **manually send** selected “On request” tasks.
   - Can **automatically send** all “Pending” tasks at 07:00 each day.
   - Provides a **preview** so nothing surprising gets sent.

---

## Features

### Manual “On request” sending

- UI dialog that lists all rows with status matching **“On request”**.
- Shows:
  - Product name
  - Action detail (monospaced box)
  - Direct link back to that row in the sheet
- You can:
  - Select/unselect rows
  - See the **combined LINE message preview**
  - Click **Send selected** to push everything to a LINE OA group

---

### Auto @ 07:00 “Pending” sending

- Optional mode (toggle in the UI) that:
  - Installs a **time-based trigger** to run at **07:00** daily.
  - Each morning:
    - Finds all rows whose status contains **“pending”**.
    - Builds a nicely formatted message.
    - Splits it into chunks if it’s too long.
    - Pushes it to LINE OA.
- You can turn this on/off without touching code.

---

### Pending Preview dialog

- Menu item for **“Task Preview – Pending”**.
- Opens a read-only dialog:
  - Displays how many matching tasks were found.
  - Shows the **exact text** that would be sent.
  - Buttons:
    - **Copy All**
    - **Download .txt**
    - **Close**

---

### LINE-safe message splitting

- Uses `CONFIG.LINE_CHAR_LIMIT` (default `4900`) as a safe limit.
- Splits long text by line (`\n`) when possible so each LINE message is readable.
- Each chunk is sent using LINE Messaging API via push or broadcast.

---

### Robust header detection

- Allows multiple header variations per field (English + Thai).
- Uses both:
  - `CONFIG.HEADERS` – primary expected header names
  - `CONFIG.ALT_HEADERS` – alternative header names
- If required headers are missing, the script throws a clear error instead of silently failing.

---

## Requirements

- **Google account** with:
  - Access to **Google Sheets**
  - Access to **Google Apps Script**
- **LINE Official Account** (OA) with:
  - Messaging API enabled
  - Channel access token
  - A LINE group where the OA account is added
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

The function `robustIndexMap_()` matches both primary and alternative headers.  

Mandatory logical fields:  
`status`, `date`, `productName`  
If any of those can’t be found, the script throws a helpful error.

---

## Installation

### 1. Create or prepare the sheet

1. Create a new Google Sheet (or use an existing one).
2. On your **task tab**, ensure you have a header row with the required columns.
3. Put the headers on **row 1**, or adjust `CONFIG.HEADER_ROW` if you prefer another row.

---

### 2. Open Apps Script

1. In your Sheet, go to **Extensions → Apps Script**.
2. Delete any default `Code.gs` content.
3. Create a file (e.g. `main.gs`) and paste the full script from this project.

---

### 3. Configure `CONFIG`

At the top of the file you will see:

```js
const CONFIG = {
  HEADER_ROW: 1,
  TIMEZONE: 'Asia/Bangkok',
  SHEET_NAME: '',

  LINE_CHAR_LIMIT: 4900,
  SEPARATOR_LINE: '------------------',

  TITLES: { onrequest: 'On request', pending: 'Pending' },

  HEADERS: { ... },
  ALT_HEADERS: { ... },

  PROP_AUTO_PENDING: 'auto_pending_enabled'
};
```
---
Update:

* `HEADER_ROW`
  Row index containing the headers (e.g. `1`, `2`, …).

* `TIMEZONE`
  Usually `Asia/Bangkok` if your team is based there.

* `SHEET_NAME`

  * Empty string (`''`): use the currently active sheet.
  * Or set to a specific tab name, e.g. `"Changes"`.

* `LINE_CHAR_LIMIT`
  Max characters per LINE message chunk (recommended ≤ 4900).

* `SEPARATOR_LINE`
  Text separator between task blocks in the message.

* `TITLES`
  Title line at the top of the message for each kind of send:

  * `onrequest`: manual sends
  * `pending`: auto 07:00 / preview

---

### 4. Set Script Properties for secrets

> 🔒 **Never** hard-code real LINE tokens into the script if you plan to commit it to GitHub.

You should store them in **Script Properties** instead.

In the Apps Script editor:

1. Click **Project Settings** (gear icon).
2. Under **Script properties**, click **Add property**.
3. Add:

| Property Key              | Example Value                       |
| ------------------------- | ----------------------------------- |
| `OA_CHANNEL_ACCESS_TOKEN` | `your-long-channel-access-token`    |
| `OA_GROUP_ID`             | `Cxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx` |

These are read by a helper like:

```js
function loadSecrets_() {
  const sp = PropertiesService.getScriptProperties();
  return {
    OA_CHANNEL_ACCESS_TOKEN: sp.getProperty('OA_CHANNEL_ACCESS_TOKEN') || '',
    OA_GROUP_ID:             sp.getProperty('OA_GROUP_ID') || ''
  };
}
```

and used in `sendLineOA_()` (your script should already be wired similarly).

---

### 5. Authorize the script

1. In the Apps Script editor, choose a function like `openPreviewPending` or `pendingAutoSendRunner_` from the dropdown.
2. Click **Run**.
3. A Google auth dialog will appear:

   * Review requested permissions (Sheets, external requests, etc.).
   * Click **Allow**.

---

### 6. Reload the sheet

After the script is saved and authorized:

* Reload your Sheet.
* You should see a new menu called **“Sent Task”**.

---

## Configuration

### Status detection

Function `matchStatus_()` controls which rows belong to which mode:

```js
function matchStatus_(cell, kind) {
  const t = normText_(cell); // lowercased, trimmed
  if (!t) return false;
  if (kind === 'pending')   return /pending/.test(t);
  return /on\s*request/.test(t);
}
```

* **Pending** tasks: status contains `pending` (case-insensitive).
* **On request** tasks: status contains `on request` (with any space pattern).

Change these regexes if your statuses use a different language or wording.

---

### Date formatting

`normalizeDateDisp_(v, tz)` tries to handle:

* `Date` objects
* Text in `dd/MM/yyyy` or `yyyy-MM-dd`
* Excel serial date numbers

It outputs `d/M/yy` (e.g. `1/1/25`).

You can adjust the format inside:

```js
Utilities.formatDate(d, tz, 'd/M/yy')
```

---

### Action field formatting

For multi-line `Action` values:

```js
function formatActionForBlock_(s){
  const raw = String(s==null?'':s).replace(/\r\n/g,'\n').replace(/\u2028|\u2029/g,'\n');
  const lines = raw.split('\n').map(x=>x.replace(/\s+$/,''));
  if (lines.length===0 || (lines.length===1 && !lines[0])) return 'Action: -';
  if (lines.length===1) return 'Action: ' + lines[0];
  return 'Action: ' + lines[0] + '\n ' + lines.slice(1).join('\n ');
}
```

So the first line appears as:

```text
Action: First line
```

and subsequent lines are indented:

```text
 Second line
 Third line
```

---

## LINE OA Setup & Secrets

### 1. Set up LINE OA

In [LINE Developers](https://developers.line.biz/):

1. Create or select your **Messaging API** channel.
2. Make sure the bot is **added to your target LINE group**.
3. Issue a **Channel access token**.
4. Copy the token.

---

### 2. Get the LINE group ID

You need the `groupId` where messages should be pushed. Common approaches:

* Use a temporary webhook / debug endpoint to log `groupId` from incoming events.
* Use existing infra if you already have an endpoint behind your bot.

Once you have something like `Cbc3fe9748cbbf290bcf82de607c57627`, store it as `OA_GROUP_ID` in Script Properties.

---

### 3. Security tips

* Treat tokens like passwords.
* **Never commit** tokens or group IDs to GitHub.
* If you accidentally expose a token:

  * Immediately revoke / re-issue it in LINE Developers.
  * Update Script Properties with the new token.

---

## Usage

### Menu items

When the sheet opens, `onOpen()` registers:

#### 1. `Sent Task → ▶ Task Preview – Pending`

* Runs `openPreviewPending()`.
* Steps:

  1. Collects rows where status is **Pending**.
  2. Builds the text for LINE.
  3. Shows a modal with:

     * Total rows count
     * Text preview
     * Buttons:

       * **Copy All**
       * **Download .txt**
       * **Close**

This preview does **not** send anything to LINE.

---

#### 2. `Sent Task → ▶ Manually sent task`

* Runs `openSendPicker()`.

* Shows:

  * **Auto 07:00 (Pending)** toggle:

    * Checkbox + ON/OFF pill indicator.
    * Toggling calls `setAutoPendingEnabled(true/false)` → creates/removes the 07:00 trigger.

  * **Cards** for each row with status “On request”:

    * Checkbox
    * Product Name
    * Link to that row in the sheet (`openUrl`)
    * Action details in monospaced, scrollable box.

  * Bottom preview area:

    * Shows combined message = `head + joined blocks` for selected rows.

* Buttons:

  * `Select all` – checks all cards.
  * `Clear` – unchecks all cards.
  * `Send selected` – sends selected rows to LINE via `sendSelected_("onrequest", rows)`.
  * `Close` – closes the dialog.

---

### Automatic 07:00 sends (Pending)

If auto mode is ON:

* A time-based trigger (`pendingAutoSendRunner_`) runs daily at 07:00.
* It:

  1. Calls `getRowsForUi_("pending")` to collect Pending rows.
  2. Builds a `head` (title + separator).
  3. Joins row `block`s separated by `CONFIG.SEPARATOR_LINE`.
  4. Sends the text using `splitAndSendToLine_()`.

You can see / manage triggers:
**Apps Script → Triggers**.

---

### Message example

For a single row, the block looks like:

```text
Date: 1/1/25
Request By: Alice
Product ID: T-001
Product Name: Amazing Tour
Platform: Website
Changed Type: Price
Impact: High
Action: Increase price from 1000 → 1200
```

Multiple blocks are separated by `------------------` (or whatever `SEPARATOR_LINE` you set), preceded by a title:

```text
Pending
------------------
Date: 1/1/25
Request By: Alice
...

------------------
Date: 2/1/25
Request By: Bob
...
```

---

## How It Works (Under the Hood)

### `getRowsForUi_(kind)`

* `kind` can be `'pending'` or `'onrequest'`.
* Flow:

  1. Get target sheet via `getTargetSheet_()`:

     * If `CONFIG.SHEET_NAME` is set, uses that.
     * Else, uses the active sheet.
  2. Read header row and build index map via `robustIndexMap_()`.
  3. Read all data rows below the header.
  4. For each row:

     * Check if the status matches `kind` using `matchStatus_()`.
     * If yes:

       * Normalize fields (dates, text).
       * Build `obj.block` (the text block for LINE).
       * Build `obj.openUrl` (link to the row in the sheet).
  5. Return an array of row objects.

---

### `sendSelected_(kind, pickedRows)`

* Called from the UI when sending selected rows.
* Steps:

  1. Build `head` using `buildHead_(kind)`, e.g.:

     ```text
     On request
     ------------------
     ```

  2. Join all `pickedRows[i].block` with `SEPARATOR_LINE`.

  3. Combine into a single large `text`.

  4. Call `splitAndSendToLine_(text, CONFIG.LINE_CHAR_LIMIT)` to send in chunks.

---

### `splitAndSendToLine_(text, limit)`

* Splits the text into <= `limit` characters per chunk.
* Attempts to split at the last newline (`\n`) before the limit.
* Sends each chunk via `sendLineOA_()`.
* Returns the number of chunks sent.

---

### `sendLineOA_(message)`

* Loads secrets (OA token and group ID) via Script Properties.
* Calls LINE Messaging API:

  * If `OA_GROUP_ID` looks like `C...`:

    * Uses `https://api.line.me/v2/bot/message/push` with `to: groupId`.
  * Otherwise:

    * Uses `https://api.line.me/v2/bot/message/broadcast`.
* Throws detailed error if the response code is not 2xx, including hints for:

  * `401` – invalid token
  * `403` – permission/plan issue
  * `404` – wrong group or bot not in the group

---

### Auto toggle

* `setAutoPendingEnabled(enabled)`:

  * Stores flag in Script Properties (`auto_pending_enabled`).
  * Installs or removes time trigger via:

    * `installPendingTrigger_()` – atHour(7).inTimezone(CONFIG.TIMEZONE)
    * `removePendingTrigger_()`

* `getAutoPendingState_()`:

  * Reads current state from Script Properties.

---

## Troubleshooting

### “Header ... not found at row X”

* Check `CONFIG.HEADER_ROW` is correct (usually 1).
* Confirm your sheet has headers that match either:

  * `CONFIG.HEADERS.*`
  * Or their `CONFIG.ALT_HEADERS.*`
* Look for:

  * Extra spaces
  * Non-breaking spaces
  * Typos

---

### No rows in “Pending” / “On request”

* Check your **Status** values.
* By default:

  * Pending mode requires `status` to contain `pending`.
  * On request mode requires `status` to contain `on request`.
* Adjust `matchStatus_()` if you use different wording.

---

### No messages sent at 07:00

* Confirm auto mode is actually enabled:

  * Open **Manually sent task** dialog and see if the pill shows **ON**.
* Check Apps Script triggers:

  * There should be a time-based trigger for `pendingAutoSendRunner_`.
* Look at execution log:

  * See if `pendingAutoSendRunner_` ran and whether it found any rows.

---

### LINE errors

Common causes:

* **401 Unauthorized**

  * Invalid or expired `OA_CHANNEL_ACCESS_TOKEN`.
  * Re-issue token and update Script Properties.

* **404 Not Found**

  * Bot not in that group.
  * Wrong `OA_GROUP_ID`.

* **403 Forbidden**

  * Plan or permissions don’t allow this type of push/broadcast.

The script tries to parse LINE’s error message and add hints.

---
