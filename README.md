# Event Check-In App

A Google Apps Script + React-based event management and live gate check-in system built on Google Sheets.

## Overview

This project has two main components:

**1. Apps Script Backend (Google Sheets automation)**
- Sends registration, confirmation (with QR code), reminder, and feedback emails
- Manages attendee capacity tracking and payment status dropdowns
- Exposes a Web App endpoint powering the mobile gate scanner

**2. Frontend Clients**
- `ScannerApp.html` — mobile-first gate scanner served as a Google Apps Script Web App
- `App.jsx` — standalone React check-in dashboard that talks to the Web App via fetch

---

## Architecture

```
Google Form → Sheet1 (responses)
                 ↕
         Code.js (email engine, triggers)
                 ↕
         Sheet2 (config: costs, capacity, email templates)

GateBackend.js → doGet() → ScannerApp.html (mobile gate UI)
GateBackend.js → doGet/doPost() → App.jsx (React dashboard via fetch)
```

---

## Files

| File | Purpose |
|------|---------|
| `Config.js` | All configuration loaded from Script Properties; `setupProperties()` initialiser |
| `Code.js` | Email automation engine (registration, confirmation, reminders, feedback) |
| `GateBackend.js` | Web App backend: ticket verification, check-in processing, directory API |
| `ScannerApp.html` | Mobile gate scanner UI (served by `doGet`) |
| `App.jsx` | React check-in dashboard (standalone, connects via `GOOGLE_WEB_APP_URL`) |
| `appsscript.json` | Apps Script manifest |

---

## Prerequisites

- A Google Account
- A Google Spreadsheet with:
  - **Sheet1** — Google Form responses (see column layout below)
  - **Sheet2** — Configuration table (costs, capacity, email templates)
- Node.js + `@google/clasp` (optional, for CLI-based deployment)

---

## Setup — Step by Step

### Option A: Manual (script.google.com)

1. Open your Google Spreadsheet → **Extensions → Apps Script**
2. Create the following script files and paste the contents of each:
   - `Config.js`
   - `Code.js`
   - `GateBackend.js`
3. Create an HTML file named `ScannerApp` and paste `ScannerApp.html` contents
4. Replace the default `appsscript.json` with the one from this repo
5. Run **`setupProperties()`** once to initialise Script Properties (see below)
6. Run **`installAutomatedTriggers()`** to activate form-submit and daily triggers

### Option B: Using clasp

```bash
npm install -g @google/clasp
clasp login
clasp clone <YOUR_SCRIPT_ID>
# Copy all .js and .html files into the cloned directory
clasp push
```

---

## Script Properties Reference

Run `setupProperties()` from the Apps Script editor once, then update each key in:
**Project Settings → Script Properties**

| Property Key | Description | Example |
|---|---|---|
| `ORGANIZER_EMAILS` | Comma-separated notification emails | `you@example.com,team@example.com` |
| `EVENT_SENDER_NAME` | Display name for outgoing emails | `Event Management Team` |
| `EVENT_NAME` | Event name used in emails and menus | `Community Navratri 2026` |
| `ORGANIZATION_NAME` | Organization name in email footers | `Community Association` |
| `DEFAULT_MAX_ATTENDEES` | Capacity ceiling | `250` |
| `PAYMENT_DUE_DAYS` | Days before payment reminder fires | `3` |
| `SHEET_NAME_RESPONSES` | Name of form responses sheet | `Sheet1` |
| `SHEET_NAME_TEMPLATES` | Name of config/templates sheet | `Sheet2` |
| `SHEET_NAME_CHECKINS` | Name of gate check-in log sheet | `CheckIns` |
| `TICKET_CODE_PREFIX` | Prefix for generated ticket codes | `T` |
| `QR_API_BASE_URL` | QR code generation API base URL | `https://quickchart.io/qr` |

---

## Sheet Column Layout (Sheet1 — Form Responses)

| Col | Letter | Field |
|-----|--------|-------|
| 1 | A | Timestamp |
| 2 | B | Guest Name |
| 3 | C | Email |
| 4 | D | Phone |
| 5 | E | Adults |
| 6 | F | Children |
| 7 | G | Infants |
| 8 | H | Consent |
| 12 | L | Ticket Code (auto-generated) |
| 13 | M | Payment Status (Paid / NotPaid dropdown) |
| 14 | N | Registration Email Status |
| 15 | O | Confirmation Email Status |
| 16 | P | Reminder Email Status |
| 19 | S | Adults Checked In |
| 20 | T | Children Checked In |
| 21 | U | Infants Checked In |
| 22 | V | Gate Status |

---

## Sheet2 — Configuration Table

| Row | Col B | Purpose |
|-----|-------|---------|
| 11 | Adult ticket cost | e.g. `15` |
| 12 | Child ticket cost | e.g. `10` |
| 13 | Confirmed attendee count (auto-updated) | |
| 14 | Reserved attendee count (auto-updated) | |
| 15 | Maximum capacity | e.g. `250` |

Email templates (Subject in Col A, Body in Col B):

| Row | Template |
|-----|----------|
| 2 | Confirmation email |
| 4 | Registration / payment request email |
| 7 | Feedback email |
| 17 | Payment reminder email |

Supported template tokens: `<name>`, `<TicketCode>`, `<AdultTicketCost>`, `<ChildTicketCost>`, `<TotalCost>`, `<EventName>`

---

## Web App Deployment (Gate Scanner)

1. In Apps Script editor: **Deploy → New Deployment → Web App**
2. Set **Execute as**: `Me`
3. Set **Who has access**: `Anyone` (or `Anyone within your organisation`)
4. Copy the deployment URL

For the React dashboard (`App.jsx`), paste the URL into:
```js
const GOOGLE_WEB_APP_URL = "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec";
```

---

## Canceling a Registration

Strike through the entire row in Sheet1 (select row → Format → Strikethrough).
All functions automatically skip strikethrough rows. The gate scanner will show **ENTRY DENIED** for any canceled ticket.

---

## License

MIT — see [LICENSE](LICENSE)
