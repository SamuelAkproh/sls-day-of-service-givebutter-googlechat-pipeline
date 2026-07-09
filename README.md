# sls-day-of-service-givebutter-googlechat-pipeline
An automated data pipeline built with Google Apps Script that connects Givebutter webhooks to Google Sheets, dynamically tracking fundraising metrics, cohort leaderboards, and broadcasting weekly recaps to Google Chat.

Architecture Overview

The system operates as a lightweight serverless data pipeline structured across three distinct operational layers:

[ Givebutter Webhook ]
          │
          ▼
  ( POST Payload )
          │
          ▼
[ Apps Script: doPost() ] ──────► (Real-Time Parsing & Sanitization )
│
├───► [ Donors Tab ] ──────► [ Pivot Tables / Live Metrics Row ]
│                                        │
└───► [ Volunteers Tab ]                 ▼
                                [ sendWeeklyRecap() ]
                                          │
                                          ▼ 
                                ( Automated Triggers )
                            [ Google Chat Space Webhooks ]


1. **Ingestion Layer (`doPost`):** Intercepts real-time transactional data from incoming webhook triggers. It performs text sanitization (trimming, case normalization) and evaluates emails against a state array to skip duplicate entries.
2. **Storage & Analytics Layer (Google Sheets):** Directs records to dedicated relational data tables (`Donors` and `Volunteers`). Advanced array sorting and matrix functions parse historical row logs down to identify Top 5 individual fundraisers and return team with the most money raised (`Atlanta` vs. `Athens`).
3. **Distribution Layer (`sendWeeklyRecap`):** Formats aggregated metadata into a high-visibility engagement message and publishes it directly via Google Chat space integrations.

---

##  Features

*  **Live Leaderboard Tracking:** Dynamically ranks cohort teams by total fundraising progress and explicit contributor volume.
*  **MVP leaderboard:** Scans raw data blocks to isolate and crown individual high-performers for each competitive cohort.
*  **Top 5 Individual Leaderboards:** Dynamically groups individual supporters with multiline transaction entries to showcase overall campaign leaders.
*  **Snapshot History Archiving:** Automatically captures weekly metrics to an isolated archive table (`Snapshot History`) on execution, tracking week-over-week growth metrics dynamically without dashboard corruption.
*  **Cybersecurity First Architecture:** Formatted completely using centralized environment config patterns to safeguard production database keys and incoming endpoint routes during public deployment.

---

##  Codebase Structure

* `Code.js`: Contains core system orchestration lines, the transactional web traffic handler, mathematical calculations, and webhook routing vectors.
* `README.md`: System documentation, setup blueprints, and portfolio tracking information.

---

##  Deployment & Configuration

To deploy this architecture within your workspace infrastructure:

### 1. Spreadsheet Initialization
The destination workbook requires three specific tab environments to map database targets correctly:
* `Donors`: Columns tracking `Timestamp`, `Full Name`, `Email`, `Team Member`, and `Amount`.
* `Volunteers`: Columns tracking `Timestamp`, `Full Name`, `Email`, and `Team Member`.
* `Weekly Report`: Workspace summary calculating metrics in Row 2, driven by an open-ended Pivot Table array spanning columns `I:K`.
* `Snapshot History`: Historical log capturing tracking updates (`Week Label`, `Total Amount Raised`, `Total Volunteers Count`, `Archive Timestamp`).

### 2. Environment Variables
Initialize the `CONFIG` data block located at the head of `Code.js` with your active environment routing variables:

```javascript
var CONFIG = {
  SPREADSHEET_ID: "YOUR_GOOGLE_SPREADSHEET_ID_HERE",
  GOAL_AMOUNT: 5000.00,
  
  URL_SANDBOX_CHAT: "YOUR_GOOGLE_CHAT_SANDBOX_WEBHOOK_URL",
  URL_MAIN_CHAT: "YOUR_GOOGLE_CHAT_PRODUCTION_WEBHOOK_URL"
};
```
3. Automated Time-Driven Triggers
To schedule the weekly execution sequence, configure an internal Apps Script project trigger:

Open the Apps Script side panel and select Triggers (⏰).

Select Add Trigger.

Set the function to run to: sendWeeklyRecap.

Configure the event source type to: Time-driven.

Set the type of time-based trigger to: Weekly timer.

Set the day of the week to: Every Friday (Targeting your preferred evening communication hours).
