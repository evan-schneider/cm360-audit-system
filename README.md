
# CM360 Audit System

CM360 Audit System is a robust Google Apps Script solution that automates Campaign Manager 360 (CM360) campaign audits and operational workflows. It aggregates daily delivery files, merges and flags issues against configurable thresholds, honors per‑team exclusions, and distributes results via formatted Excel attachments and summary emails. A simple spreadsheet UI provides configuration management, ad‑hoc run controls, and request handling.

## Overview
The system is designed for reliability and low‑maintenance operation inside Google Workspace:
- Gmail label–based ingestion of daily audit files (ZIP, CSV, XLSX)
- Deterministic merge + header detection + data cleaning
- Flexible flagging by thresholds (clicks > impressions, out‑of‑flight, pixel‑size mismatch, default ad serving)
- Per‑config recipients with optional “withhold no‑flag emails” behavior
- Exclusions by placement ID, site name, or name fragment
- Daily summary email with status across configs
- External configuration spreadsheet support, including writable INSTRUCTIONS columns

## Architecture and data flow
1. Delivery files arrive in Gmail under labels like `Daily Audits/CM360/<CONFIG>`.
2. The script pulls today’s attachments, extracts ZIPs, converts Excel (via Advanced Drive API), and stages files in Drive.
3. Files are merged into a canonical sheet; headers are detected and normalized; pixel sizes are standardized.
4. Flagging logic evaluates thresholds and exclusions; flagged rows are sorted to the top.
5. Stakeholders receive a formatted Excel export and an HTML summary for each config; a global daily summary is also sent.

## Key components
- `Code.js` — Core logic: ingestion, merge, flagging, email/reporting, configuration writers, UI/menu wiring
- `ConfigPicker.html` — Modal dialog for selecting and running audits by config
- `Dashboard.html` — Sidebar dashboard for quick access and status
- `appsscript.json` — Apps Script manifest (scopes, settings)

## Configuration
- External config sheet: set `EXTERNAL_CONFIG_SHEET_ID` in `Code.js`.
- Sheet names (defaults):
	- Thresholds: `Audit Thresholds`
	- Exclusions: `Audit Exclusions`
	- Recipients: `Audit Recipients`
	- Requests: `Audit Requests`
- INSTRUCTIONS writers: the script can write curated two‑column guidance into fixed columns on each sheet.
- Staging and recipients:
	- `STAGING_MODE`: when `Y`, emails route to `ADMIN_EMAIL`.
	- Recipients sheet supports `Withhold No‑Flag Emails` to suppress emails on clean runs.
- Gmail labels: each config in `auditConfigs` references a label such as `Daily Audits/CM360/PST01`.

## Setup
You can operate purely in the Apps Script editor, or use `clasp` for local development.

Option A — Apps Script editor
1) Create/open the Apps Script project bound to your spreadsheet.
2) Add files: paste `Code.js` and add HTML files (`ConfigPicker.html`, `Dashboard.html`).
3) Replace the manifest with `appsscript.json` (scopes will auto‑prompt).
4) In Services, enable Advanced Drive API (required for Excel→Sheets conversion). If disabled, CSV flows still work; Excel imports will be skipped.
5) Save and authorize the script on first run as prompted.

Option B — Local with `clasp`
```powershell
npm install -g @google/clasp
clasp login
git clone https://github.com/evan-schneider/cm360-audit-system.git
Set-Location .\cm360-audit-system
clasp clone <scriptId>   # or clasp pull if already linked
clasp push
```

## Operating
- On spreadsheet open, a custom menu is added with environment setup, sheet management, request tools, batch runners, and utilities.
- Use “Prepare Environment” to create Gmail labels and Drive folders for all configs.
- Use the batch runner functions (`runDailyAuditsBatch1..N`) via the menu to run audits, or install daily triggers with “Setup & Install Batch Triggers”.
- “Update External Config Instructions” writes curated guidance to INSTRUCTIONS columns in the external spreadsheet.
- A daily summary email is sent once all configs finish (lock‑protected to avoid duplicates).

## Thresholds and exclusions
- Thresholds: per‑flag minimum impressions/clicks (e.g., for clicks > impressions, out‑of‑flight, pixel mismatch, default ad serving).
- Exclusions: by placement ID, site, or name fragment; can be limited by flag type and config or applied globally.
- Recipients: control primary/CC, active status, and no‑flag email suppression per config.

## Permissions and APIs
- Built‑in services: GmailApp, DriveApp, SpreadsheetApp, HtmlService, CacheService, LockService, PropertiesService, UrlFetchApp.
- Advanced Drive API: required for Excel import (conversion). If not enabled, the script skips Excel conversion and proceeds with available CSVs.

## Troubleshooting
- Missing Gmail label: ensure filters route messages to the expected label path; the setup tool can create labels but not Gmail filters.
- Drive API not enabled: you’ll see “Skipped: Drive API not enabled”; enable Advanced Drive API in Services.
- Header not found: confirm your reports include canonical headers (e.g., Advertiser, Campaign, Site, Placement, Placement ID, Start/End Date, Creative, Impressions, Clicks).
- No files found: verify delivery occurred today and the label path matches the config.
- Email quota: the script caches and reports remaining quota; large recipient lists can exhaust daily limits.

## Version control and releases (optional)
When histories diverge (e.g., `main` vs `master`), use a clean clone and merge with unrelated histories:
```powershell
git fetch origin
git checkout -b merge-master
git merge origin/main --allow-unrelated-histories
$files = git diff --name-only --diff-filter=U
foreach ($f in $files) { git checkout --theirs -- $f; git add $f }
git commit -m "Merge origin/main into master: resolved conflicts by taking origin/main versions"
git push origin HEAD:master
```

## Contributing
Bug reports and PRs are welcome. For larger changes, open an issue to discuss approach and scope.

## License
MIT (see `LICENSE`).
