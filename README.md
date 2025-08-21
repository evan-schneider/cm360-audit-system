
# CM360 Audit System

## Overview
CM360 Audit System is a robust, enterprise-grade Google Apps Script solution for automating Campaign Manager 360 (CM360) campaign audits. It streamlines compliance, reporting, and QA by merging, flagging, and distributing audit results across multiple teams, with advanced exclusions management and automated daily summaries.

## Key Features
- **Automated Campaign Auditing**: Audits all configured CM360 campaigns on a daily schedule, supporting multiple teams and configurations.
- **Advanced Exclusions Management**: Flexible exclusions sheet for placement, site, and name fragment filtering, with real-time validation and auto-population.
- **Placement Name Lookup**: Automatically updates placement names from CM360 reports for accurate flagging and reporting.
- **Excel & Email Reporting**: Generates detailed, formatted Excel reports and sends summary/status emails to stakeholders.
- **Configurable Team Workflows**: Supports per-team thresholds, recipients, and batch processing for scalable operations.
- **Dashboard & UI**: Intuitive dashboard and configuration picker for easy management and monitoring.
- **Error Handling & Logging**: Comprehensive logging, error reporting, and status tracking for reliable operation.

## Components
- `Code.js` - Main Apps Script logic (auditing, batching, email, Drive/Sheets integration)
- `ConfigPicker.html` - Team configuration picker UI
- `Dashboard.html` - Main dashboard interface
- `appsscript.json` - Apps Script project manifest

## Setup & Installation
1. **Create a Google Apps Script Project**
	- In Google Drive, select New > More > Google Apps Script.
2. **Copy Source Files**
	- Paste the contents of `Code.js` into the script editor.
	- Add `ConfigPicker.html` and `Dashboard.html` as HTML files.
	- Replace the default `appsscript.json` with the provided manifest.
3. **Configure API Access**
	- Enable the CM360 API and Advanced Drive API in the Apps Script project.
	- Set up required OAuth scopes as prompted.
4. **Initial Sheet Setup**
	- Run the setup menu to auto-create thresholds, recipients, and exclusions sheets.
	- Populate with your team's configuration and recipient details.
5. **Deploy Triggers**
	- Use the menu to install daily batch triggers for automated audits.

## Usage Workflow
1. **Open the Dashboard**
	- Use the custom menu to access the dashboard and configuration picker.
2. **Configure Teams & Thresholds**
	- Edit the thresholds, recipients, and exclusions sheets as needed.
3. **Run or Schedule Audits**
	- Audits run automatically via triggers, or can be run manually from the menu.
4. **Review Results**
	- Receive summary emails and Excel reports; review flagged placements and campaign issues.
5. **Manage Exclusions**
	- Update the exclusions sheet to refine audit logic and reduce false positives.

## File Descriptions
- **Code.js**: Core logic for batch processing, audit execution, email/report generation, and UI integration.
- **ConfigPicker.html**: Modal dialog for selecting and managing team configurations.
- **Dashboard.html**: Sidebar/dashboard for audit status and quick actions.
- **appsscript.json**: Project manifest (defines script settings, scopes, and add-on config).

## Support & Contribution
- For support, open an issue on the project's GitHub repository or contact the maintainer.
- Contributions are welcome! Please fork the repo and submit a pull request with your improvements.

## License
This project is licensed under the MIT License. See `LICENSE` for details.

---
CM360 Audit System is designed for digital marketing teams, agencies, and enterprises seeking reliable, automated QA and compliance for Campaign Manager 360 operations.

## How to run locally (with clasp)

Prerequisites:
- Node.js and npm installed
- `clasp` installed globally or available via `npx` (we use `npx @google/clasp` in examples)
- You must be authorized with the Google account that owns or can edit the Apps Script project

1. Authenticate clasp (one-time):

```powershell
npx @google/clasp login
```

2. Pull the remote project (if you're working with the cloud project):

```powershell
npx @google/clasp pull
```

3. Push local changes to Apps Script:

```powershell
npx @google/clasp push
```

4. Create a version and (optionally) deploy an API executable if you need `clasp run`:

```powershell
npx @google/clasp version "v1"
npx @google/clasp deploy --versionNumber 1 --description "api-exec"
```

Notes:
- Running `clasp run <functionName>` requires an API-executable deployment and may require linking a user-managed GCP project for some projects; if you see a message about GCP projects in the Apps Script UI, follow the editor guidance.
- If your function uses Gmail, Drive, or other sensitive scopes, you'll need to accept OAuth consent when running from the Apps Script editor or when creating a deployment.
- Use the Apps Script editor to run functions interactively and inspect logs if you prefer not to deploy.
