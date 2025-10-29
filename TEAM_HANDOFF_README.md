# CM360 Daily Audit System - Team Handoff Documentation

**Last Updated:** October 29, 2025  
**Prepared By:** Evan Schneider (evschneider@horizonmedia.com)  
**Repository:** https://github.com/evan-schneider/cm360-audit-system  
**Branch:** integrate-2025-09-29

---

## Table of Contents

1. [System Overview](#system-overview)
2. [Critical First Steps](#critical-first-steps)
3. [Key Spreadsheets](#key-spreadsheets)
4. [Admin Controls Menu Reference](#admin-controls-menu-reference)
5. [External Config Spreadsheet Guide](#external-config-spreadsheet-guide)
6. [Daily Operations](#daily-operations)
7. [Configuration Management](#configuration-management)
8. [Automated Processes](#automated-processes)
9. [Troubleshooting](#troubleshooting)
10. [Manual Operations](#manual-operations)
11. [Maintenance Tasks](#maintenance-tasks)
12. [Technical Architecture](#technical-architecture)
13. [Common Scenarios](#common-scenarios)
14. [Emergency Contacts](#emergency-contacts)

---

## System Overview

### What This System Does

The CM360 Daily Audit System automatically:
- Fetches daily CM360 reports from Gmail
- Merges multiple report files into consolidated spreadsheets
- Flags discrepancies based on configurable thresholds
- Sends email alerts to designated recipients
- Archives reports in Google Drive
- Cleans up old files and emails (60-day and 90-day retention)

### Key Components

- **Admin Spreadsheet** (bound to Apps Script): Configuration interface
- **External Config Spreadsheet**: Centralized configuration storage
- **Apps Script Project**: Core automation logic (Code.js)
- **Gmail Labels**: Report routing (Daily Audits/CM360/*)
- **Drive Folders**: Report storage and archival
- **Deletion Log**: Audit trail for all automated deletions

---

## Critical First Steps

### System Ownership Structure

**IMPORTANT:** This system runs under the shared service account **platformsolutionshmi@gmail.com**, NOT under Evan's personal account.

**What this means:**
- ✅ Triggers will continue working after Evan leaves
- ✅ CM360 report emails go to platformsolutionshmi@gmail.com
- ✅ All Drive files stored in platformsolutionshmi@gmail.com Drive
- ✅ No disruption to automated processes
- ✅ No need to reinstall triggers or transfer spreadsheet ownership

**What DOES need to be updated:** ADMIN_EMAIL for error notifications and staging mode

**Critical:** ADMIN_EMAIL is used for ALL system communications when staging mode is active, plus error/diagnostic emails in production.

---

### 1. Update Admin Email (ONLY Required Change)

The current admin email is **evschneider@horizonmedia.com**. This should be changed before Evan's departure.

**Where to Update:**

#### Code.js (Line ~29)
1. Open the Admin Spreadsheet
2. Go to **Extensions > Apps Script**
3. Open **Code.js** file
4. Find line ~29 in the ADMIN_EMAIL constant:
   ```javascript
   return 'evschneider@horizonmedia.com';
   ```
5. Change to:
   ```javascript
   return 'newadmin@horizonmedia.com';
   ```
6. **Save** the file (Ctrl+S / Cmd+S)
7. **Deploy**: Click the blue Deploy button or use existing deployment

**What uses ADMIN_EMAIL:**
- **Staging Mode:** ALL audit emails redirected to ADMIN_EMAIL when staging mode enabled
- **Error alerts** and failure notifications
- **Health check** reports (daily at 5:04 AM EST)
- **Watchdog alerts** (timeout/stuck batch notifications)
- **System diagnostic** emails
- **Admin BCC** on production audit emails (optional)
- **Summary email** (included in distribution list)
- **Test emails** from Admin Controls menu

**Note:** In STAGING mode, ADMIN_EMAIL receives ALL emails (audit + system). In PRODUCTION mode, ADMIN_EMAIL only receives system/error emails.

---

### 2. Understand Service Account Setup

**platformsolutionshmi@gmail.com is the system owner:**

**What it owns:**
- Apps Script project and all triggers
- Admin Spreadsheet (bound to Apps Script)
- External Config Spreadsheet
- All Drive folders (Project Log Files/CM360 Daily Audits/*)
- All Gmail labels (Daily Audits/CM360/*)
- Deletion Log spreadsheet

**Access control:**
- Evan currently has access to platformsolutionshmi@gmail.com
- Before departure: Evan transfers platformsolutionshmi@gmail.com credentials to team
- Team gains access to service account (password, 2FA recovery)
- Recommended: Multiple team members should have recovery access

**Why this design:**
- Prevents disruption when individuals leave
- Centralizes system ownership
- No trigger reinstallation needed
- Cleaner access management

---

### 3. Pre-Departure Handoff Checklist

**For Evan to complete before leaving:**

- [ ] Update ADMIN_EMAIL in Code.js (line ~29)
- [ ] Deploy updated Code.js to Apps Script
- [ ] Transfer platformsolutionshmi@gmail.com credentials to team:
  - [ ] Password (via secure method)
  - [ ] 2FA backup codes
  - [ ] Recovery email/phone access
- [ ] Update summary email distribution list (Code.js line ~1061) if needed
- [ ] Document platformsolutionshmi@gmail.com access with IT/Leadership
- [ ] Verify new admin can:
  - [ ] Access platformsolutionshmi@gmail.com Gmail
  - [ ] Open Admin Spreadsheet
  - [ ] Open Apps Script editor
  - [ ] View Drive folders
- [ ] Test ADMIN_EMAIL receives error notifications
- [ ] Remove Evan's direct access to spreadsheets (if separate from service account)

**For new team lead:**

- [ ] Verify receipt of error notifications at new ADMIN_EMAIL
- [ ] Bookmark Admin Spreadsheet URL
- [ ] Bookmark External Config Spreadsheet URL
- [ ] Bookmark this README
- [ ] Test manual operations (run single audit)
- [ ] Monitor first week of automated runs

---

### 4. Emergency Access

**If platformsolutionshmi@gmail.com credentials lost:**

**Immediate actions:**
1. Contact IT/DevOps to reset password
2. Check with Evan (if still available) for recovery info
3. Check internal password manager/vault
4. Review Google Workspace admin console (if IT has access)

**Worst case recovery:**
- System will continue running (triggers still work)
- But cannot modify configurations or access Drive files
- IT must reset account or create new service account
- May require spreadsheet ownership transfer
- Contact Platform Solutions leadership immediately

---

## Key Spreadsheets

### Admin Spreadsheet (Bound Script)

**Owner:** platformsolutionshmi@gmail.com  
**Location:** Extensions > Apps Script attached to this spreadsheet

**Critical Sheets:**
- Audit Recipients - Email distribution lists per config
- Audit Thresholds - Flagging criteria per config
- Audit Exclusions - Items to ignore per config

**Access:** 
- Shared with team members (edit access)
- Must have access to platformsolutionshmi@gmail.com to modify Apps Script
- All team members can view/edit sheets

### External Config Spreadsheet

**Owner:** platformsolutionshmi@gmail.com  
**ID:** 1-566gqkyZRNDeNtXWUjKDB_H8A9XbhCu8zL-uaZdGT8  
**URL:** https://docs.google.com/spreadsheets/d/1-566gqkyZRNDeNtXWUjKDB_H8A9XbhCu8zL-uaZdGT8

**Purpose:** Centralized configuration that syncs TO the Admin spreadsheet nightly

**Same Sheets:**
- Audit Recipients
- Audit Thresholds
- Audit Exclusions

**Best Practice:** Edit this spreadsheet, not the Admin spreadsheet. Changes sync automatically at 2:00 AM EST daily.

**Access:**
- Shared with team members who need to edit configurations
- No Apps Script access required

### CM360 Deletion Log

**Owner:** platformsolutionshmi@gmail.com  
**Location:** Google Drive at Project Log Files > CM360 Daily Audits > Deletion Log  
**File Name:** CM360 Deletion Log

**Sheets:**
- Temp Daily Reports - Temporary files deleted after 60 days
- Merged Reports - Merged files deleted after 60 days
- Gmail Emails - Email threads deleted after 90 days

**Purpose:** Audit trail for all automated deletions

---

## Admin Controls Menu Reference

### How to Access

The **Admin Controls** menu appears at the top of the Admin Spreadsheet when you open it.

**Location:** Top menu bar > **Admin Controls**

If you don't see the menu:
1. Refresh the page (Ctrl+R / Cmd+R)
2. If still missing: Admin Controls > **⚙️ Prepare Environment**
3. Last resort: Extensions > Apps Script > Run > `forceCreateMenu`

### Menu Categories

The Admin Controls menu is organized into functional groups:

---

### Setup & Configuration

#### ⚙️ Prepare Environment
**Purpose:** Initial setup for new configurations  
**What it does:**
- Creates missing Gmail labels for all configs
- Creates missing Drive folders
- Summarizes labels without recent mail
- Verifies folder structure

**When to use:**
- After adding a new client configuration
- After system restore/migration
- When troubleshooting "folder not found" errors

---

#### 📄 Thresholds (create/open)
**Purpose:** Manage threshold configuration sheet  
**What it does:**
- Opens existing Audit Thresholds sheet
- Creates sheet if missing
- Applies formatting and data validations
- Sets up column headers

**Sheet columns:**
- Config Name - Unique identifier
- Various threshold fields (Impressions, Clicks, etc.)
- Tolerance percentages for flagging discrepancies

---

#### 🚫 Exclusions (create/open)
**Purpose:** Manage exclusions configuration sheet  
**What it does:**
- Opens existing Audit Exclusions sheet
- Creates sheet if missing
- Protects Placement Name column (read-only)
- Applies data validations

**Sheet columns:**
- Config Name - Which audit this exclusion applies to
- Placement ID - ID from CM360
- Placement Name - Auto-populated (protected)
- Match Mode - Exact/Contains/Regex
- Reason - Why this is excluded

---

#### 📧 Recipients (create/open)
**Purpose:** Manage email recipients sheet  
**What it does:**
- Opens existing Audit Recipients sheet
- Creates sheet if missing
- Sets up email distribution columns
- Applies formatting

**Sheet columns:**
- Config Name - Unique identifier
- To - Primary recipients (comma-separated)
- CC - Carbon copy recipients
- Gmail Label - Where to find reports
- Withhold Mode - Silent/Normal (controls email sending)

---

#### 🧩 CM360 Config Builder…
**Purpose:** Guided wizard for adding new configurations  
**What it does:**
- Opens interactive sidebar
- Guides through config creation
- Provides next steps checklist
- Shows admin hints

**Use this when:** Adding a brand new client to the system

---

### External Config Sync

#### 📤 Sync TO External Config
**Purpose:** Push Admin spreadsheet changes to External Config  
**What it does:**
- Copies Recipients, Thresholds, Exclusions, Requests sheets
- FROM: Admin Spreadsheet
- TO: External Config Spreadsheet
- Preserves formatting and validations

**When to use:**
- Rarely needed (External Config is the source of truth)
- Only if you made changes in Admin and want to preserve them
- Emergency backup scenario

**Warning:** Overwrites External Config - use with caution!

---

#### 📥 Sync FROM External Config
**Purpose:** Pull latest configuration from External Config  
**What it does:**
- Copies Recipients, Thresholds, Exclusions, Requests sheets
- FROM: External Config Spreadsheet
- TO: Admin Spreadsheet
- Updates configurations used by audit runs
- Preserves formatting, validations, dimensions

**When to use:**
- After editing External Config (to apply changes immediately)
- When testing configuration changes
- To force-sync before automated nightly sync

**This runs automatically at 2:00 AM EST daily**

---

### Audit Requests

#### 📝 Create Audit Request...
**Purpose:** Submit one-off audit request  
**What it does:**
- Opens config picker sidebar
- Adds request to Audit Requests sheet
- Request gets processed on next trigger

**Use case:** Run audit for specific date/config outside normal schedule

---

#### ▶️ Process Audit Requests
**Purpose:** Execute pending one-off requests manually  
**What it does:**
- Reads Audit Requests sheet
- Processes unexecuted requests
- Updates request status
- Sends audit emails

**When to use:** Process requests immediately without waiting for trigger

---

#### 🛠️ Fix Audit Requests Sheet
**Purpose:** Repair corrupted Requests sheet  
**What it does:**
- Reapplies headers
- Fixes data validations
- Repairs formatting

**When to use:** If Requests sheet becomes corrupted or malformed

---

### Tools & Diagnostics

#### 🔁 Update Placement Names
**Purpose:** Auto-populate placement names in Exclusions  
**What it does:**
- Reads latest merged audit reports
- Finds Placement IDs
- Fills Placement Name column in EXTERNAL Exclusions sheet
- Only updates rows with IDs but missing names

**When to use:**
- After adding new Placement IDs to Exclusions
- Monthly maintenance to keep names current

---

#### 🔐 Check Authorization
**Purpose:** Verify script permissions  
**What it does:**
- Tests Gmail access
- Tests Drive access
- Tests Spreadsheet access
- Sends result email to current user

**When to use:**
- After new admin takes over
- Troubleshooting "authorization required" errors
- Verifying scope grants

---

#### 🧾 Validate Configs
**Purpose:** Check configuration integrity  
**What it does:**
- Validates all audit configs
- Checks for missing Gmail labels
- Checks for duplicate config names
- Logs findings to console

**When to use:**
- After bulk config changes
- Troubleshooting audit failures
- Monthly maintenance

---

#### ⏱️ Install All Triggers
**Purpose:** Reinstall automation triggers  
**What it does:**
- Deletes existing triggers (except batch stubs)
- Creates new triggers:
  - Daily audit batches (8-9 AM EST)
  - Nightly maintenance (2:24 AM EST)
  - External sync (2:00 AM EST)
  - Daily summary (9:25 AM EST)
  - Watchdog monitoring
  - Health checks

**When to use:**
- **CRITICAL:** If triggers accidentally deleted or disabled
- After service account authorization expires (rare)
- If system stops running and logs show "trigger not found"
- After major Google Workspace changes

**Note:** Under normal circumstances, triggers do NOT need reinstallation when team members change, since they belong to platformsolutionshmi@gmail.com service account.

---

#### 🔄 Sync Delivery Mode Now
**Purpose:** Update delivery mode indicator  
**What it does:**
- Reads STAGING_MODE from Script Properties
- Updates "Delivery Mode" instruction line
- Updates both Admin and External Recipients sheets

**When to use:**
- After changing STAGING_MODE property
- To verify current mode

---

#### 📮 Debug Email Delivery
**Purpose:** Check email system status  
**What it does:**
- Logs delivery mode (STAGING/PRODUCTION)
- Shows admin email
- Shows remaining daily email quota (limit: ~1,500)

**When to use:**
- Troubleshooting email sending issues
- Checking if quota exhausted

---

#### ✉️ Send Test Admin Email
**Purpose:** Verify email plumbing works  
**What it does:**
- Sends simple test message to ADMIN_EMAIL
- Confirms email sending functional

**When to use:**
- After admin email change
- Verifying email delivery works
- Testing after authorization changes

---

#### 👀 Preview Daily Summary
**Purpose:** See daily summary without sending  
**What it does:**
- Builds daily summary email HTML
- Shows preview in modal dialog
- Does NOT send email

**When to use:**
- Checking what would be in summary
- Verifying summary formatting
- Debugging summary content

---

#### 🔎 Silent Withhold Check…
**Purpose:** Test email withhold logic  
**What it does:**
- Pick a config
- Simulates audit email decision
- Shows whether email would be sent or withheld
- Does NOT run actual audit or send emails

**When to use:**
- Testing withhold mode settings
- Verifying silent mode behavior
- Debugging why emails not sending

---

#### 🩺 Run Health Check (Admin)
**Purpose:** System health diagnostic  
**What it does:**
- Fast read-only checks:
  - Config validity
  - Gmail label existence
  - Drive folder accessibility
  - Trigger status
  - Email quota
- Emails report to admin

**When to use:**
- Daily/weekly proactive monitoring
- Before deployments
- After system changes
- Troubleshooting

**Runs automatically at 5:04 AM EST daily**

---

#### 🧪 Test Thresholds…
**Purpose:** Debug threshold flagging  
**What it does:**
- Pick a config
- Runs full audit
- Logs detailed threshold evaluation for each row
- Shows what was flagged and why

**When to use:**
- Debugging why items flagged/not flagged
- Tuning threshold values
- Understanding threshold logic

---

### Manual Audit Execution

#### 🧪 [TEST] Run Batch or Config
**Purpose:** Test batch execution  
**What it does:**
- Opens picker: select batch (1-12) or specific config
- Runs selected batch/config immediately
- Use for testing without waiting for scheduled triggers

**When to use:**
- Testing new configurations
- Debugging batch issues
- Verifying fixes

---

#### ▶️ Run Audit for...
**Purpose:** Run single config on demand  
**What it does:**
- Opens config picker
- Runs full audit for selected config
- Fetches reports, merges, flags, sends emails

**When to use:**
- One-off audit runs
- Re-running failed audit
- Testing configuration changes

---

### Monitoring & Access

#### 📦 Batch Assignments
**Purpose:** View batch distribution  
**What it does:**
- Shows which configs assigned to each batch (1-12)
- Displays batch balance
- Modal dialog view

**When to use:**
- Understanding batch distribution
- Troubleshooting why config not running
- Verifying batch rebalancing

---

#### ⏰ Install Health Check Trigger
**Purpose:** Enable daily health reports  
**What it does:**
- Installs daily trigger (5:04 AM EST)
- Runs health check and emails admin
- Only needed if trigger deleted

**When to use:** After trigger deletion or new admin setup

---

#### 🛡️ Install Audit Watchdog Trigger
**Purpose:** Enable timeout monitoring  
**What it does:**
- Installs 3-hour interval trigger
- Detects stuck/timed-out batch runs
- Sends alert emails

**When to use:** After trigger deletion or new admin setup

---

#### ℹ️ About Admin Controls…
**Purpose:** Help documentation  
**What it does:**
- Shows this reference guide
- Lists all menu items with descriptions

**When to use:** Quick reference for menu functions

---

## External Config Spreadsheet Guide

### Overview

**Spreadsheet ID:** `1-566gqkyZRNDeNtXWUjKDB_H8A9XbhCu8zL-uaZdGT8`  
**Direct URL:** https://docs.google.com/spreadsheets/d/1-566gqkyZRNDeNtXWUjKDB_H8A9XbhCu8zL-uaZdGT8

**Purpose:** Centralized configuration storage that multiple people can edit without needing Apps Script access.

**Key Principle:** External Config is the **SOURCE OF TRUTH** for all configuration.

### Why External Config Exists

**Problem:** Not everyone needs/should have Apps Script access, but many people need to edit configurations.

**Solution:** 
- Configuration stored in separate spreadsheet (External Config)
- Anyone with edit access can change configs
- Changes sync automatically to Admin Spreadsheet nightly
- Audit system reads from Admin Spreadsheet (synced copy)

**Benefit:** 
- Team members can manage configs without Apps Script permissions
- Reduces risk of accidental code changes
- Cleaner access control

---

### How Configuration Updates Work

#### The Update Flow

```
1. You edit External Config Spreadsheet
   ↓
2. Wait for nightly sync (2:00 AM EST automatic)
   OR
   Run sync manually (Admin Controls > 📥 Sync FROM External Config)
   ↓
3. Changes copied to Admin Spreadsheet
   ↓
4. Next audit run (8-9 AM EST) uses updated configuration
```

#### Timing Examples

**Example 1: Automatic Sync**
- 3:00 PM Monday: You add new threshold to External Config
- 2:00 AM Tuesday: Automatic sync copies change to Admin
- 8:00 AM Tuesday: Morning audits use new threshold ✅

**Example 2: Immediate Sync (Manual)**
- 3:00 PM Monday: You add new threshold to External Config
- 3:05 PM Monday: You run Admin Controls > 📥 Sync FROM External Config
- 3:10 PM Monday: You run test audit - uses new threshold ✅
- 2:00 AM Tuesday: Automatic sync runs (no changes, already synced)
- 8:00 AM Tuesday: Morning audits continue using threshold ✅

**Example 3: Same-Day Updates**
- 7:00 AM Tuesday: You update External Config
- 8:00 AM Tuesday: Morning audits run - **uses OLD config** ❌ (sync hasn't run)
- 2:00 AM Wednesday: Automatic sync copies change
- 8:00 AM Wednesday: Morning audits use NEW config ✅

**Solution for same-day:** Run manual sync immediately after editing

---

### External Config Sheets

The External Config Spreadsheet contains 4 configuration sheets:

#### 1. Audit Recipients

**Purpose:** Define who receives audit emails for each configuration

**Columns:**

| Column | Required | Description | Example |
|--------|----------|-------------|---------|
| Config Name | ✅ Yes | Unique identifier (no spaces recommended) | `ACMECorp` |
| To | ✅ Yes | Primary recipients (comma-separated emails) | `client@acme.com,am@horizonmedia.com` |
| CC | No | Carbon copy recipients | `manager@acme.com` |
| Gmail Label | ✅ Yes | Where to find reports in Gmail | `Daily Audits/CM360/ACMECorp` |
| Withhold Mode | No | `Silent` or leave blank for normal | `Silent` |

**Special Features:**

**Delivery Mode Instruction (Row 1):**
- First row shows current delivery mode
- Updates automatically when mode changes
- Format: `🟢 PRODUCTION MODE` or `🟡 STAGING MODE`

**Withhold Mode (Silent):**
- Set to `Silent` to suppress emails unless discrepancies found
- Blank or any other value = Normal (always send)
- Use for high-volume configs where "all clear" emails not needed

**Best Practices:**
- Always include at least one Horizon email in To/CC
- Test new configs with single recipient first
- Use Config Names without spaces (easier debugging)

---

#### 2. Audit Thresholds

**Purpose:** Define minimum volume thresholds for flagging issues

**Columns:**

| Column | Description | Example |
|--------|-------------|---------|
| Config Name | Must match Recipients sheet | `ACMECorp` |
| Flag Type | Type of issue to flag | `clicks_greater_than_impressions` |
| Min Impressions | Minimum impressions required to flag | `100` |
| Min Clicks | Minimum clicks required to flag | `20` |
| Active | TRUE to enable, FALSE to disable | `TRUE` |

**Flag Types:**
- `clicks_greater_than_impressions` - Flags rows where clicks > impressions (data quality issue)
- `out_of_flight_dates` - Flags rows with dates outside placement flight dates
- `pixel_size_mismatch` - Flags rows where placement pixel ≠ creative pixel
- `default_ad_serving` - Flags rows with "default" ad type

**Each config needs 4 threshold rows** (one per flag type) for complete coverage.

**How Thresholds Work:**

Thresholds are **minimum volume requirements**, not percentage tolerances. A row is only flagged if:
1. The issue exists (e.g., clicks > impressions, pixel mismatch, etc.)
2. The row meets the minimum volume threshold

**Minimum Volume Threshold Logic:**
- System checks whether **clicks or impressions** is higher
- Uses the higher metric's threshold (minClicks if clicks > impressions, minImpressions if impressions > clicks)
- Only flags the row if the higher volume meets/exceeds the threshold

**Example 1: Clicks > Impressions (High Volume)**
- Row data: 1,000 clicks, 500 impressions
- Threshold: minClicks = 100, minImpressions = 50
- Evaluation: Clicks (1,000) > Impressions (500), so check minClicks threshold
- Result: **FLAGGED** (1,000 >= 100 threshold met, and clicks > impressions is an issue)

**Example 2: Clicks > Impressions (Low Volume - Not Flagged)**
- Row data: 50 clicks, 25 impressions  
- Threshold: minClicks = 100, minImpressions = 50
- Evaluation: Clicks (50) > Impressions (25), so check minClicks threshold
- Result: **NOT FLAGGED** (50 < 100 threshold not met, even though clicks > impressions)

**Example 3: Pixel Mismatch (High Volume)**
- Row data: 800 impressions, 20 clicks, placement pixel = 300x250, creative pixel = 300x600
- Threshold: minImpressions = 100, minClicks = 10
- Evaluation: Impressions (800) > Clicks (20), so check minImpressions threshold
- Result: **FLAGGED** (800 >= 100 threshold met, and pixel mismatch exists)

**Purpose of Volume Thresholds:**
- Prevents flagging low-volume noise (e.g., 2 clicks, 1 impression)
- Focuses attention on issues that impact significant traffic
- Different configs may need different thresholds based on typical volumes

**Best Practices:**
- Set thresholds based on what volume level matters for your client
- Lower thresholds = more sensitive (flag smaller issues)
- Higher thresholds = less sensitive (only flag high-volume issues)
- Typical range: minImpressions 50-500, minClicks 10-100
- Adjust based on client's typical daily volumes

**Common Settings:**
- High-volume configs: minImpressions 500, minClicks 50
- Medium-volume configs: minImpressions 100, minClicks 20
- Low-volume configs: minImpressions 50, minClicks 10

---

#### 3. Audit Exclusions

**Purpose:** Define items to ignore in audits (known discrepancies, test placements, etc.)

**Columns:**

| Column | Required | Description | Example |
|--------|----------|-------------|---------|
| Config Name | ✅ Yes | Which audit this applies to | `ACMECorp` |
| Placement ID | ✅ Yes | CM360 Placement ID | `123456789` |
| Placement Name | ⚠️ Auto-filled | Populated by system (read-only) | `Homepage Banner` |
| Match Mode | No | `Exact`, `Contains`, `Regex` | `Exact` |
| Reason | Recommended | Why excluded (for documentation) | `Test placement` |

**Special Features:**

**Auto-Population of Placement Names:**
- Add Placement ID
- Leave Placement Name blank
- Run: Admin Controls > 🔁 Update Placement Names
- System reads latest audit reports
- Fills in Placement Name automatically (function runs nightly)

**Match Modes:**
- `Exact` - Must match exactly (default)
- `Contains` - Partial match (e.g., "Test" matches "Test Banner 123")
- `Regex` - Regular expression pattern (advanced)

**Use Cases:**
- Test placements (not live)
- Known discrepancies (can't fix)
- Rotational creatives (expected differences)
- Seasonal campaigns (temporary)

**Best Practices:**
- Always document Reason
- Review exclusions quarterly
- Remove obsolete exclusions
- Use specific IDs, not broad patterns

---

#### 4. Audit Requests

**Purpose:** One-off audit requests outside normal schedule

**Columns:**

| Column | Description | Example |
|--------|-------------|---------|
| Request Date | When request was made | `2025-10-29` |
| Config Name | Which config to audit | `ACMECorp` |
| Audit Date | Date of data to audit | `2025-10-28` |
| Status | Pending/Completed/Failed | `Pending` |
| Requested By | Who submitted request | `evschneider@horizonmedia.com` |

**How It Works:**
1. Add row to Requests sheet (or use Admin Controls > 📝 Create Audit Request)
2. Set Status to `Pending`
3. System processes on next trigger OR run Admin Controls > ▶️ Process Audit Requests
4. Status updates to `Completed` or `Failed`

**Use Cases:**
- Re-run failed audit
- Audit specific past date
- On-demand client request

---

### External Config Best Practices

#### Editing Workflow

**Recommended:**
1. ✅ Edit External Config Spreadsheet
2. ✅ Run manual sync if urgent (Admin Controls > 📥 Sync FROM)
3. ✅ Test with single config before broad deployment
4. ✅ Monitor first few audit runs after change

**Not Recommended:**
1. ❌ Editing Admin Spreadsheet directly (changes overwritten at 2 AM)
2. ❌ Making changes at 7-8 AM without manual sync (won't apply same day)
3. ❌ Deleting rows (breaks array formulas - clear cells instead)

#### Change Management

**For Small Changes** (single threshold adjustment):
- Edit External Config
- Run manual sync
- Test with single config
- Monitor next audit

**For Large Changes** (new client, restructuring):
- Edit External Config
- Enable STAGING_MODE (redirects all emails to admin)
- Run manual sync
- Test thoroughly
- Disable STAGING_MODE
- Monitor closely

**For Emergency Changes**:
- Edit External Config
- Run manual sync immediately
- Test relevant config
- Document in change log

---

### Access Control

**Who needs access:**
- **Edit Access:** AdOps team, Platform Solutions, managers who update configs
- **View Access:** Leadership, auditors, anyone who needs visibility

**Owner:** platformsolutionshmi@gmail.com (system service account)

**Sharing:**
- Already shared with team members
- Add new team members: File > Share > Add people
- Grant "Editor" role for config management
- No ownership transfer needed when people leave

---

### Backup & Recovery

**Automatic Backups:**
- Google Sheets has version history (File > Version history)
- Can restore previous versions if needed

**Manual Backup:**
- File > Download > Microsoft Excel (.xlsx)
- Store in secure location
- Do monthly for critical configs

**Recovery Scenarios:**

**If External Config accidentally deleted:**
1. Check Google Drive trash (in platformsolutionshmi@gmail.com Drive)
2. Restore from trash if within 30 days
3. If not recoverable, use Admin Spreadsheet as temporary source
4. Create new External Config Spreadsheet (as platformsolutionshmi@gmail.com)
5. Update EXTERNAL_CONFIG_SHEET_ID in Code.js Script Properties
6. Run Admin Controls > 📤 Sync TO External Config

**If External Config corrupted:**
1. File > Version history > See version history
2. Restore last known good version
3. Run manual sync (Admin Controls > 📥 Sync FROM External Config)
4. Verify configs

**Note:** All file operations happen in platformsolutionshmi@gmail.com account, so recovery requires access to that account.

---

### Troubleshooting External Config

**Problem:** Changes not applying to audits

**Solutions:**
1. Verify you edited External Config (not Admin)
2. Check sync ran: Admin Controls > 📥 Sync FROM External Config
3. Check audit used synced config (timing issue if before 2 AM)
4. Verify Config Name matches exactly (case-sensitive)

**Problem:** "External Config not found" error

**Solutions:**
1. Verify EXTERNAL_CONFIG_SHEET_ID in Script Properties
2. Check spreadsheet still exists
3. Verify script has access to spreadsheet
4. Check spreadsheet not deleted/renamed

**Problem:** Sync taking very long or timing out

**Solutions:**
1. Check spreadsheet size (very large = slow sync)
2. Remove obsolete configs
3. Reduce complexity (heavy formulas slow sync)

---

## Daily Operations

### Normal Daily Flow

**8:00 AM EST** (Morning audit run)
1. All 12 batch triggers fire simultaneously (`runDailyAuditsBatch1` through `runDailyAuditsBatch12`)
2. Each batch processes 2 configs independently
3. For each config in batch:
   - Fetches CM360 reports from Gmail labels
   - Merges multiple report files into consolidated spreadsheet
   - Flags issues based on volume thresholds (clicks > impressions, pixel mismatches, etc.)
   - Sends individual audit email to configured recipients
4. Accumulates results in cache for daily summary
5. Typical execution: All 12 batches complete within 8:00-9:00 AM EST window

**9:25 AM EST**
1. `sendDailySummaryFailover` trigger fires
2. Sends consolidated summary email to:
   - evschneider@horizonmedia.com (UPDATE THIS!)
   - bmuller@horizonmedia.com
   - bkaufman@horizonmedia.com
   - ewarburton@horizonmedia.com

**2:24 AM EST** (Next day - Nightly maintenance)
1. `runNightlyMaintenance` performs comprehensive housekeeping:
   - **Syncs External Config → Admin spreadsheet** (Recipients, Thresholds, Exclusions)
   - Rebalances audit batches for even distribution
   - Updates placement names in Exclusions from latest reports
   - Cleans up Drive files older than 60 days
   - Deletes Gmail emails older than 90 days
   - Clears daily script properties (cache reset)
2. Typical execution: ~6 minutes

**5:04 AM EST**
1. `runHealthCheckAndEmail` performs system diagnostics
2. Checks config validity, Gmail labels, Drive folders, trigger status, email quota
3. Sends health report to ADMIN_EMAIL

**Continuous (Every 1-4 hours)**
- **Every hour:** `forwardGASFailureNotificationsToAdmin` forwards script failures to ADMIN_EMAIL
- **Every 3 hours:** `auditWatchdogCheck` detects stuck/timed-out batch runs
- **Every 3 hours:** `runDeliveryModeSync` updates staging/production mode indicators
- **Every 4 hours:** `autoFixRequestsSheet_` maintains Audit Requests sheet integrity

### What to Monitor Daily

**Morning Checks (9:30 AM EST):**
- Check your inbox for summary email (subject: "CM360 Daily Audit Summary") - arrives ~9:25-9:30 AM
- Verify no error alerts from system
- Spot-check a few audit emails were received by clients (sent 8:00-9:00 AM)

**If Summary Email Missing:**
- Check spam/trash folders
- Run manually: previewDailySummaryNow() to see if data exists
- Check Apps Script logs: Extensions > Apps Script > Executions

**If Audit Emails Missing:**
- Check Gmail labels: Daily Audits/CM360/[ConfigName]
- Verify reports were delivered from CM360
- Check for error emails from system

---

## Configuration Management

### Adding a New Audit Configuration

1. **External Config Spreadsheet** > Audit Recipients sheet
2. Add new row with:
   - Config Name - Unique identifier (no spaces recommended)
   - To - Primary recipients (comma-separated emails)
   - CC - Optional CC recipients
   - Gmail Label - e.g., Daily Audits/CM360/NewClient

3. **External Config Spreadsheet** > Audit Thresholds sheet
4. Add matching row with same Config Name:
   - Set threshold values for flagging
   - Common thresholds: Impressions, Clicks, Pixel Size mismatches

5. **External Config Spreadsheet** > Audit Exclusions (if needed)
6. Add rows for items to ignore in this config

7. **Wait for nightly sync** (2:00 AM) OR **Force sync now:**
   - Open Admin Spreadsheet
   - Extensions > Apps Script
   - Run: syncFromExternalConfig()

8. **Create Gmail label** if it doesn't exist:
   - Gmail > Settings > Labels > Create new label
   - Name: Daily Audits/CM360/NewClient

9. **Ensure CM360 reports are labeled correctly** in Gmail

### Modifying Existing Configuration

**Best Practice:** Edit External Config Spreadsheet, not Admin Spreadsheet

**To Apply Changes Immediately:**
1. Edit External Config Spreadsheet
2. Open Admin Spreadsheet > Extensions > Apps Script
3. Run: syncFromExternalConfig()
4. Changes take effect on next audit run

**Changes Apply Automatically:** After 2:24 AM EST nightly maintenance (includes external config sync)

### Removing an Audit Configuration

1. **Do NOT delete rows** - this can break array formulas
2. Instead, **clear the Gmail Label** column in Recipients sheet
3. System will skip configs with blank labels
4. After 30 days of inactivity, safe to delete row entirely

---

## Automated Processes

### Trigger Schedule

| Trigger | Function | Frequency | Purpose |
|---------|----------|-----------|---------|
| Morning Audits | `runDailyAuditsBatch1-12` | Daily 8:00 AM EST | 12 batches run simultaneously, each processes 2 configs |
| Daily Summary | `sendDailySummaryFailover` | Daily 9:25 AM EST | Send consolidated summary email |
| Nightly Maintenance | `runNightlyMaintenance` | Daily 2:24 AM EST | External Config sync, cleanup, rebalancing, email deletion |
| Health Check | `runHealthCheckAndEmail` | Daily 5:04 AM EST | System diagnostics report to ADMIN_EMAIL |
| Watchdog | `auditWatchdogCheck` | Every 3 hours | Detect stuck/timed-out batch runs |
| Delivery Mode Sync | `runDeliveryModeSync` | Every 3 hours | Update staging/production mode indicators |
| Requests Sheet Fix | `autoFixRequestsSheet_` | Every 4 hours | Maintain Audit Requests sheet integrity |
| Failure Forwarder | `forwardGASFailureNotificationsToAdmin` | Every 1 hour | Forward script failures to ADMIN_EMAIL |

### Staging Mode: Complete Guide

#### What is Staging Mode?

**Purpose:**
- Test environment that prevents emails from reaching stakeholders during testing
- Routes ALL system emails (audit reports, error notifications, summaries, health checks, etc.) to ADMIN_EMAIL only
- Allows safe testing of configuration changes, code deployments, and troubleshooting

**Critical Behavior:**
- When STAGING_MODE = Y: Every email goes to ADMIN_EMAIL, no exceptions
- When STAGING_MODE = N: Emails go to configured recipients (production mode)
- This affects audit reports, daily summaries, error alerts, health checks, watchdog notifications, test emails - everything

#### When to Use Staging Mode

**✅ USE STAGING MODE FOR:**
- Testing new threshold configurations before rolling to production
- Validating External Config changes (new recipients, exclusions, etc.)
- Testing code deployments or script modifications
- Troubleshooting issues without spamming stakeholders
- Training new team members on the system
- Debugging email delivery problems
- Verifying audit logic changes

**❌ DO NOT USE STAGING MODE FOR:**
- Normal daily operations
- When stakeholders need audit reports delivered
- Extended periods (prevents stakeholders from receiving notifications)

#### How to Check Current Mode

**Method 1: Admin Spreadsheet (Fastest)**
- Open Admin Spreadsheet (Instructions tab)
- Check Row 1: `🟢 PRODUCTION MODE` or `🟡 STAGING MODE`
- This indicator updates automatically via triggers

**Method 2: Script Properties (Authoritative)**
1. Open Admin Spreadsheet
2. Extensions > Apps Script
3. Project Settings (gear icon) > Script Properties
4. Look for STAGING_MODE property:
   - Value = `Y` → Staging Mode ACTIVE
   - Value = `N` → Production Mode ACTIVE
   - Missing → Defaults to Production Mode (`N`)

**Method 3: Admin Controls Menu**
- Admin Controls > Debug Email Delivery
- Opens dialog showing current delivery mode and ADMIN_EMAIL
- Useful for quick verification

#### How to Enable Staging Mode

**Step 1: Set Script Property**
1. Admin Spreadsheet > Extensions > Apps Script
2. Project Settings (gear icon) > Script Properties
3. If STAGING_MODE exists: Click edit, change value to `Y`
4. If STAGING_MODE missing: Click "Add script property"
   - Property: `STAGING_MODE`
   - Value: `Y`
5. Click Save

**Step 2: Update Delivery Mode Indicator**
1. Return to Admin Spreadsheet
2. Admin Controls > Sync Delivery Mode Now
3. Row 1 should now show: `🟡 STAGING MODE - All emails to ADMIN_EMAIL`
4. Verify ADMIN_EMAIL is correct (should appear in instruction line)

**Step 3: Verify Staging Mode Active**
1. Admin Controls > Debug Email Delivery
2. Confirm dialog shows: "Current Mode: STAGING"
3. Confirm ADMIN_EMAIL address is correct
4. Optionally: Admin Controls > Send Test Email (should arrive at ADMIN_EMAIL only)

#### How to Disable Staging Mode (Return to Production)

**Step 1: Set Script Property**
1. Admin Spreadsheet > Extensions > Apps Script
2. Project Settings > Script Properties
3. Find STAGING_MODE property
4. Change value to `N`
5. Click Save

**Step 2: Update Delivery Mode Indicator**
1. Return to Admin Spreadsheet
2. Admin Controls > Sync Delivery Mode Now
3. Row 1 should now show: `🟢 PRODUCTION MODE - Live email delivery`

**Step 3: Verify Production Mode Active**
1. Admin Controls > Debug Email Delivery
2. Confirm dialog shows: "Current Mode: PRODUCTION"
3. Optionally: Send test email to confirm proper delivery

**⚠️ IMPORTANT:** Always disable staging mode after testing is complete! Leaving staging mode enabled prevents stakeholders from receiving audit reports.

#### Staging Mode Testing Workflow

**Recommended Process:**

1. **Before Making Changes:**
   - Enable staging mode (set STAGING_MODE = Y)
   - Sync delivery mode indicator
   - Verify staging active via Debug Email Delivery

2. **Make Your Changes:**
   - Update External Config Spreadsheet (thresholds, recipients, etc.)
   - Modify code if needed
   - Update configurations

3. **Test Changes:**
   - If testing config changes: Admin Controls > Run Single Config Audit
   - If testing full system: Wait for next scheduled batch (8-9 AM EST)
   - If testing code changes: Admin Controls > Send Test Email
   - Check ADMIN_EMAIL inbox for results

4. **Verify Results:**
   - Review audit reports at ADMIN_EMAIL
   - Check for errors or unexpected behavior
   - Validate thresholds triggered correctly
   - Confirm email formatting looks good

5. **Return to Production:**
   - Disable staging mode (set STAGING_MODE = N)
   - Sync delivery mode indicator
   - Verify production mode active
   - Monitor first production email for correctness

6. **Document:**
   - Note what was tested in changelog or notes
   - Record any issues discovered
   - Update team on changes if needed

#### Production vs Staging Mode Email Behavior

**Production Mode (STAGING_MODE = N):**
- **Audit Reports:** Sent to configured recipients from External Config Recipients sheet
- **CC/BCC:** Honored as configured; ADMIN_EMAIL typically BCC'd on audit emails
- **Daily Summary:** Sent to distribution list in code
- **Error Notifications:** Sent to ADMIN_EMAIL
- **Health Checks:** Sent to ADMIN_EMAIL
- **Watchdog Alerts:** Sent to ADMIN_EMAIL
- **Test Emails:** Sent to ADMIN_EMAIL

**Staging Mode (STAGING_MODE = Y):**
- **Everything:** ALL emails (audit, summary, error, health, watchdog, test) redirect to ADMIN_EMAIL only
- **No CC/BCC:** Original recipients never receive emails
- **No Distribution List:** Summary emails go to ADMIN_EMAIL only
- **Complete Isolation:** Stakeholders receive nothing while staging mode is active

#### Safety Considerations

**Risk Management:**
- Staging mode completely silences stakeholder communications
- If left enabled during business hours, stakeholders miss audit reports
- Always disable staging mode before leaving for the day
- Set calendar reminder if enabling staging mode for extended testing

**Best Practices:**
1. Use staging mode only when actively testing
2. Minimize staging mode duration (hours, not days)
3. Notify team if staging mode will be active during normal audit hours (8-9 AM EST)
4. Double-check production mode restored before end of business day
5. Document staging mode usage (what was tested, when, results)

**Recovery:**
- If you accidentally leave staging mode enabled overnight, stakeholders miss morning audit reports
- Resolution: Disable staging mode, consider running manual audits if time-sensitive
- Next day's scheduled audits will resume normal delivery automatically

#### Code Reference

**Key Functions:**
- `getStagingMode_()` (Code.js line ~12): Reads STAGING_MODE property, defaults to 'N'
- `safeSendEmail()` (Code.js line ~755): Checks staging mode, redirects emails if Y
- Email routing logic (Code.js line ~5272-5273): Returns ADMIN_EMAIL when staging mode active

**Script Property:**
- Property Name: `STAGING_MODE`
- Valid Values: `Y` (staging) or `N` (production)
- Default: `N` if property missing
- Location: Project Settings > Script Properties

### Batch Configuration

**Current Setting:** 2 configs per batch (BATCH_SIZE = 2)

**Why Batching:**
- Prevents timeout errors (6-minute Apps Script limit)
- Spreads load throughout the day
- Allows retry logic for failed configs

**How Batches Are Compiled:**

The system automatically rebalances batches **every night at 2:24 AM EST** during nightly maintenance to distribute workload evenly across the 12 batches.

**Rebalancing Algorithm (High-Low Pairing):**

1. **Collect Metrics:** System reads previous day's flagged counts for each config
2. **Sort Configs:** Sorts configs by flagged counts (highest to lowest)
3. **Pair High with Low:** Uses alternating pattern to distribute load evenly:
   - Batch 1: Highest flagged config + Lowest flagged config
   - Batch 2: 2nd highest + 2nd lowest
   - Batch 3: 3rd highest + 3rd lowest
   - ... continues until all configs assigned

**Example:**
- If you have 24 configs with flagged counts: [500, 450, 400, 350, 300, 250, 200, 150, 100, 90, 80, 70, ...]
- Batch assignments:
  - Batch 1: Config with 500 flags + Config with 70 flags
  - Batch 2: Config with 450 flags + Config with 80 flags
  - Batch 3: Config with 400 flags + Config with 90 flags
  - And so on...

**Why This Matters:**
- Configs with more flags take longer to process (more rows to evaluate, larger emails)
- Pairing high-volume with low-volume configs balances execution time across batches
- Prevents any single batch from taking significantly longer than others
- Reduces risk of timeout errors

**When Rebalancing Happens:**
- **Automatic:** Every night at 2:24 AM EST as part of `runNightlyMaintenance`
- **Manual:** Admin Controls > Install All Triggers (reinstalls triggers and rebalances)
- **Fallback:** If all configs have same flagged counts (or no previous data), uses alphabetical order

**Special Cases:**
- **New configs:** Assigned metric value of 100 (mid-range) until first run completes
- **All metrics tied:** System retains existing custom order if available, otherwise uses alphabetical
- **Config added/removed:** Next nightly rebalance redistributes all configs

**To View Current Batch Assignments:**
- Admin Controls > 📦 Batch Assignments
- Shows which configs are in each batch (1-12)
- Displays balance across batches

**Code Reference:**
- Function: `rebalanceAuditBatchesUsingSummary()` (Code.js line 6015)
- Called by: `runNightlyMaintenance()` (Code.js line 7709)
- Stores order in: Script Properties > `CM360_CUSTOM_CONFIG_ORDER`

**To Change Batch Size:**
1. Edit Code.js line ~68: const BATCH_SIZE = 2;
2. Deploy changes
3. Reinstall triggers: installAllAutomationTriggers()

---

## Troubleshooting

### Common Issues

#### 1. "No files found" - Audit Skipped

**Cause:** No reports in Gmail label for today

**Resolution:**
- Verify CM360 report scheduled correctly
- Check Gmail label matches configuration exactly
- Confirm report emails have attachments (.xlsx or .zip)
- Check spam folder for CM360 emails

#### 2. "Header not found" - Import Failed

**Cause:** Report structure changed, missing required columns

**Resolution:**
- Open the temp file in Drive (check error email for path)
- Verify columns: Advertiser, Campaign, Site, Placement ID, Placement, Dates, Creative, Pixel Sizes, Date, Impressions, Clicks
- If CM360 report template changed, contact Platform Solutions team

#### 3. Triggers Not Firing

**Cause:** Rare scenario - service account issue or authorization problem

**Resolution:**
1. Log into platformsolutionshmi@gmail.com
2. Open Admin Spreadsheet
3. Extensions > Apps Script > Triggers (clock icon)
4. Verify triggers exist and are enabled
5. If missing: Run installAllAutomationTriggers()
6. If authorization errors: Re-authorize all scopes

**Note:** Since system runs under service account, triggers should persist across team member changes. Only service account issues would cause trigger failures.

#### 4. "Script timeout" Errors

**Cause:** Too many configs in one batch, or large files

**Resolution:**
- Reduce BATCH_SIZE: Change const BATCH_SIZE = 2; to = 1;
- Split large config into smaller configs
- Check for extremely large report files (>50MB)

#### 5. Emails Not Sending

**Cause:** Quota exceeded or authorization issue

**Resolution:**
- Check daily quota: Run MailApp.getRemainingDailyQuota()
- Quota limit: ~1,500 emails/day
- If exceeded, audits accumulate and send next day
- Check Executions log for authorization errors

#### 6. "ADMIN_EMAIL not found" in Recipients

**Cause:** ADMIN_EMAIL constant not defined or code error

**Resolution:**
1. Extensions > Apps Script
2. Open Code.js
3. Find ADMIN_EMAIL constant definition (line ~29)
4. Verify it returns a valid email address
5. Save and deploy if changed

---

## Manual Operations

### Running Audits Manually

**Process Single Config:**
`javascript
// In Apps Script editor
function testSingleConfig() {
  const config = getAuditConfigs().find(c => c.name === 'YourConfigName');
  if (config) {
    processSingleAuditConfig(config, getRecipientsData(), getThresholdsData(), getExclusionsData());
  }
}
`

**Process All Configs:**
`javascript
runBatchedDailyAudits()
`

**Send Summary Email Now:**
`javascript
previewDailySummaryNow()  // Preview first
attemptSendDailySummary_({ allowPlaceholders: true, reason: 'Manual trigger' })  // Send
`

### Cleanup Operations

**Delete Old Gmail Emails (90+ days):**
`javascript
deleteOldAuditEmails()
`

**Delete Old Drive Files (60+ days):**
`javascript
cleanupOldAuditFiles()
`

**Force Config Sync:**
`javascript
syncFromExternalConfig()
`

### Testing Functions

**Test Email Sending:**
`javascript
safeSendEmail({
  to: 'your.email@horizonmedia.com',
  subject: 'Test Email',
  plainBody: 'This is a test',
  htmlBody: '<p>This is a test</p>'
}, 'Manual test');
`

**Check Staging Mode:**
`javascript
getStagingMode_()  // Returns 'Y' or 'N'
`

**View Audit Results Cache:**
`javascript
getCombinedAuditResults_()
`

### Accessing Logs

**View Execution Logs:**
1. Extensions > Apps Script
2. Left sidebar: **Executions** (list icon)
3. Click any execution to see logs
4. Filter by status: Success, Failed, Timeout

**View Logger Output:**
1. Extensions > Apps Script
2. Run any function
3. Bottom panel: **Execution log** tab
4. Shows real-time Logger.log() output

---

## Maintenance Tasks

### Daily (Automated - Monitor Only)

**Audit Batch Runs (8:00 AM - 9:00 AM EST):**
- ✅ `runDailyAuditsBatch1` through `runDailyAuditsBatch12` (12 batches total)
- All batches run within 1-hour window in the morning
- Each batch processes 2 configs
- Typical execution: 1-4 minutes per batch

**Nightly Maintenance (2:00 AM - 3:00 AM EST):**
- ✅ `runNightlyMaintenance` @ 2:24 AM (~6 minutes)
  - Rebalances audit batches
  - Updates placement names
  - Cleans up Drive files (60+ days old)
  - Deletes Gmail emails (90+ days old)
  - Clears daily script properties
- ✅ `cleanupOldAuditFiles` @ 2:32 AM (~5 minutes)
  - Continuation of Drive file cleanup if needed

**Monitoring & Alerting:**
- ✅ `auditWatchdogCheck` - Every 3 hours (checks for stuck audits)
- ✅ `forwardGASFailureNotificationsToAdmin` - Hourly (forwards script failures)
- ✅ `sendDailySummaryFailover` @ 9:25 AM EST (sends consolidated daily summary)

**Configuration Sync:**
- ✅ `runDeliveryModeSync` - Every 3 hours (syncs staging/production mode)
- ✅ `autoFixRequestsSheet_` - Every 4 hours (maintenance for requests sheet)

**Health Checks:**
- ✅ `runHealthCheckAndEmail` @ 5:04 AM EST (system health report)

**Your Task:** Review summary email (arrives ~9:30 AM), respond to errors, monitor execution logs for failures

### Weekly (Manual)

**Monday Morning:**
- Review deletion log for anomalies
- Check Drive storage usage
- Verify all configs ran successfully last week

### Monthly (Manual)

**First Monday:**
- Audit recipient list accuracy
- Review threshold settings for effectiveness
- Check for orphaned Gmail labels
- Verify Drive folder structure intact

**Action Items:**
1. Run: previewDailySummaryNow() - spot check stats
2. Open Deletion Log - verify cleanup is working
3. Check Drive folder: Project Log Files > CM360 Daily Audits
4. Review email quota usage trend

### Quarterly (Manual)

**Cleanup Tasks:**
- Remove obsolete configurations
- Archive old deletion logs (1+ years old)
- Review and update recipient email addresses
- Update threshold values based on campaign changes

**Documentation:**
- Update this README with any process changes
- Document new configurations added
- Note any recurring issues and resolutions

---

## Technical Architecture

### Code Structure

**Code.js** (~9,300 lines) organized in sections:

1. **Configuration Constants** (Lines 1-100)
   - Admin email, paths, batch size
   - Sheet names, cleanup settings

2. **Helper Functions** (Lines 100-1000)
   - Folder/file operations
   - Email sending
   - Text normalization
   - Error handling

3. **Audit Core Logic** (Lines 1000-3000)
   - Report fetching from Gmail
   - Excel/CSV merging
   - Threshold checking
   - Email generation

4. **Configuration Management** (Lines 3000-4500)
   - Config sheet reading
   - External sync
   - Validation

5. **Cleanup Operations** (Lines 4500-5500)
   - Drive file deletion
   - Gmail email deletion
   - Deletion logging

6. **Trigger Management** (Lines 5500-6500)
   - Trigger installation
   - Batch orchestration
   - State management

7. **UI Functions** (Lines 6500+)
   - Dashboard rendering
   - Admin controls
   - Preview functions

### Data Flow

`
CM360 Report (Email) 
   Gmail Label (Daily Audits/CM360/Config)
   Apps Script Fetch (fetchDailyAuditAttachments)
   Temp Drive Folder
   Excel/CSV Merge (mergeDailyAuditExcels)
   Threshold Check (auditMergedReport)
   Flag Discrepancies
   Generate HTML Email
   Send to Recipients
   Archive Merged Report
   Log Statistics
`

### Storage Locations

**Drive Folder Structure:**
`
Google Drive (Root)
 Project Log Files
     CM360 Daily Audits
         Deletion Log
            CM360 Deletion Log (Spreadsheet)
         To Trash After 60 Days
             Temp Daily Reports
                [ConfigName]
                    Temp_CM360_[timestamp] folders
             Merged Reports
                 [ConfigName]
                     CM360_Merged_Audit_[CONFIG]_[DATE].xlsx
`

**Script Properties:**
- ADMIN_EMAIL - Admin notification address
- STAGING_MODE - Y/N for testing vs production
- TRASH_ROOT_PATH - JSON array: ["Project Log Files", "CM360 Daily Audits"]
- CM360_LATEST_REPORT_URLS - JSON map of latest merged report URLs
- CM360_LAST_COUNTS - JSON object of previous day flagged counts
- CM360_AUDIT_RUN_STATE_V1_[ID] - Batch execution state tracking
- CM360_CLEANUP_STATE_V1 - Cleanup continuation state
- CM360_SUMMARY_SENT - Daily flag (clears after 6 hours)

### Gmail Label Structure

`
Daily Audits
 CM360
     ConfigName1
     ConfigName2
     ConfigName3
     ... (one per audit config)
`

---

## Common Scenarios

### Scenario 1: Adding a New Client

**Context:** New client "ACME Corp" needs daily audits

**Steps:**
1. Create Gmail label: Daily Audits/CM360/ACMECorp
2. External Config > Recipients sheet:
   - Config Name: ACMECorp
   - To: cme.client@example.com
   - CC: cme.am@horizonmedia.com
   - Gmail Label: Daily Audits/CM360/ACMECorp
3. External Config > Thresholds sheet:
   - Config Name: ACMECorp
   - Set thresholds (copy from similar client)
4. External Config > Exclusions sheet (if needed):
   - Add any placements/creatives to ignore
5. Force sync: Run syncFromExternalConfig()
6. Set up CM360 report to be labeled correctly
7. Wait for next audit run (or run manually to test)

### Scenario 2: Client Stops Service

**Context:** Client "OldCo" contract ended

**Steps:**
1. External Config > Recipients sheet:
   - Clear the Gmail Label cell for OldCo row
   - DO NOT delete row yet
2. Audits will automatically skip this config
3. After 30 days:
   - Delete row from Recipients, Thresholds, Exclusions
   - Archive or delete Gmail label
   - Archive or delete Drive folders for this config

### Scenario 3: Threshold Tuning

**Context:** Client getting too many/too few flags

**Steps:**
1. Review recent audit emails for the client
2. Note which thresholds are triggering
3. External Config > Thresholds sheet:
   - Find client row
   - Adjust threshold values:
     - **Too many flags:** Increase thresholds
     - **Too few flags:** Decrease thresholds
4. Common adjustments:
   - Impressions: 10-20% tolerance typical
   - Clicks: Often needs wider tolerance (30%)
   - Pixel sizes: Usually exact match required
5. Force sync if immediate: syncFromExternalConfig()
6. Monitor next few days of audits

### Scenario 4: Report Structure Changed

**Context:** CM360 changed report template, audits failing

**Steps:**
1. Check error emails for "Header not found"
2. Download one of the failing reports
3. Compare headers to expected columns (see Code.js line ~829)
4. **If minor change** (renamed column):
   - Update getExpectedHeaderSpec_() function in Code.js
   - Add new name as alias: ['Old Name', 'New Name']
5. **If major change**:
   - Contact Platform Solutions team
   - May require code refactoring
6. Test with one config before deploying to all

### Scenario 5: Emergency Stop

**Context:** System sending incorrect alerts, need to stop immediately

**Steps:**
1. **Stop all triggers:**
   - Extensions > Apps Script > Triggers
   - Delete ALL triggers manually
2. **Enable staging mode:**
   - Project Settings > Script Properties
   - Set STAGING_MODE to Y
3. **Investigate issue:**
   - Review Executions logs
   - Check recent code changes
   - Test with single config
4. **Resume when fixed:**
   - Set STAGING_MODE to N
   - Run installAllAutomationTriggers()

### Scenario 6: Transition When Evan Leaves

**Context:** Evan leaving, team taking over system

**Key Point:** System runs under platformsolutionshmi@gmail.com, so **no disruption** to automated processes.

**Steps:**

1. **Update ADMIN_EMAIL** (see Critical First Steps)
   - Code.js line ~29: Update email address
   - Deploy updated code to Apps Script

2. **Transfer platformsolutionshmi@gmail.com access:**
   - Evan provides credentials to team (secure handoff)
   - Team lead stores credentials securely
   - Update recovery email/phone if needed
   - Document 2FA backup codes

3. **Update summary email distribution (optional):**
   - Code.js line ~1061: Update recipient list if needed
   - Deploy changes if modified

4. **Test as new admin:**
   - Log into platformsolutionshmi@gmail.com
   - Open Admin Spreadsheet
   - Open Apps Script editor
   - Run a manual audit (Admin Controls > ▶️ Run Audit for...)
   - Verify error notifications go to new ADMIN_EMAIL

5. **Monitor first week:**
   - Check daily summary emails arrive (~9:30 AM EST)
   - Verify audit emails sending to clients
   - Review execution logs for errors
   - Confirm nightly maintenance runs

**What does NOT need to happen:**
- ❌ No trigger reinstallation required
- ❌ No spreadsheet ownership transfer required
- ❌ No Drive folder migration required
- ❌ No Gmail label recreation required

**Why it's seamless:**
- All resources owned by service account
- Triggers belong to platformsolutionshmi@gmail.com
- Only admin notification email changes

---

## Emergency Contacts

### Primary Contacts

**Platform Solutions Team:**
- Role: Technical support for system issues
- Email: [To be filled in]
- Escalation: [To be filled in]

**AdOps Leadership:**
- Role: Business decisions on configurations
- Email: [To be filled in]

**IT/DevOps:**
- Role: Google Workspace, permissions, access issues
- Email: [To be filled in]

### Summary Email Recipients (Update These!)

Current distribution list (Code.js line ~1061):
- evschneider@horizonmedia.com (Consider updating to new team lead)
- bmuller@horizonmedia.com
- bkaufman@horizonmedia.com
- ewarburton@horizonmedia.com

**Note:** This is separate from ADMIN_EMAIL. These are recipients of the daily summary, while ADMIN_EMAIL receives error notifications.

### Escalation Path

1. **Minor issues** (one config failing):
   - Review error email
   - Check config settings
   - Verify CM360 report delivery
   - Fix and monitor

2. **Multiple configs failing**:
   - Check system-wide settings
   - Review recent code changes
   - Verify trigger status
   - Contact Platform Solutions if persistent

3. **System completely down**:
   - Check trigger ownership
   - Verify spreadsheet access
   - Review Script Properties
   - Emergency stop if needed (see Scenario 5)
   - Escalate to Platform Solutions immediately

4. **Data integrity concerns**:
   - Stop system immediately
   - Review deletion logs
   - Check Drive storage
   - Escalate to leadership + Platform Solutions

---

## Additional Resources

### Google Apps Script Documentation

- Main docs: https://developers.google.com/apps-script
- SpreadsheetApp: https://developers.google.com/apps-script/reference/spreadsheet
- GmailApp: https://developers.google.com/apps-script/reference/gmail
- DriveApp: https://developers.google.com/apps-script/reference/drive

### Useful Scripts Documentation

**Clasp** (Command Line Apps Script):
- Install: 
pm install -g @google/clasp
- Login: clasp login
- Push code: clasp push
- Pull code: clasp pull

### Code Repository

**GitHub:** https://github.com/evan-schneider/cm360-audit-system  
**Branch:** integrate-2025-09-29

**To Clone:**
`ash
git clone https://github.com/evan-schneider/cm360-audit-system.git
cd cm360-audit-system
git checkout integrate-2025-09-29
`

**To Deploy Changes:**
`ash
# Edit Code.js locally
git add Code.js
git commit -m "Description of changes"
git push origin integrate-2025-09-29

# Push to Apps Script
npx clasp push
`

---

## Change Log Template

Use this format when documenting system changes:

`
### [Date] - [Your Name]

**Change:** Brief description

**Reason:** Why this change was needed

**Impact:** What users/configs are affected

**Testing:** How you verified it works

**Rollback:** How to undo if needed
`

**Example:**
`
### 2025-10-29 - Evan Schneider

**Change:** Added Gmail email cleanup function (deleteOldAuditEmails)

**Reason:** Gmail storage approaching limit from daily audit emails

**Impact:** Emails older than 90 days automatically deleted nightly

**Testing:** Ran manually, verified deletion log entries, confirmed email notification

**Rollback:** Remove from runNightlyMaintenance(), remove trigger invocation
`

---

## Final Notes from Evan

### What Works Well

- Batched execution prevents timeouts
- External config sync allows changes without script access
- Deletion logging provides audit trail
- Email suppression in staging mode makes testing safe
- Automatic retry logic handles transient failures

### Known Limitations

- 6-minute Apps Script execution limit (hence batching)
- 1,500 email/day quota limit
- Can't permanently delete Gmail (only trash)
- Excel import sometimes flaky with formatting
- Large files (>50MB) can cause memory issues

### Future Improvements (Backlog)

- [ ] Web dashboard for real-time monitoring
- [ ] Slack integration for alerts
- [ ] Configurable retention periods per config
- [ ] Automatic threshold tuning based on historical data
- [ ] Multi-region support for international clients
- [ ] API integration for CM360 (remove email dependency)
