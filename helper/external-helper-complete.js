// CM360 Config Helper - External (external-user-safe)
// Enhanced helper for external editors: request audits, view config status, and create new configs.

const ADMIN_EMAIL = 'evschneider@horizonmedia.com';
const AUDIT_REQUESTS_URL = 'https://docs.google.com/spreadsheets/d/1MUDE5geWlO9Flmy3vtfCNRrsnpDAMcz0z1uA0Lu2Ilw/edit?gid=951444608#gid=951444608';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CM360 Config Helper')
    .addItem('🔍 Run Config Audit', 'showConfigAuditRunner')
    .addSeparator()
    .addItem('📄 Show Config Summary', 'showConfigSummary')
    .addSeparator()
    .addItem('🧩 Create New Config…', 'createNewConfig')
    .addToUi();
}

/* -----------------------
   Create New Config Function
   ----------------------- */
function createNewConfig() {
  const ui = SpreadsheetApp.getUi();
  
  // Step 1: Get Config ID
  const configResponse = ui.prompt(
    'Create New Config - Step 1 of 2',
    'Enter a unique alphanumeric Config ID (e.g., AMC01, NEXT01, PST01):\n\n• Use the same Config ID if you want multiple networks/reports in the same email\n• Please confirm your chosen Config ID with the admin',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (configResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const configId = String(configResponse.getResponseText() || '').trim().toUpperCase();
  if (!configId || !/^[A-Z0-9]+$/.test(configId)) {
    ui.alert('Invalid Config ID', 'Please enter an alphanumeric Config ID (e.g., PST01, ENT01, NEXT01).', ui.ButtonSet.OK);
    return;
  }
  
  // Step 2: Get Recipients and Thresholds
  const formHtml = getCreateConfigFormHtml_(configId);
  const htmlOutput = HtmlService.createHtmlOutput(formHtml)
    .setWidth(600)
    .setHeight(700);
  ui.showModalDialog(htmlOutput, 'Create New Config: ' + configId);
}

function getCreateConfigFormHtml_(configId) {
  return `
    <div style="font-family:Arial,sans-serif; padding:20px; max-width:600px;">
      <h3>Create New Config: ${configId}</h3>
      
      <div style="margin:15px 0;">
        <label><strong>Recipients (required):</strong></label><br>
        <input type="email" id="recipients" style="width:100%; padding:5px;" placeholder="email1@company.com, email2@company.com" required>
      </div>
      
      <div style="margin:15px 0;">
        <label><strong>CC Recipients (optional):</strong></label><br>
        <input type="email" id="cc" style="width:100%; padding:5px;" placeholder="cc1@company.com, cc2@company.com">
      </div>
      
      <h4>Thresholds for Flag Types:</h4>
      
      <div style="display:grid; grid-template-columns:2fr 1fr 1fr; gap:10px; margin:10px 0; font-weight:bold;">
        <div>Flag Type</div>
        <div style="text-align:center;">Min Impressions</div>
        <div style="text-align:center;">Min Clicks</div>
      </div>
      
      <div style="display:grid; grid-template-columns:2fr 1fr 1fr; gap:10px; margin:5px 0; align-items:center;">
        <div>Clicks Greater Than Impressions</div>
        <input type="number" id="t_cgti_i" value="0" min="0" style="padding:3px;">
        <input type="number" id="t_cgti_c" value="0" min="0" style="padding:3px;">
      </div>
      
      <div style="display:grid; grid-template-columns:2fr 1fr 1fr; gap:10px; margin:5px 0; align-items:center;">
        <div>Out of Flight Dates</div>
        <input type="number" id="t_oofd_i" value="0" min="0" style="padding:3px;">
        <input type="number" id="t_oofd_c" value="0" min="0" style="padding:3px;">
      </div>
      
      <div style="display:grid; grid-template-columns:2fr 1fr 1fr; gap:10px; margin:5px 0; align-items:center;">
        <div>Pixel Size Mismatch</div>
        <input type="number" id="t_psm_i" value="0" min="0" style="padding:3px;">
        <input type="number" id="t_psm_c" value="0" min="0" style="padding:3px;">
      </div>
      
      <div style="display:grid; grid-template-columns:2fr 1fr 1fr; gap:10px; margin:5px 0; align-items:center;">
        <div>Default Ad Serving</div>
        <input type="number" id="t_das_i" value="0" min="0" style="padding:3px;">
        <input type="number" id="t_das_c" value="0" min="0" style="padding:3px;">
      </div>
      
      <div style="background:#f8f9fa; padding:15px; margin:20px 0; border-left:4px solid #1a73e8; font-size:13px;">
        <h4 style="margin-top:0;">CM360 Daily Reports Requirements</h4>
        
        <p><strong>Basic Info:</strong><br>
        Name: NETWORKNAME_ImpClickReport_DailyAudit_${configId}<br>
        • Network Name should be the name of the DCM network (can use shorthand notation)</p>
        
        <p><strong>Date Range:</strong> Yesterday</p>
        
        <p><strong>Fields (in order):</strong><br>
        <strong>Dimensions:</strong> Advertiser (FILTER IF NEEDED), Campaign, Site (CM360), Placement ID, Placement, Placement Start Date, Placement End Date, Ad Type (FILTER OUT SA360 & DART Search), Creative, Placement Pixel Size, Creative Pixel Size, Date<br>
        <strong>Metrics:</strong> Impressions, Clicks</p>
        
        <p><strong>Scheduling:</strong><br>
        Time zone: Eastern Time (GMT-4:00)<br>
        Repeats: Daily<br>
        Every: 1 day<br>
        Starts: today<br>
        Ends: as late as possible<br>
        Format: Excel (Attachment)<br>
        Share with: platformsolutionshmi@gmail.com</p>
        
        <p>Reach out to ${ADMIN_EMAIL} with any questions</p>
      </div>
      
      <div style="margin-top:20px;">
        <button onclick="submitForm()" style="background:#1a73e8; color:white; padding:10px 20px; border:none; border-radius:4px; cursor:pointer;">
          Submit New Config
        </button>
        <button onclick="google.script.host.close()" style="background:#6c757d; color:white; padding:10px 20px; border:none; border-radius:4px; cursor:pointer; margin-left:10px;">
          Cancel
        </button>
      </div>
    </div>
    
    <script>
      function validateEmail(email) {
        return /^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/.test(email);
      }
      
      function validateEmails(emailString) {
        if (!emailString.trim()) return false;
        const emails = emailString.split(',').map(e => e.trim());
        return emails.every(email => validateEmail(email));
      }
      
      function submitForm() {
        const configId = '${configId}';
        const recipients = document.getElementById('recipients').value.trim();
        const cc = document.getElementById('cc').value.trim();
        
        // Validation
        if (!recipients) {
          alert('Please enter at least one recipient email address.');
          return;
        }
        
        if (!validateEmails(recipients)) {
          alert('Please enter valid recipient email addresses (comma-separated).');
          return;
        }
        
        if (cc && !validateEmails(cc)) {
          alert('Please enter valid CC email addresses (comma-separated).');
          return;
        }
        
        const formData = {
          configId: configId,
          recipients: recipients,
          cc: cc,
          thresholds: {
            clicks_greater_than_impressions: {
              minImpressions: Number(document.getElementById('t_cgti_i').value || 0),
              minClicks: Number(document.getElementById('t_cgti_c').value || 0)
            },
            out_of_flight_dates: {
              minImpressions: Number(document.getElementById('t_oofd_i').value || 0),
              minClicks: Number(document.getElementById('t_oofd_c').value || 0)
            },
            pixel_size_mismatch: {
              minImpressions: Number(document.getElementById('t_psm_i').value || 0),
              minClicks: Number(document.getElementById('t_psm_c').value || 0)
            },
            default_ad_serving: {
              minImpressions: Number(document.getElementById('t_das_i').value || 0),
              minClicks: Number(document.getElementById('t_das_c').value || 0)
            }
          }
        };
        
        google.script.run
          .withSuccessHandler(function(message) {
            alert(message);
            google.script.host.close();
          })
          .withFailureHandler(function(error) {
            alert('Failed to create config: ' + error.message);
          })
          .submitNewConfigFromForm(formData);
      }
    </script>
  `;
}

function submitNewConfigFromForm(formData) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('[helper] submitNewConfigFromForm called with configId=' + (formData && formData.configId));
  
  try {
    const submitter = (Session && Session.getActiveUser ? Session.getActiveUser().getEmail() : '') || 'Unknown';
    const timestamp = new Date();
    const formattedDate = timestamp.getFullYear() + '-' + 
                         String(timestamp.getMonth() + 1).padStart(2, '0') + '-' + 
                         String(timestamp.getDate()).padStart(2, '0');
    
    // Add to Audit Recipients
    let recipientsSheet = ss.getSheetByName('Audit Recipients');
    if (!recipientsSheet) {
      recipientsSheet = ss.insertSheet('Audit Recipients');
      const headers = ['Config Name', 'Primary Recipients', 'CC Recipients', 'Active', 'Withhold No-Flag Emails', 'Last Updated'];
      recipientsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      recipientsSheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#4285f4')
        .setFontColor('#ffffff');
    }
    
    recipientsSheet.appendRow([
      formData.configId,
      formData.recipients,
      formData.cc || '',
      'TRUE',
      'FALSE',
      formattedDate
    ]);
    
    // Add to Audit Thresholds
    let thresholdsSheet = ss.getSheetByName('Audit Thresholds');
    if (!thresholdsSheet) {
      thresholdsSheet = ss.insertSheet('Audit Thresholds');
      const headers = ['Config Name', 'Flag Type', 'Min Impressions', 'Min Clicks', 'Active'];
      thresholdsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      thresholdsSheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#4285f4')
        .setFontColor('#ffffff');
    }
    
    const flagTypes = ['clicks_greater_than_impressions', 'out_of_flight_dates', 'pixel_size_mismatch', 'default_ad_serving'];
    flagTypes.forEach(flagType => {
      const threshold = formData.thresholds[flagType];
      thresholdsSheet.appendRow([
        formData.configId,
        flagType,
        threshold.minImpressions,
        threshold.minClicks,
        'TRUE'
      ]);
    });
    
    // Send notification to admin (robust: MailApp first, GmailApp fallback)
    const thresholdsSummary = flagTypes.map(ft => {
      const t = formData.thresholds[ft];
      return `  ${ft}: ${t.minImpressions} impressions, ${t.minClicks} clicks`;
    }).join('\n');

    const subject = `New CM360 Config Created: ${formData.configId}`;
    const plainBody = `New CM360 config created via Helper Menu\n\n` +
      `Config ID: ${formData.configId}\n` +
      `Created by: ${submitter}\n` +
      `Recipients: ${formData.recipients}\n` +
      `CC: ${formData.cc || 'None'}\n\n` +
      `Thresholds:\n${thresholdsSummary}\n\n` +
      `Next steps:\n` +
      ` 1) Admin Controls → Sync FROM External Config\n` +
      ` 2) Admin Controls → Prepare Environment (labels, Drive folders)\n` +
      ` 3) Create Gmail filters for Daily Audits/CM360/${formData.configId}`;

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width:600px;">
        <h3 style="color:#1a73e8;">New CM360 Config Created</h3>
        <p>A new configuration has been created via the Helper Menu:</p>
        <table style="border-collapse:collapse; margin:10px 0; width:100%;">
          <tr><td style="padding:8px; font-weight:bold; border:1px solid #ddd;">Config ID:</td><td style="padding:8px; border:1px solid #ddd;">${formData.configId}</td></tr>
          <tr><td style="padding:8px; font-weight:bold; border:1px solid #ddd;">Created by:</td><td style="padding:8px; border:1px solid #ddd;">${submitter}</td></tr>
          <tr><td style="padding:8px; font-weight:bold; border:1px solid #ddd;">Recipients:</td><td style="padding:8px; border:1px solid #ddd;">${formData.recipients}</td></tr>
          <tr><td style="padding:8px; font-weight:bold; border:1px solid #ddd;">CC:</td><td style="padding:8px; border:1px solid #ddd;">${formData.cc || 'None'}</td></tr>
        </table>
        <h4>Thresholds Added:</h4>
        <pre style="background:#f8f9fa; padding:10px; border-left:3px solid #1a73e8;">${thresholdsSummary}</pre>
        <h4>Next Steps:</h4>
        <ol>
          <li><strong>SYNC FROM External Config:</strong> Use Admin Controls → "📥 Sync FROM External Config" to pull this new config into the main system</li>
          <li><strong>Create Gmail Label & Drive Folders:</strong> Use Admin Controls → "⚙️ Prepare Environment" to create the Gmail label and Drive folder structure</li>
          <li>Set up Gmail filters to route reports to the new label: Daily Audits/CM360/${formData.configId}</li>
        </ol>
        <p>The config has been added to the <strong>Audit Recipients</strong> and <strong>Audit Thresholds</strong> tabs on the Helper Menu sheet.</p>
      </div>
    `;

    let notified = false;
    Logger.log('[helper] Attempting to notify admin: ' + ADMIN_EMAIL + ' subject=' + subject);
    try {
      MailApp.sendEmail({ to: ADMIN_EMAIL, subject, body: plainBody, htmlBody });
      notified = true;
    } catch (e1) {
      Logger.log('[helper] MailApp.sendEmail failed: ' + (e1 && e1.message));
      try {
        GmailApp.sendEmail(ADMIN_EMAIL, subject, plainBody, { htmlBody });
        notified = true;
      } catch (e2) {
        Logger.log('Admin notification failed via MailApp and GmailApp: ' + (e2 && e2.message));
      }
    }
    
    return `Config ${formData.configId} has been successfully created!

Added to Helper Menu sheet:
• Audit Recipients: 1 row
• Audit Thresholds: ${flagTypes.length} rows

${notified ? 'Admin has been notified and will:' : 'Admin notification failed to send automatically. Please forward this message to ' + ADMIN_EMAIL + ' and include the details above. Admin will:'}
1. Sync the config into the main system
2. Create Gmail labels and Drive folders
3. Set up email filters

You should receive confirmation once setup is complete.`;
    
  } catch (error) {
    Logger.log('submitNewConfigFromForm error: ' + error.message);
    throw new Error('Failed to create config: ' + error.message);
  }
}

/* -----------------------
   Audit Requests (write)
   ----------------------- */
function _buildRequestsInstructions_() {
  return [
    ['How to create a request:', 'Use CM360 Config Helper → Run Config Audit. The system adds a row automatically. Do NOT add rows manually.'],
    ['', ''],
    ['When to use this tab:', 'This is a log/queue of requests—external users should not edit Status directly.'],
    ['', ''],
    ['Troubleshooting:', 'If requests stay PENDING, confirm access and ask an admin to check logs.'],
    ['Security:', 'Only admins should change Status values.'],
    ['', ''],
    ['Usage:', 'Leave entries to the system unless directed by an admin.']
  ];
}

function _ensureAuditRequestsHeader_(sheet) {
  const headers = ['Config Name', 'Requested By', 'Request Time', 'Status', 'Notes'];

  // Only set A–E headers (don't clear anything)
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0] || [];
  let mismatch = false;
  for (let i = 0; i < headers.length; i++) {
    if (String(current[i] || '').trim() !== headers[i]) { mismatch = true; break; }
  }
  if (mismatch || sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    try {
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#4285f4')
        .setFontColor('#ffffff');
    } catch (e) {}
  }

  // Ensure INSTRUCTIONS header exists (after a blank spacer), but don't overwrite existing instructions
  try {
    const lastCol = Math.max(sheet.getLastColumn(), headers.length);
    const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];

    let instrColIndex = headerRow.findIndex(h => String(h || '').trim().toUpperCase() === 'INSTRUCTIONS');
    if (instrColIndex === -1) {
      // Find last non-empty header, add blank spacer + INSTRUCTIONS
      let lastHeaderCol = 0;
      for (let c = headerRow.length - 1; c >= 0; c--) {
        if (String(headerRow[c] || '').trim() !== '') { lastHeaderCol = c + 1; break; }
      }
      instrColIndex = lastHeaderCol + 2; // 1-based target column
      sheet.getRange(1, instrColIndex, 1, 1).setValue('INSTRUCTIONS');
    } else {
      instrColIndex = instrColIndex + 1; // convert 0-based to 1-based
    }

    // Style INSTRUCTIONS header
    try {
      sheet.getRange(1, instrColIndex, 1, 1)
        .setFontWeight('bold')
        .setBackground('#ff9900')
        .setFontColor('#ffffff');
    } catch (e) {}

    // Add default instructions only if none exist yet
    try {
      const rowsBelow = Math.max(1, sheet.getLastRow() - 1);
      const instrArea = sheet.getRange(2, instrColIndex, rowsBelow, 2).getValues();
      const hasAnyInstr = instrArea.some(r => r.some(v => String(v || '').trim() !== ''));
      if (!hasAnyInstr) {
        const instr = _buildRequestsInstructions_();
        sheet.getRange(2, instrColIndex, instr.length, 2).setValues(instr);
        try {
          sheet.getRange(2, instrColIndex, instr.length, 2)
            .setFontSize(10)
            .setVerticalAlignment('top')
            .setWrap(false);
        } catch (e) {}
      }
    } catch (e) {}
  } catch (e) {}

  try { sheet.autoResizeColumns(1, headers.length); } catch (e) {}
}

function getOrCreateAuditRequestSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Audit Requests');
  if (!sheet) {
    sheet = ss.insertSheet('Audit Requests');
  }
  _ensureAuditRequestsHeader_(sheet);
  return sheet;
}

function requestConfigAudit(configName) {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = getOrCreateAuditRequestSheet();
    const timestamp = new Date().toISOString();
    const requester = (Session && Session.getActiveUser ? Session.getActiveUser().getEmail() : '') || '';

    // Compute first empty row inside columns A–E only
    const maxRow = Math.max(sheet.getLastRow(), 1);
    const aToE = sheet.getRange(1, 1, maxRow, 5).getValues(); // includes header row
    let writeRow = -1;
    let lastDataRowAE = 1;
    for (let r = 1; r < aToE.length; r++) {
      const rowVals = aToE[r];
      const hasDataInAE = rowVals.some(v => String(v || '').trim() !== '');
      if (hasDataInAE) {
        lastDataRowAE = r + 1; // 1-based
      } else if (writeRow === -1) {
        writeRow = r + 1;
      }
    }
    if (writeRow === -1) writeRow = lastDataRowAE + 1;

    sheet.getRange(writeRow, 1, 1, 5)
         .setValues([[configName, requester, timestamp, 'PENDING', 'Requested via Helper Menu']]);

    ui.alert(
      'Audit Request Submitted',
      `Audit request for ${configName} has been submitted.\n\nRequest ID: ${timestamp}\nStatus: PENDING\n\nYou will receive an email when the audit completes.`,
      ui.ButtonSet.OK
    );

    // Notify admin (best-effort with fallback)
    const subject = `CM360 Audit Request: ${configName}`;
    const plainBody = `A new CM360 audit has been requested.\n\n` +
      `Config: ${configName}\n` +
      `Requested by: ${requester}\n` +
      `Time: ${timestamp}\n\n` +
      `Audit Requests: ${AUDIT_REQUESTS_URL}`;
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width:600px;">
        <h3 style="color:#1a73e8;">CM360 Audit Request</h3>
        <p>A new audit has been requested:</p>
        <table style="border-collapse:collapse; margin:10px 0;">
          <tr><td style="padding:5px; font-weight:bold;">Config:</td><td style="padding:5px;">${configName}</td></tr>
          <tr><td style="padding:5px; font-weight:bold;">Requested by:</td><td style="padding:5px;">${requester}</td></tr>
          <tr><td style="padding:5px; font-weight:bold;">Time:</td><td style="padding:5px;">${timestamp}</td></tr>
        </table>
        <p>Please check the <a href="${AUDIT_REQUESTS_URL}" style="color:#1a73e8; text-decoration:none;">Audit Requests sheet</a>.</p>
      </div>
    `;
    Logger.log('[helper] Notifying admin of audit request: ' + ADMIN_EMAIL + ' subject=' + subject);
    try {
      MailApp.sendEmail({to: ADMIN_EMAIL, subject, body: plainBody, htmlBody});
    } catch (e1) {
      Logger.log('[helper] MailApp failed, trying GmailApp: ' + (e1 && e1.message));
      try {
        GmailApp.sendEmail(ADMIN_EMAIL, subject, plainBody, { htmlBody });
      } catch (e2) {
        Logger.log('Could not send notification email: ' + (e2 && e2.message));
      }
    }

  } catch (error) {
    ui.alert(
      'Request Failed',
      `Failed to submit audit request: ${error.message}\n\nPlease contact the administrator.`,
      ui.ButtonSet.OK
    );
  }
}

/* -----------------------
   Read-only sheet access helpers (do NOT create or overwrite)
   ----------------------- */
function getRecipientsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Audit Recipients') || null;
}
function getThresholdsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Audit Thresholds') || null;
}
function getExclusionsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Audit Exclusions') || null;
}

/* -----------------------
   Picker & runner for external users (reads recipients sheet)
   ----------------------- */
function showConfigAuditRunner() {
  const ui = SpreadsheetApp.getUi();
  const recipientsSheet = getRecipientsSheet();
  if (!recipientsSheet) {
    ui.alert(
      'No Data Found',
      'The Audit Recipients sheet is missing in this external config. Please ask an admin to sync configuration data.',
      ui.ButtonSet.OK
    );
    return;
  }
  const data = recipientsSheet.getDataRange().getValues();
  if (data.length <= 1) {
    ui.alert(
      'No Data Found',
      `The Audit Recipients sheet appears empty or only has headers.\n\nData rows found: ${Math.max(0, data.length - 1)}\n\nPlease ask admin to populate the configuration data.`,
      ui.ButtonSet.OK
    );
    return;
  }

  const configs = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const configName = row[0];
    const activeStatus = row[3];
    const isActive = (activeStatus === true) || (String(activeStatus || '').toUpperCase() === 'TRUE');
    if (configName && isActive) {
      configs.push({
        name: configName,
        recipients: row[1] || '',
        cc: row[2] || '',
        withhold: (String(row[4] || '').toUpperCase() === 'TRUE')
      });
    }
  }

  if (configs.length === 0) {
    ui.alert(
      'No Active Configurations Found',
      'No active configurations found in Audit Recipients sheet. Please ask admin to enable at least one configuration.',
      ui.ButtonSet.OK
    );
    return;
  }

  const configOptions = configs.map((config, index) => {
    const recipientCount = config.recipients.split(',').filter(r => r.trim()).length;
    const ccCount = config.cc ? config.cc.split(',').filter(r => r.trim()).length : 0;
    return `${index + 1}. ${config.name} (📧 ${recipientCount} recipients${ccCount > 0 ? ', ' + ccCount + ' CC' : ''}${config.withhold ? ', withholds no-flag emails' : ''})`;
  }).join('\n');

  const response = ui.prompt(
    'Select Configuration to Audit',
    `Available configurations:\n\n${configOptions}\n\nEnter the number (1-${configs.length}) of the configuration to audit:`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const selectedIndex = parseInt(response.getResponseText().trim(), 10) - 1;
  if (isNaN(selectedIndex) || selectedIndex < 0 || selectedIndex >= configs.length) {
    ui.alert('Invalid Selection', `Please enter a valid number between 1 and ${configs.length}.`, ui.ButtonSet.OK);
    return;
  }
  const selectedConfig = configs[selectedIndex];
  requestConfigAudit(selectedConfig.name);
}

/* -----------------------
   Summary (read-only)
   ----------------------- */

function showConfigSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let summary = 'Configuration Summary:\n\n';

  const recipients = getRecipientsSheet();
  if (recipients) {
    const data = recipients.getDataRange().getValues();
    const configs = new Set();
    for (let i = 1; i < data.length; i++) {
      const active = data[i][3];
      if (data[i][0] && ((active === true) || (String(active || '').toUpperCase() === 'TRUE'))) configs.add(data[i][0]);
    }
    summary += `📧 Recipients: ${configs.size} active configs\n`;
  } else {
    summary += `📧 Recipients: sheet missing\n`;
  }

  const thresholds = getThresholdsSheet();
  if (thresholds) {
    const data = thresholds.getDataRange().getValues();
    const configs = new Set();
    for (let i = 1; i < data.length; i++) {
      const active = data[i][4];
      if (data[i][0] && ((active === true) || (String(active || '').toUpperCase() === 'TRUE'))) configs.add(data[i][0]);
    }
    summary += `📊 Thresholds: ${configs.size} active configs\n`;
  } else {
    summary += `📊 Thresholds: sheet missing\n`;
  }

  const exclusions = getExclusionsSheet();
  if (exclusions) {
    const data = exclusions.getDataRange().getValues();
    // Count rows (skip header) where the Active column (K -> index 10) is TRUE
    let activeRows = 0;
    for (let i = 1; i < data.length; i++) {
      const activeVal = data[i][10];
      if (activeVal === true || (String(activeVal || '').toUpperCase() === 'TRUE')) {
        activeRows++;
      }
    }
    summary += `📋 Exclusions: ${activeRows} active rules\n`;
  } else {
    summary += `📋 Exclusions: sheet missing\n`;
  }

  ui.alert('Configuration Summary', summary, ui.ButtonSet.OK);
}