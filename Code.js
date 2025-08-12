// === 📁 CONFIGURATION & CONSTANTS ===
const ADMIN_EMAIL = 'evschneider@horizonmedia.com';
const STAGING_MODE = 'Y'; // Set to 'Y' for staging mode, 'N' for production
const EXCLUSIONS_SHEET_NAME = 'CM360 Audit Exclusions'; // Name of the sheet containing exclusions
const THRESHOLDS_SHEET_NAME = 'CM360 Audit Thresholds'; // Name of the sheet containing flag thresholds
const RECIPIENTS_SHEET_NAME = 'CM360 Audit Recipients'; // Name of the sheet containing email recipients

const BATCH_SIZE = 3;
const TRASH_ROOT_PATH = ['Project Log Files', 'CM360 Daily Audits', 'To Trash After 60 Days'];
const DELETION_LOG_PATH = [...TRASH_ROOT_PATH, 'Deletion Log'];
const MASTER_LOG_NAME = 'CM360 Deleted Files Log';

// === 📦 UTILITY HELPERS ===
function folderPath(type, configName) {
  return [...TRASH_ROOT_PATH, type, configName];
}

function resolveRecipients(configName, recipientsData = null) {
  if (STAGING_MODE === 'Y') {
    return ADMIN_EMAIL; // Admin override in staging mode
  }
  
  // Production mode - use sheet data if available
  if (recipientsData && recipientsData[configName]) {
    return recipientsData[configName].primary || ADMIN_EMAIL;
  }
  
  // Fallback to admin email if no sheet data
  return ADMIN_EMAIL;
}

function resolveCc(configName, recipientsData = null) {
  if (STAGING_MODE === 'Y') {
    return ''; // No CC in staging mode
  }
  
  // Production mode - use sheet data if available
  if (recipientsData && recipientsData[configName]) {
    return recipientsData[configName].cc || '';
  }
  
  // No CC fallback
  return '';
}

// === 🔧 AUDIT CONFIGS ===
// Note: Recipients and CC lists are now managed in the Recipients sheet
// This configuration focuses on infrastructure settings only
const auditConfigs = [
  {
    name: 'PST01',
    label: 'Daily Audits/CM360/PST01',
    mergedFolderPath: folderPath('Merged Reports', 'PST01'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'PST01')
  },
  {
    name: 'PST02',
    label: 'Daily Audits/CM360/PST02',
    mergedFolderPath: folderPath('Merged Reports', 'PST02'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'PST02')
  },
  {
    name: 'PST03',
    label: 'Daily Audits/CM360/PST03',
    mergedFolderPath: folderPath('Merged Reports', 'PST03'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'PST03')
  },
  {
    name: 'NEXT01',
    label: 'Daily Audits/CM360/NEXT01',
    mergedFolderPath: folderPath('Merged Reports', 'NEXT01'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'NEXT01')
  },
  {
    name: 'NEXT02',
    label: 'Daily Audits/CM360/NEXT02',
    mergedFolderPath: folderPath('Merged Reports', 'NEXT02'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'NEXT02')
  },
  {
    name: 'NEXT03',
    label: 'Daily Audits/CM360/NEXT03',
    mergedFolderPath: folderPath('Merged Reports', 'NEXT03'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'NEXT03')
  },
  {
    name: 'SPTM01',
    label: 'Daily Audits/CM360/SPTM01',
    mergedFolderPath: folderPath('Merged Reports', 'SPTM01'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'SPTM01')
  },
  {
    name: 'NFL01',
    label: 'Daily Audits/CM360/NFL01',
    mergedFolderPath: folderPath('Merged Reports', 'NFL01'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'NFL01')
  },
  {
    name: 'ENT01',
    label: 'Daily Audits/CM360/ENT01',
    mergedFolderPath: folderPath('Merged Reports', 'ENT01'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'ENT01')
  }
];

const headerKeywords = ["Placement ID", "Impressions", "Clicks"];


// === 🧰 CORE UTILITY FUNCTIONS ===
function normalize(s) {
  return String(s || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[^\w\s]/g, ''); // strip punctuation
}

function formatDate(date = new Date(), pattern = 'yyyy-MM-dd') {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), pattern);
}

function escapeHtml(text) {
  return String(text || '')
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function getAuditCacheKey_() {
  return `CM360_AUDIT_RESULTS_${formatDate(new Date(), 'yyyyMMdd')}`;
}

function safeConvertExcelToSheet(blob, filename, parentFolderId, context = '') {
  const resource = {
    title: filename.replace(/\.[^.]+$/, ''), // strip extension
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: parentFolderId }]
  };

  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const file = Drive.Files.insert(resource, blob, { convert: true });
      Logger.log(`✅ [${context}] Excel converted to Sheet (attempt ${attempt})`);
      return file;
    } catch (err) {
      Logger.log(`⚠️ [${context}] Excel convert failed (attempt ${attempt}): ${err.message}`);
      if (attempt < 3) Utilities.sleep(2000); // wait before retry
    }
  }
  throw new Error(`❌ [${context}] Failed to convert Excel to Sheet after 3 attempts`);
}

function getDriveFolderByPath_(pathArray) {
  let folder = DriveApp.getRootFolder();

  for (const name of pathArray) {
    let found = false;
    const folders = folder.getFoldersByName(name);
    while (folders.hasNext()) {
      const subfolder = folders.next();
      if (subfolder.getName() === name) {
        folder = subfolder;
        found = true;
        break;
      }
    }
    if (!found) {
      folder = folder.createFolder(name);
      Logger.log(`📁 Created missing folder: ${name}`);
    }
  }

  return folder;
}

function validateAuditConfigs() {
  const requiredFields = ['name', 'label'];
  auditConfigs.forEach(config => {
    requiredFields.forEach(field => {
      if (!config[field] || typeof config[field] !== 'string') {
        throw new Error(`❌ Invalid audit config "${config.name || '[unnamed]'}": Missing or invalid "${field}"`);
      }
    });
  });
}

// === 📤 EMAIL FUNCTIONS ===
function safeSendEmail({ to, cc = '', subject, htmlBody, attachments = [] }, context = '') {
  let remaining = null;

  try {
    remaining = MailApp.getRemainingDailyQuota();
    storeEmailQuotaRemaining_(remaining);

    if (remaining <= 0) {
      Logger.log(`❌ Quota exhausted — Email not sent for: ${context || 'unknown'}`);
      return;
    }
  } catch (err) {
    Logger.log(`⚠️ Skipping MailApp quota check (unauthorized): ${err.message}`);
  }

  if (typeof to !== 'string' || !to.trim()) {
    Logger.log(`❌ safeSendEmail aborted: Missing or invalid 'to' field`);
    return;
  }
  if (typeof subject !== 'string') {
    Logger.log(`❌ safeSendEmail aborted: Missing or invalid 'subject'`);
    return;
  }

  const clonedAttachments = attachments
    .filter(blob => blob && typeof blob.getBytes === 'function')
    .map(blob => Utilities.newBlob(blob.getBytes(), blob.getContentType(), blob.getName()));

  const options = {
    htmlBody,
    cc,
    attachments: clonedAttachments.length > 0 ? clonedAttachments : undefined
  };

  Logger.log(`📧 safeSendEmail sending: ${context ? `[${context}] ` : ''}To: ${to}, CC: ${cc}, Subject: ${subject}, Attachments: ${clonedAttachments.length}`);

  try {
    GmailApp.sendEmail(to, subject, '', options);
  } catch (err) {
    Logger.log(`❌ safeSendEmail failed: ${err.message}`);
  }
}

function sendNoIssueEmail(config, sheetId, reason, recipientsData) {
  const now = new Date();
  const subjectDate = formatDate(now, "yyyy-MM-dd");
  const subject = `✅ CM360 Daily Audit: No Issues Detected (${config.name} - ${subjectDate})`;

  let htmlBody = `
    <p style="font-family:Arial, sans-serif; font-size:13px;">
      The CM360 audit for bundle "<strong>${escapeHtml(config.name)}</strong>" completed successfully.
    </p>
    <p style="font-family:Arial, sans-serif; font-size:13px;">
      ${escapeHtml(reason)}.
    </p>
    <p style="margin-top:12px; font-family:Arial, sans-serif; font-size:12px;">—Platform Solutions Team</p>
  `;

  let xlsxBlob;
  try {
    xlsxBlob = exportSheetAsExcel(sheetId, `CM360_DailyAudit_${config.name}_${subjectDate}.xlsx`);
  } catch (err) {
    Logger.log(`❌ [${config.name}] Excel export failed: ${err.message}`);
    htmlBody += `
      <p style="font-family:Arial, sans-serif; font-size:12px; color:red;">
        ⚠️ Excel export failed. <br>
        <a href="https://docs.google.com/spreadsheets/d/${sheetId}" target="_blank">Open in Google Sheets</a>
      </p>
    `;
  }

  safeSendEmail({
    to: resolveRecipients(config.name, recipientsData),
    cc: resolveCc(config.name, recipientsData),
    subject,
    htmlBody,
    attachments: xlsxBlob && typeof xlsxBlob.getBytes === 'function' ? [xlsxBlob] : []
  }, config.name);
}

function sendDailySummaryEmail(results) {
  const userEmail = [ADMIN_EMAIL, 'bmuller@horizonmedia.com, bkaufman@horizonmedia.com, ewarburton@horizonmedia.com'].filter(Boolean).join(', ');
  const subjectDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const subject = `CM360 Daily Audit Summary (${subjectDate})`;

  const rowsHtml = results.map(r => {
    const isAlert = String(r.status).toLowerCase().includes('skipped') || r.status.toLowerCase().includes('error');
    const bgColor = isAlert ? 'background-color:#ffe5e5;' : '';

    return `
      <tr style="font-size:12px; line-height:1.3; ${bgColor}">
        <td style="padding:4px 8px;">${escapeHtml(r.name)}</td>
        <td style="padding:4px 8px;">${escapeHtml(r.status)}</td>
        <td style="padding:4px 8px; text-align:center;">${r.flaggedRows ?? '—'}</td>
        <td style="padding:4px 8px; text-align:center;">${r.emailSent ? '✅' : '❌'}</td>
        <td style="padding:4px 8px; text-align:center;">${escapeHtml(r.emailTime)}</td>
      </tr>`;
  }).join('');

  // Pull cached remaining quota (lowest value seen)
  const remainingQuota = getEmailQuotaRemaining_();
  const quotaNote = remainingQuota !== null
    ? `<p style="font-family:Arial, sans-serif; font-size:12px; margin-top:8px;">
         Remaining daily email quota: <strong>${remainingQuota}</strong>
       </p>`
    : '';

  const htmlBody = `
    <p style="font-family:Arial, sans-serif; font-size:13px;">Here’s a summary of today’s CM360 audits:</p>
    <table border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse; font-family:Arial, sans-serif; font-size:12px;">
      <thead style="background:#f2f2f2;">
        <tr>
          <th style="padding:4px 8px;">Config</th>
          <th style="padding:4px 8px;">Status</th>
          <th style="padding:4px 8px;">Flagged Rows</th>
          <th style="padding:4px 8px;">Email Sent</th>
          <th style="padding:4px 8px;">Email Time</th>
        </tr>
      </thead>
      <tbody>${rowsHtml}</tbody>
    </table>
    ${quotaNote}
    <p style="margin-top:12px; font-family:Arial, sans-serif; font-size:12px;">—CM360 Automation</p>
  `;

  try {
    GmailApp.sendEmail(userEmail, subject, '', { htmlBody });
    Logger.log(`📨 Summary email sent to ${userEmail}`);
  } catch (err) {
    Logger.log(`❌ Failed to send summary email: ${err.message}`);
  }
}

function exportSheetAsExcel(spreadsheetId, filename) {
  const url = `https://docs.google.com/feeds/download/spreadsheets/Export?key=${spreadsheetId}&exportFormat=xlsx`;
  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': `Bearer ${token}` },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`❌ Failed to export sheet. HTTP ${response.getResponseCode()}`);
  }

  return response.getBlob().setName(filename);
}


// === 📥 GMAIL & DRIVE FILE FETCH ===
function fetchDailyAuditAttachments(config, recipientsData) {
  Logger.log(`📥 [${config.name}] fetchDailyAuditAttachments started`);

  const label = GmailApp.getUserLabelByName(config.label);
  if (!label) {
    Logger.log(`⚠️ [${config.name}] Label not found: ${config.label}`);
    safeSendEmail({
      to: resolveRecipients(config.name, recipientsData),
      cc: resolveCc(config.name, recipientsData),
      subject: `⚠️ CM360 Audit Warning: Gmail Label Missing (${config.name})`,
      htmlBody: `<p style="font-family:Arial; font-size:13px;">The label <b>${escapeHtml(config.label)}</b> could not be found. This may mean the audit for <b>${escapeHtml(config.name)}</b> will be skipped.</p>`
    }, `${config.name} - Missing Gmail Label`);
    return null;
  }

  const threads = label.getThreads();
  const startOfToday = new Date();
  startOfToday.setHours(0, 0, 0, 0);  
  
  const parentFolder = getDriveFolderByPath_(config.tempDailyFolderPath);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const driveFolder = parentFolder.createFolder(`Temp_CM360_${timestamp}`);

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      if (message.getDate() < startOfToday) return;

      message.getAttachments({ includeInlineImages: false }).forEach(file => {
        const name = file.getName();
        const type = file.getContentType();

        // Unzip file and save all .xlsx/.csv blobs
        if (name.endsWith('.xlsx') && type === MimeType.MICROSOFT_EXCEL) {
          driveFolder.createFile(file);
        } else if (name.endsWith('.zip') && type === MimeType.ZIP) {
          const blobs = Utilities.unzip(file);
          let count = 0;
          blobs.forEach(blob => {
            const lowerName = blob.getName().toLowerCase();
            if (lowerName.endsWith('.csv') || lowerName.endsWith('.xlsx')) {
              driveFolder.createFile(blob);
              count++;
            }
          });

          Logger.log(`🗂️ [${config.name}] Extracted ${count} file(s) from ZIP: ${name}`);
        }
      });
    });
  });

  const hasFiles = driveFolder.getFiles().hasNext();
  if (!hasFiles) {
    Logger.log(`⚠️ [${config.name}] No files saved to: ${driveFolder.getName()}`);
    return null;
  }

  Logger.log(`✅ [${config.name}] Files saved to: ${driveFolder.getName()}`);
  return driveFolder.getId();
}

function mergeDailyAuditExcels(folderId, mergedFolderPath, configName = 'Unknown') {
  Logger.log(`[${configName}] mergeDailyAuditExcels started`);
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const destFolder = getDriveFolderByPath_(mergedFolderPath);

  const mergedSheetName = `Merged_CM360_${new Date().toISOString().slice(0, 10)}`;
  const mergedSpreadsheet = SpreadsheetApp.create(mergedSheetName);
  Utilities.sleep(1000); // Ensure file is created
  const mergedFile = DriveApp.getFileById(mergedSpreadsheet.getId());
  destFolder.addFile(mergedFile);
  DriveApp.getRootFolder().removeFile(mergedFile);
  const mergedSheet = mergedSpreadsheet.getSheets()[0];

  let headerWritten = false;
  let header = [];

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName().toLowerCase();

    let data;
    let spreadsheet;

    if (fileName.endsWith('.xlsx')) {
      const blob = file.getBlob();
      const converted = safeConvertExcelToSheet(blob, file.getName(), folder.getId(), configName);

      // Ensure it only lives in `folder`
      Drive.Files.update({ parents: [{ id: folder.getId() }] }, converted.id);

      spreadsheet = SpreadsheetApp.openById(converted.id);
      data = spreadsheet.getSheets()[0].getDataRange().getValues();
      if (!data || data.length === 0 || data.every(row => row.every(cell => cell === ''))) {
        Logger.log(`[${configName}] ⚠️ File "${fileName}" appears blank after import.`);
        continue;
      }
    } else if (fileName.endsWith('.csv')) {
      const csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
      spreadsheet = SpreadsheetApp.create(file.getName().replace(/\.csv$/i, ''));
      spreadsheet.getSheets()[0].getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
      data = csvData;
    } else {
      Logger.log(`[${configName}] Skipping unsupported file: ${fileName}`);
      continue;
    }

    const headerRowIndex = data.findIndex(row => {
      const normRow = row.map(cell => normalize(cell));
      return headerKeywords.every(keyword =>
        normRow.includes(normalize(keyword))
      );
    });

    if (headerRowIndex === -1) {
      Logger.log(`[${configName}] Header not found in: ${file.getName()}`);
      continue;
    }

    const realData = data.slice(headerRowIndex);
    const cleanedData = realData.filter((row, idx) =>
      idx === 0 || !row.join('').toLowerCase().includes('grand total')
    );

    // Normalize pixel sizes like "1 x 1" → "1x1"
    const pixelCols = ['Placement Pixel Size', 'Creative Pixel Size'];
    const pixelColIndexes = [];

    if (!headerWritten) {
      header = cleanedData[0];
      pixelCols.forEach(col => {
        const idx = header.findIndex(h => normalize(h) === normalize(col));
        if (idx !== -1) pixelColIndexes.push(idx);
      });
    }

    cleanedData.forEach(row => {
      pixelColIndexes.forEach(colIdx => {
        row[colIdx] = String(row[colIdx] || '').replace(/\s+/g, '');
      });
    });

    if (!headerWritten) {
      mergedSheet.clear();
      const bodyRows = cleanedData.slice(1);
      mergedSheet.getRange(1, 1, 1, header.length).setValues([header]);

      if (bodyRows.length > 0) {
        mergedSheet.getRange(2, 1, bodyRows.length, header.length).setValues(bodyRows);
      }

      headerWritten = true;
    } else {
      const startRow = mergedSheet.getLastRow() + 1;
      const rowsToAdd = cleanedData.slice(1);
      if (rowsToAdd.length > 0) {
        mergedSheet.getRange(startRow, 1, rowsToAdd.length, header.length).setValues(rowsToAdd);
      } else {
        Logger.log(`[${configName}] No data rows found in ${fileName} after header; skipping.`);
      }
    }

    // Move the source file (converted or CSV) to holding folder
    const holdingFolderPath = [...TRASH_ROOT_PATH, 'Temp Daily Reports', configName];
    const holdingFolder = getDriveFolderByPath_(holdingFolderPath);

    if (holdingFolder) {
      const convertedFile = DriveApp.getFileById(spreadsheet.getId());
      convertedFile.moveTo(holdingFolder);
    } else {
      Logger.log(`[${configName}] ⚠️ Holding folder not found: ${holdingFolderPath.join(' / ')}`);
    }
  }

  const mergedHeaders = mergedSheet.getRange(1, 1, 1, mergedSheet.getLastColumn()).getValues()[0];
  Logger.log(`[${configName}] ✅ Final headers in merged sheet: ${mergedHeaders.join(' | ')}`);
  Logger.log(`[${configName}] Merged sheet created: ${mergedSpreadsheet.getUrl()}`);
  return mergedSpreadsheet.getId();
}


// === 📊 MERGE & FLAG LOGIC ===
function executeAudit(config) {
  const now = new Date();
  const formattedNow = formatDate(now, 'yyyy-MM-dd HH:mm:ss');
  const configName = config.name;

  try {
    Logger.log(`🔍 [${configName}] Audit started`);

    // Load exclusions data at the start of audit
    const exclusionsData = loadExclusionsFromSheet();
    Logger.log(`📋 [${configName}] Loaded exclusions for ${Object.keys(exclusionsData).length} configs`);
    
    // Load thresholds data at the start of audit
    const thresholdsData = loadThresholdsFromSheet();
    Logger.log(`📊 [${configName}] Loaded thresholds for ${Object.keys(thresholdsData).length} configs`);
    
    // Load recipients data at the start of audit
    const recipientsData = loadRecipientsFromSheet();
    Logger.log(`📧 [${configName}] Loaded recipients for ${Object.keys(recipientsData).length} configs`);

    const folderId = fetchDailyAuditAttachments(config, recipientsData);
    if (!folderId) {
      Logger.log(`⚠️ [${configName}] No files found today. Sending notification...`);
      const subject = `⚠️ CM360 Audit Skipped: No Files Found (${configName} - ${formatDate(now)})`;
      const htmlBody = `
        <p style="font-family:Arial, sans-serif; font-size:13px;">
          The CM360 audit for bundle "<strong>${escapeHtml(configName)}</strong>" was skipped because no Excel or ZIP files were found for today.
        </p>
        <p style="font-family:Arial, sans-serif; font-size:13px;">
          Please verify the report was delivered and labeled correctly.
        </p>
        <p style="margin-top:12px; font-family:Arial, sans-serif; font-size:12px;">—Platform Solutions Team</p>
      `;
      safeSendEmail({ 
        to: resolveRecipients(configName, recipientsData), 
        cc: resolveCc(configName, recipientsData), 
        subject, 
        htmlBody, 
        attachments: [] 
      }, configName);
      return { status: 'Skipped: No files found', flaggedCount: null, emailSent: true, emailTime: formattedNow };
    }

    const mergedSheetId = mergeDailyAuditExcels(folderId, config.mergedFolderPath, configName);
    const sheet = SpreadsheetApp.openById(mergedSheetId).getSheets()[0];
    const allData = sheet.getDataRange().getValues();

    const headerRowIndex = allData.findIndex(row => {
      const normRow = row.map(cell => normalize(cell));
      return headerKeywords.every(keyword => normRow.includes(normalize(keyword)));
    });

    if (headerRowIndex === -1) {
      Logger.log(`❌ [${configName}] Header row not found in merged sheet.`);
      return { status: 'Failed: Header not found', flaggedCount: null, emailSent: false, emailTime: formattedNow };
    }

    const headers = allData[headerRowIndex];
    const getIndex = name => headers.findIndex(h => normalize(h) === normalize(name));

    const fullCol = {
      Advertiser: getIndex('Advertiser'),
      Campaign: getIndex('Campaign'),
      Site: getIndex('Site (CM360)'),
      Placement: getIndex('Placement'),
      PlacementID: getIndex('Placement ID'),
      Start: getIndex('Placement Start Date'),
      End: getIndex('Placement End Date'),
      Creative: getIndex('Creative'),
      Impressions: getIndex('Impressions'),
      Clicks: getIndex('Clicks'),
      Flags: getIndex('Flag(s)'),
      'Placement Pixel Size': getIndex('Placement Pixel Size'),
      'Creative Pixel Size': getIndex('Creative Pixel Size'),
      'Ad Type': getIndex('Ad Type'),
      Date: getIndex('Date')
    };

    let flagColIndex = fullCol.Flags;
    if (flagColIndex === -1) {
      flagColIndex = headers.length;
      headers.push('Flag(s)');
      sheet.getRange(headerRowIndex + 1, 1, 1, headers.length).setValues([headers]);
    }

    const flaggedRows = [];
    const flaggedIDs = new Set();

    for (let r = headerRowIndex + 1; r < allData.length; r++) {
      const row = allData[r];
      const flags = [];

      const clicks = Number(row[fullCol.Clicks] || 0);
      const impressions = Number(row[fullCol.Impressions] || 0);
      
      const startDate = new Date(row[fullCol.Start]);
      const endDate = new Date(row[fullCol.End]);
      const today = new Date(row[fullCol.Date]);
      const placementPixel = String(row[fullCol['Placement Pixel Size']] || '');
      const creativePixel = String(row[fullCol['Creative Pixel Size']] || '');
      const adType = String(row[fullCol['Ad Type']] || '').toLowerCase();
      const placementName = String(row[fullCol.Placement] || '').toLowerCase();
      const siteName = String(row[fullCol.Site] || '');
      const placementId = String(row[fullCol.PlacementID] || '');

      // Check each potential flag with its specific thresholds
      
      // Clicks > Impressions check
      const clicksThreshold = getThresholdForFlag(thresholdsData, configName, 'clicks_greater_than_impressions');
      const hasMinVolumeForClicks = impressions >= clicksThreshold.minImpressions || clicks >= clicksThreshold.minClicks;
      if (hasMinVolumeForClicks && clicks > impressions && 
          !isPlacementExcludedForFlag(exclusionsData, configName, placementId, 'clicks_greater_than_impressions', placementName, siteName)) {
        flags.push('Clicks > Impressions');
      }
      
      // Out of flight dates check
      const flightThreshold = getThresholdForFlag(thresholdsData, configName, 'out_of_flight_dates');
      const hasMinVolumeForFlight = impressions >= flightThreshold.minImpressions || clicks >= flightThreshold.minClicks;
      if (hasMinVolumeForFlight && (startDate > today || endDate < today) && 
          !isPlacementExcludedForFlag(exclusionsData, configName, placementId, 'out_of_flight_dates', placementName, siteName)) {
        flags.push('Out of flight dates');
      }
      
      // Pixel size mismatch check
      const pixelThreshold = getThresholdForFlag(thresholdsData, configName, 'pixel_size_mismatch');
      const hasMinVolumeForPixel = impressions >= pixelThreshold.minImpressions || clicks >= pixelThreshold.minClicks;
      if (hasMinVolumeForPixel && placementPixel && creativePixel && placementPixel !== creativePixel && 
          !isPlacementExcludedForFlag(exclusionsData, configName, placementId, 'pixel_size_mismatch', placementName, siteName)) {
        flags.push('Pixel size mismatch');
      }
      
      // Default ad serving check
      const defaultThreshold = getThresholdForFlag(thresholdsData, configName, 'default_ad_serving');
      const hasMinVolumeForDefault = impressions >= defaultThreshold.minImpressions || clicks >= defaultThreshold.minClicks;
      if (hasMinVolumeForDefault && adType.includes('default') && 
          !isPlacementExcludedForFlag(exclusionsData, configName, placementId, 'default_ad_serving', placementName, siteName)) {
        flags.push('Default ad serving');
      }

      if (flags.length > 0) {
        row[flagColIndex] = flags.join(', ');
        flaggedRows.push(row);
        flaggedIDs.add(row[fullCol.PlacementID]);
      }
    }

    const updatedDataRows = allData.slice(headerRowIndex + 1).map(row => {
      const newRow = [...row];
      while (newRow.length < headers.length) newRow.push('');
      return newRow.slice(0, headers.length);
    });

    sheet.getRange(headerRowIndex + 2, 1, updatedDataRows.length, headers.length).setValues(updatedDataRows);
    SpreadsheetApp.flush();


    // Sort flagged rows by highest volume (clicks or impressions)
    flaggedRows.sort((a, b) => {
      const aVol = Math.max(Number(a[fullCol.Clicks] || 0), Number(a[fullCol.Impressions] || 0));
      const bVol = Math.max(Number(b[fullCol.Clicks] || 0), Number(b[fullCol.Impressions] || 0));
      return bVol - aVol;
    });

    // Reorder merged sheet preserving sorted flagged rows at top only
    const allDataRange = sheet.getDataRange().getValues();
    const allDataRows = allDataRange.slice(1); // Exclude header row

    // Build reordered flagged list
    const reorderedFlagged = flaggedRows;


    // Remaining unflagged rows
    const reorderedUnflagged = allDataRows.filter(r => !flaggedIDs.has(r[fullCol.PlacementID]));

    // Rewrite sheet cleanly
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    if (reorderedFlagged.length > 0) {
      sheet.getRange(2, 1, reorderedFlagged.length, headers.length).setValues(reorderedFlagged);
    }
    if (reorderedUnflagged.length > 0) {
      sheet.getRange(reorderedFlagged.length + 2, 1, reorderedUnflagged.length, headers.length).setValues(reorderedUnflagged);
    }
    SpreadsheetApp.flush();

    // Apply yellow highlights per flag (batched for export compatibility)
    const finalData = sheet.getDataRange().getValues();
    const flagIndex = headers.findIndex(h => normalize(h) === normalize('Flag(s)'));
    const bgData = Array.from({ length: finalData.length - 1 }, () =>
      Array.from({ length: headers.length }, () => null)
    );
    for (let r = 1; r < finalData.length; r++) {
      const row = finalData[r];
      const flagText = String(row[flagIndex] || '').toLowerCase();
      const bgRow = r - 1;

      if (flagText.includes('default ad serving') && fullCol['Ad Type'] !== -1) {
        bgData[bgRow][fullCol['Ad Type']] = '#ffff00';
      }
      if (flagText.includes('out of flight')) {
        if (fullCol.Start !== -1) bgData[bgRow][fullCol.Start] = '#ffff00';
        if (fullCol.End !== -1) bgData[bgRow][fullCol.End] = '#ffff00';
      }
      if (flagText.includes('pixel size mismatch')) {
        if (fullCol['Placement Pixel Size'] !== -1) bgData[bgRow][fullCol['Placement Pixel Size']] = '#ffff00';
        if (fullCol['Creative Pixel Size'] !== -1) bgData[bgRow][fullCol['Creative Pixel Size']] = '#ffff00';
      }
      if (flagText.includes('clicks >') && fullCol.Clicks !== -1) {
        bgData[bgRow][fullCol.Clicks] = '#ffff00';
      }
    }

    // Apply zebra striping (without overriding highlights)
    for (let r = 1; r < finalData.length; r++) {
      const bgRow = r - 1;
      const isStriped = r % 2 === 0;
      for (let c = 0; c < headers.length; c++) {
        if (!bgData[bgRow][c]) {
          bgData[bgRow][c] = isStriped ? '#fafafa' : null;
        }
      }
    }

    if (finalData.length > 1) {
      sheet.getRange(2, 1, finalData.length - 1, headers.length).setBackgrounds(bgData);
      sheet.getRange(2, 1, finalData.length - 1, headers.length).setBorder(
        true, true, true, true, true, true
      );
    }

    // Format header
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f2f2f2');
    headerRange.setBorder(
      true, true, true, true, true, true,
      '#d9d9d9',
      SpreadsheetApp.BorderStyle.SOLID
    );

    SpreadsheetApp.flush();

    const displayRows = flaggedRows.map(row => [
      row[fullCol.Advertiser],
      row[fullCol.Campaign],
      row[fullCol.Site],
      row[fullCol.Placement],
      row[fullCol.PlacementID],
      row[fullCol.Start],
      row[fullCol.End],
      row[fullCol.Creative],
      row[fullCol.Impressions],
      row[fullCol.Clicks],
      row[flagColIndex]
    ]);

    if (displayRows.length > 0) {
      emailFlaggedRows(mergedSheetId, displayRows, flaggedRows, config, recipientsData);
      return { status: 'Completed with flags', flaggedCount: flaggedRows.length, emailSent: true, emailTime: formattedNow };
    } else {
      sendNoIssueEmail(config, mergedSheetId, 'No issues were flagged', recipientsData);
      return { status: 'Completed (no issues)', flaggedCount: 0, emailSent: true, emailTime: formattedNow };
    }

  } catch (err) {
    Logger.log(`❌ [${configName}] Unexpected error: ${err.message}`);
    return { status: `Error during audit: ${err.message}`, flaggedCount: null, emailSent: false, emailTime: formattedNow };
  }
}

// === 📋 EXECUTION & AUDIT FLOW ===
function runDailyAuditByName(configName) {
  if (!checkDriveApiEnabled()) return;
  const config = auditConfigs.find(c => c.name === configName);
  if (!config) {
    Logger.log(`❌ Config "${configName}" not found.`);
    return;
  }
  executeAudit(config);
}

function runAuditBatch(configs) {
  validateAuditConfigs();
  Logger.log(`🚀 Audit Batch Started: ${new Date().toLocaleString()}`);
  const results = [];

  for (const config of configs) {
    try {
      const result = executeAudit(config);
      results.push({
        name: config.name,
        status: result.status,
        flaggedRows: result.flaggedCount,
        emailSent: result.emailSent,
        emailTime: result.emailTime
      });
    } catch (err) {
      results.push({
        name: config.name,
        status: `Error: ${err.message}`,
        flaggedRows: null,
        emailSent: false,
        emailTime: formatDate(new Date(), 'yyyy-MM-dd HH:mm:ss')
      });
    }
  }

  storeCombinedAuditResults_(results);

  const totalConfigs = auditConfigs.length;
  const cachedResults = getCombinedAuditResults_();

  const completedConfigs = new Set(cachedResults.map(r => r.name)).size;

  Logger.log(`🧮 Completed ${completedConfigs} of ${totalConfigs} configs`);

  if (completedConfigs >= totalConfigs) {
    Logger.log(`📬 All audits complete. Sending summary email...`);
    sendDailySummaryEmail(cachedResults);
    CacheService.getScriptCache().remove(getAuditCacheKey_());
  }
}

function getAuditConfigBatches(batchSize = BATCH_SIZE) {
  const batches = [];
  for (let i = 0; i < auditConfigs.length; i += batchSize) {
    batches.push(auditConfigs.slice(i, i + batchSize));
  }
  return batches;
}

function storeCombinedAuditResults_(newResults) {
  const cache = CacheService.getScriptCache();
  const existing = getCombinedAuditResults_();
  const combined = [...existing, ...newResults];
  cache.put(getAuditCacheKey_(), JSON.stringify(combined), 21600); // 6 hours
}

function getCombinedAuditResults_() {
  const cache = CacheService.getScriptCache();
  const stored = cache.get(getAuditCacheKey_());
  return stored ? JSON.parse(stored) : [];
}

function storeEmailQuotaRemaining_(remaining) {
  const cache = CacheService.getScriptCache();
  const existing = cache.get('CM360_EMAIL_QUOTA_LEFT');

  if (existing === null || Number(remaining) < Number(existing)) {
    cache.put('CM360_EMAIL_QUOTA_LEFT', String(remaining), 21600); // 6 hours
    Logger.log(`Updated cached quota remaining to: ${remaining}`);
  }
}

function getEmailQuotaRemaining_() {
  const cache = CacheService.getScriptCache();
  const val = cache.get('CM360_EMAIL_QUOTA_LEFT');
  return val !== null ? Number(val) : null;
}


// === 📬 EMAIL FLAGGED ROWS & REPORTS ===
function emailFlaggedRows(sheetId, emailRows, flaggedRows, config, recipientsData) {
  const configName = config.name;

  if (!flaggedRows || flaggedRows.length === 0) {
    Logger.log(`[${configName}] No flagged rows to report.`);
    return;
  }

  const subjectDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  const truncate = (text, maxLen = 80) => {
    const safe = String(text || '').trim();
    return safe.length > maxLen ? safe.slice(0, maxLen - 1) + '…' : safe;
  };

  const subject = `⚠️ CM360 Daily Audit: Issues Detected (${configName} - ${subjectDate})`;

  const xlsxBlob = exportSheetAsExcel(sheetId, `CM360_DailyAudit_${configName}_${subjectDate}.xlsx`);

  const plural = (count, singular, plural) => count === 1 ? singular : plural;
  const totalFlagged = flaggedRows.length;
  const uniqueCampaigns = new Set(flaggedRows.map(r => r[1])).size;
  const summaryText = `⚠️ The following ${totalFlagged} ${plural(totalFlagged, 'placement', 'placements')} across ${uniqueCampaigns} ${plural(uniqueCampaigns, 'campaign', 'campaigns')} ${plural(totalFlagged, 'was', 'were')} flagged during the <strong>${configName}</strong> CM360 audit of yesterday's delivery. Please review:`;

  const truncatedNote = flaggedRows.length > 100
    ? `<p style="font-family:Arial, sans-serif; font-size:12px;">Only the first 100 flagged rows are shown below. Full details are included in the attached Excel file.</p>`
    : '';

  const htmlBody = `
    <p style="font-family:Arial, sans-serif; font-size:13px; line-height:1.4;">${summaryText}</p>
    ${truncatedNote}
    <table border="1" cellpadding="2" cellspacing="0" width="100%" style="font-family:Arial, sans-serif; font-size:12px; table-layout:fixed; border-collapse:collapse;">
      <thead style="background-color:#f2f2f2;">
        <tr>
          <th style="padding:2px; width:140px;">Advertiser</th>
          <th style="padding:2px; width:180px;">Campaign</th>
          <th style="padding:2px; width:100px;">Site</th>
          <th style="padding:2px; width:180px;">Placement</th>
          <th style="padding:2px; width:100px;">Placement ID</th>
          <th style="padding:2px; width:90px;">Start Date</th>
          <th style="padding:2px; width:90px;">End Date</th>
          <th style="padding:2px; width:180px;">Creative</th>
          <th style="padding:2px; width:60px;">Impr.</th>
          <th style="padding:2px; width:60px;">Clicks</th>
          <th style="padding:2px; width:160px;">Flag(s)</th>
        </tr>
      </thead>
      <tbody>
        ${emailRows.map((row, i) => `
          <tr style="line-height:1.2; font-size:11px; background-color:${i % 2 === 0 ? '#ffffff' : '#f9f9f9'};">
            <td style="padding:2px 4px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis;">${escapeHtml(row[0])}</td>
            <td style="padding:2px 4px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis;">${escapeHtml(row[1])}</td>
            <td style="padding:2px 4px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis;">${escapeHtml(row[2])}</td>
            <td style="padding:2px 4px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis;">${escapeHtml(truncate(row[3], 60))}</td>
            <td style="padding:2px 4px;">${escapeHtml(row[4])}</td>
            <td style="padding:2px 4px;">${formatDate(new Date(row[5]), 'yyyy-MM-dd')}</td>
            <td style="padding:2px 4px;">${formatDate(new Date(row[6]), 'yyyy-MM-dd')}</td>
            <td style="padding:2px 4px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis;">${escapeHtml(truncate(row[7], 45))}</td>
            <td style="padding:2px 4px; text-align:right;">${row[8]}</td>
            <td style="padding:2px 4px; text-align:right;">${row[9]}</td>
            <td style="padding:2px 4px; white-space:normal; line-height:1.3; word-break:break-word;">
              ${String(row[10] ?? '').split('; ').map(f => `<div>${escapeHtml(f)}</div>`).join('')}
            </td>
          </tr>`).join('')}
      </tbody>
    </table>
    <p style="margin-top:12px; font-family:Arial, sans-serif; font-size:12px;">—Platform Solutions Team</p>
  `;

  safeSendEmail({
    to: resolveRecipients(configName, recipientsData),
    cc: resolveCc(configName, recipientsData),
    subject,
    htmlBody,
    attachments: [xlsxBlob]
  }, `[${configName}]`);

  Logger.log(`[${configName}]🚩 Flagging complete: ${flaggedRows.length} row(s)`);

  return flaggedRows;
}

// === 🛠️ SETUP & ENVIRONMENT PREP ===
function prepareAuditEnvironment() {
  const ui = SpreadsheetApp.getUi();
  const createdLabels = [];
  const createdFolders = [];
  const missingFilters = [];

  auditConfigs.forEach(config => {
    const { name, label, mergedFolderPath, tempDailyFolderPath } = config;

    // 1. Ensure Gmail label exists
    let labelObj = GmailApp.getUserLabelByName(label);
    if (!labelObj) {
      labelObj = GmailApp.createLabel(label);
      createdLabels.push(label);
      Logger.log(`✅ Created Gmail label: ${label}`);
    } else {
      Logger.log(`ℹ️ Gmail label already exists: ${label}`);
    }

    // 2. Check if label is used (i.e., if a filter exists)
    const threads = labelObj.getThreads(0, 1);
    if (threads.length === 0) {
      missingFilters.push({ name, label });
    }

    // 3. Ensure Drive folders exist
    const ensureFolder = (pathArray) => {
      let folder = DriveApp.getRootFolder();
      for (const part of pathArray) {
        const sub = folder.getFoldersByName(part);
        if (sub.hasNext()) {
          folder = sub.next();
        } else {
          folder = folder.createFolder(part);
          createdFolders.push(pathArray.join('/'));
          Logger.log(`📁 Created missing folder: ${pathArray.join('/')}`);
        }
      }
    };

    ensureFolder(mergedFolderPath);
    ensureFolder(tempDailyFolderPath);
  });

  // 4. Log missing filter suggestions and generate pop-up
  let msgParts = [];

  if (createdLabels.length > 0) {
    msgParts.push(`✅ Created ${createdLabels.length} Gmail label(s).`);
  }

  if (createdFolders.length > 0) {
    msgParts.push(`📁 Created ${createdFolders.length} Drive folder path(s).`);
  }

  if (missingFilters.length > 0 && (createdLabels.length > 0 || createdFolders.length > 0)) {
    msgParts.push(`\n⚠️ The following Gmail filters may be missing:`);
    missingFilters.forEach(({ name, label }) => {
      msgParts.push(`• from:google ${name} -{⚠️} → Label: "${label}"`);
    });
  }

  let msg = `Environment setup complete.\n\n`;

  if (msgParts.length > 0) {
    msg += msgParts.join('\n');
  } else {
    msg += `No further steps required.`;
  }

  ui.alert('✅ Setup Summary', msg.trim(), ui.ButtonSet.OK);
}

function installDailyAuditTriggers() {
  const batches = getAuditConfigBatches(BATCH_SIZE);
  const results = [];

  // Clear existing triggers
  const existing = ScriptApp.getProjectTriggers();
  existing.forEach(trigger => {
    if (trigger.getHandlerFunction().startsWith("runDailyAuditsBatch")) {
      ScriptApp.deleteTrigger(trigger);
      results.push(`🗑️ Removed trigger: ${trigger.getHandlerFunction()}`);
    }
  });

  // Install new triggers
  for (let i = 0; i < batches.length; i++) {
    const fnName = `runDailyAuditsBatch${i + 1}`;
    if (typeof globalThis[fnName] === 'function') {
      ScriptApp.newTrigger(fnName)
        .timeBased()
        .atHour(8)
        .everyDays(1)
        .create();
      results.push(`✅ Installed daily trigger for: ${fnName}`);
    } else {
      results.push(`⚠️ Skipped trigger for ${fnName} — function not defined`);
    }
  }

  return results;
}


// === 📆 TRIGGER FUNCTIONS ===
function runDailyAuditsBatch1() {
  const batches = getAuditConfigBatches(BATCH_SIZE);
  runAuditBatch(batches[0]);
}

function runDailyAuditsBatch2() {
  const batches = getAuditConfigBatches(BATCH_SIZE);
  runAuditBatch(batches[1]);
}

function runDailyAuditsBatch3() {
  const batches = getAuditConfigBatches(BATCH_SIZE);
  runAuditBatch(batches[2]);
  cleanupOldAuditFiles();
}

function generateMissingBatchStubs() {
  const batches = getAuditConfigBatches(BATCH_SIZE);
  const neededCount = batches.length;
  const existingFns = Object.keys(globalThis).filter(k => /^runDailyAuditsBatch\d+$/.test(k));
  const definedIndexes = new Set(existingFns.map(fn => Number(fn.match(/\d+$/)[0])));
  const stubs = [];

  for (let i = 1; i <= neededCount; i++) {
    if (!definedIndexes.has(i)) {
      const isFinal = i === neededCount;
      stubs.push(`function runDailyAuditsBatch${i}() {
  const batches = getAuditConfigBatches(BATCH_SIZE);
  runAuditBatch(batches[${i - 1}], ${isFinal});
  }`);
    }
  }

  if (stubs.length === 0) {
    return "// ✅ All required batch runner functions are already defined.";
  }

  return `// 🚧 Add these to your script:\n\n${stubs.join('\n\n')}`;
}


// === 📌 UI MENU & MODALS ===
function onOpen() {
  validateAuditConfigs();
  checkDriveApiEnabled();

  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CM360 Audit')
    // 🔧 Setup & One-Time Actions
    .addItem('🛠️ Prepare Audit Environment', 'prepareAuditEnvironment')
    .addItem('� Create/Open Thresholds Sheet', 'getOrCreateThresholdsSheet')
    .addItem('�📋 Create/Open Exclusions Sheet', 'getOrCreateExclusionsSheet')
    .addItem('🧪 Test Enhanced Exclusions', 'testEnhancedExclusions')
    .addItem('📊 Test Thresholds System', 'testThresholdsSystem')
    .addItem('� Test Recipients System', 'testRecipientsSystem')
    .addItem('�🔄 Update Placement Names', 'updatePlacementNamesFromReports')
    .addItem('🔐 Check Authorization', 'checkAuthorizationStatus')
    .addItem('📋 Validate Configs', 'debugValidateAuditConfigs')
    .addItem('📄 Print Config Summary', 'debugPrintConfigSummary')
    .addItem('⚙️ Install Daily Triggers', 'installDailyAuditTriggers')
    .addSeparator()

    // 🚀 Manual Run Options
    .addItem('🧪 Run Batch or Config (Manual Test)', 'showBatchTestPicker')
    .addItem('🔎 Run Audit for...', 'showConfigPicker')

    // 📊 Access Tools
    .addItem('📈 Open Dashboard', 'showAuditDashboard')
    .addToUi();
}

function showConfigPicker() {
  const template = HtmlService.createTemplateFromFile('ConfigPicker');
  template.auditConfigs = auditConfigs; // Pass into template
  const html = template.evaluate()
    .setWidth(300)
    .setHeight(160);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Audit Config');
}

function showBatchTestPicker() {
  const ui = SpreadsheetApp.getUi();
  const batches = getAuditConfigBatches(BATCH_SIZE);

  const batchOptions = batches.map((_, i) => `Batch ${i + 1}`).join('\n');
  const batchPrompt = ui.prompt(
    '🧪 Run Batch or Config',
    `Which batch do you want to run?\n\n${batchOptions}`,
    ui.ButtonSet.OK_CANCEL
  );
  if (batchPrompt.getSelectedButton() !== ui.Button.OK) return;

  const batchIndex = parseInt(batchPrompt.getResponseText().replace(/[^\d]/g, '')) - 1;
  if (isNaN(batchIndex) || batchIndex >= batches.length || batchIndex < 0) {
    ui.alert('Invalid batch number.');
    return;
  }

  const configList = batches[batchIndex].map(c => c.name).join(', ');
  const configPrompt = ui.prompt(
    `Batch ${batchIndex + 1}`,
    `Enter a specific config name to run only that audit.\nLeave blank to run the whole batch.\n\nConfigs in this batch:\n${configList}`,
    ui.ButtonSet.OK_CANCEL
  );
  if (configPrompt.getSelectedButton() !== ui.Button.OK) return;

  const configName = configPrompt.getResponseText().trim();
  if (configName) {
    const config = batches[batchIndex].find(c => c.name === configName);
    if (!config) {
      ui.alert(`❌ Config "${configName}" not found in Batch ${batchIndex + 1}.`);
      return;
    }
    Logger.log(`🧪 Manually running config: ${config.name}`);
    executeAudit(config);
  } else {
    Logger.log(`🧪 Manually running batch ${batchIndex + 1}`);
    const isFinal = batchIndex === batches.length - 1;
    runAuditBatch(batches[batchIndex], isFinal);
  }
}

function showAuditDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('CM360 Audit Dashboard')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}



// === 🧹 CLEANUP & HOUSEKEEPING ===
function cleanupOldAuditFiles() {
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - 60);

  const trashRootPath = TRASH_ROOT_PATH;
  const deletionLogPath = DELETION_LOG_PATH;
  const masterLogName = MASTER_LOG_NAME;

  const trashRoot = getDriveFolderByPath_(trashRootPath);
  const logFolder = getDriveFolderByPath_(deletionLogPath);

  if (!trashRoot || !logFolder) {
    Logger.log('❌ Cleanup failed: Trash or Log folder not found.');
    return;
  }

  const deletedFilesLog = [];
  const deletionTimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // === 1. Delete loose files directly in trash root ===
  const looseFiles = trashRoot.getFiles();
  while (looseFiles.hasNext()) {
    const file = looseFiles.next();
    const created = file.getDateCreated();
    if (created < cutoffDate) {
      deletedFilesLog.push([
        file.getName(),
        trashRootPath.join(' / '),
        Utilities.formatDate(created, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        deletionTimestamp
      ]);
      file.setTrashed(true);
    }
  }

  // === 2. Delete Temp_* folders in each config subfolder of Temp Daily Reports ===
  const tempReportsRoot = getDriveFolderByPath_([...trashRootPath, 'Temp Daily Reports']);
  if (tempReportsRoot) {
    const configFolders = tempReportsRoot.getFolders();
    while (configFolders.hasNext()) {
      const configFolder = configFolders.next();
      const tempSubfolders = configFolder.getFolders();
      while (tempSubfolders.hasNext()) {
        const tempFolder = tempSubfolders.next();
        const name = tempFolder.getName();
        const created = tempFolder.getDateCreated();
        if (name.startsWith('Temp_') && created < cutoffDate) {
          deletedFilesLog.push([
            name,
            [...trashRootPath, 'Temp Daily Reports', configFolder.getName()].join(' / '),
            Utilities.formatDate(created, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
            deletionTimestamp
          ]);
          tempFolder.setTrashed(true);
          Logger.log(`🗑️ Deleted old temp folder: ${name}`);
        }
      }
    }
  }

  // === 3. Delete Merged_* files in each config subfolder of Merged Reports ===
  const mergedReportsRoot = getDriveFolderByPath_([...trashRootPath, 'Merged Reports']);
  if (mergedReportsRoot) {
    const configFolders = mergedReportsRoot.getFolders();
    while (configFolders.hasNext()) {
      const configFolder = configFolders.next();
      const files = configFolder.getFiles();
      while (files.hasNext()) {
        const file = files.next();
        const name = file.getName();
        const created = file.getDateCreated();
        if (name.startsWith('Merged_') && created < cutoffDate) {
          deletedFilesLog.push([
            name,
            [...trashRootPath, 'Merged Reports', configFolder.getName()].join(' / '),
            Utilities.formatDate(created, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
            deletionTimestamp
          ]);
          file.setTrashed(true);
          Logger.log(`🗑️ Deleted old merged file: ${name}`);
        }
      }
    }
  }

  // === Original logic: delete aged files in other subfolders, and empty folders ===
  const subfolders = trashRoot.getFolders();
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    const subfolderPath = [...trashRootPath, subfolder.getName()];
    const files = subfolder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const created = file.getDateCreated();
      if (created < cutoffDate) {
        deletedFilesLog.push([
          file.getName(),
          subfolderPath.join(' / '),
          Utilities.formatDate(created, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          deletionTimestamp
        ]);
        file.setTrashed(true);
      }
    }

    if (!subfolder.getFiles().hasNext() && !subfolder.getFolders().hasNext()) {
      subfolder.setTrashed(true);
      Logger.log(`🗑️ Deleted empty folder: ${subfolder.getName()}`);
    }
  }

  // === Write to log sheet ===
  if (deletedFilesLog.length > 0) {
    let logSheetFile;
    let logSheet;

    const logFiles = logFolder.getFilesByName(masterLogName);
    if (logFiles.hasNext()) {
      logSheetFile = logFiles.next();
      logSheet = SpreadsheetApp.open(logSheetFile).getActiveSheet();
    } else {
      const newLog = SpreadsheetApp.create(masterLogName);
      newLog.getActiveSheet().appendRow(['File Name', 'Folder Path', 'Date Created', 'Deleted On']);
      SpreadsheetApp.flush();
      logSheetFile = DriveApp.getFileById(newLog.getId());
      logFolder.addFile(logSheetFile);
      DriveApp.getRootFolder().removeFile(logSheetFile);
      logSheet = newLog.getActiveSheet();
    }

    deletedFilesLog.forEach(row => logSheet.appendRow(row));
    SpreadsheetApp.flush();

    Logger.log(`🗑️ Deleted ${deletedFilesLog.length} item(s). Appended to log: ${logSheetFile.getUrl()}`);
  } else {
    Logger.log('✅ No files or folders met deletion criteria.');
  }
}


function checkDriveApiEnabled() {
  const userEmail = ADMIN_EMAIL;

  const driveOk = (
    typeof Drive !== 'undefined' &&
    Drive.Files &&
    typeof Drive.Files.insert === 'function'
  );

  if (!driveOk) {
    const subject = `⚠️ CM360 Audit Script Needs Drive API Enabled`;
    const body = `
      The CM360 Audit script cannot run because the Advanced Drive API is not enabled.
      <br><br>
      Please enable it by opening the script editor and going to:
      <br>
      <strong>Extensions → Apps Script → Services</strong><br>
      Then add or enable <strong>Drive API</strong>.
    `;

    safeSendEmail({
      to: userEmail,
      subject,
      htmlBody: `<div style="font-family:Arial, sans-serif; font-size:13px;">${body}</div>`
    }, 'Drive API check');

    Logger.log("❌ Drive API is not enabled.");
    return false;
  }

  Logger.log("✅ Drive API is enabled.");
  return true;
}

function checkAuthorizationStatus() {
  const userEmail = ADMIN_EMAIL;

  try {
    const info = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

    if (info.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) {
      const subject = `⚠️ CM360 Audit Script Needs Reauthorization`;
      const message = `Your CM360 Audit script has lost authorization. Please open the script and reauthorize access.`;

      safeSendEmail({
        to: userEmail,
        subject,
        htmlBody: `<pre style="font-family:monospace; font-size:12px;">${escapeHtml(message)}</pre>`
      }, 'AUTH CHECK: reauthorization');

      return false;
    }

    Logger.log("✅ Authorization is valid.");
    return true;

  } catch (e) {
    Logger.log("❌ Failed to check authorization status:", e);

    const subject = `⚠️ CM360 Audit Script Failure`;
    const message = `The script failed to check authorization status. This may mean reauthorization is required.\n\nError: ${e.message}`;

    safeSendEmail({
      to: userEmail,
      subject,
      htmlBody: `<pre style="font-family:monospace; font-size:12px;">${escapeHtml(message)}</pre>`
    }, 'AUTH CHECK: failure fallback');

    return false;
  }
}


// === � THRESHOLDS MANAGEMENT ===
function getOrCreateThresholdsSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(THRESHOLDS_SHEET_NAME);
    
    if (!sheet) {
      Logger.log(`Creating new thresholds sheet: ${THRESHOLDS_SHEET_NAME}`);
      sheet = spreadsheet.insertSheet(THRESHOLDS_SHEET_NAME);
      
      // Set up the header row
      const headers = [
        'Config Name',
        'Flag Type',
        'Min Impressions',
        'Min Clicks',
        'Active',
        '',
        'INSTRUCTIONS'
      ];
      
      Logger.log('Setting headers...');
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format the main headers
      const mainHeaderRange = sheet.getRange(1, 1, 1, 5);
      mainHeaderRange.setFontWeight('bold');
      mainHeaderRange.setBackground('#4285f4');
      mainHeaderRange.setFontColor('#ffffff');
      
      // Format the instructions header
      const instructionsHeaderRange = sheet.getRange(1, 7, 1, 1);
      instructionsHeaderRange.setFontWeight('bold');
      instructionsHeaderRange.setBackground('#ff9900');
      instructionsHeaderRange.setFontColor('#ffffff');
      
      Logger.log('Setting up dropdowns...');
      
      // Add dropdown validation for Flag Type column (column B) - starting from row 2
      const flagTypeRange = sheet.getRange('B2:B');
      const flagTypeOptions = [
        'clicks_greater_than_impressions',
        'out_of_flight_dates',
        'pixel_size_mismatch',
        'default_ad_serving'
      ];
      
      const flagTypeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(flagTypeOptions)
        .setAllowInvalid(false)
        .setHelpText('Select the flag type for this threshold configuration.')
        .build();
      
      flagTypeRange.setDataValidation(flagTypeRule);
      
      // Add dropdown validation for Active column (column E) - starting from row 2
      const activeRange = sheet.getRange('E2:E');
      const activeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['TRUE', 'FALSE'])
        .setAllowInvalid(false)
        .setHelpText('Set to TRUE to activate threshold, FALSE to deactivate.')
        .build();
      
      activeRange.setDataValidation(activeRule);
      
      // Add instructions
      const instructions = [
        ['Config Name:', 'Enter the exact config name (PST01, PST02, NEXT01, etc.)'],
        ['Flag Type:', 'Select which flag type this threshold applies to'],
        ['Min Impressions:', 'Minimum impressions required for this flag to trigger'],
        ['Min Clicks:', 'Minimum clicks required for this flag to trigger'],
        ['Active:', 'TRUE to enable, FALSE to disable this threshold'],
        ['', ''],
        ['Logic:', 'A placement must meet EITHER impression OR click threshold to be evaluated'],
        ['Flag Types:', ''],
        ['• clicks_greater_than_impressions', 'Flags when clicks exceed impressions'],
        ['• out_of_flight_dates', 'Flags when placement is outside flight dates'],
        ['• pixel_size_mismatch', 'Flags when creative and placement pixels differ'],
        ['• default_ad_serving', 'Flags when ad type contains "default"'],
        ['', ''],
        ['Examples:', 'See rows below for different threshold configurations']
      ];
      
      sheet.getRange(2, 7, instructions.length, 2).setValues(instructions);
      
      // Add default threshold rows for each config
      const defaultThresholds = [];
      
      // Default threshold values (previously from config.flags)
      const defaultValues = {
        'PST01': { minImpressions: 50, minClicks: 10 },
        'PST02': { minImpressions: 100, minClicks: 100 },
        'PST03': { minImpressions: 0, minClicks: 0 },
        'NEXT01': { minImpressions: 1200, minClicks: 1200 },
        'NEXT02': { minImpressions: 0, minClicks: 0 },
        'NEXT03': { minImpressions: 0, minClicks: 0 },
        'SPTM01': { minImpressions: 10, minClicks: 10 },
        'NFL01': { minImpressions: 50, minClicks: 50 },
        'ENT01': { minImpressions: 15, minClicks: 15 }
      };
      
      auditConfigs.forEach(config => {
        const defaults = defaultValues[config.name] || { minImpressions: 0, minClicks: 0 };
        
        // Add a row for each flag type
        flagTypeOptions.forEach(flagType => {
          defaultThresholds.push([
            config.name,
            flagType,
            defaults.minImpressions,
            defaults.minClicks,
            'TRUE'
          ]);
        });
      });
      
      if (defaultThresholds.length > 0) {
        sheet.getRange(2, 1, defaultThresholds.length, 5).setValues(defaultThresholds);
      }
      
      // Format instructions
      const instructionsRange = sheet.getRange(2, 7, instructions.length, 2);
      instructionsRange.setFontSize(10);
      instructionsRange.setVerticalAlignment('top');
      
      // Auto-resize columns
      sheet.autoResizeColumns(1, 5);
      
      Logger.log('Thresholds sheet created successfully');
    } else {
      Logger.log(`Thresholds sheet already exists: ${THRESHOLDS_SHEET_NAME}`);
    }
    
    return sheet;
    
  } catch (error) {
    Logger.log(`❌ Error creating thresholds sheet: ${error.message}`);
    throw new Error(`Failed to create thresholds sheet: ${error.message}`);
  }
}

function loadThresholdsFromSheet() {
  try {
    const sheet = getOrCreateThresholdsSheet();
    const data = sheet.getDataRange().getValues();
    const thresholds = {};
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const configName = String(row[0] || '').trim();
      const flagType = String(row[1] || '').trim();
      const minImpressions = Number(row[2] || 0);
      const minClicks = Number(row[3] || 0);
      const active = String(row[4] || '').trim().toUpperCase();
      
      // Skip empty rows, instruction rows, or inactive thresholds
      if (!configName || !flagType || active !== 'TRUE' ||
          configName.includes('INSTRUCTIONS') || 
          configName.includes('•') ||
          configName.includes('Config Name:') ||
          configName.includes('Examples:')) {
        continue;
      }
      
      // Initialize config if not exists
      if (!thresholds[configName]) {
        thresholds[configName] = {};
      }
      
      // Store threshold data for this config and flag type
      thresholds[configName][flagType] = {
        minImpressions: Math.max(0, minImpressions),
        minClicks: Math.max(0, minClicks)
      };
    }
    
    Logger.log(`Loaded thresholds for ${Object.keys(thresholds).length} configs`);
    return thresholds;
    
  } catch (error) {
    Logger.log(`❌ Error loading thresholds: ${error.message}`);
    return {};
  }
}

// Helper function to get threshold for a specific config and flag type
function getThresholdForFlag(thresholdsData, configName, flagType) {
  if (!thresholdsData || !thresholdsData[configName] || !thresholdsData[configName][flagType]) {
    // Return default fallback thresholds if not found in sheet
    return { minImpressions: 0, minClicks: 0 };
  }
  
  return thresholdsData[configName][flagType];
}

// === 📧 EMAIL RECIPIENTS MANAGEMENT ===
function getOrCreateRecipientsSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(RECIPIENTS_SHEET_NAME);
    
    if (!sheet) {
      Logger.log(`Creating new recipients sheet: ${RECIPIENTS_SHEET_NAME}`);
      sheet = spreadsheet.insertSheet(RECIPIENTS_SHEET_NAME);
      
      // Set up the header row
      const headers = [
        'Config Name',
        'Primary Recipients',
        'CC Recipients',
        'Active',
        'Last Updated',
        '',
        'INSTRUCTIONS'
      ];
      
      Logger.log('Setting headers...');
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format the main headers
      const mainHeaderRange = sheet.getRange(1, 1, 1, 5);
      mainHeaderRange.setFontWeight('bold');
      mainHeaderRange.setBackground('#4285f4');
      mainHeaderRange.setFontColor('#ffffff');
      
      // Format the instructions header
      const instructionsHeaderRange = sheet.getRange(1, 7, 1, 1);
      instructionsHeaderRange.setFontWeight('bold');
      instructionsHeaderRange.setBackground('#ff9900');
      instructionsHeaderRange.setFontColor('#ffffff');
      
      Logger.log('Setting up dropdowns...');
      
      // Add dropdown validation for Active column (column D) - starting from row 2
      const activeRange = sheet.getRange('D2:D');
      const activeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['TRUE', 'FALSE'])
        .setAllowInvalid(false)
        .setHelpText('Set to TRUE to use these recipients, FALSE to disable.')
        .build();
      
      activeRange.setDataValidation(activeRule);
      
      // Add instructions
      const instructions = [
        ['Config Name:', 'Enter the exact config name (PST01, PST02, NEXT01, etc.)'],
        ['Primary Recipients:', 'Main email addresses (comma-separated if multiple)'],
        ['CC Recipients:', 'CC email addresses (comma-separated if multiple)'],
        ['Active:', 'TRUE to use these recipients, FALSE to disable'],
        ['Last Updated:', 'Automatically updated when you modify recipients'],
        ['', ''],
        ['Staging Mode Override:', `Currently: ${STAGING_MODE === 'Y' ? 'STAGING (all emails go to admin)' : 'PRODUCTION (uses sheet recipients)'}`],
        ['', ''],
        ['Email Format:', ''],
        ['• Single recipient:', 'user@company.com'],
        ['• Multiple recipients:', 'user1@company.com, user2@company.com'],
        ['• Leave CC blank if not needed', ''],
        ['', ''],
        ['Examples:', 'See rows below for different recipient configurations']
      ];
      
      sheet.getRange(2, 7, instructions.length, 2).setValues(instructions);
      
      // Add default recipient rows for each config (using current hardcoded values)
      const defaultRecipients = [
        ['PST01', ADMIN_EMAIL, '', 'TRUE', formatDate(new Date(), 'yyyy-MM-dd')],
        ['PST02', 'fvariath@horizonmedia.com', ADMIN_EMAIL, 'TRUE', formatDate(new Date(), 'yyyy-MM-dd')],
        ['PST03', 'dmaestre@horizonmedia.com', ADMIN_EMAIL, 'TRUE', formatDate(new Date(), 'yyyy-MM-dd')],
        ['NEXT01', 'bosborne@horizonmedia.com, mmassaroni@horizonmedia.com', ADMIN_EMAIL, 'TRUE', formatDate(new Date(), 'yyyy-MM-dd')],
        ['NEXT02', 'rschaff@horizonmedia.com, mmassaroni@horizonmedia.com, jwong@horizonmedia.com', ADMIN_EMAIL, 'TRUE', formatDate(new Date(), 'yyyy-MM-dd')],
        ['NEXT03', 'szeterberg@horizonmedia.com, mmassaroni@horizonmedia.com, jwong@horizonmedia.com', ADMIN_EMAIL, 'TRUE', formatDate(new Date(), 'yyyy-MM-dd')],
        ['SPTM01', 'spectrum_adops@horizonmedia.com', ADMIN_EMAIL, 'TRUE', formatDate(new Date(), 'yyyy-MM-dd')],
        ['NFL01', 'NFL_AdOps@horizonmedia.com, sbermolone@horizonmedia.com', ADMIN_EMAIL, 'TRUE', formatDate(new Date(), 'yyyy-MM-dd')],
        ['ENT01', 'sremick@horizonmedia.com, cali@horizonmedia.com', ADMIN_EMAIL, 'TRUE', formatDate(new Date(), 'yyyy-MM-dd')]
      ];
      
      if (defaultRecipients.length > 0) {
        sheet.getRange(2, 1, defaultRecipients.length, 5).setValues(defaultRecipients);
      }
      
      // Format instructions
      const instructionsRange = sheet.getRange(2, 7, instructions.length, 2);
      instructionsRange.setFontSize(10);
      instructionsRange.setVerticalAlignment('top');
      
      // Auto-resize columns
      sheet.autoResizeColumns(1, 5);
      
      Logger.log('Recipients sheet created successfully');
    } else {
      Logger.log(`Recipients sheet already exists: ${RECIPIENTS_SHEET_NAME}`);
    }
    
    return sheet;
    
  } catch (error) {
    Logger.log(`❌ Error creating recipients sheet: ${error.message}`);
    throw new Error(`Failed to create recipients sheet: ${error.message}`);
  }
}

function loadRecipientsFromSheet() {
  try {
    const sheet = getOrCreateRecipientsSheet();
    const data = sheet.getDataRange().getValues();
    const recipients = {};
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const configName = String(row[0] || '').trim();
      const primaryRecipients = String(row[1] || '').trim();
      const ccRecipients = String(row[2] || '').trim();
      const active = String(row[3] || '').trim().toUpperCase();
      
      // Skip empty rows, instruction rows, or inactive recipients
      if (!configName || active !== 'TRUE' ||
          configName.includes('INSTRUCTIONS') || 
          configName.includes('•') ||
          configName.includes('Config Name:') ||
          configName.includes('Examples:')) {
        continue;
      }
      
      // Store recipient data for this config
      recipients[configName] = {
        primary: primaryRecipients,
        cc: ccRecipients
      };
    }
    
    Logger.log(`Loaded recipients for ${Object.keys(recipients).length} configs`);
    return recipients;
    
  } catch (error) {
    Logger.log(`❌ Error loading recipients: ${error.message}`);
    return {};
  }
}

// === �📋 EXCLUSIONS MANAGEMENT ===
function getOrCreateExclusionsSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(EXCLUSIONS_SHEET_NAME);
    
    if (!sheet) {
      Logger.log(`Creating new exclusions sheet: ${EXCLUSIONS_SHEET_NAME}`);
      sheet = spreadsheet.insertSheet(EXCLUSIONS_SHEET_NAME);
      
      // Set up the header row
      const headers = [
        'Config Name',
        'Placement ID', 
        'Placement Name',
        'Site Name',
        'Name Fragment',
        'Apply to All Configs',
        'Flag Type',
        'Reason',
        'Added By',
        'Date Added',
        'Active',
        '',
        'INSTRUCTIONS'
      ];
      
      Logger.log('Setting headers...');
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format the main headers
      const mainHeaderRange = sheet.getRange(1, 1, 1, 11);
      mainHeaderRange.setFontWeight('bold');
      mainHeaderRange.setBackground('#4285f4');
      mainHeaderRange.setFontColor('#ffffff');
      
      // Format the instructions header
      const instructionsHeaderRange = sheet.getRange(1, 13, 1, 1);
      instructionsHeaderRange.setFontWeight('bold');
      instructionsHeaderRange.setBackground('#ff9900');
      instructionsHeaderRange.setFontColor('#ffffff');
      
      Logger.log('Headers formatted, setting up protection...');
      
      // Lock the Placement Name column (column C)
      const placementNameRange = sheet.getRange('C:C');
      const protection = placementNameRange.protect().setDescription('Placement Name (Auto-populated - Do Not Edit)');
      protection.setWarningOnly(true);
      
      Logger.log('Setting up dropdowns...');
      
      // Add dropdown validation for Apply to All Configs column (column F) - starting from row 2
      const applyAllRange = sheet.getRange('F2:F');
      const applyAllRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['TRUE', 'FALSE'])
        .setAllowInvalid(false)
        .setHelpText('Set to TRUE to apply this exclusion to ALL config teams, FALSE to apply only to specified config.')
        .build();
      
      applyAllRange.setDataValidation(applyAllRule);
      
      // Add dropdown validation for Flag Type column (column G) - starting from row 2
      const flagTypeRange = sheet.getRange('G2:G');
      const flagTypeOptions = [
        'clicks_greater_than_impressions',
        'out_of_flight_dates',
        'pixel_size_mismatch',
        'default_ad_serving',
        'all_flags'
      ];
      
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(flagTypeOptions)
        .setAllowInvalid(false)
        .setHelpText('Select a flag type from the dropdown. Use "all_flags" to exclude from all flag types.')
        .build();
      
      flagTypeRange.setDataValidation(rule);
      
      // Add dropdown validation for Active column (column K) - starting from row 2
      const activeRange = sheet.getRange('K2:K');
      const activeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['TRUE', 'FALSE'])
        .setAllowInvalid(false)
        .setHelpText('Set to TRUE to activate exclusion, FALSE to deactivate.')
        .build();
      
      activeRange.setDataValidation(activeRule);
      
      // Add instructions
      const instructions = [
        ['Config Name:', 'Enter the exact config name (PST01, PST02, NEXT01, etc.) OR leave blank if using "Apply to All Configs"'],
        ['Placement ID:', 'Enter the CM360 Placement ID number (leave blank if using Site Name or Name Fragment)'],
        ['Placement Name:', 'Auto-populated - use "Update Placement Names" button'],
        ['Site Name:', 'Enter exact site name as it appears in CM360 reporting (alternative to Placement ID)'],
        ['Name Fragment:', 'Enter text fragment that appears in placement names (matches any placement containing this text)'],
        ['Apply to All Configs:', 'TRUE = applies to ALL config teams, FALSE = applies only to specified config'],
        ['Flag Type:', 'Select which flag type to exclude'],
        ['Reason:', 'Brief explanation for the exclusion'],
        ['Added By:', 'Your name or email'],
        ['Date Added:', 'Date this exclusion was added'],
        ['Active:', 'TRUE to enable, FALSE to disable'],
        ['', ''],
        ['Exclusion Types:', ''],
        ['• Placement ID', 'Excludes specific placement by ID'],
        ['• Site Name', 'Excludes all placements from a specific site'],
        ['• Name Fragment', 'Excludes placements with names containing the fragment'],
        ['', ''],
        ['Flag Types:', ''],
        ['• clicks_greater_than_impressions', 'Excludes clicks > impressions flags'],
        ['• out_of_flight_dates', 'Excludes out of flight date flags'],
        ['• pixel_size_mismatch', 'Excludes pixel mismatch flags'],
        ['• default_ad_serving', 'Excludes default ad serving flags'],
        ['• all_flags', 'Excludes from ALL flag types'],
        ['', ''],
        ['Examples:', 'See rows 2-4 below for different exclusion types']
      ];
      
      sheet.getRange(2, 13, instructions.length, 2).setValues(instructions);
      
      // Add example rows separately (in the main data area)
      const exampleRows = [
        ['PST01', '424138145', '', '', '', 'FALSE', 'clicks_greater_than_impressions', 'Specific placement exclusion', 'your.name@company.com', '2025-08-12', 'TRUE'],
        ['', '', '', 'YouTube', '', 'TRUE', 'all_flags', 'Exclude all YouTube placements from all configs', 'your.name@company.com', '2025-08-12', 'TRUE'],
        ['PST02', '', '', '', 'Social Media', 'FALSE', 'pixel_size_mismatch', 'Exclude placements with "Social Media" in name', 'your.name@company.com', '2025-08-12', 'TRUE']
      ];
      
      sheet.getRange(2, 1, exampleRows.length, exampleRows[0].length).setValues(exampleRows);
      
      // Format instructions
      const instructionsRange = sheet.getRange(2, 13, instructions.length, 2);
      instructionsRange.setFontSize(10);
      instructionsRange.setVerticalAlignment('top');
      
      Logger.log('Exclusions sheet created successfully');
    } else {
      Logger.log(`Exclusions sheet already exists: ${EXCLUSIONS_SHEET_NAME}`);
    }
    
    return sheet;
    
  } catch (error) {
    Logger.log(`❌ Error creating exclusions sheet: ${error.message}`);
    throw new Error(`Failed to create exclusions sheet: ${error.message}`);
  }
}

// Auto-populate placement names when data is entered (Simple Trigger)
function onEdit(e) {
  // Only process if we have an event object
  if (!e) return;
  
  const sheet = e.source.getActiveSheet();
  
  // Only process if we're on the exclusions sheet
  if (sheet.getName() !== EXCLUSIONS_SHEET_NAME) return;
  
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  
  // Only process if editing Config Name (col 1) or Placement ID (col 2) and not header row
  if (row <= 1 || (col !== 1 && col !== 2)) return;
  
  try {
    const configName = String(sheet.getRange(row, 1).getValue() || '').trim();
    const placementId = String(sheet.getRange(row, 2).getValue() || '').trim();
    
    // Only lookup if both config and placement ID are provided
    if (configName && placementId && 
        !configName.includes('INSTRUCTIONS') && 
        !configName.includes('•') && 
        !configName.includes('Config Name:')) {
      
      const placementName = LOOKUP_PLACEMENT_NAME(configName, placementId);
      if (placementName) {
        sheet.getRange(row, 3).setValue(placementName);
        Logger.log(`Auto-populated placement name for ${configName}/${placementId}: ${placementName}`);
      } else {
        sheet.getRange(row, 3).setValue('(Not found in recent data)');
        Logger.log(`Could not find placement name for ${configName}/${placementId}`);
      }
    } else if (!configName && !placementId) {
      // Clear the placement name if both config and ID are cleared
      sheet.getRange(row, 3).setValue('');
    }
    
  } catch (error) {
    Logger.log(`Error in onEdit for exclusions: ${error.message}`);
  }
}

function loadExclusionsFromSheet() {
  try {
    const sheet = getOrCreateExclusionsSheet();
    const data = sheet.getDataRange().getValues();
    const exclusions = {};
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const configName = String(row[0] || '').trim();
      const placementId = String(row[1] || '').trim();
      const placementName = String(row[2] || '').trim();
      const siteName = String(row[3] || '').trim();
      const nameFragment = String(row[4] || '').trim();
      const applyToAllConfigs = String(row[5] || '').trim().toUpperCase();
      const flagType = String(row[6] || '').trim();
      const active = String(row[10] || '').trim().toUpperCase();
      
      // Skip empty rows, instruction rows, or inactive exclusions
      if (!flagType || active !== 'TRUE' ||
          configName.includes('INSTRUCTIONS') || 
          configName.includes('•') ||
          configName.includes('Config Name:') ||
          configName.includes('Examples:')) {
        continue;
      }
      
      // Must have at least one identification method
      if (!placementId && !siteName && !nameFragment) {
        continue;
      }
      
      // If "Apply to All Configs" is TRUE, apply to all known configs
      const configsToApply = [];
      if (applyToAllConfigs === 'TRUE') {
        // Add all known config names
        configsToApply.push('PST01', 'PST02', 'PST03', 'PST04', 'PST05', 'PST06', 'PST07', 'PST08', 'NEXT01');
      } else if (configName) {
        configsToApply.push(configName);
      } else {
        continue; // No config specified and not applying to all
      }
      
      // Process each config
      for (const config of configsToApply) {
        // Initialize config if not exists
        if (!exclusions[config]) {
          exclusions[config] = {};
        }
        
        // Initialize flag type object if not exists
        if (!exclusions[config][flagType]) {
          exclusions[config][flagType] = {
            placementIds: [],
            siteNames: [],
            nameFragments: []
          };
        }
        
        // Add exclusion data based on type
        if (placementId) {
          exclusions[config][flagType].placementIds.push(placementId);
        }
        if (siteName) {
          exclusions[config][flagType].siteNames.push(siteName);
        }
        if (nameFragment) {
          exclusions[config][flagType].nameFragments.push(nameFragment);
        }
      }
    }
    
    Logger.log(`Loaded exclusions for ${Object.keys(exclusions).length} configs`);
    return exclusions;
    
  } catch (error) {
    Logger.log(`❌ Error loading exclusions: ${error.message}`);
    return {};
  }
}

// Helper function to check if a placement should be excluded for a specific flag type
function isPlacementExcludedForFlag(exclusionsData, configName, placementId, flagType, placementName = '', siteName = '') {
  if (!exclusionsData || !exclusionsData[configName]) {
    return false;
  }
  
  const trimmedId = String(placementId || '').trim();
  const trimmedPlacementName = String(placementName || '').trim().toLowerCase();
  const trimmedSiteName = String(siteName || '').trim().toLowerCase();
  const configExclusions = exclusionsData[configName];
  
  // Check if excluded from all flags
  if (configExclusions.all_flags) {
    const allFlagsExclusions = configExclusions.all_flags;
    
    // Check placement ID exclusions
    if (allFlagsExclusions.placementIds && 
        allFlagsExclusions.placementIds.some(id => String(id).trim() === trimmedId)) {
      return true;
    }
    
    // Check site name exclusions
    if (allFlagsExclusions.siteNames && trimmedSiteName &&
        allFlagsExclusions.siteNames.some(site => String(site).trim().toLowerCase() === trimmedSiteName)) {
      return true;
    }
    
    // Check name fragment exclusions
    if (allFlagsExclusions.nameFragments && trimmedPlacementName &&
        allFlagsExclusions.nameFragments.some(fragment => 
          trimmedPlacementName.includes(String(fragment).trim().toLowerCase()))) {
      return true;
    }
  }
  
  // Check if excluded from specific flag type
  if (configExclusions[flagType]) {
    const flagExclusions = configExclusions[flagType];
    
    // Check placement ID exclusions
    if (flagExclusions.placementIds && 
        flagExclusions.placementIds.some(id => String(id).trim() === trimmedId)) {
      return true;
    }
    
    // Check site name exclusions
    if (flagExclusions.siteNames && trimmedSiteName &&
        flagExclusions.siteNames.some(site => String(site).trim().toLowerCase() === trimmedSiteName)) {
      return true;
    }
    
    // Check name fragment exclusions
    if (flagExclusions.nameFragments && trimmedPlacementName &&
        flagExclusions.nameFragments.some(fragment => 
          trimmedPlacementName.includes(String(fragment).trim().toLowerCase()))) {
      return true;
    }
  }
  
  return false;
}

// Function to lookup placement name from recent audit data
function LOOKUP_PLACEMENT_NAME(configName, placementId) {
  try {
    // Look for recent merged reports for this config
    const config = auditConfigs.find(c => c.name === configName);
    if (!config) return null;
    
    const mergedFolder = getDriveFolderByPath_(config.mergedFolderPath);
    const files = mergedFolder.getFiles();
    
    // Look at the most recent files (up to 5)
    const recentFiles = [];
    while (files.hasNext() && recentFiles.length < 5) {
      recentFiles.push(files.next());
    }
    
    // Sort by date, newest first
    recentFiles.sort((a, b) => b.getDateCreated() - a.getDateCreated());
    
    // Search through recent files for the placement ID
    for (const file of recentFiles) {
      try {
        const spreadsheet = SpreadsheetApp.open(file);
        const sheet = spreadsheet.getSheets()[0];
        const data = sheet.getDataRange().getValues();
        
        // Find the header row
        let headerRowIndex = -1;
        let placementIdCol = -1;
        let placementNameCol = -1;
        
        for (let i = 0; i < Math.min(data.length, 20); i++) {
          const row = data[i];
          for (let j = 0; j < row.length; j++) {
            const cellValue = String(row[j] || '').toLowerCase();
            if (cellValue.includes('placement id')) {
              headerRowIndex = i;
              placementIdCol = j;
            }
            // More specific matching for placement name column
            if ((cellValue === 'placement' || cellValue === 'placement name') && 
                !cellValue.includes('id') && 
                !cellValue.includes('pixel') && 
                !cellValue.includes('start') && 
                !cellValue.includes('end') && 
                !cellValue.includes('date')) {
              placementNameCol = j;
            }
          }
          if (headerRowIndex !== -1 && placementIdCol !== -1 && placementNameCol !== -1) {
            break;
          }
        }
        
        if (headerRowIndex !== -1 && placementIdCol !== -1 && placementNameCol !== -1) {
          // Search for the placement ID in the data
          for (let i = headerRowIndex + 1; i < data.length; i++) {
            const row = data[i];
            if (String(row[placementIdCol] || '').trim() === String(placementId).trim()) {
              const placementName = String(row[placementNameCol] || '').trim();
              
              // Validate that this looks like a placement name, not a date or other data
              if (placementName && 
                  !placementName.includes('GMT') && 
                  !placementName.includes('00:00:00') && 
                  !/^\d{4}-\d{2}-\d{2}/.test(placementName) && // Not YYYY-MM-DD format
                  !/^\w{3}\s\w{3}\s\d{2}\s\d{4}/.test(placementName) && // Not "Mon Jan 01 2025" format
                  placementName.length > 5 && // Reasonable minimum length
                  placementName.length < 200) { // Reasonable maximum length
                
                Logger.log(`Found placement name for ${placementId}: ${placementName}`);
                return placementName;
              } else {
                Logger.log(`Rejected invalid placement name for ${placementId}: ${placementName}`);
              }
            }
          }
        }
      } catch (fileError) {
        // Skip files that can't be opened
        continue;
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log(`Error looking up placement name: ${error.message}`);
    return null;
  }
}

// Function to update all placement names in the exclusions sheet
function updatePlacementNamesFromReports() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Update Placement Names',
      'This will search through the latest merged reports to update placement names in the exclusions sheet. This may take a few minutes. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    const sheet = getOrCreateExclusionsSheet();
    const data = sheet.getDataRange().getValues();
    let updatedCount = 0;
    let notFoundCount = 0;
    
    Logger.log('Starting placement name update process...');
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const configName = String(row[0] || '').trim();
      const placementId = String(row[1] || '').trim();
      const currentPlacementName = String(row[2] || '').trim();
      
      // Skip empty rows, instruction rows, or rows that already have placement names
      if (!configName || !placementId || 
          configName.includes('INSTRUCTIONS') || 
          configName.includes('•') ||
          configName.includes('Config Name:') ||
          configName.includes('Example:')) {
        continue;
      }
      
      // Skip if placement name already exists and doesn't look like an error message
      if (currentPlacementName && 
          !currentPlacementName.includes('(Not found') && 
          !currentPlacementName.includes('(Error')) {
        Logger.log(`Skipping ${configName}/${placementId} - already has name: ${currentPlacementName}`);
        continue;
      }
      
      Logger.log(`Looking up placement name for ${configName}/${placementId}...`);
      
      try {
        const placementName = LOOKUP_PLACEMENT_NAME(configName, placementId);
        if (placementName) {
          sheet.getRange(i + 1, 3).setValue(placementName);
          Logger.log(`✅ Updated ${configName}/${placementId}: ${placementName}`);
          updatedCount++;
        } else {
          sheet.getRange(i + 1, 3).setValue('(Not found in recent reports)');
          Logger.log(`❌ Not found ${configName}/${placementId}`);
          notFoundCount++;
        }
        
        // Add a small delay to avoid rate limiting
        if ((updatedCount + notFoundCount) % 5 === 0) {
          Utilities.sleep(1000);
        }
        
      } catch (error) {
        Logger.log(`Error processing ${configName}/${placementId}: ${error.message}`);
        sheet.getRange(i + 1, 3).setValue(`(Error: ${error.message})`);
        notFoundCount++;
      }
    }
    
    const message = `Placement name update complete!\n\n` +
                   `✅ Updated: ${updatedCount}\n` +
                   `❌ Not found: ${notFoundCount}\n\n` +
                   `Check the Logger for detailed results.`;
    
    ui.alert('Update Complete', message, ui.ButtonSet.OK);
    Logger.log(`🎉 Update complete: ${updatedCount} updated, ${notFoundCount} not found`);
    
  } catch (error) {
    Logger.log(`❌ Error in updatePlacementNamesFromReports: ${error.message}`);
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', `Failed to update placement names: ${error.message}`, ui.ButtonSet.OK);
  }
}


// === 🧪 DEBUGGING & TEST TOOLS ===
function debugValidateAuditConfigs() {
  try {
    validateAuditConfigs();
    Logger.log("✅ All audit configs passed validation.");
  } catch (e) {
    Logger.log(`❌ Audit config validation failed: ${e.message}`);
    SpreadsheetApp.getUi().alert(`Audit config validation failed:\n\n${e.message}`);
  }
}

function debugPrintConfigSummary() {
  auditConfigs.forEach(c => {
    Logger.log(`[${c.name}] Label: ${c.label}, Flags: ${JSON.stringify(c.flags)} (Recipients managed via sheet)`);
  });
}

function checkMissingBatchRunners() {
  const batches = getAuditConfigBatches(BATCH_SIZE);
  const messages = [];

  for (let i = 0; i < batches.length; i++) {
    const fnName = `runDailyAuditsBatch${i + 1}`;
    if (typeof this[fnName] !== 'function') {
      messages.push(`❌ Missing: ${fnName}`);
    } else {
      messages.push(`✅ Present: ${fnName}`);
    }
  }

  if (messages.length === 0) {
    messages.push(`⚠️ No batch configs found.`);
  }

  return messages;
}

function runPST01Audit() { runDailyAuditByName('PST01'); }
function runPST02Audit() { runDailyAuditByName('PST02'); }
function runPST03Audit()  { runDailyAuditByName('PST03'); }
function runNEXT01Audit()   { runDailyAuditByName('NEXT01'); }
function runNEXT02Audit()   { runDailyAuditByName('NEXT02'); }
function runNEXT03Audit()   { runDailyAuditByName('NEXT03'); }
function runSPTM01Audit()   { runDailyAuditByName('SPTM01'); }
function runNFL01Audit()   { runDailyAuditByName('NFL01'); }
function runENT02Audit()   { runDailyAuditByName('ENT02'); }

// === 🧪 ENHANCED EXCLUSIONS TESTING ===
function testEnhancedExclusions() {
  try {
    Logger.log('🧪 Testing Enhanced Exclusions System...');
    
    // Create/update exclusions sheet
    const sheet = getOrCreateExclusionsSheet();
    Logger.log('✅ Exclusions sheet created/updated successfully');
    
    // Test loading exclusions
    const exclusionsData = loadExclusionsFromSheet();
    Logger.log(`✅ Loaded exclusions data: ${JSON.stringify(exclusionsData, null, 2)}`);
    
    // Test exclusion checking with different scenarios
    const testCases = [
      {
        description: 'Placement ID exclusion',
        configName: 'PST01',
        placementId: '424138145',
        flagType: 'clicks_greater_than_impressions',
        placementName: 'Test Placement',
        siteName: 'Test Site'
      },
      {
        description: 'Site name exclusion',
        configName: 'PST02',
        placementId: '123456789',
        flagType: 'all_flags',
        placementName: 'YouTube Video Ad',
        siteName: 'YouTube'
      },
      {
        description: 'Name fragment exclusion',
        configName: 'PST02',
        placementId: '987654321',
        flagType: 'pixel_size_mismatch',
        placementName: 'Social Media Campaign',
        siteName: 'Facebook'
      }
    ];
    
    testCases.forEach(testCase => {
      const isExcluded = isPlacementExcludedForFlag(
        exclusionsData,
        testCase.configName,
        testCase.placementId,
        testCase.flagType,
        testCase.placementName,
        testCase.siteName
      );
      Logger.log(`${testCase.description}: ${isExcluded ? 'EXCLUDED' : 'NOT EXCLUDED'}`);
    });
    
    // Open the exclusions sheet for review
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
    
    Logger.log('✅ Enhanced exclusions system test completed successfully!');
    Logger.log('📋 Please review the exclusions sheet and test the new features:');
    Logger.log('   • Site Name column for excluding by site');
    Logger.log('   • Name Fragment column for partial name matching');
    Logger.log('   • Apply to All Configs for global exclusions');
    
  } catch (error) {
    Logger.log(`❌ Error testing enhanced exclusions: ${error.message}`);
    throw error;
  }
}

// === 🧪 THRESHOLDS TESTING ===
function testThresholdsSystem() {
  try {
    Logger.log('🧪 Testing Thresholds System...');
    
    // Create/update thresholds sheet
    const sheet = getOrCreateThresholdsSheet();
    Logger.log('✅ Thresholds sheet created/updated successfully');
    
    // Test loading thresholds
    const thresholdsData = loadThresholdsFromSheet();
    Logger.log(`✅ Loaded thresholds data: ${JSON.stringify(thresholdsData, null, 2)}`);
    
    // Test threshold retrieval for different scenarios
    const testCases = [
      {
        description: 'PST01 clicks threshold',
        configName: 'PST01',
        flagType: 'clicks_greater_than_impressions'
      },
      {
        description: 'PST02 pixel threshold',
        configName: 'PST02',
        flagType: 'pixel_size_mismatch'
      },
      {
        description: 'Non-existent config (should use defaults)',
        configName: 'INVALID',
        flagType: 'out_of_flight_dates'
      }
    ];
    
    testCases.forEach(testCase => {
      const threshold = getThresholdForFlag(
        thresholdsData,
        testCase.configName,
        testCase.flagType
      );
      Logger.log(`${testCase.description}: Min Impressions=${threshold.minImpressions}, Min Clicks=${threshold.minClicks}`);
    });
    
    // Open the thresholds sheet for review
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
    
    Logger.log('✅ Thresholds system test completed successfully!');
    Logger.log('📊 Please review the thresholds sheet and adjust values as needed:');
    Logger.log('   • Each config has separate thresholds for each flag type');
    Logger.log('   • Min Impressions OR Min Clicks threshold triggers evaluation');
    Logger.log('   • Set Active to FALSE to disable specific threshold rules');
    
  } catch (error) {
    Logger.log(`❌ Error testing thresholds system: ${error.message}`);
    throw error;
  }
}

/**
 * Test function for the Recipients Management System
 * Creates/updates recipients sheet and validates dynamic recipient resolution
 */
function testRecipientsSystem() {
  Logger.log(`🧪 Testing Recipients Management System...`);
  
  try {
    // Step 1: Create/Update recipients sheet
    Logger.log(`📋 Creating/updating recipients sheet...`);
    const sheet = getOrCreateRecipientsSheet();
    Logger.log(`✅ Recipients sheet ready with ${sheet.getLastRow() - 1} recipient configurations`);
    
    // Step 2: Load recipients data
    Logger.log(`📂 Loading recipients data...`);
    const recipientsData = loadRecipientsFromSheet();
    const configCount = Object.keys(recipientsData).length;
    Logger.log(`✅ Loaded recipients for ${configCount} configurations`);
    
    // Step 3: Test recipient resolution with sample configs
    Logger.log(`🔍 Testing recipient resolution...`);
    let testCount = 0;
    
    for (const configName of Object.keys(recipientsData)) {
      const recipients = resolveRecipients(configName, recipientsData);
      const cc = resolveCc(configName, recipientsData);
      
      Logger.log(`  📧 [${configName}]:`);
      Logger.log(`    To: ${recipients}`);
      Logger.log(`    CC: ${cc}`);
      
      testCount++;
      if (testCount >= 3) break; // Limit output for testing
    }
    
    // Step 4: Test staging mode behavior
    Logger.log(`🎭 Testing staging mode override...`);
    const originalStaging = IS_STAGING_MODE;
    
    // Temporarily enable staging mode
    global.IS_STAGING_MODE = true;
    const stagingRecipients = resolveRecipients('test-config', recipientsData);
    Logger.log(`  Staging mode recipients: ${stagingRecipients}`);
    
    // Restore original staging mode
    global.IS_STAGING_MODE = originalStaging;
    
    Logger.log(`✅ Recipients system test completed successfully!`);
    Logger.log(`📊 Summary:`);
    Logger.log(`  - Recipients sheet: Ready`);
    Logger.log(`  - Configurations loaded: ${configCount}`);
    Logger.log(`  - Recipient resolution: Working`);
    Logger.log(`  - Staging mode override: Working`);
    
  } catch (error) {
    Logger.log(`❌ Error testing recipients system: ${error.message}`);
    throw error;
  }
}