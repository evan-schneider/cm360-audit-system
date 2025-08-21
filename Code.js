// === CONFIGURATION & CONSTANTS ===
const ADMIN_EMAIL = 'evschneider@horizonmedia.com';
const STAGING_MODE = 'Y'; // Set to 'Y' for staging mode, 'N' for product

//
// Configuration data source: use external sheet if EXTERNAL_CONFIG_SHEET_ID is set
const EXTERNAL_CONFIG_SHEET_ID = '1-566gqkyZRNDeNtXWUjKDB_H8A9XbhCu8zL-uaZdGT8'; // External config sheet ID

const EXCLUSIONS_SHEET_NAME = 'Audit Exclusions'; // Name of the sheet containing exclusions
const THRESHOLDS_SHEET_NAME = 'Audit Thresholds'; // Name of the sheet containing flag thresholds
const RECIPIENTS_SHEET_NAME = 'Audit Recipients'; // Name of the sheet containing email recipients

const BATCH_SIZE = 2;
const TRASH_ROOT_PATH = ['Project Log Files', 'CM360 Daily Audits', 'To Trash After 60 Days'];
const DELETION_LOG_PATH = [...TRASH_ROOT_PATH, 'Deletion Log'];
const ADMIN_LOG_NAME = 'CM360 Deleted Files Log';

// Ensure folderPath exists before globals use it
var folderPath = (typeof folderPath === 'function') ? folderPath : function(type, configName) {
	var root = Array.isArray(TRASH_ROOT_PATH) ? TRASH_ROOT_PATH : [];
	return root.concat([type, configName]);
};

// === AUDIT CONFIGS (source of truth for paths/labels) ===
// Each config carries:
// - name: short code for the team/bundle
// - label: Gmail label to pull daily attachments from (adjust to match your filters)
// - mergedFolderPath: Drive path array where merged reports are saved
// - tempDailyFolderPath: Drive path array where daily inbound files are staged
// Paths default under TRASH_ROOT_PATH to align with cleanup logic.
function makeAuditConfig_(name, label) {
	return {
		name: name,
		label: label || name,
		mergedFolderPath: [...TRASH_ROOT_PATH, 'Merged Reports', name],
		tempDailyFolderPath: [...TRASH_ROOT_PATH, 'Temp Daily Reports', name]
	};
}

// NOTE: Labels reflect your existing nested Gmail labels like "Daily Audits/CM360/<CONFIG>".
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

// === UTILITY HELPERS ===
function getConfigSpreadsheet() {
	// Return external config spreadsheet if configured; otherwise active spreadsheet
	return EXTERNAL_CONFIG_SHEET_ID
		? SpreadsheetApp.openById(EXTERNAL_CONFIG_SHEET_ID)
		: SpreadsheetApp.getActiveSpreadsheet();
}

function folderPath(type, configName) {
 return [...TRASH_ROOT_PATH, type, configName];
}

function resolveRecipients(configName, recipientsData = null) {
	if (STAGING_MODE === 'Y') return ADMIN_EMAIL; // staging -> admin only
	var name = String(configName || '');
	if (recipientsData && recipientsData[name]) return recipientsData[name].primary || ADMIN_EMAIL;
	return ADMIN_EMAIL;
}

function resolveCc(configName, recipientsData = null) {
	if (STAGING_MODE === 'Y') return ''; // staging -> no CC
	var name = String(configName || '');
	return recipientsData && recipientsData[name] ? (recipientsData[name].cc || '') : '';
}

// Resolve a Drive folder by path array, creating missing levels when possible.
function getDriveFolderByPath_(pathArray) {
	try {
		if (!Array.isArray(pathArray) || pathArray.length === 0) return null;
		let folder = DriveApp.getRootFolder();
		for (const part of pathArray) {
			const name = String(part || '').trim();
			if (!name) continue;
			const it = folder.getFoldersByName(name);
			folder = it.hasNext() ? it.next() : folder.createFolder(name);
		}
		return folder;
	} catch (e) {
		Logger.log('getDriveFolderByPath_ error: ' + e.message);
		return null;
	}
}

// Small helper wrapper for consistent date formatting across the script
function formatDate(dateObj, pattern = 'yyyy-MM-dd', timeZone = Session.getScriptTimeZone()) {
	try {
		return Utilities.formatDate(dateObj instanceof Date ? dateObj : new Date(dateObj), timeZone, pattern);
	} catch (e) {
		Logger.log('formatDate error: ' + e.message);
		return String(dateObj);
	}
}

// Cache key for storing combined audit results
function getAuditCacheKey_() {
	return 'CM360_AUDIT_COMBINED_RESULTS_V1';
}

// Robust Gmail label resolver: tries exact match, normalized match, and suffix match by config name
function findGmailLabel_(desiredLabel, configName) {
	// 1) Fast path: exact
	var lbl = GmailApp.getUserLabelByName(desiredLabel);
	if (lbl) return lbl;

	// Build a canonical form: trim, collapse spaces, unify slashes
	function canon(s) {
		return String(s || '')
			.replace(/\s*\/\s*/g, '/')
			.replace(/\s+/g, ' ')
			.trim()
			.toLowerCase();
	}

	var desiredCanon = canon(desiredLabel);
	var labels = GmailApp.getUserLabels();

	// 2) Case/whitespace-insensitive match
	for (var i = 0; i < labels.length; i++) {
		var name = labels[i].getName();
		if (canon(name) === desiredCanon) return labels[i];
	}

	// 3) Suffix match by config name to catch slight path mismatches
	var suffix = String(configName || '').toLowerCase();
	if (suffix) {
		for (var j = 0; j < labels.length; j++) {
			var nm = labels[j].getName();
			var nmCanon = canon(nm);
			if (nmCanon.endsWith('/' + suffix) || nmCanon === suffix) {
				Logger.log(`[LABEL] Using closest matching label for ${configName}: ${nm}`);
				return labels[j];
			}
		}
	}

	return null;
}

// Safe email sender wrapper - centralizes staging behavior, logging and error handling
function safeSendEmail(opts = {}, context = '') {
	try {
		const to = opts.to || ADMIN_EMAIL;
		const subject = opts.subject || '';
		const plainBody = opts.plainBody || '';
		const mailOpts = {};
		if (opts.htmlBody) mailOpts.htmlBody = opts.htmlBody;
		if (opts.cc) mailOpts.cc = opts.cc;
		if (opts.attachments) mailOpts.attachments = opts.attachments;

		// In staging mode send only to admin unless explicitly overridden
		if (STAGING_MODE === 'Y' && to !== ADMIN_EMAIL) {
			Logger.log(`[EMAIL] Staging mode active - redirecting to admin instead of ${to} (${context})`);
			GmailApp.sendEmail(ADMIN_EMAIL, subject, plainBody, mailOpts);
			return true;
		}

		GmailApp.sendEmail(to, subject, plainBody, mailOpts);
		Logger.log(`[EMAIL] Sent to ${to} (${context})`);
		return true;
	} catch (e) {
		Logger.log(`safeSendEmail error (${context}): ${e.message}`);
		return false;
	}
}

// Escape text for safe HTML insertion
function escapeHtml(input) {
	if (input === null || typeof input === 'undefined') return '';
	return String(input)
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&#39;');
}

// Normalize cell/header text for comparison
function normalize(value) {
	if (value === null || typeof value === 'undefined') return '';
	try {
		return String(value).trim().toLowerCase().replace(/\s+/g, ' ');
	} catch (e) {
		return String(value || '').toLowerCase();
	}
}

// Normalize specifically for header names: remove all non-alphanumeric
function headerNormalize(value) {
	if (value === null || typeof value === 'undefined') return '';
	try {
		return String(value).toLowerCase().replace(/[^a-z0-9]+/g, '');
	} catch (e) {
		return String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
	}
}

// Normalize pixel size text like "1 x 1", "1X1", or using the multiply sign "×"
function normalizePixelSize(value) {
	if (value === null || typeof value === 'undefined') return '';
	try {
		return String(value)
			.toLowerCase()
			.replace(/\u00d7/g, 'x') // replace multiply sign × with x
			.replace(/\s+/g, '')     // remove all whitespace
			.trim();
	} catch (e) {
		return String(value || '')
			.toLowerCase()
			.replace(/\u00d7/g, 'x')
			.replace(/\s+/g, '')
			.trim();
	}
}

// Default header keywords used to locate the header row in imported reports
const headerKeywords = [
	'advertiser', 'campaign', 'site', 'placement', 'placement id',
	'placement start date', 'placement end date', 'creative', 'impressions', 'clicks'
];
const headerKeywordsCanon = headerKeywords.map(k => headerNormalize(k));

// Convert uploaded Excel/CSV blob into a Google Sheet and return an object with id
function safeConvertExcelToSheet(blob, filename, parentFolderId, configName) {
	try {
		// Ensure Advanced Drive API is available
		if (typeof Drive === 'undefined' || !Drive.Files) {
			throw new Error('Advanced Drive API not available');
		}

		const resource = {
			title: filename,
			mimeType: blob.getContentType() || MimeType.MICROSOFT_EXCEL,
			parents: parentFolderId ? [{ id: parentFolderId }] : []
		};

		// Use Drive.Files.insert with convert=true to import as Google Sheets
		const created = Drive.Files.insert(resource, blob, { convert: true });

		// created.id contains the new file id
		return { id: created.id };
	} catch (e) {
		Logger.log(`safeConvertExcelToSheet error (${configName} / ${filename}): ${e.message}`);
		throw e;
	}
}

function sendDailySummaryEmail(results) {
 const userEmail = [ADMIN_EMAIL, 'bmuller@horizonmedia.com, bkaufman@horizonmedia.com, ewarburton@horizonmedia.com'].filter(Boolean).join(', ');
 const subjectDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
 const subject = `CM360 Daily Audit Summary (${subjectDate})`;

 const rowsHtml = results.map(r => {
 const isAlert = String(r.status).toLowerCase().includes('skipped') || r.status.toLowerCase().includes('error');
 const bgColor = isAlert ? 'background-color:#ffe5e5;' : '';

 // Email status with withhold indicator
	let emailStatus = r.emailSent ? '✅' : '❌';
 if (r.emailWithheld) {
	emailStatus = '⏸️'; // Paused/withheld indicator
 }

 return `
 <tr style="font-size:12px; line-height:1.3; ${bgColor}">
 <td style="padding:4px 8px;">${escapeHtml(r.name)}</td>
 <td style="padding:4px 8px;">${escapeHtml(r.status)}</td>
 <td style="padding:4px 8px; text-align:center;">${r.flaggedRows ?? '-'}</td>
 <td style="padding:4px 8px; text-align:center;">${emailStatus}</td>
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
 <p style="font-family:Arial, sans-serif; font-size:13px;">Here's a summary of today's CM360 audits:</p>
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
 <p style="font-family:Arial, sans-serif; font-size:11px; margin-top:8px; color:#666;">
	Email Status: ✅ Sent | ❌ Failed | ⏸️ Withheld (no-flag emails disabled)
 </p>
 ${quotaNote}
 <p style="margin-top:12px; font-family:Arial, sans-serif; font-size:12px;">&mdash; Platform Solutions Team</p>
 `;

 try {
 GmailApp.sendEmail(userEmail, subject, '', { htmlBody });
 Logger.log(`[EMAIL] Summary email sent to ${userEmail}`);
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
 throw new Error(` Failed to export sheet. HTTP ${response.getResponseCode()}`);
 }

 return response.getBlob().setName(filename);
}

// === GMAIL & DRIVE FILE FETCH ===
function fetchDailyAuditAttachments(config, recipientsData) {
 Logger.log(`[IN] [${config.name}] fetchDailyAuditAttachments started`);

 const label = findGmailLabel_(config.label, config.name);
 if (!label) {
	Logger.log(`⚠️ [${config.name}] Label not found: ${config.label}`);
 safeSendEmail({
	to: resolveRecipients(config.name, recipientsData),
	cc: resolveCc(config.name, recipientsData),
	subject: `⚠️ CM360 Audit: Gmail Label Missing (${config.name})`,
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

 Logger.log(`[EXTRACT] [${config.name}] Extracted ${count} file(s) from ZIP: ${name}`);
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

 const mergedSheetName = `CM360_Merged_Audit_${new Date().toISOString().slice(0, 10)}`;
 const mergedSpreadsheet = SpreadsheetApp.create(mergedSheetName);
 Utilities.sleep(500); // Reduced from 1000ms - just need brief pause for file creation
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
 Logger.log(`[${configName}] File "${fileName}" appears blank after import.`);
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

 // Primary header detection: require all keywords
 let headerRowIndex = data.findIndex(row => {
	 const normRow = row.map(cell => normalize(cell));
	 const canonRowSet = new Set(row.map(cell => headerNormalize(cell)));
	 // strict: every keyword must appear either as normalized or canonical match
	 return headerKeywords.every(keyword => normRow.includes(normalize(keyword))) &&
					headerKeywordsCanon.every(keyword => canonRowSet.has(keyword));
 });

 // Fallback: allow partial match (at least half of keywords) if strict check failed
 if (headerRowIndex === -1) {
	 const minHits = Math.max(3, Math.ceil(headerKeywords.length / 2));
		 headerRowIndex = data.findIndex(row => {
			 const normRow = row.map(cell => normalize(cell));
			 const canonRowSet = new Set(row.map(cell => headerNormalize(cell)));
			 let hits = 0;
			 for (var k = 0; k < headerKeywords.length; k++) {
				 if (normRow.includes(normalize(headerKeywords[k])) || canonRowSet.has(headerKeywordsCanon[k])) hits++;
			 }
			 return hits >= minHits;
		 });
 }

 if (headerRowIndex === -1) {
	 Logger.log(`[${configName}] Header not found in: ${file.getName()} (strict + fallback)`);
	 continue;
 }

 const realData = data.slice(headerRowIndex);
 const cleanedData = realData.filter((row, idx) =>
	 idx === 0 || !String(row.join('') || '').toLowerCase().includes('grand total')
 );

 // Normalize pixel sizes like "1 x 1" -> "1x1" consistently for each file's data
 const pixelCols = ['Placement Pixel Size', 'Creative Pixel Size'];
 const localHeader = cleanedData[0] || [];
 const pixelColIndexes = pixelCols
	 .map(col => localHeader.findIndex(h => normalize(h) === normalize(col)))
	 .filter(idx => idx !== -1);

 // Capture header for this file when first writing
 if (!headerWritten) {
	 header = localHeader;
 }

 // Apply normalization to data rows only (do not mutate header text)
 for (let i = 1; i < cleanedData.length; i++) {
	 const row = cleanedData[i];
	 pixelColIndexes.forEach(colIdx => {
		 row[colIdx] = normalizePixelSize(row[colIdx]);
	 });
 }

 if (!headerWritten) {
	 mergedSheet.clear();
	 const bodyRows = cleanedData.slice(1);
	 if (header.length > 0) {
		 mergedSheet.getRange(1, 1, 1, header.length).setValues([header]);
		 if (bodyRows.length > 0) {
			 mergedSheet.getRange(2, 1, bodyRows.length, header.length).setValues(bodyRows);
		 }
		 headerWritten = true;
	 } else {
		 Logger.log(`[${configName}] ⚠️ Skipping write: detected header row is empty after cleaning.`);
	 }
 } else {
 const startRow = mergedSheet.getLastRow() + 1;
 const rowsToAdd = cleanedData.slice(1);
 if (rowsToAdd.length > 0 && header.length > 0) {
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
	Logger.log(`⚠️ [${configName}] Holding folder not found: ${holdingFolderPath.join(' / ')}`);
 }
 }

 const lastCol = mergedSheet.getLastColumn();
 if (lastCol && lastCol > 0) {
	 const mergedHeaders = mergedSheet.getRange(1, 1, 1, lastCol).getValues()[0];
	 Logger.log(`✅ [${configName}] Final headers in merged sheet: ${mergedHeaders.join(' | ')}`);
 } else {
	 Logger.log(`⚠️ [${configName}] No headers written; merged sheet appears empty.`);
 }
 Logger.log(`[${configName}] Merged sheet created: ${mergedSpreadsheet.getUrl()}`);
 return mergedSpreadsheet.getId();
}

// === MERGE & FLAG LOGIC ===
function executeAudit(config) {
 const now = new Date();
 const formattedNow = formatDate(now, 'yyyy-MM-dd HH:mm:ss');
 const configName = config.name;

 try {
 // Ensure Advanced Drive API is enabled before proceeding. If not, skip gracefully.
 if (!checkDriveApiEnabled()) {
	Logger.log(`⚠️ [${configName}] Skipping audit: Drive API not enabled.`);
 return { status: 'Skipped: Drive API not enabled', flaggedCount: null, emailSent: false, emailTime: formattedNow };
 }
 Logger.log(`[TEST] [${configName}] Audit started`);

 // Load configuration data for this specific config
 const exclusionsData = loadExclusionsFromSheet();
	Logger.log(`ℹ️ [${configName}] Loaded exclusions for ${Object.keys(exclusionsData).length} configs`);

 const thresholdsData = loadThresholdsFromSheet();
	Logger.log(`ℹ️ [${configName}] Loaded thresholds for ${Object.keys(thresholdsData).length} configs`);

 const recipientsData = loadRecipientsFromSheet();
	Logger.log(`ℹ️ [${configName}] Loaded recipients for ${Object.keys(recipientsData).length} configs`);

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
 <p style="margin-top:12px; font-family:Arial, sans-serif; font-size:12px;">&mdash; Platform Solutions Team</p>
 `;
	safeSendEmail({
	to: resolveRecipients(config.name, recipientsData),
	cc: resolveCc(config.name, recipientsData),
 subject, 
 htmlBody, 
 attachments: [] 
 }, configName);
 return { status: 'Skipped: No files found', flaggedCount: null, emailSent: true, emailTime: formattedNow };
 }

 const mergedSheetId = mergeDailyAuditExcels(folderId, config.mergedFolderPath, configName);
 const sheet = SpreadsheetApp.openById(mergedSheetId).getSheets()[0];
 const allData = sheet.getDataRange().getValues();

 // Detect header row using canonical, substring-aware matching (strict + fallback)
 let headerRowIndex = allData.findIndex(row => {
	 const canonCells = row.map(cell => headerNormalize(cell));
	 return headerKeywordsCanon.every(k => canonCells.some(c => c.indexOf(k) >= 0));
 });

 if (headerRowIndex === -1) {
	 const minHits = Math.max(3, Math.ceil(headerKeywordsCanon.length / 2));
	 headerRowIndex = allData.findIndex(row => {
		 const canonCells = row.map(cell => headerNormalize(cell));
		 let hits = 0;
		 for (var i = 0; i < headerKeywordsCanon.length; i++) {
			 if (canonCells.some(c => c.indexOf(headerKeywordsCanon[i]) >= 0)) hits++;
		 }
		 return hits >= minHits;
	 });
 }

 if (headerRowIndex === -1) {
	 Logger.log(`❌ [${configName}] Header row not found in merged sheet (strict + fallback).`);
	 return { status: 'Failed: Header not found', flaggedCount: null, emailSent: false, emailTime: formattedNow };
 }

 const headers = allData[headerRowIndex];
 const getIndex = name => {
	 const n = normalize(name);
	 const c = headerNormalize(name);
	 for (var i = 0; i < headers.length; i++) {
		 const h = headers[i];
		 if (normalize(h) === n || headerNormalize(h) === c) return i;
	 }
	 return -1;
 };

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
 const placementPixel = normalizePixelSize(row[fullCol['Placement Pixel Size']]);
 const creativePixel = normalizePixelSize(row[fullCol['Creative Pixel Size']]);
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

 if (updatedDataRows.length > 0) {
	 sheet.getRange(headerRowIndex + 2, 1, updatedDataRows.length, headers.length).setValues(updatedDataRows);
 }
 // Flush moved to end of formatting section for better performance

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
 // Flush moved to end of formatting section for better performance

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
 // Check if recipients have opted out of no-flag emails
 const configRecipients = recipientsData[config.name];
 
 if (!configRecipients || configRecipients.withholdNoFlagEmails) {
	Logger.log(`ℹ️ [${config.name}] No-issue email withheld: Recipients opted out of no-flag emails`);
 return { status: 'Completed (no issues)', flaggedCount: 0, emailSent: false, emailTime: formattedNow, emailWithheld: true };
 } else {
 sendNoIssueEmail(config, mergedSheetId, 'No issues were flagged', recipientsData);
 return { status: 'Completed (no issues)', flaggedCount: 0, emailSent: true, emailTime: formattedNow };
 }
 }

 } catch (err) {
 Logger.log(` [${configName}] Unexpected error: ${err.message}`);
 return { status: `Error during audit: ${err.message}`, flaggedCount: null, emailSent: false, emailTime: formattedNow };
 }
}

// === EXECUTION & AUDIT FLOW ===
function runDailyAuditByName(configName) {
 if (!checkDriveApiEnabled()) return;
 const config = auditConfigs.find(c => c.name === configName);
 if (!config) {
	Logger.log(`❌ Config "${configName}" not found.`);
 return;
 }
 executeAudit(config);
}

function runAuditBatch(configs, isFinal = false) {
 validateAuditConfigs();
 Logger.log(` Audit Batch Started: ${new Date().toLocaleString()}`);
 const results = [];

 for (const config of configs) {
 try {
 const result = executeAudit(config);
 results.push({
 name: config.name,
 status: result.status,
 flaggedRows: result.flaggedCount,
 emailSent: result.emailSent,
 emailTime: result.emailTime,
 emailWithheld: result.emailWithheld || false
 });
 } catch (err) {
 results.push({
 name: config.name,
 status: `Error: ${err.message}`,
 flaggedRows: null,
 emailSent: false,
 emailTime: formatDate(new Date(), 'yyyy-MM-dd HH:mm:ss'),
 emailWithheld: false
 });
 }
 }

 storeCombinedAuditResults_(results);

 const totalConfigs = auditConfigs.length;
 const cachedResults = getCombinedAuditResults_();

 const completedConfigs = new Set(cachedResults.map(r => r.name)).size;

	Logger.log(`✅ Completed ${completedConfigs} of ${totalConfigs} configs`);

 // Send the summary once when all configs are done, regardless of batch order
 if (completedConfigs >= totalConfigs) {
 const cache = CacheService.getScriptCache();
 const alreadySent = cache.get('CM360_SUMMARY_SENT');
 if (alreadySent === '1') {
	Logger.log('⚠️ Summary already sent by another batch. Skipping.');
 return;
 }
 const lock = LockService.getScriptLock();
 if (lock.tryLock(5000)) {
 try {
 const recheck = cache.get('CM360_SUMMARY_SENT');
 if (recheck !== '1') {
 Logger.log(`[EMAIL] All audits complete. Sending summary email...`);
 sendDailySummaryEmail(cachedResults);
 cache.put('CM360_SUMMARY_SENT', '1', 21600); // 6 hours
 CacheService.getScriptCache().remove(getAuditCacheKey_());
 } else {
	Logger.log('⚠️ Summary already sent after acquiring lock. Skipping.');
 }
 } finally {
 lock.releaseLock();
 }
 } else {
	Logger.log('⚠️ Could not acquire lock to send summary; another batch likely sending it.');
 }
 }
}

function getAuditConfigBatches(batchSize = BATCH_SIZE) {
 const batches = [];
 for (let i = 0; i < auditConfigs.length; i += batchSize) {
 batches.push(auditConfigs.slice(i, i + batchSize));
 }
 return batches;
}

function validateAuditConfigs() {
	if (!Array.isArray(auditConfigs) || auditConfigs.length === 0) {
		throw new Error('auditConfigs is empty. Define configs with name, label, mergedFolderPath, tempDailyFolderPath.');
	}
	auditConfigs.forEach((c, idx) => {
		const missing = [];
		if (!c || typeof c !== 'object') missing.push('config object');
		if (!c.name) missing.push('name');
		if (!c.label) missing.push('label');
		if (!Array.isArray(c.mergedFolderPath) || c.mergedFolderPath.length === 0) missing.push('mergedFolderPath');
		if (!Array.isArray(c.tempDailyFolderPath) || c.tempDailyFolderPath.length === 0) missing.push('tempDailyFolderPath');
		if (missing.length) {
			throw new Error(`auditConfigs[${idx}] missing: ${missing.join(', ')}`);
		}
	});
	Logger.log(`validateAuditConfigs: ${auditConfigs.length} configs ready.`);
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

// === EMAIL FLAGGED ROWS & REPORTS ===
function emailFlaggedRows(sheetId, emailRows, flaggedRows, config, recipientsData) {
 const configName = config.name;

 if (!flaggedRows || flaggedRows.length === 0) {
 Logger.log(`[${configName}] No flagged rows to report.`);
 return;
 }

 const subjectDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

 const truncate = (text, maxLen = 80) => {
 const safe = String(text || '').trim();
 return safe.length > maxLen ? safe.slice(0, maxLen - 1) + '...' : safe;
 };

	const subject = `⚠️ CM360 Daily Audit: Issues Detected (${configName} - ${subjectDate})`;

 const xlsxBlob = exportSheetAsExcel(sheetId, `CM360_DailyAudit_${configName}_${subjectDate}.xlsx`);

 const plural = (count, singular, plural) => count === 1 ? singular : plural;
 const totalFlagged = flaggedRows.length;
 const uniqueCampaigns = new Set(flaggedRows.map(r => r[1])).size;
 const verb = totalFlagged === 1 ? 'was' : 'were';
	const summaryText = `⚠️ The following ${totalFlagged} ${plural(totalFlagged, 'placement', 'placements')} across ${uniqueCampaigns} ${plural(uniqueCampaigns, 'campaign', 'campaigns')} ${verb} flagged during the <strong>${configName}</strong> CM360 audit of yesterday's delivery. Please review:`;

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
 <p style="margin-top:12px; font-family:Arial, sans-serif; font-size:12px;">&mdash; Platform Solutions Team</p>
 `;

 safeSendEmail({
 to: resolveRecipients(configName, recipientsData),
 cc: resolveCc(configName, recipientsData),
 subject,
 htmlBody,
 attachments: [xlsxBlob]
 }, `[${configName}]`);

 Logger.log(`[${configName}](c) Flagging complete: ${flaggedRows.length} row(s)`);

 return flaggedRows;
}

// === SETUP & ENVIRONMENT PREP ===
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
 Logger.log(`... Created Gmail label: ${label}`);
 } else {
 Logger.log(`[LABEL] Gmail label already exists: ${label}`);
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
 Logger.log(`[FOLDER] Created missing folder: ${pathArray.join('/')}`);
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
 msgParts.push(`[FOLDER] Created ${createdFolders.length} Drive folder path(s).`);
 }

 if (missingFilters.length > 0 && (createdLabels.length > 0 || createdFolders.length > 0)) {
	msgParts.push(`\n⚠️ The following Gmail filters may be missing:`);
 missingFilters.forEach(({ name, label }) => {
	msgParts.push(`- from:google ${name} - Label: "${label}"`);
 });
 }

 let msg = `Environment setup complete.\n\n`;

 if (msgParts.length > 0) {
 msg += msgParts.join('\n');
 } else {
 msg += `No further steps required.`;
 }

 ui.alert('Setup Summary', msg.trim(), ui.ButtonSet.OK);
}

function installDailyAuditTriggers() {
 const batches = getAuditConfigBatches(BATCH_SIZE);
 const results = [];

 // Clear existing triggers
 const existing = ScriptApp.getProjectTriggers();
 existing.forEach(trigger => {
 if (trigger.getHandlerFunction().startsWith("runDailyAuditsBatch")) {
 ScriptApp.deleteTrigger(trigger);
 results.push(` Removed trigger: ${trigger.getHandlerFunction()}`);
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
	results.push(`⚠️ Skipped trigger for ${fnName} - function not defined`);
 }
 }

 return results;
}

// === TRIGGER FUNCTIONS ===
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
}

function runDailyAuditsBatch4() {
 const batches = getAuditConfigBatches(BATCH_SIZE);
 runAuditBatch(batches[3]);
}

function runDailyAuditsBatch5() {
 const batches = getAuditConfigBatches(BATCH_SIZE);
 runAuditBatch(batches[4]);
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
 return "// All required batch runner functions are already defined.";
 }

 return `// Add these to your script:\n\n${stubs.join('\n\n')}`;
}

// === UI MENU & MODALS ===
function onOpen() {
 try {
 validateAuditConfigs();
 checkDriveApiEnabled();

 // Check if UI is available (only works when spreadsheet is actually opened)
 try {
	const ui = SpreadsheetApp.getUi();
	createAuditMenu(ui);
	// Ensure all menu/sidebar function names are present as callables to avoid missing-function runtime errors
	ensureMenuFunctionsPresent();

	// Show admin refresh prompt once per session (user property guards)
	try {
	 const props = PropertiesService.getUserProperties();
	 const seen = props.getProperty('CM360_ADMIN_REFRESH_SEEN');
	 if (!seen) {
		 const html = HtmlService.createHtmlOutputFromFile('AdminRefreshPrompt').setWidth(360).setHeight(140);
		 ui.showSidebar(html);
		 props.setProperty('CM360_ADMIN_REFRESH_SEEN', '1');
	 }
	} catch (e) {
	 Logger.log('Could not show admin refresh prompt: ' + e.message);
	}
 } catch (uiError) {
	Logger.log('⚠️ UI not available in this context - skipping menu creation');
 }
 
 } catch (error) {
	Logger.log(`❌ Error in onOpen: ${error.message}`);
 }
}

/** Force-create the UI menu (for testing) */
function forceCreateMenu() {
 try {
 const ui = SpreadsheetApp.getUi();
 createAuditMenu(ui);
	Logger.log('✅ Menu created successfully');
 } catch (error) {
	Logger.log(`❌ Error creating menu: ${error.message}`);
 }
}

function createAuditMenu(ui) {
 ui.createMenu('Admin Controls')
 // Setup
 .addItem('⚙️  Prepare Environment', 'prepareAuditEnvironment')
 .addSeparator()
 // Sheets — create/open
 .addItem('📄  Thresholds (create/open)', 'getOrCreateThresholdsSheet')
 .addItem('🚫  Exclusions (create/open)', 'getOrCreateExclusionsSheet')
 .addItem('📧  Recipients (create/open)', 'getOrCreateRecipientsSheet')
 .addSeparator()
 // External config
 .addItem('🔁  External Config: Setup Menu', 'promptSetupExternalConfigMenu')
 .addItem('📝  Ensure External Sheet Instructions', 'ensureExternalConfigInstructions')
 .addItem('⬆️  Update External Config Instructions', 'updateExternalConfigInstructions')
 .addItem('📤  Sync TO External Config', 'syncToExternalConfig')
 .addItem('📥  Sync FROM External Config', 'syncFromExternalConfig')
 .addItem('⚡  Populate External Config', 'populateExternalConfigWithDefaults')
 .addSeparator()
 // Requests
 .addItem('📝  Create Audit Request...', 'showCreateAuditRequestPicker')
 .addItem('▶️  Process Audit Requests', 'processAuditRequests')
 .addItem('🛠️  Fix Audit Requests Sheet', 'fixAuditRequestsSheet')
 .addSeparator()
 // Tools & Utilities
 .addItem('🔧  Refresh External Header Styles', 'refreshExternalHeaderStyles')
 .addItem('🔁  Update Placement Names', 'updatePlacementNamesFromReports')
 .addItem('🔐  Check Authorization', 'checkAuthorizationStatus')
 .addItem('🧾  Validate Configs', 'debugValidateAuditConfigs')
 .addItem('⏱️  Setup & Install Batch Triggers', 'setupAndInstallBatchTriggers')
 .addSeparator()
 // Manual Run Options
 .addItem('🧪  [TEST] Run Batch or Config', 'showBatchTestPicker')
 .addItem('▶️  Run Audit for...', 'showConfigPicker')
 .addSeparator()
 // Access Tools
 .addItem('📊  Open Dashboard', 'showAuditDashboard')
 .addItem('🔄  Refresh Admin Headers & Instructions', 'refreshAdminHeadersAndInstructions')
 .addItem('🧰  Buttons...', 'showButtonsSidebar')
 .addToUi();
}

/**
 * Refresh header rows, formatting, validations and INSTRUCTIONS on active/external sheets.
 * Non-destructive: preserves table data.
 */
function refreshAdminHeadersAndInstructions() {
	const ui = SpreadsheetApp.getUi();
	try {
		const targets = [SpreadsheetApp.getActiveSpreadsheet()];
		if (EXTERNAL_CONFIG_SHEET_ID) {
			try {
				targets.push(SpreadsheetApp.openById(EXTERNAL_CONFIG_SHEET_ID));
			} catch (e) {
				Logger.log('Could not open external config: ' + e.message);
			}
		}

		targets.forEach(ss => {
			try {
				// AUDIT THRESHOLDS
				let th = ss.getSheetByName(THRESHOLDS_SHEET_NAME);
				const thHeaders = ['Config Name', 'Flag Type', 'Min Impressions', 'Min Clicks', 'Active', '', 'INSTRUCTIONS'];
				if (!th) th = ss.insertSheet(THRESHOLDS_SHEET_NAME);
				th.getRange(1, 1, 1, thHeaders.length).setValues([thHeaders]);
				try { th.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff'); } catch(e) {}
				try { th.autoResizeColumns(1, 5); } catch(e) {}
				_ensureInstructionsOnSheet_(th, _buildThresholdsInstructions_());

				// AUDIT RECIPIENTS
				let rc = ss.getSheetByName(RECIPIENTS_SHEET_NAME);
				const rcHeaders = ['Config Name', 'Primary Recipients', 'CC Recipients', 'Active', 'Withhold No-Flag Emails', 'Last Updated', '', 'INSTRUCTIONS'];
				if (!rc) rc = ss.insertSheet(RECIPIENTS_SHEET_NAME);
				rc.getRange(1, 1, 1, rcHeaders.length).setValues([rcHeaders]);
				try { rc.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff'); } catch(e) {}
				try { rc.autoResizeColumns(1, 6); } catch(e) {}
				_ensureInstructionsOnSheet_(rc, _buildRecipientsInstructions_());

				// AUDIT EXCLUSIONS
				let ex = ss.getSheetByName(EXCLUSIONS_SHEET_NAME);
				const exHeaders = ['Config Name','Placement ID','Placement Name','Site Name','Name Fragment','Apply to All Configs','Flag Type','Reason','Added By','Date Added','Active','','INSTRUCTIONS'];
				if (!ex) ex = ss.insertSheet(EXCLUSIONS_SHEET_NAME);
				ex.getRange(1, 1, 1, exHeaders.length).setValues([exHeaders]);
				try { ex.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff'); } catch(e) {}
				try { ex.autoResizeColumns(1, 11); } catch(e) {}
				_ensureInstructionsOnSheet_(ex, _buildExclusionsInstructions_());

				// AUDIT REQUESTS (header only, preserve data)
				let req = ss.getSheetByName('Audit Requests');
				const reqHeaders = ['Config Name','Requested By','Requested At','Status','Notes','','INSTRUCTIONS'];
				if (!req) {
					req = ss.insertSheet('Audit Requests');
					req.getRange(1, 1, 1, reqHeaders.length).setValues([reqHeaders]);
					try { req.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff'); } catch(e) {}
				} else {
					req.getRange(1, 1, 1, reqHeaders.length).setValues([reqHeaders]);
					try { req.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff'); } catch(e) {}
				}
				_ensureInstructionsOnSheet_(req, _buildRequestsInstructions_());

				Logger.log(`refreshAdminHeadersAndInstructions: refreshed sheets for ${ss.getName()}`);
			} catch (inner) {
				Logger.log('Error refreshing headers for spreadsheet: ' + inner.message);
			}
		});

		ui.alert('Refresh Complete', 'Headers and INSTRUCTIONS refreshed on target spreadsheets. No table data was modified.', ui.ButtonSet.OK);
	} catch (err) {
	Logger.log('refreshAdminHeadersAndInstructions error: ' + err.message);
		ui.alert('Refresh Failed', `Failed to refresh headers: ${err.message}`, ui.ButtonSet.OK);
	}
}

/** Update only the external configuration spreadsheet's INSTRUCTIONS columns.
 * This is a non-destructive write to the external sheet so you can pull changes back to your local/admin copy with Sync FROM External.
 */
function refreshExternalConfigInstructions() {
	if (!EXTERNAL_CONFIG_SHEET_ID) {
		const ui = SpreadsheetApp.getUi();
		ui.alert('No External Config Sheet', 'EXTERNAL_CONFIG_SHEET_ID is not configured. Cannot update external instructions.', ui.ButtonSet.OK);
		return;
	}

	try {
		const ss = SpreadsheetApp.openById(EXTERNAL_CONFIG_SHEET_ID);

		// Helper: attempt to remove or relax protections that intersect a range so we can write instructions.
		function tryRelaxProtections(sheet, startRow, startCol, numRows, numCols) {
			try {
				const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).concat(sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET));
				const targetEndRow = startRow + Math.max(1, numRows) - 1;
				const targetEndCol = startCol + Math.max(1, numCols) - 1;
				protections.forEach(p => {
					try {
						const range = p.getRange ? p.getRange() : null;
						if (range) {
							const rStart = range.getRow();
							const rEnd = rStart + range.getNumRows() - 1;
							const cStart = range.getColumn();
							const cEnd = cStart + range.getNumColumns() - 1;
							const intersects = !(rEnd < startRow || rStart > targetEndRow || cEnd < startCol || cStart > targetEndCol);
							if (intersects) {
								// If protection is already warning-only, leave it. Otherwise try to set to warning-only so writes succeed.
								try {
									if (!p.isWarningOnly && typeof p.isWarningOnly === 'function' && !p.isWarningOnly()) {
										p.setWarningOnly(true);
										Logger.log('Relaxed protection to warning-only for write.');
									}
								} catch (inner) {
									// If we cannot change, attempt to remove (may fail depending on permissions)
									try { p.remove(); Logger.log('Removed protection to allow write.'); } catch (remErr) { Logger.log('Could not relax/remove protection: ' + remErr.message); }
								}
							}
						}
					} catch (pe) {
						Logger.log('Error evaluating protection: ' + pe.message);
					}
				});
			} catch (outer) {
				Logger.log('tryRelaxProtections: ' + outer.message);
			}
		}

		// Generic writer that ensures header cell and then writes the builder array into the fixed INSTRUCTIONS columns
		function writeInstructionsToFixedColumns(sheet, headerCol, startCol, instructionsArray) {
			try {
				// Ensure header cell says INSTRUCTIONS
				try { sheet.getRange(1, headerCol).setValue('INSTRUCTIONS').setFontWeight('bold').setBackground('#ff9900').setFontColor('#ffffff'); } catch (h) { Logger.log('Could not set INSTRUCTIONS header cell: ' + h.message); }

				const rows = instructionsArray.length || 0;
				if (rows === 0) return;

				// Attempt to relax protections covering the write range
				tryRelaxProtections(sheet, 2, startCol, rows, 2);

				// Clear the target area first (safe non-destructive for content-only ranges)
				try { sheet.getRange(2, startCol, rows, 2).clearContent(); } catch (c) { Logger.log('Could not clear target instruction range: ' + c.message); }

				// Finally write the exact builder array into the fixed columns
				try {
					sheet.getRange(2, startCol, rows, 2).setValues(instructionsArray);
				} catch (w) {
					Logger.log('Failed to write fixed-range instructions (startCol=' + startCol + '): ' + w.message);
				}
			} catch (err) {
				Logger.log('writeInstructionsToFixedColumns error: ' + err.message);
			}
		}

		// AUDIT THRESHOLDS
		let th = ss.getSheetByName(THRESHOLDS_SHEET_NAME);
		const thHeaders = ['Config Name', 'Flag Type', 'Min Impressions', 'Min Clicks', 'Active', '', 'INSTRUCTIONS'];
		if (!th) th = ss.insertSheet(THRESHOLDS_SHEET_NAME);
		th.getRange(1, 1, 1, thHeaders.length).setValues([thHeaders]);
		try { th.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff'); } catch(e) {}
		try { th.autoResizeColumns(1, 5); } catch(e) {}
		const thresholdsInstructions = _buildThresholdsInstructions_();
		_ensureInstructionsOnSheet_(th, thresholdsInstructions);
		// Write exact builder array into fixed INSTRUCTIONS columns (G=7)
		writeInstructionsToFixedColumns(th, 7, 7, thresholdsInstructions);

		// AUDIT RECIPIENTS
		let rc = ss.getSheetByName(RECIPIENTS_SHEET_NAME);
		const rcHeaders = ['Config Name', 'Primary Recipients', 'CC Recipients', 'Active', 'Withhold No-Flag Emails', 'Last Updated', '', 'INSTRUCTIONS'];
		if (!rc) rc = ss.insertSheet(RECIPIENTS_SHEET_NAME);
		rc.getRange(1, 1, 1, rcHeaders.length).setValues([rcHeaders]);
		try { rc.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff'); } catch(e) {}
		try { rc.autoResizeColumns(1, 6); } catch(e) {}
		const recipientsInstructions = _buildRecipientsInstructions_();
		_ensureInstructionsOnSheet_(rc, recipientsInstructions);
		// Write exact builder array into fixed INSTRUCTIONS columns (H=8)
		writeInstructionsToFixedColumns(rc, 8, 8, recipientsInstructions);

		// AUDIT EXCLUSIONS
		let ex = ss.getSheetByName(EXCLUSIONS_SHEET_NAME);
		const exHeaders = ['Config Name','Placement ID','Placement Name','Site Name','Name Fragment','Apply to All Configs','Flag Type','Reason','Added By','Date Added','Active','','INSTRUCTIONS'];
		if (!ex) ex = ss.insertSheet(EXCLUSIONS_SHEET_NAME);
		ex.getRange(1, 1, 1, exHeaders.length).setValues([exHeaders]);
		try { ex.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff'); } catch(e) {}
		try { ex.autoResizeColumns(1, 11); } catch(e) {}
		const exclusionsInstructions = _buildExclusionsInstructions_();
		_ensureInstructionsOnSheet_(ex, exclusionsInstructions);
		// Write exact builder array into fixed INSTRUCTIONS columns (M=13)
		writeInstructionsToFixedColumns(ex, 13, 13, exclusionsInstructions);

		// AUDIT REQUESTS (header only)
		let req = ss.getSheetByName('Audit Requests');
		const reqHeaders = ['Config Name','Requested By','Requested At','Status','Notes','','INSTRUCTIONS'];
		if (!req) {
			req = ss.insertSheet('Audit Requests');
			req.getRange(1, 1, 1, reqHeaders.length).setValues([reqHeaders]);
			try { req.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff'); } catch(e) {}
		} else {
			req.getRange(1, 1, 1, reqHeaders.length).setValues([reqHeaders]);
			try { req.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff'); } catch(e) {}
		}
		const requestsInstructions = _buildRequestsInstructions_();
		_ensureInstructionsOnSheet_(req, requestsInstructions);
		// Write exact builder array into fixed INSTRUCTIONS columns (G=7)
		writeInstructionsToFixedColumns(req, 7, 7, requestsInstructions);

		Logger.log('refreshExternalConfigInstructions: refreshed headers and instructions on external config (fixed columns written)');
		return true;
	} catch (e) {
		Logger.log('refreshExternalConfigInstructions error: ' + e.message);
		throw e;
	}
}

/** Debug function to test the external config instructions writing */
function debugExternalInstructionsWrite() {
	if (!EXTERNAL_CONFIG_SHEET_ID) {
		Logger.log('❌ EXTERNAL_CONFIG_SHEET_ID is not configured');
		return;
	}

	try {
		Logger.log('🔍 Starting debug of external instructions write...');
		Logger.log('📊 External sheet ID: ' + EXTERNAL_CONFIG_SHEET_ID);
		
		const ss = SpreadsheetApp.openById(EXTERNAL_CONFIG_SHEET_ID);
		Logger.log('✅ Successfully opened external spreadsheet: ' + ss.getName());

		// Test thresholds instructions
		const thresholdsInstructions = _buildThresholdsInstructions_();
		Logger.log('📝 Thresholds instructions built. Length: ' + thresholdsInstructions.length + ' rows');
		Logger.log('📝 First 3 instructions: ' + JSON.stringify(thresholdsInstructions.slice(0, 3)));

		let th = ss.getSheetByName(THRESHOLDS_SHEET_NAME);
		if (!th) {
			Logger.log('⚠️ Thresholds sheet not found, creating...');
			th = ss.insertSheet(THRESHOLDS_SHEET_NAME);
		}
		Logger.log('✅ Thresholds sheet found/created: ' + th.getName());

		// Try to write to column G (7)
		try {
			Logger.log('🎯 Attempting to write to Thresholds column G (7), rows 2-' + (1 + thresholdsInstructions.length));
			
			// First clear the area
			const clearRange = th.getRange(2, 7, Math.max(thresholdsInstructions.length, 10), 2);
			clearRange.clearContent();
			Logger.log('🧹 Cleared target range');

			// Now write
			const writeRange = th.getRange(2, 7, thresholdsInstructions.length, 2);
			writeRange.setValues(thresholdsInstructions);
			Logger.log('✅ Successfully wrote ' + thresholdsInstructions.length + ' instruction rows to Thresholds G:H');

			// Verify the write
			const verification = th.getRange(2, 7, 3, 2).getValues();
			Logger.log('✅ Verification read back: ' + JSON.stringify(verification));

		} catch (writeError) {
			Logger.log('❌ Failed to write thresholds instructions: ' + writeError.message);
			Logger.log('❌ Error details: ' + writeError.toString());
		}

		// Test recipients too
		Logger.log('🔄 Testing Recipients sheet...');
		const recipientsInstructions = _buildRecipientsInstructions_();
		Logger.log('📝 Recipients instructions built. Length: ' + recipientsInstructions.length + ' rows');

		let rc = ss.getSheetByName(RECIPIENTS_SHEET_NAME);
		if (!rc) {
			Logger.log('⚠️ Recipients sheet not found, creating...');
			rc = ss.insertSheet(RECIPIENTS_SHEET_NAME);
		}

		try {
			Logger.log('🎯 Attempting to write to Recipients column H (8), rows 2-' + (1 + recipientsInstructions.length));
			const writeRange = rc.getRange(2, 8, recipientsInstructions.length, 2);
			writeRange.setValues(recipientsInstructions);
			Logger.log('✅ Successfully wrote recipients instructions to H:I');
		} catch (writeError) {
			Logger.log('❌ Failed to write recipients instructions: ' + writeError.message);
		}

		// Test exclusions - this is the one you mentioned didn't update
		Logger.log('🔄 Testing Exclusions sheet...');
		const exclusionsInstructions = _buildExclusionsInstructions_();
		Logger.log('📝 Exclusions instructions built. Length: ' + exclusionsInstructions.length + ' rows');
		Logger.log('📝 First 3 exclusions instructions: ' + JSON.stringify(exclusionsInstructions.slice(0, 3)));

		let ex = ss.getSheetByName(EXCLUSIONS_SHEET_NAME);
		if (!ex) {
			Logger.log('⚠️ Exclusions sheet not found, creating...');
			ex = ss.insertSheet(EXCLUSIONS_SHEET_NAME);
		}

		try {
			Logger.log('🎯 Attempting to write to Exclusions column M (13), rows 2-' + (1 + exclusionsInstructions.length));
			const writeRange = ex.getRange(2, 13, exclusionsInstructions.length, 2);
			writeRange.setValues(exclusionsInstructions);
			Logger.log('✅ Successfully wrote exclusions instructions to M:N');
			
			// Verify the write - specifically look for placement name instruction
			const verification = ex.getRange(2, 13, 5, 2).getValues();
			Logger.log('✅ Exclusions verification read back: ' + JSON.stringify(verification));
			
		} catch (writeError) {
			Logger.log('❌ Failed to write exclusions instructions: ' + writeError.message);
		}

		// Test requests too
		Logger.log('🔄 Testing Requests sheet...');
		const requestsInstructions = _buildRequestsInstructions_();
		Logger.log('📝 Requests instructions built. Length: ' + requestsInstructions.length + ' rows');

		let req = ss.getSheetByName('Audit Requests');
		if (!req) {
			Logger.log('⚠️ Requests sheet not found, creating...');
			req = ss.insertSheet('Audit Requests');
		}

		try {
			Logger.log('🎯 Attempting to write to Requests column G (7), rows 2-' + (1 + requestsInstructions.length));
			const writeRange = req.getRange(2, 7, requestsInstructions.length, 2);
			writeRange.setValues(requestsInstructions);
			Logger.log('✅ Successfully wrote requests instructions to G:H');
		} catch (writeError) {
			Logger.log('❌ Failed to write requests instructions: ' + writeError.message);
		}

		Logger.log('🎉 Debug complete. Check the external spreadsheet manually now.');
		
	} catch (error) {
		Logger.log('❌ Debug failed: ' + error.message);
		Logger.log('❌ Stack trace: ' + error.stack);
	}
}

/**
 * User-friendly wrapper for updating external config instructions.
 * This function provides detailed logging and user feedback for the instruction update process.
 */
function updateExternalConfigInstructions() {
	const ui = SpreadsheetApp.getUi();
	
	if (!EXTERNAL_CONFIG_SHEET_ID) {
		ui.alert('No External Config Sheet', 'EXTERNAL_CONFIG_SHEET_ID is not configured. Cannot update external instructions.', ui.ButtonSet.OK);
		return;
	}

	try {
		// Show a progress message
		ui.alert('Updating Instructions', 'Updating INSTRUCTIONS columns in the external configuration spreadsheet. This may take a moment...', ui.ButtonSet.OK);
		
		// Call the debug function to do the actual work
		debugExternalInstructionsWrite();
		
		// Show success message
		ui.alert('Update Complete', 
			'Successfully updated INSTRUCTIONS columns in the external configuration spreadsheet.\n\n' +
			'Updated sheets:\n' +
			'• Audit Thresholds (columns G:H)\n' +
			'• Audit Recipients (columns H:I)\n' +
			'• Audit Exclusions (columns M:N)\n' +
			'• Audit Requests (columns G:H)\n\n' +
			'Check the execution logs for detailed information.',
			ui.ButtonSet.OK
		);
		
	} catch (error) {
		Logger.log('updateExternalConfigInstructions error: ' + error.message);
		ui.alert('Update Failed', 
			`Failed to update external instructions: ${error.message}\n\n` +
			'Check the execution logs for more details.',
			ui.ButtonSet.OK
		);
	}
}

/** Ensure all menu functions exist; create safe stubs for missing ones. */
function ensureMenuFunctionsPresent() {
	const fnNames = [
		'prepareAuditEnvironment','getOrCreateThresholdsSheet','getOrCreateExclusionsSheet',
		'promptSetupExternalConfigMenu','ensureExternalConfigInstructions','updateExternalConfigInstructions','syncToExternalConfig',
		'syncFromExternalConfig','populateExternalConfigWithDefaults','showCreateAuditRequestPicker',
		'processAuditRequests','fixAuditRequestsSheet','refreshExternalHeaderStyles',
		'updatePlacementNamesFromReports','checkAuthorizationStatus','debugValidateAuditConfigs',
		'setupAndInstallBatchTriggers','showBatchTestPicker','showConfigPicker','showAuditDashboard',
		'showButtonsSidebar','createMissingThresholds','createMissingRecipients','createMissingExclusions'
	];

	fnNames.forEach(fn => {
		try {
			if (typeof globalThis[fn] !== 'function') {
				Logger.log(`ensureMenuFunctionsPresent: Creating stub for missing function: ${fn}`);
				globalThis[fn] = function() {
					const ui = SpreadsheetApp.getUi();
					ui.alert('Missing Function', `The function "${fn}" is not implemented in this script. Please contact the administrator.`, ui.ButtonSet.OK);
				};
			}
		} catch (e) {
			Logger.log(`ensureMenuFunctionsPresent: error ensuring ${fn}: ${e.message}`);
		}
	});
}

function showButtonsSidebar() {
	const html = HtmlService.createHtmlOutputFromFile('ButtonsSidebar')
		.setTitle('CM360 Helper Buttons')
		.setWidth(300);
	SpreadsheetApp.getUi().showSidebar(html);
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
 ' Run Batch or Config',
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
 ui.alert(` Config "${configName}" not found in Batch ${batchIndex + 1}.`);
 return;
 }
 Logger.log(` Manually running config: ${config.name}`);
 executeAudit(config);
 } else {
 Logger.log(` Manually running batch ${batchIndex + 1}`);
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

// === CREATE AUDIT REQUEST (external helper bypass) ===
function showCreateAuditRequestPicker() {
 const ui = SpreadsheetApp.getUi();

 if (!EXTERNAL_CONFIG_SHEET_ID) {
 ui.alert('External Config Not Set', 'EXTERNAL_CONFIG_SHEET_ID is not configured.', ui.ButtonSet.OK);
 return;
 }

 const recipientsDataSheet = getOrCreateRecipientsSheet();
 const data = recipientsDataSheet.getDataRange().getValues();
 if (data.length <= 1) {
 ui.alert('No Data Found', 'Audit Recipients is empty or headers-only. Populate recipients first.', ui.ButtonSet.OK);
 return;
 }

 const configs = [];
 for (let i = 1; i < data.length; i++) {
 const row = data[i];
 const configName = row[0];
 const activeStatus = row[3];
 if (configName && (String(activeStatus).toUpperCase() === 'TRUE')) {
 configs.push(configName);
 }
 }
 if (configs.length === 0) {
 ui.alert('No Active Configurations Found', 'Mark at least one configuration Active (column D) in Audit Recipients.', ui.ButtonSet.OK);
 return;
 }

 const options = configs.map((name, i) => `${i + 1}. ${name}`).join('\n');
 const res = ui.prompt('Create Audit Request', 'Select a configuration to request:\n\n' + options + '\n\nEnter number (1-' + configs.length + '):', ui.ButtonSet.OK_CANCEL);
 if (res.getSelectedButton() !== ui.Button.OK) return;
 const index = parseInt(res.getResponseText().trim(), 10) - 1;
 if (isNaN(index) || index < 0 || index >= configs.length) {
 ui.alert('Invalid Selection', 'Please enter a valid number between 1 and ' + configs.length + '.', ui.ButtonSet.OK);
 return;
 }

 const configName = configs[index];
 const requester = (Session && Session.getActiveUser ? Session.getActiveUser().getEmail() : '') || '';
 const ok = createAuditRequestInExternal_(configName, requester);
 if (ok) {
 ui.alert('Request Submitted', `Added a PENDING request for ${configName} in the external sheet.`, ui.ButtonSet.OK);
 } else {
 ui.alert('Request Failed', 'Could not add a request. See logs for details.', ui.ButtonSet.OK);
 }
}

function createAuditRequestInExternal_(configName, requester) {
 try {
 const ss = SpreadsheetApp.openById(EXTERNAL_CONFIG_SHEET_ID);
 let sheet = ss.getSheetByName('Audit Requests');
 if (!sheet) {
 sheet = ss.insertSheet('Audit Requests');
 }

 // Ensure headers present (A-E + spacer + INSTRUCTIONS)
 const expectedHeaders = ['Config Name','Requested By','Requested At','Status','Notes','','INSTRUCTIONS'];
 const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), expectedHeaders.length)).getValues()[0];
 const needsHeader = currentHeaders.slice(0, expectedHeaders.length).some((h, i) => String(h || '') !== expectedHeaders[i]);
 if (needsHeader || sheet.getLastRow() === 0) {
 sheet.clear();
 sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
 sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
 }

 // Compute first empty row within A-E or append after last A-E data row
 const maxRow = Math.max(sheet.getLastRow(), 1);
 const aToE = maxRow > 0 ? sheet.getRange(1, 1, maxRow, 5).getValues() : [["Config Name","Requested By","Requested At","Status","Notes"]];
 let writeRow = -1;
 let lastDataRowAE = 1; // header row index
 for (let r = 1; r < aToE.length; r++) { // skip header
 const rowVals = aToE[r];
 const hasDataInAE = rowVals.some(v => String(v || '').trim() !== '');
 if (hasDataInAE) {
 lastDataRowAE = r + 1; // 1-based
 } else if (writeRow === -1) {
 writeRow = r + 1; // first gap
 }
 }
 if (writeRow === -1) writeRow = lastDataRowAE + 1;

 // Write request row
 const now = new Date();
 sheet.getRange(writeRow, 1, 1, 5).setValues([[configName, requester || '', now, 'PENDING', '']]);
 return true;
 } catch (e) {
 Logger.log('Failed to create audit request: ' + e.message);
 return false;
 }
}

// === CLEANUP & HOUSEKEEPING ===
function cleanupOldAuditFiles() {
 const cutoffDate = new Date();
 cutoffDate.setDate(cutoffDate.getDate() - 60);

 const trashRootPath = TRASH_ROOT_PATH;
 const deletionLogPath = DELETION_LOG_PATH;
 const adminLogName = ADMIN_LOG_NAME;

 const trashRoot = getDriveFolderByPath_(trashRootPath);
 const logFolder = getDriveFolderByPath_(deletionLogPath);

 if (!trashRoot || !logFolder) {
 Logger.log(' Cleanup failed: Trash or Log folder not found.');
 return;
 }

 const deletedFilesLog = [];
 const deletionTimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

// === DELETE LOOSE FILES IN TRASH ROOT ===
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

// === DELETE TEMP_* FOLDERS IN TEMP DAILY REPORTS ===
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
 Logger.log(`-' Deleted old temp folder: ${name}`);
 }
 }
 }
 }

// === DELETE MERGED_* FILES IN MERGED REPORTS ===
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
 Logger.log(`-' Deleted old merged file: ${name}`);
 }
 }
 }
 }

// === DELETE AGED FILES IN OTHER SUBFOLDERS & EMPTY FOLDERS ===
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
 Logger.log(`-' Deleted empty folder: ${subfolder.getName()}`);
 }
 }

// === WRITE TO LOG SHEET ===
 if (deletedFilesLog.length > 0) {
 let logSheetFile;
 let logSheet;

 const logFiles = logFolder.getFilesByName(adminLogName);
 if (logFiles.hasNext()) {
 logSheetFile = logFiles.next();
 logSheet = SpreadsheetApp.open(logSheetFile).getActiveSheet();
 } else {
 const newLog = SpreadsheetApp.create(adminLogName);
 newLog.getActiveSheet().appendRow(['File Name', 'Folder Path', 'Date Created', 'Deleted On']);
 SpreadsheetApp.flush();
 logSheetFile = DriveApp.getFileById(newLog.getId());
 logFolder.addFile(logSheetFile);
 DriveApp.getRootFolder().removeFile(logSheetFile);
 logSheet = newLog.getActiveSheet();
 }

 deletedFilesLog.forEach(row => logSheet.appendRow(row));
 SpreadsheetApp.flush();

 Logger.log(`-' Deleted ${deletedFilesLog.length} item(s). Appended to log: ${logSheetFile.getUrl()}`);
 } else {
 Logger.log('... No files or folders met deletion criteria.');
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
 const subject = ` CM360 Audit Script Needs Drive API Enabled`;
 const body = `
 The CM360 Audit script cannot run because the Advanced Drive API is not enabled.
 <br><br>
 Please enable it by opening the script editor and going to:
 <br>
 <strong>Extensions ' Apps Script ' Services</strong><br>
 Then add or enable <strong>Drive API</strong>.
 `;

 safeSendEmail({
 to: userEmail,
 subject,
 htmlBody: `<div style="font-family:Arial, sans-serif; font-size:13px;">${body}</div>`
 }, 'Drive API check');

 Logger.log(" Drive API is not enabled.");
 return false;
 }

 Logger.log("... Drive API is enabled.");
 return true;
}

function checkAuthorizationStatus() {
 const userEmail = ADMIN_EMAIL;

 try {
 const info = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

 if (info.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) {
 const subject = ` CM360 Audit Script Needs Reauthorization`;
 const message = `Your CM360 Audit script has lost authorization. Please open the script and reauthorize access.`;

 safeSendEmail({
 to: userEmail,
 subject,
 htmlBody: `<pre style="font-family:monospace; font-size:12px;">${escapeHtml(message)}</pre>`
 }, 'AUTH CHECK: reauthorization');

 return false;
 }

 Logger.log("... Authorization is valid.");
 return true;

 } catch (e) {
 Logger.log(" Failed to check authorization status:", e);

 const subject = ` CM360 Audit Script Failure`;
 const message = `The script failed to check authorization status. This may mean reauthorization is required.\n\nError: ${e.message}`;

 safeSendEmail({
 to: userEmail,
 subject,
 htmlBody: `<pre style="font-family:monospace; font-size:12px;">${escapeHtml(message)}</pre>`
 }, 'AUTH CHECK: failure fallback');

 return false;
 }
}

/**
 * Forward recent Google Apps Script failure summary emails to the configured admin.
 * Labels threads after forwarding to avoid duplicates.
 */
function forwardGASFailureNotificationsToAdmin() {
	try {
		const admin = ADMIN_EMAIL;
		const forwardedLabelName = 'GAS-Failure-Forwarded';

		// Ensure a label exists to mark forwarded threads (prevents double-forwarding)
		let forwardedLabel = GmailApp.getUserLabelByName(forwardedLabelName);
		if (!forwardedLabel) forwardedLabel = GmailApp.createLabel(forwardedLabelName);

		// Query: failure summary emails from the Apps Script notifier that are not yet labeled as forwarded
		const query = 'from:noreply-apps-scripts-notifications@google.com subject:"Summary of failures for Google Apps Script" -label:' + forwardedLabelName;
		const threads = GmailApp.search(query, 0, 100);

		if (!threads || threads.length === 0) {
			Logger.log('forwardGASFailureNotificationsToAdmin: No new GAS failure notifications found.');
			return;
		}

		threads.forEach(thread => {
			try {
				const messages = thread.getMessages();
				messages.forEach(msg => {
					// Only forward messages that are recent (configurable - here 7 days)
					const ageDays = (Date.now() - msg.getDate().getTime()) / (1000 * 60 * 60 * 24);
					if (ageDays > 7) return; // skip older messages

					const subject = msg.getSubject();
					const htmlBody = `
						<p>Forwarded Google Apps Script failure notification received by the script:</p>
						<p><strong>From:</strong> ${escapeHtml(msg.getFrom())}<br>
						<strong>Date:</strong> ${escapeHtml(String(msg.getDate()))}<br>
						<strong>Subject:</strong> ${escapeHtml(subject)}</p>
						<hr/>
						<div>${msg.getBody()}</div>
						<hr/>
						<p style="font-size:11px;color:#666;">This message was auto-forwarded by the CM360 Audit script.</p>
					`;

					// Use safeSendEmail so quota and error handling are consistent with other sends
					safeSendEmail({
						to: admin,
						subject: `GAS Failure Notification: ${subject}`,
						htmlBody: htmlBody
					}, 'GAS failure forward');
				});

				// Mark thread as forwarded
				thread.addLabel(forwardedLabel);
			} catch (inner) {
				Logger.log('forwardGASFailureNotificationsToAdmin: thread-level error: ' + inner.message);
			}
		});
		Logger.log(`forwardGASFailureNotificationsToAdmin: forwarded ${threads.length} thread(s)`);
	} catch (err) {
		Logger.log('forwardGASFailureNotificationsToAdmin error: ' + err.message);
	}
}

/**
 * Install/refresh an hourly time-driven trigger for the GAS failure forwarder.
 * Run once from the editor while signed in as the admin to create the trigger.
 */
function installGASFailureNotifierTrigger() {
	const fnName = 'forwardGASFailureNotificationsToAdmin';
	// Remove existing triggers for the function to avoid duplicates
	const existing = ScriptApp.getProjectTriggers();
	existing.forEach(t => {
		try {
			if (t.getHandlerFunction && t.getHandlerFunction() === fnName) {
				ScriptApp.deleteTrigger(t);
				Logger.log('installGASFailureNotifierTrigger: removed existing trigger for ' + fnName);
			}
		} catch (e) {
			Logger.log('installGASFailureNotifierTrigger: error while removing triggers: ' + e.message);
		}
	});

	// Create a new hourly trigger (adjust frequency if desired)
	try {
		ScriptApp.newTrigger(fnName)
			.timeBased()
			.everyHours(1)
			.create();
		Logger.log('installGASFailureNotifierTrigger: hourly trigger installed for ' + fnName);
		return 'Trigger installed: forward GAS failure notifications hourly.';
	} catch (err) {
		Logger.log('installGASFailureNotifierTrigger error: ' + err.message);
		return 'Failed to install trigger: ' + err.message;
	}
}

// ...existing code...

// === THRESHOLDS MANAGEMENT ===
function getOrCreateThresholdsSheet() {
 try {
 const spreadsheet = getConfigSpreadsheet();
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
 ['LOGIC EXPLANATION:', ''],
 ['How Evaluation Works:', 'The system compares impression vs click volume for each placement.'],
 ['', 'Whichever metric is HIGHER determines the pricing model used:'],
 ['- If Clicks > Impressions', ' Uses Min Clicks'],
 ['- If Impressions > Clicks', ' Uses Min Impressions'],
 ['', ''],
 ['EXAMPLE:', ''],
 ['Placement Data:', 'Impressions: 1,500 | Clicks: 75'],
 ['Result:', 'Since 1,500 impressions > 75 clicks Uses Min Impressions'],
 ['Threshold Check:', 'Compares against "Min Impressions" value only'],
 ['', ''],
 ['Another Example:', 'Impressions: 200 | Clicks: 850'],
 ['Result:', 'Since 850 clicks > 200 impressions Uses Min Clicks'],
 ['Threshold Check:', 'Compares against "Min Clicks" value only'],
 ['', ''],
 ['Flag Types:', ''],
 ['- clicks_greater_than_impressions', 'Flags when clicks exceed impressions'],
 ['- out_of_flight_dates', 'Flags when placement is outside flight dates'],
 ['- pixel_size_mismatch', 'Flags when creative and placement pixels differ'],
 ['- default_ad_serving', 'Flags when ad type contains "default"'],
 ['', ''],
 ['Usage:', 'Add your threshold values as needed']
 ];
 
 sheet.getRange(2, 7, instructions.length, 2).setValues(instructions);
 
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
 configName.includes('-') ||
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

// === EMAIL RECIPIENTS MANAGEMENT ===
function getOrCreateRecipientsSheet() {
 try {
 const spreadsheet = getConfigSpreadsheet();
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
 'Withhold No-Flag Emails',
 'Last Updated',
 '',
 'INSTRUCTIONS'
 ];
 
 Logger.log('Setting headers...');
 sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
 
 // Format the main headers
 const mainHeaderRange = sheet.getRange(1, 1, 1, 6);
 mainHeaderRange.setFontWeight('bold');
 mainHeaderRange.setBackground('#4285f4');
 mainHeaderRange.setFontColor('#ffffff');
 
 // Format the instructions header
 const instructionsHeaderRange = sheet.getRange(1, 8, 1, 1);
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
 
 // Add dropdown validation for Withhold No-Flag Emails column (column E) - starting from row 2
 const withholdRange = sheet.getRange('E2:E');
 const withholdRule = SpreadsheetApp.newDataValidation()
 .requireValueInList(['TRUE', 'FALSE'])
 .setAllowInvalid(false)
 .setHelpText('Set to TRUE to withhold "no issues" emails, FALSE to always send emails.')
 .build();
 
 activeRange.setDataValidation(activeRule);
 withholdRange.setDataValidation(withholdRule);
 
 // Add instructions
 const instructions = [
 ['Config Name:', 'Enter the exact config name (PST01, PST02, NEXT01, etc.)'],
 ['Primary Recipients:', 'Main email addresses (comma-separated if multiple)'],
 ['CC Recipients:', 'CC email addresses (comma-separated if multiple)'],
 ['Active:', 'TRUE to use these recipients, FALSE to disable'],
 ['Withhold No-Flag Emails:', 'TRUE to skip emails when 0 flags found, FALSE to always send emails'],
 ['Last Updated:', 'Automatically updated when you modify recipients'],
 ['', ''],
 ['Staging Mode Override:', `Currently: ${STAGING_MODE === 'Y' ? 'STAGING (all emails go to admin)' : 'PRODUCTION (uses sheet recipients)'}`],
 ['', ''],
 ['Email Format:', ''],
 ['- Single recipient:', 'user@company.com'],
 ['- Multiple recipients:', 'user1@company.com, user2@company.com'],
 ['- Leave CC blank if not needed', ''],
 ['', ''],
 ['Usage:', 'Add your recipients below as needed']
 ];
 
 sheet.getRange(2, 8, instructions.length, 2).setValues(instructions);
 
 // Format instructions
 const instructionsRange = sheet.getRange(2, 8, instructions.length, 2);
 instructionsRange.setFontSize(10);
 instructionsRange.setVerticalAlignment('top');
 
 // Auto-resize columns
 sheet.autoResizeColumns(1, 6);
 
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
 const withholdNoFlagEmails = String(row[4] || '').trim().toUpperCase();
 
 // Skip empty rows, instruction rows, or inactive recipients
 if (!configName || active !== 'TRUE' ||
 configName.includes('INSTRUCTIONS') || 
 configName.includes('-') ||
 configName.includes('Config Name:') ||
 configName.includes('Examples:')) {
 continue;
 }
 
 // Store recipient data for this config
 recipients[configName] = {
 primary: primaryRecipients,
 cc: ccRecipients,
 withholdNoFlagEmails: withholdNoFlagEmails === 'TRUE'
 };
 }
 
 Logger.log(`Loaded recipients for ${Object.keys(recipients).length} configs`);
 return recipients;
 
 } catch (error) {
	Logger.log(`❌ Error loading recipients: ${error.message}`);
 return {};
 }
}

// === EXCLUSIONS MANAGEMENT ===
function getOrCreateExclusionsSheet() {
 try {
 const spreadsheet = getConfigSpreadsheet();
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
 ['Placement Name:', 'Auto-populated - will update the following day as long as placement ID filled in with an active CM360 placement ID for the appropriate config'],
 ['Site Name:', 'Enter exact site name as it appears in CM360 reporting (alternative to Placement ID)'],
 ['Name Fragment:', 'Enter text fragment that appears in placement names (matches any placement containing this text)'],
 ['Apply to All Configs:', 'TRUE = applies to ALL config teams, FALSE = applies only to specified config'],
 ['Flag Type:', 'Select which flag type to exclude'],
 ['Reason:', 'Brief explanation for the exclusion'],
 ['Added By:', 'Your email'],
 ['Date Added:', 'Date this exclusion was added'],
 ['Active:', 'TRUE to enable, FALSE to disable'],
 ['', ''],
 ['Exclusion Types:', ''],
 ['- Placement ID', 'Excludes specific placement by ID'],
 ['- Site Name', 'Excludes all placements from a specific site'],
 ['- Name Fragment', 'Excludes placements with names containing the fragment'],
 ['', ''],
 ['Flag Types:', ''],
 ['- clicks_greater_than_impressions', 'Excludes clicks > impressions flags'],
 ['- out_of_flight_dates', 'Excludes out of flight date flags'],
 ['- pixel_size_mismatch', 'Excludes pixel mismatch flags'],
 ['- default_ad_serving', 'Excludes default ad serving flags'],
 ['- all_flags', 'Excludes from ALL flag types'],
 ['', ''],
 ['Usage:', 'Add your exclusion rules - fill in only the columns you need']
 ];
 
 sheet.getRange(2, 13, instructions.length, 2).setValues(instructions);
 
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
 !configName.includes('-') && 
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
 configName.includes('-') ||
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
 // Add all known config names from auditConfigs
 configsToApply.push(...auditConfigs.map(c => c.name));
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
 exclusions[config][flagType].siteNames.push(siteName.toLowerCase());
 }
 if (nameFragment) {
 exclusions[config][flagType].nameFragments.push(nameFragment.toLowerCase());
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
 flagExclusions.siteNames.some(site => String(site).trim() === trimmedSiteName)) {
 return true;
 }
 
 // Check name fragment exclusions
 if (flagExclusions.nameFragments && trimmedPlacementName &&
 flagExclusions.nameFragments.some(fragment => 
 trimmedPlacementName.includes(String(fragment).trim()))) {
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
  
	 // Collect target rows: Placement ID present, Placement Name blank
	 const targetsByConfig = {};
	 const missingConfigRows = [];
	 for (let i = 1; i < data.length; i++) {
		 const row = data[i];
		 const configName = String(row[0] || '').trim();
		 const placementId = String(row[1] || '').trim();
		 const currentPlacementName = String(row[2] || '').trim();
		 const headerish = configName.includes('INSTRUCTIONS') || configName.includes('Config Name:') || configName.includes('Example:');
		 if (headerish) continue;
		 if (!placementId) continue; // need a Placement ID to look up
		 if (currentPlacementName) continue; // already has a name
		 if (!configName) {
			 missingConfigRows.push(i + 1); // 1-based row index
			 continue;
		 }
		 if (!targetsByConfig[configName]) targetsByConfig[configName] = [];
		 targetsByConfig[configName].push({ rowIndex: i + 1, placementId });
	 }

	 // Fill errors for rows missing config
	 if (missingConfigRows.length > 0) {
		 missingConfigRows.forEach(r => sheet.getRange(r, 3).setValue('Error: Config Name is required'));
	 }

	 // Helper: open latest merged sheet for a config and build a map of ID -> Name
	 function buildIdToNameMap_(configName) {
		 try {
			 const cfg = auditConfigs.find(c => c.name === configName);
			 if (!cfg) return null;
			 const folder = getDriveFolderByPath_(cfg.mergedFolderPath);
			 if (!folder) return null;
			 const it = folder.getFiles();
			 const files = [];
			 while (it.hasNext()) files.push(it.next());
			 if (files.length === 0) return null;
			 files.sort((a, b) => b.getDateCreated() - a.getDateCreated());
			 // Prefer files named like our merged pattern, else fall back to newest
			 const preferred = files.find(f => String(f.getName() || '').startsWith('CM360_Merged_Audit_')) || files[0];
			 const ss = SpreadsheetApp.open(preferred);
			 const sh = ss.getSheets()[0];
			 const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
			 const findCol = (names) => {
				 for (let i = 0; i < header.length; i++) {
					 const h = header[i];
					 for (const n of names) {
						 if (normalize(h) === normalize(n) || headerNormalize(h) === headerNormalize(n)) return i;
					 }
				 }
				 return -1;
			 };
			 const idCol = findCol(['Placement ID']);
			 const nameCol = findCol(['Placement', 'Placement Name']);
			 if (idCol === -1 || nameCol === -1) return null;
			 const values = sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), sh.getLastColumn()).getValues();
			 const map = new Map();
			 for (const r of values) {
				 const id = String(r[idCol] || '').trim();
				 const nm = String(r[nameCol] || '').trim();
				 if (id) map.set(id, nm);
			 }
			 return map;
		 } catch (e) {
			 Logger.log(`buildIdToNameMap_ error for ${configName}: ${e.message}`);
			 return null;
		 }
	 }

	 // Process each config group
	 for (const [configName, rows] of Object.entries(targetsByConfig)) {
		 const idToName = buildIdToNameMap_(configName);
		 if (!idToName) {
			 // No report or no headers; mark all as not found in last CM360 report
			 rows.forEach(({ rowIndex }) => sheet.getRange(rowIndex, 3).setValue('Placement ID not found in last CM360 report'));
			 notFoundCount += rows.length;
			 continue;
		 }
		 rows.forEach(({ rowIndex, placementId }) => {
			 const name = idToName.get(String(placementId).trim());
			 if (name) {
				 sheet.getRange(rowIndex, 3).setValue(name);
				 updatedCount++;
			 } else {
				 sheet.getRange(rowIndex, 3).setValue('Placement ID not found in last CM360 report');
				 notFoundCount++;
			 }
		 });
		 // Throttle a bit between configs
		 Utilities.sleep(200);
	 }

	 const message = `Placement name update complete!\n\n` +
		`Updated: ${updatedCount}\n` +
		`Not found: ${notFoundCount}`;
	 ui.alert('Update Complete', message, ui.ButtonSet.OK);
	 Logger.log(`Update complete: ${updatedCount} updated, ${notFoundCount} not found`);
 
	} catch (error) {
	 Logger.log(` Error in updatePlacementNamesFromReports: ${error.message}`);
	 const ui = SpreadsheetApp.getUi();
	 ui.alert('Error', `Failed to update placement names: ${error.message}`, ui.ButtonSet.OK);
	}
}

// === DEBUGGING & TEST TOOLS ===
function debugValidateAuditConfigs() {
 try {
 validateAuditConfigs();
 Logger.log("... All audit configs passed validation.");
 } catch (e) {
 Logger.log(` Audit config validation failed: ${e.message}`);
 SpreadsheetApp.getUi().alert(`Audit config validation failed:\n\n${e.message}`);
 }
}

function checkMissingBatchRunners() {
 const batches = getAuditConfigBatches(BATCH_SIZE);
 const messages = [];

 for (let i = 0; i < batches.length; i++) {
 const fnName = `runDailyAuditsBatch${i + 1}`;
 if (typeof this[fnName] !== 'function') {
 messages.push(` Missing: ${fnName}`);
 } else {
 messages.push(`... Present: ${fnName}`);
 }
 }

 if (messages.length === 0) {
 messages.push(` No batch configs found.`);
 }

 return messages;
}

function runPST01Audit() { runDailyAuditByName('PST01'); }
function runPST02Audit() { runDailyAuditByName('PST02'); }
function runPST03Audit() { runDailyAuditByName('PST03'); }
function runNEXT01Audit() { runDailyAuditByName('NEXT01'); }
function runNEXT02Audit() { runDailyAuditByName('NEXT02'); }
function runNEXT03Audit() { runDailyAuditByName('NEXT03'); }
function runSPTM01Audit() { runDailyAuditByName('SPTM01'); }
function runNFL01Audit() { runDailyAuditByName('NFL01'); }

// === ENHANCED EXCLUSIONS TESTING ===
function testEnhancedExclusions() {
 try {
 Logger.log(' Testing Enhanced Exclusions System...');
 
 // Create/update exclusions sheet
 const sheet = getOrCreateExclusionsSheet();
 Logger.log('... Exclusions sheet created/updated successfully');
 
 // Test loading exclusions
 const exclusionsData = loadExclusionsFromSheet();
 Logger.log(`... Loaded exclusions data: ${JSON.stringify(exclusionsData, null, 2)}`);
 
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
 
 Logger.log('... Enhanced exclusions system test completed successfully!');
 Logger.log('" Please review the exclusions sheet and test the new features:');
 Logger.log(' Site Name column for excluding by site');
 Logger.log(' Name Fragment column for partial name matching');
 Logger.log(' Apply to All Configs for global exclusions');
 
 } catch (error) {
 Logger.log(` Error testing enhanced exclusions: ${error.message}`);
 throw error;
 }
}

// === THRESHOLDS TESTING ===
function testThresholdsSystem() {
 try {
 Logger.log(' Testing Thresholds System...');
 
 // Create/update thresholds sheet
 const sheet = getOrCreateThresholdsSheet();
 Logger.log('... Thresholds sheet created/updated successfully');
 
 // Test loading thresholds
 const thresholdsData = loadThresholdsFromSheet();
 Logger.log(`... Loaded thresholds data: ${JSON.stringify(thresholdsData, null, 2)}`);
 
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
 
 Logger.log('... Thresholds system test completed successfully!');
 Logger.log('" Please review the thresholds sheet and adjust values as needed:');
 Logger.log(' Each config has separate thresholds for each flag type');
 Logger.log(' Min Impressions OR Min Clicks threshold triggers evaluation');
 Logger.log(' Set Active to FALSE to disable specific threshold rules');
 
 } catch (error) {
 Logger.log(` Error testing thresholds system: ${error.message}`);
 throw error;
 }
}

/** Test Recipients Management System: create/update recipients sheet and validate resolution */
function testRecipientsSystem() {
 Logger.log(` Testing Recipients Management System...`);
 
 try {
 // Step 1: Create/Update recipients sheet
 Logger.log(`" Creating/updating recipients sheet...`);
 const sheet = getOrCreateRecipientsSheet();
 Logger.log(`... Recipients sheet ready with ${sheet.getLastRow() - 1} recipient configurations`);
 
 // Step 2: Load recipients data
 Logger.log(`" Loading recipients data...`);
 const recipientsData = loadRecipientsFromSheet();
 const configCount = Object.keys(recipientsData).length;
 Logger.log(`... Loaded recipients for ${configCount} configurations`);
 
 // Step 3: Test recipient resolution with sample configs
 Logger.log(`" Testing recipient resolution...`);
 let testCount = 0;
 
 for (const configName of Object.keys(recipientsData)) {
 const recipients = resolveRecipients(configName, recipientsData);
 const cc = resolveCc(configName, recipientsData);
 
 Logger.log(` " [${configName}]:`);
 Logger.log(` To: ${recipients}`);
 Logger.log(` CC: ${cc}`);
 
 testCount++;
 if (testCount >= 3) break; // Limit output for testing
 }
 
 // Step 4: Test staging mode behavior
 Logger.log(` Testing staging mode override...`);
 Logger.log(` Current staging mode: ${STAGING_MODE}`);
 
 // Show current mode recipients
 const currentRecipients = resolveRecipients('test-config', recipientsData);
 Logger.log(` Current mode recipients: ${currentRecipients}`);
 
 // Note about staging mode
 if (STAGING_MODE === 'Y') {
 Logger.log(` ... Staging mode is ENABLED - all emails go to admin`);
 } else {
 Logger.log(` ... Production mode is ENABLED - emails use sheet recipients`);
 }
 
 Logger.log(`... Recipients system test completed successfully!`);
 Logger.log(`" Summary:`);
 Logger.log(` - Recipients sheet: Ready`);
 Logger.log(` - Configurations loaded: ${configCount}`);
 Logger.log(` - Recipient resolution: Working`);
 Logger.log(` - Staging mode: ${STAGING_MODE === 'Y' ? 'ENABLED (Admin override)' : 'DISABLED (Sheet recipients)'}`);
 
 } catch (error) {
 Logger.log(` Error testing recipients system: ${error.message}`);
 throw error;
 }
}

// === " CONFIGURATION SEPARATION SETUP ===
/** Create an external configuration spreadsheet (copying configuration sheets). */
function createExternalConfigSheet() {
 try {
 Logger.log(`" Creating external configuration spreadsheet...`);
 
 // Create new spreadsheet
 const configSpreadsheet = SpreadsheetApp.create('CM360 Audit Configuration');
 const configId = configSpreadsheet.getId();
 
 Logger.log(`" Created spreadsheet: ${configId}`);
 Logger.log(`"- URL: https://docs.google.com/spreadsheets/d/${configId}`);
 
 // Copy configuration sheets from current spreadsheet
 const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 
 // Copy Recipients sheet
 try {
 const recipientsSheet = currentSpreadsheet.getSheetByName(RECIPIENTS_SHEET_NAME);
 if (recipientsSheet) {
 recipientsSheet.copyTo(configSpreadsheet).setName(RECIPIENTS_SHEET_NAME);
 Logger.log(`... Copied Recipients sheet`);
 }
 } catch (e) {
 Logger.log(` No Recipients sheet found, will create new one`);
 }
 
 // Copy Thresholds sheet
 try {
 const thresholdsSheet = currentSpreadsheet.getSheetByName(THRESHOLDS_SHEET_NAME);
 if (thresholdsSheet) {
 thresholdsSheet.copyTo(configSpreadsheet).setName(THRESHOLDS_SHEET_NAME);
 Logger.log(`... Copied Thresholds sheet`);
 }
 } catch (e) {
 Logger.log(` No Thresholds sheet found, will create new one`);
 }
 
 // Copy Exclusions sheet
 try {
 const exclusionsSheet = currentSpreadsheet.getSheetByName(EXCLUSIONS_SHEET_NAME);
 if (exclusionsSheet) {
 exclusionsSheet.copyTo(configSpreadsheet).setName(EXCLUSIONS_SHEET_NAME);
 Logger.log(`... Copied Exclusions sheet`);
 }
 } catch (e) {
 Logger.log(` No Exclusions sheet found, will create new one`);
 }
 
 // Remove default "Sheet1" if it exists
 try {
 const defaultSheet = configSpreadsheet.getSheetByName('Sheet1');
 if (defaultSheet && configSpreadsheet.getSheets().length > 1) {
 configSpreadsheet.deleteSheet(defaultSheet);
 }
 } catch (e) {
 // Ignore if Sheet1 doesn't exist
 }
 
 Logger.log(`\n SETUP INSTRUCTIONS:`);
 Logger.log(`1. Update the EXTERNAL_CONFIG_SHEET_ID constant in your source code:`);
 Logger.log(` const EXTERNAL_CONFIG_SHEET_ID = '${configId}';`);
 Logger.log(`2. Give edit access to users who need to modify configurations`);
 Logger.log(`3. Keep the source code spreadsheet private`);
 Logger.log(`4. Run setupExternalConfigMenu('${configId}') to add helper menu to config sheet`);
 Logger.log(`5. Users can now edit configurations without seeing source code`);
 
 return configId;
 
 } catch (error) {
 Logger.log(` Error creating external config sheet: ${error.message}`);
 throw error;
 }
}

/** Install a helper menu in the external configuration spreadsheet by ID. */
function setupExternalConfigMenu(configSheetId) {
 try {
 Logger.log(`" Setting up menu for external config sheet...`);
 
 if (!configSheetId) {
 throw new Error('Config sheet ID is required');
 }
 
 const configSpreadsheet = SpreadsheetApp.openById(configSheetId);
 
 // Create a simple Apps Script project for the config sheet
 const scriptProject = {
 title: 'CM360 Config Helper',
 files: [{
 name: 'Code',
 type: 'SERVER_JS',
 source: `
// === CM360 Configuration Helper Menu ===
function onOpen() {
 const ui = SpreadsheetApp.getUi();
 ui.createMenu('CM360 Config Helper')
 .addItem(' Run Config Audit', 'showConfigAuditRunner')
 .addSeparator()
 .addItem('" Create Missing Thresholds', 'createMissingThresholds')
 .addItem('" Create Missing Recipients', 'createMissingRecipients') 
 .addItem('" Create Missing Exclusions', 'createMissingExclusions')
 .addSeparator()
 .addItem('... Validate Configuration', 'validateConfiguration')
 .addItem('" Show Config Summary', 'showConfigSummary')
 .addToUi();
}

function showConfigAuditRunner() {
 const ui = SpreadsheetApp.getUi();
 
 // Get available configs from recipients sheet
 const recipientsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Audit Recipients');
 if (!recipientsSheet) {
 ui.alert('Error', 'Audit Recipients sheet not found. Please ask admin to sync configuration data.', ui.ButtonSet.OK);
 return;
 }
 
 const data = recipientsSheet.getDataRange().getValues();
 
 // Check if sheet has data
 if (data.length <= 1) {
 ui.alert(
 'No Data Found', 
 'The Audit Recipients sheet appears to be empty or only has headers.\\n\\nData rows found: ' + (data.length - 1) + '\\n\\nPlease ask admin to populate the configuration data.',
 ui.ButtonSet.OK
 );
 return;
 }
 
 const configs = [];
 
 for (let i = 1; i < data.length; i++) {
 const row = data[i];
 const configName = row[0];
 const activeStatus = row[3];
 
 if (configName && (activeStatus === 'TRUE' || activeStatus === 'true' || activeStatus === true)) {
 configs.push({
 name: configName,
 recipients: row[1] || '',
 cc: row[2] || '',
 withhold: row[4] === 'TRUE'
 });
 }
 }
 
 if (configs.length === 0) {
 ui.alert(
 'No Active Configurations Found', 
 'No active configurations found in Audit Recipients sheet.\\n\\nTotal rows: ' + data.length + '\\nPlease ensure some configurations are marked as Active (TRUE) in column D.',
 ui.ButtonSet.OK
 );
 return;
 }
 
 // Create a simple dropdown selection
 const configOptions = configs.map((config, index) => {
 const recipientCount = config.recipients.split(',').length;
 const ccCount = config.cc ? config.cc.split(',').length : 0;
 return \`\${index + 1}. \${config.name} (" \${recipientCount} recipients\${ccCount > 0 ? ', ' + ccCount + ' CC' : ''}\${config.withhold ? ', withholds no-flag emails' : ''})\`;
 }).join('\\n');
 
 const response = ui.prompt(
 'Select Configuration to Audit',
 'Available configurations:\\n\\n' + configOptions + '\\n\\nEnter the number (1-' + configs.length + ') of the configuration to audit:',
 ui.ButtonSet.OK_CANCEL
 );
 
 if (response.getSelectedButton() !== ui.Button.OK) {
 return;
 }
 
 const selectedIndex = parseInt(response.getResponseText().trim()) - 1;
 
 if (isNaN(selectedIndex) || selectedIndex < 0 || selectedIndex >= configs.length) {
 ui.alert('Invalid Selection', 'Please enter a valid number between 1 and ' + configs.length + '.', ui.ButtonSet.OK);
 return;
 }
 
 const selectedConfig = configs[selectedIndex];

 // Create request row locally in this sheet (first available row in A:E)
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 let reqSheet = ss.getSheetByName('Audit Requests');
 if (!reqSheet) {
 reqSheet = ss.insertSheet('Audit Requests');
 const headers = ['Config Name','Requested By','Requested At','Status','Notes','','INSTRUCTIONS'];
 reqSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
 reqSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
 }

 // Find the first available (empty) row within columns A-E so we fill gaps.
 // IMPORTANT: Do not let instructions (G:...) extend the write row; base only on A-E.
 const maxRow = Math.max(reqSheet.getLastRow(), 1);
 const aToE = reqSheet.getRange(1, 1, maxRow, 5).getValues();
 let writeRow = -1;
 let lastDataRowAE = 1; // track the last row with any data in A-E (1-based; header row starts at 1)

 // Scan top-down from row 2 to find the first row where A-E are all blank;
 // also track the last row that contains any A-E data to compute a safe append row.
 for (let r = 1; r < aToE.length; r++) { // aToE[0] is header
 const rowVals = aToE[r];
 const hasDataInAE = rowVals.some(v => String(v || '').trim() !== '');
 if (hasDataInAE) {
 lastDataRowAE = r + 1; // convert to 1-based row index
 } else if (writeRow === -1) {
 writeRow = r + 1; // first gap found
 }
 }

 // If no internal gap, append immediately after the last A-E data row
 if (writeRow === -1) writeRow = lastDataRowAE + 1;
 const requester = (Session && Session.getActiveUser ? Session.getActiveUser().getEmail() : '') || '';
 const now = new Date();
 reqSheet.getRange(writeRow, 1, 1, 5).setValues([[selectedConfig.name, requester, now, 'PENDING', '']]);

 SpreadsheetApp.getUi().alert('Request Submitted', 'Your audit request was added to the queue.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// (moved repair function definitions to top-level scope below)

function createMissingThresholds() {
 const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Audit Thresholds');
 if (!sheet) {
 SpreadsheetApp.getUi().alert('Audit Thresholds sheet not found. Please create it first.');
 return;
 }
 SpreadsheetApp.getUi().alert('... Audit Thresholds sheet is available for editing.');
}

function createMissingRecipients() {
 const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Audit Recipients');
 if (!sheet) {
 SpreadsheetApp.getUi().alert('Audit Recipients sheet not found. Please create it first.');
 return;
 }
 SpreadsheetApp.getUi().alert('... Audit Recipients sheet is available for editing.');
}

function createMissingExclusions() {
 const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Audit Exclusions');
 if (!sheet) {
 SpreadsheetApp.getUi().alert('Audit Exclusions sheet not found. Please create it first.');
 return;
 }
 SpreadsheetApp.getUi().alert('... Audit Exclusions sheet is available for editing.');
}

function validateConfiguration() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const ui = SpreadsheetApp.getUi();
 
 const sheets = ['Audit Thresholds', 'Audit Recipients', 'Audit Exclusions'];
 const found = sheets.filter(name => ss.getSheetByName(name) !== null);
 const missing = sheets.filter(name => ss.getSheetByName(name) === null);
 
 let message = 'Configuration Validation:\\n\\n';
 if (found.length > 0) {
 message += '... Found sheets:\\n' + found.map(s => ' ' + s).join('\\n') + '\\n\\n';
 }
 if (missing.length > 0) {
 message += ' Missing sheets:\\n' + missing.map(s => ' ' + s).join('\\n') + '\\n\\n';
 }
 
 ui.alert('Validation Results', message, ui.ButtonSet.OK);
}

function showConfigSummary() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const ui = SpreadsheetApp.getUi();
 
 let summary = 'Configuration Summary:\\n\\n';
 
 const thresholds = ss.getSheetByName('Audit Thresholds');
 if (thresholds) {
 const data = thresholds.getDataRange().getValues();
 const configs = new Set();
 for (let i = 1; i < data.length; i++) {
 if (data[i][0] && data[i][4] === 'TRUE') configs.add(data[i][0]);
 }
 summary += \`" Thresholds: \${configs.size} active configs\\n\`;
 }
 
 const recipients = ss.getSheetByName('Audit Recipients');
 if (recipients) {
 const data = recipients.getDataRange().getValues();
 const configs = new Set();
 for (let i = 1; i < data.length; i++) {
 if (data[i][0] && data[i][3] === 'TRUE') configs.add(data[i][0]);
 }
 summary += \`" Recipients: \${configs.size} active configs\\n\`;
 }
 
 const exclusions = ss.getSheetByName('Audit Exclusions');
 if (exclusions) {
 const data = exclusions.getDataRange().getValues();
 const activeRows = data.slice(1).filter(row => row[0] && row[4] === 'TRUE').length;
 summary += \`" Exclusions: \${activeRows} active rules\\n\`;
 }
 
 ui.alert('Configuration Summary', summary, ui.ButtonSet.OK);
}
`
 }]
 };
 
 Logger.log(`... Helper menu setup completed for config sheet`);
 Logger.log(`" To add the menu:`);
 Logger.log(`1. Open the config spreadsheet: https://docs.google.com/spreadsheets/d/${configSheetId}`);
 Logger.log(`2. Go to Extensions ' Apps Script`);
 Logger.log(`3. Replace the default code with the helper menu code`);
 Logger.log(`4. Save the project and refresh the spreadsheet`);
 
 return true;
 
 } catch (error) {
 Logger.log(` Error setting up external config menu: ${error.message}`);
 throw error;
 }
}

/** Verify access to external configuration and load sample data */
function testExternalConfig() {
 if (!EXTERNAL_CONFIG_SHEET_ID) {
 Logger.log(` EXTERNAL_CONFIG_SHEET_ID is not set. Using current spreadsheet.`);
 return;
 }
 
 try {
 Logger.log(` Testing external configuration access...`);
 const configSpreadsheet = getConfigSpreadsheet();
 Logger.log(`... Successfully accessed external config: ${configSpreadsheet.getName()}`);
 
 // Test loading each type of configuration
 const recipients = loadRecipientsFromSheet();
 const thresholds = loadThresholdsFromSheet();
 const exclusions = loadExclusionsFromSheet();
 
 Logger.log(`... External configuration test passed!`);
 Logger.log(`" Loaded: ${Object.keys(recipients).length} recipients, ${Object.keys(thresholds).length} thresholds, ${Object.keys(exclusions).length} exclusions`);
 
 } catch (error) {
 Logger.log(` External configuration test failed: ${error.message}`);
 throw error;
 }
}

/** Prompt user with instructions to install the external config helper menu */
function promptSetupExternalConfigMenu() {
 const ui = SpreadsheetApp.getUi();
 
 if (!EXTERNAL_CONFIG_SHEET_ID) {
 ui.alert(
 'External Config Not Set', 
 'No external configuration sheet is currently configured.\n\nRun "createExternalConfigSheet()" first to create an external config sheet.', 
 ui.ButtonSet.OK
 );
 return;
 }
 
 const response = ui.alert(
 'Setup External Config Menu',
 `This will provide instructions to add a helper menu to your external configuration spreadsheet.\n\nConfig Sheet ID: ${EXTERNAL_CONFIG_SHEET_ID}\n\nContinue?`,
 ui.ButtonSet.YES_NO
 );
 
 if (response === ui.Button.YES) {
 setupExternalConfigMenu(EXTERNAL_CONFIG_SHEET_ID);
 ui.alert(
 'Setup Instructions Logged',
 'Check the execution log for detailed setup instructions.\n\nThe helper menu will provide basic validation and summary functions for your external config sheet.',
 ui.ButtonSet.OK
 );
 }
}

// === " SYNC FUNCTIONS FOR EXTERNAL CONFIG ===

// === Repair tools for Audit Requests sheet (external config) ===
function fixAuditRequestsSheet() {
 const ui = SpreadsheetApp.getUi();
 if (!EXTERNAL_CONFIG_SHEET_ID) {
 ui.alert('No External Config Sheet', 'EXTERNAL_CONFIG_SHEET_ID is not set.', ui.ButtonSet.OK);
 return;
 }

 try {
 const ss = SpreadsheetApp.openById(EXTERNAL_CONFIG_SHEET_ID);
 const sheet = ss.getSheetByName('Audit Requests') || ss.insertSheet('Audit Requests');

 // Ensure header exists (A-E + spacer + INSTRUCTIONS)
 const expectedHeaders = ['Config Name','Requested By','Requested At','Status','Notes','','INSTRUCTIONS'];
 const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), expectedHeaders.length)).getValues()[0];
 const needsHeader = currentHeaders.slice(0, expectedHeaders.length).some(function(h, i){ return String(h||'') !== expectedHeaders[i]; });
 if (needsHeader || sheet.getLastRow() === 0) {
 sheet.clear();
 sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
 sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
 }

 // Re-apply multi-row instructions for readability
 (function(){
 const lastCol = sheet.getLastColumn() || 1;
 const headerRow = sheet.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
 let instrColIndex = -1;
 for (var c = 0; c < headerRow.length; c++) {
 if (String(headerRow[c] || '').trim().toUpperCase() === 'INSTRUCTIONS') { instrColIndex = c + 1; break; }
 }
 if (instrColIndex === -1) {
 var lastHeaderCol = 0;
 for (var c2 = headerRow.length - 1; c2 >= 0; c2--) {
 if (String(headerRow[c2] || '').trim() !== '') { lastHeaderCol = c2 + 1; break; }
 }
 instrColIndex = Math.max(1, lastHeaderCol) + 2;
 sheet.getRange(1, instrColIndex, 1, 1).setValue('INSTRUCTIONS');
 }
 // Style header
 sheet.getRange(1, instrColIndex, 1, 1).setFontWeight('bold').setBackground('#ff9900').setFontColor('#ffffff');
 // Write instructions (two columns)
 const requestsInstructions = _buildRequestsInstructions_();
 sheet.getRange(2, instrColIndex, requestsInstructions.length, 2).clearContent();
 sheet.getRange(2, instrColIndex, requestsInstructions.length, 2).setValues(requestsInstructions);
 sheet.getRange(2, instrColIndex, requestsInstructions.length, 2).setFontSize(10).setVerticalAlignment('top').setWrap(false);
 // Reset first few row heights in case row 2 ballooned
 try { sheet.setRowHeights(1, Math.min(30, sheet.getMaxRows()), 21); } catch(e) {}
 })();

 // Repack A-E rows to remove gaps
 _repackPrimaryColumnsAE_(sheet);

 ui.alert('Audit Requests Sheet Repaired', 'Instructions refreshed (multi-row) and A-E rows repacked to remove gaps.', ui.ButtonSet.OK);
 } catch (err) {
 ui.alert('Repair Failed', 'Error repairing Audit Requests: ' + err.message, ui.ButtonSet.OK);
 }
}

function _repackPrimaryColumnsAE_(sheet) {
 // Read all data starting at row 2, columns A-E
 const lastRow = Math.max(sheet.getLastRow(), 1);
 const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 5).getValues() : [];
 if (rows.length === 0) return;

 const dataRows = rows.filter(function(r){ return r.some(function(v){ return String(v || '').trim() !== ''; }); });
 // Clear all A-E below header
 if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
 // Write back packed data starting at row 2
 if (dataRows.length > 0) sheet.getRange(2, 1, dataRows.length, 5).setValues(dataRows);
}

// Optional: installable time-based auto-fix trigger for the external sheet
function installAutoFixForRequestsSheet() {
 if (!EXTERNAL_CONFIG_SHEET_ID) {
 SpreadsheetApp.getUi().alert('No External Config Sheet', 'EXTERNAL_CONFIG_SHEET_ID is not set.', SpreadsheetApp.getUi().ButtonSet.OK);
 return;
 }
 // Remove existing triggers for this handler
 ScriptApp.getProjectTriggers().forEach(function(t){
 if (t.getHandlerFunction && t.getHandlerFunction() === 'autoFixRequestsSheet_') {
 ScriptApp.deleteTrigger(t);
 }
 });
 ScriptApp.newTrigger('autoFixRequestsSheet_').timeBased().everyHours(4).create();
 SpreadsheetApp.getUi().alert('Auto-Fix Installed', 'A time-based trigger will compact the Audit Requests sheet every 4 hours.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function autoFixRequestsSheet_() {
 try {
 fixAuditRequestsSheet();
 } catch (e) {
 Logger.log('Auto-fix error: ' + e.message);
 }
}
/** Sync configuration sheets (Recipients, Thresholds, Exclusions) TO external config spreadsheet */
function syncToExternalConfig() {
 if (!EXTERNAL_CONFIG_SHEET_ID) {
 const ui = SpreadsheetApp.getUi();
 ui.alert(
 'No External Config Sheet',
 'No external configuration sheet is configured.\n\nRun "createExternalConfigSheet()" first to create one.',
 ui.ButtonSet.OK
 );
 return;
 }

 try {
 Logger.log(`" Starting sync to external config sheet: ${EXTERNAL_CONFIG_SHEET_ID}`);
 
 const mainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 const externalSpreadsheet = SpreadsheetApp.openById(EXTERNAL_CONFIG_SHEET_ID);
 
 const sheetsToSync = [
 { name: RECIPIENTS_SHEET_NAME, description: 'Recipients' },
 { name: THRESHOLDS_SHEET_NAME, description: 'Thresholds' },
 { name: EXCLUSIONS_SHEET_NAME, description: 'Exclusions' }
 ];
 
 const syncResults = [];
 
 for (const sheetInfo of sheetsToSync) {
 try {
 const mainSheet = mainSpreadsheet.getSheetByName(sheetInfo.name);
 if (!mainSheet) {
 syncResults.push(` ${sheetInfo.description}: Source sheet not found`);
 continue;
 }
 
 // Get or create target sheet in external spreadsheet
 let externalSheet = externalSpreadsheet.getSheetByName(sheetInfo.name);
 if (!externalSheet) {
 externalSheet = externalSpreadsheet.insertSheet(sheetInfo.name);
 }
 
 // Clear existing content
 externalSheet.clear();
 
 // Copy data from main to external
 const mainData = mainSheet.getDataRange().getValues();
 if (mainData.length > 0) {
 externalSheet.getRange(1, 1, mainData.length, mainData[0].length).setValues(mainData);
 }
 
 // Copy formatting and validations manually (copyTo cannot be used across spreadsheets)
 const mainRange = mainSheet.getDataRange();
 const numRows = mainRange.getNumRows();
 const numCols = mainRange.getNumColumns();
 const externalRange = externalSheet.getRange(1, 1, numRows, numCols);

 try {
	 // Basic formatting
	 try {
		 externalRange.setBackgrounds(mainRange.getBackgrounds());
		 externalRange.setFontColors(mainRange.getFontColors());
		 externalRange.setFontFamilies(mainRange.getFontFamilies());
		 externalRange.setFontSizes(mainRange.getFontSizes());
		 externalRange.setFontWeights(mainRange.getFontWeights());
		 externalRange.setFontStyles(mainRange.getFontStyles());
		 externalRange.setNumberFormats(mainRange.getNumberFormats());
		 externalRange.setHorizontalAlignments(mainRange.getHorizontalAlignments());
		 externalRange.setVerticalAlignments(mainRange.getVerticalAlignments());
	 } catch (formatErr) {
		 Logger.log(` Could not copy some formatting for ${sheetInfo.name}: ${formatErr.message}`);
	 }

	 // Copy data validations (whole-range if supported)
	 try {
		 const validations = mainRange.getDataValidations();
		 if (validations) externalRange.setDataValidations(validations);
	 } catch (valErr) {
		 Logger.log(` Could not copy validations for ${sheetInfo.name}: ${valErr.message}`);
	 }

	 // Copy column widths
	 try {
		 for (let col = 1; col <= numCols; col++) {
			 externalSheet.setColumnWidth(col, mainSheet.getColumnWidth(col));
		 }
	 } catch (widthErr) {
		 Logger.log(` Could not copy column widths for ${sheetInfo.name}: ${widthErr.message}`);
	 }

	 // Copy row heights for the first N rows (instructions and headers typically in the top rows)
	 try {
		 const maxRowCopy = Math.min(50, numRows);
		 for (let r = 1; r <= maxRowCopy; r++) {
			 externalSheet.setRowHeight(r, mainSheet.getRowHeight(r));
		 }
	 } catch (heightErr) {
		 Logger.log(` Could not copy row heights for ${sheetInfo.name}: ${heightErr.message}`);
	 }

	 // Copy protections (range protections only)
	 try {
		 const protections = mainSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
		 protections.forEach(p => {
			 try {
				 const rng = p.getRange();
				 const target = externalSheet.getRange(rng.getA1Notation());
				 const newProt = target.protect();
				 newProt.setDescription(p.getDescription());
				 if (p.isWarningOnly()) newProt.setWarningOnly(true);
			 } catch (pe) {
				 Logger.log(` Could not copy a protection for ${sheetInfo.name}: ${pe.message}`);
			 }
		 });
	 } catch (protErr) {
		 Logger.log(` Could not copy protections for ${sheetInfo.name}: ${protErr.message}`);
	 }
 } catch (err) {
	 Logger.log(` Error copying formats/validations to external for ${sheetInfo.name}: ${err.message}`);
 }
 
 syncResults.push(`... ${sheetInfo.description}: Synced ${mainData.length} rows`);
 
 } catch (error) {
 syncResults.push(` ${sheetInfo.description}: Error - ${error.message}`);
 Logger.log(` Error syncing ${sheetInfo.name}: ${error.message}`);
 }
 }
 
 // Show results
 const ui = SpreadsheetApp.getUi();
 const resultMessage = syncResults.join('\n');
 ui.alert(
 'Sync Results',
 `Sync to external config sheet completed:\n\n${resultMessage}\n\nExternal sheet: CM360 Audit Configuration - Helper Menu`,
 ui.ButtonSet.OK
 );
 
 Logger.log(`... Sync completed. Results:\n${resultMessage}`);
 
 } catch (error) {
 Logger.log(` Error during sync: ${error.message}`);
 const ui = SpreadsheetApp.getUi();
 ui.alert(
 'Sync Error',
 `Failed to sync to external config sheet:\n\n${error.message}`,
 ui.ButtonSet.OK
 );
 }
}

/**
 * Syncs data from external config sheet back to main spreadsheet
 * Preserves formatting, instructions, validations, and includes Audit Requests sheet
 */
function syncFromExternalConfig() {
	// Detect whether a Spreadsheet UI is available (true when run via menu/button)
	let ui = null;
	try {
		ui = SpreadsheetApp.getUi();
	} catch (e) {
		ui = null; // Running from trigger or other non-UI context
	}

	if (!EXTERNAL_CONFIG_SHEET_ID) {
		if (ui) {
			ui.alert(
				'No External Config Sheet',
				'No external configuration sheet is configured.',
				ui.ButtonSet.OK
			);
		} else {
			Logger.log('syncFromExternalConfig: EXTERNAL_CONFIG_SHEET_ID is not configured.');
		}
		return;
	}

 try {
 Logger.log(`" Starting sync from external config sheet: ${EXTERNAL_CONFIG_SHEET_ID}`);
 
 const mainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 const externalSpreadsheet = SpreadsheetApp.openById(EXTERNAL_CONFIG_SHEET_ID);
 
 const sheetsToSync = [
 { name: RECIPIENTS_SHEET_NAME, description: 'Recipients' },
 { name: THRESHOLDS_SHEET_NAME, description: 'Thresholds' },
 { name: EXCLUSIONS_SHEET_NAME, description: 'Exclusions' },
 { name: 'Audit Requests', description: 'Audit Requests' }
 ];
 
 const syncResults = [];
 
 for (const sheetInfo of sheetsToSync) {
 try {
 const externalSheet = externalSpreadsheet.getSheetByName(sheetInfo.name);
 if (!externalSheet) {
 syncResults.push(` ${sheetInfo.description}: External sheet not found`);
 continue;
 }
 
 // Get or create target sheet in main spreadsheet
 let mainSheet = mainSpreadsheet.getSheetByName(sheetInfo.name);
 if (!mainSheet) {
 mainSheet = mainSpreadsheet.insertSheet(sheetInfo.name);
 }
 
 // Clear existing content
 mainSheet.clear();
 
 // Get the data range from external sheet
 const externalRange = externalSheet.getDataRange();
 const numRows = externalRange.getNumRows();
 const numCols = externalRange.getNumColumns();
 
 if (numRows > 0) {
 const targetRange = mainSheet.getRange(1, 1, numRows, numCols);
 
 // Copy values
 const values = externalRange.getValues();
 targetRange.setValues(values);
 
 // Copy formatting manually since copyTo doesn't work across spreadsheets
 try {
 // Copy basic formatting
 const backgrounds = externalRange.getBackgrounds();
 const fontColors = externalRange.getFontColors();
 const fontFamilies = externalRange.getFontFamilies();
 const fontSizes = externalRange.getFontSizes();
 const fontWeights = externalRange.getFontWeights();
 const fontStyles = externalRange.getFontStyles();
 const horizontalAlignments = externalRange.getHorizontalAlignments();
 const verticalAlignments = externalRange.getVerticalAlignments();
 
 targetRange.setBackgrounds(backgrounds);
 targetRange.setFontColors(fontColors);
 targetRange.setFontFamilies(fontFamilies);
 targetRange.setFontSizes(fontSizes);
 targetRange.setFontWeights(fontWeights);
 targetRange.setFontStyles(fontStyles);
 targetRange.setHorizontalAlignments(horizontalAlignments);
 targetRange.setVerticalAlignments(verticalAlignments);
 } catch (formatError) {
 Logger.log(` Could not copy some formatting for ${sheetInfo.name}: ${formatError.message}`);
 }
 
 // Copy data validations manually
 try {
 for (let row = 1; row <= numRows; row++) {
 for (let col = 1; col <= numCols; col++) {
 const externalCell = externalSheet.getRange(row, col);
 const validation = externalCell.getDataValidation();
 if (validation) {
 const mainCell = mainSheet.getRange(row, col);
 mainCell.setDataValidation(validation);
 }
 }
 }
 } catch (validationError) {
 Logger.log(` Could not copy validations for ${sheetInfo.name}: ${validationError.message}`);
 }
 
 // Copy column widths
 try {
 for (let col = 1; col <= numCols; col++) {
 const width = externalSheet.getColumnWidth(col);
 mainSheet.setColumnWidth(col, width);
 }
 } catch (widthError) {
 Logger.log(` Could not copy column widths for ${sheetInfo.name}: ${widthError.message}`);
 }
 
 // Copy row heights for first 20 rows (where instructions typically are)
 try {
 for (let row = 1; row <= Math.min(20, numRows); row++) {
 const height = externalSheet.getRowHeight(row);
 mainSheet.setRowHeight(row, height);
 }
 } catch (heightError) {
 Logger.log(` Could not copy row heights for ${sheetInfo.name}: ${heightError.message}`);
 }
 
 // Copy sheet-level protections
 try {
 const externalProtections = externalSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
 externalProtections.forEach(protection => {
 try {
 const range = protection.getRange();
 const mainRange = mainSheet.getRange(range.getA1Notation());
 const newProtection = mainRange.protect();
 newProtection.setDescription(protection.getDescription());
 if (protection.isWarningOnly()) {
 newProtection.setWarningOnly(true);
 }
 } catch (protectionError) {
 Logger.log(` Could not copy protection: ${protectionError.message}`);
 }
 });
 } catch (protectionsError) {
 Logger.log(` Could not copy protections for ${sheetInfo.name}: ${protectionsError.message}`);
 }
 }
 
 syncResults.push(`... ${sheetInfo.description}: Synced ${numRows} rows with formatting`);
 
 } catch (error) {
 syncResults.push(` ${sheetInfo.description}: Error - ${error.message}`);
 Logger.log(` Error syncing ${sheetInfo.name}: ${error.message}`);
 }
 }
 
		// Show results (only when UI available)
		const resultMessage = syncResults.join('\n');
		if (ui) {
			ui.alert(
				'Sync Results',
				`Sync from external config sheet completed:\n\n${resultMessage}\n\nAll formatting, validations, and instructions preserved.`,
				ui.ButtonSet.OK
			);
		} else {
			Logger.log(`syncFromExternalConfig: completed. Results:\n${resultMessage}`);
		}
 
 Logger.log(`... Sync completed. Results:\n${resultMessage}`);
 
	} catch (error) {
		Logger.log(` Error during sync: ${error.message}`);
		if (ui) {
			ui.alert(
				'Sync Error',
				`Failed to sync from external config sheet:\n\n${error.message}`,
				ui.ButtonSet.OK
			);
		} else {
			// When run from a trigger, notify admin via email so failures are visible
			safeSendEmail({
				to: ADMIN_EMAIL,
				subject: `CM360: syncFromExternalConfig failed`,
				htmlBody: `<pre style="font-family:monospace">${escapeHtml(error.message)}</pre>`
			}, 'syncFromExternalConfig error');
		}
	}
}

/** Ensure external config tabs contain standardized INSTRUCTIONS sections. */
function ensureExternalConfigInstructions() {
 const ui = SpreadsheetApp.getUi();
 if (!EXTERNAL_CONFIG_SHEET_ID) {
 ui.alert(
 'No External Config Sheet',
 'EXTERNAL_CONFIG_SHEET_ID is not set. Create or set an external configuration sheet first.',
 ui.ButtonSet.OK
 );
 return;
 }

 const ss = SpreadsheetApp.openById(EXTERNAL_CONFIG_SHEET_ID);

 // Ensure all known tabs exist (Recipients, Thresholds, Exclusions, Audit Requests)
 const sheets = {
 [RECIPIENTS_SHEET_NAME]: ss.getSheetByName(RECIPIENTS_SHEET_NAME),
 [THRESHOLDS_SHEET_NAME]: ss.getSheetByName(THRESHOLDS_SHEET_NAME),
 [EXCLUSIONS_SHEET_NAME]: ss.getSheetByName(EXCLUSIONS_SHEET_NAME),
 'Audit Requests': ss.getSheetByName('Audit Requests')
 };

 // Create missing sheets with basic headers so we can place instructions
 if (!sheets[RECIPIENTS_SHEET_NAME]) {
 sheets[RECIPIENTS_SHEET_NAME] = ss.insertSheet(RECIPIENTS_SHEET_NAME);
 const headers = [
 'Config Name',
 'Primary Recipients',
 'CC Recipients',
 'Active',
 'Withhold No-Flag Emails',
 'Last Updated',
 '',
 'INSTRUCTIONS'
 ];
 sheets[RECIPIENTS_SHEET_NAME].getRange(1, 1, 1, headers.length).setValues([headers]);
 const headerRange = sheets[RECIPIENTS_SHEET_NAME].getRange(1, 1, 1, 6);
 headerRange.setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
 const instrHeader = sheets[RECIPIENTS_SHEET_NAME].getRange(1, 8, 1, 1);
 instrHeader.setFontWeight('bold').setBackground('#ff9900').setFontColor('#ffffff');
 }

 if (!sheets[THRESHOLDS_SHEET_NAME]) {
 sheets[THRESHOLDS_SHEET_NAME] = ss.insertSheet(THRESHOLDS_SHEET_NAME);
 const headers = [
 'Config Name',
 'Flag Type',
 'Min Impressions',
 'Min Clicks',
 'Active',
 '',
 'INSTRUCTIONS'
 ];
 sheets[THRESHOLDS_SHEET_NAME].getRange(1, 1, 1, headers.length).setValues([headers]);
 const headerRange = sheets[THRESHOLDS_SHEET_NAME].getRange(1, 1, 1, 5);
 headerRange.setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
 const instrHeader = sheets[THRESHOLDS_SHEET_NAME].getRange(1, 7, 1, 1);
 instrHeader.setFontWeight('bold').setBackground('#ff9900').setFontColor('#ffffff');
 }

 if (!sheets[EXCLUSIONS_SHEET_NAME]) {
 sheets[EXCLUSIONS_SHEET_NAME] = ss.insertSheet(EXCLUSIONS_SHEET_NAME);
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
 sheets[EXCLUSIONS_SHEET_NAME].getRange(1, 1, 1, headers.length).setValues([headers]);
 const headerRange = sheets[EXCLUSIONS_SHEET_NAME].getRange(1, 1, 1, 11);
 headerRange.setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
 const instrHeader = sheets[EXCLUSIONS_SHEET_NAME].getRange(1, 13, 1, 1);
 instrHeader.setFontWeight('bold').setBackground('#ff9900').setFontColor('#ffffff');
 }

 if (!sheets['Audit Requests']) {
 sheets['Audit Requests'] = ss.insertSheet('Audit Requests');
 const headers = [
 'Config Name',
 'Requested By',
 'Requested At',
 'Status',
 'Notes',
 '',
 'INSTRUCTIONS'
 ];
 sheets['Audit Requests'].getRange(1, 1, 1, headers.length).setValues([headers]);
 const headerRange = sheets['Audit Requests'].getRange(1, 1, 1, 5);
 headerRange.setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
 const instrHeader = sheets['Audit Requests'].getRange(1, 7, 1, 1);
 instrHeader.setFontWeight('bold').setBackground('#ff9900').setFontColor('#ffffff');
 }

 // Prepare instruction payloads matching Exclusions-style descriptiveness
 const exclusionsInstructions = _buildExclusionsInstructions_();
 const thresholdsInstructions = _buildThresholdsInstructions_();
 const recipientsInstructions = _buildRecipientsInstructions_();
 const requestsInstructions = _buildRequestsInstructions_();

 // Apply or update instruction blocks
 _ensureInstructionsOnSheet_(sheets[EXCLUSIONS_SHEET_NAME], exclusionsInstructions);
 _ensureInstructionsOnSheet_(sheets[THRESHOLDS_SHEET_NAME], thresholdsInstructions);
 _ensureInstructionsOnSheet_(sheets[RECIPIENTS_SHEET_NAME], recipientsInstructions);
 // Use standard multi-row instructions for Audit Requests to preserve readability
 _ensureInstructionsOnSheet_(sheets['Audit Requests'], requestsInstructions);

 ui.alert(
 'Instructions Updated',
 'All external config tabs now include a standardized INSTRUCTIONS section matching the Exclusions tab style.',
 ui.ButtonSet.OK
 );
}

/** Reapply header styling (visual only) in the external configuration spreadsheet. */
function refreshExternalHeaderStyles() {
	if (!EXTERNAL_CONFIG_SHEET_ID) {
		Logger.log('No EXTERNAL_CONFIG_SHEET_ID configured; cannot refresh external headers.');
		return;
	}

	try {
		const ss = SpreadsheetApp.openById(EXTERNAL_CONFIG_SHEET_ID);

		const sheetNames = ['Audit Requests', 'Audit Thresholds', 'Audit Recipients', 'Audit Exclusions'];

		sheetNames.forEach(name => {
			try {
				const sh = ss.getSheetByName(name);
				if (!sh) {
					Logger.log(`refreshExternalHeaderStyles: sheet not found: ${name}`);
					return;
				}

				const lastCol = Math.max(1, sh.getLastColumn());
				const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0] || [];

				// Find the first blank header cell which separates the table from INSTRUCTIONS
				let spacerIndex = headerRow.findIndex(h => String(h || '').trim() === '');
				let dataCols = 0;
				let instrCol = -1;

				if (spacerIndex === -1) {
					// No explicit blank column found; treat entire header as the table
					dataCols = lastCol;
					// Try to find an INSTRUCTIONS header anywhere; if present, note its column
					const instrFound = headerRow.findIndex(h => String(h || '').trim().toUpperCase() === 'INSTRUCTIONS');
					if (instrFound !== -1) instrCol = instrFound + 1; // 1-based
				} else {
					// spacerIndex is 0-based index of the blank spacer column
					dataCols = Math.max(0, spacerIndex); // number of data columns to the left of the spacer
					// INSTRUCTIONS expected to be one column to the right of the spacer
					const candidate = spacerIndex + 2; // convert to 1-based and skip spacer
					if (candidate <= lastCol && String(sh.getRange(1, candidate).getValue() || '').trim().toUpperCase() === 'INSTRUCTIONS') {
						instrCol = candidate;
					} else {
						// As a fallback, search the header row for INSTRUCTIONS
						const instrFound = headerRow.findIndex(h => String(h || '').trim().toUpperCase() === 'INSTRUCTIONS');
						if (instrFound !== -1) instrCol = instrFound + 1;
					}
				}

				// Clear background and reset font color across the whole header row first
				try {
					sh.getRange(1, 1, 1, lastCol).setBackground(null).setFontColor('#000000');
				} catch (e) {
					// ignore
				}

				// Apply blue header to columns 1..dataCols (only if there are data columns)
				if (dataCols > 0) {
					try {
						const headerRange = sh.getRange(1, 1, 1, dataCols);
						headerRange.setFontWeight('bold');
						headerRange.setBackground('#4285f4');
						headerRange.setFontColor('#ffffff');
					} catch (e) {
						Logger.log(`refreshExternalHeaderStyles: could not style header range for ${name}: ${e.message}`);
					}
				}

				// Style the INSTRUCTIONS header cell (if found)
				if (instrCol > 0 && instrCol <= lastCol) {
					try {
						const instrRange = sh.getRange(1, instrCol, 1, 1);
						instrRange.setFontWeight('bold');
						instrRange.setBackground('#ff9900');
						instrRange.setFontColor('#ffffff');
					} catch (e) {
						Logger.log(`refreshExternalHeaderStyles: could not style INSTRUCTIONS for ${name}: ${e.message}`);
					}
				}

				// Ensure any header cells after the INSTRUCTIONS column have no fill
				try {
					if (instrCol > 0 && instrCol < lastCol) {
						sh.getRange(1, instrCol + 1, 1, lastCol - instrCol).setBackground(null).setFontColor('#000000');
					}
				} catch (e) { /* ignore */ }

				// Try auto-resize only the data columns for readability
				try { if (dataCols > 0) sh.autoResizeColumns(1, dataCols); } catch (e) { /* ignore */ }

				Logger.log(`refreshExternalHeaderStyles: styled ${name} (dataCols: ${dataCols}, instrCol: ${instrCol > 0 ? instrCol : 'n/a'})`);
			} catch (inner) {
				Logger.log(`refreshExternalHeaderStyles: failed for ${name}: ${inner.message}`);
			}
		});

		Logger.log('refreshExternalHeaderStyles: completed');
	} catch (err) {
		Logger.log(`refreshExternalHeaderStyles: error opening external sheet: ${err.message}`);
	}
}

// Internal helper: place an INSTRUCTIONS header after a blank column, then write two-column instructions
function _ensureInstructionsOnSheet_(sheet, instructions) {
 if (!sheet) return;
 const lastCol = sheet.getLastColumn() || 1;
 const headerRow = sheet.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
 let instrColIndex = -1;
 for (let c = 0; c < headerRow.length; c++) {
 if (String(headerRow[c] || '').trim().toUpperCase() === 'INSTRUCTIONS') {
 instrColIndex = c + 1; // 1-based
 break;
 }
 }

 if (instrColIndex === -1) {
 // Find last non-empty header cell
 let lastHeaderCol = 0;
 for (let c = headerRow.length - 1; c >= 0; c--) {
 if (String(headerRow[c] || '').trim() !== '') { lastHeaderCol = c + 1; break; }
 }
 // Leave one blank column and place INSTRUCTIONS header
 instrColIndex = Math.max(1, lastHeaderCol) + 2;
 sheet.getRange(1, instrColIndex, 1, 1).setValue('INSTRUCTIONS');
 }

 // Style the INSTRUCTIONS header
 sheet.getRange(1, instrColIndex, 1, 1)
 .setFontWeight('bold')
 .setBackground('#ff9900')
 .setFontColor('#ffffff');

 // Write instructions (two columns: label + description)
 if (instructions && instructions.length > 0) {
	// Clear existing instruction area first to avoid duplicates
	try { sheet.getRange(2, instrColIndex, Math.max(sheet.getMaxRows() - 1, instructions.length), 2).clearContent(); } catch(e) {}
	sheet.getRange(2, instrColIndex, instructions.length, 2).setValues(instructions);
	const instrRange = sheet.getRange(2, instrColIndex, instructions.length, 2);
	instrRange.setFontSize(10).setVerticalAlignment('top');
 }
}

// Specialized helper: compress Audit Requests instructions into a single wrapped cell and clear extra rows
// (removed single-cell instructions helper; using standard multi-row instructions)

// Build detailed instructions arrays
function _buildExclusionsInstructions_() {
 return [
 ['Config Name:', 'Enter the exact config name (PST01, PST02, NEXT01, etc.) OR leave blank if using "Apply to All Configs"'],
 ['Placement ID:', 'Enter the CM360 Placement ID number (leave blank if using Site Name or Name Fragment)'],
 ['Placement Name:', 'Auto-populated - will update the following day as long as placement ID filled in with an active CM360 placement ID for the appropriate config'],
 ['Site Name:', 'Enter exact site name as it appears in CM360 reporting (alternative to Placement ID)'],
 ['Name Fragment:', 'Enter text fragment that appears in placement names (matches any placement containing this text)'],
 ['Apply to All Configs:', 'TRUE = applies to ALL config teams, FALSE = applies only to specified config'],
 ['Flag Type:', 'Select which flag type to exclude'],
 ['Reason:', 'Brief explanation for the exclusion'],
 ['Added By:', 'Your email'],
 ['Date Added:', 'Date this exclusion was added'],
 ['Active:', 'TRUE to enable, FALSE to disable'],
 ['', ''],
 ['Exclusion Types:', ''],
 ['- Placement ID', 'Excludes specific placement by ID'],
 ['- Site Name', 'Excludes all placements from a specific site'],
 ['- Name Fragment', 'Excludes placements with names containing the fragment'],
 ['', ''],
 ['Flag Types:', ''],
 ['- clicks_greater_than_impressions', 'Excludes clicks > impressions flags'],
 ['- out_of_flight_dates', 'Excludes out of flight date flags'],
 ['- pixel_size_mismatch', 'Excludes pixel mismatch flags'],
 ['- default_ad_serving', 'Excludes default ad serving flags'],
 ['- all_flags', 'Excludes from ALL flag types'],
 ['', ''],
 ['Usage:', 'Add your exclusion rules - fill in only the columns you need']
 ];
}

function _buildThresholdsInstructions_() {
 return [
 ['Config Name:', 'Enter the exact config name (PST01, PST02, NEXT01, etc.)'],
 ['Flag Type:', 'Select which flag type this threshold applies to'],
 ['Min Impressions:', 'Minimum impressions required for this flag to trigger'],
 ['Min Clicks:', 'Minimum clicks required for this flag to trigger'],
 ['Active:', 'TRUE to enable, FALSE to disable this threshold'],
 ['', ''],
 ['LOGIC EXPLANATION:', ''],
 ['How Evaluation Works:', 'The system compares impression vs click volume for each placement.'],
 ['', 'Whichever metric is HIGHER determines the pricing model used:'],
 ['- If Clicks > Impressions', ' Uses Min Clicks'],
 ['- If Impressions > Clicks', ' Uses Min Impressions'],
 ['', ''],
 ['EXAMPLE:', ''],
 ['Placement Data:', 'Impressions: 1,500 | Clicks: 75'],
 ['Result:', 'Since 1,500 impressions > 75 clicks Uses Min Impressions'],
 ['Threshold Check:', 'Compares against "Min Impressions" value only'],
 ['', ''],
 ['Another Example:', 'Impressions: 200 | Clicks: 850'],
 ['Result:', 'Since 850 clicks > 200 impressions Uses Min Clicks'],
 ['Threshold Check:', 'Compares against "Min Clicks" value only'],
 ['', ''],
 ['Flag Types:', ''],
 ['- clicks_greater_than_impressions', 'Flags when clicks exceed impressions'],
 ['- out_of_flight_dates', 'Flags when placement is outside flight dates'],
 ['- pixel_size_mismatch', 'Flags when creative and placement pixels differ'],
 ['- default_ad_serving', 'Flags when ad type contains "default"'],
 ['', ''],
 ['Usage:', 'Modify your configuration thresholds as needed']
 ];
}

function _buildRecipientsInstructions_() {
 return [
 ['Config Name:', 'Enter the exact config name (PST01, PST02, NEXT01, etc.)'],
 ['Primary Recipients:', 'Main email addresses (comma-separated if multiple)'],
 ['CC Recipients:', 'CC email addresses (comma-separated if multiple)'],
 ['Active:', 'TRUE to use these recipients, FALSE to disable'],
 ['Withhold No-Flag Emails:', 'TRUE to skip emails when 0 flags found, FALSE to always send emails'],
 ['Last Updated:', 'Automatically updated when you modify recipients'],
 ['', ''],
 ['Staging Mode Override:', `Currently: ${STAGING_MODE === 'Y' ? 'STAGING (all emails go to admin)' : 'PRODUCTION (uses sheet recipients)'}`],
 ['', ''],
 ['Email Format:', ''],
 ['- Single recipient:', 'user@company.com'],
 ['- Multiple recipients:', 'user1@company.com, user2@company.com'],
 ['- Leave CC blank if not needed', ''],
 ['', ''],
 ['Usage:', 'Modify your configuration recipients as needed']
 ];
}

function _buildRequestsInstructions_() {
	return [
		['How to create a request:', 'Use CM360 Config Helper → Run Config Audit in the helper menu. The system will add a row for you automatically in the next available slot. Do NOT add rows manually.'],
		['', ''],
		['When to use this tab:', 'This sheet is a log/queue of requests — it is not intended for direct user data entry.'],
		['', ''],
		['Troubleshooting:', 'If requests stay PENDING, confirm EXTERNAL_CONFIG_SHEET_ID, permissions, and check the execution logs for errors.'],
		['Security:', 'Only admins should manually change Status values.'],
		['', ''],
		['Usage:', 'Auto-populated and maintained by the helper — leave entries to the system unless instructed by an admin.']
	];
}

/** Populate external config sheet with basic configurations */
function populateExternalConfigWithDefaults() {
 if (!EXTERNAL_CONFIG_SHEET_ID) {
	 const ui = SpreadsheetApp.getUi();
	 ui.alert(
		 'No External Config Sheet',
		 'No external configuration sheet is configured.',
		 ui.ButtonSet.OK
	 );
	 return;
 }

 try {
 Logger.log(`" Populating external config with basic data...`);
 
 const externalSpreadsheet = SpreadsheetApp.openById(EXTERNAL_CONFIG_SHEET_ID);
 Logger.log(`... Successfully opened external spreadsheet: ${EXTERNAL_CONFIG_SHEET_ID}`);
 
 // Populate Recipients sheet
 const recipientsSheet = externalSpreadsheet.getSheetByName(RECIPIENTS_SHEET_NAME);
 if (recipientsSheet) {
 Logger.log(`" Found Recipients sheet: ${RECIPIENTS_SHEET_NAME}`);
 
 const defaultRecipients = [
 ['PST01', ADMIN_EMAIL, '', 'TRUE', 'FALSE', formatDate(new Date(), 'yyyy-MM-dd')],
 ['PST02', 'fvariath@horizonmedia.com', ADMIN_EMAIL, 'TRUE', 'FALSE', formatDate(new Date(), 'yyyy-MM-dd')],
 ['PST03', 'dmaestre@horizonmedia.com', ADMIN_EMAIL, 'TRUE', 'FALSE', formatDate(new Date(), 'yyyy-MM-dd')],
 ['NEXT01', 'bosborne@horizonmedia.com, mmassaroni@horizonmedia.com', ADMIN_EMAIL, 'TRUE', 'FALSE', formatDate(new Date(), 'yyyy-MM-dd')],
 ['NEXT02', 'rschaff@horizonmedia.com, mmassaroni@horizonmedia.com, jwong@horizonmedia.com', ADMIN_EMAIL, 'TRUE', 'FALSE', formatDate(new Date(), 'yyyy-MM-dd')],
 ['NEXT03', 'szeterberg@horizonmedia.com, mmassaroni@horizonmedia.com, jwong@horizonmedia.com', ADMIN_EMAIL, 'TRUE', 'FALSE', formatDate(new Date(), 'yyyy-MM-dd')],
 ['SPTM01', 'spectrum_adops@horizonmedia.com', ADMIN_EMAIL, 'TRUE', 'FALSE', formatDate(new Date(), 'yyyy-MM-dd')],
 ['NFL01', 'NFL_AdOps@horizonmedia.com, sbermolone@horizonmedia.com', ADMIN_EMAIL, 'TRUE', 'FALSE', formatDate(new Date(), 'yyyy-MM-dd')],
 ['ENT01', 'sremick@horizonmedia.com, cali@horizonmedia.com', ADMIN_EMAIL, 'TRUE', 'FALSE', formatDate(new Date(), 'yyyy-MM-dd')]
 ];
 
 // Clear existing data and add defaults
 recipientsSheet.clear();
 const headers = [
 'Config Name',
 'Primary Recipients',
 'CC Recipients',
 'Active',
 'Withhold No-Flag Emails',
 'Last Updated'
 ];
 recipientsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
 recipientsSheet.getRange(2, 1, defaultRecipients.length, 6).setValues(defaultRecipients);
 
 // Format headers
 const headerRange = recipientsSheet.getRange(1, 1, 1, headers.length);
 headerRange.setFontWeight('bold');
 headerRange.setBackground('#4285f4');
 headerRange.setFontColor('#ffffff');
 
 Logger.log(`... Populated Recipients with ${defaultRecipients.length} configs`);
 } else {
 Logger.log(` Recipients sheet not found: ${RECIPIENTS_SHEET_NAME}`);
 }
 
 // Populate Thresholds sheet
 const thresholdsSheet = externalSpreadsheet.getSheetByName(THRESHOLDS_SHEET_NAME);
 if (thresholdsSheet) {
 const flagTypeOptions = [
 'clicks_greater_than_impressions',
 'out_of_flight_dates',
 'pixel_size_mismatch',
 'default_ad_serving'
 ];
 
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
 
 const defaultThresholds = [];
 const configNames = ['PST01', 'PST02', 'PST03', 'NEXT01', 'NEXT02', 'NEXT03', 'SPTM01', 'NFL01', 'ENT01'];
 configNames.forEach(configName => {
 const defaults = defaultValues[configName] || { minImpressions: 0, minClicks: 0 };
 flagTypeOptions.forEach(flagType => {
 defaultThresholds.push([
 configName,
 flagType,
 defaults.minImpressions,
 defaults.minClicks,
 'TRUE'
 ]);
 });
 });
 
 // Clear and populate
 thresholdsSheet.clear();
 const headers = [
 'Config Name',
 'Flag Type',
 'Min Impressions',
 'Min Clicks',
 'Active'
 ];
 thresholdsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
 thresholdsSheet.getRange(2, 1, defaultThresholds.length, 5).setValues(defaultThresholds);
 
 // Format headers
 const headerRange = thresholdsSheet.getRange(1, 1, 1, headers.length);
 headerRange.setFontWeight('bold');
 headerRange.setBackground('#4285f4');
 headerRange.setFontColor('#ffffff');
 
 Logger.log(`... Populated Thresholds with ${defaultThresholds.length} entries`);
 }
 
 const ui = SpreadsheetApp.getUi();
 ui.alert(
 'External Config Populated',
 'External configuration sheet has been populated with default configurations.\n\nRecipients can now use the helper menu to run audits.',
 ui.ButtonSet.OK
 );
 
 Logger.log(`... External config population completed`);
 
 } catch (error) {
 Logger.log(` Error populating external config: ${error.message}`);
 const ui = SpreadsheetApp.getUi();
 ui.alert(
 'Population Error',
 `Failed to populate external config: ${error.message}`,
 ui.ButtonSet.OK
 );
 }
}

/**
 * Debug function to check external config data
 */
function debugExternalConfigData() {
 if (!EXTERNAL_CONFIG_SHEET_ID) {
 const ui = SpreadsheetApp.getUi();
 ui.alert(
 'No External Config Sheet',
 'No external configuration sheet is configured.',
 ui.ButtonSet.OK
 );
 return;
 }

 try {
 Logger.log(`" Debugging external config sheet: ${EXTERNAL_CONFIG_SHEET_ID}`);
 
 const externalSpreadsheet = SpreadsheetApp.openById(EXTERNAL_CONFIG_SHEET_ID);
 Logger.log(`... Successfully opened external spreadsheet: ${externalSpreadsheet.getName()}`);
 
 // Check Recipients sheet
 const recipientsSheet = externalSpreadsheet.getSheetByName(RECIPIENTS_SHEET_NAME);
 let recipientsInfo = '';
 
 if (recipientsSheet) {
 const data = recipientsSheet.getDataRange().getValues();
 Logger.log(`" Recipients sheet found with ${data.length} rows`);
 
 recipientsInfo = `Recipients Sheet (${RECIPIENTS_SHEET_NAME}):\n`;
 recipientsInfo += ` Total rows: ${data.length}\n`;
 recipientsInfo += ` Data rows: ${data.length - 1}\n\n`;
 
 if (data.length > 0) {
 recipientsInfo += `Headers: ${data[0].join(', ')}\n\n`;
 
 // Check each data row
 for (let i = 1; i < Math.min(data.length, 6); i++) {
 const row = data[i];
 const configName = row[0];
 const recipients = row[1];
 const active = row[3];
 recipientsInfo += `Row ${i}: "${configName}" | Recipients: "${recipients}" | Active: "${active}" (type: ${typeof active})\n`;
 }
 
 if (data.length > 6) {
 recipientsInfo += `... and ${data.length - 6} more rows\n`;
 }
 
 // Count active configs with detailed logging
 let activeCount = 0;
 for (let i = 1; i < data.length; i++) {
 const configName = data[i][0];
 const activeValue = data[i][3];
 const activeString = String(activeValue).trim();
 const isActive = (activeString === 'TRUE' || activeString === 'true');
 
 Logger.log(`Row ${i}: Config="${configName}" | Active="${activeValue}" | Type=${typeof activeValue} | String="${activeString}" | IsActive=${isActive}`);
 
 if (configName && isActive) {
 activeCount++;
 }
 }
 recipientsInfo += `\nActive configurations: ${activeCount}`;
 }
 } else {
 recipientsInfo = ` Recipients sheet "${RECIPIENTS_SHEET_NAME}" not found`;
 Logger.log(recipientsInfo);
 }
 
 // Check Thresholds sheet
 const thresholdsSheet = externalSpreadsheet.getSheetByName(THRESHOLDS_SHEET_NAME);
 let thresholdsInfo = '';
 
 if (thresholdsSheet) {
 const data = thresholdsSheet.getDataRange().getValues();
 thresholdsInfo = `\n\nThresholds Sheet (${THRESHOLDS_SHEET_NAME}):\n`;
 thresholdsInfo += ` Total rows: ${data.length}\n`;
 } else {
 thresholdsInfo = `\n\n Thresholds sheet "${THRESHOLDS_SHEET_NAME}" not found`;
 }
 
 // Show all sheet names
 const allSheets = externalSpreadsheet.getSheets().map(sheet => sheet.getName());
 const sheetsInfo = `\n\nAll sheets in external config:\n ${allSheets.join('\n ')}`;
 
 const ui = SpreadsheetApp.getUi();
 const fullReport = recipientsInfo + thresholdsInfo + sheetsInfo;
 
 ui.alert(
 'External Config Debug Report',
 fullReport,
 ui.ButtonSet.OK
 );
 
 Logger.log(`" Debug report:\n${fullReport}`);
 
 } catch (error) {
 Logger.log(` Error debugging external config: ${error.message}`);
 const ui = SpreadsheetApp.getUi();
 ui.alert(
 'Debug Error',
 `Failed to debug external config: ${error.message}`,
 ui.ButtonSet.OK
 );
 }
}

/**
 * Fix case sensitivity issues in external config data
 */
function fixCaseIssuesInExternalConfig() {
 if (!EXTERNAL_CONFIG_SHEET_ID) {
 const ui = SpreadsheetApp.getUi();
 ui.alert(
 'No External Config Sheet',
 'No external configuration sheet is configured.',
 ui.ButtonSet.OK
 );
 return;
 }

 try {
 Logger.log(`" Fixing case issues in external config sheet: ${EXTERNAL_CONFIG_SHEET_ID}`);
 
 const externalSpreadsheet = SpreadsheetApp.openById(EXTERNAL_CONFIG_SHEET_ID);
 
 // Fix Recipients sheet
 const recipientsSheet = externalSpreadsheet.getSheetByName(RECIPIENTS_SHEET_NAME);
 if (recipientsSheet) {
 const data = recipientsSheet.getDataRange().getValues();
 let changes = 0;
 
 for (let i = 1; i < data.length; i++) {
 let changed = false;
 
 // Fix Active column (column D, index 3)
 if (data[i][3] === 'true') {
 data[i][3] = 'TRUE';
 changed = true;
 } else if (data[i][3] === 'false') {
 data[i][3] = 'FALSE';
 changed = true;
 }
 
 // Fix Withhold No-Flag Emails column (column E, index 4)
 if (data[i][4] === 'true') {
 data[i][4] = 'TRUE';
 changed = true;
 } else if (data[i][4] === 'false') {
 data[i][4] = 'FALSE';
 changed = true;
 }
 
 if (changed) changes++;
 }
 
 if (changes > 0) {
 recipientsSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
 Logger.log(`... Fixed ${changes} case issues in Recipients sheet`);
 }
 }
 
 // Fix Thresholds sheet
 const thresholdsSheet = externalSpreadsheet.getSheetByName(THRESHOLDS_SHEET_NAME);
 if (thresholdsSheet) {
 const data = thresholdsSheet.getDataRange().getValues();
 let changes = 0;
 
 for (let i = 1; i < data.length; i++) {
 // Fix Active column (column E, index 4)
 if (data[i][4] === 'true') {
 data[i][4] = 'TRUE';
 changes++;
 } else if (data[i][4] === 'false') {
 data[i][4] = 'FALSE';
 changes++;
 }
 }
 
 if (changes > 0) {
 thresholdsSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
 Logger.log(`... Fixed ${changes} case issues in Thresholds sheet`);
 }
 }
 
 const ui = SpreadsheetApp.getUi();
 ui.alert(
 'Case Issues Fixed',
 'All lowercase "true"/"false" values have been converted to uppercase "TRUE"/"FALSE".\\n\\nThe helper menu should now work correctly.',
 ui.ButtonSet.OK
 );
 
 Logger.log(`... Case fix completed`);
 
 } catch (error) {
 Logger.log(` Error fixing case issues: ${error.message}`);
 const ui = SpreadsheetApp.getUi();
 ui.alert(
 'Fix Error',
 `Failed to fix case issues: ${error.message}`,
 ui.ButtonSet.OK
 );
 }
}

/**
 * Process audit requests from the external config helper menu
 */
function processAuditRequests() {
 if (!EXTERNAL_CONFIG_SHEET_ID) {
 const ui = SpreadsheetApp.getUi();
 ui.alert(
 'No External Config Sheet',
 'No external configuration sheet is configured.',
 ui.ButtonSet.OK
 );
 return;
 }

 try {
		// Ensure main spreadsheet is synced from the external config before processing requests.
		try {
			syncFromExternalConfig();
			Logger.log('Synced main spreadsheet from external config before processing requests.');
		} catch (syncErr) {
			Logger.log(`Failed to sync from external config prior to processing requests: ${syncErr.message}`);
			// Proceed anyway; requests will still be read from the external sheet.
		}
 Logger.log(`" Processing audit requests from external config sheet: ${EXTERNAL_CONFIG_SHEET_ID}`);
 
 const externalSpreadsheet = SpreadsheetApp.openById(EXTERNAL_CONFIG_SHEET_ID);
 const requestsSheet = externalSpreadsheet.getSheetByName('Audit Requests');
 
 if (!requestsSheet) {
 const ui = SpreadsheetApp.getUi();
 ui.alert(
 'No Requests Found',
 'No "Audit Requests" sheet found in the external config spreadsheet.\\n\\nRequests are created when users use the helper menu.',
 ui.ButtonSet.OK
 );
 return;
 }
 
 const data = requestsSheet.getDataRange().getValues();
 if (data.length <= 1) {
 const ui = SpreadsheetApp.getUi();
 ui.alert(
 'No Pending Requests',
 'No audit requests found in the external config sheet.',
 ui.ButtonSet.OK
 );
 return;
 }
 
 const pendingRequests = [];
 const processedRequests = [];
 
 // Find pending requests
 for (let i = 1; i < data.length; i++) {
 const row = data[i];
 const configName = row[0];
 const requestedBy = row[1];
 const requestTime = row[2];
 const status = row[3];
 
 if (status === 'PENDING' && configName) {
 pendingRequests.push({
 row: i + 1,
 configName: configName,
 requestedBy: requestedBy,
 requestTime: requestTime
 });
 }
 }
 
 if (pendingRequests.length === 0) {
 const ui = SpreadsheetApp.getUi();
 ui.alert(
 'No Pending Requests',
 `Found ${data.length - 1} total requests, but none are pending.\\n\\nAll requests have already been processed.`,
 ui.ButtonSet.OK
 );
 return;
 }
 
 Logger.log(`" Found ${pendingRequests.length} pending audit requests`);
 
 // Process each pending request
 for (const request of pendingRequests) {
 Logger.log(`" Processing request for config: ${request.configName}`);
 
 try {
 // Update status to "PROCESSING"
 requestsSheet.getRange(request.row, 4).setValue('PROCESSING');
 requestsSheet.getRange(request.row, 5).setValue(`Started processing at ${new Date().toISOString()}`);
 
 // Run the audit
 let result = false;
 let errorMessage = '';
 
 try {
 // Check if config exists first
 const config = auditConfigs.find(c => c.name === request.configName);
 if (!config) {
 throw new Error(`Configuration "${request.configName}" not found in auditConfigs`);
 }
 
 // Run the audit
 executeAudit(config);
 result = true; // If no exception thrown, assume success
 
 } catch (auditError) {
 result = false;
 errorMessage = auditError.message;
 Logger.log(` Audit execution failed for ${request.configName}: ${errorMessage}`);
 }
 
 // Update status based on result
 if (result) {
 requestsSheet.getRange(request.row, 4).setValue('COMPLETED');
 requestsSheet.getRange(request.row, 5).setValue(`Completed successfully at ${new Date().toISOString()}`);
 processedRequests.push({
 configName: request.configName,
 status: 'COMPLETED',
 requestedBy: request.requestedBy
 });
 } else {
 requestsSheet.getRange(request.row, 4).setValue('FAILED');
 requestsSheet.getRange(request.row, 5).setValue(`Failed at ${new Date().toISOString()}: ${errorMessage || 'Unknown error'}`);
 processedRequests.push({
 configName: request.configName,
 status: 'FAILED',
 requestedBy: request.requestedBy,
 error: errorMessage
 });
 }
 
 Logger.log(`... Completed processing request for config: ${request.configName}`);
 
 } catch (error) {
 Logger.log(` Error processing request for ${request.configName}: ${error.message}`);
 requestsSheet.getRange(request.row, 4).setValue('ERROR');
 requestsSheet.getRange(request.row, 5).setValue(`Error at ${new Date().toISOString()}: ${error.message}`);
 processedRequests.push({
 configName: request.configName,
 status: 'ERROR',
 requestedBy: request.requestedBy,
 error: error.message
 });
 }
 }
 
 // Send summary email to admin
 if (processedRequests.length > 0) {
 const summaryLines = processedRequests.map(req => 
 ` ${req.configName}: ${req.status}${req.error ? ` (${req.error})` : ''} - Requested by: ${req.requestedBy}`
 );
 
 try {
 safeSendEmail({
 to: ADMIN_EMAIL,
 subject: `CM360 Audit Requests Processed - ${processedRequests.length} requests`,
 htmlBody: `
 <h3>Audit Request Processing Summary</h3>
 <p>Processed ${processedRequests.length} audit requests:</p>
 <ul>
 ${summaryLines.map(line => `<li>${line}</li>`).join('')}
 </ul>
 <p>Time: ${new Date().toISOString()}</p>
 <p>Check the external config sheet's "Audit Requests" tab for detailed status.</p>
 `
 }, 'Audit Request Processing Summary');
 } catch (emailError) {
 Logger.log(` Could not send summary email: ${emailError.message}`);
 }
 }
 
 const ui = SpreadsheetApp.getUi();
 const completedCount = processedRequests.filter(r => r.status === 'COMPLETED').length;
 const failedCount = processedRequests.filter(r => r.status === 'FAILED' || r.status === 'ERROR').length;
 
 ui.alert(
 'Audit Requests Processed',
 `Processed ${processedRequests.length} audit requests:\n\n... Completed: ${completedCount}\n Failed: ${failedCount}\n\nCheck your email and the external config sheet for details.`,
 ui.ButtonSet.OK
 );
 
 Logger.log(`... Audit request processing completed. ${completedCount} successful, ${failedCount} failed.`);
 
 } catch (error) {
 Logger.log(` Error processing audit requests: ${error.message}`);
 const ui = SpreadsheetApp.getUi();
 ui.alert(
 'Processing Error',
 `Failed to process audit requests: ${error.message}`,
 ui.ButtonSet.OK
 );
 }
}

// === BATCH TRIGGERS SETUP & MANAGEMENT ===
/**
 * Comprehensive batch management function that checks for missing batch functions,
 * creates them if needed, and installs daily triggers
 */
function setupAndInstallBatchTriggers() {
 const ui = SpreadsheetApp.getUi();
 
 try {
 Logger.log(' Starting batch triggers setup and installation...');
 
 // Step 1: Check current batch status
 const batches = getAuditConfigBatches(BATCH_SIZE);
 const neededCount = batches.length;
 const existingFns = Object.keys(globalThis).filter(k => /^runDailyAuditsBatch\d+$/.test(k));
 const definedIndexes = new Set(existingFns.map(fn => Number(fn.match(/\d+$/)[0])));
 
 let missingFunctions = [];
 for (let i = 1; i <= neededCount; i++) {
 if (!definedIndexes.has(i)) {
 missingFunctions.push(`runDailyAuditsBatch${i}`);
 }
 }
 
 // Step 2: Report status and get user confirmation
 let statusMessage = ` Batch Status Analysis:\n\n`;
 statusMessage += `- Total configs: ${auditConfigs.length}\n`;
 statusMessage += `- Batch size: ${BATCH_SIZE}\n`;
 statusMessage += `- Batches needed: ${neededCount}\n`;
 statusMessage += `- Existing batch functions: ${existingFns.length}\n`;
 statusMessage += `- Missing batch functions: ${missingFunctions.length}\n\n`;
 
 if (missingFunctions.length > 0) {
	statusMessage += `❌ Missing functions:\n${missingFunctions.map(fn => `- ${fn}`).join('\n')}\n\n`;
	statusMessage += `⚠️ Missing batch functions must be manually added to the script.\n\n`;
 statusMessage += `This tool will:\n`;
 statusMessage += `1. Show you the missing function code to copy\n`;
 statusMessage += `2. Install/update daily triggers for existing functions\n\n`;
 statusMessage += `Continue?`;
 } else {
	statusMessage += `✅ All batch functions exist!\n\n`;
 statusMessage += `This tool will install/update daily triggers.\n\n`;
 statusMessage += `Continue?`;
 }
 
 const response = ui.alert(
 'Setup Batch Triggers',
 statusMessage,
 ui.ButtonSet.YES_NO
 );
 
 if (response !== ui.Button.YES) {
 return;
 }
 
 // Step 3: Show missing function code if needed
 if (missingFunctions.length > 0) {
 const missingCode = generateMissingBatchStubs();
 
	const codeMessage = `❌ MISSING BATCH FUNCTIONS:\n\nCopy this code and paste it at the end of your script:\n\n${missingCode}\n\nAfter adding the functions, run this tool again to install triggers.`;
 
 ui.alert(
 'Missing Functions Code',
 codeMessage,
 ui.ButtonSet.OK
 );
 
	Logger.log('❌ Missing batch functions code:');
 Logger.log(missingCode);
 return;
 }
 
 // Step 4: Install triggers for existing functions
 Logger.log(' Installing daily triggers...');
 const triggerResults = installDailyAuditTriggers();
 
 // Step 5: Report final results
	let finalMessage = `✅ Batch Triggers Setup Complete!\n\n`;
 finalMessage += ` Summary:\n`;
 finalMessage += `- Batch functions: ${neededCount}/${neededCount} available\n`;
 finalMessage += `- Triggers installed: ${triggerResults.length}\n`;
 finalMessage += `- Configs per batch: ${BATCH_SIZE}\n\n`;
 finalMessage += ` Batches:\n`;
 
 batches.forEach((batch, index) => {
 finalMessage += `- Batch ${index + 1}: ${batch.map(c => c.name).join(', ')}\n`;
 });
 
	finalMessage += `\n✅ Daily triggers are now active!`;
 
 ui.alert(
 'Setup Complete',
 finalMessage,
 ui.ButtonSet.OK
 );
 
	Logger.log('✅ Batch triggers setup completed successfully');
 Logger.log(` Installed ${triggerResults.length} triggers for ${neededCount} batches`);
 
 } catch (error) {
	Logger.log(`❌ Error in setupAndInstallBatchTriggers: ${error.message}`);
 ui.alert(
 'Setup Error',
 `Failed to setup batch triggers:\n\n${error.message}`,
 ui.ButtonSet.OK
 );
 }
}








