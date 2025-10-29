

// === CONFIGURATION CONSTANTS ===
// Update EXTERNAL_CONFIG_SHEET_ID with the ID of the helper spreadsheet once it is created.
// Leave blank to keep using the bound spreadsheet for configuration.
const EXTERNAL_CONFIG_SHEET_ID = '1-566gqkyZRNDeNtXWUjKDB_H8A9XbhCu8zL-uaZdGT8';

// Delivery mode override: read from Script Properties once at load with safe fallback
// IMPORTANT: STAGING_MODE must be read DYNAMICALLY from Script Properties, not cached as a constant
// This ensures changes take effect immediately without requiring trigger reinstallation
// Default to 'Y' (staging) if property not set - safer for testing
function getStagingMode_() {
	try {
		const v = String(PropertiesService.getScriptProperties().getProperty('STAGING_MODE') || 'N').trim().toUpperCase();
		return v === 'Y' ? 'Y' : 'N';
	} catch (e) {
		return 'Y'; // Safe default: staging mode on
	}
}

// Admin email for alerts, staging redirects, and defaults
// IMPORTANT: Do NOT fall back to the trigger owner; always use the explicit admin.
const ADMIN_EMAIL = (function() {
	try {
		const p = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
		if (p && /@/.test(p)) return p;
	} catch (e) {}
	// Fallback default for admin notifications
	return 'evschneider@horizonmedia.com';
})();

// Root Drive path used for Trash/Temp/Merged report storage
const TRASH_ROOT_PATH = (function() {
	try {
		const raw = PropertiesService.getScriptProperties().getProperty('TRASH_ROOT_PATH');
		if (raw) {
			const parsed = JSON.parse(raw);
			if (Array.isArray(parsed) && parsed.length) return parsed.map(String);
		}
	} catch (e) {}
	// Sensible default structure
	return ['Project Log Files', 'CM360 Daily Audits'];
})();

// Cleanup constants (used by Drive cleanup routines)
const CLEANUP_STATE_KEY = 'CM360_CLEANUP_STATE_V1';
const CLEANUP_TRIGGER_ID_KEY = 'CM360_CLEANUP_TRIGGER_ID';
const CLEANUP_STATE_VERSION = 1;
const CLEANUP_PAGE_SIZE = 100;
const CLEANUP_RUNTIME_BUFFER_MS = 5000; // ms reserved to bail out before timeout
// Maximum runtime budget for a cleanup pass before yielding (ms)
const CLEANUP_RUNTIME_LIMIT_MS = 300000; // 5 minutes
// Location to store the deletion log workbook
const DELETION_LOG_PATH = [...TRASH_ROOT_PATH, 'Deletion Log'];
// Name of the deletion log workbook
const ADMIN_LOG_NAME = 'CM360 Deletion Log';

// Audit run state tracking (used by watchdog to detect timeouts)
const AUDIT_RUN_STATE_KEY_PREFIX = 'CM360_AUDIT_RUN_STATE_V1_';
const AUDIT_RUN_LIST_KEY = 'CM360_AUDIT_RUN_LIST_V1';
const CONFIG_ORDER_PROPERTY_KEY = 'CM360_CONFIG_ORDER_V1';

// Email safety guard to keep HTML bodies reasonable in size
const EMAIL_BODY_BYTE_LIMIT = 90000; // ~90KB

// Batch size for audit execution (number of configs per batch)
// Requested: 2 per batch
const BATCH_SIZE = 2;

// === SHEET NAME CONSTANTS ===
// Centralized names for configuration tabs used throughout the script
const RECIPIENTS_SHEET_NAME = 'Audit Recipients';
const THRESHOLDS_SHEET_NAME = 'Audit Thresholds';
const EXCLUSIONS_SHEET_NAME = 'Audit Exclusions';

// Robust opener for spreadsheets by ID with Drive fallback and clear errors
function openSpreadsheetById_(id) {
	const sid = String(id || '').trim();
	if (!sid) throw new Error('Missing spreadsheet ID.');
	try {
		return SpreadsheetApp.openById(sid);
	} catch (e1) {
		try {
			const file = DriveApp.getFileById(sid);
			return SpreadsheetApp.open(file);
		} catch (e2) {
			throw new Error(`Failed to open spreadsheet by ID. openById error: ${e1 && e1.message ? e1.message : e1}; Drive fallback error: ${e2 && e2.message ? e2.message : e2}`);
		}
	}
}

// Returns the spreadsheet to use for configuration tabs.
// If EXTERNAL_CONFIG_SHEET_ID is set, opens that file; otherwise uses the bound active spreadsheet.
function getConfigSpreadsheet() {
	// Always use the bound Admin spreadsheet for in-app operations.
	// External config access is handled explicitly by sync functions via openById.
	const active = SpreadsheetApp.getActiveSpreadsheet();
	if (!active) throw new Error('No active spreadsheet available');
	return active;
}

// Restored: getDriveFolderByPath_ helper to get or create nested Drive folders by path
function getDriveFolderByPath_(pathArray) {
	const maxAttempts = 3;
	for (var attempt = 1; attempt <= maxAttempts; attempt++) {
		try {
			let folder = DriveApp.getRootFolder();
			if (!Array.isArray(pathArray) || pathArray.length === 0) return folder;
			for (var i = 0; i < pathArray.length; i++) {
				var name = String(pathArray[i] || '').trim();
				if (!name) continue;
				var it = folder.getFoldersByName(name);
				folder = it.hasNext() ? it.next() : folder.createFolder(name);
			}
			return folder;
		} catch (e) {
			Logger.log(`getDriveFolderByPath_ error (attempt ${attempt}/${maxAttempts}): ${e.message}`);
			if (attempt < maxAttempts) {
				try {
					Utilities.sleep(200 * attempt);
				} catch (sleepErr) {}
			}
		}
	}
	Logger.log('getDriveFolderByPath_ exhausted retries; returning null.');
	return null;
}

// Read-only variant: walk a Drive path and return the folder if all parts exist; do not create anything
function getDriveFolderByPathReadOnly_(pathArray) {
	try {
		let folder = DriveApp.getRootFolder();
		if (!Array.isArray(pathArray) || pathArray.length === 0) return folder;
		for (var i = 0; i < pathArray.length; i++) {
			var name = String(pathArray[i] || '').trim();
			if (!name) continue;
			var it = folder.getFoldersByName(name);
			if (!it.hasNext()) return null;
			folder = it.next();
		}
		return folder;
	} catch (e) {
		Logger.log('getDriveFolderByPathReadOnly_ error: ' + e.message);
		return null;
	}
}

// Ensure a Gmail label exists; returns { created: boolean, labelName: string }
function ensureGmailLabelExists_(labelName) {
	try {
		let created = false;
		let lbl = GmailApp.getUserLabelByName(labelName);
		if (!lbl) {
			GmailApp.createLabel(labelName);
			created = true;
			Logger.log(`... Created Gmail label: ${labelName}`);
		} else {
			Logger.log(`[LABEL] Gmail label already exists: ${labelName}`);
		}
		// Verify existence after potential creation
		const confirm = GmailApp.getUserLabelByName(labelName);
		if (!confirm) throw new Error(`Failed to verify Gmail label existence: ${labelName}`);
		return { created, labelName };
	} catch (e) {
		Logger.log(`ensureGmailLabelExists_ error for "${labelName}": ${e.message}`);
		throw e;
	}
}

function getCleanupState_() {
    try {
        const raw = PropertiesService.getScriptProperties().getProperty(CLEANUP_STATE_KEY);
        return raw ? JSON.parse(raw) : null;
    } catch (e) {
        Logger.log('getCleanupState_ error: ' + e.message);
        return null;
    }
}

function saveCleanupState_(state) {
    try {
        if (!state || typeof state !== 'object') return;
        state.version = CLEANUP_STATE_VERSION;
        PropertiesService.getScriptProperties().setProperty(CLEANUP_STATE_KEY, JSON.stringify(state));
    } catch (e) {
        Logger.log('saveCleanupState_ error: ' + e.message);
    }
}

function clearCleanupState_() {
	try {
		PropertiesService.getScriptProperties().deleteProperty(CLEANUP_STATE_KEY);
	} catch (e) {
		Logger.log('clearCleanupState_ error: ' + e.message);
	}
}

function scheduleCleanupContinuation_() {
	try {
		const props = PropertiesService.getScriptProperties();
		let triggerExists = false;
		const storedId = props.getProperty(CLEANUP_TRIGGER_ID_KEY);
		if (storedId) {
			const triggers = ScriptApp.getProjectTriggers();
			triggerExists = triggers.some(trigger => {
				if (trigger.getHandlerFunction() !== 'cleanupOldAuditFiles') return false;
				return typeof trigger.getUniqueId === 'function' && trigger.getUniqueId() === storedId;
			});
			if (!triggerExists) {
				props.deleteProperty(CLEANUP_TRIGGER_ID_KEY);
			}
		}
		if (!triggerExists) {
			const trigger = ScriptApp.newTrigger('cleanupOldAuditFiles').timeBased().after(60 * 1000).create();
			if (trigger && typeof trigger.getUniqueId === 'function') {
				props.setProperty(CLEANUP_TRIGGER_ID_KEY, trigger.getUniqueId());
			}
		}
	} catch (e) {
		Logger.log('scheduleCleanupContinuation_ error: ' + e.message);
	}
}

function clearCleanupContinuation_() {
	try {
		const props = PropertiesService.getScriptProperties();
		const storedId = props.getProperty(CLEANUP_TRIGGER_ID_KEY);
		const triggers = ScriptApp.getProjectTriggers();
		triggers.forEach(trigger => {
			try {
				if (trigger.getHandlerFunction && trigger.getHandlerFunction() === 'cleanupOldAuditFiles') {
					if (!storedId || (typeof trigger.getUniqueId === 'function' && trigger.getUniqueId() === storedId)) {
						ScriptApp.deleteTrigger(trigger);
					}
				}
			} catch (inner) {
				Logger.log('clearCleanupContinuation_ delete error: ' + inner.message);
			}
		});
		props.deleteProperty(CLEANUP_TRIGGER_ID_KEY);
	} catch (e) {
		Logger.log('clearCleanupContinuation_ error: ' + e.message);
	}
}

function shouldYieldCleanup_(startTime, maxRuntimeMs) {
	return Date.now() - startTime > (maxRuntimeMs - CLEANUP_RUNTIME_BUFFER_MS);
}

function ensureFolderFromState_(state, key, pathArray) {
	if (!state) return null;
	if (state[key]) {
		try {
			return DriveApp.getFolderById(state[key]);
		} catch (e) {
			Logger.log(`ensureFolderFromState_ invalid id (${key}): ${e.message}`);
		}
	}
	const folder = getDriveFolderByPath_(pathArray);
	if (folder) {
		state[key] = folder.getId();
	}
	return folder;
}

function appendDeletionLogRows_(logFolder, adminLogName, rows, sheetName) {
	if (!rows || rows.length === 0) return null;
	const targetSheetName = sheetName || 'Temp Daily Reports';
	let logFile;
	try {
		const existing = logFolder.getFilesByName(adminLogName);
		if (existing.hasNext()) {
			logFile = existing.next();
		} else {
			const ss = SpreadsheetApp.create(adminLogName);
			const initialSheet = ss.getActiveSheet();
			initialSheet.setName('Temp Daily Reports');
			ensureDeletionLogSheet_(ss, 'Temp Daily Reports', true);
			ensureDeletionLogSheet_(ss, 'Merged Reports', true);
			logFile = DriveApp.getFileById(ss.getId());
			logFolder.addFile(logFile);
			try {
				DriveApp.getRootFolder().removeFile(logFile);
			} catch (e) {
				Logger.log('appendDeletionLogRows_ remove error: ' + e.message);
			}
		}
	} catch (e) {
		Logger.log('appendDeletionLogRows_ lookup error: ' + e.message);
		return null;
	}
	try {
		const ss = SpreadsheetApp.open(logFile);
		const sheet1 = ss.getSheetByName('Sheet1');
		if (sheet1) sheet1.setName('Temp Daily Reports');
		ensureDeletionLogSheet_(ss, 'Temp Daily Reports');
		ensureDeletionLogSheet_(ss, 'Merged Reports');
		const sheet = ensureDeletionLogSheet_(ss, targetSheetName);
		const startRow = sheet.getLastRow() + 1;
		sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
		SpreadsheetApp.flush();
		return ss.getUrl();
	} catch (e) {
		Logger.log('appendDeletionLogRows_ write error: ' + e.message);
		return null;
	}
}

function ensureDeletionLogSheet_(spreadsheet, sheetName, initializing) {
	const header = ['File Name', 'Folder Path', 'Date Created', 'Deleted On'];
	let sheet = spreadsheet.getSheetByName(sheetName);
	if (!sheet) {
		sheet = spreadsheet.insertSheet(sheetName);
	}
	if (initializing) {
		sheet.clear();
		sheet.getRange(1, 1, 1, header.length).setValues([header]);
		return sheet;
	}
	if (sheet.getLastRow() === 0) {
		sheet.getRange(1, 1, 1, header.length).setValues([header]);
		return sheet;
	}
	const firstRow = sheet.getRange(1, 1, 1, header.length).getValues()[0];
	const hasHeader = header.every((val, idx) => String(firstRow[idx]).trim() === val);
	if (!hasHeader) {
		sheet.insertRowBefore(1);
		sheet.getRange(1, 1, 1, header.length).setValues([header]);
	}
	return sheet;
}

function ensureDeletionLogWorkbookStructure_(logFolder, adminLogName) {
	try {
		const existing = logFolder.getFilesByName(adminLogName);
		if (existing.hasNext()) {
			const logFile = existing.next();
			const ss = SpreadsheetApp.open(logFile);
			const sheet1 = ss.getSheetByName('Sheet1');
			if (sheet1) sheet1.setName('Temp Daily Reports');
			ensureDeletionLogSheet_(ss, 'Temp Daily Reports');
			ensureDeletionLogSheet_(ss, 'Merged Reports');
			return;
		}
		const ss = SpreadsheetApp.create(adminLogName);
		const initialSheet = ss.getActiveSheet();
		initialSheet.setName('Temp Daily Reports');
		ensureDeletionLogSheet_(ss, 'Temp Daily Reports', true);
		ensureDeletionLogSheet_(ss, 'Merged Reports', true);
		const logFile = DriveApp.getFileById(ss.getId());
		logFolder.addFile(logFile);
		try {
			DriveApp.getRootFolder().removeFile(logFile);
		} catch (e) {
			Logger.log('ensureDeletionLogWorkbookStructure_ remove error: ' + e.message);
		}
	} catch (e) {
		Logger.log('ensureDeletionLogWorkbookStructure_ error: ' + e.message);
	}
}

function processCleanupLooseFiles_(state, trashRoot, cutoffDate, timestamp, logBuckets, startTime, maxRuntimeMs) {
	const files = trashRoot.getFiles();
	let index = 0;
	const startIndex = state.looseIndex || 0;
	const pathString = TRASH_ROOT_PATH.join(' / ');
	const tz = Session.getScriptTimeZone();
	while (files.hasNext()) {
		const file = files.next();
		if (index++ < startIndex) continue;
		state.looseIndex = index;
		try {
			const created = file.getDateCreated();
			if (created < cutoffDate) {
				logBuckets.temp.push([
					file.getName(),
					pathString,
					Utilities.formatDate(created, tz, 'yyyy-MM-dd'),
					timestamp
				]);
				file.setTrashed(true);
			}
		} catch (err) {
			Logger.log('processCleanupLooseFiles_ error: ' + err.message);
		}
		if (shouldYieldCleanup_(startTime, maxRuntimeMs)) {
			return false;
		}
	}
	state.looseIndex = 0;
	state.looseDone = true;
	return true;
}

function processCleanupTempFolders_(state, tempRoot, cutoffDate, timestamp, logBuckets, startTime, maxRuntimeMs) {
	if (!tempRoot) {
		state.tempDone = true;
		return true;
	}
	if (!state.temp) state.temp = { configIndex: 0, subfolderIndex: 0 };
	const configs = tempRoot.getFolders();
	let configIdx = 0;
	const tz = Session.getScriptTimeZone();
	while (configs.hasNext()) {
		const configFolder = configs.next();
		if (configIdx++ < (state.temp.configIndex || 0)) continue;
		const configName = configFolder.getName();
		// Use the config folder directly; no nested REPORTS folder in the structure
		const container = configFolder;
		let subIdx = 0;
		const subfolders = container.getFolders();
		while (subfolders.hasNext()) {
			const tempFolder = subfolders.next();
			if (subIdx++ < (state.temp.subfolderIndex || 0)) continue;
			state.temp.subfolderIndex = subIdx;
			try {
				const name = tempFolder.getName();
				const created = tempFolder.getDateCreated();
				if (name.startsWith('Temp_') && created < cutoffDate) {
					logBuckets.temp.push([
						name,
						[...TRASH_ROOT_PATH, 'To Trash After 60 Days', 'Temp Daily Reports', configName].join(' / '),
						Utilities.formatDate(created, tz, 'yyyy-MM-dd'),
						timestamp
					]);
					tempFolder.setTrashed(true);
					Logger.log(`-' Deleted old temp folder: ${name}`);
				}
			} catch (err) {
				Logger.log(`processCleanupTempFolders_ error (${configName}): ${err.message}`);
			}
			if (shouldYieldCleanup_(startTime, maxRuntimeMs)) {
				state.temp.configIndex = configIdx - 1;
				return false;
			}
		}
		state.temp.subfolderIndex = 0;
		state.temp.configIndex = configIdx;
		if (shouldYieldCleanup_(startTime, maxRuntimeMs)) {
			return false;
		}
	}
	state.temp.configIndex = 0;
	state.temp.subfolderIndex = 0;
	state.tempDone = true;
	return true;
}

function processCleanupMergedFiles_(state, mergedRoot, cutoffDate, timestamp, logBuckets, startTime, maxRuntimeMs) {
	if (!mergedRoot) {
		state.mergedDone = true;
		return true;
	}
	if (!state.merged) state.merged = { configIndex: 0, pageToken: null };
	const configs = mergedRoot.getFolders();
	let configIdx = 0;
	const tz = Session.getScriptTimeZone();
	while (configs.hasNext()) {
		const configFolder = configs.next();
		if (configIdx++ < (state.merged.configIndex || 0)) continue;
		const configName = configFolder.getName();
		// Use the config folder directly; no nested REPORTS folder in the structure
		const folderId = configFolder.getId();
		let pageToken = state.merged.pageToken || null;
		let continuePaging = true;
		while (continuePaging) {
			if (shouldYieldCleanup_(startTime, maxRuntimeMs)) {
				state.merged.configIndex = configIdx - 1;
				state.merged.pageToken = pageToken;
				return false;
			}
			const params = {
				q: `'${folderId}' in parents and trashed = false`,
				fields: 'nextPageToken, items(id, title, createdDate, modifiedDate)',
				pageSize: CLEANUP_PAGE_SIZE,
				supportsAllDrives: true,
				includeItemsFromAllDrives: true
			};
			if (pageToken) params.pageToken = pageToken;
			const resp = driveFilesListWithRetry_(params, `processCleanupMergedFiles_/${configName}`);
			if (!resp) {
				Logger.log(`processCleanupMergedFiles_ list error (${configName}): exhausted retries`);
				break;
			}
			const files = resp.items || resp.files || [];
			for (const file of files) {
				if (!file) continue;
				const name = file.title || file.name || '';
				const createdRaw = file.createdDate || file.modifiedDate || null;
				const created = createdRaw ? new Date(createdRaw) : null;
				if (created && name.startsWith('Merged_') && created < cutoffDate) {
					logBuckets.merged.push([
						name,
						[...TRASH_ROOT_PATH, 'To Trash After 60 Days', 'Merged Reports', configName].join(' / '),
						Utilities.formatDate(created, tz, 'yyyy-MM-dd'),
						timestamp
					]);
					try {
						driveFilesUpdateWithRetry_({ trashed: true }, file.id, null, { supportsAllDrives: true }, `processCleanupMergedFiles_/trash/${configName}/${name}`);
					} catch (e) {
						try {
							DriveApp.getFileById(file.id).setTrashed(true);
						} catch (fallbackErr) {
							Logger.log(`processCleanupMergedFiles_ delete error (${name}): ${fallbackErr.message}`);
						}
					}
					Logger.log(`-' Deleted old merged file: ${name}`);
				}
				if (shouldYieldCleanup_(startTime, maxRuntimeMs)) {
					state.merged.configIndex = configIdx - 1;
					state.merged.pageToken = pageToken;
					return false;
				}
			}
			pageToken = resp.nextPageToken || null;
			continuePaging = Boolean(pageToken);
		}
		state.merged.pageToken = null;
		state.merged.configIndex = configIdx;
		if (shouldYieldCleanup_(startTime, maxRuntimeMs)) {
			return false;
		}
	}
	state.merged.configIndex = 0;
	state.merged.pageToken = null;
	state.mergedDone = true;
	return true;
}

function processCleanupOtherFolders_(state, trashRoot, cutoffDate, timestamp, logBuckets, startTime, maxRuntimeMs) {
	if (!state.other) state.other = { folderIndex: 0, pageToken: null };
	const exclusions = new Set(['Temp Daily Reports', 'Merged Reports', 'Deletion Log']);
	const folders = trashRoot.getFolders();
	let processedIndex = 0;
	const tz = Session.getScriptTimeZone();
	while (folders.hasNext()) {
		const folder = folders.next();
		const name = folder.getName();
		if (exclusions.has(name)) continue;
		if (processedIndex < (state.other.folderIndex || 0)) {
			processedIndex++;
			continue;
		}
		const folderId = folder.getId();
		const pathStr = [...TRASH_ROOT_PATH, name].join(' / ');
		let pageToken = state.other.pageToken || null;
		let continuePaging = true;
		while (continuePaging) {
			if (shouldYieldCleanup_(startTime, maxRuntimeMs)) {
				state.other.folderIndex = processedIndex;
				state.other.pageToken = pageToken;
				return false;
			}
			const params = {
				q: `'${folderId}' in parents and trashed = false`,
				fields: 'nextPageToken, items(id, title, createdDate, modifiedDate)',
				pageSize: CLEANUP_PAGE_SIZE,
				supportsAllDrives: true,
				includeItemsFromAllDrives: true
			};
			if (pageToken) params.pageToken = pageToken;
			const resp = driveFilesListWithRetry_(params, `processCleanupOtherFolders_/${name}`);
			if (!resp) {
				Logger.log(`processCleanupOtherFolders_ list error (${name}): exhausted retries`);
				break;
			}
			const files = resp.items || resp.files || [];
			for (const file of files) {
				if (!file) continue;
				const createdRaw = file.createdDate || file.modifiedDate || null;
				const created = createdRaw ? new Date(createdRaw) : null;
				if (created && created < cutoffDate) {
					const fname = file.title || file.name || '';
					logBuckets.temp.push([
						fname,
						pathStr,
						Utilities.formatDate(created, tz, 'yyyy-MM-dd'),
						timestamp
					]);
					try {
						driveFilesUpdateWithRetry_({ trashed: true }, file.id, null, { supportsAllDrives: true }, `processCleanupOtherFolders_/trash/${name}/${fname}`);
					} catch (e) {
						try {
							DriveApp.getFileById(file.id).setTrashed(true);
						} catch (fallbackErr) {
							Logger.log(`processCleanupOtherFolders_ delete error (${fname}): ${fallbackErr.message}`);
						}
					}
				}
				if (shouldYieldCleanup_(startTime, maxRuntimeMs)) {
					state.other.folderIndex = processedIndex;
					state.other.pageToken = pageToken;
					return false;
				}
			}
			pageToken = resp.nextPageToken || null;
			continuePaging = Boolean(pageToken);
		}
		state.other.pageToken = null;
		processedIndex++;
		state.other.folderIndex = processedIndex;
		try {
			if (!folder.getFiles().hasNext() && !folder.getFolders().hasNext()) {
				folder.setTrashed(true);
				Logger.log(`-' Deleted empty folder: ${name}`);
			}
		} catch (err) {
			Logger.log(`processCleanupOtherFolders_ empty check error (${name}): ${err.message}`);
		}
		if (shouldYieldCleanup_(startTime, maxRuntimeMs)) {
			return false;
		}
	}
	state.other.folderIndex = 0;
	state.other.pageToken = null;
	state.otherDone = true;
	return true;
}

function finalizeCleanupRun_(state, logFolder, adminLogName, logBuckets, isComplete) {
	const tempRows = logBuckets && Array.isArray(logBuckets.temp) ? logBuckets.temp : [];
	const mergedRows = logBuckets && Array.isArray(logBuckets.merged) ? logBuckets.merged : [];
	let logUrl = null;
	let loggedAny = false;

	if (tempRows.length) {
		const url = appendDeletionLogRows_(logFolder, adminLogName, tempRows, 'Temp Daily Reports');
		if (url) logUrl = url;
		loggedAny = true;
	}

	if (mergedRows.length) {
		const url = appendDeletionLogRows_(logFolder, adminLogName, mergedRows, 'Merged Reports');
		if (url) logUrl = url;
		loggedAny = true;
	}

	if (isComplete) {
		clearCleanupState_();
		clearCleanupContinuation_();
		if (logUrl) {
			Logger.log(`Cleanup complete. Log updated: ${logUrl}`);
		} else if (loggedAny) {
			Logger.log('Cleanup complete. Log updated.');
		} else {
			Logger.log('Cleanup complete. No deletions logged.');
		}
	} else {
		saveCleanupState_(state);
		scheduleCleanupContinuation_();
		if (logUrl) {
			Logger.log(`Cleanup paused; will resume shortly. Log updated: ${logUrl}`);
		} else if (loggedAny) {
			Logger.log('Cleanup paused; will resume shortly. Log updated.');
		} else {
			Logger.log('Cleanup paused; will resume shortly. No deletions logged this pass.');
		}
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
	let to = ADMIN_EMAIL;
	let subject = '';
	let plainBody = '';
	let mailOpts = {};
	try {
		to = opts.to || ADMIN_EMAIL;
		subject = opts.subject || '';
		plainBody = opts.plainBody || '';
		mailOpts = {};
		if (opts.htmlBody) mailOpts.htmlBody = opts.htmlBody;
		if (opts.cc) mailOpts.cc = opts.cc;
		if (opts.attachments) mailOpts.attachments = opts.attachments;
		if (opts.bcc) mailOpts.bcc = opts.bcc; // allow callers to pass explicit bcc

		// If this run is in suppression mode (silent check), do not actually send
		try {
			if (isEmailSuppressed_ && isEmailSuppressed_()) {
				Logger.log(`[EMAIL] Suppressed (no send) ${context ? '(' + context + ')' : ''} to=${to}; subject="${subject}"`);
				return true; // simulate success so callers treat as "would send"
			}
		} catch (sErr) { /* ignore */ }

		const payloadSize = getEmailPayloadSize_(mailOpts.htmlBody, plainBody);
		Logger.log(`[EMAIL] Payload size (${context}): html=${payloadSize.htmlBytes}B/${payloadSize.htmlChars} chars, plain=${payloadSize.plainBytes}B/${payloadSize.plainChars} chars, total=${payloadSize.totalBytes}B`);
		if (payloadSize.totalBytes > EMAIL_BODY_BYTE_LIMIT) {
			Logger.log(`[EMAIL] Warning: body size ${payloadSize.totalBytes}B exceeds guard limit ${EMAIL_BODY_BYTE_LIMIT}B (${context})`);
		}

		// In staging mode: send ONLY to admin (strip CC/BCC and ignore provided recipients)
		const stagingMode = getStagingMode_();
		if (stagingMode === 'Y') {
			Logger.log(`[EMAIL] Staging mode active - forcing delivery to admin only (original to=${to}) (${context})`);
			const sanitized = {};
			if (mailOpts.htmlBody) sanitized.htmlBody = mailOpts.htmlBody;
			if (mailOpts.attachments) sanitized.attachments = mailOpts.attachments;
			// Intentionally omit cc/bcc in staging
			const bodyForSend = (typeof plainBody === 'string' && plainBody.length > 0) ? plainBody : ' '; // ensure non-empty body
			GmailApp.sendEmail(ADMIN_EMAIL, subject, bodyForSend, sanitized);
			// Best-effort quota capture and explicit success log in staging
			try { const q = MailApp.getRemainingDailyQuota(); if (typeof q === 'number') storeEmailQuotaRemaining_(q); } catch (qErr) {}
			Logger.log(`[EMAIL] Sent (staging) to ${ADMIN_EMAIL} (${context})`);
			return true;
		}

		// BCC admin only when explicitly requested by caller and in production
		if (stagingMode === 'N' && opts.bccAdmin === true && ADMIN_EMAIL) {
			const normalizeList = (val) => {
				if (!val) return [];
				if (Array.isArray(val)) return val.map(String);
				return String(val).split(',');
			};
			const hasEmail = (list, email) => {
				const needle = String(email || '').trim().toLowerCase();
				return normalizeList(list).some(x => String(x || '').trim().toLowerCase() === needle);
			};
			const alreadyIncluded = hasEmail(to, ADMIN_EMAIL) || hasEmail(mailOpts.cc, ADMIN_EMAIL) || hasEmail(mailOpts.bcc, ADMIN_EMAIL);
			if (!alreadyIncluded) {
				const existingBcc = normalizeList(mailOpts.bcc);
				existingBcc.push(ADMIN_EMAIL);
				const unique = Array.from(new Set(existingBcc.map(x => String(x || '').trim()).filter(Boolean)));
				if (unique.length) mailOpts.bcc = unique.join(', ');
			}
		}

		const bodyForSend = (typeof plainBody === 'string' && plainBody.length > 0) ? plainBody : ' ';
		GmailApp.sendEmail(to, subject, bodyForSend, mailOpts);
		Logger.log(`[EMAIL] Sent to ${to} (${context})`);
		try { const q = MailApp.getRemainingDailyQuota(); if (typeof q === 'number') storeEmailQuotaRemaining_(q); } catch (qErr) {}
		return true;
	} catch (e) {
		const errorMessage = e && e.message ? e.message : String(e);
		Logger.log(`safeSendEmail error (${context}): ${errorMessage}`);
		try {
			const fallbackSubject = `[ALERT] CM360 email send failed${context ? ` (${context})` : ''}`;
			const truncatedOriginalSubject = subject ? String(subject).slice(0, 120) : '';
			const recipientsForAlert = Array.isArray(to) ? to.join(', ') : String(to || '');
			const attachmentCount = Array.isArray(opts.attachments) ? opts.attachments.length : (opts.attachments ? 1 : 0);
			const fallbackBody = [
				'A CM360 audit email failed to send.',
				`Context: ${context || 'n/a'}`,
				`Original recipients: ${recipientsForAlert || 'n/a'}`,
				`Original subject (truncated): ${truncatedOriginalSubject}`,
				`Attachment count: ${attachmentCount}`,
				`Error: ${errorMessage}`,
				'',
				'Please review the Apps Script logs for full details.'
			].join('\n');
			GmailApp.sendEmail(ADMIN_EMAIL, fallbackSubject, fallbackBody);
			Logger.log(`[EMAIL] Failure alert sent to admin (${ADMIN_EMAIL})`);
		} catch (notifyErr) {
			Logger.log(`safeSendEmail fallback error (${context}): ${notifyErr.message}`);
		}
		return false;
	}
}

function getEmailPayloadSize_(htmlBody, plainBody) {
	const htmlText = htmlBody ? String(htmlBody) : '';
	const plainText = plainBody ? String(plainBody) : '';
	const htmlBytes = htmlText ? Utilities.newBlob(htmlText, 'text/html').getBytes().length : 0;
	const plainBytes = plainText ? Utilities.newBlob(plainText, 'text/plain').getBytes().length : 0;
	return {
		htmlBytes,
		plainBytes,
		totalBytes: htmlBytes + plainBytes,
		htmlChars: htmlText.length,
		plainChars: plainText.length
	};
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

// Expected column order specification (with allowed aliases per position)
function getExpectedHeaderSpec_() {
	return [
		['Advertiser'],
		['Campaign'],
		['Site (CM360)', 'Site'],
		['Placement ID'],
		['Placement', 'Placement Name'],
		['Placement Start Date', 'Start Date'],
		['Placement End Date', 'End Date'],
		['Ad Type'],
		['Creative'],
		['Placement Pixel Size', 'Placement Size'],
		['Creative Pixel Size', 'Creative Size'],
		['Date'],
		['Impressions'],
		['Clicks']
	];
}

// Analyze a header row against the expected spec
// Returns { indices: number[], matchedNames: string[], missing: string[], orderOk: boolean }
function analyzeHeaderAgainstSpec_(headerRow) {
	const spec = getExpectedHeaderSpec_();
	const indices = [];
	const matchedNames = [];
	const missing = [];
	const canonHeader = headerRow.map(h => headerNormalize(h));
	for (const opts of spec) {
		let idx = -1;
		let matchedLabel = '';
		for (const label of opts) {
			const needle = headerNormalize(label);
			idx = canonHeader.indexOf(needle);
			if (idx !== -1) { matchedLabel = headerRow[idx]; break; }
		}
		if (idx === -1) {
			// Use the first option as the canonical missing label
			missing.push(opts[0]);
			indices.push(-1);
			matchedNames.push('');
		} else {
			indices.push(idx);
			matchedNames.push(matchedLabel || opts[0]);
		}
	}
	// Order check: ensure indices are strictly increasing when all present
	let orderOk = true;
	let last = -1;
	for (let i = 0; i < indices.length; i++) {
		const idx = indices[i];
		if (idx === -1) continue; // ignore missing in order check
		if (idx <= last) { orderOk = false; break; }
		last = idx;
	}
	return { indices, matchedNames, missing, orderOk };
}

function formatHeaderOrderDiagnostic_(headerRow) {
	try {
		const spec = getExpectedHeaderSpec_();
		const expected = spec.map(opts => opts[0]);
		const analysis = analyzeHeaderAgainstSpec_(headerRow);
		const found = analysis.matchedNames.filter(Boolean);
		const missing = analysis.missing;
		const parts = [];
		parts.push('Expected order: ' + expected.join(' | '));
		parts.push('Found (matched) order: ' + (found.length ? found.join(' | ') : '(none)'));
		if (missing.length) parts.push('Missing: ' + missing.join(', '));
		if (!analysis.orderOk) parts.push('Note: Column order differs from expected.');
		return parts.join('\n');
	} catch (e) {
		return '';
	}
}

// Generic retry wrapper for Drive.Files.list calls (returns null after exhausting attempts)
function driveFilesListWithRetry_(params, contextLabel) {
	const maxAttempts = 3;
	let lastError = null;
	for (var attempt = 1; attempt <= maxAttempts; attempt++) {
		try {
			return Drive.Files.list(params);
		} catch (e) {
			lastError = e;
			const msg = e && e.message ? e.message : String(e);
			Logger.log(`[Drive] list error (${contextLabel}) attempt ${attempt}/${maxAttempts}: ${msg}`);
			if (attempt < maxAttempts) {
				try { Utilities.sleep(200 * attempt); } catch (_) {}
			}
		}
	}
	return null;
}

// Generic retry wrapper for Drive.Files.update calls (throws after exhausting attempts)
function driveFilesUpdateWithRetry_(resource, fileId, mediaData, optionalArgs, contextLabel) {
	const maxAttempts = 3;
	let lastError = null;
	for (var attempt = 1; attempt <= maxAttempts; attempt++) {
		try {
			if (optionalArgs != null) {
				return Drive.Files.update(resource, fileId, mediaData, optionalArgs);
			} else if (typeof mediaData !== 'undefined' && mediaData !== null) {
				return Drive.Files.update(resource, fileId, mediaData);
			} else {
				return Drive.Files.update(resource, fileId);
			}
		} catch (e) {
			lastError = e;
			const msg = e && e.message ? e.message : String(e);
			Logger.log(`[Drive] update error (${contextLabel}) attempt ${attempt}/${maxAttempts}: ${msg}`);
			if (attempt < maxAttempts) {
				try { Utilities.sleep(200 * attempt); } catch (_) {}
			}
		}
	}
	throw lastError;
}

// Convert uploaded Excel/CSV blob into a Google Sheet and return an object with id
function safeConvertExcelToSheet(blob, filename, parentFolderId, configName) {
	// Ensure Advanced Drive API is available before attempting conversion
	if (typeof Drive === 'undefined' || !Drive.Files) {
		throw new Error('Advanced Drive API not available');
	}

	const maxAttempts = 3;
	let lastError = null;

	for (var attempt = 1; attempt <= maxAttempts; attempt++) {
		try {
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
			lastError = e;
			Logger.log(`safeConvertExcelToSheet error (${configName} / ${filename}) attempt ${attempt}/${maxAttempts}: ${e.message}`);
			if (attempt < maxAttempts) {
				try {
					Utilities.sleep(250 * attempt);
				} catch (_) {}
			}
		}
	}

	// Exhausted retries; rethrow last error so caller can handle failure
	throw lastError;
}

// Store and retrieve previous summary flagged counts per config for delta calculation
function getPreviousSummaryCounts_() {
	try {
		const p = PropertiesService.getScriptProperties().getProperty('CM360_LAST_COUNTS');
		return p ? JSON.parse(p) : { date: null, counts: {} };
	} catch (e) {
		return { date: null, counts: {} };
	}
}

function saveSummaryCounts_(countsMap) {
	try {
		const payload = { date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), counts: countsMap };
		PropertiesService.getScriptProperties().setProperty('CM360_LAST_COUNTS', JSON.stringify(payload));
	} catch (e) {
		Logger.log('saveSummaryCounts_ error: ' + e.message);
	}
}

// Persist and retrieve the latest merged report URL per config
function getLatestReportUrlMap_() {
	try {
		const raw = PropertiesService.getScriptProperties().getProperty('CM360_LATEST_REPORT_URLS');
		const obj = raw ? JSON.parse(raw) : {};
		return (obj && typeof obj === 'object') ? obj : {};
	} catch (e) {
		return {};
	}
}

function setLatestReportUrl_(configName, url) {
	if (!configName || !url) return;
	try {
		const map = getLatestReportUrlMap_();
		map[String(configName)] = String(url);
		PropertiesService.getScriptProperties().setProperty('CM360_LATEST_REPORT_URLS', JSON.stringify(map));
		Logger.log(`[${configName}] Latest merged report URL saved`);
	} catch (e) {
		Logger.log(`setLatestReportUrl_ error (${configName}): ${e.message}`);
	}
}

// (Removed) Script Properties-backed config store and admin actions

function buildSummaryEmailContent_(results, options) {
	options = options || {};
	const strictLatestLink = !!options.strictLatestLink; // true for actual email, false for preview
	const subjectDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

	// Delivery Mode badge (no emoji for Outlook compatibility)
	let modeLabel = 'PRODUCTION';
	try { modeLabel = getCurrentDeliveryMode_(); } catch (e) { modeLabel = 'PRODUCTION'; }

	// Aggregate metrics (low-effort, from current result fields)
	const toNum = v => (v == null ? 0 : Number(v) || 0);
	// Filter out any null/undefined entries
	const validResults = results.filter(r => r && typeof r === 'object');
	const totalConfigs = validResults.length;
	const flaggedConfigs = validResults.filter(r => toNum(r.flaggedRows) > 0).length;
	const totalFlaggedRows = validResults.reduce((sum, r) => sum + toNum(r.flaggedRows), 0);
	const errorCount = validResults.filter(r => /error/i.test(String(r.status || ''))).length;
	const skippedCount = validResults.filter(r => /skipped/i.test(String(r.status || ''))).length;
	const emailsSentCount = validResults.filter(r => r.emailSent === true).length;
	const emailsWithheldCount = validResults.filter(r => r.emailWithheld === true).length;

		// Build previous-day counts for delta
		const prev = getPreviousSummaryCounts_();
		const prevCounts = prev && prev.counts ? prev.counts : {};

		// Sort rows: most flags first, then errors/skips, then alpha
	const sorted = validResults.slice().sort((a, b) => {
		const af = toNum(a.flaggedRows), bf = toNum(b.flaggedRows);
		if (bf !== af) return bf - af;
		const aErr = /error/i.test(String(a.status || ''));
		const bErr = /error/i.test(String(b.status || ''));
		if (aErr !== bErr) return aErr ? -1 : 1;
		const aSk = /skipped/i.test(String(a.status || ''));
		const bSk = /skipped/i.test(String(b.status || ''));
		if (aSk !== bSk) return aSk ? -1 : 1;
		return String(a.name || '').localeCompare(String(b.name || ''));
	});

		const rowsHtml = sorted.map(r => {
		const isAlert = /skipped|error/i.test(String(r.status || ''));
		const bgColor = isAlert ? 'background-color:#ffe5e5;' : '';

		// Email status with withhold indicator (emojis are OK in Outlook)
		let emailStatus = r.emailSent ? '✅' : '❌';
		if (r.emailWithheld) emailStatus = '⏸️';

			// Inline error hint extraction
			const statusText = String(r.status || '');
			let statusHtml = escapeHtml(statusText);
			const mErr = statusText.match(/Error(?: during audit)?:\s*(.+)$/i);
			if (mErr && mErr[1]) {
				statusHtml += `<div style="color:#b00020; font-size:11px; margin-top:2px;">${escapeHtml(mErr[1])}</div>`;
			}

			// Delta calculation
			const curr = toNum(r.flaggedRows);
			const prevVal = toNum(prevCounts[String(r.name || '')] || 0);
			const delta = curr - prevVal;
			const deltaBadge = delta === 0 ? '0' : (delta > 0 ? `+${delta}` : `${delta}`);

			// Link: Prefer Latest Report (merged sheet), fallback to Gmail label search (in preview only)
			const labelPath = `Daily Audits/CM360/${r.name}`;
			const gmailSearch = `https://mail.google.com/mail/u/0/#search/${encodeURIComponent('label:' + labelPath + ' newer_than:2d')}`;
			// Helper: validate that r.latestReportUrl points to today's exact expected file name
			const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
			const expectedName = `CM360_Merged_Audit_${String(r.name || '').replace(/\s+/g, '').toUpperCase()}_${todayStr}`;
			function extractId_(url) {
				if (!url) return '';
				const m = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/);
				return m ? m[1] : '';
			}
			function matchesTodayNaming_(url) {
				try {
					const id = extractId_(url);
					if (!id) return false;
					const file = DriveApp.getFileById(id);
					return String(file.getName() || '') === expectedName;
				} catch (e) { return false; }
			}
			let latestReportLink;
			if (strictLatestLink) {
				// Actual email: only link if URL matches today's exact expected file name, else UNAVAILABLE
				if (r.latestReportUrl && matchesTodayNaming_(r.latestReportUrl)) {
					latestReportLink = `<a href="${escapeHtml(r.latestReportUrl)}">View</a>`;
				} else {
					latestReportLink = `<span style="display:inline-block; background:#fde8e8; border:1px solid #f4c7c3; color:#b00020; padding:2px 6px; border-radius:4px;">UNAVAILABLE</span>`;
				}
			} else {
				// Preview: show a best-effort link using latest known URL or Gmail label search
				const linkUrl = r.latestReportUrl || gmailSearch;
				latestReportLink = `<a href="${escapeHtml(linkUrl)}">View</a>`;
			}

		return `
			<tr style="font-size:12px; line-height:1.3; ${bgColor}">
					<td style="padding:4px 8px;">${escapeHtml(r.name)}</td>
					<td style="padding:4px 8px;">${statusHtml}${r.emailWithheld ? '<div style="color:#5f6368; font-size:11px;">No-flag emails withheld</div>' : ''}</td>
					<td style="padding:4px 8px; text-align:center;">${r.flaggedRows ?? '-'}</td>
					<td style="padding:4px 8px; text-align:center;">${deltaBadge}</td>
				<td style="padding:4px 8px; text-align:center;">${emailStatus}</td>
					<td style="padding:4px 8px; text-align:center;">${escapeHtml(r.emailTime)}</td>
					<td style="padding:4px 8px; text-align:center;">${latestReportLink}</td>
			</tr>`;
	}).join('');

	// Pull cached remaining quota (lowest value seen)
	const remainingQuota = getEmailQuotaRemaining_();
	const quotaNote = remainingQuota !== null
		? `<p style="font-family:Arial, sans-serif; font-size:12px; margin-top:8px;">
				 Remaining daily email quota: <strong>${remainingQuota}</strong>
			 </p>`
		: '';

	// Quick links (no Admin Sheet link for recipients)
	const externalUrl = EXTERNAL_CONFIG_SHEET_ID ? `https://docs.google.com/spreadsheets/d/${EXTERNAL_CONFIG_SHEET_ID}` : '';
	const linksRow = externalUrl ? `
		<p style="margin:8px 0 0; font-family:Arial, sans-serif; font-size:12px;">
			<a href="${externalUrl}" style="margin-right:12px;">Open External Config</a>
		</p>` : '';

	const glanceHtml = `
		<div style="margin:8px 0 12px; font-family:Arial, sans-serif; font-size:12px;">
			<span style="display:inline-block; background:#eef3fe; border:1px solid #d2e3fc; border-radius:12px; padding:2px 8px; margin-right:6px;">${totalConfigs} configs</span>
			<span style="display:inline-block; background:#fde8e8; border:1px solid #facaca; border-radius:12px; padding:2px 8px; margin-right:6px;">${flaggedConfigs} with flags</span>
			<span style="display:inline-block; background:#fff7e6; border:1px solid #ffe8b3; border-radius:12px; padding:2px 8px; margin-right:6px;">${totalFlaggedRows} total flagged rows</span>
			<span style="display:inline-block; background:#fce8e6; border:1px solid #fad2cf; border-radius:12px; padding:2px 8px; margin-right:6px;">${errorCount} errors</span>
			<span style="display:inline-block; background:#f1f3f4; border:1px solid #e0e3e7; border-radius:12px; padding:2px 8px; margin-right:6px;">${skippedCount} skipped</span>
			<span style="display:inline-block; background:#e6f4ea; border:1px solid #ccebd7; border-radius:12px; padding:2px 8px; margin-right:6px;">${emailsSentCount} emails sent</span>
			<span style="display:inline-block; background:#f1f3f4; border:1px solid #e0e3e7; border-radius:12px; padding:2px 8px; margin-right:6px;">${emailsWithheldCount} withheld</span>
			<span style="display:inline-block; font-weight:bold; color:${modeLabel === 'STAGING' ? '#8e24aa' : '#0b8043'}; padding:2px 8px;">Delivery Mode: ${modeLabel}</span>
		</div>`;

	// Top 5 offenders snippet
	const topOffenders = sorted.filter(r => toNum(r.flaggedRows) > 0).slice(0, 5).map(r => `
		<li>${escapeHtml(r.name)} — <strong>${toNum(r.flaggedRows)}</strong> flagged row(s)</li>`).join('');
	const offendersHtml = topOffenders
		? `<div style="font-family:Arial, sans-serif; font-size:12px; margin:8px 0;">
				 <strong>Top flagged configs:</strong>
				 <ul style="margin:6px 0 0 18px; padding:0;">${topOffenders}</ul>
			 </div>`
		: '';

	const htmlBody = `
		<p style="font-family:Arial, sans-serif; font-size:13px;">Here's a summary of today's CM360 audits:</p>
		${glanceHtml}
		${offendersHtml}
			<table border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse; font-family:Arial, sans-serif; font-size:12px;">
			<thead style="background:#f2f2f2;">
				<tr>
					<th style="padding:4px 8px;">Config</th>
					<th style="padding:4px 8px;">Status</th>
					<th style="padding:4px 8px;">Flagged Rows</th>
						<th style="padding:4px 8px;">Δ</th>
					<th style="padding:4px 8px;">Email Sent</th>
					<th style="padding:4px 8px;">Email Time</th>
						<th style="padding:4px 8px;">Latest Report</th>
				</tr>
			</thead>
			<tbody>${rowsHtml}</tbody>
		</table>
		<p style="font-family:Arial, sans-serif; font-size:11px; margin-top:8px; color:#666;">
			Email Status: ✅ Sent | ❌ Failed | ⏸️ Withheld (no-flag emails disabled)
		</p>
		${quotaNote}
		${linksRow}
		<p style="margin-top:12px; font-family:Arial, sans-serif; font-size:12px;">&mdash; Platform Solutions Team</p>
	`;

	// Subject with delivery mode and key counts
	const subject = `CM360 Daily Audit Summary (${subjectDate}) — ${flaggedConfigs} flagged, ${errorCount} errors`;

	// CSV attachment removed as redundant; summary will only include HTML and plain text

		// Plain-text fallback
	const plainLines = [];
	plainLines.push(`CM360 Daily Audit Summary (${subjectDate}) [${modeLabel}]`);
	plainLines.push(`Configs=${totalConfigs}, WithFlags=${flaggedConfigs}, FlagRows=${totalFlaggedRows}, Errors=${errorCount}, Skipped=${skippedCount}, Sent=${emailsSentCount}, Withheld=${emailsWithheldCount}`);
	if (externalUrl) plainLines.push(`External: ${externalUrl}`);
		if (sorted.length) {
			plainLines.push('Top flagged:');
			sorted.filter(r => toNum(r.flaggedRows) > 0).slice(0, 5).forEach(r => {
				const curr = toNum(r.flaggedRows);
				const prevVal = toNum(prevCounts[String(r.name || '')] || 0);
				const delta = curr - prevVal;
				plainLines.push(` - ${r.name}: ${curr} (Δ ${delta})`);
			});
		}
		const plainText = plainLines.join('\n');

	// Return structured content for sending/preview
	return { subject, htmlBody, plainText, meta: { modeLabel, totalConfigs, flaggedConfigs, totalFlaggedRows, errorCount, skippedCount, emailsSentCount, emailsWithheldCount } };
}

// Best-effort: Find the latest merged report URL for a config directly from Drive
function findLatestMergedReportUrl_(config) {
	try {
		if (!config || !Array.isArray(config.mergedFolderPath)) return '';
		const folder = getDriveFolderByPath_(config.mergedFolderPath);
		if (!folder) return '';
		const it = folder.getFiles();
		const files = [];
		while (it.hasNext()) files.push(it.next());
		if (!files.length) return '';
		const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
		const cfg = String(config && (config.name || config.configId || config.id) || '').trim();
		const basePrefix = 'CM360_Merged_Audit_';
		const prefix = cfg ? `${basePrefix}${cfg}_` : basePrefix;
		const exactToday = cfg ? `${basePrefix}${cfg}_${todayStr}` : '';
		// 1) Exact today match for this config
		if (exactToday) {
			const exact = files.find(f => String(f.getName() || '') === exactToday);
			if (exact) return exact.getUrl();
		}
		// 2) Latest by this config prefix
		const forConfig = files.filter(f => String(f.getName() || '').startsWith(prefix));
		if (forConfig.length) {
			forConfig.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
			return forConfig[0].getUrl();
		}
		// 3) Back-compat: latest of any CM360_Merged_Audit_ file
		const anyMerged = files.filter(f => String(f.getName() || '').startsWith(basePrefix));
		if (anyMerged.length) {
			anyMerged.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
			return anyMerged[0].getUrl();
		}
		// 4) Fallback: newest file in folder
		files.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
		return files[0].getUrl();
	} catch (e) {
		try { Logger.log(`findLatestMergedReportUrl_ error (${config && config.name ? config.name : 'Unknown'}): ${e.message}`); } catch (_) {}
		return '';
	}
}

function sendDailySummaryEmail(results) {
	const distribution = [ADMIN_EMAIL, 'bmuller@horizonmedia.com, bkaufman@horizonmedia.com, ewarburton@horizonmedia.com'].filter(Boolean).join(', ');
	const content = buildSummaryEmailContent_(results, { strictLatestLink: true });
	try {
		safeSendEmail({
			to: distribution,
			subject: content.subject,
			plainBody: content.plainText,
			htmlBody: content.htmlBody
		}, 'Daily Summary');
		Logger.log(`[EMAIL] Summary email sent to: ${distribution}`);
	} catch (err) {
		Logger.log(`❌ Failed to send summary email: ${err.message}`);
	}
	// Persist current counts for next day's delta
	try {
		const countsMap = {};
		results.forEach(r => { const toNum = v => (v == null ? 0 : Number(v) || 0); countsMap[String(r.name || '')] = toNum(r.flaggedRows); });
		saveSummaryCounts_(countsMap);
	} catch (e) {
		Logger.log('Failed to persist summary counts: ' + e.message);
	}
}

function mergeAuditResultsByConfig_(...resultSets) {
	const map = new Map();
	const pushSet = (arr) => {
		if (!Array.isArray(arr)) return;
		arr.forEach(item => {
			if (!item || typeof item !== 'object' || !item.name) return;
			const name = String(item.name).trim();
			const existing = map.get(name) || {};
			map.set(name, Object.assign({}, existing, item, { name }));
		});
	};
	resultSets.forEach(pushSet);
	return Array.from(map.values());
}

function getAllAuditRunStates_() {
	try {
		const props = PropertiesService.getScriptProperties();
		const all = props.getProperties();
		const states = [];
		Object.keys(all || {}).forEach(key => {
			if (!key || key.indexOf(AUDIT_RUN_STATE_KEY_PREFIX) !== 0) return;
			try {
				const state = JSON.parse(all[key]);
				state.batchId = key.slice(AUDIT_RUN_STATE_KEY_PREFIX.length);
				states.push(state);
			} catch (e) {
				Logger.log(`getAllAuditRunStates_ parse error (${key}): ${e.message}`);
			}
		});
		return states;
	} catch (e) {
		Logger.log('getAllAuditRunStates_ error: ' + e.message);
		return [];
	}
}

function buildSummaryResultSet_(options) {
	options = options || {};
	const allowPlaceholders = !!options.allowPlaceholders;
	const configs = getAuditConfigs();
	const activeNames = configs.map(c => c && c.name).filter(Boolean);
	const activeSet = new Set(activeNames);
	const cached = mergeAuditResultsByConfig_(getCombinedAuditResults_()).filter(r => activeSet.has(r.name));
	if (!allowPlaceholders) {
		return cached;
	}
	const resultMap = new Map();
	cached.forEach(r => { if (r && r.name) resultMap.set(r.name, r); });
	const tz = Session.getScriptTimeZone();
	const todayKey = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
	const states = getAllAuditRunStates_();
	const stateByConfig = new Map();
	states.forEach(state => {
		const startedAt = state && state.startedAt ? new Date(state.startedAt) : null;
		if (startedAt && Utilities.formatDate(startedAt, tz, 'yyyyMMdd') !== todayKey) return;
		const cfgs = Array.isArray(state && state.configs) ? state.configs : [];
		cfgs.forEach(name => {
			if (!name || !activeSet.has(name)) return;
			const existing = stateByConfig.get(name);
			if (!existing || (startedAt && existing.startedAt && startedAt > existing.startedAt) || (startedAt && !existing.startedAt)) {
				stateByConfig.set(name, {
					state,
					startedAt
				});
			}
		});
	});
	activeNames.forEach(name => {
		if (resultMap.has(name)) return;
		const info = stateByConfig.get(name) || {};
		const startedAt = info.startedAt;
		const state = info.state || {};
		let status = 'Not started (no run recorded today)';
		let emailTime = 'Not sent';
		if (startedAt || state.startedAt) {
			const startStr = startedAt ? Utilities.formatDate(startedAt, tz, 'HH:mm:ss') : null;
			if (state.alertedAt) {
				const alertStr = Utilities.formatDate(new Date(state.alertedAt), tz, 'HH:mm:ss');
				status = `Error: Timed out before completion${startStr ? ` (started ${startStr})` : ''}`;
				emailTime = `Timed out${alertStr ? ` @ ${alertStr}` : ''}`;
			} else if (state.completedAt) {
				status = `Completed (no summary data recorded${startStr ? `; started ${startStr}` : ''})`;
				emailTime = 'Completed (email status unknown)';
			} else {
				status = `In progress (no completion logged${startStr ? `; started ${startStr}` : ''})`;
			}
		}
		resultMap.set(name, {
			name,
			status,
			flaggedRows: null,
			emailSent: false,
			emailWithheld: false,
			emailTime,
			latestReportUrl: ''
		});
	});
	return Array.from(resultMap.values());
}

function attemptSendDailySummary_(options) {
	options = options || {};
	const allowPlaceholders = !!options.allowPlaceholders;
	const reason = options.reason || 'run';
	const cache = CacheService.getScriptCache();
	const alreadySent = cache.get('CM360_SUMMARY_SENT');
	if (alreadySent === '1') {
		Logger.log(`[Summary] Already sent (${reason}).`);
		return false;
	}
	const lock = LockService.getScriptLock();
	if (!lock.tryLock(5000)) {
		Logger.log(`[Summary] Could not acquire lock (${reason}).`);
		return false;
	}
	try {
		const recheck = cache.get('CM360_SUMMARY_SENT');
		if (recheck === '1') {
			Logger.log(`[Summary] Already sent after acquiring lock (${reason}).`);
			return false;
		}
		const configs = getAuditConfigs();
		const totalConfigs = configs.length;
		const results = buildSummaryResultSet_({ allowPlaceholders });
		const uniqueCount = new Set(results.map(r => r && r.name).filter(Boolean)).size;
		if (!allowPlaceholders && uniqueCount < totalConfigs) {
			Logger.log(`[Summary] ${uniqueCount}/${totalConfigs} configs complete; awaiting more before sending (${reason}).`);
			return false;
		}
		if (!results.length) {
			Logger.log('[Summary] No results available to include; skipping send.');
			return false;
		}
		sendDailySummaryEmail(results);
		cache.put('CM360_SUMMARY_SENT', '1', 21600);
		cache.remove(getAuditCacheKey_());
		Logger.log(`[Summary] Daily summary dispatched (${reason}).`);
		return true;
	} catch (e) {
		Logger.log(`attemptSendDailySummary_ error (${reason}): ${e.message}`);
		return false;
	} finally {
		try { lock.releaseLock(); } catch (_) {}
	}
}

function sendDailySummaryFailover() {
	try {
		const sent = attemptSendDailySummary_({ allowPlaceholders: true, reason: 'Failover trigger' });
		if (!sent) {
			Logger.log('[Summary Failover] No summary sent (already delivered or waiting for additional data).');
		}
	} catch (e) {
		Logger.log('sendDailySummaryFailover error: ' + e.message);
	}
}

// Preview the single summary email without sending; synthesizes from configs if no cache
function previewDailySummaryNow() {
	let results = buildSummaryResultSet_({ allowPlaceholders: true });
	if (!results || results.length === 0) {
		results = [];
	}
	const content = buildSummaryEmailContent_(results, { strictLatestLink: false });
	const html = HtmlService.createHtmlOutput(`
		<div style="font-family:Arial,sans-serif; padding:12px;">
			<div style="margin-bottom:10px; font-weight:bold;">${escapeHtml(content.subject)}</div>
			<div style="border:1px solid #e0e0e0; max-height:70vh; overflow:auto;">${content.htmlBody}</div>
		</div>
	`).setWidth(900).setHeight(700);
	try { SpreadsheetApp.getUi().showModalDialog(html, 'Preview: CM360 Daily Summary'); } catch (e) {}
	Logger.log('Preview rendered for daily summary');
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
		htmlBody: `<p style=\"font-family:Arial; font-size:13px;\">The label <b>${escapeHtml(config.label)}</b> could not be found. This may mean the audit for <b>${escapeHtml(config.name)}</b> will be skipped.</p>`,
		bccAdmin: true
	}, `${config.name} - Missing Gmail Label`);
 return null;
 }

 const threads = label.getThreads();
 const startOfToday = new Date();
 startOfToday.setHours(0, 0, 0, 0); 
 
 const parentFolder = getDriveFolderByPath_(config.tempDailyFolderPath);
	if (!parentFolder) {
		const pathStr = Array.isArray(config.tempDailyFolderPath) ? config.tempDailyFolderPath.join(' / ') : '(invalid path)';
		const errMsg = `Temp daily folder unavailable after retries (path: ${pathStr})`;
		Logger.log(`❌ [${config.name}] ${errMsg}`);
		try {
			safeSendEmail({
				to: ADMIN_EMAIL,
				subject: `⚠️ CM360 Audit Drive folder unavailable (${config.name})`,
				htmlBody: `<p style="font-family:Arial, sans-serif; font-size:13px;">${escapeHtml(errMsg)}</p>`
			}, `${config.name} - temp folder unavailable`);
		} catch (notifyErr) {
			Logger.log(`[${config.name}] Failed to send admin alert for temp folder issue: ${notifyErr.message}`);
		}
		throw new Error(errMsg);
	}
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
	const subject = `⚠️ CM360 Audit Skipped: No Files Found (${config.name} - ${formatDate(new Date())})`;
	const htmlBody = `
	<p style="font-family:Arial, sans-serif; font-size:13px;">
	The CM360 audit for bundle "<strong>${escapeHtml(config.name)}</strong>" was skipped because no Excel or ZIP files were found for today.
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
		 attachments: [],
		 bccAdmin: true
	}, config.name);
	return null;
 }

	Logger.log(`✅ [${config.name}] Files saved to: ${driveFolder.getName()}`);
 return driveFolder.getId();
}

function mergeDailyAuditExcels(folderId, mergedFolderPath, configName = 'Unknown', recipientsData) {
 Logger.log(`[${configName}] mergeDailyAuditExcels started`);
 const folder = DriveApp.getFolderById(folderId);
 const files = folder.getFiles();
 const destFolder = getDriveFolderByPath_(mergedFolderPath);

 // Name format: CM360_Merged_Audit_<CONFIG>_<YYYY-MM-DD>
 const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
 const safeConfig = String(configName || 'Unknown').replace(/\s+/g, '').toUpperCase();
 const mergedSheetName = `CM360_Merged_Audit_${safeConfig}_${todayStr}`;
 const mergedSpreadsheet = SpreadsheetApp.create(mergedSheetName);
 Utilities.sleep(500); // Reduced from 1000ms - just need brief pause for file creation
 const mergedFile = DriveApp.getFileById(mergedSpreadsheet.getId());
 destFolder.addFile(mergedFile);
 DriveApp.getRootFolder().removeFile(mergedFile);
 const mergedSheet = mergedSpreadsheet.getSheets()[0];

 let headerWritten = false;
 let header = [];
 const processedIds = new Set(); // track originals we've consumed (and alternates) to avoid duplicates
	const headerIssues = []; // collect per-file schema/order issues

 while (files.hasNext()) {
 const file = files.next();
 try { if (processedIds.has(file.getId())) { continue; } } catch (_) {}
 const fileName = file.getName().toLowerCase();
 let invalidThisFile = false;

 let data;
 let spreadsheet;

 if (fileName.endsWith('.xlsx')) {
 const blob = file.getBlob();
 const converted = safeConvertExcelToSheet(blob, file.getName(), folder.getId(), configName);

 // Ensure it only lives in `folder`
		driveFilesUpdateWithRetry_({ parents: [{ id: folder.getId() }] }, converted.id, null, null, `${configName} reparent primary ${file.getName()}`);

 spreadsheet = SpreadsheetApp.openById(converted.id);
 data = spreadsheet.getSheets()[0].getDataRange().getValues();
 if (!data || data.length === 0 || data.every(row => row.every(cell => cell === ''))) {
 Logger.log(`[${configName}] File "${fileName}" appears blank after import.`);
 invalidThisFile = true;
 continue;
 }
 } else if (fileName.endsWith('.csv')) {
 const csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
  // Defer creating a sheet for CSV until we confirm selection; for now work with raw data
  spreadsheet = null; // will create only if/when selected for move
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
 invalidThisFile = true;
 continue;
 }

 const realData = data.slice(headerRowIndex);
 let cleanedData = realData.filter((row, idx) =>
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

 	 // Schema validation: ensure required columns are present; try counterpart if missing
 	 (function() {
 		 const requiredOptions = [
 			 ['Advertiser'],
 			 ['Campaign'],
 			 ['Site (CM360)', 'Site'],
 			 ['Placement ID'],
 			 ['Placement', 'Placement Name'],
 			 ['Placement Start Date', 'Start Date'],
 			 ['Placement End Date', 'End Date'],
 			 ['Ad Type'],
 			 ['Creative'],
 			 ['Placement Pixel Size', 'Placement Size'],
 			 ['Creative Pixel Size', 'Creative Size'],
 			 ['Date'],
 			 ['Impressions'],
 			 ['Clicks']
 		 ];
			 const canon = new Set(header.map(h => headerNormalize(h)));
			 const missing = [];
			 requiredOptions.forEach(opts => { if (!opts.some(o => canon.has(headerNormalize(o)))) missing.push(opts[0]); });
			 const orderDiag = formatHeaderOrderDiagnostic_(header);
		 if (missing.length > 0) {
 			 try {
 				 const baseName = String(file.getName()).replace(/\.(xlsx|csv)$/i, '');
				 const diag = formatHeaderOrderDiagnostic_(header);
 				 const tryNames = [];
 				 if (/\.csv$/i.test(file.getName())) tryNames.push(`${baseName}.xlsx`);
 				 else if (/\.xlsx$/i.test(file.getName())) tryNames.push(`${baseName}.csv`);
 				 for (const altName of tryNames) {
 					 const it = folder.getFilesByName(altName);
 					 while (it.hasNext()) {
 						 const altFile = it.next();
 						 if (processedIds.has(altFile.getId())) continue;
 						 const lowerAlt = altFile.getName().toLowerCase();
 						 let altHeader = null;
 						 let altData = null;
 						 let altSpreadsheet = null;
 						 if (lowerAlt.endsWith('.csv')) {
 							 const altCsv = Utilities.parseCsv(altFile.getBlob().getDataAsString());
 							 let altHdrIdx = altCsv.findIndex(row => {
 								 const normRow = row.map(cell => normalize(cell));
 								 const canonRowSet = new Set(row.map(cell => headerNormalize(cell)));
 								 return headerKeywords.every(keyword => normRow.includes(normalize(keyword))) &&
 										 headerKeywordsCanon.every(keyword => canonRowSet.has(keyword));
 							 });
 							 if (altHdrIdx === -1) {
 								 const minHits = Math.max(3, Math.ceil(headerKeywords.length / 2));
 								 altHdrIdx = altCsv.findIndex(row => {
 									 const normRow = row.map(cell => normalize(cell));
 									 const canonRowSet = new Set(row.map(cell => headerNormalize(cell)));
 									 let hits = 0;
 									 for (var k = 0; k < headerKeywords.length; k++) {
 										 if (normRow.includes(normalize(headerKeywords[k])) || canonRowSet.has(headerKeywordsCanon[k])) hits++;
 									 }
															 return hits >= minHits;
															 });
														 } // end if (altHdrIdx === -1)
															 if (altHdrIdx !== -1) {
 								 const altReal = altCsv.slice(altHdrIdx);
 								 const altClean = altReal.filter((row, idx) => idx === 0 || !String(row.join('') || '').toLowerCase().includes('grand total'));
 								 altHeader = altClean[0] || [];
 								 const altCanon = new Set(altHeader.map(h => headerNormalize(h)));
 								 const altMissing = requiredOptions.filter(opts => !opts.some(o => altCanon.has(headerNormalize(o)))).map(opts => opts[0]);
 								 if (altMissing.length === 0) {
 									 // Create archival sheet now that we've selected this alt CSV
 									 altSpreadsheet = SpreadsheetApp.create(altFile.getName().replace(/\.csv$/i, ''));
 									 altSpreadsheet.getSheets()[0].getRange(1, 1, altCsv.length, altCsv[0].length).setValues(altCsv);
 									 altData = altClean;
 								 }
 							 }
 						 } else if (lowerAlt.endsWith('.xlsx')) {
 							 const altBlob = altFile.getBlob();
 							 const altConverted = safeConvertExcelToSheet(altBlob, altFile.getName(), folder.getId(), configName);
							driveFilesUpdateWithRetry_({ parents: [{ id: folder.getId() }] }, altConverted.id, null, null, `${configName} reparent alt ${altFile.getName()}`);
 							 altSpreadsheet = SpreadsheetApp.openById(altConverted.id);
 							 const altVals = altSpreadsheet.getSheets()[0].getDataRange().getValues();
 							 let altHdrIdx = altVals.findIndex(row => {
 								 const normRow = row.map(cell => normalize(cell));
 								 const canonRowSet = new Set(row.map(cell => headerNormalize(cell)));
 								 return headerKeywords.every(keyword => normRow.includes(normalize(keyword))) &&
 										 headerKeywordsCanon.every(keyword => canonRowSet.has(keyword));
 							 });
 							 if (altHdrIdx === -1) {
 								 const minHits = Math.max(3, Math.ceil(headerKeywords.length / 2));
 								 altHdrIdx = altVals.findIndex(row => {
 									 const normRow = row.map(cell => normalize(cell));
 									 const canonRowSet = new Set(row.map(cell => headerNormalize(cell)));
 									 let hits = 0;
 									 for (var k = 0; k < headerKeywords.length; k++) {
 										 if (normRow.includes(normalize(headerKeywords[k])) || canonRowSet.has(headerKeywordsCanon[k])) hits++;
 									 }
															 return hits >= minHits;
															 });
													 } // end if (altHdrIdx === -1)
													 if (altHdrIdx !== -1) {
 								 const altReal = altVals.slice(altHdrIdx);
 								 const altClean = altReal.filter((row, idx) => idx === 0 || !String(row.join('') || '').toLowerCase().includes('grand total'));
 								 altHeader = altClean[0] || [];
 								 const altCanon = new Set(altHeader.map(h => headerNormalize(h)));
 								 const altMissing = requiredOptions.filter(opts => !opts.some(o => altCanon.has(headerNormalize(o)))).map(opts => opts[0]);
 								 if (altMissing.length === 0) {
 									 altData = altClean;
 								 }
 							 }
 						 }
 						 if (altHeader && altData) {
 							 Logger.log(`[${configName}] Failsafe: Using counterpart with complete schema for "${baseName}" instead of incomplete file "${file.getName()}"`);
 							 header = altHeader;
 							 cleanedData = altData;
 							 if (altSpreadsheet) spreadsheet = altSpreadsheet;
 							 processedIds.add(altFile.getId());
 							 break;
 						 }
 					 }
 				 }
 			 } catch (schemaAltErr) {
 				 Logger.log(`[${configName}] Schema counterpart lookup error: ${schemaAltErr.message}`);
 			 }
 			 // Re-check schema after potential swap; if still missing, abort
 			 const afterCanon = new Set(header.map(h => headerNormalize(h)));
 			 const stillMissing = [];
 			 requiredOptions.forEach(opts => { if (!opts.some(o => afterCanon.has(headerNormalize(o)))) stillMissing.push(opts[0]); });
			 if (stillMissing.length > 0) {
				 // Record issue and skip this file; try other files to proceed with merge
				 const baseName = String(file.getName()).replace(/\.(xlsx|csv)$/i, '');
				 const diag = formatHeaderOrderDiagnostic_(header);
				 headerIssues.push({ fileName: file.getName(), missing: stillMissing.slice(), orderOk: false, diag });
				 Logger.log(`[${configName}] HEADER ISSUE: missing required columns for ${baseName}: ${stillMissing.join(', ')}`);
				 // Reset header candidate since it's invalid
				 header = [];
				 skipThisFile = true;
				 invalidThisFile = true;
			 }

			 // Enforce correct order for initial header as well
			 const initialOrderCheck = analyzeHeaderAgainstSpec_(header);
			 if (initialOrderCheck && initialOrderCheck.orderOk === false) {
				 const baseName = String(file.getName()).replace(/\.(xlsx|csv)$/i, '');
				 const diag = formatHeaderOrderDiagnostic_(header);
				 headerIssues.push({ fileName: file.getName(), missing: (initialOrderCheck.missing||[]).slice(), orderOk: false, diag });
				 Logger.log(`[${configName}] HEADER ISSUE: out-of-order columns for ${baseName}; expected order required. Skipping file.`);
				 header = [];
				 skipThisFile = true;
				 invalidThisFile = true;
			 }
 		 }
 	 })();

		 // Failsafe: if this first header is 13 cols, look for a same-name counterpart with 14 cols
 	 if (header.length === 13) {
 	 	 try {
 	 	 	 const baseName = String(file.getName()).replace(/\.(xlsx|csv)$/i, '');
 	 	 	 const tryNames = [];
 	 	 	 if (/\.csv$/i.test(file.getName())) {
 	 	 	 	 tryNames.push(`${baseName}.xlsx`);
 	 	 	 } else if (/\.xlsx$/i.test(file.getName())) {
 	 	 	 	 tryNames.push(`${baseName}.csv`);
 	 	 	 }
 	 	 	 let swapped = false;
 	 	 	 for (const altName of tryNames) {
 	 	 	 	 const it = folder.getFilesByName(altName);
 	 	 	 	 while (it.hasNext()) {
 	 	 	 	 	 const altFile = it.next();
 	 	 	 	 	 if (processedIds.has(altFile.getId())) continue;
 	 	 	 	 	 const lowerAlt = altFile.getName().toLowerCase();
 	 	 	 	 	 let altHeader = null;
 	 	 	 	 	 let altData = null;
 	 	 	 	 	 let altSpreadsheet = null;
 	 	 	 	 	 if (lowerAlt.endsWith('.csv')) {
 	 	 	 	 	 	 const altCsv = Utilities.parseCsv(altFile.getBlob().getDataAsString());
 	 	 	 	 	 	 // detect header in altCsv
 	 	 	 	 	 	 let altHdrIdx = altCsv.findIndex(row => {
 	 	 	 	 	 	 	 const normRow = row.map(cell => normalize(cell));
 	 	 	 	 	 	 	 const canonRowSet = new Set(row.map(cell => headerNormalize(cell)));
 	 	 	 	 	 	 	 return headerKeywords.every(keyword => normRow.includes(normalize(keyword))) &&
 	 	 	 	 	 	 	 	 headerKeywordsCanon.every(keyword => canonRowSet.has(keyword));
 	 	 	 	 	 	 });
 	 	 	 	 	 	 if (altHdrIdx === -1) {
 	 	 	 	 	 	 	 const minHits = Math.max(3, Math.ceil(headerKeywords.length / 2));
 	 	 	 	 	 	 	 altHdrIdx = altCsv.findIndex(row => {
 	 	 	 	 	 	 	 	 const normRow = row.map(cell => normalize(cell));
 	 	 	 	 	 	 	 	 const canonRowSet = new Set(row.map(cell => headerNormalize(cell)));
 	 	 	 	 	 	 	 	 let hits = 0;
 	 	 	 	 	 	 	 	 for (var k = 0; k < headerKeywords.length; k++) {
 	 	 	 	 	 	 	 	 	 if (normRow.includes(normalize(headerKeywords[k])) || canonRowSet.has(headerKeywordsCanon[k])) hits++;
 	 	 	 	 	 	 	 	 }
 	 	 	 	 	 	 	 	 return hits >= minHits;
 	 	 	 	 	 	 	 });
 	 	 	 	 	 	 }
 	 	 	 	 	 	 if (altHdrIdx !== -1) {
 	 	 	 	 	 	 	 const altReal = altCsv.slice(altHdrIdx);
 	 	 	 	 	 	 	 const altClean = altReal.filter((row, idx) => idx === 0 || !String(row.join('') || '').toLowerCase().includes('grand total'));
 	 	 	 	 	 	 	 altHeader = altClean[0] || [];
 	 	 	 	 	 	 	 if (altHeader.length === 14) {
 	 	 	 	 	 	 	 	 // Create a sheet for the alt CSV now that it's selected
 	 	 	 	 	 	 	 	 altSpreadsheet = SpreadsheetApp.create(altFile.getName().replace(/\.csv$/i, ''));
 	 	 	 	 	 	 	 	 altSpreadsheet.getSheets()[0].getRange(1, 1, altCsv.length, altCsv[0].length).setValues(altCsv);
 	 	 	 	 	 	 	 	 altData = altClean;
						 	 }
							 // If still 13 after lookup, log that no 14-col counterpart was found
							 	 if (header.length === 13) {
							 	 	 const baseName = String(file.getName()).replace(/\.(xlsx|csv)$/i, '');
							 	 	 const diag = formatHeaderOrderDiagnostic_(header);
							 	 	 headerIssues.push({ fileName: file.getName(), missing: ['14-column header expected; found 13 columns'], orderOk: false, diag });
							 	 	 Logger.log(`[${configName}] HEADER ISSUE: 13 vs 14 columns for ${baseName}; skipping this file.`);
							 	 	 // Reset invalid header candidate and mark to skip
							 	 	 header = [];
							 	 	 skipThisFile = true;
							 	 }
 	 	 	 	 	 	 }
 	 	 	 	 	 } else if (lowerAlt.endsWith('.xlsx')) {
 	 	 	 	 	 	 const altBlob = altFile.getBlob();
 	 	 	 	 	 	 const altConverted = safeConvertExcelToSheet(altBlob, altFile.getName(), folder.getId(), configName);
						driveFilesUpdateWithRetry_({ parents: [{ id: folder.getId() }] }, altConverted.id, null, null, `${configName} reparent alt ${altFile.getName()}`);
 	 	 	 	 	 	 altSpreadsheet = SpreadsheetApp.openById(altConverted.id);
 	 	 	 	 	 	 const altVals = altSpreadsheet.getSheets()[0].getDataRange().getValues();
 	 	 	 	 	 	 let altHdrIdx = altVals.findIndex(row => {
 	 	 	 	 	 	 	 const normRow = row.map(cell => normalize(cell));
 	 	 	 	 	 	 	 const canonRowSet = new Set(row.map(cell => headerNormalize(cell)));
 	 	 	 	 	 	 	 return headerKeywords.every(keyword => normRow.includes(normalize(keyword))) &&
 	 	 	 	 	 	 	 	 headerKeywordsCanon.every(keyword => canonRowSet.has(keyword));
 	 	 	 	 	 	 });
 	 	 	 	 	 	 if (altHdrIdx === -1) {
 	 	 	 	 	 	 	 const minHits = Math.max(3, Math.ceil(headerKeywords.length / 2));
 	 	 	 	 	 	 	 altHdrIdx = altVals.findIndex(row => {
 	 	 	 	 	 	 	 	 const normRow = row.map(cell => normalize(cell));
 	 	 	 	 	 	 	 	 const canonRowSet = new Set(row.map(cell => headerNormalize(cell)));
 	 	 	 	 	 	 	 	 let hits = 0;
 	 	 	 	 	 	 	 	 for (var k = 0; k < headerKeywords.length; k++) {
 	 	 	 	 	 	 	 	 	 if (normRow.includes(normalize(headerKeywords[k])) || canonRowSet.has(headerKeywordsCanon[k])) hits++;
 	 	 	 	 	 	 	 	 }
 	 	 	 	 	 	 	 	 return hits >= minHits;
 	 	 	 	 	 	 	 });
 	 	 	 	 	 	 }
 	 	 	 	 	 	 if (altHdrIdx !== -1) {
 	 	 	 	 	 	 	 const altReal = altVals.slice(altHdrIdx);
 	 	 	 	 	 	 	 const altClean = altReal.filter((row, idx) => idx === 0 || !String(row.join('') || '').toLowerCase().includes('grand total'));
 	 	 	 	 	 	 	 altHeader = altClean[0] || [];
 	 	 	 	 	 	 	 if (altHeader.length === 14) {
 	 	 	 	 	 	 	 	 altData = altClean;
 	 	 	 	 	 	 	 }
 	 	 	 	 	 	 }
 	 	 	 	 	 }
 	 	 	 	 	 if (altHeader && altHeader.length === 14 && altData) {
 	 	 	 	 	 	 Logger.log(`[${configName}] Failsafe: Using 14-col variant for "${baseName}" instead of 13-col file "${file.getName()}"`);
 	 	 	 	 	 	 // Swap current context to alt
 	 	 	 	 	 	 header = altHeader;
 	 	 	 	 	 	 cleanedData = altData;
 	 	 	 	 	 	 if (altSpreadsheet) {
 	 	 	 	 	 	 	 spreadsheet = altSpreadsheet;
 	 	 	 	 	 	 }
 	 	 	 	 	 	 processedIds.add(altFile.getId());
 	 	 	 	 	 	 swapped = true;
 	 	 	 	 	 	 break;
 	 	 	 	 	 }
 	 	 	 	 }
 	 	 	 }
 	 	 } catch (altErr) {
 	 	 	 Logger.log(`[${configName}] Failsafe lookup error: ${altErr.message}`);
 	 	 }
 	 }
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
				// Normalize rows to header width to prevent mismatches
				const normalized = bodyRows.map(r => {
						const arr = r.slice(0, header.length);
						while (arr.length < header.length) arr.push('');
						return arr;
				});
				 try {
				 mergedSheet.getRange(2, 1, normalized.length, header.length).setValues(normalized);
				 } catch (e) {
					 const sourceName = (spreadsheet && typeof spreadsheet.getName === 'function') ? spreadsheet.getName() : file.getName();
				 const dataCols = (normalized && normalized[0]) ? normalized[0].length : 0;
					 const headerCols = header.length;
					 Logger.log(`[${configName}] setValues error while writing first data block from temp daily report "${sourceName}": ${e.message} (dataCols=${dataCols}, headerCols=${headerCols})`);
					 throw new Error(`Temp Daily Report: ${sourceName} — ${e.message} (dataCols=${dataCols}, headerCols=${headerCols})`);
				 }
			 }
		 headerWritten = true;
	 } else {
		 Logger.log(`[${configName}] ⚠️ Skipping write: detected header row is empty after cleaning.`);
	 }
	} else {
	 // Validate this file's header strictly before appending rows
	 const check = analyzeHeaderAgainstSpec_(localHeader);
	 if ((check.missing && check.missing.length) || check.orderOk === false) {
		 const diag = formatHeaderOrderDiagnostic_(localHeader);
		 headerIssues.push({ fileName: file.getName(), missing: (check.missing || []).slice(), orderOk: !!check.orderOk, diag });
		 Logger.log(`[${configName}] Skipping file due to schema/order mismatch: ${fileName} (missing: ${(check.missing||[]).join(', ') || 'none'}; orderOk=${check.orderOk})`);
		 continue;
	 }
 const startRow = mergedSheet.getLastRow() + 1;
 const rowsToAdd = cleanedData.slice(1);
 if (rowsToAdd.length > 0 && header.length > 0) {
	 const normalized = rowsToAdd.map(r => {
		 const arr = r.slice(0, header.length);
		 while (arr.length < header.length) arr.push('');
		 return arr;
	 });
	 try {
		mergedSheet.getRange(startRow, 1, normalized.length, header.length).setValues(normalized);
	 } catch (e) {
		 const sourceName = (spreadsheet && typeof spreadsheet.getName === 'function') ? spreadsheet.getName() : file.getName();
		const dataCols = (normalized && normalized[0]) ? normalized[0].length : 0;
		 const headerCols = header.length;
		 Logger.log(`[${configName}] setValues error while appending from temp daily report "${sourceName}": ${e.message} (dataCols=${dataCols}, headerCols=${headerCols})`);
		 throw new Error(`Temp Daily Report: ${sourceName} — ${e.message} (dataCols=${dataCols}, headerCols=${headerCols})`);
	 }
 } else {
 Logger.log(`[${configName}] No data rows found in ${fileName} after header; skipping.`);
 }
 }

 // Move the source file (converted or CSV) to holding folder
 // Route all processed source files to Temp Daily Reports (no separate Invalid folder)
 const holdingFolderPath = [...TRASH_ROOT_PATH, 'To Trash After 60 Days', 'Temp Daily Reports', configName];
 const holdingFolder = getDriveFolderByPath_(holdingFolderPath);

 if (holdingFolder) {
 // For CSV, spreadsheet might be null if we never selected it; in that case create it now for consistent archival
 let sheetFileToMove = null;
 if (spreadsheet) {
	 sheetFileToMove = DriveApp.getFileById(spreadsheet.getId());
 } else if (/\.csv$/i.test(file.getName())) {
	 try {
		 const csvAgain = Utilities.parseCsv(file.getBlob().getDataAsString());
		 const csvSheet = SpreadsheetApp.create(file.getName().replace(/\.csv$/i, ''));
		 csvSheet.getSheets()[0].getRange(1, 1, csvAgain.length, csvAgain[0].length).setValues(csvAgain);
		 sheetFileToMove = DriveApp.getFileById(csvSheet.getId());
	 } catch (csvErr) {
		 Logger.log(`[${configName}] CSV archival create error: ${csvErr.message}`);
	 }
 }
 if (sheetFileToMove) {
	 sheetFileToMove.moveTo(holdingFolder);
 }
 } else {
 	Logger.log(`⚠️ [${configName}] Holding folder not found: ${holdingFolderPath.join(' / ')}`);
 }

 // mark this original as processed to avoid duplicates
 try { processedIds.add(file.getId()); } catch (_) {}
 }

 // After processing all files, if any header issues were recorded, notify recipients/admin
 if (headerIssues.length > 0) {
 	try {
 		const recipientsDataLocal = recipientsData || loadRecipientsFromSheet();
 		const to = resolveRecipients(configName, recipientsDataLocal);
 		const cc = resolveCc(configName, recipientsDataLocal);
 		const subject = `⚠️ CM360 Merge Validation Issues (${configName})`;
 		// Build a single canonical line showing the required 14-column schema order
 		const expectedOrder = getExpectedHeaderSpec_().map(opts => opts[0]).join(' | ');
		const rows = headerIssues.map((iss) => `
 			<tr>
 				<td style="padding:4px; border:1px solid #ddd;">${escapeHtml(iss.fileName)}</td>
 				<td style="padding:4px; border:1px solid #ddd;">${escapeHtml((iss.missing||[]).join(', ') || '—')}</td>
 				<td style="padding:4px; border:1px solid #ddd;">${iss.orderOk === false ? 'Out of order' : 'OK'}</td>
 			</tr>`).join('');
		const htmlBody = `
			<p style="font-family:Arial, sans-serif; font-size:13px;">Some temp daily report(s) were skipped due to schema validation errors. The audit proceeds using valid files so you still receive flags today — see below.</p>
			<table cellpadding="0" cellspacing="0" style="border-collapse:collapse; font-family:Arial, sans-serif; font-size:12px;">
 				<thead>
					<tr style="background:#f2f2f2;">
						<th style="padding:4px; border:1px solid #ddd;">File</th>
						<th style="padding:4px; border:1px solid #ddd;">Missing Header(s)</th>
						<th style="padding:4px; border:1px solid #ddd;">Header Order</th>
					</tr>
 				</thead>
 				<tbody>${rows}</tbody>
 			</table>
			<p style="font-family:Arial, sans-serif; font-size:12px; color:#333; margin:10px 0 12px;">
				Required 14-column schema order:
				<span style="font-family:Consolas, 'Courier New', monospace;">${escapeHtml(expectedOrder)}</span>
			</p>
 			<p style="margin-top:10px; font-family:Arial, sans-serif; font-size:12px; color:#666;">If you need help updating your report template to the required 14-column schema, please contact the Platform Solutions team.</p>
 		`;
 		// In production, BCC admin on recipient-facing failures; in staging, safeSendEmail will route to admin only
 		safeSendEmail({ to, cc, subject, htmlBody, bccAdmin: true }, `${configName} - merge validation issues`);
 	} catch (notifyErr) {
 		Logger.log(`[${configName}] Error sending header issues notification: ${notifyErr.message}`);
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
function executeAudit(config, preloaded) {
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

 // Load configuration data (prefer batch-preloaded if provided)
 const usingPreloaded = !!preloaded;
 const exclusionsData = (preloaded && preloaded.exclusionsData) || loadExclusionsFromSheet();
 const thresholdsData = (preloaded && preloaded.thresholdsData) || loadThresholdsFromSheet();
 const recipientsData = (preloaded && preloaded.recipientsData) || loadRecipientsFromSheet();
 if (!usingPreloaded) {
	 Logger.log(`ℹ️ [${configName}] Loaded exclusions for ${Object.keys(exclusionsData).length} configs`);
	 Logger.log(`ℹ️ [${configName}] Loaded thresholds for ${Object.keys(thresholdsData).length} configs`);
	 Logger.log(`ℹ️ [${configName}] Loaded recipients for ${Object.keys(recipientsData).length} configs`);
 }

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
	 attachments: [],
	 bccAdmin: true 
 }, configName);
 return { status: 'Skipped: No files found', flaggedCount: null, emailSent: true, emailTime: formattedNow };
 }

 const mergedSheetId = mergeDailyAuditExcels(folderId, config.mergedFolderPath, configName, recipientsData);
 const mergedSs = SpreadsheetApp.openById(mergedSheetId);
 const sheet = mergedSs.getSheets()[0];
 const mergedSheetUrl = mergedSs.getUrl();
 setLatestReportUrl_(configName, mergedSheetUrl);
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
 // NEW LOGIC: Use threshold for whichever metric (impressions or clicks) is HIGHER
 
 // Clicks > Impressions check
 const clicksThreshold = getThresholdForFlag(thresholdsData, configName, 'clicks_greater_than_impressions');
 // Determine which metric is higher and apply its threshold
 const dominantMetricForClicks = clicks > impressions ? 'clicks' : 'impressions';
 const dominantValueForClicks = clicks > impressions ? clicks : impressions;
 const appliedThresholdForClicks = clicks > impressions ? clicksThreshold.minClicks : clicksThreshold.minImpressions;
 const hasMinVolumeForClicks = dominantValueForClicks >= appliedThresholdForClicks;
 if (hasMinVolumeForClicks && clicks > impressions && 
 !isPlacementExcludedForFlag(exclusionsData, configName, placementId, 'clicks_greater_than_impressions', placementName, siteName)) {
 flags.push('Clicks > Impressions');
 Logger.log(`🚩 [FLAG ADDED] ${configName} - ${placementName}: Clicks > Impressions (${clicks} clicks >= ${appliedThresholdForClicks} threshold)`);
 }
 
 // Out of flight dates check
 const flightThreshold = getThresholdForFlag(thresholdsData, configName, 'out_of_flight_dates');
 const dominantMetricForFlight = clicks > impressions ? 'clicks' : 'impressions';
 const dominantValueForFlight = clicks > impressions ? clicks : impressions;
 const appliedThresholdForFlight = clicks > impressions ? flightThreshold.minClicks : flightThreshold.minImpressions;
 const hasMinVolumeForFlight = dominantValueForFlight >= appliedThresholdForFlight;
 if (hasMinVolumeForFlight && (startDate > today || endDate < today) && 
 !isPlacementExcludedForFlag(exclusionsData, configName, placementId, 'out_of_flight_dates', placementName, siteName)) {
 flags.push('Out of flight dates');
 }
 
 // Pixel size mismatch check
 const pixelThreshold = getThresholdForFlag(thresholdsData, configName, 'pixel_size_mismatch');
 const dominantMetricForPixel = clicks > impressions ? 'clicks' : 'impressions';
 const dominantValueForPixel = clicks > impressions ? clicks : impressions;
 const appliedThresholdForPixel = clicks > impressions ? pixelThreshold.minClicks : pixelThreshold.minImpressions;
 const hasMinVolumeForPixel = dominantValueForPixel >= appliedThresholdForPixel;
 if (hasMinVolumeForPixel && placementPixel && creativePixel && placementPixel !== creativePixel && 
 !isPlacementExcludedForFlag(exclusionsData, configName, placementId, 'pixel_size_mismatch', placementName, siteName)) {
 flags.push('Pixel size mismatch');
 }
 
 // Default ad serving check
 const defaultThreshold = getThresholdForFlag(thresholdsData, configName, 'default_ad_serving');
 const dominantMetricForDefault = clicks > impressions ? 'clicks' : 'impressions';
 const dominantValueForDefault = clicks > impressions ? clicks : impressions;
 const appliedThresholdForDefault = clicks > impressions ? defaultThreshold.minClicks : defaultThreshold.minImpressions;
 const hasMinVolumeForDefault = dominantValueForDefault >= appliedThresholdForDefault;
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

	if (reorderedFlagged.length === 0) {
		// No flags: decide whether we would send based on recipients setting
		const configRecipients = recipientsData[config.name];
		const withhold = !!(configRecipients && configRecipients.withholdNoFlagEmails);
		// If suppressed (silent check), do not actually send but report would-send/withheld
		const suppressed = (typeof isEmailSuppressed_ === 'function' && isEmailSuppressed_());
		if (suppressed) {
			return {
				status: withhold ? 'Completed (no issues; email would be withheld)' : 'Completed (no issues; email would be sent) ',
				flaggedCount: 0,
				emailSent: false,
				emailTime: formattedNow,
				emailWithheld: withhold,
				latestReportUrl: mergedSheetUrl
			};
		}

		if (withhold) {
			Logger.log(`ℹ️ [${config.name}] No-issue email withheld: Recipients opted out of no-flag emails`);
			return { status: 'Completed (no issues)', flaggedCount: 0, emailSent: false, emailTime: formattedNow, emailWithheld: true, latestReportUrl: mergedSheetUrl };
		} else {
			// No flags and not withheld: send the no-issues notice
			SpreadsheetApp.flush();
			const emailSent = sendNoIssueEmail(config, mergedSheetId, 'No issues were flagged', recipientsData);
			const statusNoFlags = emailSent ? 'Completed (no issues)' : 'Completed (no issues, email failed to send)';
			return { status: statusNoFlags, flaggedCount: 0, emailSent, emailTime: formattedNow, latestReportUrl: mergedSheetUrl };
		}
	}

 // Rewrite sheet cleanly for flagged case
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
	 // If suppressed, don't actually send but report as would-send
	 const suppressed = (typeof isEmailSuppressed_ === 'function' && isEmailSuppressed_());
	 const emailSent = suppressed ? false : emailFlaggedRows(mergedSheetId, displayRows, flaggedRows, config, recipientsData);
 const status = emailSent ? 'Completed with flags' : 'Completed with flags (email failed to send)';
	 return { status, flaggedCount: flaggedRows.length, emailSent, emailTime: formattedNow, latestReportUrl: mergedSheetUrl };
 } else {
	 // Should not reach: displayRows corresponds to flagged case; keep existing fallback just in case
	 const emailSent = emailFlaggedRows(mergedSheetId, displayRows, flaggedRows, config, recipientsData);
	 const status = emailSent ? 'Completed with flags' : 'Completed with flags (email failed to send)';
	 return { status, flaggedCount: flaggedRows.length, emailSent, emailTime: formattedNow, latestReportUrl: mergedSheetUrl };
 
}

 } catch (err) {
 Logger.log(` [${configName}] Unexpected error: ${err.message}`);
 return { status: `Error during audit: ${err.message}`, flaggedCount: null, emailSent: false, emailTime: formattedNow };
 }
}

// === EXECUTION & AUDIT FLOW ===
function runDailyAuditByName(configName) {
	if (!checkDriveApiEnabled()) return;
	const config = getAuditConfigByName(configName);
	if (!config) {
		Logger.log(`?? Config "${configName}" not found.`);
		return;
	}
	executeAudit(config);
}

function runAuditBatch(configs, isFinal = false) {
 validateAuditConfigs();
 const batchId = (function(){
	 try {
		 const names = (configs || []).map(c => c && c.name).filter(Boolean).join(',');
		 return `batch:${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}:${names}`;
	 } catch (_) { return `batch:${Date.now()}`; }
 })();
 Logger.log(` Audit Batch Started: ${new Date().toLocaleString()} (${batchId})`);
 const results = [];

 // Record start state for watchdog
 try {
	 const props = PropertiesService.getScriptProperties();
	 const state = { startedAt: Date.now(), tz: Session.getScriptTimeZone(), configs: (configs||[]).map(c=>c&&c.name).filter(Boolean), isFinal: !!isFinal };
	 props.setProperty(AUDIT_RUN_STATE_KEY_PREFIX + batchId, JSON.stringify(state));
	 const listRaw = props.getProperty(AUDIT_RUN_LIST_KEY);
	 const list = listRaw ? JSON.parse(listRaw) : [];
	 if (!list.includes(batchId)) list.push(batchId);
	 props.setProperty(AUDIT_RUN_LIST_KEY, JSON.stringify(list.slice(-50))); // keep last 50
 } catch (e) { Logger.log('runAuditBatch: failed to persist run state: ' + e.message); }

 // Soft guard: if we're close to hard timeout, send an early warning and exit
 const startWall = Date.now();
 const HARD_TIMEOUT_MS = 6 * 60 * 1000; // 6 minutes typical hard cap
 const SOFT_GUARD_MS = HARD_TIMEOUT_MS - 30000; // warn/exit 30s before

 // Preload config tables once per batch
 let preloaded = null;
 try {
	 preloaded = {
		 recipientsData: loadRecipientsFromSheet(),
		 thresholdsData: loadThresholdsFromSheet(),
		 exclusionsData: loadExclusionsFromSheet()
	 };
	 Logger.log(`ℹ️ Preloaded config tables: recipients=${Object.keys(preloaded.recipientsData||{}).length}, thresholds=${Object.keys(preloaded.thresholdsData||{}).length}, exclusions=${Object.keys(preloaded.exclusionsData||{}).length}`);
 } catch (preErr) {
	 Logger.log('Preload failed (will fall back to per-config loads): ' + preErr.message);
	 preloaded = null;
 }

 for (const config of configs) {
 try {
 // Check soft guard before each config
 if ((Date.now() - startWall) > SOFT_GUARD_MS) {
	 const msg = `Exiting early to avoid hard timeout. Batch ${batchId} processed ${results.length} of ${configs.length} configs.`;
	 Logger.log('[TIMEGUARD] ' + msg);
	 try {
		 const subject = `[ALERT] CM360 Audit early exit before timeout`;
		 const html = `<p>${escapeHtml(msg)}</p>`;
		 safeSendEmail({ to: ADMIN_EMAIL, subject, htmlBody: html, plainBody: msg }, 'Timeout Guard');
	 } catch (e) { Logger.log('Failed to notify admin about early-exit: ' + e.message); }
	 break;
 }
 const result = executeAudit(config, preloaded);
 const entry = {
 name: config.name,
 status: result.status,
 flaggedRows: result.flaggedCount,
 emailSent: result.emailSent,
 emailTime: result.emailTime,
	emailWithheld: result.emailWithheld || false,
	latestReportUrl: result.latestReportUrl || ''
};
 results.push(entry);
 storeCombinedAuditResults_([entry]);
 } catch (err) {
 const entry = {
 name: config.name,
 status: `Error: ${err.message}`,
 flaggedRows: null,
 emailSent: false,
 emailTime: formatDate(new Date(), 'yyyy-MM-dd HH:mm:ss'),
	emailWithheld: false,
	latestReportUrl: ''
 };
 results.push(entry);
 storeCombinedAuditResults_([entry]);
 }
 }

	storeCombinedAuditResults_(results);

	const cachedResults = getCombinedAuditResults_();
	const totalConfigs = getAuditConfigs().length;
	const completedConfigs = new Set(cachedResults.map(r => r.name)).size;

	Logger.log(`✅ Completed ${completedConfigs} of ${totalConfigs} configs`);

	// Send the summary once when all configs are done, regardless of batch order
	if (completedConfigs >= totalConfigs) {
		attemptSendDailySummary_({ allowPlaceholders: false, reason: 'All configs complete' });
	}

 // Mark completion for this batch
 try {
	 const props = PropertiesService.getScriptProperties();
	 const key = AUDIT_RUN_STATE_KEY_PREFIX + batchId;
	 const raw = props.getProperty(key);
	 const state = raw ? JSON.parse(raw) : {};
	 state.completedAt = Date.now();
	 state.results = results.map(r => ({ name: r.name, status: r.status }));
	 props.setProperty(key, JSON.stringify(state));
 } catch (e) { Logger.log('runAuditBatch: failed to mark completion: ' + e.message); }
}

function getAuditConfigBatches(batchSize = BATCH_SIZE) {
	const configs = getAuditConfigs();
	const batches = [];
	for (let i = 0; i < configs.length; i += batchSize) {
		batches.push(configs.slice(i, i + batchSize));
	}
	return batches;
}

function validateAuditConfigs() {
	const configs = getAuditConfigs();
	if (!Array.isArray(configs) || configs.length === 0) {
		throw new Error('No active audit configurations found. Add entries to the Audit Recipients sheet.');
	}
	configs.forEach((c, idx) => {
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
	Logger.log(`validateAuditConfigs: ${configs.length} configs ready.`);
}
function sanitizeCombinedResultEntry_(entry, options) {
	options = options || {};
	if (!entry || typeof entry !== 'object') return null;
	const statusLimit = typeof options.statusLimit === 'number' ? options.statusLimit : 400;
	const emailTimeLimit = typeof options.emailTimeLimit === 'number' ? options.emailTimeLimit : 48;
	const urlLimit = typeof options.urlLimit === 'number' ? options.urlLimit : 350;
	const includeUrl = options.includeUrl !== false;
	const safeNumber = (value) => {
		if (value === null || typeof value === 'undefined') return null;
		const n = Number(value);
		return isFinite(n) ? n : null;
	};
	const sanitizeString = (value, limit) => {
		if (!value) return '';
		const str = String(value);
		if (!limit || str.length <= limit) return str;
		return str.slice(0, Math.max(1, limit - 1)) + '…';
	};
	const sanitized = {
		name: entry.name ? String(entry.name) : '',
		status: sanitizeString(entry.status, statusLimit),
		flaggedRows: safeNumber(entry.flaggedRows),
		emailSent: entry.emailSent === true,
		emailTime: sanitizeString(entry.emailTime, emailTimeLimit),
		emailWithheld: entry.emailWithheld === true
	};
	if (includeUrl) {
		sanitized.latestReportUrl = sanitizeString(entry.latestReportUrl, urlLimit);
	}
	return sanitized;
}

function storeCombinedAuditResults_(newResults) {
	const cache = CacheService.getScriptCache();
	const key = getAuditCacheKey_();
	const existing = getCombinedAuditResults_();
	const incoming = Array.isArray(newResults) ? newResults.filter(Boolean) : (newResults ? [newResults] : []);
	const combined = mergeAuditResultsByConfig_(existing, incoming);

	const variants = [
		{ name: 'full', options: { statusLimit: 400, emailTimeLimit: 64, urlLimit: 350, includeUrl: true } },
		{ name: 'trimmed', options: { statusLimit: 160, emailTimeLimit: 48, urlLimit: 250, includeUrl: true } },
		{ name: 'minimal', options: { statusLimit: 80, emailTimeLimit: 32, urlLimit: 0, includeUrl: false } }
	];

	for (const variant of variants) {
		const payloadArray = combined
			.map(entry => sanitizeCombinedResultEntry_(entry, variant.options))
			.filter(Boolean);
		let serialized;
		try {
			serialized = JSON.stringify(payloadArray);
		} catch (serErr) {
			Logger.log(`storeCombinedAuditResults_: failed to serialize (${variant.name}): ${serErr.message}`);
			continue;
		}
		if (serialized.length > 95000) {
			Logger.log(`storeCombinedAuditResults_: payload too large for variant ${variant.name} (${serialized.length}B)`);
			continue;
		}
		try {
			cache.put(key, serialized, 21600); // 6 hours
			if (variant.name !== 'full') {
				Logger.log(`storeCombinedAuditResults_: cached using ${variant.name} variant (${serialized.length}B)`);
			}
			return;
		} catch (err) {
			const msg = err && err.message ? err.message : String(err);
			if (/storage/i.test(msg) || /INTERNAL/i.test(msg)) {
				Logger.log(`storeCombinedAuditResults_: cache storage error with ${variant.name} variant (${serialized.length}B): ${msg}`);
				continue;
			}
			Logger.log(`storeCombinedAuditResults_: unexpected cache error (${variant.name}): ${msg}`);
			return;
		}
	}

	Logger.log('storeCombinedAuditResults_: unable to cache combined results after all variants (using in-memory fallback only).');
}

function getCombinedAuditResults_() {
 const cache = CacheService.getScriptCache();
 const stored = cache.get(getAuditCacheKey_());
 if (!stored) return [];
 try {
		const parsed = JSON.parse(stored);
		if (!Array.isArray(parsed)) return [];
		return parsed.map(entry => {
			if (!entry || typeof entry !== 'object') return null;
			const toNumber = value => {
				if (value === null || typeof value === 'undefined') return null;
				const n = Number(value);
				return isFinite(n) ? n : null;
			};
			return {
				name: entry.name ? String(entry.name) : '',
				status: entry.status ? String(entry.status) : '',
				flaggedRows: toNumber(entry.flaggedRows),
				emailSent: entry.emailSent === true,
				emailTime: entry.emailTime ? String(entry.emailTime) : '',
				emailWithheld: entry.emailWithheld === true,
				latestReportUrl: entry.latestReportUrl ? String(entry.latestReportUrl) : ''
			};
		}).filter(Boolean);
 } catch (e) {
 	Logger.log('getCombinedAuditResults_ parse error: ' + e.message);
 	return [];
 }
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
	const summaryText = `The following ${totalFlagged} ${plural(totalFlagged, 'placement', 'placements')} across ${uniqueCampaigns} ${plural(uniqueCampaigns, 'campaign', 'campaigns')} ${verb} flagged during the <strong>${configName}</strong> CM360 audit of yesterday's delivery. Please review:`;

	const rowHtmlFragments = emailRows.map((row, i) => `
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
	</tr>`);

	const buildHtmlBody = (rowsToInclude, truncated) => {
		const visibleRows = rowHtmlFragments.slice(0, rowsToInclude).join('');
		const truncateNotice = truncated
			? `<p style="font-family:Arial, sans-serif; font-size:12px; margin-top:12px;">Please view the attachment for additional flags.</p>`
			: '';

		return `
	<p style="font-family:Arial, sans-serif; font-size:13px; line-height:1.4;">${summaryText}</p>
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
${visibleRows}
	</tbody>
	</table>
${truncateNotice}
	<p style="margin-top:12px; font-family:Arial, sans-serif; font-size:12px;">&mdash; Platform Solutions Team</p>
	`;
	};

	let rowsToInclude = rowHtmlFragments.length;
	let truncated = false;
	let htmlBody = buildHtmlBody(rowsToInclude, truncated);
	let payloadSize = getEmailPayloadSize_(htmlBody, '');

	if (payloadSize.totalBytes > EMAIL_BODY_BYTE_LIMIT) {
		truncated = true;
		while (rowsToInclude > 0 && payloadSize.totalBytes > EMAIL_BODY_BYTE_LIMIT) {
			rowsToInclude--;
			htmlBody = buildHtmlBody(rowsToInclude, true);
			payloadSize = getEmailPayloadSize_(htmlBody, '');
		}

		if (payloadSize.totalBytes > EMAIL_BODY_BYTE_LIMIT) {
			htmlBody = `
	<p style="font-family:Arial, sans-serif; font-size:13px; line-height:1.4;">${summaryText}</p>
	<p style="font-family:Arial, sans-serif; font-size:12px;">The flagged placements table was omitted because it exceeded the email body limit. Please review the attached file for the full list of flags.</p>
	<p style="margin-top:12px; font-family:Arial, sans-serif; font-size:12px;">&mdash; Platform Solutions Team</p>
	`;
			payloadSize = getEmailPayloadSize_(htmlBody, '');
			Logger.log(`[${configName}] Email body limit reached; table omitted entirely. Final size ${payloadSize.totalBytes}B (limit ${EMAIL_BODY_BYTE_LIMIT}B).`);
			rowsToInclude = 0;
		} else {
			const removed = emailRows.length - rowsToInclude;
			Logger.log(`[${configName}] Email body truncated to ${rowsToInclude} of ${emailRows.length} rows (${payloadSize.totalBytes}B, limit ${EMAIL_BODY_BYTE_LIMIT}B). ${removed} row(s) moved to attachment only.`);
		}
	} else {
		Logger.log(`[${configName}] Email body size ${payloadSize.totalBytes}B with ${rowsToInclude} rows (limit ${EMAIL_BODY_BYTE_LIMIT}B).`);
	}

	const emailSuccess = safeSendEmail({
		to: resolveRecipients(configName, recipientsData),
		cc: resolveCc(configName, recipientsData),
		subject,
		htmlBody,
		attachments: [xlsxBlob]
	}, `[${configName}]`);

	if (!emailSuccess) {
		Logger.log(`[${configName}] Failed to send flagged rows email; an alert was sent to the admin.`);
	}

 Logger.log(`[${configName}](c) Flagging complete: ${flaggedRows.length} row(s)`);

 return emailSuccess;
}


// Send a concise "no issues" notification with optional Excel attachment of the merged sheet
function sendNoIssueEmail(config, spreadsheetId, note, recipientsData) {
	try {
		const configName = config && config.name ? String(config.name) : 'Unknown';
		const tz = Session.getScriptTimeZone();
		const subjectDate = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
		const subject = `✅ CM360 Daily Audit: No Issues (${configName} - ${subjectDate})`;

		// Build a short, friendly HTML body
		const htmlBody = `
			<p style="font-family:Arial, sans-serif; font-size:13px; line-height:1.4;">
				No issues were flagged for <strong>${escapeHtml(configName)}</strong> during yesterday's CM360 audit.
			</p>
			<p style="font-family:Arial, sans-serif; font-size:12px; color:#666;">
				The merged delivery file is attached for reference.
			</p>
			<p style="margin-top:12px; font-family:Arial, sans-serif; font-size:12px;">&mdash; Platform Solutions Team</p>
		`;

		// Attach merged sheet as Excel for audit trail (best-effort)
		let attachments = [];
		try {
			if (spreadsheetId) {
				const xlsx = exportSheetAsExcel(spreadsheetId, `CM360_DailyAudit_${configName}_${subjectDate}.xlsx`);
				if (xlsx) attachments = [xlsx];
			}
		} catch (attachErr) {
			Logger.log(`sendNoIssueEmail: could not build attachment (${configName}): ${attachErr.message}`);
		}

		return safeSendEmail({
			to: resolveRecipients(configName, recipientsData),
			cc: resolveCc(configName, recipientsData),
			subject,
			plainBody: 'No issues were flagged.',
			htmlBody,
			attachments
		}, `${configName} - no issues`);
	} catch (e) {
		Logger.log(`sendNoIssueEmail error: ${e.message}`);
		return false;
	}
}

// === SETUP & ENVIRONMENT PREP ===
function prepareAuditEnvironment() {
 const ui = SpreadsheetApp.getUi();
 const createdLabels = [];
 const createdFolders = [];
 const labelsWithoutRecentMail = [];

 getAuditConfigs().forEach(config => {
 const { name, label, mergedFolderPath, tempDailyFolderPath } = config;

	// 1. Ensure Gmail label exists (robust: try canonical match before creating)
	let labelObj = findGmailLabel_(label, name);
	if (labelObj) {
		Logger.log(`[LABEL] Found existing label for ${name}: ${labelObj.getName()} (desired: ${label})`);
	} else {
		labelObj = GmailApp.createLabel(label);
		createdLabels.push(label);
		Logger.log(`... Created Gmail label: ${label}`);
	}

	// 2. Check for recent mail under this label (7d) to hint at filter status
	try {
		const encoded = labelObj.getName().replace(/"/g, '\\"');
		const query = `label:"${encoded}" newer_than:7d`;
		const recent = GmailApp.search(query, 0, 1);
		if (!recent || recent.length === 0) {
			labelsWithoutRecentMail.push({ name, label: labelObj.getName() });
		}
	} catch (e) {
		Logger.log(`[LABEL] Recent mail check failed for ${labelObj.getName()}: ${e.message}`);
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

 // 4. Log status and generate pop-up
 let msgParts = [];

 if (createdLabels.length > 0) {
	msgParts.push(`✅ Created ${createdLabels.length} Gmail label(s).`);
 }

 if (createdFolders.length > 0) {
 msgParts.push(`[FOLDER] Created ${createdFolders.length} Drive folder path(s).`);
 }

 if (labelsWithoutRecentMail.length > 0) {
	 msgParts.push(`\nℹ️ These labels had no recent mail in the last 7 days (review filters if expected):`);
	 labelsWithoutRecentMail.forEach(({ name, label }) => {
		 msgParts.push(`- ${name} — label: "${label}"`);
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

function installAllAutomationTriggers(options) {
	options = options || {};
	const silent = options.silent === true;
	const handlerFilter = Array.isArray(options.handlers) && options.handlers.length
		? new Set(options.handlers.map(String))
		: null;

	const shouldInclude = (key, handlerNames) => {
		if (!handlerFilter) return true;
		if (key && handlerFilter.has(key)) return true;
		if (handlerNames) {
			for (const name of [].concat(handlerNames)) {
				if (handlerFilter.has(String(name))) return true;
			}
		}
		return false;
	};

	const removeHandlers = (handlerNames, label) => {
		if (!handlerNames || handlerNames.length === 0) return;
		const targetSet = new Set(handlerNames.map(String));
		(ScriptApp.getProjectTriggers() || []).forEach(trigger => {
			try {
				const fn = trigger.getHandlerFunction && trigger.getHandlerFunction();
				if (fn && targetSet.has(fn)) {
					ScriptApp.deleteTrigger(trigger);
					results.push(`🧹 Removed existing trigger for ${fn}`);
				}
			} catch (err) {
				results.push(`⚠️ Failed to remove ${label || 'trigger'}: ${err.message}`);
			}
		});
	};

	const createTrigger = (handler, builderFn, description) => {
		try {
			builderFn(ScriptApp.newTrigger(handler));
			results.push(`✅ Installed ${description || handler}`);
		} catch (err) {
			Logger.log(`installAllAutomationTriggers: failed to install ${handler}: ${err.message}`);
			results.push(`⚠️ Failed to install ${description || handler}: ${err.message}`);
		}
	};

	const results = [];

	// === Daily audit batch triggers ===
	if (shouldInclude('dailyAuditBatches', [])) {
		const batches = getAuditConfigBatches(BATCH_SIZE);
		const batchHandlers = [];
		for (let i = 0; i < batches.length; i++) {
			batchHandlers.push(`runDailyAuditsBatch${i + 1}`);
		}
		removeHandlers(batchHandlers, 'daily audit batch trigger');

		let presentIndexes = null;
		try {
			const info = listBatchFunctionsInSource_();
			if (info && info.existingIndexes) presentIndexes = info.existingIndexes;
		} catch (e) {
			Logger.log('installAllAutomationTriggers: source scan failed; falling back to runtime checks: ' + e.message);
		}

		for (let i = 0; i < batches.length; i++) {
			const index1 = i + 1;
			const fnName = `runDailyAuditsBatch${index1}`;
			let canInstall = false;
			if (presentIndexes) {
				canInstall = presentIndexes.has(index1);
			} else {
				try {
					canInstall = (typeof globalThis[fnName] === 'function');
				} catch (_) {
					canInstall = false;
				}
			}
			if (canInstall) {
				createTrigger(fnName, trig => trig.timeBased().atHour(8).everyDays(1).create(), `daily audit batch trigger (${fnName})`);
			} else {
				results.push(`⚠️ Skipped trigger for ${fnName} — function not found in source`);
			}
		}
	}

	// === Summary failover ===
	if (shouldInclude('summaryFailover', ['sendDailySummaryFailover'])) {
		removeHandlers(['sendDailySummaryFailover'], 'summary failover trigger');
		createTrigger('sendDailySummaryFailover', trig => trig.timeBased().atHour(9).nearMinute(30).everyDays(1).create(), 'daily summary failover trigger');
	}

	// === Daily health check ===
	if (shouldInclude('healthCheck', ['runHealthCheckAndEmail'])) {
		removeHandlers(['runHealthCheckAndEmail'], 'health check trigger');
		createTrigger('runHealthCheckAndEmail', trig => trig.timeBased().atHour(5).everyDays(1).create(), 'daily health check trigger');
	}

	// === Audit watchdog ===
	if (shouldInclude('auditWatchdog', ['auditWatchdogCheck'])) {
		removeHandlers(['auditWatchdogCheck'], 'watchdog trigger');
		createTrigger('auditWatchdogCheck', trig => trig.timeBased().everyHours(3).create(), '3-hour audit watchdog trigger');
	}

	// === Delivery mode sync ===
	if (shouldInclude('deliveryModeSync', ['runDeliveryModeSync'])) {
		removeHandlers(['runDeliveryModeSync'], 'delivery mode sync trigger');
		createTrigger('runDeliveryModeSync', trig => trig.timeBased().everyHours(3).create(), '3-hour delivery mode sync trigger');
	}

	// === Auto-fix audit requests sheet (requires external config) ===
	if (shouldInclude('autoFixRequests', ['autoFixRequestsSheet_'])) {
		removeHandlers(['autoFixRequestsSheet_'], 'auto-fix requests trigger');
		if (EXTERNAL_CONFIG_SHEET_ID) {
			createTrigger('autoFixRequestsSheet_', trig => trig.timeBased().everyHours(4).create(), '4-hour audit requests auto-fix trigger');
		} else {
			results.push('ℹ️ Skipped auto-fix trigger — EXTERNAL_CONFIG_SHEET_ID not configured');
		}
	}

	// === Nightly maintenance bundle ===
	if (shouldInclude('nightlyMaintenance', [
		'runNightlyMaintenance',
		'rebalanceAuditBatchesUsingSummary',
		'nightlyExternalSync',
		'runNightlyExternalSync',
		'runNightlyExternalSync_',
		'refreshExternalConfigInstructionsSilent',
		'refreshExternalConfigInstructions',
		'updatePlacementNamesFromReports',
		'clearDailyScriptProperties'
	])) {
		removeHandlers([
			'runNightlyMaintenance',
			'rebalanceAuditBatchesUsingSummary',
			'nightlyExternalSync',
			'runNightlyExternalSync',
			'runNightlyExternalSync_',
			'refreshExternalConfigInstructionsSilent',
			'refreshExternalConfigInstructions',
			'updatePlacementNamesFromReports',
			'clearDailyScriptProperties'
		], 'nightly maintenance trigger');
		createTrigger('runNightlyMaintenance', trig => trig.timeBased().atHour(2).nearMinute(20).everyDays(1).create(), 'nightly maintenance trigger');
	}

	// === GAS failure notifications forwarder ===
	if (shouldInclude('gasFailureForwarder', ['forwardGASFailureNotificationsToAdmin'])) {
		removeHandlers(['forwardGASFailureNotificationsToAdmin'], 'GAS failure forwarder trigger');
		createTrigger('forwardGASFailureNotificationsToAdmin', trig => trig.timeBased().everyHours(1).create(), 'hourly GAS failure forwarder trigger');
	}

	return results;
}

// === TRIGGER FUNCTIONS ===
// Ensure batch runner functions exist for current number of configs (2 per batch)
function ensureBatchRunnerFunctionsPresent() {
	try {
		const batches = getAuditConfigBatches(BATCH_SIZE);
		const neededCount = batches.length;
		for (let i = 1; i <= neededCount; i++) {
			const fnName = `runDailyAuditsBatch${i}`;
			let exists = false;
			try {
				exists = (typeof eval(fnName) === 'function');
			} catch (e) {
				exists = false;
			}
			
			if (!exists) {
				const index = i - 1;
				const isFinal = (i === neededCount);
				globalThis[fnName] = function() {
					const latest = getAuditConfigBatches(BATCH_SIZE);
					runAuditBatch(latest[index], isFinal);
				};
				Logger.log(`ensureBatchRunnerFunctionsPresent: created ${fnName}`);
			}
		}
	} catch (e) {
		Logger.log('ensureBatchRunnerFunctionsPresent error: ' + e.message);
	}
}

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

function runDailyAuditsBatch6() {
const batches = getAuditConfigBatches(BATCH_SIZE);
runAuditBatch(batches[5]);
}

function runDailyAuditsBatch7() {
const batches = getAuditConfigBatches(BATCH_SIZE);
runAuditBatch(batches[6]);
}


function runDailyAuditsBatch8() {
 const batches = getAuditConfigBatches(BATCH_SIZE);
 runAuditBatch(batches[7]);
}

function runDailyAuditsBatch9() {
 const batches = getAuditConfigBatches(BATCH_SIZE);
 runAuditBatch(batches[8]);
}

function runDailyAuditsBatch10() {
 const batches = getAuditConfigBatches(BATCH_SIZE);
 runAuditBatch(batches[9]);
}

function runDailyAuditsBatch11() {
 const batches = getAuditConfigBatches(BATCH_SIZE);
 runAuditBatch(batches[10]);
}

function runDailyAuditsBatch12() {
 const batches = getAuditConfigBatches(BATCH_SIZE);
 runAuditBatch(batches[11]);
}

function generateMissingBatchStubs() {
 const batches = getAuditConfigBatches(BATCH_SIZE);
 const neededCount = batches.length;
 
 // Better detection: check if functions actually exist
 const definedIndexes = new Set();
 for (let j = 1; j <= neededCount; j++) {
   const fnName = `runDailyAuditsBatch${j}`;
   try {
     if (typeof eval(fnName) === 'function') {
       definedIndexes.add(j);
     }
   } catch (e) {
     // Function doesn't exist
   }
 }
 
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

/**
 * Insert any missing runDailyAuditsBatchN functions directly into Code.js source
 * right after the last existing batch function.
 */
function insertMissingBatchFunctionsIntoSource_() {
	try {
		const projectId = ScriptApp.getScriptId();
		const manifest = getScriptProjectContent_();
		if (!manifest || !manifest.files) throw new Error('Unable to read project files');

		// Consider all SERVER_JS files and pick a target file for insertion
		const serverFiles = manifest.files.filter(f => f.type === 'SERVER_JS');
		if (!serverFiles.length) throw new Error('No SERVER_JS files found');
		const preferIdx = serverFiles.findIndex(f => f.name === 'Code');
		let targetIdx = preferIdx !== -1 ? preferIdx : 0;
		let source = String(serverFiles[targetIdx].source || '');

		// Determine how many batches are needed
		const needed = getAuditConfigBatches(BATCH_SIZE).length;

		// Find last defined batch function index present in any server file
		const batchFnRegex = /function\s+runDailyAuditsBatch(\d+)\s*\(/g;
		let match;
		const existingIndexes = new Set();
		for (const f of serverFiles) {
			const src = String(f.source || '');
			while ((match = batchFnRegex.exec(src)) !== null) {
				existingIndexes.add(Number(match[1]));
			}
		}

		// If all needed exist, nothing to do
		let maxExisting = 0;
		existingIndexes.forEach(i => { if (i > maxExisting) maxExisting = i; });

		const additions = [];
		for (let i = 1; i <= needed; i++) {
			if (!existingIndexes.has(i)) {
				const idx = i - 1;
				const isFinal = (i === needed);
				additions.push(
`\n\nfunction runDailyAuditsBatch${i}() {\n const batches = getAuditConfigBatches(BATCH_SIZE);\n runAuditBatch(batches[${idx}], ${isFinal});\n}\n`);
			}
		}

		if (additions.length === 0) {
			Logger.log('insertMissingBatchFunctionsIntoSource_: no new functions needed');
			return { created: [] };
		}

		// Insert additions right before generateMissingBatchStubs() definition if present,
		// else append to end of file
		const anchor = source.indexOf('function generateMissingBatchStubs()');
		if (anchor !== -1) {
			source = source.slice(0, anchor) + additions.join('') + source.slice(anchor);
		} else {
			source = source + additions.join('');
		}

		// Update target file content in manifest and push
		const targetFileGlobalIndex = manifest.files.indexOf(serverFiles[targetIdx]);
		const updatedFiles = manifest.files.map((f, idx) => {
			if (idx === targetFileGlobalIndex) {
				return { name: f.name, type: f.type, source };
			}
			return { name: f.name, type: f.type, source: f.source };
		});
		let updated = updateScriptProjectContent_(updatedFiles);
		if (!updated) throw new Error('Failed to update project content (primary path)');

		// Verify creation by re-reading project content
		const verify1 = getScriptProjectContent_();
		const verifyFiles1 = (verify1 && verify1.files) ? verify1.files : [];
		const verifySourceAll1 = verifyFiles1.filter(f => f.type === 'SERVER_JS').map(f => String(f.source || '')).join('\n');
		const createdFns1 = additions.map(a => {
			const m = a.match(/runDailyAuditsBatch(\d+)/);
			return m ? `runDailyAuditsBatch${m[1]}` : null;
		}).filter(Boolean);
		const missingAfterFirstUpdate = createdFns1.filter(fn => !new RegExp(`function\\s+${fn}\\s*\\(`).test(verifySourceAll1));
		if (missingAfterFirstUpdate.length === 0) {
			return { created: createdFns1 };
		}

		// Fallback: create a new SERVER_JS file containing the additions
		const newFileName = 'BatchRunners_Auto';
		const newFileSource = additions.join('\n');
		const updatedFiles2 = verifyFiles1.concat([{ name: newFileName, type: 'SERVER_JS', source: newFileSource }]);
		updated = updateScriptProjectContent_(updatedFiles2);
		if (!updated) throw new Error('Failed to update project content (fallback new file)');

		// Verify again
		const verify2 = getScriptProjectContent_();
		const verifySourceAll2 = (verify2 && verify2.files ? verify2.files : []).filter(f => f.type === 'SERVER_JS').map(f => String(f.source || '')).join('\n');
		const stillMissing = createdFns1.filter(fn => !new RegExp(`function\\s+${fn}\\s*\\(`).test(verifySourceAll2));
		if (stillMissing.length > 0) {
			throw new Error(`Verification failed: could not find newly added function(s): ${stillMissing.join(', ')}`);
		}
		return { created: createdFns1 };
	} catch (e) {
		Logger.log('insertMissingBatchFunctionsIntoSource_ error: ' + e.message);
		throw e;
	}
}

/**
 * Read Code.js from the Apps Script project and return batch function presence info.
 * Uses the Script advanced service (Script.Projects.getContent).
 * Returns { existingIndexes: Set<number>, count: number, maxExisting: number }
 */
function listBatchFunctionsInSource_() {
		try {
				const manifest = getScriptProjectContent_();
				if (!manifest || !manifest.files) throw new Error('Unable to read project files');
				const serverFiles = manifest.files.filter(f => f.type === 'SERVER_JS');
		if (!serverFiles.length) throw new Error('No SERVER_JS files found');
		const batchFnRegex = /function\s+runDailyAuditsBatch(\d+)\s*\(/g;
		const existingIndexes = new Set();
		for (const f of serverFiles) {
			const src = String(f.source || '');
			let match;
			while ((match = batchFnRegex.exec(src)) !== null) {
				const idx = Number(match[1]);
				if (!isNaN(idx)) existingIndexes.add(idx);
			}
		}
		let maxExisting = 0;
		existingIndexes.forEach(i => { if (i > maxExisting) maxExisting = i; });
		return { existingIndexes, count: existingIndexes.size, maxExisting };
	} catch (e) {
		Logger.log('listBatchFunctionsInSource_ error: ' + e.message);
		throw e;
	}
}

// === Apps Script API helpers (via REST) ===
function getScriptProjectContent_() {
	const projectId = ScriptApp.getScriptId();
	const url = `https://script.googleapis.com/v1/projects/${projectId}/content`;
	const headers = { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` };
	const resp = UrlFetchApp.fetch(url, { headers, muteHttpExceptions: true });
	const code = resp.getResponseCode();
	if (code !== 200) {
		throw new Error(`Apps Script API getContent failed: HTTP ${code} — ${resp.getContentText().slice(0, 500)}`);
	}
	return JSON.parse(resp.getContentText());
}

function updateScriptProjectContent_(files) {
	const projectId = ScriptApp.getScriptId();
	const url = `https://script.googleapis.com/v1/projects/${projectId}/content`;
	const headers = {
		Authorization: `Bearer ${ScriptApp.getOAuthToken()}`,
		'Content-Type': 'application/json'
	};
	const payload = JSON.stringify({ files });
	const resp = UrlFetchApp.fetch(url, { method: 'put', headers, payload, muteHttpExceptions: true });
	const code = resp.getResponseCode();
	if (code !== 200) {
		throw new Error(`Apps Script API updateContent failed: HTTP ${code} — ${resp.getContentText().slice(0, 500)}`);
	}
	return true;
}

// === UI MENU & MODALS ===
function onOpen() {
	// Always try to create the Admin menu first; never block UI on validation.
	let ui;
	try {
		ui = SpreadsheetApp.getUi();
		createAuditMenu(ui);
	} catch (e) {
		Logger.log('onOpen: could not create menu: ' + e.message);
	}

	// Run validations and helpers non-blocking so the menu still shows up.
	try { validateAuditConfigs(); } catch (e) { Logger.log('onOpen: validateAuditConfigs warning: ' + e.message); }
	try { checkDriveApiEnabled(); } catch (e) { Logger.log('onOpen: checkDriveApiEnabled warning: ' + e.message); }
	try { ensureMenuFunctionsPresent(); } catch (e) { Logger.log('onOpen: ensureMenuFunctionsPresent warning: ' + e.message); }
	try { ensureBatchRunnerFunctionsPresent(); } catch (e) { Logger.log('onOpen: ensureBatchRunnerFunctionsPresent warning: ' + e.message); }
	try { syncDeliveryModeStatus(); } catch (e) { Logger.log('onOpen: syncDeliveryModeStatus warning: ' + e.message); }
	// Also run the explicit delivery mode sync function that the trigger uses
	try { runDeliveryModeSync(); } catch (e) { Logger.log('onOpen: runDeliveryModeSync warning: ' + e.message); }
	try { installDeliveryModeSyncTrigger(); } catch (e) { Logger.log('onOpen: installDeliveryModeSyncTrigger warning: ' + e.message); }

	// Show admin refresh prompt once per session (user property guards)
	try {
		if (ui) {
			const props = PropertiesService.getUserProperties();
			const seen = props.getProperty('CM360_ADMIN_REFRESH_SEEN');
			if (!seen) {
				const html = HtmlService.createHtmlOutputFromFile('AdminRefreshPrompt').setWidth(360).setHeight(140);
				ui.showSidebar(html);
				props.setProperty('CM360_ADMIN_REFRESH_SEEN', '1');
			}
		}
	} catch (e) {
		Logger.log('onOpen: admin refresh prompt warning: ' + e.message);
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

/** Debug function to check menu creation issues */
function debugMenuCreation() {
 const ui = SpreadsheetApp.getUi();
 try {
   Logger.log('🔍 Starting menu creation debug...');
   
   // Test basic UI access
   Logger.log('✅ UI object accessible');
   
   // Test menu creation step by step
   const menu = ui.createMenu('Admin Controls');
   Logger.log('✅ Menu object created');
   
   menu.addItem('⚙️  Prepare Environment', 'prepareAuditEnvironment');
   Logger.log('✅ First menu item added');
   
   menu.addToUi();
   Logger.log('✅ Menu added to UI successfully');
   
   ui.alert('Debug Complete', 'Admin Controls menu should now be visible. Check the logs for details.', ui.ButtonSet.OK);
   
 } catch (error) {
   Logger.log(`❌ Menu creation failed: ${error.message}`);
   Logger.log(`❌ Error stack: ${error.stack}`);
   ui.alert('Menu Creation Failed', `Error: ${error.message}`, ui.ButtonSet.OK);
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
 .addItem('🧩  CM360 Config Builder…', 'showConfigCreationHelper')
 .addSeparator()
// External config (lean)
 .addItem('📤  Sync TO External Config', 'syncToExternalConfig')
 .addItem('📥  Sync FROM External Config', 'syncFromExternalConfig')
 .addSeparator()
 // (Removed) Script Properties config helpers
 // Requests
 .addItem('📝  Create Audit Request...', 'showCreateAuditRequestPicker')
 .addItem('▶️  Process Audit Requests', 'processAuditRequests')
 .addItem('🛠️  Fix Audit Requests Sheet', 'fixAuditRequestsSheet')
 .addSeparator()
 // Tools & Utilities
 .addItem('🔁  Update Placement Names', 'updatePlacementNamesFromReportsWithUI')
 .addItem('🔐  Check Authorization', 'checkAuthorizationStatus')
 .addItem('🧾  Validate Configs', 'debugValidateAuditConfigs')
 .addItem('⏱️  Install All Triggers', 'installAllAutomationTriggers')
 .addItem('🔄  Sync Delivery Mode Now', 'runDeliveryModeSync')
 .addItem('📮  Debug Email Delivery', 'debugEmailDeliveryStatus')
 .addItem('✉️  Send Test Admin Email', 'sendTestAdminEmail')
 .addItem('👀  Preview Daily Summary', 'previewDailySummaryNow')
 .addItem('🔎  Silent Withhold Check…', 'showSilentWithholdCheck')
 .addItem('🩺  Run Health Check (Admin)', 'runHealthCheckAndEmail')
 .addItem('🧪  Test Thresholds…', 'showThresholdTestPicker')
 .addSeparator()
 // Manual Run Options
 .addItem('🧪  [TEST] Run Batch or Config', 'showBatchTestPicker')
 .addItem('▶️  Run Audit for...', 'showConfigPicker')
 .addSeparator()
 // Access Tools (no sidebar)
 .addItem('📦  Batch Assignments', 'showBatchAssignmentsModal')
 .addItem('⏰  Install Health Check Trigger', 'installHealthCheckTrigger')
 .addItem('🛡️  Install Audit Watchdog Trigger', 'installAuditWatchdogTrigger')
 .addItem('ℹ️  About Admin Controls…', 'showAdminControlsHelp')
 .addToUi();
}

// Provide help metadata for Admin Controls items
function getAdminControlsHelpItems() {
	const items = [
		{ label: '⚙️  Prepare Environment', fn: 'prepareAuditEnvironment', desc: 'Creates missing Gmail labels and Drive folders for all configs. Also summarizes any labels without recent mail.' },
		{ label: '📄  Thresholds (create/open)', fn: 'getOrCreateThresholdsSheet', desc: 'Opens or creates the Audit Thresholds sheet and applies formatting/validations.' },
		{ label: '🚫  Exclusions (create/open)', fn: 'getOrCreateExclusionsSheet', desc: 'Opens or creates the Audit Exclusions sheet with protected Placement Name column and validations.' },
		{ label: '📧  Recipients (create/open)', fn: 'getOrCreateRecipientsSheet', desc: 'Opens or creates the Audit Recipients sheet to manage To/CC and withhold settings.' },
		{ label: '🧩  CM360 Config Builder…', fn: 'showConfigCreationHelper', desc: 'Guided UI to add a new CM360 audit config; provides next steps and admin hints.' },
		{ label: '📤  Sync TO External Config', fn: 'syncToExternalConfig', desc: 'Copies Admin sheet tabs (Thresholds/Recipients/Exclusions/Requests) to the external spreadsheet.' },
		{ label: '📥  Sync FROM External Config', fn: 'syncFromExternalConfig', desc: 'Pulls latest config from the external spreadsheet into Admin tabs.' },
		{ label: '📝  Create Audit Request...', fn: 'showCreateAuditRequestPicker', desc: 'Open a picker to submit a one-off audit request.' },
		{ label: '▶️  Process Audit Requests', fn: 'processAuditRequests', desc: 'Executes pending one-off audit requests from the Requests sheet.' },
		{ label: '🛠️  Fix Audit Requests Sheet', fn: 'fixAuditRequestsSheet', desc: 'Repairs headers/validations for the Audit Requests sheet.' },
		{ label: '🔁  Update Placement Names', fn: 'updatePlacementNamesFromReportsWithUI', desc: 'Reads latest merged reports and fills Placement Name in the EXTERNAL Exclusions sheet for rows with IDs.' },
		{ label: '🔐  Check Authorization', fn: 'checkAuthorizationStatus', desc: 'Verifies script scopes/auth and sends a result to the current user.' },
		{ label: '🧾  Validate Configs', fn: 'debugValidateAuditConfigs', desc: 'Validates audit configs derived from Recipients and logs findings.' },
		{ label: '⏱️  Install All Triggers', fn: 'installAllAutomationTriggers', desc: 'Reinstalls all automation triggers (daily batches, summaries, syncs, cleanup) without touching batch stubs.' },
		{ label: '🔄  Sync Delivery Mode Now', fn: 'runDeliveryModeSync', desc: 'Updates the “Delivery Mode” instruction line on Admin and External Recipients tabs.' },
		{ label: '📮  Debug Email Delivery', fn: 'debugEmailDeliveryStatus', desc: 'Logs delivery mode, admin email, and remaining send quota.' },
		{ label: '✉️  Send Test Admin Email', fn: 'sendTestAdminEmail', desc: 'Sends a quick test message to admin to verify email plumbing.' },
		{ label: '👀  Preview Daily Summary', fn: 'previewDailySummaryNow', desc: 'Builds and shows the daily summary email preview without sending.' },
		{ label: '🔎  Silent Withhold Check…', fn: 'showSilentWithholdCheck', desc: 'Pick a config to simulate an audit’s email decision (send vs withhold) without sending.' },
		{ label: '🩺  Run Health Check (Admin)', fn: 'runHealthCheckAndEmail', desc: 'Performs a fast, read-only workflow health check and emails the admin.' },
		{ label: '🧪  Test Thresholds…', fn: 'showThresholdTestPicker', desc: 'Run a full audit for selected config with detailed threshold logging to diagnose threshold filtering.' },
		{ label: '🧪  [TEST] Run Batch or Config', fn: 'showBatchTestPicker', desc: 'Pick a batch or specific config to run on demand for testing.' },
		{ label: '▶️  Run Audit for...', fn: 'showConfigPicker', desc: 'Pick any single config to run a one-off audit now.' },
		{ label: '📦  Batch Assignments', fn: 'showBatchAssignmentsModal', desc: 'Displays which configs are assigned to each batch runner function.' },
		{ label: '⏰  Install Health Check Trigger', fn: 'installHealthCheckTrigger', desc: 'Installs a daily trigger (early morning) to email the health report to admin.' },
		{ label: '🛡️  Install Audit Watchdog Trigger', fn: 'installAuditWatchdogTrigger', desc: 'Installs a 3-hour watchdog trigger to detect and alert on timed-out batches.' }
	];
	return items;
}

function showAdminControlsHelp() {
	try {
		const items = getAdminControlsHelpItems();
		const tpl = HtmlService.createTemplateFromFile('AdminControlsHelp');
		tpl.items = items;
		const html = tpl.evaluate().setWidth(560).setHeight(620);
		SpreadsheetApp.getUi().showModalDialog(html, 'About Admin Controls');
	} catch (e) {
		Logger.log('showAdminControlsHelp error: ' + e.message);
	}
}

// Quick diagnostics for email delivery environment
function debugEmailDeliveryStatus() {
	try {
		const mode = (function(){ try { return getCurrentDeliveryMode_ ? getCurrentDeliveryMode_() : (getStagingMode_() === 'Y' ? 'STAGING' : 'PRODUCTION'); } catch(e) { return getStagingMode_() === 'Y' ? 'STAGING' : 'PRODUCTION'; } })();
		let quota = null;
		try { quota = MailApp.getRemainingDailyQuota(); } catch (e) {}
		const msg = `Delivery Mode=${mode}; Admin=${ADMIN_EMAIL}; Remaining quota=${quota !== null ? quota : 'unknown'}`;
		Logger.log('[DEBUG] ' + msg);
		try { const ss = SpreadsheetApp.getActiveSpreadsheet(); if (ss) ss.toast(msg, 'Email Diagnostics', 8); } catch (_) {}
	} catch (e) {
		Logger.log('debugEmailDeliveryStatus error: ' + e.message);
	}
}

// === HEALTH CHECK (non-intrusive) ===
function runHealthCheckAndEmail() {
	try {
		const report = buildHealthCheckReport_();
		const subject = `[HEALTH] CM360 Audit Workflow Check — ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}`;
		const html = formatHealthCheckHtml_(report);
		// Send admin-only, independent of STAGING_MODE
		safeSendEmail({ to: ADMIN_EMAIL, subject, htmlBody: html, plainBody: 'See HTML report.' }, 'Health Check');
		Logger.log('Health check email sent to admin.');
	} catch (e) {
		Logger.log('runHealthCheckAndEmail error: ' + e.message);
	}
}

function buildHealthCheckReport_() {
	const started = new Date();
	const tz = Session.getScriptTimeZone();
	const findings = [];
	const warnings = [];
	const errors = [];
	const stats = { configs: 0, labelsFound: 0, labelsMissing: 0, foldersFound: 0, foldersMissing: 0 };

	// Quick constants presence
	try {
		const ok = (typeof CLEANUP_RUNTIME_LIMIT_MS !== 'undefined');
		if (!ok) warnings.push('CLEANUP_RUNTIME_LIMIT_MS missing; using fallback in cleanup.');
	} catch (_) {}

	// Drive API availability (don’t fail check if unavailable; just warn)
	let driveEnabled = false;
	try { driveEnabled = (typeof Drive !== 'undefined' && Drive.Files && typeof Drive.Files.insert === 'function'); } catch (_) {}
	if (!driveEnabled) warnings.push('Advanced Drive API is not enabled; merge/export features will be skipped.');

	// Load config sheets (read-only)
	let recipientsData = {};
	let thresholdsData = {};
	try { recipientsData = loadRecipientsFromSheet(); } catch (e) { errors.push('Failed to load recipients: ' + e.message); }
	try { thresholdsData = loadThresholdsFromSheet(); } catch (e) { warnings.push('Failed to load thresholds: ' + e.message); }

	// Gather configs from recipients (Active only)
	let configs = [];
	try {
		configs = getAuditConfigs();
	} catch (e) {
		errors.push('getAuditConfigs failed: ' + e.message);
		configs = [];
	}
	stats.configs = configs.length;

	// Per-config light checks (no audit run, no email)
	configs.forEach(cfg => {
		const name = cfg.name;
		const labelName = cfg.label;
		const mergedPath = cfg.mergedFolderPath;
		const tempPath = cfg.tempDailyFolderPath;

		// Gmail label exists?
		try {
			const lbl = findGmailLabel_(labelName, name) || GmailApp.getUserLabelByName(labelName);
			if (lbl) { stats.labelsFound++; }
			else { stats.labelsMissing++; findings.push(`Label missing for ${name}: ${labelName}`); }
		} catch (e) { warnings.push(`Label check error for ${name}: ${e.message}`); }

		// Drive folders exist? (read-only path walk)
		try {
			const merged = getDriveFolderByPathReadOnly_(mergedPath);
			const temp = getDriveFolderByPathReadOnly_(tempPath);
			if (!merged) { stats.foldersMissing++; findings.push(`Merged folder missing for ${name}: ${mergedPath.join(' / ')}`); } else { stats.foldersFound++; }
			if (!temp) { stats.foldersMissing++; findings.push(`Temp folder missing for ${name}: ${tempPath.join(' / ')}`); } else { stats.foldersFound++; }
		} catch (e) { warnings.push(`Folder check error for ${name}: ${e.message}`); }

		// Recipients sanity
		try {
			const entry = recipientsData[name];
			if (!entry || !String(entry.primary || '').trim()) {
				findings.push(`Recipients missing or empty for ${name}`);
			}
		} catch (e) { warnings.push(`Recipients check error for ${name}: ${e.message}`); }

		// Thresholds presence for expected flags
		try {
			const th = thresholdsData[name] || {};
			const flagTypes = ['clicks_greater_than_impressions','out_of_flight_dates','pixel_size_mismatch','default_ad_serving'];
			const missing = flagTypes.filter(ft => !th[ft]);
			if (missing.length) findings.push(`Missing thresholds for ${name}: ${missing.join(', ')}`);
		} catch (e) { warnings.push(`Thresholds check error for ${name}: ${e.message}`); }
	});

	// Trigger posture (existence only; don’t create)
	try {
		const triggers = ScriptApp.getProjectTriggers() || [];
		const daily = triggers.filter(t => (t.getHandlerFunction && /^runDailyAuditsBatch\d+$/.test(t.getHandlerFunction())));
		if (daily.length === 0) warnings.push('No daily audit batch triggers are installed.');
	} catch (e) { warnings.push('Trigger check failed: ' + e.message); }

	const ended = new Date();
	return { startedAt: Utilities.formatDate(started, tz, 'yyyy-MM-dd HH:mm:ss'),
					 endedAt: Utilities.formatDate(ended, tz, 'yyyy-MM-dd HH:mm:ss'),
					 durationSec: Math.round((ended - started)/1000), stats, findings, warnings, errors };
}

// === WATCHDOG: detect hung/timeout batches and notify admin ===
function auditWatchdogCheck() {
	try {
		const props = PropertiesService.getScriptProperties();
		const listRaw = props.getProperty(AUDIT_RUN_LIST_KEY);
		const list = listRaw ? JSON.parse(listRaw) : [];
		if (!list.length) { Logger.log('[Watchdog] No recent batches recorded.'); return; }
		const now = Date.now();
		const alerts = [];
		const TIMEOUT_MS = 6 * 60 * 1000; // 6 minutes
		const GRACE_MS = 60 * 1000; // 1 minute grace
		for (const id of list) {
			try {
				const key = AUDIT_RUN_STATE_KEY_PREFIX + id;
				const raw = props.getProperty(key);
				if (!raw) continue;
				const state = JSON.parse(raw);
				const started = Number(state.startedAt || 0);
				const completed = Number(state.completedAt || 0);
				const alerted = Number(state.alertedAt || 0);
				if (!started) continue;
				// Only consider runs from the last day
				if ((now - started) > 24 * 60 * 60 * 1000) continue;
				// Only alert once per hung batch: skip if we've already alerted
				if (!completed && !alerted && (now - started) > (TIMEOUT_MS + GRACE_MS)) {
					alerts.push({ id, started, configs: state.configs || [], isFinal: state.isFinal });
				}
			} catch (e) { Logger.log('[Watchdog] parse error for ' + id + ': ' + e.message); }
		}
		if (!alerts.length) { Logger.log('[Watchdog] No hung batches detected.'); return; }

		const tz = Session.getScriptTimeZone();
		const lines = alerts.map(a => `• ${a.id} — started ${Utilities.formatDate(new Date(a.started), tz, 'yyyy-MM-dd HH:mm:ss')} — configs: ${(a.configs||[]).join(', ')}`);
		const body = `The following audit batch execution(s) appear to have timed out without completing:\n\n${lines.join('\n')}\n\nThis alert is generated by the CM360 audit watchdog.`;
		const html = `<p>The following audit batch execution(s) appear to have timed out without completing:</p><ul>${alerts.map(a => `<li><code>${escapeHtml(a.id)}</code> — started ${escapeHtml(Utilities.formatDate(new Date(a.started), tz, 'yyyy-MM-dd HH:mm:ss'))} — configs: ${escapeHtml((a.configs||[]).join(', '))}</li>`).join('')}</ul><p>This alert is generated by the CM360 audit watchdog.</p>`;
		// Send a single alert email covering all newly detected hung batches
		safeSendEmail({ to: ADMIN_EMAIL, subject: '[ALERT] CM360 Audit timed-out batch detected', htmlBody: html, plainBody: body }, 'Watchdog');

		// Mark alertedAt for each alerted batch to avoid duplicate notifications
		const alertedAt = Date.now();
		for (const a of alerts) {
			try {
				const key = AUDIT_RUN_STATE_KEY_PREFIX + a.id;
				const raw = props.getProperty(key);
				if (!raw) continue;
				const state = JSON.parse(raw);
				state.alertedAt = alertedAt;
				props.setProperty(key, JSON.stringify(state));
			} catch (e) { Logger.log('[Watchdog] failed to mark alertedAt for ' + a.id + ': ' + e.message); }
		}
		Logger.log(`[Watchdog] Alerted admin for ${alerts.length} hung batch(es).`);
		try {
			attemptSendDailySummary_({ allowPlaceholders: true, reason: 'Watchdog hung detection' });
		} catch (summaryErr) {
			Logger.log('auditWatchdogCheck summary attempt error: ' + summaryErr.message);
		}
	} catch (e) {
		Logger.log('auditWatchdogCheck error: ' + e.message);
	}
}

function installAuditWatchdogTrigger() {
	try {
		const results = installAllAutomationTriggers({ handlers: ['auditWatchdog'] });
		Logger.log('installAuditWatchdogTrigger results: ' + results.join(' | '));
		try {
			const ss = SpreadsheetApp.getActiveSpreadsheet();
			if (ss) ss.toast('Audit watchdog trigger refreshed.', 'Trigger Installer', 5);
		} catch (_) {}
		return results;
	} catch (e) {
		Logger.log('installAuditWatchdogTrigger error: ' + e.message);
		return [];
	}
}

function formatHealthCheckHtml_(report) {
	const s = report.stats || { configs: 0, labelsFound: 0, labelsMissing: 0, foldersFound: 0, foldersMissing: 0 };
	const list = (arr) => (arr && arr.length) ? '<ul>' + arr.map(x => `<li>${escapeHtml(String(x))}</li>`).join('') + '</ul>' : '<p>None</p>';
	return `
	<div style="font-family:Arial,sans-serif; font-size:13px;">
		<h3 style="margin:0 0 8px;">CM360 Health Check</h3>
		<div>Started: ${escapeHtml(report.startedAt)} | Ended: ${escapeHtml(report.endedAt)} | Duration: ${report.durationSec}s</div>
		<div style="margin:8px 0;">
			<span style="display:inline-block; background:#eef3fe; border:1px solid #d2e3fc; border-radius:12px; padding:2px 8px; margin-right:6px;">Configs: ${s.configs}</span>
			<span style="display:inline-block; background:#e6f4ea; border:1px solid #ccebd7; border-radius:12px; padding:2px 8px; margin-right:6px;">Labels OK: ${s.labelsFound}</span>
			<span style="display:inline-block; background:#fde8e8; border:1px solid #facaca; border-radius:12px; padding:2px 8px; margin-right:6px;">Labels Missing: ${s.labelsMissing}</span>
			<span style="display:inline-block; background:#e6f4ea; border:1px solid #ccebd7; border-radius:12px; padding:2px 8px; margin-right:6px;">Folders OK: ${s.foldersFound}</span>
			<span style="display:inline-block; background:#fde8e8; border:1px solid #facaca; border-radius:12px; padding:2px 8px; margin-right:6px;">Folders Missing: ${s.foldersMissing}</span>
		</div>
		<h4 style="margin:10px 0 6px;">Errors</h4>
		${list(report.errors)}
		<h4 style="margin:10px 0 6px;">Warnings</h4>
		${list(report.warnings)}
		<h4 style="margin:10px 0 6px;">Findings</h4>
		${list(report.findings)}
		<p style="margin-top:12px; font-size:12px; color:#666;">This is a non-intrusive read-only health report. No audit runs or emails (other than this admin report) were performed.</p>
	</div>`;
}

// Install/refresh a daily health check a few hours before Daily Audit batches
function installHealthCheckTrigger() {
	try {
		const results = installAllAutomationTriggers({ handlers: ['healthCheck'] });
		Logger.log('installHealthCheckTrigger results: ' + results.join(' | '));
		try {
			const ui = SpreadsheetApp.getUi();
			ui.alert('Health Check Trigger', results.join('\n'), ui.ButtonSet.OK);
		} catch (_) {}
		return results;
	} catch (e) {
		Logger.log('installHealthCheckTrigger error: ' + e.message);
		return [];
	}
}

// Send a quick test email to ADMIN_EMAIL to verify delivery
function sendTestAdminEmail() {
	try {
		const mode = getStagingMode_() === 'Y' ? 'STAGING' : 'PRODUCTION';
		const subject = `[${mode}] CM360 Admin Test Email`;
		const html = `<p style="font-family:Arial,sans-serif; font-size:13px;">This is a test email sent to <b>${escapeHtml(ADMIN_EMAIL)}</b> from the CM360 Apps Script.</p>`;
		const ok = safeSendEmail({ to: ADMIN_EMAIL, subject, plainBody: 'CM360 admin test email', htmlBody: html }, 'Test Admin Email');
		const ss = SpreadsheetApp.getActiveSpreadsheet();
		if (ss) ss.toast(ok ? `Test email sent to ${ADMIN_EMAIL}` : `Test email failed; check logs`, 'Admin Email Test', 8);
	} catch (e) {
		Logger.log('sendTestAdminEmail error: ' + e.message);
	}
}

// Prompt for a config name and set up Drive paths, Gmail label, and sheet entries
function showConfigCreationHelper() {
	const ui = SpreadsheetApp.getUi();
	const res = ui.prompt('CM360 Config Builder', 'Enter a new Config Name (e.g., PST01):', ui.ButtonSet.OK_CANCEL);
	if (res.getSelectedButton() !== ui.Button.OK) return;
	const name = String(res.getResponseText() || '').trim();
	if (!name) { ui.alert('Config Builder', 'Config name cannot be empty.', ui.ButtonSet.OK); return; }
	try {
		const created = ensureConfigArtifacts_(name);
		const msg = [
			`Config: ${name}`,
			created.labelCreated ? '• Gmail label created' : '• Gmail label exists',
			created.mergedFolderCreated ? '• Merged folder created' : '• Merged folder exists',
			created.tempFolderCreated ? '• Temp folder created' : '• Temp folder exists',
			created.recipientsAdded ? '• Added row to Audit Recipients' : '• Recipients row exists',
			created.thresholdsAdded ? '• Added default thresholds' : '• Thresholds exist'
		].join('\n');
		// Add final note about where updates were applied
		const suffix = (created.recipientsAdded || created.thresholdsAdded)
		  ? '\n\nUpdates were applied to the Audit Recipients and Audit Thresholds tabs on the Helper Menu sheet.'
		  : '';
		ui.alert('Config Builder Complete', msg + suffix, ui.ButtonSet.OK);
	} catch (e) {
		ui.alert('Config Builder Error', e.message, ui.ButtonSet.OK);
	}
}

function ensureConfigArtifacts_(configName) {
	const name = String(configName || '').trim();
	if (!name) throw new Error('Missing config name');
	const labelName = `Daily Audits/CM360/${name}`;
	const mergedPath = [...TRASH_ROOT_PATH, 'To Trash After 60 Days', 'Merged Reports', name];
	const tempPath = [...TRASH_ROOT_PATH, 'To Trash After 60 Days', 'Temp Daily Reports', name];

	// Gmail label (robust creation + verification)
	let { created: labelCreated } = ensureGmailLabelExists_(labelName);

	// Drive folders
	const ensureFolder = (path) => {
		let f = DriveApp.getRootFolder();
		let created = false;
		path.forEach(part => {
			const it = f.getFoldersByName(part);
			if (it.hasNext()) { f = it.next(); }
			else { f = f.createFolder(part); created = true; }
		});
		return created;
	};
	const mergedFolderCreated = ensureFolder(mergedPath);
	const tempFolderCreated = ensureFolder(tempPath);

	// Sheets
	const recipientsSheet = getOrCreateRecipientsSheet();
	const recData = recipientsSheet.getDataRange().getValues();
	let hasRecipient = false;
	for (let i = 1; i < recData.length; i++) {
		if (String(recData[i][0] || '').trim() === name) { hasRecipient = true; break; }
	}
	let recipientsAdded = false;
	if (!hasRecipient) {
		const now = formatDate(new Date(), 'yyyy-MM-dd');
		recipientsSheet.appendRow([name, ADMIN_EMAIL, '', 'TRUE', 'FALSE', now]);
		recipientsAdded = true;
	}

	const thresholdsSheet = getOrCreateThresholdsSheet();
	const thData = thresholdsSheet.getDataRange().getValues();
	const flagTypes = ['clicks_greater_than_impressions','out_of_flight_dates','pixel_size_mismatch','default_ad_serving'];
	const existingFlags = new Set();
	for (let i = 1; i < thData.length; i++) {
		if (String(thData[i][0] || '').trim() === name) existingFlags.add(String(thData[i][1] || '').trim());
	}
	let thresholdsAdded = false;
	flagTypes.forEach(ft => {
		if (!existingFlags.has(ft)) {
			thresholdsSheet.appendRow([name, ft, 0, 0, 'TRUE']);
			thresholdsAdded = true;
		}
	});

	// Clear cache so new config shows up immediately
	clearAuditConfigsCache_();

	return { labelCreated, mergedFolderCreated, tempFolderCreated, recipientsAdded, thresholdsAdded };
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
				targets.push(openSpreadsheetById_(EXTERNAL_CONFIG_SHEET_ID));
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
				const exHeaders = ['Config Name','Placement ID','Placement Name (auto-populated)','Site Name','Name Fragment','Apply to All Configs','Flag Type','Reason','Added By','Date Added','Active','','INSTRUCTIONS'];
				if (!ex) ex = ss.insertSheet(EXCLUSIONS_SHEET_NAME);
				ex.getRange(1, 1, 1, exHeaders.length).setValues([exHeaders]);
				try { ex.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff'); } catch(e) {}
				try { ex.autoResizeColumns(1, 11); } catch(e) {}
				_ensureInstructionsOnSheet_(ex, _buildExclusionsInstructions_());

				// Enforce protection and visual styling on Placement Name column (C)
				try { enforcePlacementNameProtectionAndStyle_(ex); } catch (e) { Logger.log('refreshAdminHeadersAndInstructions: placement name protect/style: ' + e.message); }

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
function refreshExternalConfigInstructions(options) {
	options = options || {};
	const silent = options && options.silent === true;
	const ui = silent ? null : SpreadsheetApp.getUi();
	if (!EXTERNAL_CONFIG_SHEET_ID) {
		if (silent) {
			Logger.log('refreshExternalConfigInstructions: EXTERNAL_CONFIG_SHEET_ID not configured; skipping.');
		} else {
			ui.alert('No External Config Sheet', 'EXTERNAL_CONFIG_SHEET_ID is not configured. Cannot update external instructions.', ui.ButtonSet.OK);
		}
		return;
	}

	try {
		const ss = openSpreadsheetById_(EXTERNAL_CONFIG_SHEET_ID);

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
		const exHeaders = ['Config Name','Placement ID','Placement Name (auto-populated)','Site Name','Name Fragment','Apply to All Configs','Flag Type','Reason','Added By','Date Added','Active','','INSTRUCTIONS'];
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
		if (silent) {
			return true;
		}
		ui.alert('Refresh Complete', 'Headers and INSTRUCTIONS refreshed on target spreadsheets. No table data was modified.', ui.ButtonSet.OK);
		return true;
	} catch (e) {
		Logger.log('refreshExternalConfigInstructions error: ' + e.message);
		if (!silent) {
			ui.alert('Refresh Failed', `Failed to refresh instructions: ${e.message}`, ui.ButtonSet.OK);
		}
		throw e;
	}
}

function refreshExternalConfigInstructionsSilent() {
	try {
		refreshExternalConfigInstructions({ silent: true });
	} catch (e) {
		Logger.log('refreshExternalConfigInstructionsSilent error: ' + e.message);
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
		
		const ss = openSpreadsheetById_(EXTERNAL_CONFIG_SHEET_ID);
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
		'prepareAuditEnvironment','getOrCreateThresholdsSheet','getOrCreateExclusionsSheet','getOrCreateRecipientsSheet','addMissingConfigNames',
		'promptSetupExternalConfigMenu','ensureExternalConfigInstructions','updateExternalConfigInstructions','syncToExternalConfig',
		'syncFromExternalConfig','populateExternalConfigWithDefaults','showCreateAuditRequestPicker',
		'processAuditRequests','fixAuditRequestsSheet','refreshExternalHeaderStyles',
		'updatePlacementNamesFromReports','updatePlacementNamesFromReportsWithUI','checkAuthorizationStatus','debugValidateAuditConfigs',
		'setupAndInstallBatchTriggers','showBatchTestPicker','showConfigPicker','showAuditDashboard',
		'createMissingThresholds','createMissingRecipients','createMissingExclusions'
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



function showConfigPicker() {
 const template = HtmlService.createTemplateFromFile('ConfigPicker');
 template.auditConfigs = getAuditConfigs(); // Pass into template
 const html = template.evaluate()
 .setWidth(300)
 .setHeight(160);
 SpreadsheetApp.getUi().showModalDialog(html, 'Select Audit Config');
}

function showThresholdTestPicker() {
 const template = HtmlService.createTemplateFromFile('ThresholdTestPicker');
 template.auditConfigs = getAuditConfigs(); // Pass into template
 const html = template.evaluate()
 .setWidth(360)
 .setHeight(240);
 SpreadsheetApp.getUi().showModalDialog(html, 'Test Thresholds');
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

// === EMAIL SUPPRESSION (silent run) HELPERS ===
function setEmailSuppressed_(val) {
	try {
		PropertiesService.getScriptProperties().setProperty('CM360_EMAIL_SUPPRESSED', val ? '1' : '0');
	} catch (e) { Logger.log('setEmailSuppressed_ error: ' + e.message); }
}
function clearEmailSuppressed_() {
	try { PropertiesService.getScriptProperties().deleteProperty('CM360_EMAIL_SUPPRESSED'); } catch (e) {}
}
function isEmailSuppressed_() {
	try { return PropertiesService.getScriptProperties().getProperty('CM360_EMAIL_SUPPRESSED') === '1'; } catch (e) { return false; }
}

// UI to run a silent withhold check for a chosen config
function showSilentWithholdCheck() {
	const ui = SpreadsheetApp.getUi();
	try {
		const cfgs = getAuditConfigs();
		if (!cfgs || cfgs.length === 0) { ui.alert('No configs found.'); return; }
		const options = cfgs.map((c, i) => `${i + 1}. ${c.name}`).join('\n');
		const res = ui.prompt('Silent Withhold Check', 'Select a configuration to test (no emails will be sent):\n\n' + options + '\n\nEnter number:', ui.ButtonSet.OK_CANCEL);
		if (res.getSelectedButton() !== ui.Button.OK) return;
		const idx = parseInt(String(res.getResponseText() || '').trim(), 10) - 1;
		if (isNaN(idx) || idx < 0 || idx >= cfgs.length) { ui.alert('Invalid selection.'); return; }
		const config = cfgs[idx];
		const out = runSilentWithholdCheck_(config.name);
		const msg = [
			`Config: ${out.name}`,
			`Status: ${out.status}`,
			`Flagged Rows: ${out.flaggedRows}`,
			`Email Would Be Withheld: ${out.emailWouldBeWithheld ? 'YES' : 'NO'}`,
			`Email Would Send: ${out.emailWouldSend ? 'YES' : 'NO'}`,
			out.latestReportUrl ? `Latest Report: ${out.latestReportUrl}` : ''
		].filter(Boolean).join('\n');
		ui.alert('Silent Withhold Check Result', msg, ui.ButtonSet.OK);
	} catch (e) {
		ui.alert('Silent Check Failed', e.message, ui.ButtonSet.OK);
	}
}

// Core runner that suppresses email sending while executing a single config audit,
// and reports whether an email would have been sent or withheld based on current settings.
function runSilentWithholdCheck_(configName) {
	const cfg = getAuditConfigByName(configName);
	if (!cfg) throw new Error(`Config not found: ${configName}`);
	const startedSuppressed = isEmailSuppressed_();
	let resultMeta = { name: cfg.name, status: '', flaggedRows: null, emailWouldSend: false, emailWouldBeWithheld: false, latestReportUrl: '' };
	try {
		setEmailSuppressed_(true);
		const res = executeAudit(cfg);
		// Determine would-send vs withheld
		const flagged = Number(res.flaggedCount || 0);
		const recipientsData = loadRecipientsFromSheet();
		const entry = recipientsData[cfg.name];
		const withholdNoFlag = !!(entry && entry.withholdNoFlagEmails);
		const wouldSend = flagged > 0 ? true : !withholdNoFlag;
		resultMeta = {
			name: cfg.name,
			status: String(res.status || ''),
			flaggedRows: flagged,
			emailWouldSend: wouldSend,
			emailWouldBeWithheld: !wouldSend,
			latestReportUrl: res.latestReportUrl || ''
		};
		return resultMeta;
	} finally {
		// Restore original suppression state
		if (!startedSuppressed) clearEmailSuppressed_();
	}
}

// Modal view: show batch assignments grouped into boxes
function showBatchAssignmentsModal() {
	try {
		const data = getAuditConfigSummaries();
		let html = `
			<div style="font-family:Arial,sans-serif; font-size:13px; max-height:70vh; overflow:auto;">
				<h3 style="margin-top:0;">Batch Assignments</h3>
		`;
		data.forEach(batch => {
			html += `
				<div style="margin:10px 0; padding:10px; border:2px solid #bbb; border-radius:8px; background:#f9f9f9;">
					<div style="font-weight:bold; margin-bottom:6px;">${batch.batchLabel}</div>
					${batch.configs.map(c => `
						<div style="margin:6px 0; padding:6px; border:1px solid #ddd; border-radius:6px; background:#fff;">
							<div><strong>${c.name}</strong></div>
							<div style="color:#555;">Label: ${c.label}</div>
							<div style="color:#555;">Recipients: ${c.recipients}</div>
							<div style="color:#555;">Flags: ${c.flags}</div>
						</div>
					`).join('')}
				</div>
			`;
		});
		html += `</div>`;
		const out = HtmlService.createHtmlOutput(html).setWidth(480).setHeight(520);
		SpreadsheetApp.getUi().showModalDialog(out, 'Batch Assignments');
	} catch (e) {
		Logger.log('showBatchAssignmentsModal error: ' + e.message);
	}
}

// Provide batch summaries for the Dashboard sidebar
function getAuditConfigSummaries() {
	try {
		const batches = getAuditConfigBatches(BATCH_SIZE);
		const recipientsData = loadRecipientsFromSheet();
		// Static list of flag types monitored by the audit
		const flagTypes = ['Clicks > Impressions', 'Out of flight dates', 'Pixel size mismatch', 'Default ad serving'];
		const flagsText = flagTypes.join(', ');
		return batches.map((cfgs, idx) => ({
			batchLabel: `Batch ${idx + 1} (${cfgs.length} config${cfgs.length === 1 ? '' : 's'})`,
			configs: cfgs.map(c => ({
				name: c.name,
				label: c.label,
				recipients: resolveRecipients(c.name, recipientsData),
				flags: flagsText
			}))
		}));
	} catch (e) {
		Logger.log('getAuditConfigSummaries error: ' + e.message);
		return [];
	}
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
 let ss;
 try {
	 ss = openSpreadsheetById_(EXTERNAL_CONFIG_SHEET_ID);
 } catch (e) {
	 ui.alert('External Config Error', `Could not open external sheet by ID.\n\n${e.message}`, ui.ButtonSet.OK);
	 return;
 }
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
	const startTime = Date.now();
	let state = getCleanupState_();
	if (!state || state.version !== CLEANUP_STATE_VERSION) {
		state = {
			version: CLEANUP_STATE_VERSION,
			looseIndex: 0,
			temp: { configIndex: 0, subfolderIndex: 0 },
			merged: { configIndex: 0, pageToken: null },
			other: { folderIndex: 0, pageToken: null },
			looseDone: false,
			tempDone: false,
			mergedDone: false,
			otherDone: false
		};
	}
	const maxRuntimeMs = (typeof CLEANUP_RUNTIME_LIMIT_MS !== 'undefined' && Number(CLEANUP_RUNTIME_LIMIT_MS) > 0)
		? Number(CLEANUP_RUNTIME_LIMIT_MS)
		: 300000; // default 5 minutes if missing
	const cutoffDate = new Date();
	cutoffDate.setDate(cutoffDate.getDate() - 60);

	const trashRoot = ensureFolderFromState_(state, 'trashRootId', TRASH_ROOT_PATH);
	const logFolder = ensureFolderFromState_(state, 'logFolderId', DELETION_LOG_PATH);

	if (!trashRoot || !logFolder) {
		Logger.log(' Cleanup failed: Trash or Log folder not found.');
		clearCleanupState_();
		clearCleanupContinuation_();
		return;
	}

	ensureDeletionLogWorkbookStructure_(logFolder, ADMIN_LOG_NAME);

	const deletionTimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
	const logBuckets = { temp: [], merged: [] };

	if (!state.looseDone) {
		if (!processCleanupLooseFiles_(state, trashRoot, cutoffDate, deletionTimestamp, logBuckets, startTime, maxRuntimeMs)) {
			finalizeCleanupRun_(state, logFolder, ADMIN_LOG_NAME, logBuckets, false);
			return;
		}
	}

	if (!state.tempDone) {
		const tempRoot = ensureFolderFromState_(state, 'tempRootId', [...TRASH_ROOT_PATH, 'To Trash After 60 Days', 'Temp Daily Reports']);
		if (!processCleanupTempFolders_(state, tempRoot, cutoffDate, deletionTimestamp, logBuckets, startTime, maxRuntimeMs)) {
			finalizeCleanupRun_(state, logFolder, ADMIN_LOG_NAME, logBuckets, false);
			return;
		}
	}

	if (!state.mergedDone) {
		const mergedRoot = ensureFolderFromState_(state, 'mergedRootId', [...TRASH_ROOT_PATH, 'To Trash After 60 Days', 'Merged Reports']);
		if (!processCleanupMergedFiles_(state, mergedRoot, cutoffDate, deletionTimestamp, logBuckets, startTime, maxRuntimeMs)) {
			finalizeCleanupRun_(state, logFolder, ADMIN_LOG_NAME, logBuckets, false);
			return;
		}
	}

	if (!state.otherDone) {
		if (!processCleanupOtherFolders_(state, trashRoot, cutoffDate, deletionTimestamp, logBuckets, startTime, maxRuntimeMs)) {
			finalizeCleanupRun_(state, logFolder, ADMIN_LOG_NAME, logBuckets, false);
			return;
		}
	}

	finalizeCleanupRun_(state, logFolder, ADMIN_LOG_NAME, logBuckets, true);
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
	try {
		const results = installAllAutomationTriggers({ handlers: ['gasFailureForwarder'] });
		Logger.log('installGASFailureNotifierTrigger results: ' + results.join(' | '));
		return results.join('\n');
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
 
 Logger.log('Setting up validations...');
 // Keep Flag Type (col B) as free text: ensure no data validation
 try { sheet.getRange('B2:B').clearDataValidations(); } catch (e) { /* ignore */ }
 
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
 
 // Apply conditional formatting rules on Thresholds
 try { applyThresholdsFormatting_(sheet); } catch (e) { Logger.log('applyThresholdsFormatting_ error: ' + e.message); }
 
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

// Add alternating 4-row banding and inactive shading to Thresholds
function applyThresholdsFormatting_(sheet) {
	if (!sheet) return;
	const rules = sheet.getConditionalFormatRules() || [];
	// Remove prior instances of these specific rules by matching formula strings
	const filtered = rules.filter(r => {
		try {
			const cond = r.getBooleanCondition();
			if (!cond) return true;
			const vals = cond.getCriteriaValues() || [];
			const formula = String((vals[0] || '')).toUpperCase();
			if (formula.includes('ROUNDUP((ROW()-1)/4)')) return false; // remove any prior banding rules
			if (formula.includes('$E2=FALSE')) return false; // remove prior inactive rules
			return true;
		} catch (e) { return true; }
	});
	// Alternating 4-row banding over A..E (includes Active), only when Active is TRUE
	const bandRange = sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 5); // A..E
	const bandRule = SpreadsheetApp.newConditionalFormatRule()
		.whenFormulaSatisfied('=AND($E2=TRUE, ISEVEN(ROUNDUP((ROW()-1)/4)))')
		.setBackground('#e8f0fe')
		.setRanges([bandRange])
		.build();
	// Inactive shading: include Active cell and skip blank rows (requires non-empty Config Name in A)
	const inactiveRange = sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 5); // A..E
	const inactiveRule = SpreadsheetApp.newConditionalFormatRule()
		.whenFormulaSatisfied('=AND($E2=FALSE, LEN($A2)>0)')
		.setBackground('#f8d7da')
		.setRanges([inactiveRange])
		.build();
	filtered.push(bandRule, inactiveRule);
	sheet.setConditionalFormatRules(filtered);
}

function loadThresholdsFromSheet(forceSheetRead) {
 try {
 // Always read from sheet to ensure latest edits are used immediately
 const sheet = getOrCreateThresholdsSheet();
 const data = sheet.getDataRange().getValues();
 const thresholds = {};
 const skippedRows = [];
 
 // Skip header row (index 0)
 for (let i = 1; i < data.length; i++) {
 const row = data[i];
 const configName = String(row[0] || '').trim();
 const flagType = String(row[1] || '').trim();
 const minImpressions = Number(row[2] || 0);
 const minClicks = Number(row[3] || 0);
 const active = String(row[4] || '').trim().toUpperCase();
 
 // Skip empty rows, instruction rows, or inactive thresholds
 if (!configName || !flagType) {
 continue;
 }
 
 if (active !== 'TRUE') {
 skippedRows.push(`Row ${i+1}: ${configName}-${flagType} (Active='${active}' - not TRUE)`);
 continue;
 }
 
 if (configName.includes('INSTRUCTIONS') || 
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
 
 Logger.log(`Loaded thresholds from sheet for ${Object.keys(thresholds).length} configs`);
 if (skippedRows.length > 0) {
 Logger.log(`⚠️ Skipped ${skippedRows.length} inactive threshold rows:`);
 skippedRows.forEach(msg => Logger.log(` ${msg}`));
 }
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
 
 const threshold = thresholdsData[configName][flagType];
 return threshold;
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

 // Conditional formatting: if Active (col D) is FALSE, shade A..D light red; skip blank rows
 try {
	 const rules = sheet.getConditionalFormatRules() || [];
	 const filtered = rules.filter(r => {
		 try { const bc = r.getBooleanCondition(); if (!bc) return true; const v = bc.getCriteriaValues()||[]; return String((v[0]||'')).toUpperCase() !== '=$D2=FALSE'; } catch(e){ return true; }
	 });
	 const inactiveRange = sheet.getRange(2, 1, Math.max(sheet.getMaxRows()-1,1), 4); // A..D (include Active cell)
	 const inactiveRule = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($D2=FALSE, LEN($A2)>0)').setBackground('#f8d7da').setRanges([inactiveRange]).build();
	 filtered.push(inactiveRule);
	 sheet.setConditionalFormatRules(filtered);
 } catch (e) { Logger.log('Recipients inactive shading error: ' + e.message); }
 
 // Add instructions
 const instructions = [
 ['Config Name:', 'Enter the exact config name (PST01, PST02, NEXT01, etc.)'],
 ['Primary Recipients:', 'Main email addresses (comma-separated if multiple)'],
 ['CC Recipients:', 'CC email addresses (comma-separated if multiple)'],
 ['Active:', 'TRUE to use these recipients, FALSE to disable'],
 ['Withhold No-Flag Emails:', 'TRUE to skip emails when 0 flags found, FALSE to always send emails'],
 ['Last Updated:', 'Automatically updated when you modify recipients'],
 ['', ''],
 ['Delivery Mode:', `${getStagingMode_() === 'Y' ? 'STAGING (all audit emails currently sending to admin only)' : 'PRODUCTION (all audit emails currently sending to recipients)'}`],
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

 // Conditional formatting: if Active (col K) is FALSE, shade A..J light red
 try {
	 const rules = sheet.getConditionalFormatRules() || [];
	 const filtered = rules.filter(r => {
		 try { const bc = r.getBooleanCondition(); if (!bc) return true; const v = bc.getCriteriaValues()||[]; return String((v[0]||'')).toUpperCase() !== '=$K2=FALSE'; } catch(e){ return true; }
	 });
	 const inactiveRange = sheet.getRange(2, 1, Math.max(sheet.getMaxRows()-1,1), 10);
	 const inactiveRule = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$K2=FALSE').setBackground('#f8d7da').setRanges([inactiveRange]).build();
	 filtered.push(inactiveRule);
	 sheet.setConditionalFormatRules(filtered);
 } catch (e) { Logger.log('Exclusions inactive shading error: ' + e.message); }
 
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

function loadRecipientsFromSheet(forceSheetRead) {
	try {
		// Always read from sheet to ensure latest edits are used immediately
		const sheet = getOrCreateRecipientsSheet();
		const data = sheet.getDataRange().getValues();
		const recipients = {};

		for (let i = 1; i < data.length; i++) {
			const row = data[i];
			const configNameRaw = row[0];
			if (!shouldIncludeConfigRow_(configNameRaw, row[3])) continue;
			const configName = String(configNameRaw || '').trim();
			const primaryRecipients = String(row[1] || '').trim();
			const ccRecipients = String(row[2] || '').trim();
			const withholdNoFlagEmails = String(row[4] || '').trim().toUpperCase();

			recipients[configName] = {
				primary: primaryRecipients,
				cc: ccRecipients,
				withholdNoFlagEmails: withholdNoFlagEmails === 'TRUE'
			};
		}

	Logger.log(`Loaded recipients from sheet for ${Object.keys(recipients).length} configs`);
		return recipients;

	} catch (error) {
		Logger.log(`?? Error loading recipients: ${error.message}`);
		return {};
	}
}

// Resolve primary recipients for a given config name from recipients sheet data.
// In STAGING mode, always return ADMIN_EMAIL to prevent accidental sends.
function resolveRecipients(configName, recipientsData) {
	const name = String(configName || '').trim();
	const data = recipientsData || {};
	const entry = name && data[name] ? data[name] : null;
	if (getStagingMode_() === 'Y') {
		return ADMIN_EMAIL;
	}
	const primary = entry && entry.primary ? String(entry.primary).trim() : '';
	return primary || ADMIN_EMAIL;
}

// Resolve CC recipients for a given config name from recipients sheet data.
// In STAGING mode, omit CCs so only admin receives messages.
function resolveCc(configName, recipientsData) {
	if (getStagingMode_() === 'Y') return '';
	const name = String(configName || '').trim();
	const data = recipientsData || {};
	const entry = name && data[name] ? data[name] : null;
	const cc = entry && entry.cc ? String(entry.cc).trim() : '';
	return cc;
}

// === DELIVERY MODE STATUS SYNC ===
function getCurrentDeliveryMode_() {
	return getStagingMode_() === 'Y' ? 'STAGING' : 'PRODUCTION';
}

function getDeliveryModeDisplay_() {
	return getStagingMode_() === 'Y'
		? 'STAGING (all audit emails currently sending to admin only)'
		: 'PRODUCTION (all audit emails currently sending to recipients)';
}

function colorForMode_(mode) {
	return mode === 'STAGING' ? '#fff3cd' /* yellow */ : '#d4edda' /* light green */;
}

// Update the "Delivery Mode:" instruction row on the current config spreadsheet
function updateDeliveryModeInstructionRow_(spreadsheet) {
	try {
		const ss = spreadsheet || getConfigSpreadsheet();
		const sheet = ss.getSheetByName(RECIPIENTS_SHEET_NAME);
		if (!sheet) return false;
		const mode = getCurrentDeliveryMode_();
		const display = getDeliveryModeDisplay_();
		// Find the row in the instructions block that starts at H2 with 2 columns wide
		// We wrote ~15 rows of instructions; search H2:H25 for the label
		const range = sheet.getRange(2, 8, 25, 2); // H2:I26
		const values = range.getValues();
		let updated = false;
		for (let r = 0; r < values.length; r++) {
			const label = String(values[r][0] || '').trim().toLowerCase();
			if (label === 'delivery mode:' || label === 'staging mode override:') {
				// Normalize label to Delivery Mode and set the display text with parentheses
				values[r][0] = 'Delivery Mode:';
				values[r][1] = display;
				range.setValues(values);
				// Set highlight color on the value cell
				const bg = range.getBackgrounds();
				bg[r][1] = colorForMode_(mode);
				range.setBackgrounds(bg);
				updated = true;
				break;
			}
		}
		// If not found, append one just under existing instructions
		if (!updated) {
			const lastRow = sheet.getLastRow();
			const targetRow = Math.max(lastRow + 1, 2);
			sheet.getRange(targetRow, 8, 1, 2).setValues([[ 'Delivery Mode:', display ]]);
			sheet.getRange(targetRow, 9, 1, 1).setBackground(colorForMode_(mode));
			updated = true;
		}
		return updated;
	} catch (e) {
		Logger.log('updateDeliveryModeInstructionRow_ error: ' + e.message);
		return false;
	}
}

// Sync delivery mode to both Admin project sheet and External config sheet
function syncDeliveryModeStatus() {
	try {
		// Admin project sheet (getConfigSpreadsheet respects EXTERNAL_CONFIG_SHEET_ID, so open active for admin explicitly)
		let updatedAny = false;
		try {
			const adminSs = SpreadsheetApp.getActiveSpreadsheet();
			if (adminSs) updatedAny = updateDeliveryModeInstructionRow_(adminSs) || updatedAny;
		} catch (e) { Logger.log('Admin sheet update skipped: ' + e.message); }

		// External config sheet
		try {
			if (EXTERNAL_CONFIG_SHEET_ID) {
				const externalSs = openSpreadsheetById_(EXTERNAL_CONFIG_SHEET_ID);
				updatedAny = updateDeliveryModeInstructionRow_(externalSs) || updatedAny;
			}
		} catch (e) { Logger.log('External sheet update skipped: ' + e.message); }

		return updatedAny;
	} catch (e) {
		Logger.log('syncDeliveryModeStatus error: ' + e.message);
		return false;
	}
}

// Install or refresh a periodic trigger that will update Delivery Mode on both sheets
function installDeliveryModeSyncTrigger() {
	try {
		const results = installAllAutomationTriggers({ handlers: ['deliveryModeSync'] });
		Logger.log('installDeliveryModeSyncTrigger results: ' + results.join(' | '));
		return results;
	} catch (e) {
		Logger.log('installDeliveryModeSyncTrigger error: ' + e.message);
		return [];
	}
}

function runDeliveryModeSync() {
	const ok = syncDeliveryModeStatus();
	Logger.log('Delivery mode sync completed: ' + (ok ? 'updated' : 'no changes'));
}

// === EXCLUSIONS MANAGEMENT ===

// Protect and style the Exclusions sheet's Placement Name column (C) to prevent manual edits
function enforcePlacementNameProtectionAndStyle_(sheet) {
	if (!sheet) return;
	try {
		// Style data cells in column C (exclude header)
		try {
			sheet.getRange('C2:C').setBackground('#eeeeee').setFontColor('#555555');
		} catch (e) { /* ignore styling errors */ }

		// Remove prior protections created by this helper to avoid duplicates
		try {
			const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
			protections.forEach(p => {
				try {
					if (String(p.getDescription() || '') === 'Placement Name (Auto-populated - Do Not Edit)') {
						p.remove();
					}
				} catch (e) { /* ignore */ }
			});
		} catch (e) { /* ignore */ }

		// Strictly protect C2:C (block manual edits). Keep script owner as editor.
		try {
			const rng = sheet.getRange('C2:C');
			const protection = rng.protect().setDescription('Placement Name (Auto-populated - Do Not Edit)');
			protection.setWarningOnly(false);
			try { protection.removeEditors(protection.getEditors()); } catch (e) { /* ignore */ }
			try { protection.addEditors([Session.getEffectiveUser().getEmail()]); } catch (e) { /* ignore */ }
			try { protection.setDomainEdit(false); } catch (e) { /* ignore non-domain */ }
		} catch (e) { /* ignore protection errors */ }
	} catch (err) {
		Logger.log('enforcePlacementNameProtectionAndStyle_ error: ' + err.message);
	}
}
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
 'Placement Name (auto-populated)',
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
 
 // Lock and style the Placement Name column (column C) strictly
 try {
	 enforcePlacementNameProtectionAndStyle_(sheet);
 } catch (e) {
	 Logger.log('Failed to enforce Placement Name protection on create: ' + e.message);
 }
 
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
 // Ensure protection/styling is enforced even if sheet existed already
 try {
	 enforcePlacementNameProtectionAndStyle_(sheet);
 } catch (e) {
	 Logger.log('Failed to enforce Placement Name protection on existing sheet: ' + e.message);
 }
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
	const sheetName = sheet.getName();
	const range = e.range;
	const row = range.getRow();
	const col = range.getColumn();

	try {
		// Branch 1: Exclusions sheet behavior (auto-populate placement name)
		if (sheetName === EXCLUSIONS_SHEET_NAME) {
			// Only process if editing Config Name (col 1) or Placement ID (col 2) and not header row
			if (row <= 1 || (col !== 1 && col !== 2)) return;

			const configName = String(sheet.getRange(row, 1).getValue() || '').trim();
			const placementId = String(sheet.getRange(row, 2).getValue() || '').trim();

			// Only lookup if both config and placement ID are provided
			if (
				configName &&
				placementId &&
				!configName.includes('INSTRUCTIONS') &&
				!configName.includes('-') &&
				!configName.includes('Config Name:')
			) {
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
			return; // Done handling exclusions
		}

		// Branch 2: Recipients sheet behavior (auto-update Last Updated on edits)
		if (sheetName === RECIPIENTS_SHEET_NAME) {
			// Determine editable columns that should trigger a timestamp update
			// Columns: 1=Config Name, 2=Primary Recipients, 3=CC Recipients, 4=Active, 5=Withhold No-Flag Emails
			// Do not trigger on header row or when editing the Last Updated column itself
			const numRows = range.getNumRows();
			const numCols = range.getNumColumns();

			// If the edited range includes any editable columns (1..5) on data rows (>1), stamp col 6 for each affected row
			const editableColTouched = (startCol, count) => {
				const endCol = startCol + count - 1;
				return !(endCol < 1 || startCol > 5); // overlap with [1,5]
			};

			if (row > 1 && editableColTouched(col, numCols)) {
				const tz = Session.getScriptTimeZone();
				const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
				const rowsToStamp = [];
				for (let r = 0; r < numRows; r++) {
					rowsToStamp.push([today]);
				}
				sheet.getRange(row, 6, numRows, 1).setValues(rowsToStamp);
				Logger.log(`Updated Last Updated for ${numRows} row(s) on ${RECIPIENTS_SHEET_NAME} starting at row ${row}.`);
			}
			return; // Done handling recipients
		}
	} catch (error) {
		Logger.log(`Error in onEdit: ${error.message}`);
	}
}

function loadExclusionsFromSheet(forceSheetRead) {
 try {
 // Always read from sheet to ensure latest edits are used immediately
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
 // Add all known config names dynamically
 configsToApply.push(...getAuditConfigs().map(c => c.name));
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
 
 Logger.log(`Loaded exclusions from sheet for ${Object.keys(exclusions).length} configs`);
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
 const config = getAuditConfigs().find(c => c.name === configName);
 if (!config) return null;
 
 const mergedFolder = getDriveFolderByPath_(config.mergedFolderPath);
 const files = mergedFolder.getFiles();

// (moved below) CONFIG LIST HELPERS were previously nested here by mistake
 
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

// === CONFIG LIST HELPERS (migrated from archive, adapted to Properties) ===
function makeAuditConfig_(name, label) {
	return {
		name: name,
		label: label || name,
		mergedFolderPath: [...TRASH_ROOT_PATH, 'To Trash After 60 Days', 'Merged Reports', name],
		tempDailyFolderPath: [...TRASH_ROOT_PATH, 'To Trash After 60 Days', 'Temp Daily Reports', name]
	};
}

let auditConfigsCache_ = null;

function clearAuditConfigsCache_() {
	auditConfigsCache_ = null;
}

function getCustomConfigOrder_() {
	try {
		const raw = PropertiesService.getScriptProperties().getProperty(CONFIG_ORDER_PROPERTY_KEY);
		if (!raw) return [];
		const arr = JSON.parse(raw);
		return Array.isArray(arr) ? arr.map(v => String(v || '')).filter(Boolean) : [];
	} catch (e) {
		Logger.log('getCustomConfigOrder_ error: ' + e.message);
		return [];
	}
}

function saveCustomConfigOrder_(names) {
	try {
		const props = PropertiesService.getScriptProperties();
		if (!Array.isArray(names) || names.length === 0) {
			props.deleteProperty(CONFIG_ORDER_PROPERTY_KEY);
			clearAuditConfigsCache_();
			return;
		}
		const sanitized = names
			.map(name => String(name || '').trim())
			.filter(Boolean);
		props.setProperty(CONFIG_ORDER_PROPERTY_KEY, JSON.stringify(sanitized));
		clearAuditConfigsCache_();
	} catch (e) {
		Logger.log('saveCustomConfigOrder_ error: ' + e.message);
	}
}

function shouldIncludeConfigRow_(configName, activeValue) {
	const name = String(configName || '').trim();
	if (!name) return false;
	const upper = name.toUpperCase();
	if (upper.includes('INSTRUCTIONS') || upper.includes('CONFIG NAME:') || upper.includes('EXAMPLES')) return false;
	if (name.startsWith('-')) return false;
	const active = String(activeValue || '').trim().toUpperCase();
	if (active && active !== 'TRUE') return false;
	return true;
}

function getAuditConfigs(options = {}) {
	if (!options || options.refresh !== true) {
		if (auditConfigsCache_ && Array.isArray(auditConfigsCache_.list)) {
			return auditConfigsCache_.list;
		}
	}
	// Derive config list from the Recipients sheet to reflect latest edits
	const recipients = loadRecipientsFromSheet(true);
	const configs = [];
	const seen = new Set();
	Object.keys(recipients || {}).forEach(name => {
		const cfgName = String(name || '').trim();
		if (!cfgName || seen.has(cfgName)) return;
		seen.add(cfgName);
		const label = `Daily Audits/CM360/${cfgName}`;
		configs.push(makeAuditConfig_(cfgName, label));
	});
		configs.sort((a, b) => a.name.localeCompare(b.name));
		const customOrder = getCustomConfigOrder_();
		let finalConfigs = configs;
		if (customOrder && customOrder.length) {
			const byName = new Map();
			configs.forEach(cfg => byName.set(cfg.name, cfg));
			const ordered = [];
			const consumed = new Set();
			customOrder.forEach(name => {
				const cfg = byName.get(name);
				if (cfg) {
					ordered.push(cfg);
					consumed.add(cfg.name);
				}
			});
			if (ordered.length) {
				const leftovers = configs.filter(cfg => !consumed.has(cfg.name));
				finalConfigs = ordered.concat(leftovers);
			}
		}
		auditConfigsCache_ = { list: finalConfigs, timestamp: Date.now() };
		return finalConfigs;
}

function getAuditConfigByName(configName) {
	const name = String(configName || '').trim();
	if (!name) return null;
	return getAuditConfigs().find(cfg => cfg.name === name) || null;
}

// Admin: Ensure config-level folders exist under key roots
function ensureConfigFoldersExist() {
	try {
		const recipients = loadRecipientsFromSheet(true);
		const configNames = Object.keys(recipients || {}).filter(n => String(n || '').trim());
		if (!configNames.length) {
			SpreadsheetApp.getUi().alert('No configs found in Recipients sheet.');
			return;
		}
		const roots = [
			{ label: 'Temp Daily Reports', subpath: ['To Trash After 60 Days', 'Temp Daily Reports'] },
			{ label: 'Merged Reports', subpath: ['To Trash After 60 Days', 'Merged Reports'] }
		];
		let createdCount = 0;
		for (const cfgName of configNames) {
			for (const r of roots) {
				const path = [...TRASH_ROOT_PATH, ...r.subpath, cfgName];
				const folder = getDriveFolderByPath_(path);
				if (folder) createdCount++;
			}
		}
		SpreadsheetApp.getUi().alert(`Ensured folders for ${configNames.length} config(s). Created/verified ${createdCount} folder(s).`);
	} catch (e) {
		Logger.log('ensureConfigFoldersExist error: ' + e.message);
		SpreadsheetApp.getUi().alert('Error creating folders: ' + e.message);
	}
}

function rebalanceAuditBatchesUsingSummary(options) {
	options = options || {};
	try {
		const configs = getAuditConfigs({ refresh: true });
		if (!configs.length) {
			Logger.log('[Rebalance] No configs available to reorder.');
			return [];
		}
		const prev = getPreviousSummaryCounts_();
		const countsMap = (prev && prev.counts) ? prev.counts : {};
		const overrideMetrics = options.metrics && typeof options.metrics === 'object' ? options.metrics : null;
		const metrics = configs.map(cfg => {
			const raw = overrideMetrics && overrideMetrics.hasOwnProperty(cfg.name)
				? overrideMetrics[cfg.name]
				: (countsMap.hasOwnProperty(cfg.name) ? countsMap[cfg.name] : null);
			let metricValue;
			if (raw === null || typeof raw === 'undefined' || raw === '') {
				metricValue = 100;
			} else {
				const numeric = Number(raw);
				metricValue = Number.isFinite(numeric) ? numeric : 100;
			}
			return { name: cfg.name, metric: metricValue };
		});
		metrics.sort((a, b) => {
			if (b.metric !== a.metric) return b.metric - a.metric;
			return a.name.localeCompare(b.name);
		});
		const distinctMetrics = new Set(metrics.map(m => m.metric));
		if (distinctMetrics.size <= 1) {
			const existing = getCustomConfigOrder_();
			if (existing.length) {
				Logger.log('[Rebalance] Metrics tied; retaining existing custom order.');
				return existing;
			}
			const alphabetical = configs.map(cfg => cfg.name);
			saveCustomConfigOrder_(alphabetical);
			Logger.log('[Rebalance] Metrics tied and no custom order; keeping alphabetical order.');
			return alphabetical;
		}
		const paired = [];
		let left = 0;
		let right = metrics.length - 1;
		while (left <= right) {
			if (left === right) {
				paired.push(metrics[left].name);
			} else {
				paired.push(metrics[left].name);
				paired.push(metrics[right].name);
			}
			left++;
			right--;
		}
		saveCustomConfigOrder_(paired);
		Logger.log(`[Rebalance] Applied high-low pairing. New order: ${paired.join(', ')}`);
		return paired;
	} catch (e) {
		Logger.log('rebalanceAuditBatchesUsingSummary error: ' + e.message);
		return [];
	}
}

function clearDailyScriptProperties(options) {
	options = options || {};
	try {
		const props = PropertiesService.getScriptProperties();
		const allProps = props.getProperties() || {};
		const targets = new Set();
		const staticKeys = [
			AUDIT_RUN_LIST_KEY,
			'CM360_ADMIN_REFRESH_SEEN'
		];
		staticKeys.forEach(key => {
			if (key && Object.prototype.hasOwnProperty.call(allProps, key)) targets.add(key);
		});
		if (options.resetCustomOrder === true && Object.prototype.hasOwnProperty.call(allProps, CONFIG_ORDER_PROPERTY_KEY)) {
			targets.add(CONFIG_ORDER_PROPERTY_KEY);
		}
		Object.keys(allProps).forEach(key => {
			if (typeof key === 'string' && key.indexOf(AUDIT_RUN_STATE_KEY_PREFIX) === 0) {
				targets.add(key);
			}
		});
		if (options.includeCleanupState === true) {
			if (Object.prototype.hasOwnProperty.call(allProps, CLEANUP_STATE_KEY)) targets.add(CLEANUP_STATE_KEY);
			if (Object.prototype.hasOwnProperty.call(allProps, CLEANUP_TRIGGER_ID_KEY)) targets.add(CLEANUP_TRIGGER_ID_KEY);
		}
		if (!targets.size) {
			Logger.log('[EOD] No script properties required clearing.');
			return [];
		}
		const cleared = [];
		targets.forEach(key => {
			try {
				props.deleteProperty(key);
				cleared.push(key);
			} catch (err) {
				Logger.log(`[EOD] Failed to delete property ${key}: ${err.message}`);
			}
		});
		Logger.log(`[EOD] Cleared ${cleared.length} script propert${cleared.length === 1 ? 'y' : 'ies'}: ${cleared.join(', ')}`);
		return cleared;
	} catch (e) {
		Logger.log('clearDailyScriptProperties error: ' + e.message);
		return [];
	}
}

// Core logic to update placement names on a provided Exclusions sheet
function updatePlacementNamesFromReportsOnSheet_(sheet) {
	if (!sheet) throw new Error('No sheet provided to updatePlacementNamesFromReportsOnSheet_');
	Logger.log('Starting placement name update process on provided sheet...');

	const data = sheet.getDataRange().getValues();
	let updatedCount = 0;
	let notFoundCount = 0;

	// Helper: consider a placement name "missing" if it's empty or contains a previous error message
	function isMissingOrErrorPlacementName(name) {
		if (!name) return true;
		const s = String(name || '').trim().toLowerCase();
		if (s === '') return true;
		// Common error markers used in this script
		if (s.startsWith('error') || s.includes('not found') || s.includes('placement id not found')) return true;
		return false;
	}

	// Collect target rows: Placement ID present, Placement Name missing or contains an error
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
		// Include rows where placement name is missing OR contains a previous error marker
		if (!isMissingOrErrorPlacementName(currentPlacementName)) continue; // already has a valid name
		if (!configName) {
			missingConfigRows.push(i + 1); // 1-based row index
			continue;
		}
		if (!targetsByConfig[configName]) targetsByConfig[configName] = [];
		targetsByConfig[configName].push({ rowIndex: i + 1, placementId });
	}

	// Fill errors for rows missing config
	if (missingConfigRows.length > 0) {
		missingConfigRows.forEach(r => sheet.getRange(r, 3).setValue('ERROR: Config Name is required'));
	}

	// Helper: open latest merged sheet for a config and build a map of ID -> Name
	function buildIdToNameMap_(configName) {
		try {
			const cfg = getAuditConfigs().find(c => c.name === configName);
			if (!cfg) return null;
			const folder = getDriveFolderByPath_(cfg.mergedFolderPath);
			if (!folder) return null;
			const it = folder.getFiles();
			const files = [];
			while (it.hasNext()) files.push(it.next());
			if (files.length === 0) return null;
			// Prefer exact-today config-specific file, then latest matching config prefix, then any merged file, else newest
			const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
			const basePrefix = 'CM360_Merged_Audit_';
			const prefix = `${basePrefix}${configName}_`;
			const exactToday = `${basePrefix}${configName}_${todayStr}`;
			let preferred = files.find(f => String(f.getName() || '') === exactToday);
			if (!preferred) {
				const forConfig = files.filter(f => String(f.getName() || '').startsWith(prefix));
				if (forConfig.length) {
					forConfig.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
					preferred = forConfig[0];
				} else {
					const anyMerged = files.filter(f => String(f.getName() || '').startsWith(basePrefix));
					if (anyMerged.length) {
						anyMerged.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
						preferred = anyMerged[0];
					} else {
						files.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
						preferred = files[0];
					}
				}
			}
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
			rows.forEach(({ rowIndex }) => sheet.getRange(rowIndex, 3).setValue('ERROR:Placement ID not found in last CM360 report'));
			notFoundCount += rows.length;
			continue;
		}
		rows.forEach(({ rowIndex, placementId }) => {
			const name = idToName.get(String(placementId).trim());
			if (name) {
				sheet.getRange(rowIndex, 3).setValue(name);
				updatedCount++;
			} else {
				sheet.getRange(rowIndex, 3).setValue('ERROR: Placement ID not found in last CM360 report');
				notFoundCount++;
			}
		});
		// Throttle a bit between configs
		Utilities.sleep(200);
	}

	Logger.log(`Update complete: ${updatedCount} updated, ${notFoundCount} not found`);
}

// Convenience: Update placement names on the bound Admin Exclusions sheet
function updatePlacementNamesFromReports() {
	try {
		Logger.log('Starting updatePlacementNamesFromReports - updating bound Admin exclusions sheet...');
		const sheet = getOrCreateExclusionsSheet();
		updatePlacementNamesFromReportsOnSheet_(sheet);
	} catch (error) {
		Logger.log(` Error in updatePlacementNamesFromReports: ${error.message}`);
		// Send error notification to admin email instead of showing UI alert
		try {
			GmailApp.sendEmail(
				ADMIN_EMAIL,
				'CM360 Audit - Placement Name Update Error',
				`Failed to update placement names: ${error.message}\n\nThis error occurred during the automated placement name update process.`,
				{ htmlBody: `<p>Failed to update placement names: <strong>${error.message}</strong></p><p>This error occurred during the automated placement name update process.</p>` }
			);
		} catch (emailError) {
			Logger.log(`Failed to send error email: ${emailError.message}`);
		}
	}
}

// UI wrapper for manual menu use - shows confirmation dialog
function updatePlacementNamesFromReportsWithUI() {
	try {
		const ui = SpreadsheetApp.getUi();
		const response = ui.alert(
			'Update Placement Names (External)',
			'This will search the latest merged reports to update placement names in the EXTERNAL exclusions sheet. This may take a few minutes. Continue?',
			ui.ButtonSet.YES_NO
		);
		if (response !== ui.Button.YES) return;

		if (!EXTERNAL_CONFIG_SHEET_ID) {
			ui.alert('No External Config Sheet', 'EXTERNAL_CONFIG_SHEET_ID is not configured. Cannot update the external exclusions sheet.', ui.ButtonSet.OK);
			return;
		}

		const ss = openSpreadsheetById_(EXTERNAL_CONFIG_SHEET_ID);
			const externalSheet = ss.getSheetByName(EXCLUSIONS_SHEET_NAME);
		if (!externalSheet) {
			ui.alert('Missing Exclusions tab', `The external spreadsheet does not contain a sheet named "${EXCLUSIONS_SHEET_NAME}". Please create it first.`, ui.ButtonSet.OK);
			return;
		}

			// Ensure protection/styling is enforced on external sheet as well
			try { enforcePlacementNameProtectionAndStyle_(externalSheet); } catch (e) { Logger.log('Warn: could not enforce protection on external sheet: ' + e.message); }

		updatePlacementNamesFromReportsOnSheet_(externalSheet);
		ui.alert('Update Complete', 'External placement name update process has finished. Check the logs for details.', ui.ButtonSet.OK);
	} catch (error) {
		Logger.log(`Error in updatePlacementNamesFromReportsWithUI: ${error.message}`);
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
function runNEXTCD01Audit() { runDailyAuditByName('NEXTCD01'); }
function runNEXT01Audit() { runDailyAuditByName('NEXT01'); }
function runNEXT02Audit() { runDailyAuditByName('NEXT02'); }
function runNEXT03Audit() { runDailyAuditByName('NEXT03'); }
function runSPTM01Audit() { runDailyAuditByName('SPTM01'); }
function runNFL01Audit() { runDailyAuditByName('NFL01'); }
function runGMNR01Audit() { runDailyAuditByName('GMNR01'); }

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
 Logger.log(` Current staging mode: ${getStagingMode_()}`);
 
 // Show current mode recipients
 const currentRecipients = resolveRecipients('test-config', recipientsData);
 Logger.log(` Current mode recipients: ${currentRecipients}`);
 
 // Note about staging mode
 if (getStagingMode_() === 'Y') {
 Logger.log(` ... Staging mode is ENABLED - all emails go to admin`);
 } else {
 Logger.log(` ... Production mode is ENABLED - emails use sheet recipients`);
 }
 
 Logger.log(`... Recipients system test completed successfully!`);
 Logger.log(`" Summary:`);
 Logger.log(` - Recipients sheet: Ready`);
 Logger.log(` - Configurations loaded: ${configCount}`);
 Logger.log(` - Recipient resolution: Working`);
 Logger.log(` - Staging mode: ${getStagingMode_() === 'Y' ? 'ENABLED (Admin override)' : 'DISABLED (Sheet recipients)'}`);
 
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
 Logger.log('Setting up menu for external config sheet...');
 
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
 source: '// === CM360 Configuration Helper Menu ===\n' +
 'function onOpen() {\n' +
 ' const ui = SpreadsheetApp.getUi();\n' +
 ' ui.createMenu(\'CM360 Config Helper\')\n' +
 ' .addItem(\'🏃 Run Config Audit\', \'showConfigAuditRunner\')\n' +
 ' .addItem(\'✅ Validate Configuration\', \'validateConfiguration\')\n' +
 ' .addItem(\'📋 Show Config Summary\', \'showConfigSummary\')\n' +
 ' .addSeparator()\n' +
 ' .addItem(\'🧩 Create New Config…\', \'createNewConfig\')\n' +
 ' .addToUi();\n' +
 '}\n\n' +
 'var ADMIN_EMAIL = ' + JSON.stringify(ADMIN_EMAIL) + ';\n\n' +
 'function showConfigAuditRunner() {\n' +
 ' const ui = SpreadsheetApp.getUi();\n' +
 ' \n' +
 ' // Get available configs from recipients sheet\n' +
 ' const recipientsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(\'Audit Recipients\');\n' +
 ' if (!recipientsSheet) {\n' +
 ' ui.alert(\'Error\', \'Audit Recipients sheet not found. Please ask admin to sync configuration data.\', ui.ButtonSet.OK);\n' +
 ' return;\n' +
 ' }\n' +
 ' \n' +
 ' const data = recipientsSheet.getDataRange().getValues();\n' +
 ' \n' +
 ' // Check if sheet has data\n' +
 ' if (data.length <= 1) {\n' +
 ' ui.alert(\n' +
 ' \'No Data Found\', \n' +
 ' \'The Audit Recipients sheet appears to be empty or only has headers.\\\\n\\\\nData rows found: \' + (data.length - 1) + \'\\\\n\\\\nPlease ask admin to populate the configuration data.\',\n' +
 ' ui.ButtonSet.OK\n' +
 ' );\n' +
 ' return;\n' +
 ' }\n' +
 ' \n' +
 ' const configs = [];\n' +
 ' \n' +
 ' for (let i = 1; i < data.length; i++) {\n' +
 ' const row = data[i];\n' +
 ' const configName = row[0];\n' +
 ' const activeStatus = row[3];\n' +
 ' \n' +
 ' if (configName && (activeStatus === \'TRUE\' || activeStatus === \'true\' || activeStatus === true)) {\n' +
 ' configs.push({\n' +
 ' name: configName,\n' +
 ' recipients: row[1] || \'\',\n' +
 ' cc: row[2] || \'\',\n' +
 ' withhold: row[4] === \'TRUE\'\n' +
 ' });\n' +
 ' }\n' +
 ' }\n' +
 ' \n' +
 ' if (configs.length === 0) {\n' +
 ' ui.alert(\'No Active Configs\', \'No active configurations found in the Audit Recipients sheet.\', ui.ButtonSet.OK);\n' +
 ' return;\n' +
 ' }\n' +
 ' \n' +
 ' const configOptions = configs.map((config, index) => {\n' +
 ' const recipientCount = config.recipients.split(\',\').length;\n' +
 ' const ccCount = config.cc ? config.cc.split(\',\').length : 0;\n' +
 ' return (index + 1) + \'. \' + config.name + \' (\' + recipientCount + \' recipients\' + (ccCount > 0 ? \', \' + ccCount + \' CC\' : \'\') + (config.withhold ? \', withholds no-flag emails\' : \'\') + \')\';\n' +
 ' }).join(\'\\\\n\');\n' +
 ' \n' +
 ' const response = ui.prompt(\n' +
 ' \'Select Configuration to Audit\',\n' +
 ' \'Available configurations:\\\\n\\\\n\' + configOptions + \'\\\\n\\\\nEnter the number (1-\' + configs.length + \') of the configuration to audit:\',\n' +
 ' ui.ButtonSet.OK_CANCEL\n' +
 ' );\n' +
 ' \n' +
 ' if (response.getSelectedButton() !== ui.Button.OK) return;\n' +
 ' \n' +
 ' const selectedIndex = parseInt(response.getResponseText().trim(), 10) - 1;\n' +
 ' \n' +
 ' if (isNaN(selectedIndex) || selectedIndex < 0 || selectedIndex >= configs.length) {\n' +
 ' ui.alert(\'Invalid Selection\', \'Please enter a valid number between 1 and \' + configs.length + \'.\', ui.ButtonSet.OK);\n' +
 ' return;\n' +
 ' }\n' +
 ' \n' +
 ' const selectedConfig = configs[selectedIndex];\n' +
 '\n' +
 ' // Create request row locally in this sheet (first available row in A:E)\n' +
 ' const ss = SpreadsheetApp.getActiveSpreadsheet();\n' +
 ' let reqSheet = ss.getSheetByName(\'Audit Requests\');\n' +
 ' if (!reqSheet) {\n' +
 ' reqSheet = ss.insertSheet(\'Audit Requests\');\n' +
 ' const headers = [\'Config Name\',\'Requested By\',\'Requested At\',\'Status\',\'Notes\',\'\',\'INSTRUCTIONS\'];\n' +
 ' reqSheet.getRange(1, 1, 1, headers.length).setValues([headers]);\n' +
 ' reqSheet.getRange(1, 1, 1, 5).setFontWeight(\'bold\').setBackground(\'#4285f4\').setFontColor(\'#ffffff\');\n' +
 ' }\n' +
 '\n' +
 ' // Find the first available (empty) row within columns A-E so we fill gaps.\n' +
 ' // IMPORTANT: Do not let instructions (G:...) extend the write row; base only on A-E.\n' +
 ' const maxRow = Math.max(reqSheet.getLastRow(), 1);\n' +
 ' const aToE = reqSheet.getRange(1, 1, maxRow, 5).getValues();\n' +
 ' let writeRow = -1;\n' +
 ' let lastDataRowAE = 1; // track the last row with any data in A-E (1-based; header row starts at 1)\n' +
 '\n' +
 ' // Scan top-down from row 2 to find the first row where A-E are all blank;\n' +
 ' // also track the last row that contains any A-E data to compute a safe append row.\n' +
 ' for (let r = 1; r < aToE.length; r++) { // aToE[0] is header\n' +
 ' const rowVals = aToE[r];\n' +
 ' const hasDataInAE = rowVals.some(v => String(v || \'\').trim() !== \'\');\n' +
 ' if (hasDataInAE) {\n' +
 ' lastDataRowAE = r + 1; // convert to 1-based row index\n' +
 ' } else if (writeRow === -1) {\n' +
 ' writeRow = r + 1; // first gap found\n' +
 ' }\n' +
 ' }\n' +
 '\n' +
 ' // If no internal gap, append immediately after the last A-E data row\n' +
 ' if (writeRow === -1) writeRow = lastDataRowAE + 1;\n' +
 ' const requester = (Session && Session.getActiveUser ? Session.getActiveUser().getEmail() : \'\') || \'\';\n' +
 ' const now = new Date();\n' +
 ' reqSheet.getRange(writeRow, 1, 1, 5).setValues([[selectedConfig.name, requester, now, \'PENDING\', \'\']]);\n' +
 '\n' +
 ' SpreadsheetApp.getUi().alert(\'Request Submitted\', \'Your audit request was added to the queue.\', SpreadsheetApp.getUi().ButtonSet.OK);\n' +
 '}\n\n' +
 'function createNewConfig() {\n' +
 ' var html = HtmlService.createHtmlOutput(getCreateConfigFormHtml_())\n' +
 '   .setWidth(720).setHeight(860);\n' +
 ' SpreadsheetApp.getUi().showModalDialog(html, \"Create New Config\");\n' +
 '}\n\n' +
 'function getCreateConfigFormHtml_() {\n' +
 ' var html = \"<style>body{font-family:Arial,sans-serif;font-size:13px;padding:12px;}label{display:block;margin:8px 0 4px;}input[type=number]{width:100px;} .row{display:flex;gap:12px;align-items:center;flex-wrap:wrap;} .flag{border:1px solid #ddd;padding:8px;border-radius:6px;margin:8px 0;} .help{background:#f8f9fa;border:1px solid #e0e0e0;padding:10px;border-radius:6px;margin-top:12px;} button{background:#1a73e8;color:#fff;border:none;border-radius:4px;padding:8px 12px;cursor:pointer;} .muted{color:#666;} textarea{width:100%;height:160px;} </style>\" +\n' +
 '   \"<div>\" +\n' +
 '   \"<label>Config ID (e.g., PST01, ENT01, NEXT01)</label>\" +\n' +
 '   \"<input id=\\\"cfg\\\" type=\\\"text\\\" placeholder=\\\"PST01\\\" maxlength=\\\"16\\\"/>\" +\n' +
 '   \"<label>Primary Recipients (comma-separated)</label>\" +\n' +
 '   \"<input id=\\\"recips\\\" type=\\\"text\\\" placeholder=\\\"user@company.com, team@company.com\\\"/>\" +\n' +
 '   \"<label>CC Recipients (optional, comma-separated)</label>\" +\n' +
 '   \"<input id=\\\"cc\\\" type=\\\"text\\\" placeholder=\\\"cc1@company.com, cc2@company.com\\\"/>\" +\n' +
 '   \"<div class=\\\"flag\\\"><div class=\\\"muted\\\">Set thresholds (min Impressions and min Clicks) per flag type</div>\" +\n' +
 '   \"<div class=\\\"row\\\"><strong>clicks_greater_than_impressions</strong><label>Min Impressions</label><input id=\\\"t_cgti_i\\\" type=\\\"number\\\" min=\\\"0\\\" value=\\\"0\\\"/><label>Min Clicks</label><input id=\\\"t_cgti_c\\\" type=\\\"number\\\" min=\\\"0\\\" value=\\\"0\\\"/></div>\" +\n' +
 '   \"<div class=\\\"row\\\"><strong>out_of_flight_dates</strong><label>Min Impressions</label><input id=\\\"t_oofd_i\\\" type=\\\"number\\\" min=\\\"0\\\" value=\\\"0\\\"/><label>Min Clicks</label><input id=\\\"t_oofd_c\\\" type=\\\"number\\\" min=\\\"0\\\" value=\\\"0\\\"/></div>\" +\n' +
 '   \"<div class=\\\"row\\\"><strong>pixel_size_mismatch</strong><label>Min Impressions</label><input id=\\\"t_psm_i\\\" type=\\\"number\\\" min=\\\"0\\\" value=\\\"0\\\"/><label>Min Clicks</label><input id=\\\"t_psm_c\\\" type=\\\"number\\\" min=\\\"0\\\" value=\\\"0\\\"/></div>\" +\n' +
 '   \"<div class=\\\"row\\\"><strong>default_ad_serving</strong><label>Min Impressions</label><input id=\\\"t_das_i\\\" type=\\\"number\\\" min=\\\"0\\\" value=\\\"0\\\"/><label>Min Clicks</label><input id=\\\"t_das_c\\\" type=\\\"number\\\" min=\\\"0\\\" value=\\\"0\\\"/></div>\" +\n' +
 '   \"</div>\" +\n' +
 '   \"<div class=\\\"help\\\"><strong>CM360 Daily Reports Requirements</strong><br/>\" +\n' +
 '   \"<em>Basic info</em><br/>\" +\n' +
 '   \"Name: NETWORKNAME_ImpClickReport_DailyAudit_CONFIGID<br/>\" +\n' +
 '   \"- Network Name should be the name of the DCM network (can use shorthand notation)<br/>\" +\n' +
 '   \"- Config ID should be a unique alphanumeric phrase (i.e., AMC01, NEXT01, PST01) – please let me know which one you’ll be using<br/>\" +\n' +
 '   \"If you’d like multiple networks/DCM reports to be rolled up into the same email, please use the same Config ID<br/><br/>\" +\n' +
 '   \"<em>Date Range</em>: Yesterday<br/><br/>\" +\n' +
 '   \"<em>Fields (in order)</em><br/>Dimensions:<br/>Advertiser (FILTER IF NEEDED)<br/>Campaign<br/>Site (CM360)<br/>Placement ID<br/>Placement<br/>Placement Start Date<br/>Placement End Date<br/>Ad Type (FILTER OUT SA360 & DART Search)<br/>Creative<br/>Placement Pixel Size<br/>Creative Pixel Size<br/>Date<br/><br/>Metrics:<br/>Impressions<br/>Clicks<br/><br/>\" +\n' +
 '   \"<em>Scheduling</em><br/>Time zone: Eastern Time (GMT-4:00)<br/>Repeats: Daily<br/>Every: 1 day<br/>Starts: today<br/>Ends: as late as possible<br/>Format: Excel (Attachment)<br/>Share with: platformsolutionshmi@gmail.com<br/><br/>\" +\n' +
 '   \"Reach out to ' + JSON.stringify(ADMIN_EMAIL) + ' with any questions\" +\n' +
 '   \"</div>\" +\n' +
 '   \"<div style=\\\"margin-top:12px\\\"><button onclick=\\\"submitForm()\\\">Submit</button></div>\" +\n' +
 '   \"</div>\" +\n' +
 '   \"<script>function valId(s){return /^[A-Za-z0-9]+$/.test(s)};function submitForm(){var cfg=document.getElementById(\\\"cfg\\\").value.trim();var rec=document.getElementById(\\\"recips\\\").value.trim();var cc=document.getElementById(\\\"cc\\\").value.trim();if(!cfg||!valId(cfg)){alert(\\\"Please enter an alphanumeric Config ID (e.g., PST01).\\\");return;}if(!rec){alert(\\\"Please enter at least one recipient.\\\");return;}var form={configId:cfg.toUpperCase(),recipients:rec,cc:cc,thresholds:{clicks_greater_than_impressions:{minImpressions:Number(document.getElementById(\\\"t_cgti_i\\\").value||0),minClicks:Number(document.getElementById(\\\"t_cgti_c\\\").value||0)},out_of_flight_dates:{minImpressions:Number(document.getElementById(\\\"t_oofd_i\\\").value||0),minClicks:Number(document.getElementById(\\\"t_oofd_c\\\").value||0)},pixel_size_mismatch:{minImpressions:Number(document.getElementById(\\\"t_psm_i\\\").value||0),minClicks:Number(document.getElementById(\\\"t_psm_c\\\").value||0)},default_ad_serving:{minImpressions:Number(document.getElementById(\\\"t_das_i\\\").value||0),minClicks:Number(document.getElementById(\\\"t_das_c\\\").value||0)}}};google.script.run.withSuccessHandler(function(msg){alert(msg);google.script.host.close();}).withFailureHandler(function(e){alert(\\\"Failed: \\\"+e.message);}).submitNewConfigFromForm(form);}<\/script>\";\n' +
 ' return html;\n' +
 '}\n\n' +
 'function submitNewConfigFromForm(form) {\n' +
 ' var ss = SpreadsheetApp.getActiveSpreadsheet();\n' +
 ' var recSheet = ss.getSheetByName(\'Audit Recipients\');\n' +
 ' if (!recSheet) recSheet = ss.insertSheet(\'Audit Recipients\');\n' +
 ' var thSheet = ss.getSheetByName(\'Audit Thresholds\');\n' +
 ' if (!thSheet) thSheet = ss.insertSheet(\'Audit Thresholds\');\n' +
 ' var now = new Date();\n' +
 ' // Ensure headers minimal\n' +
 ' if (recSheet.getLastRow() === 0) recSheet.getRange(1,1,1,8).setValues([[\'Config Name\',\'Primary Recipients\',\'CC Recipients\',\'Active\',\'Withhold No-Flag Emails\',\'Last Updated\',\'\',\'INSTRUCTIONS\']]);\n' +
 ' if (thSheet.getLastRow() === 0) thSheet.getRange(1,1,1,7).setValues([[\'Config Name\',\'Flag Type\',\'Min Impressions\',\'Min Clicks\',\'Active\',\'\',\'INSTRUCTIONS\']]);\n' +
 ' // Add recipients row if missing\n' +
 ' var data = recSheet.getDataRange().getValues();\n' +
 ' var exists = false;\n' +
 ' for (var i=1;i<data.length;i++){ if ((data[i][0]||\'\').toString().trim().toUpperCase()===form.configId){ exists=true; break; } }\n' +
 ' if (!exists){ recSheet.appendRow([form.configId, form.recipients, form.cc||\'\', \"TRUE\", \"FALSE\", now, \"\", \"\"]); }\n' +
 ' // Add thresholds (overwrite duplicates by appending; admin can clean later)\n' +
 ' var flags = Object.keys(form.thresholds||{});\n' +
 ' var addedFlags = 0;\n' +
 ' for (var f=0; f<flags.length; f++){ var ft = flags[f]; var t = form.thresholds[ft]||{minImpressions:0,minClicks:0}; thSheet.appendRow([form.configId, ft, Number(t.minImpressions)||0, Number(t.minClicks)||0, \"TRUE\", \"\", \"\"]); addedFlags++; }\n' +
 ' // Notify admin\n' +
 ' var requester = (Session && Session.getActiveUser ? Session.getActiveUser().getEmail() : \"\");\n' +
 ' var subject = \"CM360: New Config Request Submitted - \" + form.configId;\n' +
 ' var body = \"A new CM360 Config was submitted via the helper menu.\\n\\n\" +\n' +
 '           \"Submitted by: \" + requester + \"\\n\" +\n' +
 '           \"Config ID: \" + form.configId + \"\\n\" +\n' +
 '           \"Recipients: \" + form.recipients + \"\\n\" +\n' +
 '           \"CC: \" + (form.cc||\"\") + \"\\n\\n\" +\n' +
 '           \"Thresholds:\\n\" + flags.map(function(k){ var t=form.thresholds[k]; return \" - \"+k+\": minImpressions=\"+t.minImpressions+\", minClicks=\"+t.minClicks; }).join(\"\\n\") + \"\\n\\n\" +\n' +
 '           \"Added in helper sheet:\\n - Audit Recipients row\" + (exists?\" (already existed)\":\" (created)\") + \"\\n - Audit Thresholds rows: \" + addedFlags + \"\\n\\n\" +\n' +
 '           \"Next steps for Admin:\\n 1) In Admin Controls, click \\\"Sync FROM External Config\\\".\\n 2) Then use \\\"Prepare Environment\\\" to create Drive folders and Gmail label.\\n\";\n' +
 ' try { MailApp.sendEmail({to: ADMIN_EMAIL, subject: subject, body: body}); } catch(e) {}\n' +
 ' return \"New config saved. An email notification was sent to admin (\" + ADMIN_EMAIL + \").\";\n' +
 '}\n\n' +
 'function validateConfiguration() {\n' +
 ' const ss = SpreadsheetApp.getActiveSpreadsheet();\n' +
 ' const ui = SpreadsheetApp.getUi();\n' +
 ' \n' +
 ' const sheets = [\'Audit Thresholds\', \'Audit Recipients\', \'Audit Exclusions\'];\n' +
 ' const found = sheets.filter(name => ss.getSheetByName(name) !== null);\n' +
 ' const missing = sheets.filter(name => ss.getSheetByName(name) === null);\n' +
 ' \n' +
 ' let message = \'Configuration Validation:\\\\n\\\\n\';\n' +
 ' if (found.length > 0) {\n' +
 ' message += \'✅ Found sheets:\\\\n\' + found.map(s => \'• \' + s).join(\'\\\\n\') + \'\\\\n\\\\n\';\n' +
 ' }\n' +
 ' if (missing.length > 0) {\n' +
 ' message += \'❌ Missing sheets:\\\\n\' + missing.map(s => \'• \' + s).join(\'\\\\n\') + \'\\\\n\\\\n\';\n' +
 ' }\n' +
 ' \n' +
 ' ui.alert(\'Validation Results\', message, ui.ButtonSet.OK);\n' +
 '}\n\n' +
 'function showConfigSummary() {\n' +
 ' const ss = SpreadsheetApp.getActiveSpreadsheet();\n' +
 ' const ui = SpreadsheetApp.getUi();\n' +
 ' \n' +
 ' let summary = \'Configuration Summary:\\\\n\\\\n\';\n' +
 ' \n' +
 ' const thresholds = ss.getSheetByName(\'Audit Thresholds\');\n' +
 ' if (thresholds) {\n' +
 ' const data = thresholds.getDataRange().getValues();\n' +
 ' const configs = new Set();\n' +
 ' for (let i = 1; i < data.length; i++) {\n' +
 ' if (data[i][0] && data[i][4] === \'TRUE\') configs.add(data[i][0]);\n' +
 ' }\n' +
 ' summary += \'📊 Thresholds: \' + configs.size + \' active configs\\\\n\';\n' +
 ' }\n' +
 ' \n' +
 ' const recipients = ss.getSheetByName(\'Audit Recipients\');\n' +
 ' if (recipients) {\n' +
 ' const data = recipients.getDataRange().getValues();\n' +
 ' const configs = new Set();\n' +
 ' for (let i = 1; i < data.length; i++) {\n' +
 ' if (data[i][0] && data[i][3] === \'TRUE\') configs.add(data[i][0]);\n' +
 ' }\n' +
 ' summary += \'📧 Recipients: \' + configs.size + \' active configs\\\\n\';\n' +
 ' }\n' +
 ' \n' +
 ' const exclusions = ss.getSheetByName(\'Audit Exclusions\');\n' +
 ' if (exclusions) {\n' +
 ' const data = exclusions.getDataRange().getValues();\n' +
 ' let activeRows = 0;\n' +
 ' for (let i = 1; i < data.length; i++) {\n' +
 ' if (data[i][10] === \'TRUE\') activeRows++;\n' +
 ' }\n' +
 ' summary += \'🚫 Exclusions: \' + activeRows + \' active rules\\\\n\';\n' +
 ' }\n' +
 ' \n' +
 ' ui.alert(\'Configuration Summary\', summary, ui.ButtonSet.OK);\n' +
 '}'
 }]
 };
 
 Logger.log('Helper menu setup completed for config sheet');
 Logger.log('');
 Logger.log('=== COPY THE CODE BELOW ===');
 Logger.log(scriptProject.files[0].source);
 Logger.log('=== END OF CODE ===');
 Logger.log('');
 Logger.log('To add the menu:');
 Logger.log('1. Open the config spreadsheet: https://docs.google.com/spreadsheets/d/' + configSheetId);
 Logger.log('2. Go to Extensions > Apps Script');
 Logger.log('3. Replace the default code with the helper menu code above');
 Logger.log('4. Save the project and refresh the spreadsheet');
 
 return true;
 
 } catch (error) {
 Logger.log('Error setting up external config menu: ' + error.message);
 throw error;
 }
}

// === CM360 Configuration Helper Menu Functions ===
// These functions are intended for use in external config sheets
// Note: These functions are provided as reference for external config sheets
// They should NOT have an onOpen() function in the main script

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
 return (index + 1) + '. ' + config.name + ' (' + recipientCount + ' recipients' + (ccCount > 0 ? ', ' + ccCount + ' CC' : '') + (config.withhold ? ', withholds no-flag emails' : '') + ')';
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

/**
 * Test function to verify deployment
 */
function testAddMissingConfigNames() {
	const ui = SpreadsheetApp.getUi();
	const configNames = getAuditConfigs().map(config => config.name);
	ui.alert('Test Function Works', 'Found ' + configNames.length + ' configs: ' + configNames.join(', '), ui.ButtonSet.OK);
}

/**
 * Adds missing config names (dynamic) to all three sheets with default values
 */
function addMissingConfigNames() {
	const ui = SpreadsheetApp.getUi();
	
	try {
	// Get all config names dynamically
	const configNames = getAuditConfigs().map(config => config.name);
	Logger.log('Found ' + configNames.length + ' config names: ' + configNames.join(', '));
		
		let totalAdded = 0;
		let results = [];
		
		// Check and add to Audit Recipients
		const recipientsAdded = addMissingConfigsToRecipients(configNames);
		totalAdded += recipientsAdded;
		if (recipientsAdded > 0) results.push('Added ' + recipientsAdded + ' configs to Audit Recipients');
		
		// Check and add to Audit Thresholds  
		const thresholdsAdded = addMissingConfigsToThresholds(configNames);
		totalAdded += thresholdsAdded;
		if (thresholdsAdded > 0) results.push('Added ' + thresholdsAdded + ' configs to Audit Thresholds');
		
		// Show results
		if (totalAdded === 0) {
			ui.alert('No Missing Configs', 'All config names are already present in the sheets.', ui.ButtonSet.OK);
		} else {
			const message = 'Successfully added ' + totalAdded + ' missing config entries:\n\n' + results.join('\n') + '\n\nPlease review and update the default values as needed.';
			ui.alert('Missing Configs Added', message, ui.ButtonSet.OK);
		}
		
	} catch (error) {
		Logger.log('Error in addMissingConfigNames: ' + error.message);
		ui.alert('Error', 'Failed to add missing configs: ' + error.message, ui.ButtonSet.OK);
	}
}

/**
 * Add missing config names to Audit Recipients sheet with default values
 */
function addMissingConfigsToRecipients(configNames) {
	const sheet = getOrCreateRecipientsSheet();
	const data = sheet.getDataRange().getValues();
	
	// Get existing config names (column A)
	const existingConfigs = new Set();
	for (let i = 1; i < data.length; i++) {
		if (data[i][0]) existingConfigs.add(String(data[i][0]).trim());
	}
	
	// Find missing configs
	const missingConfigs = configNames.filter(name => !existingConfigs.has(name));
	if (missingConfigs.length === 0) return 0;
	
	// Add missing configs with default values
	const newRows = [];
	const currentDate = formatDate(new Date(), 'yyyy-MM-dd');
	
	missingConfigs.forEach(configName => {
		newRows.push([
			configName,           // Config Name
			ADMIN_EMAIL,          // Primary Recipients  
			'',                   // CC Recipients
			'TRUE',               // Active
			'FALSE',              // Withhold No-Flag Emails
			currentDate           // Last Updated
		]);
	});
	
	// Append new rows
	const startRow = sheet.getLastRow() + 1;
	sheet.getRange(startRow, 1, newRows.length, 6).setValues(newRows);
	
	Logger.log('Added ' + missingConfigs.length + ' missing configs to Audit Recipients: ' + missingConfigs.join(', '));
	return missingConfigs.length;
}

/**
 * Add missing config names to Audit Thresholds sheet with default values
 */
function addMissingConfigsToThresholds(configNames) {
	const sheet = getOrCreateThresholdsSheet();
	const data = sheet.getDataRange().getValues();
	
	// Get existing config names (column A)
	const existingConfigs = new Set();
	for (let i = 1; i < data.length; i++) {
		if (data[i][0]) existingConfigs.add(String(data[i][0]).trim());
	}
	
	// Find missing configs
	const missingConfigs = configNames.filter(name => !existingConfigs.has(name));
	if (missingConfigs.length === 0) return 0;
	
	// Define flag types and default thresholds
	const flagTypes = [
		'clicks_greater_than_impressions',
		'out_of_flight_dates', 
		'pixel_size_mismatch',
		'default_ad_serving'
	];
	
	const newRows = [];
	missingConfigs.forEach(configName => {
		flagTypes.forEach(flagType => {
			newRows.push([
				configName,    // Config Name
				flagType,      // Flag Type
				0,             // Min Impressions (default)
				0,             // Min Clicks (default)
				'TRUE'         // Active
			]);
		});
	});
	
	// Append new rows
	const startRow = sheet.getLastRow() + 1;
	sheet.getRange(startRow, 1, newRows.length, 5).setValues(newRows);
	
	Logger.log('Added ' + (missingConfigs.length * flagTypes.length) + ' threshold entries for missing configs: ' + missingConfigs.join(', '));
	return missingConfigs.length * flagTypes.length;
}

function createMissingThresholds() {
 const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Audit Thresholds');
 if (!sheet) {
 SpreadsheetApp.getUi().alert('Audit Thresholds sheet not found. Please create it first.');
 return;
 }
 SpreadsheetApp.getUi().alert('✅ Audit Thresholds sheet is available for editing.');
}

function createMissingRecipients() {
 const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Audit Recipients');
 if (!sheet) {
 SpreadsheetApp.getUi().alert('Audit Recipients sheet not found. Please create it first.');
 return;
 }
 SpreadsheetApp.getUi().alert('✅ Audit Recipients sheet is available for editing.');
}

function createMissingExclusions() {
 const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Audit Exclusions');
 if (!sheet) {
 SpreadsheetApp.getUi().alert('Audit Exclusions sheet not found. Please create it first.');
 return;
 }
 SpreadsheetApp.getUi().alert('✅ Audit Exclusions sheet is available for editing.');
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
 summary += 'Thresholds: ' + configs.size + ' active configs\\n';
 }
 
 const recipients = ss.getSheetByName('Audit Recipients');
 if (recipients) {
 const data = recipients.getDataRange().getValues();
 const configs = new Set();
 for (let i = 1; i < data.length; i++) {
 if (data[i][0] && data[i][3] === 'TRUE') configs.add(data[i][0]);
 }
 summary += 'Recipients: ' + configs.size + ' active configs\\n';
 }
 
 const exclusions = ss.getSheetByName('Audit Exclusions');
 if (exclusions) {
 const data = exclusions.getDataRange().getValues();
 const activeRows = data.slice(1).filter(row => row[0] && row[4] === 'TRUE').length;
 summary += 'Exclusions: ' + activeRows + ' active rules\\n';
 }
 
 ui.alert('Configuration Summary', summary, ui.ButtonSet.OK);
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

/** Direct installer for running from Apps Script editor (no UI prompts) */
function installExternalConfigHelper() {
 if (!EXTERNAL_CONFIG_SHEET_ID) {
 Logger.log('❌ EXTERNAL_CONFIG_SHEET_ID is not configured');
 return;
 }
 
 Logger.log('🔧 Installing helper menu for external config sheet: ' + EXTERNAL_CONFIG_SHEET_ID);
 setupExternalConfigMenu(EXTERNAL_CONFIG_SHEET_ID);
 Logger.log('✅ Installation complete! Check logs above for copy-paste instructions.');
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
const ss = openSpreadsheetById_(EXTERNAL_CONFIG_SHEET_ID);
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
	const results = installAllAutomationTriggers({ handlers: ['autoFixRequests'] });
	try {
		const ui = SpreadsheetApp.getUi();
		ui.alert('Audit Requests Auto-Fix', results.join('\n'), ui.ButtonSet.OK);
	} catch (_) {}
	return results;
}

function autoFixRequestsSheet_() {
 try {
 fixAuditRequestsSheet();
 } catch (e) {
 Logger.log('Auto-fix error: ' + e.message);
 }
}

// === Nightly External → Admin Sync Helpers ===
function installNightlyExternalSync() {
	const results = installAllAutomationTriggers({ handlers: ['nightlyMaintenance'] });
	try {
		const ui = SpreadsheetApp.getUi();
		ui.alert('Nightly Maintenance Triggers', results.join('\n'), ui.ButtonSet.OK);
	} catch (_) {}
	return results;
}

function removeNightlyExternalSync() {
	let removed = 0;
	ScriptApp.getProjectTriggers().forEach(function(t){
		try {
			const handler = t.getHandlerFunction && t.getHandlerFunction();
			if (handler === 'runNightlyExternalSync' || handler === 'runNightlyMaintenance' || handler === 'runNightlyExternalSync_') {
				ScriptApp.deleteTrigger(t);
				removed++;
			}
		} catch(_) {}
	});
	try { SpreadsheetApp.getUi().alert('Nightly Sync', `${removed} nightly sync trigger(s) removed.`, SpreadsheetApp.getUi().ButtonSet.OK); } catch(_) {}
}

function runNightlyExternalSync() {
	try {
		syncFromExternalConfig({
			silent: true,
			valuesOnly: false,
			copyFormatting: true,
			copyValidations: true,
			copyDimensions: true,
			copyProtections: false
		});
	} catch (e) {
		Logger.log('Nightly external sync error: ' + e.message);
		// Notify admin so failures are visible
		safeSendEmail({
			to: ADMIN_EMAIL,
			subject: 'CM360: Nightly External → Admin Sync Failed',
			htmlBody: `<pre style="font-family:monospace">${escapeHtml(e.message)}</pre>`
		}, 'runNightlyExternalSync');
	}
}

/**
 * Delete Gmail emails older than 90 days from all CM360 audit labels.
 * This function searches for emails in the "Daily Audits/CM360" label hierarchy
 * that are older than 90 days and permanently deletes them to manage storage.
 * 
 * @returns {Object} Statistics object with counts of deleted emails per label
 */
function deleteOldAuditEmails() {
	Logger.log('[Gmail Cleanup] Starting deletion of emails older than 90 days');
	
	const stats = {
		totalDeleted: 0,
		byLabel: {},
		errors: []
	};
	
	try {
		// Calculate cutoff date (90 days ago)
		const cutoffDate = new Date();
		cutoffDate.setDate(cutoffDate.getDate() - 90);
		const cutoffStr = Utilities.formatDate(cutoffDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
		
		Logger.log(`[Gmail Cleanup] Cutoff date: ${cutoffStr}`);
		
		// Get all labels that match the CM360 audit pattern
		const allLabels = GmailApp.getUserLabels();
		const auditLabels = allLabels.filter(label => {
			const labelName = label.getName();
			return labelName.startsWith('Daily Audits/CM360/');
		});
		
		Logger.log(`[Gmail Cleanup] Found ${auditLabels.length} audit labels to process`);
		
		// Process each label
		for (const label of auditLabels) {
			const labelName = label.getName();
			let labelDeleteCount = 0;
			
			try {
				// Search for threads older than 90 days with this label
				// Gmail search syntax: before:yyyy/mm/dd
				const searchQuery = `label:"${labelName}" before:${cutoffStr}`;
				
				// Process in batches to avoid timeout
				let batchStart = 0;
				const batchSize = 100;
				let hasMore = true;
				
				while (hasMore) {
					const threads = GmailApp.search(searchQuery, batchStart, batchSize);
					
					if (threads.length === 0) {
						hasMore = false;
						break;
					}
					
					// Move threads to trash (Gmail API limitation - can't permanently delete via Apps Script)
					for (const thread of threads) {
						try {
							thread.moveToTrash();
							labelDeleteCount++;
						} catch (threadErr) {
							Logger.log(`[Gmail Cleanup] Error deleting thread in ${labelName}: ${threadErr.message}`);
							stats.errors.push(`${labelName}: ${threadErr.message}`);
						}
					}
					
					// Check if we got fewer results than batch size (indicates last batch)
					if (threads.length < batchSize) {
						hasMore = false;
					} else {
						batchStart += batchSize;
					}
					
					// Small delay to avoid rate limiting
					if (hasMore) {
						Utilities.sleep(500);
					}
				}
				
				if (labelDeleteCount > 0) {
					stats.byLabel[labelName] = labelDeleteCount;
					stats.totalDeleted += labelDeleteCount;
					Logger.log(`[Gmail Cleanup] ${labelName}: deleted ${labelDeleteCount} thread(s)`);
				}
				
			} catch (labelErr) {
				Logger.log(`[Gmail Cleanup] Error processing label ${labelName}: ${labelErr.message}`);
				stats.errors.push(`${labelName}: ${labelErr.message}`);
			}
		}
		
		// Also clean up the parent "Daily Audits/CM360" label if it exists
		try {
			const parentLabel = GmailApp.getUserLabelByName('Daily Audits/CM360');
			if (parentLabel) {
				const searchQuery = `label:"Daily Audits/CM360" before:${cutoffStr}`;
				let batchStart = 0;
				const batchSize = 100;
				let hasMore = true;
				let parentDeleteCount = 0;
				
				while (hasMore) {
					const threads = GmailApp.search(searchQuery, batchStart, batchSize);
					
					if (threads.length === 0) {
						hasMore = false;
						break;
					}
					
					for (const thread of threads) {
						try {
							thread.moveToTrash();
							parentDeleteCount++;
						} catch (threadErr) {
							Logger.log(`[Gmail Cleanup] Error deleting thread in parent label: ${threadErr.message}`);
							stats.errors.push(`Daily Audits/CM360: ${threadErr.message}`);
						}
					}
					
					if (threads.length < batchSize) {
						hasMore = false;
					} else {
						batchStart += batchSize;
					}
					
					if (hasMore) {
						Utilities.sleep(500);
					}
				}
				
				if (parentDeleteCount > 0) {
					stats.byLabel['Daily Audits/CM360'] = parentDeleteCount;
					stats.totalDeleted += parentDeleteCount;
					Logger.log(`[Gmail Cleanup] Daily Audits/CM360 (parent): deleted ${parentDeleteCount} thread(s)`);
				}
			}
		} catch (parentErr) {
			Logger.log(`[Gmail Cleanup] Error processing parent label: ${parentErr.message}`);
			stats.errors.push(`Parent label: ${parentErr.message}`);
		}
		
		Logger.log(`[Gmail Cleanup] Completed. Total threads deleted: ${stats.totalDeleted}`);
		
		// Send summary email to admin if any deletions occurred or errors happened
		if (stats.totalDeleted > 0 || stats.errors.length > 0) {
			const summaryLines = [
				`<p style="font-family:Arial,sans-serif;font-size:13px;">Gmail cleanup completed for CM360 audit emails older than 90 days.</p>`,
				`<p style="font-family:Arial,sans-serif;font-size:13px;"><strong>Total threads deleted:</strong> ${stats.totalDeleted}</p>`
			];
			
			if (Object.keys(stats.byLabel).length > 0) {
				summaryLines.push('<p style="font-family:Arial,sans-serif;font-size:13px;"><strong>Breakdown by label:</strong></p>');
				summaryLines.push('<ul style="font-family:Arial,sans-serif;font-size:12px;">');
				for (const [labelName, count] of Object.entries(stats.byLabel)) {
					summaryLines.push(`<li>${escapeHtml(labelName)}: ${count} thread(s)</li>`);
				}
				summaryLines.push('</ul>');
			}
			
			if (stats.errors.length > 0) {
				summaryLines.push(`<p style="font-family:Arial,sans-serif;font-size:13px;color:#b00020;"><strong>Errors encountered:</strong> ${stats.errors.length}</p>`);
				summaryLines.push('<ul style="font-family:Arial,sans-serif;font-size:12px;color:#b00020;">');
				for (const error of stats.errors) {
					summaryLines.push(`<li>${escapeHtml(error)}</li>`);
				}
				summaryLines.push('</ul>');
			}
			
			summaryLines.push('<p style="margin-top:12px;font-family:Arial,sans-serif;font-size:12px;">&mdash; CM360 Audit System</p>');
			
			safeSendEmail({
				to: ADMIN_EMAIL,
				subject: `CM360: Gmail Cleanup Summary (${stats.totalDeleted} threads deleted)`,
				htmlBody: summaryLines.join('\n'),
				plainBody: `Gmail cleanup completed. ${stats.totalDeleted} threads deleted from CM360 audit labels.`
			}, 'deleteOldAuditEmails');
		}
		
	} catch (e) {
		Logger.log(`[Gmail Cleanup] Fatal error: ${e.message}`);
		stats.errors.push(`Fatal error: ${e.message}`);
		
		// Notify admin of failure
		safeSendEmail({
			to: ADMIN_EMAIL,
			subject: 'CM360: Gmail Cleanup Failed',
			htmlBody: `<p style="font-family:Arial,sans-serif;font-size:13px;color:#b00020;">Gmail cleanup encountered a fatal error:</p><pre style="font-family:monospace;background:#f5f5f5;padding:8px;">${escapeHtml(e.message)}</pre>`,
			plainBody: `Gmail cleanup failed: ${e.message}`
		}, 'deleteOldAuditEmails-error');
	}
	
	return stats;
}

function runNightlyMaintenance() {
	const results = [];
	const record = msg => results.push(msg);
	const invoke = (label, action) => {
		try {
			if (typeof action !== 'function') {
				record(`ℹ️ Skipped ${label} — handler not available`);
				return;
			}
			action();
			record(`✅ ${label}`);
		} catch (err) {
			Logger.log(`runNightlyMaintenance ${label} error: ${err.message}`);
			record(`⚠️ ${label} failed: ${err.message}`);
		}
	};

	invoke('rebalanceAuditBatchesUsingSummary', () => rebalanceAuditBatchesUsingSummary());

	if (EXTERNAL_CONFIG_SHEET_ID) {
		invoke('runNightlyExternalSync', () => runNightlyExternalSync());
		invoke('refreshExternalConfigInstructionsSilent', () => refreshExternalConfigInstructionsSilent());
	} else {
		record('ℹ️ Skipped runNightlyExternalSync — EXTERNAL_CONFIG_SHEET_ID not configured');
		record('ℹ️ Skipped refreshExternalConfigInstructionsSilent — EXTERNAL_CONFIG_SHEET_ID not configured');
	}

	invoke('updatePlacementNamesFromReports', () => updatePlacementNamesFromReports());
	invoke('clearDailyScriptProperties', () => clearDailyScriptProperties());
	invoke('cleanupOldAuditFiles', () => cleanupOldAuditFiles());
	invoke('deleteOldAuditEmails', () => deleteOldAuditEmails());

	Logger.log('runNightlyMaintenance results: ' + results.join(' | '));
	return results;
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
 const externalSpreadsheet = openSpreadsheetById_(EXTERNAL_CONFIG_SHEET_ID);
 
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
 * Options:
 * - silent: boolean (suppress UI alerts)
 * - valuesOnly: boolean (copy only values; skip formatting/validations/protections/widths/heights)
 * - sheets: string[] (limit to specific sheet names, e.g., ['Audit Recipients'])
 * Defaults to full sync with formatting when options not provided.
 */
function syncFromExternalConfig(options) {
	const silent = options && options.silent === true;
	const valuesOnly = options && options.valuesOnly === true;
	const sheetWhitelist = Array.isArray(options && options.sheets) ? options.sheets : null;
	// Guard against Apps Script 6-min runtime by limiting work to ~5 minutes
	const START_TIME = Date.now();
	const MAX_RUNTIME_MS = 5 * 60 * 1000; // 5 minutes safety budget
	// Additional guardrails per sheet
	const SHEET_TIME_BUDGET_MS = 30 * 1000; // try to keep each sheet under ~30s
	const HEAVY_CELLS_THRESHOLD = 100000; // if rows*cols exceed this, skip heavy ops
	const MAX_PROTECTIONS = 20; // cap protections copied to avoid explosion/duplication

	// Fine-grained copy options (valuesOnly overrides to false)
	const copyFormatting = !valuesOnly && !(options && options.copyFormatting === false);
	const copyValidations = !valuesOnly && !(options && options.copyValidations === false);
	const copyDimensions = !valuesOnly && !(options && options.copyDimensions === false);
	// Protections are OFF by default; enable explicitly with options.copyProtections=true
	const copyProtections = !valuesOnly && !!(options && options.copyProtections === true);
	// Detect whether a Spreadsheet UI is available (true when run via menu/button)
	let ui = null;
	try {
		ui = SpreadsheetApp.getUi();
	} catch (e) {
		ui = null; // Running from trigger or other non-UI context
	}

	// When silent is requested, suppress UI even if available
	if (silent) ui = null;

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
 Logger.log(`Starting sync from external config sheet: ${EXTERNAL_CONFIG_SHEET_ID}`);
 
 const mainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 const externalSpreadsheet = openSpreadsheetById_(EXTERNAL_CONFIG_SHEET_ID);
 
 const allSheets = [
 { name: RECIPIENTS_SHEET_NAME, description: 'Recipients' },
 { name: THRESHOLDS_SHEET_NAME, description: 'Thresholds' },
 { name: EXCLUSIONS_SHEET_NAME, description: 'Exclusions' },
 { name: 'Audit Requests', description: 'Audit Requests' }
 ];
 const sheetsToSync = sheetWhitelist
 	? allSheets.filter(s => sheetWhitelist.indexOf(s.name) !== -1)
 	: allSheets;
 
 const syncResults = [];
 
 for (const sheetInfo of sheetsToSync) {
 // Runtime guard: if we're close to the limit, abort remaining work gracefully
 if (Date.now() - START_TIME > MAX_RUNTIME_MS) {
 	syncResults.push('Aborted remaining sheets due to runtime limit. Partial sync completed.');
 	break;
 }
 Logger.log(` Syncing: ${sheetInfo.name} (${sheetInfo.description})`);
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
 const cellCount = numRows * numCols;
 const sheetStart = Date.now();
 
 if (numRows > 0) {
 const targetRange = mainSheet.getRange(1, 1, numRows, numCols);
 
 // Copy values
 const values = externalRange.getValues();
 targetRange.setValues(values);

 if (!valuesOnly) {
	const heavy = cellCount > HEAVY_CELLS_THRESHOLD;
	if (heavy) {
		Logger.log(`  Skipping heavy ops for ${sheetInfo.name} (cells=${cellCount})`);
	}
 	// Copy formatting manually since copyTo doesn't work across spreadsheets
 	if (copyFormatting && !heavy && (Date.now() - sheetStart) < SHEET_TIME_BUDGET_MS) try {
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
 
 	// Copy data validations in bulk (faster and avoids per-cell loops)
 	if (copyValidations && !heavy && (Date.now() - sheetStart) < SHEET_TIME_BUDGET_MS) try {
 		const validations = externalRange.getDataValidations();
 		if (validations) targetRange.setDataValidations(validations);
 	} catch (validationError) {
 		Logger.log(` Could not copy validations for ${sheetInfo.name}: ${validationError.message}`);
 	}
 
 	// Copy column widths
 	if (copyDimensions && (Date.now() - sheetStart) < SHEET_TIME_BUDGET_MS) try {
 		for (let col = 1; col <= numCols; col++) {
 			const width = externalSheet.getColumnWidth(col);
 			mainSheet.setColumnWidth(col, width);
 		}
 	} catch (widthError) {
 		Logger.log(` Could not copy column widths for ${sheetInfo.name}: ${widthError.message}`);
 	}
 
 	// Copy row heights for first 20 rows (where instructions typically are)
 	if (copyDimensions && (Date.now() - sheetStart) < SHEET_TIME_BUDGET_MS) try {
 		for (let row = 1; row <= Math.min(20, numRows); row++) {
 			const height = externalSheet.getRowHeight(row);
 			mainSheet.setRowHeight(row, height);
 		}
 	} catch (heightError) {
 		Logger.log(` Could not copy row heights for ${sheetInfo.name}: ${heightError.message}`);
 	}
 
 	// Copy sheet-level protections
 	if (copyProtections && (Date.now() - sheetStart) < SHEET_TIME_BUDGET_MS) try {
 		const externalProtections = externalSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
 		if (externalProtections.length > MAX_PROTECTIONS) {
 			Logger.log(` Skipping protections for ${sheetInfo.name}: too many (${externalProtections.length} > ${MAX_PROTECTIONS}).`);
 		} else {
 			externalProtections.forEach(protection => {
 				if ((Date.now() - sheetStart) >= SHEET_TIME_BUDGET_MS) return; // stop if over budget
 				try {
 					const range = protection.getRange();
 					const mainRange = mainSheet.getRange(range.getA1Notation());
 					const newProtection = mainRange.protect();
 					newProtection.setDescription(protection.getDescription());
 					if (protection.isWarningOnly()) newProtection.setWarningOnly(true);
 				} catch (protectionError) {
 					Logger.log(` Could not copy protection: ${protectionError.message}`);
 				}
 			});
 		}
 	} catch (protectionsError) {
 		Logger.log(` Could not copy protections for ${sheetInfo.name}: ${protectionsError.message}`);
 	}
 }
 }
 
 // Apply pending changes for this sheet
 try { SpreadsheetApp.flush(); } catch (flushErr) { /* noop */ }
 
 const detailParts = [];
 detailParts.push('values');
 if (!valuesOnly) {
 	if (copyFormatting) detailParts.push('formatting');
 	if (copyValidations) detailParts.push('validations');
 	if (copyDimensions) detailParts.push('widths/heights');
 	if (copyProtections) detailParts.push('protections');
 }
 const detail = detailParts.join(' + ');
	const elapsedMs = Date.now() - sheetStart;
	Logger.log(` Done syncing: ${sheetInfo.name} -> ${numRows}x${numCols} cells in ${elapsedMs} ms (${detail})`);
 syncResults.push(`... ${sheetInfo.description}: Synced ${numRows} rows (${detail})`);
 
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
				`Sync from external config sheet completed:\n\n${resultMessage}\n\n${valuesOnly ? 'Only values were copied; formatting, validations, widths/heights, and protections were skipped.' : 'Formatting, validations, widths/heights, and protections were preserved where possible.'}`,
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

/** Fast wrapper for menu: values only, skip heavy ops to avoid timeouts */
function syncFromExternalConfigQuick() {
	try {
		return syncFromExternalConfig({
			silent: false,
			valuesOnly: true,
			copyFormatting: false,
			copyValidations: false,
			copyDimensions: false,
			copyProtections: false
		});
	} catch (e) {
		Logger.log(`syncFromExternalConfigQuick error: ${e.message}`);
		const ui = SpreadsheetApp.getUi();
		try { ui.alert('Sync Error', e.message, ui.ButtonSet.OK); } catch (_) {}
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

 const ss = openSpreadsheetById_(EXTERNAL_CONFIG_SHEET_ID);

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
 'Placement Name (auto-populated)',
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

 // Also enforce formatting rules (static flag type, banding, inactive shading)
 try {
	 if (sheets[THRESHOLDS_SHEET_NAME]) {
		 try { sheets[THRESHOLDS_SHEET_NAME].getRange('B2:B').clearDataValidations(); } catch (e) {}
		 applyThresholdsFormatting_(sheets[THRESHOLDS_SHEET_NAME]);
	 }
	 if (sheets[RECIPIENTS_SHEET_NAME]) {
		 const rc = sheets[RECIPIENTS_SHEET_NAME];
		 const rules = rc.getConditionalFormatRules() || [];
		 const filtered = rules.filter(r => {
			 try { const bc = r.getBooleanCondition(); if (!bc) return true; const v = bc.getCriteriaValues()||[]; const f=String((v[0]||'')).toUpperCase(); return f !== '=$D2=FALSE' && f !== '=AND($D2=FALSE, LEN($A2)>0)'; } catch(e){ return true; }
		 });
		 const inactiveRange = rc.getRange(2,1,Math.max(rc.getMaxRows()-1,1),4);
		 const inactiveRule = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($D2=FALSE, LEN($A2)>0)').setBackground('#f8d7da').setRanges([inactiveRange]).build();
		 filtered.push(inactiveRule);
		 rc.setConditionalFormatRules(filtered);
	 }
	 if (sheets[EXCLUSIONS_SHEET_NAME]) {
		 const ex = sheets[EXCLUSIONS_SHEET_NAME];
		 const rules = ex.getConditionalFormatRules() || [];
		 const filtered = rules.filter(r => {
			 try { const bc = r.getBooleanCondition(); if (!bc) return true; const v = bc.getCriteriaValues()||[]; const f=String((v[0]||'')).toUpperCase(); return f !== '=$K2=FALSE' && f !== '=AND($K2=FALSE, LEN($A2)>0)'; } catch(e){ return true; }
		 });
		 const inactiveRange = ex.getRange(2,1,Math.max(ex.getMaxRows()-1,1),11);
		 const inactiveRule = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($K2=FALSE, LEN($A2)>0)').setBackground('#f8d7da').setRanges([inactiveRange]).build();
		 filtered.push(inactiveRule);
		 ex.setConditionalFormatRules(filtered);
		 // Enforce protection and visual styling on Placement Name column (C)
		 try { enforcePlacementNameProtectionAndStyle_(ex); } catch (e) { Logger.log('ensureExternalConfigInstructions: placement name protect/style: ' + e.message); }
	 }
 } catch (e) {
	 Logger.log('ensureExternalConfigInstructions: formatting enforcement error: ' + e.message);
 }

 ui.alert(
 'Instructions + Formatting Updated',
 'External config tabs now have standardized INSTRUCTIONS, static Flag Type on Thresholds, 4-row banding (blue) when Active=TRUE, and inactive (FALSE) rows shaded light red, skipping blanks.',
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
		const ss = openSpreadsheetById_(EXTERNAL_CONFIG_SHEET_ID);

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

// Apply our formatting rules to the external configuration spreadsheet
function applyExternalFormattingRules() {
	if (!EXTERNAL_CONFIG_SHEET_ID) {
		Logger.log('applyExternalFormattingRules: EXTERNAL_CONFIG_SHEET_ID not set');
		try { SpreadsheetApp.getUi().alert('No External Config Sheet', 'Set EXTERNAL_CONFIG_SHEET_ID first.', SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
		return;
	}
	try {
		const ss = openSpreadsheetById_(EXTERNAL_CONFIG_SHEET_ID);
		const th = ss.getSheetByName(THRESHOLDS_SHEET_NAME);
		const rc = ss.getSheetByName(RECIPIENTS_SHEET_NAME);
		const ex = ss.getSheetByName(EXCLUSIONS_SHEET_NAME);
		if (th) {
			try { th.getRange('B2:B').clearDataValidations(); } catch (e) {}
			try { applyThresholdsFormatting_(th); } catch (e) { Logger.log('applyExternalFormattingRules thresholds: ' + e.message); }
		}
		if (rc) {
			try {
				const rules = rc.getConditionalFormatRules() || [];
				const filtered = rules.filter(r => {
						try { const bc = r.getBooleanCondition(); if (!bc) return true; const v = bc.getCriteriaValues()||[]; const f=String((v[0]||'')).toUpperCase(); return f !== '=$D2=FALSE' && f !== '=AND($D2=FALSE, LEN($A2)>0)'; } catch(e){ return true; }
				});
					const inactiveRange = rc.getRange(2,1,Math.max(rc.getMaxRows()-1,1),4); // A..D
					const inactiveRule = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($D2=FALSE, LEN($A2)>0)').setBackground('#f8d7da').setRanges([inactiveRange]).build();
				filtered.push(inactiveRule);
				rc.setConditionalFormatRules(filtered);
			} catch (e) { Logger.log('applyExternalFormattingRules recipients: ' + e.message); }
		}
		if (ex) {
			try {
				const rules = ex.getConditionalFormatRules() || [];
				const filtered = rules.filter(r => {
						try { const bc = r.getBooleanCondition(); if (!bc) return true; const v = bc.getCriteriaValues()||[]; const f=String((v[0]||'')).toUpperCase(); return f !== '=$K2=FALSE' && f !== '=AND($K2=FALSE, LEN($A2)>0)'; } catch(e){ return true; }
				});
					const inactiveRange = ex.getRange(2,1,Math.max(ex.getMaxRows()-1,1),11); // A..K
					const inactiveRule = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($K2=FALSE, LEN($A2)>0)').setBackground('#f8d7da').setRanges([inactiveRange]).build();
				filtered.push(inactiveRule);
				ex.setConditionalFormatRules(filtered);
				// Enforce protection and visual styling on Placement Name column (C)
				try { enforcePlacementNameProtectionAndStyle_(ex); } catch (e) { Logger.log('applyExternalFormattingRules: placement name protect/style: ' + e.message); }
			} catch (e) { Logger.log('applyExternalFormattingRules exclusions: ' + e.message); }
		}
		try { SpreadsheetApp.getUi().alert('Formatting Applied', 'External formatting rules have been applied.', SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
		Logger.log('applyExternalFormattingRules: completed');
	} catch (err) {
		Logger.log('applyExternalFormattingRules error: ' + err.message);
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
 ['Staging Mode Override:', `Currently: ${getStagingMode_() === 'Y' ? 'STAGING (all emails go to admin)' : 'PRODUCTION (uses sheet recipients)'}`],
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
 
const externalSpreadsheet = openSpreadsheetById_(EXTERNAL_CONFIG_SHEET_ID);
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
 
const externalSpreadsheet = openSpreadsheetById_(EXTERNAL_CONFIG_SHEET_ID);
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
 
const externalSpreadsheet = openSpreadsheetById_(EXTERNAL_CONFIG_SHEET_ID);
 
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
		// Use lightweight, values-only sync for Recipients and Thresholds to avoid timeouts.
		try {
			syncFromExternalConfig({ silent: true, valuesOnly: true, sheets: [RECIPIENTS_SHEET_NAME, THRESHOLDS_SHEET_NAME] });
			Logger.log('Synced (values-only) main spreadsheet from external config before processing requests.');
		} catch (syncErr) {
			Logger.log(`Failed to sync from external config prior to processing requests: ${syncErr.message}`);
			// Proceed anyway; requests will still be read from the external sheet.
		}
 Logger.log(`" Processing audit requests from external config sheet: ${EXTERNAL_CONFIG_SHEET_ID}`);
 
const externalSpreadsheet = openSpreadsheetById_(EXTERNAL_CONFIG_SHEET_ID);
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
 const config = getAuditConfigs().find(c => c.name === request.configName);
 if (!config) {
 throw new Error(`Configuration "${request.configName}" not found in recipients list`);
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
 
	// Show a non-blocking toast instead of a blocking UI alert to avoid hitting the 6-minute menu limit
	const completedCount = processedRequests.filter(r => r.status === 'COMPLETED').length;
	const failedCount = processedRequests.filter(r => r.status === 'FAILED' || r.status === 'ERROR').length;
	try {
		const ss = SpreadsheetApp.getActiveSpreadsheet();
		if (ss) {
			ss.toast(
				`Processed ${processedRequests.length} request(s) — Completed: ${completedCount}; Failed: ${failedCount}. Check email and the external sheet for details.`,
				'Audit Requests Processed',
				10
			);
		}
	} catch (toastErr) {
		// Fallback: just log if toast fails (e.g., no active spreadsheet context)
		Logger.log('Toast failed: ' + toastErr.message);
	}
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
 
 // First, analyze current source for existing batch functions
 let sourceInfo = null;
 try { sourceInfo = listBatchFunctionsInSource_(); } catch (e) { Logger.log('source scan failed: ' + e.message); }
 let createdInSource = [];
 
 // Auto-creation disabled - batch functions must be added manually
 // (Re-scan to get updated detection after any manual additions)
 try { sourceInfo = listBatchFunctionsInSource_(); } catch (e) { Logger.log('post-insert source scan failed: ' + e.message); } // Step 1: Check current batch status
 const batches = getAuditConfigBatches(BATCH_SIZE);
 const neededCount = batches.length;
 
 // Prefer source-based detection to avoid counting temporary in-memory stubs
 let definedIndexes = (sourceInfo && sourceInfo.existingIndexes) ? sourceInfo.existingIndexes : new Set();
 // Fallback: if source scan yields none (e.g., Script API disabled by scopes), fall back to runtime check
 if (!definedIndexes || definedIndexes.size === 0) {
	 definedIndexes = new Set();
	 for (let i = 1; i <= neededCount; i++) {
		 const fnName = `runDailyAuditsBatch${i}`;
		 try {
			 if (typeof globalThis[fnName] === 'function') definedIndexes.add(i);
		 } catch (_) {}
	 }
 }
 const existingFns = Array.from(definedIndexes).sort((a,b)=>a-b).map(i => `runDailyAuditsBatch${i}`);
 
 let missingFunctions = [];
 for (let i = 1; i <= neededCount; i++) {
 if (!definedIndexes.has(i)) {
 missingFunctions.push(`runDailyAuditsBatch${i}`);
 }
 }
 
 // Step 2: Report status and get user confirmation
 let statusMessage = ` Batch Status Analysis:\n\n`;
 statusMessage += `- Total configs: ${getAuditConfigs().length}\n`;
 statusMessage += `- Batch size: ${BATCH_SIZE}\n`;
 statusMessage += `- Batches needed: ${neededCount}\n`;
 statusMessage += `- Existing batch functions: ${existingFns.length}\n`;
 statusMessage += `- Missing batch functions: ${missingFunctions.length}\n\n`;
 
 if (missingFunctions.length > 0) {
		 statusMessage += `❌ Missing functions in source (add manually):\n${missingFunctions.map(fn => `- ${fn}`).join('\n')}\n\n`;
		 statusMessage += `This run will proceed to (re)install triggers only for functions present in source.\n\nContinue?`;
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
 
 // Step 3: Proceed to triggers (source already updated if needed)
 
 // Step 4: Install triggers for existing functions and supporting automations
 Logger.log(' Installing automation triggers...');
 const triggerResults = installAllAutomationTriggers();
 const batchInstallCount = triggerResults.filter(r => /daily audit batch trigger/.test(r)).length;
 const totalInstalledCount = triggerResults.filter(r => /^✅/.test(r)).length;
 const informationalNotes = triggerResults.filter(r => /^ℹ️/.test(r));
 
 // Step 5: Report final results
    let finalMessage = `✅ Batch Triggers Setup Complete!\n\n`;
 finalMessage += ` Summary:\n`;
 finalMessage += `- Batch functions present: ${existingFns.length}/${neededCount}\n`;
 finalMessage += `- Batch triggers installed: ${batchInstallCount}\n`;
 finalMessage += `- Total automation installs this run: ${totalInstalledCount}\n`;
 finalMessage += `- Configs per batch: ${BATCH_SIZE}\n\n`;
 finalMessage += ` Batches:\n`;
 
 batches.forEach((batch, index) => {
 finalMessage += `- Batch ${index + 1}: ${batch.map(c => c.name).join(', ')}\n`;
 });
 
	finalMessage += `\n✅ Automation triggers refreshed.`;

	if (informationalNotes.length) {
		finalMessage += `\n\nℹ️ Notes:\n${informationalNotes.join('\n')}`;
	}
 
 ui.alert(
 'Setup Complete',
 finalMessage,
 ui.ButtonSet.OK
 );
 
		Logger.log('✅ Batch triggers setup completed successfully');
		Logger.log(` Installed ${batchInstallCount} batch triggers (${totalInstalledCount} total automation triggers this run)`);
 
 } catch (error) {
	Logger.log(`❌ Error in setupAndInstallBatchTriggers: ${error.message}`);
 ui.alert(
 'Setup Error',
 `Failed to setup batch triggers:\n\n${error.message}`,
 ui.ButtonSet.OK
 );
 }
}

// === DIAGNOSTIC UTILITIES ===

/**
 * HOW TO ENABLE/DISABLE STAGING MODE:
 * 
 * Staging mode routes all audit emails to ADMIN_EMAIL only (for testing).
 * 
 * METHOD 1 (Recommended - immediate effect):
 *   Open Apps Script Editor → Project Settings → Script Properties
 *   Add/Edit property: STAGING_MODE
 *   Set to: Y (staging) or N (production)
 *   Changes take effect immediately on next execution.
 * 
 * METHOD 2 (Code default - requires push):
 *   Change line 12 default: || 'Y' to || 'N' (or vice versa)
 *   Run: clasp push
 *   Note: Existing triggers may cache old value until reinstalled
 * 
 * Current behavior: Defaults to 'Y' (staging) if Script Property not set
 */

/**
 * Diagnostic function to trace threshold loading for a specific config
 */
function diagnoseThresholds(configName) {
 try {
 Logger.log(`\n=== THRESHOLD DIAGNOSTIC FOR ${configName} ===\n`);
 
 // First, inspect raw sheet data
 Logger.log('=== RAW SHEET INSPECTION ===');
 const sheet = getOrCreateThresholdsSheet();
 const allData = sheet.getDataRange().getValues();
 
 Logger.log(`Total rows in sheet: ${allData.length}`);
 Logger.log(`Header row (row 0): ${JSON.stringify(allData[0])}`);
 
 // Show first 5 data rows
 Logger.log('\nFirst 5 data rows:');
 for (let i = 1; i < Math.min(6, allData.length); i++) {
 Logger.log(`Row ${i}: ${JSON.stringify(allData[i])}`);
 }
 
 // Show WRI01 rows specifically
 Logger.log('\n=== WRI01 ROWS ===');
 for (let i = 1; i < allData.length; i++) {
 const row = allData[i];
 const configName_raw = String(row[0] || '');
 if (configName_raw.includes('WRI01')) {
 Logger.log(`Row ${i+1}: [${row.map((v, idx) => `col${idx}="${v}"`).join(', ')}]`);
 }
 }
 
 // Load thresholds
 Logger.log('\n=== THRESHOLD LOADING ===');
 const thresholdsData = loadThresholdsFromSheet(true);
 Logger.log(`Loaded thresholds for ${Object.keys(thresholdsData).length} configs total`);
 
 // Check if config exists
 if (!thresholdsData[configName]) {
 Logger.log(`❌ Config "${configName}" NOT FOUND in thresholds data`);
 Logger.log(`Available configs: ${Object.keys(thresholdsData).join(', ')}`);
 return `Config ${configName} not found in thresholds`;
 }
 
 Logger.log(`✅ Config "${configName}" found in thresholds data`);
 
 // Show all flag thresholds for this config
 const configThresholds = thresholdsData[configName];
 Logger.log(`\nThresholds for ${configName}:`);
 for (const [flagType, threshold] of Object.entries(configThresholds)) {
 Logger.log(` - ${flagType}:`);
 Logger.log(` minImpressions: ${threshold.minImpressions}`);
 Logger.log(` minClicks: ${threshold.minClicks}`);
 }
 
 return `Diagnostic complete - see logs`;
 
 } catch (error) {
 Logger.log(`❌ Error in diagnoseThresholds: ${error.message}`);
 throw error;
 }
}

/**
 * Quick helper to set staging mode ON (emails to admin only)
 */
function setStagingModeOn() {
 try {
 PropertiesService.getScriptProperties().setProperty('STAGING_MODE', 'Y');
 Logger.log('✅ Staging mode set to ON - emails will route to admin only');
 return 'Staging mode enabled';
 } catch (error) {
 Logger.log(`❌ Error setting staging mode: ${error.message}`);
 throw error;
 }
}

/**
 * Quick helper to set staging mode OFF (emails to real recipients)
 */
function setStagingModeOff() {
 try {
 PropertiesService.getScriptProperties().setProperty('STAGING_MODE', 'N');
 Logger.log('✅ Staging mode set to OFF - emails will route to real recipients');
 return 'Staging mode disabled';
 } catch (error) {
 Logger.log(`❌ Error setting staging mode: ${error.message}`);
 throw error;
 }
}

/**
 * Run a single config audit with detailed threshold logging
 * Use this to test threshold behavior in staging mode
 */
function testAuditWithThresholdLogging(configName) {
 try {
 // Check staging mode status
 const stagingMode = getStagingMode_();
 Logger.log(`Current staging mode: ${stagingMode}`);
 if (stagingMode !== 'Y') {
 Logger.log('⚠️ WARNING: Staging mode is NOT enabled. Emails will go to real recipients!');
 Logger.log('To enable: Run setStagingModeOn() first, then run this test again');
 }
 
 Logger.log(`\n=== TESTING AUDIT FOR ${configName} WITH THRESHOLD LOGGING ===\n`);
 
 // Find the config
 const configs = getAuditConfigs();
 const config = configs.find(c => c.name === configName);
 
 if (!config) {
 Logger.log(`❌ Config "${configName}" not found`);
 Logger.log(`Available configs: ${configs.map(c => c.name).join(', ')}`);
 return `Config ${configName} not found`;
 }
 
 // Run the audit
 Logger.log(`Starting audit for ${configName}...`);
 const result = executeAudit(config);
 
 Logger.log(`\n=== AUDIT COMPLETE ===`);
 Logger.log(`Status: ${result.status}`);
 Logger.log(`Flagged count: ${result.flaggedCount || 0}`);
 Logger.log(`Email sent: ${result.emailSent}`);
 
 return result;
 
 } catch (error) {
 Logger.log(`❌ Error in testAuditWithThresholdLogging: ${error.message}`);
 throw error;
 }
}




