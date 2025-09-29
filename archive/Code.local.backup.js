// Moved to archive to prevent clasp from pushing this backup file.
// Original content preserved below.

function makeAuditConfig_(name, label) {
	return {
		name: name,
		label: label || name,
		mergedFolderPath: [...TRASH_ROOT_PATH, 'Merged Reports', name],
		tempDailyFolderPath: [...TRASH_ROOT_PATH, 'Temp Daily Reports', name]
	};
}

let auditConfigsCache_ = null;

function clearAuditConfigsCache_() {
	auditConfigsCache_ = null;
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
	const sheet = getOrCreateRecipientsSheet();
	const data = sheet.getDataRange().getValues();
	const configs = [];
	const seen = new Set();
	for (let i = 1; i < data.length; i++) {
		const row = data[i];
		if (!shouldIncludeConfigRow_(row[0], row[3])) continue;
		const name = String(row[0]).trim();
		if (seen.has(name)) continue;
		seen.add(name);
		const label = `Daily Audits/CM360/${name}`;
		configs.push(makeAuditConfig_(name, label));
	}
	configs.sort((a, b) => a.name.localeCompare(b.name));
	auditConfigsCache_ = { list: configs, timestamp: Date.now() };
	return configs;
}

function getAuditConfigByName(configName) {
	const name = String(configName || '').trim();
	if (!name) return null;
	return getAuditConfigs().find(cfg => cfg.name === name) || null;
}

// NOTE: This is a snapshot; the full original file had more content. If needed, retrieve from version control.
