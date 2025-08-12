/*// === 📁 CONFIGURATION & CONSTANTS ===
const ADMIN_EMAIL = 'evschneider@horizonmedia.com';
const STAGING_MODE = 'Y'; // Set to 'Y' for staging mode, 'N' for production
const EXCLUSIONS_SHEET_NAME = 'CM360 Audit Exclusions'; // Name of the sheet containing exclusions

// === 🔧 AUDIT CONSTANTS ===
const AUDIT_CONSTANTS = {
  BATCH_SIZE: 3,
  TRASH_ROOT_PATH: ['Project Log Files', 'CM360 Daily Audits', 'To Trash After 60 Days'],
  DELETION_LOG_PATH: ['Project Log Files', 'CM360 Daily Audits', 'To Trash After 60 Days', 'Deletion Log'],
  MASTER_LOG_NAME: 'CM360 Deleted Files Log',
  MIN_VOLUME_THRESHOLD: 0,
  FLAG_TYPES: {
    CLICKS_GT_IMPRESSIONS: 'clicks_greater_than_impressions',
    OUT_OF_FLIGHT: 'out_of_flight_dates',
    PIXEL_MISMATCH: 'pixel_size_mismatch',
    DEFAULT_AD: 'default_ad_serving',
    ALL_FLAGS: 'all_flags'
  },
  EMAIL_QUOTA_CACHE_DURATION: 21600, // 6 hours
  PROCESSING_BATCH_SIZE: 100
};

const BATCH_SIZE = AUDIT_CONSTANTS.BATCH_SIZE;
const TRASH_ROOT_PATH = AUDIT_CONSTANTS.TRASH_ROOT_PATH;
const DELETION_LOG_PATH = AUDIT_CONSTANTS.DELETION_LOG_PATH;
const MASTER_LOG_NAME = AUDIT_CONSTANTS.MASTER_LOG_NAME;

// === �️ UTILITY CLASSES ===
class AuditUtils {
  static formatDate(date, format = 'yyyy-MM-dd') {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), format);
  }

  static escapeHtml(text) {
    if (text === null || text === undefined) return '';
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  static normalize(text) {
    return String(text || '').toLowerCase().trim();
  }

  static createLogger(configName) {
    return {
      log: (message) => Logger.log(`[${configName}] ${message}`),
      error: (message) => Logger.log(`❌ [${configName}] ${message}`),
      warn: (message) => Logger.log(`⚠️ [${configName}] ${message}`),
      info: (message) => Logger.log(`ℹ️ [${configName}] ${message}`),
      success: (message) => Logger.log(`✅ [${configName}] ${message}`)
    };
  }

  static truncateText(text, maxLen = 80) {
    const safe = String(text || '').trim();
    return safe.length > maxLen ? safe.slice(0, maxLen - 1) + '…' : safe;
  }

  static plural(count, singular, plural) {
    return count === 1 ? singular : plural;
  }
}

class CacheManager {
  static getExclusions() {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('CM360_EXCLUSIONS_DATA');
    if (cached) {
      const data = JSON.parse(cached);
      return new Set(data);
    }
    return new Set(); // Return empty Set instead of null
  }

  static setExclusions(exclusionsData) {
    const cache = CacheService.getScriptCache();
    const exclusionsArray = Array.from(exclusionsData); // Convert Set to Array for JSON
    cache.put('CM360_EXCLUSIONS_DATA', JSON.stringify(exclusionsArray), 3600); // 1 hour
  }

  static getEmailQuotaRemaining() {
    const cache = CacheService.getScriptCache();
    const val = cache.get('CM360_EMAIL_QUOTA_LEFT');
    return val !== null ? Number(val) : null;
  }

  static setEmailQuotaRemaining(remaining) {
    const cache = CacheService.getScriptCache();
    const existing = this.getEmailQuotaRemaining();
    
    if (existing === null || Number(remaining) < Number(existing)) {
      cache.put('CM360_EMAIL_QUOTA_LEFT', String(remaining), AUDIT_CONSTANTS.EMAIL_QUOTA_CACHE_DURATION);
      Logger.log(`Updated cached quota remaining to: ${remaining}`);
    }
  }

  static decrementEmailQuota() {
    const cache = CacheService.getScriptCache();
    const current = this.getEmailQuotaRemaining();
    
    if (current !== null && current > 0) {
      const newQuota = current - 1;
      cache.put('CM360_EMAIL_QUOTA_LEFT', String(newQuota), AUDIT_CONSTANTS.EMAIL_QUOTA_CACHE_DURATION);
      Logger.log(`Email quota decremented to: ${newQuota}`);
      return newQuota;
    }
    
    Logger.log('Email quota already at 0 or not initialized');
    return 0;
  }
}

// === �📦 UTILITY HELPERS ===
function folderPath(type, configName) {
  return [...TRASH_ROOT_PATH, type, configName];
}

function resolveRecipients(recipients) {
  return STAGING_MODE === 'Y' ? ADMIN_EMAIL : recipients;
}

function resolveCc(ccList) {
  return STAGING_MODE === 'Y' ? '' : ccList.filter(Boolean).join(', ');
}

// === 🚨 ERROR HANDLING ===
class AuditErrorHandler {
  static handleFileNotFound(config) {
    const logger = AuditUtils.createLogger(config.name);
    logger.warn('No files found today. Sending notification...');
    
    const subject = `⚠️ CM360 Audit Skipped: No Files Found (${config.name} - ${AuditUtils.formatDate(new Date())})`;
    const htmlBody = `
      <p style="font-family:Arial, sans-serif; font-size:13px;">
        The CM360 audit for bundle "<strong>${AuditUtils.escapeHtml(config.name)}</strong>" was skipped because no Excel or ZIP files were found for today.
      </p>
      <p style="font-family:Arial, sans-serif; font-size:13px;">
        Please verify the report was delivered and labeled correctly.
      </p>
      <p style="margin-top:12px; font-family:Arial, sans-serif; font-size:12px;">—Platform Solutions Team</p>
    `;
    
    safeSendEmail({ 
      to: config.recipients, 
      cc: config.cc || '', 
      subject, 
      htmlBody, 
      attachments: [] 
    }, config.name);
    
    return { 
      status: 'Skipped: No files found', 
      flaggedCount: null, 
      emailSent: true, 
      emailTime: AuditUtils.formatDate(new Date(), 'yyyy-MM-dd HH:mm:ss') 
    };
  }

  static handleProcessingError(error, context) {
    const logger = AuditUtils.createLogger(context.configName || 'Unknown');
    logger.error(`Processing error: ${error.message}`);
    
    return { 
      status: `Error during audit: ${error.message}`, 
      flaggedCount: null, 
      emailSent: false, 
      emailTime: AuditUtils.formatDate(new Date(), 'yyyy-MM-dd HH:mm:ss') 
    };
  }

  static handleHeaderNotFound(configName) {
    const logger = AuditUtils.createLogger(configName);
    logger.error('Header row not found in merged sheet');
    
    return { 
      status: 'Failed: Header not found', 
      flaggedCount: null, 
      emailSent: false, 
      emailTime: AuditUtils.formatDate(new Date(), 'yyyy-MM-dd HH:mm:ss') 
    };
  }

  static sendErrorNotification(error, config) {
    const subject = `CM360 Audit Error: ${config.name} - ${AuditUtils.formatDate(new Date())}`;
    const htmlBody = `
      <p style="font-family:Arial, sans-serif; font-size:13px;">
        An error occurred during the CM360 audit for bundle "<strong>${AuditUtils.escapeHtml(config.name)}</strong>".
      </p>
      <p style="font-family:Arial, sans-serif; font-size:13px;">
        <strong>Error:</strong> ${AuditUtils.escapeHtml(error.message)}
      </p>
      <p style="margin-top:12px; font-family:Arial, sans-serif; font-size:12px;">—Platform Solutions Team</p>
    `;
    
    safeSendEmail({ 
      to: config.recipients, 
      cc: config.cc || '', 
      subject, 
      htmlBody, 
      attachments: [] 
    }, config.name);
  }

  static handleNoFiles(config, logger) {
    logger.warn('No files found today. Sending notification...');
    
    const timestamp = AuditUtils.formatDate(new Date(), 'yyyy-MM-dd');
    const subject = `⚠️ CM360 Audit Skipped: No Files Found (${config.name} - ${timestamp})`;
    const htmlBody = `
      <p style="font-family:Arial, sans-serif; font-size:13px;">
        The CM360 audit for bundle "<strong>${AuditUtils.escapeHtml(config.name)}</strong>" was skipped because no Excel or ZIP files were found for today.
      </p>
      <p style="font-family:Arial, sans-serif; font-size:13px;">
        Please verify the report was delivered and labeled correctly.
      </p>
      <p style="margin-top:12px; font-family:Arial, sans-serif; font-size:12px;">—Platform Solutions Team</p>
    `;
    
    safeSendEmail({ 
      to: config.recipients, 
      cc: config.cc || '', 
      subject, 
      htmlBody, 
      attachments: [] 
    }, config.name);
    
    return { 
      status: 'Skipped: No files found', 
      flaggedCount: null, 
      emailSent: true, 
      emailTime: AuditUtils.formatDate(new Date(), 'yyyy-MM-dd HH:mm:ss') 
    };
  }

  static handleLabelNotFound(context, config) {
    const logger = AuditUtils.createLogger(config.name);
    logger.warn(context);
    
    const subject = `⚠️ CM360 Audit Warning: Gmail Label Missing (${config.name})`;
    const htmlBody = `<p style="font-family:Arial; font-size:13px;">The label <b>${AuditUtils.escapeHtml(config.label)}</b> could not be found. This may mean the audit for <b>${AuditUtils.escapeHtml(config.name)}</b> will be skipped.</p>`;
    
    safeSendEmail({
      to: config.recipients,
      cc: config.cc || '',
      subject,
      htmlBody
    }, `${config.name} - Missing Gmail Label`);
    
    return null;
  }

  static handleGenericError(error, context, config) {
    const logger = AuditUtils.createLogger(config.name);
    logger.error(`${context}: ${error.message}`);
    
    AuditErrorHandler.sendErrorNotification(error, config);
  }

  static handleEmailError(error, context, config) {
    const logger = AuditUtils.createLogger(config.name);
    logger.error(`Email error - ${context}: ${error.message}`);
    
    // Don't send error notification for email errors to avoid infinite loops
  }
}

// === 📁 FILE PROCESSING PIPELINE ===
class FileProcessor {
  static processGmailAttachments(config) {
    const logger = AuditUtils.createLogger(config.name);
    logger.log('fetchDailyAuditAttachments started');

    const label = GmailApp.getUserLabelByName(config.label);
    if (!label) {
      logger.warn(`Label not found: ${config.label}`);
      const subject = `⚠️ CM360 Audit Warning: Gmail Label Missing (${config.name})`;
      const htmlBody = `<p style="font-family:Arial; font-size:13px;">The label <b>${AuditUtils.escapeHtml(config.label)}</b> could not be found. This may mean the audit for <b>${AuditUtils.escapeHtml(config.name)}</b> will be skipped.</p>`;
      safeSendEmail({
        to: config.recipients,
        cc: config.cc || '',
        subject,
        htmlBody
      }, `${config.name} - Missing Gmail Label`);
      return null;
    }

    const threads = label.getThreads();
    const startOfToday = new Date();
    startOfToday.setHours(0, 0, 0, 0);  
    
    const parentFolder = getDriveFolderByPath_(config.tempDailyFolderPath);
    const timestamp = AuditUtils.formatDate(new Date(), 'yyyyMMdd_HHmmss');
    const driveFolder = parentFolder.createFolder(`Temp_CM360_${timestamp}`);

    let processedFiles = 0;
    threads.forEach(thread => {
      thread.getMessages().forEach(message => {
        if (message.getDate() < startOfToday) return;

        message.getAttachments({ includeInlineImages: false }).forEach(file => {
          const name = file.getName();
          const type = file.getContentType();
          
          if (FileProcessor.validateFileFormat(name, type)) {
            logger.log(`Processing attachment: ${name}`);
            driveFolder.createFile(file);
            processedFiles++;
          }
        });
      });
    });

    if (processedFiles === 0) {
      driveFolder.setTrashed(true);
      return null;
    }

    logger.log(`Processed ${processedFiles} files`);
    return driveFolder.getId();
  }

  static validateFileFormat(filename, contentType) {
    const name = String(filename || '').toLowerCase();
    const type = String(contentType || '').toLowerCase();
    
    return (name.endsWith('.xlsx') || name.endsWith('.xls') || name.endsWith('.zip')) ||
           (type.includes('spreadsheet') || type.includes('excel') || type.includes('zip'));
  }

  static mergeExcelFiles(folderId, mergedFolderPath, configName = 'Unknown') {
    const logger = AuditUtils.createLogger(configName);
    logger.log('FileProcessor.mergeExcelFiles started');

    try {
      logger.log(`Attempting to access folder with ID: ${folderId}`);
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFiles();
      const allData = [];

      while (files.hasNext()) {
        const file = files.next();
        const name = file.getName();
        
        if (name.toLowerCase().includes('.zip')) {
          // Handle ZIP files - extract and process
          FileProcessor.processZipFile(file, allData, logger);
        } else if (name.toLowerCase().includes('.xlsx') || name.toLowerCase().includes('.xls')) {
          // Handle Excel files directly
          FileProcessor.processExcelFile(file, allData, logger);
        }
      }

      if (allData.length === 0) {
        throw new Error('No data found in processed files');
      }

      // Create merged spreadsheet
      const mergedFolder = getDriveFolderByPath_(mergedFolderPath);
      const timestamp = AuditUtils.formatDate(new Date(), 'yyyy-MM-dd_HH-mm-ss');
      const mergedSpreadsheet = SpreadsheetApp.create(`CM360_Merged_${configName}_${timestamp}`);
      
      // Move to correct folder
      DriveApp.getFileById(mergedSpreadsheet.getId()).moveTo(mergedFolder);
      
      // Write data to sheet with standardized format
      const sheet = mergedSpreadsheet.getSheets()[0];
      if (allData.length > 0) {
        // Standardize headers and clean data
        const cleanedData = FileProcessor.standardizeReportFormat(allData);
        if (cleanedData.length > 0) {
          // Ensure all rows have the same number of columns
          const maxColumns = Math.max(...cleanedData.map(row => row.length));
          const normalizedData = cleanedData.map(row => {
            const normalizedRow = [...row];
            while (normalizedRow.length < maxColumns) {
              normalizedRow.push(''); // Pad with empty strings
            }
            return normalizedRow;
          });
          
          sheet.getRange(1, 1, normalizedData.length, maxColumns).setValues(normalizedData);
          // Format header row
          const headerRange = sheet.getRange(1, 1, 1, maxColumns);
          headerRange.setFontWeight('bold').setBackground('#E8F0FE');
        }
      }

      logger.log(`Merged sheet created: ${mergedSpreadsheet.getUrl()}`);
      return mergedSpreadsheet.getId();
    } catch (error) {
      logger.error(`Error in mergeExcelFiles: ${error.message}`);
      throw error;
    }
  }

  static processExcelFile(file, allData, logger) {
    try {
      logger.log(`Processing Excel file: ${file.getName()}`);
      
      // Convert Excel file to Google Sheets format
      const blob = file.getBlob();
      const tempFolder = DriveApp.getFolderById(file.getParents().next().getId());
      
      // Create a temporary Google Sheet from the Excel file
      const resource = {
        title: `temp_${file.getName()}`,
        mimeType: MimeType.GOOGLE_SHEETS,
        parents: [{ id: tempFolder.getId() }]
      };
      
      const tempFile = Drive.Files.insert(resource, blob, { convert: true });
      const tempSheet = SpreadsheetApp.openById(tempFile.id);
      
      // Get data from the first sheet
      const sheet = tempSheet.getSheets()[0];
      const data = sheet.getDataRange().getValues();
      
      if (data.length > 0) {
        if (allData.length === 0) {
          // First file - include headers
          allData.push(...data);
        } else {
          // Subsequent files - skip header row
          allData.push(...data.slice(1));
        }
        logger.log(`Added ${data.length} rows from ${file.getName()}`);
      }
      
      // Clean up temporary file
      DriveApp.getFileById(tempFile.id).setTrashed(true);
      
    } catch (error) {
      logger.error(`Failed to process Excel file ${file.getName()}: ${error.message}`);
    }
  }

  static standardizeReportFormat(rawData) {
    if (rawData.length === 0) return [];
    
    // Standard header format for CM360 reports
    const standardHeaders = [
      'Advertiser', 'Campaign', 'Site (CM360)', 'Placement ID', 'Placement', 
      'Placement Start Date', 'Placement End Date', 'Ad Type', 'Creative', 
      'Placement Pixel Size', 'Creative Pixel Size', 'Date', 'Impressions', 'Clicks', 'Flags'
    ];
    
    // Find the header row and map columns
    let headerRowIndex = -1;
    let columnMapping = {};
    
    for (let i = 0; i < Math.min(rawData.length, 20); i++) {
      const row = rawData[i];
      const rowText = row.join('|').toLowerCase();
      
      // Look for key indicators of a data header row
      if ((rowText.includes('advertiser') || rowText.includes('placement')) && 
          rowText.includes('impressions') && rowText.includes('clicks')) {
        headerRowIndex = i;
        
        // Map raw columns to standard positions
        standardHeaders.forEach((standardCol, stdIndex) => {
          const stdColLower = standardCol.toLowerCase();
          for (let rawIndex = 0; rawIndex < row.length; rawIndex++) {
            const rawCol = String(row[rawIndex] || '').toLowerCase();
            
            // Map common variations with more specific matching
            if ((stdColLower.includes('advertiser') && rawCol.includes('advertiser')) ||
                (stdColLower.includes('campaign') && rawCol.includes('campaign')) ||
                (stdColLower.includes('site') && rawCol.includes('site')) ||
                (stdColLower.includes('placement id') && (rawCol.includes('placement id') || rawCol.includes('placement_id'))) ||
                (stdColLower.includes('placement start date') && rawCol.includes('placement') && rawCol.includes('start') && rawCol.includes('date')) ||
                (stdColLower.includes('placement end date') && rawCol.includes('placement') && rawCol.includes('end') && rawCol.includes('date')) ||
                (stdColLower.includes('placement') && rawCol.includes('placement') && !rawCol.includes('id') && !rawCol.includes('date') && !rawCol.includes('pixel')) ||
                (stdColLower.includes('ad type') && rawCol.includes('ad') && rawCol.includes('type')) ||
                (stdColLower.includes('creative') && rawCol.includes('creative') && !rawCol.includes('pixel')) ||
                (stdColLower.includes('placement pixel') && rawCol.includes('placement') && rawCol.includes('pixel')) ||
                (stdColLower.includes('creative pixel') && rawCol.includes('creative') && rawCol.includes('pixel')) ||
                (stdColLower.includes('date') && rawCol === 'date') ||
                (stdColLower.includes('impressions') && rawCol.includes('impressions')) ||
                (stdColLower.includes('clicks') && rawCol.includes('clicks'))) {
              columnMapping[stdIndex] = rawIndex;
              break;
            }
          }
        });
        break;
      }
    }
    
    if (headerRowIndex === -1) {
      // If no proper header found, return raw data with standard header
      const headerRow = [...standardHeaders];
      const dataRows = rawData.slice(1).map(row => {
        const standardRow = new Array(standardHeaders.length).fill('');
        for (let i = 0; i < Math.min(row.length, standardRow.length - 1); i++) { // -1 to leave space for Flags
          standardRow[i] = row[i] || '';
        }
        return standardRow;
      });
      return [headerRow, ...dataRows];
    }
    
    // Process data rows
    const cleanedData = [standardHeaders];
    const dataRows = rawData.slice(headerRowIndex + 1);
    
    dataRows.forEach(row => {
      // Skip total rows, empty rows, filter info, and metadata
      const rowText = row.join('|').toLowerCase();
      if (rowText.includes('grand total') || 
          rowText.includes('total:') ||
          rowText.includes('filter') ||
          rowText.includes('report run') ||
          rowText.includes('date/time generated') ||
          rowText.includes('report time zone') ||
          rowText.includes('account id') ||
          rowText.includes('date range') ||
          rowText.includes('ad type	equals:') ||
          rowText.includes('advertiser	equals:') ||
          rowText.includes('mrc accredited') ||
          rowText.includes('reporting numbers') ||
          rowText.includes('report fields') ||
          rowText.includes('report metrics') ||
          row.every(cell => !cell || String(cell).trim() === '') ||
          (!row.some(cell => /^\d{8,}$/.test(String(cell))))) { // Must have placement ID pattern
        return;
      }
      
      // Additional check: must be actual data row with placement ID in column 3 or 4
      const hasValidPlacementId = row.some((cell, index) => {
        return index >= 2 && index <= 4 && /^\d{8,}$/.test(String(cell));
      });
      
      if (!hasValidPlacementId) return;
      
      // Skip rows where impressions and clicks are both 0 or empty
      const impressions = parseInt(String(row[12] || row[13] || '0')) || 0;
      const clicks = parseInt(String(row[13] || row[14] || '0')) || 0;
      
      // Skip if this looks like a header or summary row based on content
      const hasLongPlacementName = row.some(cell => {
        const cellStr = String(cell || '');
        return cellStr.length > 100 && cellStr.includes('_'); // Likely placement name in wrong column
      });
      
      if (hasLongPlacementName) return;
      
      // Map data to standard format
      const standardRow = new Array(standardHeaders.length).fill('');
      Object.keys(columnMapping).forEach(stdIndex => {
        const rawIndex = columnMapping[stdIndex];
        if (rawIndex < row.length) {
          standardRow[stdIndex] = row[rawIndex] || '';
        }
      });
      
      // Ensure the row has exactly the right number of columns
      while (standardRow.length < standardHeaders.length) {
        standardRow.push('');
      }
      
      cleanedData.push(standardRow);
    });
    
    return cleanedData;
  }

  static processZipFile(file, allData, logger) {
    try {
      logger.log(`Processing ZIP file: ${file.getName()}`);
      
      // Extract ZIP contents
      const blobs = Utilities.unzip(file.getBlob());
      let processedCount = 0;
      
      blobs.forEach(blob => {
        const fileName = blob.getName().toLowerCase();
        if (fileName.endsWith('.csv') || fileName.endsWith('.xlsx')) {
          logger.log(`Processing extracted file: ${blob.getName()}`);
          
          if (fileName.endsWith('.csv')) {
            // Process CSV directly
            const csvContent = blob.getDataAsString();
            const rows = csvContent.split('\n').map(row => row.split(','));
            
            if (rows.length > 0) {
              if (allData.length === 0) {
                // First file - include headers
                allData.push(...rows);
              } else {
                // Subsequent files - skip header row
                allData.push(...rows.slice(1));
              }
              processedCount++;
            }
          } else if (fileName.endsWith('.xlsx')) {
            // Convert Excel blob to temporary sheet and process
            try {
              const tempFolder = DriveApp.getFolderById(file.getParents().next().getId());
              const resource = {
                title: `temp_${blob.getName()}`,
                mimeType: MimeType.GOOGLE_SHEETS,
                parents: [{ id: tempFolder.getId() }]
              };
              
              const tempFile = Drive.Files.insert(resource, blob, { convert: true });
              const tempSheet = SpreadsheetApp.openById(tempFile.id);
              const data = tempSheet.getSheets()[0].getDataRange().getValues();
              
              if (data.length > 0) {
                if (allData.length === 0) {
                  allData.push(...data);
                } else {
                  allData.push(...data.slice(1));
                }
                processedCount++;
              }
              
              DriveApp.getFileById(tempFile.id).setTrashed(true);
            } catch (error) {
              logger.error(`Failed to process Excel from ZIP ${blob.getName()}: ${error.message}`);
            }
          }
        }
      });
      
      logger.log(`Extracted and processed ${processedCount} files from ZIP: ${file.getName()}`);
    } catch (error) {
      logger.error(`Failed to process ZIP file ${file.getName()}: ${error.message}`);
    }
  }
}

// === 🎯 AUDIT RULES ENGINE ===
class AuditRulesEngine {
  static analyzeRow(row, headers, config) {
    const flags = [];
    const rowData = AuditRulesEngine.parseRowData(row, headers);
    
    // Skip rows that don't represent actual placements
    if (!AuditRulesEngine.isValidPlacementRow(rowData)) {
      return flags; // Return empty flags array for non-placement rows
    }
    
    // Apply each audit rule
    if (config.flags.clicks_gt_impressions) {
      const clicksFlag = AuditRulesEngine.checkClicksGreaterThanImpressions(rowData);
      if (clicksFlag) flags.push(clicksFlag);
    }
    
    if (config.flags.out_of_flight) {
      const flightFlag = AuditRulesEngine.checkOutOfFlight(rowData);
      if (flightFlag) flags.push(flightFlag);
    }
    
    if (config.flags.pixel_mismatch) {
      const pixelFlag = AuditRulesEngine.checkPixelMismatch(rowData);
      if (pixelFlag) flags.push(pixelFlag);
    }
    
    if (config.flags.default_ad_serving) {
      const defaultFlag = AuditRulesEngine.checkDefaultAdServing(rowData);
      if (defaultFlag) flags.push(defaultFlag);
    }
    
    return flags;
  }
  
  static parseRowData(row, headers) {
    const data = {};
    headers.forEach((header, index) => {
      data[header] = row[index];
    });
    return data;
  }
  
  static checkClicksGreaterThanImpressions(rowData) {
    const clicks = parseInt(rowData['Clicks'] || rowData['clicks'] || '0') || 0;
    const impressions = parseInt(rowData['Impressions'] || rowData['impressions'] || '0') || 0;
    
    if (clicks > impressions && clicks > 0) {
      return {
        type: 'clicks_gt_impressions',
        severity: 'HIGH',
        message: `Clicks (${clicks}) exceed Impressions (${impressions})`,
        data: { clicks, impressions }
      };
    }
    return null;
  }
  
  static isValidPlacementRow(rowData) {
    // Check if this row represents an actual placement (not campaign-level data)
    const placementId = rowData['Placement ID'] || rowData['placement_id'] || rowData['Placement'];
    
    // Must have a placement ID that looks like a proper CM360 placement ID
    if (!placementId) return false;
    
    // CM360 placement IDs are typically 9+ digit numbers
    const placementIdStr = String(placementId).trim();
    
    // Skip if it's empty, contains letters/special chars, or is too short to be a real placement ID
    if (!placementIdStr || 
        !/^\d+$/.test(placementIdStr) || 
        placementIdStr.length < 8) {
      return false;
    }
    
    // Also check for required data fields - placements should have either impressions or clicks
    const impressions = parseInt(rowData['Impressions'] || rowData['impressions'] || '0') || 0;
    const clicks = parseInt(rowData['Clicks'] || rowData['clicks'] || '0') || 0;
    
    // Skip rows with no activity data (likely summary/header rows)
    if (impressions === 0 && clicks === 0) return false;
    
    return true;
  }
  
  static checkOutOfFlight(rowData) {
    // Reports show yesterday's data, so use yesterday's date for comparison
    const dataDate = new Date();
    dataDate.setDate(dataDate.getDate() - 1);
    // Reset time to start of day for accurate date-only comparison
    dataDate.setHours(0, 0, 0, 0);
    
    const startDate = AuditRulesEngine.parseDate(
      rowData['Placement Start Date'] || // Found in debug output!
      rowData['Campaign Start Date'] || 
      rowData['start_date'] || 
      rowData['Start Date'] || 
      rowData['Campaign Start'] ||
      rowData['Start'] ||
      rowData['Period Start Date'] ||
      rowData['Flight Start Date']
    );
    const endDate = AuditRulesEngine.parseDate(
      rowData['Placement End Date'] || // Found in debug output!
      rowData['Campaign End Date'] || 
      rowData['end_date'] || 
      rowData['End Date'] || 
      rowData['Campaign End'] ||
      rowData['End'] ||
      rowData['Period End Date'] ||
      rowData['Flight End Date']
    );
    
    // Reset time components for date-only comparison
    if (startDate) startDate.setHours(0, 0, 0, 0);
    if (endDate) endDate.setHours(0, 0, 0, 0);
    
    // Debug for the problematic placements
    const placementId = rowData['Placement ID'] || rowData['Placement'];
    if (placementId && ['424138145', '424138142'].includes(String(placementId))) {
      Logger.log(`[DEBUG] Placement ${placementId}`);
      Logger.log(`[DEBUG] Data date: ${dataDate.toDateString()} (${dataDate.getTime()})`);
      Logger.log(`[DEBUG] Start date: ${startDate ? startDate.toDateString() + ' (' + startDate.getTime() + ')' : 'null'}`);
      Logger.log(`[DEBUG] End date: ${endDate ? endDate.toDateString() + ' (' + endDate.getTime() + ')' : 'null'}`);
      Logger.log(`[DEBUG] Is out of flight? ${dataDate < startDate || dataDate > endDate}`);
    }
    
    if (startDate && endDate) {
      if (dataDate < startDate || dataDate > endDate) {
        return {
          type: 'out_of_flight',
          severity: 'MEDIUM',
          message: `Out of flight (${AuditUtils.formatDate(startDate, 'yyyy-MM-dd')} to ${AuditUtils.formatDate(endDate, 'yyyy-MM-dd')})`,
          data: { startDate, endDate, dataDate }
        };
      }
    }
    return null;
  }
  
  static checkPixelMismatch(rowData) {
    const pixelCounts = rowData['Pixel Count'] || rowData['pixel_count'];
    const expectedPixels = rowData['Expected Pixels'] || rowData['expected_pixels'];
    
    if (pixelCounts && expectedPixels && pixelCounts !== expectedPixels) {
      return {
        type: 'pixel_mismatch',
        severity: 'MEDIUM',
        message: `Pixel mismatch: Found ${pixelCounts}, Expected ${expectedPixels}`,
        data: { actual: pixelCounts, expected: expectedPixels }
      };
    }
    return null;
  }
  
  static checkDefaultAdServing(rowData) {
    const adName = (rowData['Ad Name'] || rowData['Creative Name'] || '').toLowerCase();
    const defaultIndicators = ['default', 'backup', 'fallback', 'placeholder'];
    
    if (defaultIndicators.some(indicator => adName.includes(indicator))) {
      const impressions = parseInt(rowData['Impressions'] || '0') || 0;
      if (impressions > 100) { // Threshold for concern
        return {
          type: 'default_ad_serving',
          severity: 'MEDIUM',
          message: `Default ad serving detected with ${impressions} impressions`,
          data: { adName, impressions }
        };
      }
    }
    return null;
  }
  
  static parseDate(dateInput) {
    if (!dateInput) return null;
    
    // If it's already a Date object, return it directly
    if (dateInput instanceof Date && !isNaN(dateInput.getTime())) {
      return dateInput;
    }
    
    // If it's a string, try to parse it
    const dateString = String(dateInput);
    
    // Try multiple date formats
    const formats = [
      /^(\d{4})-(\d{2})-(\d{2})$/,  // YYYY-MM-DD
      /^(\d{2})\/(\d{2})\/(\d{4})$/, // MM/DD/YYYY
      /^(\d{2})-(\d{2})-(\d{4})$/   // MM-DD-YYYY
    ];
    
    for (const format of formats) {
      const match = dateString.match(format);
      if (match) {
        if (format === formats[0]) { // YYYY-MM-DD
          return new Date(match[1], match[2] - 1, match[3]);
        } else { // MM/DD/YYYY or MM-DD-YYYY
          return new Date(match[3], match[1] - 1, match[2]);
        }
      }
    }
    
    // Fallback to JavaScript Date parsing
    const parsed = new Date(dateString);
    return isNaN(parsed.getTime()) ? null : parsed;
  }
}

// === 📊 AUDIT DATA PROCESSOR ===
class AuditDataProcessor {
  static processSheetData(spreadsheetId, config) {
    const logger = AuditUtils.createLogger(config.name);
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    
    if (data.length === 0) {
      throw new Error('No data found in spreadsheet');
    }
    
    // Find the actual header row by looking for expected columns
    let headerRowIndex = -1;
    let headers = [];
    
    // Search much deeper - CM360 reports can have lots of metadata
    for (let i = 0; i < Math.min(data.length, 100); i++) {
      const row = data[i];
      const rowText = row.join('|').toLowerCase();
      
      // Check if this row contains the key columns we expect
      const hasPlacementId = rowText.includes('placement id') || rowText.includes('placement_id') || rowText.includes('placement');
      const hasImpressions = rowText.includes('impressions') || rowText.includes('impression');
      const hasClicks = rowText.includes('clicks') || rowText.includes('click');
      
      // Also check for other common CM360 columns
      const hasAdvertiser = rowText.includes('advertiser');
      const hasCampaign = rowText.includes('campaign');
      const hasSite = rowText.includes('site');
      
      // Must have at least placement + impressions + clicks, or advertiser + campaign + impressions
      const isHeaderRow = (hasPlacementId && hasImpressions && hasClicks) || 
                         (hasAdvertiser && hasCampaign && hasImpressions);
      
      if (isHeaderRow) {
        headerRowIndex = i;
        headers = row;
        break;
      }
    }
    
    if (headerRowIndex === -1) {
      // If still not found, show more rows for debugging
      logger.log(`Could not find header row with expected columns`);
      throw new Error('Could not find header row with expected columns (Placement ID, Impressions, Clicks)');
    }
    
    const rows = data.slice(headerRowIndex + 1);
    const results = [];
    
    logger.log(`Processing ${rows.length} rows of data`);
    
    // Check exclusions
    const exclusions = CacheManager.getExclusions(config.name);
    
    rows.forEach((row, index) => {
      try {
        const rowKey = AuditDataProcessor.generateRowKey(row, headers);
        
        // Skip if in exclusions
        if (exclusions.has(rowKey)) {
          return;
        }
        
        // Apply audit rules
        const flags = AuditRulesEngine.analyzeRow(row, headers, config);
        
        if (flags.length > 0) {
          // Standardize row data for consistent structure
          const standardizedRow = AuditDataProcessor.standardizeRowData(row, headers);
          
          results.push({
            rowIndex: index + 2, // +2 for header and 0-based index
            rowData: standardizedRow,
            flags: flags,
            rowKey: rowKey
          });
        }
      } catch (error) {
        logger.error(`Error processing row ${index + 2}: ${error.message}`);
      }
    });
    
    logger.log(`Found ${results.length} flagged rows out of ${rows.length} total`);
    
    // Sort results by highest value (impressions or clicks), descending
    results.sort((a, b) => {
      const getHighestValue = (standardizedRow) => {
        // Use fixed positions for standardized data: [12] = Impressions, [13] = Clicks
        const impressions = parseInt(standardizedRow[12] || '0') || 0;
        const clicks = parseInt(standardizedRow[13] || '0') || 0;
        return Math.max(impressions, clicks);
      };
      
      const aHighest = getHighestValue(a.rowData);
      const bHighest = getHighestValue(b.rowData);
      
      return bHighest - aHighest; // Descending order
    });
    
    return {
      headers: ['Advertiser', 'Campaign', 'Site (CM360)', 'Placement ID', 'Placement', 'Placement Start Date', 'Placement End Date', 'Ad Type', 'Creative', 'Placement Pixel Size', 'Creative Pixel Size', 'Date', 'Impressions', 'Clicks', 'Flags'],
      allData: data.slice(headerRowIndex), // Include header row and all data rows (this will be standardized in Excel creation)
      flaggedRows: results,
      totalRows: rows.length,
      spreadsheetUrl: spreadsheet.getUrl()
    };
  }
  
  static generateRowKey(row, headers) {
    // Create a unique key for the row based on key columns
    const keyColumns = ['Campaign', 'Ad', 'Date', 'Placement']; // Adjust based on data
    const keyValues = [];
    
    keyColumns.forEach(col => {
      const index = headers.indexOf(col);
      if (index !== -1) {
        keyValues.push(String(row[index] || '').trim());
      }
    });
    
    return keyValues.join('|');
  }
  
  static getColumnValue(row, headers, possibleNames) {
    for (const name of possibleNames) {
      const index = headers.findIndex(h => h && h.toLowerCase().includes(name.toLowerCase()));
      if (index !== -1 && row[index] !== undefined && row[index] !== null) {
        return row[index];
      }
    }
    return null;
  }
  
  static standardizeRowData(row, originalHeaders) {
    // Map row data to standard format
    const standardHeaders = ['Advertiser', 'Campaign', 'Site (CM360)', 'Placement ID', 'Placement', 'Placement Start Date', 'Placement End Date', 'Ad Type', 'Creative', 'Placement Pixel Size', 'Creative Pixel Size', 'Date', 'Impressions', 'Clicks', 'Flags'];
    const standardRow = new Array(standardHeaders.length).fill('');
    
    // Map each standard column to the original data
    const mapping = {
      0: ['Advertiser', 'advertiser'],
      1: ['Campaign', 'campaign'],
      2: ['Site', 'site', 'Site (CM360)'],
      3: ['Placement ID', 'placement_id', 'Placement_ID'],
      4: ['Placement', 'placement'],
      5: ['Placement Start Date', 'Start Date', 'start_date'],
      6: ['Placement End Date', 'End Date', 'end_date'],
      7: ['Ad Type', 'ad_type'],
      8: ['Creative', 'creative'],
      9: ['Placement Pixel Size', 'placement_pixel'],
      10: ['Creative Pixel Size', 'creative_pixel'],
      11: ['Date', 'date'],
      12: ['Impressions', 'impressions'],
      13: ['Clicks', 'clicks'],
      14: [] // Flags - will be filled later
    };
    
    Object.keys(mapping).forEach(stdIndex => {
      const possibleNames = mapping[stdIndex];
      if (possibleNames.length > 0) {
        const value = AuditDataProcessor.getColumnValue(row, originalHeaders, possibleNames);
        standardRow[stdIndex] = value || '';
      }
    });
    
    return standardRow;
  }
}

// === 📄 RESULT FORMATTER ===
class ResultFormatter {
  static generateHtmlReport(auditResults, config) {
    const { headers, flaggedRows, totalRows, spreadsheetUrl } = auditResults;
    
    const style = ResultFormatter.getEmailStyles();
    const summary = ResultFormatter.generateSummary(flaggedRows, totalRows, headers);
    const table = ResultFormatter.generateFlaggedRowsTable(flaggedRows, headers);
    
    return `
      ${style}
      <div class="container">
        <div class="header">
          <h1>CM360 Audit Alert</h1>
          <h2>${AuditUtils.escapeHtml(config.name)}</h2>
          <p class="date">Generated on ${AuditUtils.formatDate(new Date(), 'yyyy-MM-dd HH:mm:ss')}</p>
        </div>
        ${summary}
        ${table}
        <div class="footer">
          <p style="text-align: center;"><a href="${spreadsheetUrl}">View Source Data</a></p>
          <p style="text-align: center;"><em>CM360 Audit System - Automated Quality Assurance</em></p>
          <p style="text-align: center;"><em>Platform Solutions</em></p>
        </div>
      </div>
    `;
  }
  
  static getEmailStyles() {
    return `
      <style>
        .container { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; }
        .header { background: #f8f9fa; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .header h1 { color: #1a73e8; margin: 0; }
        .header h2 { color: #5f6368; margin: 5px 0; }
        .date { color: #80868b; font-size: 14px; margin: 0; }
        .summary { background: #e8f0fe; padding: 15px; border-radius: 6px; margin-bottom: 20px; }
        .severity-high { background: #fce8e6; border-left: 4px solid #d93025; }
        .severity-medium { background: #fef7e0; border-left: 4px solid #f9ab00; }
        .severity-low { background: #e6f4ea; border-left: 4px solid #137333; }
        .flag-item { margin: 10px 0; padding: 10px; border-radius: 4px; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #dadce0; text-align: center; color: #5f6368; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; }
        th, td { padding: 8px; text-align: left; border-bottom: 1px solid #dadce0; }
        th { background: #f8f9fa; }
      </style>
    `;
  }
  
  static generateSummary(flaggedRows, totalRows, headers) {
    const total = flaggedRows.length;
    
    // Get unique campaigns from flagged rows using the correct headers
    const campaigns = flaggedRows.map(row => {
      return AuditDataProcessor.getColumnValue(row.rowData, headers, ['Campaign', 'campaign', 'Campaign Name', 'campaign_name']);
    }).filter(campaign => campaign && String(campaign).trim() !== '');
    
    const uniqueCampaigns = new Set(campaigns).size;
    
    const plural = (count, singular, plural) => count === 1 ? singular : plural;
    const summaryText = `The following ${total} ${plural(total, 'placement', 'placements')} across ${uniqueCampaigns} ${plural(uniqueCampaigns, 'campaign', 'campaigns')} ${plural(total, 'was', 'were')} flagged during the audit. Please review:`;
    
    return `
      <div class="summary">
        <p>${summaryText}</p>
        ${total > 100 ? '<p><em>Only the first 100 flagged rows are shown below. Full details are included in the attached Excel file.</em></p>' : ''}
      </div>
    `;
  }
  
  static generateFlaggedRowsTable(flaggedRows, headers) {
    const displayRows = flaggedRows.slice(0, 100); // Limit to first 100 rows
    
    let html = `
      <table border="1" cellpadding="8" cellspacing="0" style="width:100%; border-collapse:collapse; font-family:Arial,sans-serif; font-size:12px;">
        <thead style="background-color:#f2f2f2;">
          <tr>
            <th style="padding:8px; width:120px;">Advertiser</th>
            <th style="padding:8px; width:160px;">Campaign</th>
            <th style="padding:8px; width:100px;">Site</th>
            <th style="padding:8px; width:140px;">Placement</th>
            <th style="padding:8px; width:100px;">Placement ID</th>
            <th style="padding:8px; width:80px;">Start Date</th>
            <th style="padding:8px; width:80px;">End Date</th>
            <th style="padding:8px; width:120px;">Creative</th>
            <th style="padding:8px; width:80px;">Impressions</th>
            <th style="padding:8px; width:80px;">Clicks</th>
            <th style="padding:8px; width:200px;">Issues</th>
          </tr>
        </thead>
        <tbody>
    `;
    
    displayRows.forEach(result => {
      const row = result.rowData;
      
      // Use standardized column positions instead of dynamic lookup
      // Based on standard headers: Advertiser, Campaign, Site (CM360), Placement ID, Placement, 
      // Placement Start Date, Placement End Date, Ad Type, Creative, Placement Pixel Size, Creative Pixel Size, Date, Impressions, Clicks, Flags
      const advertiser = row[0] || '';
      const campaign = row[1] || '';
      const site = row[2] || '';
      const placementId = row[3] || '';
      const placement = row[4] || '';
      const startDate = row[5] || '';
      const endDate = row[6] || '';
      const creative = row[8] || ''; // Skip Ad Type (index 7)
      const impressions = parseInt(row[12] || '0') || 0;
      const clicks = parseInt(row[13] || '0') || 0;
      
      const issues = result.flags.map(flag => flag.message).join('<br>');
      
      // Format dates for display
      const formatDate = (dateStr) => {
        if (!dateStr) return '';
        try {
          const date = new Date(dateStr);
          return isNaN(date.getTime()) ? dateStr : date.toLocaleDateString();
        } catch {
          return dateStr;
        }
      };
      
      html += `
        <tr>
          <td style="padding:6px;">${AuditUtils.escapeHtml(advertiser)}</td>
          <td style="padding:6px;">${AuditUtils.escapeHtml(campaign)}</td>
          <td style="padding:6px;">${AuditUtils.escapeHtml(site)}</td>
          <td style="padding:6px;">${AuditUtils.escapeHtml(placement)}</td>
          <td style="padding:6px;">${AuditUtils.escapeHtml(placementId)}</td>
          <td style="padding:6px;">${AuditUtils.escapeHtml(formatDate(startDate))}</td>
          <td style="padding:6px;">${AuditUtils.escapeHtml(formatDate(endDate))}</td>
          <td style="padding:6px;">${AuditUtils.escapeHtml(creative)}</td>
          <td style="padding:6px; text-align:right;">${impressions.toLocaleString()}</td>
          <td style="padding:6px; text-align:right;">${clicks.toLocaleString()}</td>
          <td style="padding:6px; font-size:11px;">${issues}</td>
        </tr>
      `;
    });
    
    html += `
        </tbody>
      </table>
    `;
    
    return html;
  }
}

// === 🔧 AUDIT CONFIGS ===
// TODO: Move to external configuration sheet for better maintainability
class ConfigManager {
  static getAuditConfigs() {
    // In future: load from 'CM360 Audit Configs' sheet
    return auditConfigs;
  }
  
  static getConfigByName(name) {
    return this.getAuditConfigs().find(config => config.name === name);
  }
  
  static validateConfig(config) {
    const requiredFields = ['name', 'label', 'recipients', 'flags'];
    return requiredFields.every(field => config[field] !== undefined);
  }
}

const auditConfigs = [
  {
    name: 'PST01',
    label: 'Daily Audits/CM360/PST01',
    mergedFolderPath: folderPath('Merged Reports', 'PST01'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'PST01'),
    recipients: resolveRecipients(ADMIN_EMAIL),
    cc: resolveCc([]),
    flags: { 
      minImpThreshold: 50, 
      minClickThreshold: 10,
      clicks_gt_impressions: true,
      out_of_flight: true,
      pixel_mismatch: true,
      default_ad_serving: true
    }
  },
  {
    name: 'PST02',
    label: 'Daily Audits/CM360/PST02',
    mergedFolderPath: folderPath('Merged Reports', 'PST02'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'PST02'),
    recipients: resolveRecipients('fvariath@horizonmedia.com'),
    cc: resolveCc([ADMIN_EMAIL]),
    flags: { 
      minImpThreshold: 100, 
      minClickThreshold: 100,
      clicks_gt_impressions: true,
      out_of_flight: true,
      pixel_mismatch: true,
      default_ad_serving: true
    }
  },
  {
    name: 'PST03',
    label: 'Daily Audits/CM360/PST03',
    mergedFolderPath: folderPath('Merged Reports', 'PST03'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'PST03'),
    recipients: resolveRecipients('dmaestre@horizonmedia.com'),
    cc: resolveCc([ADMIN_EMAIL]),
    flags: { 
      minImpThreshold: 0, 
      minClickThreshold: 0,
      clicks_gt_impressions: true,
      out_of_flight: true,
      pixel_mismatch: true,
      default_ad_serving: true
    }
  },
  {
    name: 'NEXT01',
    label: 'Daily Audits/CM360/NEXT01',
    mergedFolderPath: folderPath('Merged Reports', 'NEXT01'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'NEXT01'),
    recipients: resolveRecipients('bosborne@horizonmedia.com, mmassaroni@horizonmedia.com'),
    cc: resolveCc([ADMIN_EMAIL]),
    flags: { 
      minImpThreshold: 1200, 
      minClickThreshold: 1200,
      clicks_gt_impressions: true,
      out_of_flight: true,
      pixel_mismatch: true,
      default_ad_serving: true
    }
  },
  {
    name: 'NEXT02',
    label: 'Daily Audits/CM360/NEXT02',
    mergedFolderPath: folderPath('Merged Reports', 'NEXT02'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'NEXT02'),
    recipients: resolveRecipients('rschaff@horizonmedia.com, mmassaroni@horizonmedia.com, jwong@horizonmedia.com'),
    cc: resolveCc([ADMIN_EMAIL]),
    flags: { 
      minImpThreshold: 0, 
      minClickThreshold: 0,
      clicks_gt_impressions: true,
      out_of_flight: true,
      pixel_mismatch: true,
      default_ad_serving: true
    }
  },
  {
    name: 'NEXT03',
    label: 'Daily Audits/CM360/NEXT03',
    mergedFolderPath: folderPath('Merged Reports', 'NEXT03'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'NEXT03'),
    recipients: resolveRecipients('szeterberg@horizonmedia.com, mmassaroni@horizonmedia.com, jwong@horizonmedia.com'),
    cc: resolveCc([ADMIN_EMAIL]),
    flags: { 
      minImpThreshold: 0, 
      minClickThreshold: 0,
      clicks_gt_impressions: true,
      out_of_flight: true,
      pixel_mismatch: true,
      default_ad_serving: true
    }
  },
  {
    name: 'SPTM01',
    label: 'Daily Audits/CM360/SPTM01',
    mergedFolderPath: folderPath('Merged Reports', 'SPTM01'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'SPTM01'),
    recipients: resolveRecipients('spectrum_adops@horizonmedia.com'),
    cc: resolveCc([ADMIN_EMAIL]),
    flags: { 
      minImpThreshold: 10, 
      minClickThreshold: 10,
      clicks_gt_impressions: true,
      out_of_flight: true,
      pixel_mismatch: true,
      default_ad_serving: true
    }
  },
  {
    name: 'NFL01',
    label: 'Daily Audits/CM360/NFL01',
    mergedFolderPath: folderPath('Merged Reports', 'NFL01'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'NFL01'),
    recipients: resolveRecipients('NFL_AdOps@horizonmedia.com, sbermolone@horizonmedia.com'),
    cc: resolveCc([ADMIN_EMAIL]),
    flags: { 
      minImpThreshold: 50, 
      minClickThreshold: 50,
      clicks_gt_impressions: true,
      out_of_flight: true,
      pixel_mismatch: true,
      default_ad_serving: true
    }
  },
  {
    name: 'ENT01',
    label: 'Daily Audits/CM360/ENT01',
    mergedFolderPath: folderPath('Merged Reports', 'ENT01'),
    tempDailyFolderPath: folderPath('Temp Daily Reports', 'ENT01'),
    recipients: resolveRecipients('sremick@horizonmedia.com, cali@horizonmedia.com'),
    cc: resolveCc([ADMIN_EMAIL]),
    flags: { 
      minImpThreshold: 15, 
      minClickThreshold: 15,
      clicks_gt_impressions: true,
      out_of_flight: true,
      pixel_mismatch: true,
      default_ad_serving: true
    }
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

// Legacy wrapper - use AuditUtils.formatDate instead
function formatDate(date = new Date(), pattern = 'yyyy-MM-dd') {
  return AuditUtils.formatDate(date, pattern);
}

// Legacy wrapper - use AuditUtils.escapeHtml instead
function escapeHtml(text) {
  return AuditUtils.escapeHtml(text);
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
  const requiredFields = ['name', 'label', 'recipients'];
  auditConfigs.forEach(config => {
    requiredFields.forEach(field => {
      if (!config[field] || typeof config[field] !== 'string') {
        throw new Error(`❌ Invalid audit config "${config.name || '[unnamed]'}": Missing or invalid "${field}"`);
      }
    });
  });
}

// === 📋 EXCLUSIONS MANAGEMENT ===
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
        'Flag Type',
        'Reason',
        'Added By',
        'Date Added',
        'Active',
        '', // Blank column
        'INSTRUCTIONS'
      ];
      
      Logger.log('Setting headers...');
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format the main headers
      const mainHeaderRange = sheet.getRange(1, 1, 1, 8);
      mainHeaderRange.setFontWeight('bold');
      mainHeaderRange.setBackground('#4285f4');
      mainHeaderRange.setFontColor('#ffffff');
      
      // Format the instructions header
      const instructionsHeaderRange = sheet.getRange(1, 10, 1, 1);
      instructionsHeaderRange.setFontWeight('bold');
      instructionsHeaderRange.setBackground('#ff9900');
      instructionsHeaderRange.setFontColor('#ffffff');
      
      Logger.log('Headers formatted, setting up protection...');
      
      // Lock the Placement Name column (column C)
      const placementNameRange = sheet.getRange('C:C');
      const protection = placementNameRange.protect().setDescription('Placement Name (Auto-populated - Do Not Edit)');
      protection.setWarningOnly(true);
      
      Logger.log('Setting up dropdowns...');
      
      // Add dropdown validation for Flag Type column (column D) - starting from row 2
      const flagTypeRange = sheet.getRange('D2:D');
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
      
      // Add dropdown validation for Active column (column H) - starting from row 2
      const activeRange = sheet.getRange('H2:H');
      const activeOptions = ['TRUE', 'FALSE'];
      
      const activeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(activeOptions)
        .setAllowInvalid(false)
        .setHelpText('Select TRUE to enable exclusion or FALSE to disable.')
        .build();
      
      activeRange.setDataValidation(activeRule);
      
      Logger.log('Adding sample data...');
      
      // Add some sample data
      const currentDate = formatDate(new Date());
      const sampleData = [
        ['PST01', '12345678', '', 'clicks_greater_than_impressions', 'Social placement - clicks > impressions expected', 'admin', currentDate, 'TRUE', '', 'Config Name: Must match exactly (PST01, PST02, etc.)'],
        ['PST01', '87654321', '', 'all_flags', 'Test placement - exclude from all auditing', 'admin', currentDate, 'TRUE', '', 'Placement ID: The CM360 placement ID to exclude'],
        ['NFL01', '11111111', '', 'pixel_size_mismatch', 'Rich media creative - intentional size difference', 'admin', currentDate, 'FALSE', '', 'Placement Name: Auto-populated (DO NOT EDIT)'],
        ['', '', '', '', '', '', '', '', '', 'Flag Type: Use dropdown to select flag type'],
        ['', '', '', '', '', '', '', '', '', '  Available options: clicks_greater_than_impressions,'],
        ['', '', '', '', '', '', '', '', '', '  out_of_flight_dates, pixel_size_mismatch,'],
        ['', '', '', '', '', '', '', '', '', '  default_ad_serving, or all_flags'],
        ['', '', '', '', '', '', '', '', '', 'Active: Use dropdown - TRUE to enable, FALSE to disable'],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', 'TIP: Use CM360 Audit > Refresh Placement Names'],
        ['', '', '', '', '', '', '', '', '', 'to manually update placement names if needed.'],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', 'NOTE: Placement Name column is protected'],
        ['', '', '', '', '', '', '', '', '', 'and will auto-populate when you add data.']
      ];
      
      Logger.log(`Adding ${sampleData.length} rows of sample data...`);
      sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
      
      Logger.log('Formatting columns...');
      
      // Auto-resize columns
      sheet.autoResizeColumns(1, headers.length);
      
      // Set column widths for better readability
      sheet.setColumnWidth(1, 100); // Config Name
      sheet.setColumnWidth(2, 120); // Placement ID
      sheet.setColumnWidth(3, 200); // Placement Name
      sheet.setColumnWidth(4, 180); // Flag Type
      sheet.setColumnWidth(5, 250); // Reason
      sheet.setColumnWidth(6, 100); // Added By
      sheet.setColumnWidth(7, 120); // Date Added
      sheet.setColumnWidth(8, 80);  // Active
      sheet.setColumnWidth(9, 30);  // Blank
      sheet.setColumnWidth(10, 300); // Instructions
      
      Logger.log(`✅ Created exclusions sheet: ${EXCLUSIONS_SHEET_NAME}`);
    } else {
      Logger.log(`✅ Exclusions sheet already exists: ${EXCLUSIONS_SHEET_NAME}`);
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
      sheet.getRange(row, 3).setValue(placementName); // Column 3 is Placement Name
      
      Logger.log(`Auto-populated placement name: ${configName} - ${placementId} -> ${placementName}`);
    } else if (!configName || !placementId) {
      // Clear placement name if either config or placement ID is empty
      sheet.getRange(row, 3).setValue('');
    }
    
  } catch (error) {
    Logger.log(`Error in onEdit: ${error.message}`);
    // Don't throw error to avoid disrupting user experience
  }
}

// Legacy function for backward compatibility
function onEditExclusionsSheet(e) {
  // This function is no longer needed as we use the simpler onEdit trigger
  onEdit(e);
}

function loadExclusionsFromSheet() {
  try {
    const sheet = getOrCreateExclusionsSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      Logger.log('📋 No exclusion data found in sheet');
      return {};
    }
    
    const headers = data[0];
    const configColIndex = headers.indexOf('Config Name');
    const placementIdColIndex = headers.indexOf('Placement ID');
    const flagTypeColIndex = headers.indexOf('Flag Type');
    const activeColIndex = headers.indexOf('Active');
    
    if (configColIndex === -1 || placementIdColIndex === -1 || flagTypeColIndex === -1) {
      Logger.log('❌ Required columns not found in exclusions sheet');
      Logger.log(`Found headers: ${headers.join(', ')}`);
      return {};
    }
    
    const exclusions = {};
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const configName = String(row[configColIndex] || '').trim();
      const placementId = String(row[placementIdColIndex] || '').trim();
      const flagType = String(row[flagTypeColIndex] || '').trim();
      const isActive = String(row[activeColIndex] || '').toUpperCase() === 'TRUE';
      
      // Skip empty rows, instruction rows, or inactive exclusions
      if (!configName || !placementId || !flagType || !isActive) {
        continue;
      }
      
      // Skip rows that look like instructions
      if (configName.includes('INSTRUCTIONS') || configName.includes('•') || configName.includes('Config Name:')) {
        continue;
      }
      
      // Initialize config if not exists
      if (!exclusions[configName]) {
        exclusions[configName] = {
          'clicks_greater_than_impressions': [],
          'out_of_flight_dates': [],
          'pixel_size_mismatch': [],
          'default_ad_serving': [],
          'all_flags': []
        };
      }
      
      // Add to appropriate flag type array
      if (exclusions[configName][flagType]) {
        exclusions[configName][flagType].push(placementId);
        Logger.log(`📋 Loaded exclusion: ${configName} - ${placementId} - ${flagType}`);
      } else {
        Logger.log(`⚠️ Unknown flag type "${flagType}" for ${configName} - ${placementId}`);
      }
    }
    
    Logger.log(`✅ Loaded exclusions for ${Object.keys(exclusions).length} config(s)`);
    return exclusions;
    
  } catch (error) {
    Logger.log(`❌ Error loading exclusions from sheet: ${error.message}`);
    return {};
  }
}

// Helper function to check if a placement ID should be excluded for a specific flag type
function isPlacementExcludedForFlag(exclusionsData, configName, placementId, flagType) {
  if (!exclusionsData || !exclusionsData[configName]) return false;
  
  const trimmedId = String(placementId || '').trim();
  const configExclusions = exclusionsData[configName];
  
  // Check if excluded from all flags
  if (configExclusions.all_flags && 
      configExclusions.all_flags.some(id => String(id).trim() === trimmedId)) {
    return true;
  }
  
  // Check if excluded from specific flag type
  if (configExclusions[flagType] && 
      configExclusions[flagType].some(id => String(id).trim() === trimmedId)) {
    return true;
  }
  
  return false;
}

// Function to lookup placement name from recent audit data
function LOOKUP_PLACEMENT_NAME(configName, placementId) {
  if (!configName || !placementId) return '';
  
  try {
    // Find the most recent merged report for this config
    const config = auditConfigs.find(c => c.name === configName);
    if (!config) return `Config "${configName}" not found`;
    
    const mergedFolder = getDriveFolderByPath_(config.mergedFolderPath);
    if (!mergedFolder) return 'Merged folder not found';
    
    // Get the most recent merged file
    const files = mergedFolder.getFiles();
    let mostRecentFile = null;
    let mostRecentDate = new Date(0);
    
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().startsWith('Merged_CM360_') && file.getDateCreated() > mostRecentDate) {
        mostRecentFile = file;
        mostRecentDate = file.getDateCreated();
      }
    }
    
    if (!mostRecentFile) return 'No recent audit data';
    
    // Open the spreadsheet and search for the placement ID
    const spreadsheet = SpreadsheetApp.open(mostRecentFile);
    const sheet = spreadsheet.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    
    if (data.length === 0) return 'No data found';
    
    // Find header indices
    const headers = data[0];
    const placementIdIndex = headers.findIndex(h => normalize(h) === normalize('Placement ID'));
    const placementNameIndex = headers.findIndex(h => normalize(h) === normalize('Placement'));
    
    if (placementIdIndex === -1 || placementNameIndex === -1) return 'Headers not found';
    
    // Search for the placement ID
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[placementIdIndex]).trim() === String(placementId).trim()) {
        return String(row[placementNameIndex] || 'Name not found').trim();
      }
    }
    
    return 'Placement ID not found in recent data';
    
  } catch (error) {
    Logger.log(`Error in LOOKUP_PLACEMENT_NAME: ${error.message}`);
    return `Error: ${error.message}`;
  }
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

function sendNoIssueEmail(config, sheetId, reason) {
  const now = new Date();
  const subjectDate = formatDate(now, "yyyy-MM-dd");
  const subject = `CM360 Audit Complete: No Issues Found (${config.name} - ${subjectDate})`;

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
    to: config.recipients,
    cc: config.cc || '',
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
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx`;
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
function fetchDailyAuditAttachments(config) {
  try {
    Logger.log(`📥 [${config.name}] fetchDailyAuditAttachments started`);

    const label = GmailApp.getUserLabelByName(config.label);
    if (!label) {
      const context = `Missing Gmail label "${config.label}" for config "${config.name}"`;
      return AuditErrorHandler.handleLabelNotFound(context, config);
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
  } catch (error) {
    const context = `fetchDailyAuditAttachments failed for config "${config.name}"`;
    AuditErrorHandler.handleGenericError(error, context, config);
    return null;
  }
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


// === � AUDIT RULES ENGINE ===

// === �📊 MERGE & FLAG LOGIC ===
function executeAudit(config) {
  const logger = AuditUtils.createLogger(config.name);
  const startTime = new Date();
  
  try {
    logger.log('🔍 Audit started');
    
    // Validate configuration
    if (!ConfigManager.validateConfig(config)) {
      throw new Error('Invalid configuration provided');
    }
    
    // Step 1: Process Gmail attachments
    const folderId = FileProcessor.processGmailAttachments(config);
    if (!folderId) {
      return AuditErrorHandler.handleNoFiles(config, logger);
    }
    
    // Step 2: Merge Excel files
    const mergedSheetId = FileProcessor.mergeExcelFiles(
      folderId, 
      config.mergedFolderPath, 
      config.name
    );
    
    // Step 3: Process audit data
    const auditResults = AuditDataProcessor.processSheetData(mergedSheetId, config);
    
    // Step 4: Generate and send reports
    if (auditResults.flaggedRows.length > 0) {
      const htmlReport = ResultFormatter.generateHtmlReport(auditResults, config);
      EmailTemplateEngine.sendFlaggedRowsEmail(config, auditResults, htmlReport);
      
      logger.log(`✅ Audit completed with ${auditResults.flaggedRows.length} flagged rows`);
      return {
        status: 'Completed with flags',
        flaggedCount: auditResults.flaggedRows.length,
        emailSent: true,
        emailTime: AuditUtils.formatDate(startTime, 'yyyy-MM-dd HH:mm:ss'),
        processingTime: Date.now() - startTime.getTime()
      };
    } else {
      EmailTemplateEngine.sendNoIssuesEmail(config, auditResults.spreadsheetUrl);
      
      logger.log('✅ Audit completed with no issues');
      return {
        status: 'Completed (no issues)',
        flaggedCount: 0,
        emailSent: true,
        emailTime: AuditUtils.formatDate(startTime, 'yyyy-MM-dd HH:mm:ss'),
        processingTime: Date.now() - startTime.getTime()
      };
    }
    
  } catch (error) {
    const context = `executeAudit failed for config "${config.name}"`;
    AuditErrorHandler.handleGenericError(error, context, config);
    
    return {
      status: `Error during audit: ${error.message}`,
      flaggedCount: null,
      emailSent: false,
      emailTime: AuditUtils.formatDate(startTime, 'yyyy-MM-dd HH:mm:ss'),
      processingTime: Date.now() - startTime.getTime()
    };
  }
}

// === � EMAIL TEMPLATE ENGINE ===
class EmailTemplateEngine {
  static sendFlaggedRowsEmail(config, auditResults, htmlReport) {
    const { flaggedRows, totalRows, spreadsheetUrl, allData, headers } = auditResults;
    
    // Check email quota before sending
    const quotaRemaining = CacheManager.getEmailQuotaRemaining();
    if (quotaRemaining <= 5) {
      Logger.log(`⚠️ Email quota low (${quotaRemaining}), skipping email for ${config.name}`);
      return false;
    }
    
    const timestamp = AuditUtils.formatDate(new Date(), 'yyyy-MM-dd');
    const subject = `CM360 Audit Alert: ${flaggedRows.length} Issue(s) Found (${config.name} - ${timestamp})`;
    
    // Create Excel attachment with highlighted flags
    let attachments = [];
    try {
      const excelAttachment = EmailTemplateEngine.createFlaggedExcelAttachment(allData, headers, flaggedRows, config.name, timestamp);
      if (excelAttachment) {
        attachments.push(excelAttachment);
      }
    } catch (error) {
      Logger.log(`⚠️ Failed to create Excel attachment: ${error.message}`);
    }
    
    const emailData = {
      to: config.recipients,
      cc: config.cc || '',
      subject: subject,
      htmlBody: htmlReport,
      attachments: attachments
    };
    
    try {
      safeSendEmail(emailData, `${config.name} - Flagged Rows`);
      CacheManager.decrementEmailQuota();
      return true;
    } catch (error) {
      AuditErrorHandler.handleEmailError(error, `Failed to send flagged rows email for ${config.name}`, config);
      return false;
    }
  }
  
  static createFlaggedExcelAttachment(allData, headers, flaggedRows, configName, timestamp) {
    try {
      // Create a new spreadsheet for the attachment
      const tempSpreadsheet = SpreadsheetApp.create(`CM360_Flagged_${configName}_${timestamp}`);
      const sheet = tempSpreadsheet.getSheets()[0];
      sheet.setName('CM360 Audit Results');
      
      // Standard headers with Flags column
      const standardHeaders = [
        'Advertiser', 'Campaign', 'Site (CM360)', 'Placement ID', 'Placement', 
        'Placement Start Date', 'Placement End Date', 'Ad Type', 'Creative', 
        'Placement Pixel Size', 'Creative Pixel Size', 'Date', 'Impressions', 'Clicks', 'Flags'
      ];
      
      if (flaggedRows && flaggedRows.length > 0) {
        // Build Excel data: header + flagged rows only
        const excelData = [standardHeaders];
        
        // Process each flagged row
        flaggedRows.forEach(flaggedResult => {
          const standardRow = [...flaggedResult.rowData]; // Copy the standardized row data
          
          // Add flags to the last column
          const flagMessages = flaggedResult.flags.map(flag => flag.message).join('; ');
          standardRow[14] = flagMessages; // Flags column
          
          // Ensure row has exactly 15 columns
          while (standardRow.length < 15) {
            standardRow.push('');
          }
          if (standardRow.length > 15) {
            standardRow.length = 15;
          }
          
          excelData.push(standardRow);
        });
        
        // Write data to sheet
        if (excelData.length > 1) { // Header + at least one data row
          sheet.getRange(1, 1, excelData.length, standardHeaders.length).setValues(excelData);
          
          // Format header row
          const headerRange = sheet.getRange(1, 1, 1, standardHeaders.length);
          headerRange.setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
          
          // Highlight all flagged data rows (rows 2 onwards) in yellow
          if (excelData.length > 1) {
            const flaggedRange = sheet.getRange(2, 1, excelData.length - 1, standardHeaders.length);
            flaggedRange.setBackground('#FFFF00'); // Yellow highlight
            
            // Bold the flags column for flagged rows
            const flagsRange = sheet.getRange(2, 15, excelData.length - 1, 1);
            flagsRange.setFontWeight('bold').setFontColor('#D32F2F');
          }
          
          // Auto-resize columns
          sheet.autoResizeColumns(1, standardHeaders.length);
          
          Logger.log(`✅ Created Excel attachment with ${excelData.length - 1} flagged rows`);
        }
      } else {
        // No flagged rows - create empty sheet with headers
        sheet.getRange(1, 1, 1, standardHeaders.length).setValues([standardHeaders]);
        const headerRange = sheet.getRange(1, 1, 1, standardHeaders.length);
        headerRange.setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
        
        Logger.log(`✅ Created Excel attachment with headers only (no flagged rows)`);
      }
      
      // Convert to Excel format for attachment
      const url = `https://docs.google.com/spreadsheets/d/${tempSpreadsheet.getId()}/export?format=xlsx`;
      const token = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`
        }
      });
      
      const blob = response.getBlob();
      blob.setName(`CM360_Flagged_${configName}_${timestamp}.xlsx`);
      
      // Clean up temporary file
      DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
      
      return blob;
    } catch (error) {
      Logger.log(`Error creating Excel attachment: ${error.message}`);
      return null;
    }
  }
  
  static sendNoIssuesEmail(config, spreadsheetUrl) {
    const quotaRemaining = CacheManager.getEmailQuotaRemaining();
    if (quotaRemaining <= 5) {
      Logger.log(`⚠️ Email quota low (${quotaRemaining}), skipping no-issues email for ${config.name}`);
      return false;
    }
    
    const timestamp = AuditUtils.formatDate(new Date(), 'yyyy-MM-dd');
    const subject = `✅ CM360 Audit Complete: No Issues Found (${config.name} - ${timestamp})`;
    
    const htmlBody = EmailTemplateEngine.generateNoIssuesTemplate(config, spreadsheetUrl);
    
    const emailData = {
      to: config.recipients,
      cc: config.cc || '',
      subject: subject,
      htmlBody: htmlBody
    };
    
    try {
      safeSendEmail(emailData, `${config.name} - No Issues`);
      CacheManager.decrementEmailQuota();
      return true;
    } catch (error) {
      AuditErrorHandler.handleEmailError(error, `Failed to send no-issues email for ${config.name}`, config);
      return false;
    }
  }
  
  static generateNoIssuesTemplate(config, spreadsheetUrl) {
    return `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #e8f5e8; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
          <h1 style="color: #137333; margin: 0;">✅ CM360 Audit Complete</h1>
          <h2 style="color: #5f6368; margin: 5px 0;">${AuditUtils.escapeHtml(config.name)}</h2>
          <p style="color: #80868b; font-size: 14px; margin: 0;">
            Generated on ${AuditUtils.formatDate(new Date(), 'yyyy-MM-dd HH:mm:ss')}
          </p>
        </div>
        
        <div style="background: #f8f9fa; padding: 15px; border-radius: 6px; margin-bottom: 20px;">
          <h3 style="color: #1a73e8; margin-top: 0;">🎉 Great News!</h3>
          <p>The automated CM360 audit for <strong>${AuditUtils.escapeHtml(config.name)}</strong> 
             has completed successfully with <strong>no issues found</strong>.</p>
          <p>All campaigns, placements, and creatives are operating within expected parameters.</p>
        </div>
        
        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #dadce0; text-align: center; color: #5f6368;">
          <p><a href="${spreadsheetUrl}" style="color: #1a73e8; text-decoration: none;">📊 View Audit Data</a></p>
          <p style="font-size: 12px;"><em>CM360 Audit System - Automated Quality Assurance</em></p>
        </div>
      </div>
    `;
  }
  
  static sendSummaryEmail(results) {
    const timestamp = AuditUtils.formatDate(new Date(), 'yyyy-MM-dd');
    const subject = `📊 CM360 Daily Audit Summary - ${timestamp}`;
    
    const totalConfigs = results.length;
    const successfulAudits = results.filter(r => r.emailSent).length;
    const totalFlags = results.reduce((sum, r) => sum + (r.flaggedRows || 0), 0);
    
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto;">
        <div style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
          <h1 style="color: #1a73e8; margin: 0;">📊 CM360 Daily Audit Summary</h1>
          <p style="color: #80868b; font-size: 14px; margin: 5px 0;">
            ${timestamp} - Platform Solutions Team
          </p>
        </div>
        
        <div style="background: #e8f0fe; padding: 15px; border-radius: 6px; margin-bottom: 20px;">
          <h3>📋 Overview</h3>
          <ul>
            <li><strong>Total Configurations:</strong> ${totalConfigs}</li>
            <li><strong>Successful Audits:</strong> ${successfulAudits}</li>
            <li><strong>Total Flags:</strong> ${totalFlags}</li>
          </ul>
        </div>
        
        <div style="margin-bottom: 20px;">
          <h3>🔍 Audit Results</h3>
          <table style="width: 100%; border-collapse: collapse; border: 1px solid #dadce0;">
            <thead>
              <tr style="background: #f8f9fa;">
                <th style="padding: 12px; text-align: left; border: 1px solid #dadce0;">Configuration</th>
                <th style="padding: 12px; text-align: left; border: 1px solid #dadce0;">Status</th>
                <th style="padding: 12px; text-align: left; border: 1px solid #dadce0;">Flags</th>
                <th style="padding: 12px; text-align: left; border: 1px solid #dadce0;">Email Sent</th>
              </tr>
            </thead>
            <tbody>
              ${results.map(result => `
                <tr>
                  <td style="padding: 8px; border: 1px solid #dadce0;">${AuditUtils.escapeHtml(result.name)}</td>
                  <td style="padding: 8px; border: 1px solid #dadce0;">${AuditUtils.escapeHtml(result.status)}</td>
                  <td style="padding: 8px; border: 1px solid #dadce0;">${result.flaggedRows || 0}</td>
                  <td style="padding: 8px; border: 1px solid #dadce0;">${result.emailSent ? '✅' : '❌'}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
        
        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #dadce0; text-align: center; color: #5f6368;">
          <p style="font-size: 12px;"><em>CM360 Audit System - Automated Quality Assurance</em></p>
        </div>
      </div>
    `;
    
    // Send to summary recipients
    const summaryRecipients = 'nfl-platform-solutions@fanatics.com'; // or from config
    const emailData = {
      to: summaryRecipients,
      subject: subject,
      htmlBody: htmlBody
    };
    
    try {
      safeSendEmail(emailData, 'Daily Audit Summary');
      return true;
    } catch (error) {
      Logger.log(`❌ Failed to send summary email: ${error.message}`);
      return false;
    }
  }
}

// === �📋 EXECUTION & AUDIT FLOW ===
function runDailyAuditByName(configName) {
  if (!checkDriveApiEnabled()) return;
  const config = ConfigManager.getConfigByName(configName);
  if (!config) {
    Logger.log(`❌ Config "${configName}" not found.`);
    return;
  }
  executeAudit(config);
}

function runAuditBatch(configs, isFinal = false) {
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
        emailTime: result.emailTime,
        processingTime: result.processingTime
      });
    } catch (err) {
      results.push({
        name: config.name,
        status: `Error: ${err.message}`,
        flaggedRows: null,
        emailSent: false,
        emailTime: AuditUtils.formatDate(new Date(), 'yyyy-MM-dd HH:mm:ss'),
        processingTime: null
      });
    }
  }

  storeCombinedAuditResults_(results);

  const totalConfigs = ConfigManager.getAuditConfigs().length;
  const cachedResults = getCombinedAuditResults_();

  const completedConfigs = new Set(cachedResults.map(r => r.name)).size;

  Logger.log(`🧮 Completed ${completedConfigs} of ${totalConfigs} configs`);

  if (completedConfigs >= totalConfigs) {
    Logger.log(`📬 All audits complete. Sending summary email...`);
    EmailTemplateEngine.sendSummaryEmail(cachedResults);
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
function emailFlaggedRows(sheetId, emailRows, flaggedRows, config) {
  const recipients = config.recipients;
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
    to: recipients,
    cc: config.cc || '',
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

// Install trigger for auto-populating placement names
function installExclusionsEditTrigger() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    '✅ Auto-Populate Already Active',
    'The auto-populate feature uses a built-in simple trigger (onEdit) that is automatically active.\n\n' +
    'No installation needed! Placement names will auto-fill when you:\n' +
    '• Enter a Config Name and Placement ID\n' +
    '• Edit existing Config Name or Placement ID\n\n' +
    'If auto-populate doesn\'t work, use "Refresh Placement Names" from the menu.',
    ui.ButtonSet.OK
  );
  
  return 'Simple trigger is automatically active';
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
  try {
    validateAuditConfigs();
    checkDriveApiEnabled();

    const ui = SpreadsheetApp.getUi();
  ui.createMenu('CM360 Audit')
    // 🔧 Setup & One-Time Actions
    .addItem('🛠️ Prepare Audit Environment', 'prepareAuditEnvironment')
    .addItem('🔐 Check Authorization', 'checkAuthorizationStatus')
    .addItem('📋 Validate Configs', 'debugValidateAuditConfigs')
    .addItem('📄 Print Config Summary', 'debugPrintConfigSummary')
    .addItem('⚙️ Install Daily Triggers', 'installDailyAuditTriggers')
    .addSeparator()

    // � Exclusions Management
    .addItem('📝 Manage Exclusions Sheet', 'openExclusionsSheet')
    .addItem('🔄 Test Load Exclusions', 'testLoadExclusions')
    .addItem('🔄 Refresh Placement Names', 'refreshPlacementNames')
    .addItem('🔧 Recreate Exclusions Sheet', 'recreateExclusionsSheet')
    .addItem('⚡ Check Auto-Populate Status', 'installExclusionsEditTrigger')
    .addSeparator()

    // �🚀 Manual Run Options
    .addItem('🧪 Run Batch or Config (Manual Test)', 'showBatchTestPicker')
    .addItem('🔎 Run Audit for...', 'showConfigPicker')

    // 📊 Access Tools
    .addItem('📈 Open Dashboard', 'showAuditDashboard')
    .addToUi();
    
    console.log('✅ CM360 Audit menu created successfully');
  } catch (error) {
    console.error('❌ Error creating menu:', error.toString());
    // Menu creation failed, but don't throw error to prevent script failure
  }
}

// Alternative function to create menu manually when onOpen() fails
function createMenuManually() {
  try {
    validateAuditConfigs();
    checkDriveApiEnabled();

    const ui = SpreadsheetApp.getUi();
    ui.createMenu('CM360 Audit')
      // 🔧 Setup & One-Time Actions
      .addItem('🛠️ Prepare Audit Environment', 'prepareAuditEnvironment')
      .addItem('🔐 Check Authorization', 'checkAuthorizationStatus')
      .addItem('📋 Validate Configs', 'debugValidateAuditConfigs')
      .addItem('📄 Print Config Summary', 'debugPrintConfigSummary')
      .addItem('⚙️ Install Daily Triggers', 'installDailyAuditTriggers')
      .addSeparator()

      // 📝 Exclusions Management
      .addItem('📝 Manage Exclusions Sheet', 'openExclusionsSheet')
      .addItem('🔄 Test Load Exclusions', 'testLoadExclusions')
      .addItem('🔄 Refresh Placement Names', 'refreshPlacementNames')
      .addItem('🔧 Recreate Exclusions Sheet', 'recreateExclusionsSheet')
      .addItem('⚡ Check Auto-Populate Status', 'installExclusionsEditTrigger')
      .addSeparator()

      // 🚀 Manual Run Options
      .addItem('🧪 Run Batch or Config (Manual Test)', 'showBatchTestPicker')
      .addItem('🔎 Run Audit for...', 'showConfigPicker')

      // 📊 Access Tools
      .addItem('📈 Open Dashboard', 'showAuditDashboard')
      .addToUi();
    
    console.log('✅ CM360 Audit menu created successfully via manual function');
    SpreadsheetApp.getUi().alert('✅ Success!', 'CM360 Audit menu has been created. Check the menu bar.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('❌ Error creating menu manually:', error.toString());
    throw new Error(`Failed to create menu: ${error.toString()}`);
  }
}

// Diagnostic function that works from script editor
function diagnoseMenuIssue() {
  console.log('🔍 Diagnosing menu creation issue...');
  
  try {
    // Check if we have a spreadsheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    if (sheet) {
      console.log('✅ Active spreadsheet found:', sheet.getName());
      console.log('📧 Spreadsheet ID:', sheet.getId());
      console.log('🔗 Spreadsheet URL:', sheet.getUrl());
    } else {
      console.log('❌ No active spreadsheet found');
    }
  } catch (e) {
    console.log('❌ Cannot access spreadsheet:', e.toString());
  }
  
  console.log('');
  console.log('📋 SOLUTION INSTRUCTIONS:');
  console.log('1. CLOSE this Apps Script tab');
  console.log('2. OPEN your Google Sheets file directly');
  console.log('3. Go to Extensions → Apps Script');
  console.log('4. Select "createMenuManually" and click Run');
  console.log('5. OR simply refresh the Google Sheets page');
  console.log('');
  console.log('💡 The menu creation must happen FROM the spreadsheet context, not from a standalone script editor.');
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

function openExclusionsSheet() {
  try {
    const sheet = getOrCreateExclusionsSheet();
    
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '📝 Exclusions Sheet Ready',
      `The exclusions sheet "${EXCLUSIONS_SHEET_NAME}" is ready for use.\n\n` +
      `✅ Auto-populate is automatically active\n` +
      `✅ Placement names will auto-fill when you enter Config + Placement ID\n\n` +
      `You can now add placement IDs to exclude from specific flag types.\n\n` +
      `The sheet includes sample data and instructions to get you started.`,
      ui.ButtonSet.OK
    );
    
    // Activate the sheet so user can see it
    sheet.activate();
    
  } catch (error) {
    Logger.log(`❌ Error in openExclusionsSheet: ${error.message}`);
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Error Creating Exclusions Sheet',
      `Failed to create or open the exclusions sheet.\n\nError: ${error.message}\n\nPlease check the logs for more details.`,
      ui.ButtonSet.OK
    );
  }
}

function testLoadExclusions() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const exclusions = loadExclusionsFromSheet();
    
    const configCount = Object.keys(exclusions).length;
    let summary = `Loaded exclusions for ${configCount} config(s):\n\n`;
    
    Object.keys(exclusions).forEach(configName => {
      const configExclusions = exclusions[configName];
      summary += `${configName}:\n`;
      
      Object.keys(configExclusions).forEach(flagType => {
        const ids = configExclusions[flagType];
        if (ids.length > 0) {
          summary += `  • ${flagType}: ${ids.join(', ')}\n`;
        }
      });
      summary += '\n';
    });
    
    if (configCount === 0) {
      // Check if it's a structure issue
      const sheet = getOrCreateExclusionsSheet();
      const data = sheet.getDataRange().getValues();
      if (data.length > 0) {
        const headers = data[0];
        summary = `No active exclusions found.\n\nSheet headers found: ${headers.join(', ')}\n\nMake sure:\n• Headers are correct\n• Data has Active = TRUE\n• Config names match exactly`;
      } else {
        summary = 'No exclusions sheet data found.';
      }
    }
    
    ui.alert('🔄 Exclusions Test Results', summary, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error Testing Exclusions', `Failed to load exclusions: ${error.message}`, ui.ButtonSet.OK);
  }
}

function refreshPlacementNames() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getOrCreateExclusionsSheet();
  
  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      ui.alert('No Data', 'No exclusion data found to refresh.', ui.ButtonSet.OK);
      return;
    }
    
    const headers = data[0];
    const configColIndex = headers.indexOf('Config Name');
    const placementIdColIndex = headers.indexOf('Placement ID');
    const placementNameColIndex = headers.indexOf('Placement Name');
    
    if (configColIndex === -1 || placementIdColIndex === -1 || placementNameColIndex === -1) {
      ui.alert('Error', `Required columns not found in exclusions sheet.\nFound headers: ${headers.join(', ')}`, ui.ButtonSet.OK);
      return;
    }
    
    let updated = 0;
    
    // Update placement names for each row
    for (let i = 2; i <= data.length; i++) { // Start from row 2 (skipping header)
      const configName = String(data[i-1][configColIndex] || '').trim();
      const placementId = String(data[i-1][placementIdColIndex] || '').trim();
      
      if (configName && placementId && 
          !configName.includes('INSTRUCTIONS') && 
          !configName.includes('•') && 
          !configName.includes('Config Name:')) {
        
        const placementName = LOOKUP_PLACEMENT_NAME(configName, placementId);
        sheet.getRange(i, placementNameColIndex + 1).setValue(placementName);
        updated++;
        
        // Add a small delay to prevent timeout
        if (updated % 10 === 0) {
          Utilities.sleep(100);
        }
      }
    }
    
    ui.alert(
      '🔄 Placement Names Updated', 
      `Successfully refreshed placement names for ${updated} entries.`, 
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('Error', `Failed to refresh placement names: ${error.message}`, ui.ButtonSet.OK);
  }
}

function recreateExclusionsSheet() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if sheet exists and ask for confirmation
  const existingSheet = spreadsheet.getSheetByName(EXCLUSIONS_SHEET_NAME);
  if (existingSheet) {
    const response = ui.alert(
      'Recreate Exclusions Sheet',
      `The sheet "${EXCLUSIONS_SHEET_NAME}" already exists.\n\nThis will DELETE the existing sheet and create a new one with the updated structure.\n\nDo you want to continue?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // Delete the existing sheet
    spreadsheet.deleteSheet(existingSheet);
  }
  
  // Create the new sheet
  const newSheet = getOrCreateExclusionsSheet();
  
  ui.alert(
    '✅ Sheet Recreated',
    `The exclusions sheet has been recreated with the new structure including:\n\n• Placement Name column (auto-populated)\n• Instructions moved to the right\n• Enhanced protection and formatting`,
    ui.ButtonSet.OK
  );
  
  // Activate the new sheet
  newSheet.activate();
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
    Logger.log(`[${c.name}] Label: ${c.label}, Recipients: ${c.recipients}, Flags: ${JSON.stringify(c.flags)}`);
  });
}

function checkMissingBatchRunners() {
  const batches = getAuditConfigBatches(BATCH_SIZE);
  const messages = [];

  for (let i = 0; i < batches.length; i++) {
    const fnName = `runDailyAuditsBatch${i + 1}`;
  if (typeof globalThis[fnName] !== 'function') {
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
function runPST03Audit() { runDailyAuditByName('PST03'); }
function runNEXT01Audit() { runDailyAuditByName('NEXT01'); }
function runNEXT02Audit() { runDailyAuditByName('NEXT02'); }
function runNEXT03Audit() { runDailyAuditByName('NEXT03'); }
function runSPTM01Audit() { runDailyAuditByName('SPTM01'); }
function runNFL01Audit() { runDailyAuditByName('NFL01'); }
function runENT01Audit() { runDailyAuditByName('ENT01'); }

// Summaries for the Dashboard sidebar
function getAuditConfigSummaries() {
  const batches = getAuditConfigBatches(BATCH_SIZE);
  return batches.map((configs, idx) => ({
    batchLabel: `Batch ${idx + 1}`,
    configs: configs.map(c => ({
      name: c.name,
      label: c.label,
      recipients: c.recipients,
      flags: JSON.stringify(c.flags || {}, null, 2)
    }))
  }));
}

function testDateParsing() {
  // Test the date parsing with Date objects like those from spreadsheets
  const testDate1 = new Date('2025-06-30'); // Past end date
  const testDate2 = new Date('2025-12-31'); // Future end date
  
  Logger.log(`Testing parseDate with Date object: ${testDate1}`);
  const parsed1 = AuditRulesEngine.parseDate(testDate1);
  Logger.log(`Parsed result: ${parsed1}`);
  
  Logger.log(`Testing parseDate with Date object: ${testDate2}`);
  const parsed2 = AuditRulesEngine.parseDate(testDate2);
  Logger.log(`Parsed result: ${parsed2}`);
  
  // Test out-of-flight detection (remember: data represents yesterday's activity)
  const testRowData = {
    'Placement Start Date': new Date('2025-06-01'),
    'Placement End Date': new Date('2025-06-30'),
    'Campaign': 'Wind Creek Bethlehem Test'
  };
  
  Logger.log(`Testing out-of-flight with test data (data represents yesterday's activity)`);
  const result = AuditRulesEngine.checkOutOfFlight(testRowData);
  Logger.log(`Out-of-flight result: ${JSON.stringify(result)}`);
}
*/