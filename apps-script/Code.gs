function configValue_(key, fallback) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  return value === null || value === '' ? fallback : value;
}

function numberConfigValue_(key, fallback) {
  const value = Number(configValue_(key, fallback));
  return Number.isFinite(value) ? value : fallback;
}

const CONFIG = {
  spreadsheetId: configValue_('SPREADSHEET_ID', ''),
  masterSheetName: configValue_('MASTER_SHEET_NAME', 'Master Capture'),
  floorSheetName: configValue_('FLOOR_SHEET_NAME', 'Floor Schedule'),
  photoFolderId: configValue_('PHOTO_FOLDER_ID', ''),
  weeklyOpenOrdersSpreadsheetId: configValue_('WEEKLY_OPEN_ORDERS_SPREADSHEET_ID', ''),
  weeklyOpenOrdersSheetName: configValue_('WEEKLY_OPEN_ORDERS_SHEET_NAME', ''),
  powderMasterSpreadsheetId: configValue_('POWDER_MASTER_SPREADSHEET_ID', ''),
  powderMasterSheetName: configValue_('POWDER_MASTER_SHEET_NAME', 'Powder Master'),
  maskedPartsSpreadsheetId: configValue_('MASKED_PARTS_SPREADSHEET_ID', ''),
  maskedPartsSheetName: configValue_('MASKED_PARTS_SHEET_NAME', 'Sheet1'),
  processedFolderName: configValue_('PROCESSED_FOLDER_NAME', 'Processed'),
  reportFolderName: configValue_('REPORT_FOLDER_NAME', 'Open Order Reports'),
  processedReportFolderName: configValue_('PROCESSED_REPORT_FOLDER_NAME', 'Processed Reports'),
  maskedPartsFolderId: configValue_('MASKED_PARTS_FOLDER_ID', ''),
  formerSchedulesFolderName: configValue_('FORMER_SCHEDULES_FOLDER_NAME', 'Former Schedules'),
  reviewStatus: configValue_('REVIEW_STATUS', 'Needs Review'),
  defaultStatus: configValue_('DEFAULT_STATUS', 'On Hold'),
  digestLimit: numberConfigValue_('DIGEST_LIMIT', 50),
  reportImportLimit: numberConfigValue_('REPORT_IMPORT_LIMIT', 10),
  geminiModel: configValue_('GEMINI_MODEL', 'gemini-2.5-flash'),
  timezone: configValue_('TIMEZONE', 'America/Indiana/Indianapolis')
};

const MASTER_HEADERS = [
  'Schedule Status',
  'Customer',
  'Part Number',
  'Printed Quantity',
  'Powder Code + Color',
  'Date Received',
  'Due Date',
  'Capture Date',
  'Control / Work Order Number',
  'Part Description',
  'Process',
  'Net / E-coat Quantity Received',
  'Film Build Spec',
  'Job / Lot Number',
  'Tooling / Density Text',
  'Special Handling',
  'Conflict Flag',
  'Notes / Unreadable Field Flags',
  'Source Photo',
  'Processed Timestamp',
  'Source Photo ID',
  'OCR Text',
  'Review Status',
  'Digest Batch',
  'Sent to Shop Timestamp',
  'Shop Batch',
  'Masked'
];

const FLOOR_HEADERS = ['Paint Color', 'Hot', 'Customer', 'Part Number', 'Quantity', 'Due Date', 'Control'];
const OPEN_STATUSES = ['On Hold', 'Run Tomorrow', 'Hot List'];

function doGet() {
  return HtmlService
    .createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('Creative Coatings Powder Schedule')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getDashboardState() {
  syncMaskedFlags_();
  const rows = readMasterRows_().filter(isOpenRow_);
  return {
    generatedAt: new Date().toISOString(),
    masterUrl: getSpreadsheet_().getUrl(),
    floorUrl: getSpreadsheet_().getUrl() + '#gid=' + getFloorSheet_().getSheetId(),
    readerStatus: getReaderStatus(),
    counts: buildCounts_(rows),
    rows: rows.slice(0, 300)
  };
}

function updateScheduleStatus(rowNumber, status) {
  const allowed = ['On Hold', 'Run Tomorrow', 'Hot List'];
  if (!allowed.includes(status)) {
    throw new Error('Unsupported status: ' + status);
  }
  const sheet = getMasterSheet_();
  sheet.getRange(Number(rowNumber), 1).setValue(status);
  return getDashboardState();
}

function updateTravelerRow(rowNumber, patch) {
  const columnMap = {
    customer: 2,
    partNumber: 3,
    quantity: 4,
    powder: 5,
    dateReceived: 6,
    dueDate: 7,
    maskedStatus: 27,
    notes: 18
  };
  const sheet = getMasterSheet_();
  Object.keys(patch || {}).forEach(key => {
    if (!columnMap[key]) return;
    sheet.getRange(Number(rowNumber), columnMap[key]).setValue(String(patch[key] || '').trim());
  });
  if (patch && Object.prototype.hasOwnProperty.call(patch, 'partNumber') && !Object.prototype.hasOwnProperty.call(patch, 'maskedStatus')) {
    applyMaskedFlagToRow_(sheet, Number(rowNumber), String(patch.partNumber || '').trim(), null);
  }
  return getDashboardState();
}

function markReviewed(rowNumber) {
  const sheet = getMasterSheet_();
  sheet.getRange(Number(rowNumber), 17).setValue('');
  sheet.getRange(Number(rowNumber), 23).setValue('');
  return getDashboardState();
}

function completeSendToShop() {
  ensureWorkbook_();
  const sheet = getMasterSheet_();
  const rows = readMasterRows_().filter(isOpenRow_);
  const selected = rows.filter(row => row.status === 'Run Tomorrow' || row.status === 'Hot List');
  const shopRows = sortShopRows_(selected);
  const batchId = 'shop-' + Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyyMMdd-HHmmss');
  const timestamp = Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd HH:mm:ss');
  const snapshot = createScheduleSnapshot_(shopRows, batchId, timestamp);

  shopRows.forEach(row => {
    sheet.getRange(row.rowNumber, 1).setValue('Sent to Shop');
    sheet.getRange(row.rowNumber, 25).setValue(timestamp);
    sheet.getRange(row.rowNumber, 26).setValue(batchId);
  });

  const floorSheet = ensureFloorSheet_();
  floorSheet.getRange('I1').setValue('Shop Batch');
  floorSheet.getRange('J1').setValue(batchId);
  floorSheet.getRange('I2').setValue('Sent');
  floorSheet.getRange('J2').setValue(timestamp);

  return {
    batchId,
    timestamp,
    sentCount: shopRows.length,
    shopRows,
    snapshotUrl: snapshot.url,
    floorUrl: getSpreadsheet_().getUrl() + '#gid=' + floorSheet.getSheetId(),
    state: getDashboardState()
  };
}

function previewShopSchedule() {
  ensureWorkbook_();
  const rows = readMasterRows_().filter(isOpenRow_);
  const selected = rows.filter(row => row.status === 'Run Tomorrow' || row.status === 'Hot List');
  const shopRows = sortShopRows_(selected);
  return {
    batchId: 'preview-' + Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyyMMdd-HHmmss'),
    timestamp: Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd HH:mm:ss'),
    sentCount: shopRows.length,
    shopRows,
    floorUrl: getSpreadsheet_().getUrl() + '#gid=' + getFloorSheet_().getSheetId(),
    state: getDashboardState()
  };
}

function digestPhotosNow() {
  ensureWorkbook_();
  return digestNewPhotos_('manual');
}

function importOpenOrderReports() {
  ensureWorkbook_();
  return importOpenOrderReports_('manual');
}

function repairNeedsReview() {
  ensureWorkbook_();
  const rows = readMasterRows_()
    .filter(row => isOpenRow_(row) && row.reviewStatus && row.sourcePhotoId)
    .slice(0, 25);
  const repaired = [];
  const failed = [];

  rows.forEach(row => {
    try {
      const metadata = Drive.Files.get(row.sourcePhotoId);
      const file = makeDriveApiFile_(metadata);
      const readResult = readTraveler_(file);
      updateMasterRowFromParsed_(row.rowNumber, readResult.parsed, readResult.rawText, readResult.reader);
      repaired.push({
        rowNumber: row.rowNumber,
        name: file.getName(),
        partNumber: readResult.parsed.partNumber,
        controlNumber: readResult.parsed.controlNumber,
        reader: readResult.reader
      });
    } catch (error) {
      failed.push({
        rowNumber: row.rowNumber,
        sourcePhotoId: row.sourcePhotoId,
        error: String(error && error.message ? error.message : error)
      });
    }
  });

  return {
    repaired,
    failed,
    state: getDashboardState()
  };
}

function digestNewPhotosScheduled() {
  ensureWorkbook_();
  digestNewPhotos_('scheduled');
}

function importOpenOrderReportsScheduled() {
  ensureWorkbook_();
  importOpenOrderReports_('scheduled');
}

function installTriggers() {
  ensureWorkbook_();
  const handlerNames = ['digestNewPhotosScheduled', 'importOpenOrderReportsScheduled'];
  ScriptApp.getProjectTriggers()
    .filter(trigger => handlerNames.includes(trigger.getHandlerFunction()))
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('digestNewPhotosScheduled')
    .timeBased()
    .everyMinutes(5)
    .create();

  ScriptApp.newTrigger('digestNewPhotosScheduled')
    .timeBased()
    .atHour(15)
    .nearMinute(0)
    .everyDays(1)
    .inTimezone(CONFIG.timezone)
    .create();

  ScriptApp.newTrigger('importOpenOrderReportsScheduled')
    .timeBased()
    .atHour(15)
    .nearMinute(5)
    .everyDays(1)
    .inTimezone(CONFIG.timezone)
    .create();

  return 'Installed 5-minute photo intake, 3 PM photo sweep, and 3:05 PM weekly open-order sync.';
}

function authorizeServices() {
  const ss = getSpreadsheet_();
  const folder = Drive.Files.get(CONFIG.photoFolderId);
  const reader = getReaderStatus();
  UrlFetchApp.fetch('https://generativelanguage.googleapis.com', { muteHttpExceptions: true });
  return {
    spreadsheet: ss.getName(),
    folder: folder.title || folder.name || CONFIG.photoFolderId,
    openOrderSource: CONFIG.weeklyOpenOrdersSpreadsheetId,
    reader: reader.reader,
    model: reader.model
  };
}

function digestNewPhotos_(mode) {
  const processedFolderId = getOrCreateProcessedFolderId_();
  const known = getKnownKeys_();
  const files = listPhotoFiles_();
  const batchId = Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyyMMdd-HHmmss') + '-' + mode;
  const appended = [];
  const skipped = [];
  const failed = [];

  for (let index = 0; index < files.length && appended.length < CONFIG.digestLimit; index += 1) {
    const file = makeDriveApiFile_(files[index]);
    const mimeType = file.getMimeType();
    if (!mimeType || !mimeType.startsWith('image/')) {
      skipped.push({ name: file.getName(), reason: 'not an image' });
      continue;
    }
    if (known.photoIds.has(file.getId())) {
      skipped.push({ name: file.getName(), reason: 'already captured' });
      moveToProcessed_(file, processedFolderId);
      continue;
    }

    try {
      const readResult = readTraveler_(file);
      const parsed = readResult.parsed;
      if (parsed.controlNumber && known.controlNumbers.has(parsed.controlNumber)) {
        skipped.push({ name: file.getName(), reason: 'duplicate control number ' + parsed.controlNumber });
        moveToProcessed_(file, processedFolderId);
        continue;
      }
      appendMasterRow_(parsed, readResult.rawText, file, batchId);
      known.photoIds.add(file.getId());
      if (parsed.controlNumber) known.controlNumbers.add(parsed.controlNumber);
      appended.push({ name: file.getName(), partNumber: parsed.partNumber, controlNumber: parsed.controlNumber, reader: readResult.reader });
      moveToProcessed_(file, processedFolderId);
    } catch (error) {
      failed.push({ name: file.getName(), error: String(error && error.message ? error.message : error) });
    }
  }

  return {
    batchId,
    mode,
    appended,
    skipped,
    failed,
    state: getDashboardState()
  };
}

function importOpenOrderReports_(mode) {
  const batchId = Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyyMMdd-HHmmss') + '-' + mode + '-open-orders';
  const imported = [];
  const skipped = [];
  const failed = [];
  let updated = 0;
  let removed = 0;

  try {
    const weeklySource = getWeeklyOpenOrderSource_();
    const rows = readWeeklyOpenOrderRows_(weeklySource);
    const existingImportedRows = getExistingOpenOrderImportRows_();
    const existingByKey = bucketExistingOpenOrderRows_(existingImportedRows);
    let added = 0;

    rows.forEach(parsed => {
      const bucket = existingByKey[parsed.importKey] || [];
      if (bucket.length) {
        const match = bucket.shift();
        updateOpenOrderImportedRow_(match.rowNumber, match.status, parsed, batchId);
        updated += 1;
        return;
      }
      appendMasterRow_(parsed, JSON.stringify(parsed), weeklySource.file, batchId);
      added += 1;
    });

    const staleRows = Object.keys(existingByKey)
      .reduce((all, key) => all.concat(existingByKey[key]), [])
      .sort((a, b) => b.rowNumber - a.rowNumber);

    if (staleRows.length) {
      deleteMasterRows_(staleRows.map(row => row.rowNumber));
      removed = staleRows.length;
    }

    imported.push({ name: weeklySource.file.getName(), rows: added });
  } catch (error) {
    failed.push({
      name: 'Weekly Open Orders',
      error: String(error && error.message ? error.message : error)
    });
  }

  return {
    batchId,
    mode,
    imported,
    updated,
    removed,
    skipped,
    failed,
    state: getDashboardState()
  };
}

function getWeeklyOpenOrderSource_() {
  const ss = SpreadsheetApp.openById(CONFIG.weeklyOpenOrdersSpreadsheetId);
  const sheet = ss.getSheetByName(CONFIG.weeklyOpenOrdersSheetName) || ss.getSheets()[0];
  if (!sheet) throw new Error('Weekly open-order sheet not found.');
  return {
    ss,
    sheet,
    file: {
      getId: () => ss.getId(),
      getName: () => 'Weekly Open Orders',
      getUrl: () => ss.getUrl(),
      getMimeType: () => 'application/vnd.google-apps.spreadsheet'
    }
  };
}

function readWeeklyOpenOrderRows_(source) {
  const values = source.sheet.getDataRange().getDisplayValues();
  const powderLookup = getPowderMasterLookup_();
  const maskedParts = getMaskedPartLookup_();
  return parseWeeklyOpenOrderValues_(values, source.file, powderLookup, maskedParts);
}

function parseWeeklyOpenOrderValues_(values, file, powderLookup, maskedParts) {
  if (!values || values.length < 2) return [];
  const headers = values[0].map(header => normalizeHeader_(header));
  const customerColumn = findHeaderColumn_(headers, ['name', 'customer']);
  const partColumn = findHeaderColumn_(headers, ['partnumber']);
  const quantityColumn = findHeaderColumn_(headers, ['quantity', 'partsreceivedbydept', 'partsleavingdept']);
  const dueDateColumn = findHeaderColumn_(headers, ['dateshipped', 'duedate', 'shipdate']);
  const receivedDateColumn = findHeaderColumn_(headers, ['recshipdate', 'partreceiveddate', 'datereceived', 'receiveddate']);
  const jobColumn = findHeaderColumn_(headers, ['jobnumber']);
  const poColumn = findHeaderColumn_(headers, ['ponumber', 'po']);
  const descriptionColumn = findHeaderColumn_(headers, ['description', 'partdescription']);
  const processColumn = findHeaderColumn_(headers, ['process', 'process.', 'baseprocessname', 'departementsused', 'departmentsused']);
  const departmentColumn = findHeaderColumn_(headers, ['departmentcaptured', 'department', 'dept']);
  const netQuantityColumn = findHeaderColumn_(headers, ['netquantity', 'partsleavingdept']);
  const powderColumn = findHeaderColumn_(headers, ['powder', 'powdercodecolor', 'powdercode', 'color']);
  const requiredColumns = {
    Customer: customerColumn,
    PartNumber: partColumn,
    Quantity: quantityColumn,
    ReceivedDate: receivedDateColumn
  };
  const missing = Object.keys(requiredColumns).filter(name => requiredColumns[name] < 0);
  if (missing.length) {
    throw new Error('Weekly open-order sheet is missing expected columns: ' + missing.join(', '));
  }

  const rows = [];
  const seenBaseCounts = {};
  const captureDate = Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd');

  for (let index = 1; index < values.length; index += 1) {
    const row = values[index];
    const customer = cleanText_(row[customerColumn]);
    const partNumber = cleanText_(row[partColumn]);
    const printedQuantity = cleanText_(row[quantityColumn]);
    const dueDate = dueDateColumn >= 0 ? cleanText_(row[dueDateColumn]) : '';
    const dateReceived = cleanText_(row[receivedDateColumn]);
    if (!customer && !partNumber && !printedQuantity && !dueDate && !dateReceived) continue;
    if (!customer || !partNumber) continue;

    const powderEntry = getPowderMasterEntry_(partNumber, powderLookup);
    const sourcePowder = powderColumn >= 0 ? cleanText_(row[powderColumn]) : '';
    const powder = normalizePowderCode_(sourcePowder || (powderEntry ? powderEntry.powder : ''));
    const baseKey = buildOpenOrderBaseKey_(customer, partNumber, printedQuantity, dateReceived, dueDate);
    seenBaseCounts[baseKey] = (seenBaseCounts[baseKey] || 0) + 1;
    const process = processColumn >= 0 ? cleanText_(row[processColumn]) : '';
    const department = departmentColumn >= 0 ? cleanText_(row[departmentColumn]) : '';

    rows.push({
      status: CONFIG.defaultStatus,
      customer,
      partNumber,
      printedQuantity,
      powder,
      dateReceived,
      dueDate,
      captureDate,
      controlNumber: jobColumn >= 0 ? cleanText_(row[jobColumn]) : '',
      partDescription: descriptionColumn >= 0 ? cleanText_(row[descriptionColumn]) : '',
      process: [department, process].filter(Boolean).join(' / '),
      netQuantity: netQuantityColumn >= 0 ? cleanText_(row[netQuantityColumn]) : '',
      filmBuildSpec: '',
      jobLotNumber: poColumn >= 0 ? cleanText_(row[poColumn]) : '',
      toolingText: '',
      specialHandling: '',
      conflictFlag: '',
      notes: 'Imported from open-order report: ' + file.getName(),
      reviewStatus: '',
      sourcePhotoUrl: file.getUrl(),
      sourcePhotoId: file.getId(),
      sourceLabel: file.getName(),
      maskedStatus: isMaskedPart_(partNumber, maskedParts) ? 'Masked' : '',
      importKey: baseKey + '|' + seenBaseCounts[baseKey]
    });
  }
  return rows;
}

function findHeaderColumn_(headers, aliases) {
  const normalizedAliases = aliases.map(alias => normalizeHeader_(alias));
  for (const alias of normalizedAliases) {
    const exact = headers.indexOf(alias);
    if (exact >= 0) return exact;
  }
  for (const alias of normalizedAliases) {
    const containing = headers.findIndex(header => header.indexOf(alias) >= 0);
    if (containing >= 0) return containing;
  }
  return -1;
}

function getPowderMasterLookup_() {
  const ss = SpreadsheetApp.openById(CONFIG.powderMasterSpreadsheetId);
  const sheet = ss.getSheetByName(CONFIG.powderMasterSheetName) || ss.getSheets()[0];
  const values = sheet.getDataRange().getDisplayValues();
  const lookup = {};
  if (!values || values.length < 2) return lookup;
  const headers = values[0].map(header => normalizeHeader_(header));
  const partIndex = headers.indexOf('partnumber');
  const powderIndex = headers.indexOf('powder');
  if (partIndex === -1 || powderIndex === -1) {
    throw new Error('Powder Parts Master is missing Part Number or Powder columns.');
  }

  for (let index = 1; index < values.length; index += 1) {
    const row = values[index];
    const partNumber = cleanText_(row[partIndex]);
    const key = normalizePartNumberKey_(partNumber);
    if (!key) continue;
    const powder = normalizePowderCode_(row[powderIndex]);
    if (!lookup[key] || (!lookup[key].powder && powder)) {
      lookup[key] = { partNumber, powder };
    }
  }
  return lookup;
}

function getPowderMasterEntry_(partNumber, lookup) {
  const key = normalizePartNumberKey_(partNumber);
  return key ? (lookup[key] || null) : null;
}

function readOpenOrderReport_(file) {
  const tempSheet = getReadableOpenOrderSheet_(file);
  try {
    const ss = SpreadsheetApp.openById(tempSheet.id);
    const sheet = ss.getSheets()[0];
    const values = sheet.getDataRange().getDisplayValues();
    return parseOpenOrderReportValues_(values, file);
  } finally {
    if (tempSheet.cleanup) {
      Drive.Files.trash(tempSheet.id);
    }
  }
}

function getReadableOpenOrderSheet_(file) {
  const mimeType = String(file.getMimeType() || '').toLowerCase();
  if (mimeType === 'application/vnd.google-apps.spreadsheet') {
    return { id: file.getId(), cleanup: false };
  }

  throw new Error(
    'Unsupported report format for "' + file.getName() + '". Please convert it to a Google Sheet and place that Google Sheet in Open Order Reports.'
  );
}

function parseOpenOrderReportValues_(values, file) {
  if (!values || values.length < 2) return [];
  const maskedParts = getMaskedPartLookup_();
  const headers = values[0].map(header => normalizeHeader_(header));
  const required = ['jobnumber', 'name', 'dateshipped', 'recshipdate', 'baseprocessname', 'text57', 'text59'];
  const missing = required.filter(header => headers.indexOf(header) === -1);
  if (missing.length) {
    throw new Error('Open-order report is missing expected columns: ' + missing.join(', '));
  }

  const column = name => headers.indexOf(normalizeHeader_(name));
  const rows = [];
  for (let index = 1; index < values.length; index += 1) {
    const row = values[index];
    const controlNumber = cleanText_(row[column('jobnumber')]);
    const customer = cleanText_(row[column('name')]);
    const partNumber = cleanText_(row[column('text57')]);
    if (!controlNumber && !customer && !partNumber) continue;
    if (!customer || !partNumber) continue;

    const process = cleanText_(row[column('baseprocessname')]);
    const priority = cleanText_(row[column('priority')]);
    const poNumber = cleanText_(row[column('ponumber')]);
    const printedQuantity = cleanText_(row[column('text59')]);
    const netQuantity = column('text72') >= 0 ? cleanText_(row[column('text72')]) : '';
    rows.push({
      status: CONFIG.defaultStatus,
      customer,
      partNumber,
      printedQuantity,
      powder: '',
      dateReceived: cleanText_(row[column('dateshipped')]),
      dueDate: cleanText_(row[column('recshipdate')]),
      captureDate: Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd'),
      controlNumber,
      partDescription: '',
      process,
      netQuantity,
      filmBuildSpec: '',
      jobLotNumber: poNumber,
      toolingText: '',
      specialHandling: priority && priority !== 'Normal' ? priority : '',
      conflictFlag: '',
      notes: [
        'Imported from open-order report: ' + file.getName(),
        priority ? 'Priority: ' + priority : '',
        'Report process: ' + process,
        netQuantity ? 'Text72 quantity: ' + netQuantity : ''
      ].filter(Boolean).join(' | '),
      reviewStatus: '',
      sourcePhotoUrl: file.getUrl(),
      sourcePhotoId: file.getId(),
      maskedStatus: isMaskedPart_(partNumber, maskedParts) ? 'Masked' : ''
    });
  }
  return rows;
}

function readTraveler_(file) {
  if (getGeminiApiKey_()) {
    try {
      const vision = runGeminiTravelerVision_(file);
      return {
        reader: 'Gemini Vision',
        rawText: JSON.stringify(vision, null, 2),
        parsed: normalizeVisionTraveler_(vision, file)
      };
    } catch (error) {
      const ocrText = runDriveOcr_(file);
      const parsed = parseTravelerText_(ocrText, file);
      parsed.conflictFlag = parsed.conflictFlag || 'Vision fallback';
      parsed.notes = [parsed.notes, 'Gemini vision failed; used Drive OCR fallback: ' + String(error.message || error)]
        .filter(Boolean)
        .join(' | ');
      return { reader: 'Drive OCR fallback', rawText: ocrText, parsed };
    }
  }

  const ocrText = runDriveOcr_(file);
  return {
    reader: 'Drive OCR',
    rawText: ocrText,
    parsed: parseTravelerText_(ocrText, file)
  };
}

function runGeminiTravelerVision_(file) {
  const apiKey = getGeminiApiKey_();
  const blob = file.getBlob();
  const payload = {
    contents: [{
      parts: [
        { text: buildTravelerVisionPrompt_() },
        {
          inline_data: {
            mime_type: blob.getContentType() || file.getMimeType() || 'image/jpeg',
            data: Utilities.base64Encode(blob.getBytes())
          }
        }
      ]
    }],
    generationConfig: {
      temperature: 0,
      responseMimeType: 'application/json',
      responseJsonSchema: getTravelerSchema_()
    }
  };
  const response = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/' + CONFIG.geminiModel + ':generateContent?key=' + encodeURIComponent(apiKey),
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );
  const status = response.getResponseCode();
  const body = response.getContentText();
  if (status < 200 || status >= 300) {
    throw new Error('Gemini API ' + status + ': ' + body.slice(0, 500));
  }
  const json = JSON.parse(body);
  const text = json.candidates &&
    json.candidates[0] &&
    json.candidates[0].content &&
    json.candidates[0].content.parts &&
    json.candidates[0].content.parts.map(part => part.text || '').join('');
  if (!text) throw new Error('Gemini returned no text payload.');
  return JSON.parse(text);
}

function buildTravelerVisionPrompt_() {
  return [
    'You are extracting production traveler data for Creative Coatings powder scheduling.',
    'Read the whole image carefully. Return only the requested JSON fields.',
    'Capture the traveler truth as written. Do not infer a value if it is not visible.',
    'Paint color means powder color. Preserve powder/color codes such as YN1117, PCS99109, MED GL BLACK, GL BLACK.',
    'Control/work order is usually a five digit number around 50xxx/51xxx.',
    'Printed quantity is the traveler quantity, not tooling counts, PC/LB, PC/HR, dates, or batch numbers.',
    'Tooling text should capture setup/rack/density/process lines verbatim when visible.',
    'Special handling should capture pack out, foam, mask, protect, prehang, clearance, and similar instructions.',
    'If a field is unreadable, return an empty string and list it in unreadableFields.'
  ].join('\n');
}

function getTravelerSchema_() {
  return {
    type: 'object',
    properties: {
      customer: { type: 'string' },
      partNumber: { type: 'string' },
      controlNumber: { type: 'string' },
      partDescription: { type: 'string' },
      process: { type: 'string' },
      printedQuantity: { type: 'string' },
      netQuantity: { type: 'string' },
      powder: { type: 'string' },
      filmBuildSpec: { type: 'string' },
      dateReceived: { type: 'string' },
      dueDate: { type: 'string' },
      jobLotNumber: { type: 'string' },
      toolingText: { type: 'string' },
      specialHandling: { type: 'string' },
      unreadableFields: {
        type: 'array',
        items: { type: 'string' }
      },
      notes: { type: 'string' }
    },
    required: [
      'customer',
      'partNumber',
      'controlNumber',
      'partDescription',
      'process',
      'printedQuantity',
      'netQuantity',
      'powder',
      'filmBuildSpec',
      'dateReceived',
      'dueDate',
      'jobLotNumber',
      'toolingText',
      'specialHandling',
      'unreadableFields',
      'notes'
    ]
  };
}

function normalizeVisionTraveler_(vision, file) {
  const unreadable = Array.isArray(vision.unreadableFields) ? vision.unreadableFields.filter(Boolean) : [];
  ['customer', 'partNumber', 'printedQuantity', 'powder', 'controlNumber'].forEach(field => {
    if (!String(vision[field] || '').trim()) unreadable.push(field);
  });
  return {
    status: CONFIG.defaultStatus,
    customer: cleanText_(vision.customer),
    partNumber: cleanText_(vision.partNumber),
    printedQuantity: cleanText_(vision.printedQuantity),
    powder: normalizePowderCode_(vision.powder),
    dateReceived: cleanText_(vision.dateReceived),
    dueDate: cleanText_(vision.dueDate),
    captureDate: Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd'),
    controlNumber: cleanText_(vision.controlNumber),
    partDescription: cleanText_(vision.partDescription),
    process: cleanText_(vision.process),
    netQuantity: cleanText_(vision.netQuantity),
    filmBuildSpec: cleanText_(vision.filmBuildSpec),
    jobLotNumber: cleanText_(vision.jobLotNumber),
    toolingText: cleanText_(vision.toolingText),
    specialHandling: cleanText_(vision.specialHandling),
    conflictFlag: unreadable.length ? 'Review vision' : '',
    notes: [vision.notes, unreadable.length ? 'Unreadable or low-confidence fields: ' + unique_(unreadable).join(', ') : '']
      .filter(Boolean)
      .join(' | '),
    reviewStatus: unreadable.length ? CONFIG.reviewStatus : '',
    sourcePhotoUrl: file.getUrl(),
    sourcePhotoId: file.getId(),
    maskedStatus: ''
  };
}

function cleanText_(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function unique_(values) {
  return [...new Set(values.map(value => String(value || '').trim()).filter(Boolean))];
}

function getGeminiApiKey_() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
}

function setGeminiApiKey(apiKey) {
  const value = String(apiKey || '').trim();
  if (!value) throw new Error('Missing Gemini API key.');
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', value);
  return 'Gemini vision reader is configured.';
}

function getReaderStatus() {
  return {
    reader: getGeminiApiKey_() ? 'Gemini Vision' : 'Drive OCR fallback',
    model: CONFIG.geminiModel
  };
}

function runDriveOcr_(file) {
  const tempDoc = Drive.Files.copy(
    {
      title: 'OCR - ' + file.getName(),
      mimeType: MimeType.GOOGLE_DOCS
    },
    file.getId(),
    {
      ocr: true,
      ocrLanguage: 'en',
      convert: true
    }
  );

  try {
    const doc = DocumentApp.openById(tempDoc.id);
    return doc.getBody().getText().trim();
  } finally {
    Drive.Files.trash(tempDoc.id);
  }
}

function parseTravelerText_(text, file) {
  const lines = normalizeOcrLines_(text);
  const compact = lines.join('\n');
  const firstNumbers = compact.match(/\b[A-Z0-9][A-Z0-9-]{4,}\b/g) || [];
  const dates = compact.match(/\b(?:0?[1-9]|1[0-2])[\/-](?:0?[1-9]|[12]\d|3[01])[\/-](?:20)?\d{2}\b/g) || [];

  const customer = findCustomer_(lines);
  const process = /e\s*coat|ecoat/i.test(compact) && /powder/i.test(compact)
    ? 'Ecoat, Powder Coat'
    : (/powder/i.test(compact) ? 'Powder Coat' : '');
  const powder = normalizePowderCode_(findPowder_(compact));
  const quantity = findQuantity_(lines);
  const controlNumber = findControlNumber_(lines, firstNumbers);
  const partNumber = findPartNumber_(lines, firstNumbers, controlNumber);
  const toolingText = findToolingText_(lines);
  const specialHandling = findSpecialHandling_(lines);

  const unreadable = [];
  if (!customer) unreadable.push('Customer');
  if (!partNumber) unreadable.push('Part Number');
  if (!quantity) unreadable.push('Printed Quantity');
  if (!powder) unreadable.push('Powder Code + Color');
  if (!controlNumber) unreadable.push('Control / Work Order Number');

  return {
    status: CONFIG.defaultStatus,
    customer,
    partNumber,
    printedQuantity: quantity,
    powder,
    dateReceived: dates[0] || '',
    dueDate: dates[1] || '',
    captureDate: Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd'),
    controlNumber,
    partDescription: findDescription_(lines, partNumber),
    process,
    netQuantity: '',
    filmBuildSpec: findFilmSpec_(compact),
    jobLotNumber: findJobLot_(lines),
    toolingText,
    specialHandling,
    conflictFlag: unreadable.length ? 'Review OCR' : '',
    notes: unreadable.length ? 'Unreadable or low-confidence fields: ' + unreadable.join(', ') : '',
    reviewStatus: unreadable.length ? CONFIG.reviewStatus : '',
    sourcePhotoUrl: file.getUrl(),
    sourcePhotoId: file.getId(),
    maskedStatus: ''
  };
}

function appendMasterRow_(parsed, ocrText, file, batchId) {
  const sheet = getMasterSheet_();
  const sourceLabel = parsed.sourceLabel || file.getName();
  const row = [
    parsed.status,
    parsed.customer,
    parsed.partNumber,
    parsed.printedQuantity,
    normalizePowderCode_(parsed.powder),
    parsed.dateReceived,
    parsed.dueDate,
    parsed.captureDate,
    parsed.controlNumber,
    parsed.partDescription,
    parsed.process,
    parsed.netQuantity,
    parsed.filmBuildSpec,
    parsed.jobLotNumber,
    parsed.toolingText,
    parsed.specialHandling,
    parsed.conflictFlag,
    parsed.notes,
    '=HYPERLINK("' + parsed.sourcePhotoUrl + '","' + String(sourceLabel || '').replace(/"/g, '""') + '")',
    Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd HH:mm:ss'),
    parsed.sourcePhotoId,
    ocrText,
    parsed.reviewStatus,
    batchId,
    '',
    '',
    parsed.maskedStatus || ''
  ];
  sheet.appendRow(row);
}

function updateOpenOrderImportedRow_(rowNumber, preservedStatus, parsed, batchId) {
  const sheet = getMasterSheet_();
  const sourceLabel = parsed.sourceLabel || 'Weekly Open Orders';
  const row = [[
    preservedStatus || CONFIG.defaultStatus,
    parsed.customer,
    parsed.partNumber,
    parsed.printedQuantity,
    normalizePowderCode_(parsed.powder),
    parsed.dateReceived,
    parsed.dueDate,
    parsed.captureDate,
    parsed.controlNumber,
    parsed.partDescription,
    parsed.process,
    parsed.netQuantity,
    parsed.filmBuildSpec,
    parsed.jobLotNumber,
    parsed.toolingText,
    parsed.specialHandling,
    parsed.conflictFlag,
    parsed.notes,
    '=HYPERLINK("' + parsed.sourcePhotoUrl + '","' + String(sourceLabel || '').replace(/"/g, '""') + '")',
    Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd HH:mm:ss'),
    parsed.sourcePhotoId,
    JSON.stringify(parsed),
    parsed.reviewStatus,
    batchId,
    '',
    '',
    parsed.maskedStatus || ''
  ]];
  sheet.getRange(rowNumber, 1, 1, MASTER_HEADERS.length).setValues(row);
}

function updateMasterRowFromParsed_(rowNumber, parsed, rawText, reader) {
  const sheet = getMasterSheet_();
  const values = [
    parsed.customer,
    parsed.partNumber,
    parsed.printedQuantity,
    normalizePowderCode_(parsed.powder),
    parsed.dateReceived,
    parsed.dueDate
  ];
  sheet.getRange(rowNumber, 2, 1, values.length).setValues([values]);
  sheet.getRange(rowNumber, 9, 1, 10).setValues([[
    parsed.controlNumber,
    parsed.partDescription,
    parsed.process,
    parsed.netQuantity,
    parsed.filmBuildSpec,
    parsed.jobLotNumber,
    parsed.toolingText,
    parsed.specialHandling,
    parsed.conflictFlag,
    [parsed.notes, 'Re-read with ' + reader + ' at ' + Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd HH:mm:ss')]
      .filter(Boolean)
      .join(' | ')
  ]]);
  sheet.getRange(rowNumber, 22).setValue(rawText);
  sheet.getRange(rowNumber, 23).setValue(parsed.reviewStatus);
}

function ensureWorkbook_() {
  const master = getMasterSheet_();
  ensureHeaders_(master, MASTER_HEADERS);
  ensureStatusValidation_(master);
  ensureFloorSheet_();
}

function ensureHeaders_(sheet, headers) {
  ensureSheetSize_(sheet, 1, headers.length);
  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (existing.join('|') !== headers.join('|')) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function ensureStatusValidation_(sheet) {
  ensureSheetSize_(sheet, 1000, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['On Hold', 'Run Tomorrow', 'Hot List', 'Sent to Shop'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
}

function ensureFloorSheet_() {
  const sheet = getFloorSheet_();
  ensureSheetSize_(sheet, 1000, 10);
  ensureHeaders_(sheet, FLOOR_HEADERS);
  sheet.getRange(2, 1).setFormula(
    '=SORT(FILTER({ARRAYFORMULA(IFERROR(REGEXEXTRACT(\'' + CONFIG.masterSheetName + '\'!E2:E,"(?i)\\\\b(?:YN|PCS)\\\\d+\\\\b"),\'' +
      CONFIG.masterSheetName + '\'!E2:E)),IF(\'' +
      CONFIG.masterSheetName + '\'!A2:A="Hot List","HOT",""),\'' +
      CONFIG.masterSheetName + '\'!B2:B,\'' +
      CONFIG.masterSheetName + '\'!C2:C,\'' +
      CONFIG.masterSheetName + '\'!D2:D,\'' +
      CONFIG.masterSheetName + '\'!G2:G,\'' +
      CONFIG.masterSheetName + '\'!I2:I},((\'' +
      CONFIG.masterSheetName + '\'!A2:A="Run Tomorrow")+(\''
      + CONFIG.masterSheetName + '\'!A2:A="Hot List"))*(\'' +
      CONFIG.masterSheetName + '\'!B2:B<>"")),1,TRUE,2,FALSE,3,TRUE,4,TRUE)'
  );
  return sheet;
}

function getFloorSheet_() {
  const ss = getSpreadsheet_();
  return ss.getSheetByName(CONFIG.floorSheetName) || ss.insertSheet(CONFIG.floorSheetName);
}

function createScheduleSnapshot_(shopRows, batchId, timestamp) {
  const dateStamp = Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd');
  const snapshotName = 'Floor Schedule ' + dateStamp + ' ' + batchId;
  const snapshot = SpreadsheetApp.create(snapshotName);
  const sheet = snapshot.getSheets()[0];
  sheet.setName('Floor Schedule');

  const values = [
    ['Creative Coatings Floor Schedule', '', '', '', '', '', ''],
    ['Schedule Date', dateStamp, '', '', '', '', ''],
    ['Shop Batch', batchId, '', '', '', '', ''],
    ['Sent Timestamp', timestamp, '', '', '', '', ''],
    ['', '', '', '', '', '', ''],
    FLOOR_HEADERS
  ];
  shopRows.forEach(row => {
    values.push([
      row.powder,
      row.status === 'Hot List' ? 'HOT' : '',
      row.customer,
      row.partNumber,
      row.quantity,
      row.dueDate,
      row.controlNumber
    ]);
  });

  sheet.getRange(1, 1, values.length, FLOOR_HEADERS.length).setValues(values);
  sheet.getRange(1, 1, 1, FLOOR_HEADERS.length).merge();
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(16);
  sheet.getRange(6, 1, 1, FLOOR_HEADERS.length).setFontWeight('bold').setBackground('#dceadd');
  sheet.setFrozenRows(6);
  sheet.autoResizeColumns(1, FLOOR_HEADERS.length);

  const folderId = getOrCreateFormerSchedulesFolderId_();
  moveDriveFileToFolder_(snapshot.getId(), folderId);
  return {
    id: snapshot.getId(),
    url: snapshot.getUrl()
  };
}

function getOrCreateFormerSchedulesFolderId_() {
  const parentId = getScheduleParentFolderId_();
  const query = [
    '\'' + parentId + '\' in parents',
    'title = \'' + escapeDriveQuery_(CONFIG.formerSchedulesFolderName) + '\'',
    'mimeType = \'application/vnd.google-apps.folder\'',
    'trashed = false'
  ].join(' and ');
  const existing = Drive.Files.list({ q: query, maxResults: 1 }).items || [];
  if (existing.length) return existing[0].id;

  const folder = Drive.Files.insert({
    title: CONFIG.formerSchedulesFolderName,
    mimeType: 'application/vnd.google-apps.folder',
    parents: [{ id: parentId }]
  });
  return folder.id;
}

function getOrCreateOpenOrderReportsFolderId_() {
  const parentId = getScheduleParentFolderId_();
  return getOrCreateChildFolder_(parentId, CONFIG.reportFolderName);
}

function getOrCreateProcessedReportFolderId_(reportFolderId) {
  return getOrCreateChildFolder_(reportFolderId, CONFIG.processedReportFolderName);
}

function getOrCreateChildFolder_(parentId, folderName) {
  const query = [
    '\'' + parentId + '\' in parents',
    'title = \'' + escapeDriveQuery_(folderName) + '\'',
    'mimeType = \'application/vnd.google-apps.folder\'',
    'trashed = false'
  ].join(' and ');
  const existing = Drive.Files.list({ q: query, maxResults: 1 }).items || [];
  if (existing.length) return existing[0].id;

  const folder = Drive.Files.insert({
    title: folderName,
    mimeType: 'application/vnd.google-apps.folder',
    parents: [{ id: parentId }]
  });
  return folder.id;
}

function getScheduleParentFolderId_() {
  const photoFolder = Drive.Files.get(CONFIG.photoFolderId);
  const parents = photoFolder.parents || [];
  return parents.length ? parents[0].id : 'root';
}

function moveDriveFileToFolder_(fileId, folderId) {
  const metadata = Drive.Files.get(fileId);
  const removeParents = (metadata.parents || []).map(parent => parent.id).filter(Boolean).join(',');
  const options = { addParents: folderId };
  if (removeParents) options.removeParents = removeParents;
  Drive.Files.patch({}, fileId, options);
}

function moveDriveFile_(fileId, fromFolderId, toFolderId) {
  Drive.Files.patch(
    {},
    fileId,
    {
      addParents: toFolderId,
      removeParents: fromFolderId
    }
  );
}

function sortShopRows_(rows) {
  return rows.slice().sort((a, b) => {
    const colorCompare = compareText_(a.powder, b.powder);
    if (colorCompare !== 0) return colorCompare;
    const hotCompare = Number(b.status === 'Hot List') - Number(a.status === 'Hot List');
    if (hotCompare !== 0) return hotCompare;
    const customerCompare = compareText_(a.customer, b.customer);
    if (customerCompare !== 0) return customerCompare;
    return compareText_(a.partNumber, b.partNumber);
  });
}

function compareText_(a, b) {
  return String(a || '').localeCompare(String(b || ''), undefined, { sensitivity: 'base', numeric: true });
}

function ensureSheetSize_(sheet, minRows, minColumns) {
  if (sheet.getMaxRows() < minRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), minRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < minColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), minColumns - sheet.getMaxColumns());
  }
}

function readMasterRows_() {
  const sheet = getMasterSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, MASTER_HEADERS.length).getDisplayValues();
  return values
    .map((row, index) => rowToObject_(row, index + 2))
    .filter(row => row.customer || row.partNumber || row.sourcePhotoId);
}

function rowToObject_(row, rowNumber) {
  return {
    rowNumber,
    status: row[0],
    customer: row[1],
    partNumber: row[2],
    quantity: row[3],
    powder: normalizePowderCode_(row[4]),
    dateReceived: row[5],
    dueDate: row[6],
    captureDate: row[7],
    controlNumber: row[8],
    description: row[9],
    process: row[10],
    conflictFlag: row[16],
    notes: row[17],
    sourcePhoto: row[18],
    processedAt: row[19],
    sourcePhotoId: row[20],
    sourcePhotoUrl: row[20] ? 'https://drive.google.com/file/d/' + row[20] + '/view' : '',
    reviewStatus: row[22],
    batchId: row[23],
    sentToShopAt: row[24],
    shopBatch: row[25],
    maskedStatus: row[26] || ''
  };
}

function buildCounts_(rows) {
  return rows.reduce((counts, row) => {
    counts.total += 1;
    if (row.status === 'Hot List') counts.hot += 1;
    if (row.status === 'Run Tomorrow') counts.tomorrow += 1;
    if (row.status === 'On Hold') counts.hold += 1;
    if (row.reviewStatus) counts.review += 1;
    return counts;
  }, { total: 0, hot: 0, tomorrow: 0, hold: 0, review: 0 });
}

function isOpenRow_(row) {
  return OPEN_STATUSES.includes(row.status || '');
}

function getKnownKeys_() {
  const sheet = getMasterSheet_();
  const lastRow = sheet.getLastRow();
  const keys = {
    photoIds: new Set(),
    controlNumbers: new Set(),
    controlPartKeys: new Set()
  };
  if (lastRow < 2) return keys;
  sheet.getRange(2, 9, lastRow - 1, 13).getValues()
    .forEach(row => {
      const controlNumber = row[0];
      const sourcePhotoId = row[12];
      if (controlNumber) keys.controlNumbers.add(String(controlNumber));
      if (sourcePhotoId) keys.photoIds.add(String(sourcePhotoId));
    });
  sheet.getRange(2, 3, lastRow - 1, 7).getValues()
    .forEach(row => {
      const partNumber = row[0];
      const controlNumber = row[6];
      const key = buildControlPartKey_(controlNumber, partNumber);
      if (key) keys.controlPartKeys.add(key);
    });
  return keys;
}

function getExistingOpenOrderImportRows_() {
  return readMasterRows_()
    .filter(row => isOpenRow_(row) && isOpenOrderImportRow_(row))
    .sort((a, b) => a.rowNumber - b.rowNumber);
}

function isOpenOrderImportRow_(row) {
  const sourceId = String(row.sourcePhotoId || '');
  const notes = String(row.notes || '');
  const sourceLabel = String(row.sourcePhoto || '');
  return sourceId === CONFIG.weeklyOpenOrdersSpreadsheetId
    || notes.indexOf('Imported from open-order report:') >= 0
    || /open order/i.test(sourceLabel);
}

function bucketExistingOpenOrderRows_(rows) {
  const baseCounts = {};
  const buckets = {};
  rows.forEach(row => {
    const baseKey = buildOpenOrderBaseKey_(row.customer, row.partNumber, row.quantity, row.dateReceived, row.dueDate);
    baseCounts[baseKey] = (baseCounts[baseKey] || 0) + 1;
    const importKey = baseKey + '|' + baseCounts[baseKey];
    if (!buckets[importKey]) buckets[importKey] = [];
    buckets[importKey].push(row);
  });
  return buckets;
}

function buildOpenOrderBaseKey_(customer, partNumber, quantity, dateReceived, dueDate) {
  return [
    normalizePartNumberKey_(customer),
    normalizePartNumberKey_(partNumber),
    normalizePartNumberKey_(quantity),
    normalizePartNumberKey_(dateReceived),
    normalizePartNumberKey_(dueDate)
  ].join('|');
}

function deleteMasterRows_(rowNumbers) {
  const sheet = getMasterSheet_();
  rowNumbers
    .slice()
    .sort((a, b) => b - a)
    .forEach(rowNumber => sheet.deleteRow(rowNumber));
}

function syncMaskedFlags_() {
  const sheet = getMasterSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const maskedParts = getMaskedPartLookup_();
  const values = sheet.getRange(2, 3, lastRow - 1, 25).getDisplayValues();
  const updates = [];
  values.forEach((row, index) => {
    const partNumber = row[0];
    const currentMasked = row[24];
    const nextMasked = isMaskedPart_(partNumber, maskedParts) ? 'Masked' : '';
    if (currentMasked !== nextMasked) {
      updates.push({ rowNumber: index + 2, value: nextMasked });
    }
  });
  updates.forEach(update => sheet.getRange(update.rowNumber, 27).setValue(update.value));
}

function applyMaskedFlagToRow_(sheet, rowNumber, partNumber, maskedParts) {
  const lookup = maskedParts || getMaskedPartLookup_();
  const value = isMaskedPart_(partNumber, lookup) ? 'Masked' : '';
  sheet.getRange(rowNumber, 27).setValue(value);
}

function getMaskedPartLookup_() {
  const spreadsheetId = getMaskedPartsSpreadsheetId_();
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(CONFIG.maskedPartsSheetName) || ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  const lookup = new Set();
  if (lastRow < 2) return lookup;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(header => normalizeHeader_(header));
  const partColumn = headers.indexOf('partnumber');
  if (partColumn === -1) {
    throw new Error('Masked Parts Database is missing a Part Number column.');
  }
  sheet.getRange(2, partColumn + 1, lastRow - 1, 1).getDisplayValues().forEach(row => {
    const normalized = normalizePartNumberKey_(row[0]);
    if (normalized) lookup.add(normalized);
  });
  return lookup;
}

function getMaskedPartsSpreadsheetId_() {
  if (CONFIG.maskedPartsSpreadsheetId) return CONFIG.maskedPartsSpreadsheetId;
  const query = [
    '\'' + CONFIG.maskedPartsFolderId + '\' in parents',
    'mimeType = \'application/vnd.google-apps.spreadsheet\'',
    'trashed = false'
  ].join(' and ');
  const files = Drive.Files.list({ q: query, maxResults: 1, orderBy: 'modifiedDate desc' }).items || [];
  if (!files.length) {
    throw new Error('No Google Sheet found in the Masked Parts folder.');
  }
  return files[0].id;
}

function normalizePartNumberKey_(value) {
  return String(value || '').toUpperCase().replace(/\s+/g, '').trim();
}

function isMaskedPart_(partNumber, maskedParts) {
  const key = normalizePartNumberKey_(partNumber);
  return Boolean(key) && maskedParts.has(key);
}

function getOrCreateProcessedFolderId_() {
  const query = [
    '\'' + CONFIG.photoFolderId + '\' in parents',
    'title = \'' + escapeDriveQuery_(CONFIG.processedFolderName) + '\'',
    'mimeType = \'application/vnd.google-apps.folder\'',
    'trashed = false'
  ].join(' and ');
  const existing = Drive.Files.list({ q: query, maxResults: 1 }).items || [];
  if (existing.length) return existing[0].id;

  const folder = Drive.Files.insert({
    title: CONFIG.processedFolderName,
    mimeType: 'application/vnd.google-apps.folder',
    parents: [{ id: CONFIG.photoFolderId }]
  });
  return folder.id;
}

function listPhotoFiles_() {
  const query = [
    '\'' + CONFIG.photoFolderId + '\' in parents',
    'trashed = false'
  ].join(' and ');
  const result = Drive.Files.list({
    q: query,
    maxResults: CONFIG.digestLimit,
    orderBy: 'createdDate'
  });
  return result.items || [];
}

function listOpenOrderReportFiles_(reportFolderId) {
  const query = [
    '\'' + reportFolderId + '\' in parents',
    'trashed = false',
    'mimeType != \'application/vnd.google-apps.folder\''
  ].join(' and ');
  const result = Drive.Files.list({
    q: query,
    maxResults: CONFIG.reportImportLimit,
    orderBy: 'createdDate'
  });
  return result.items || [];
}

function makeDriveApiFile_(metadata) {
  const id = metadata.id;
  const title = metadata.title || metadata.name || id;
  const mimeType = metadata.mimeType || '';
  return {
    getId: () => id,
    getName: () => title,
    getMimeType: () => mimeType,
    getUrl: () => metadata.alternateLink || metadata.webViewLink || ('https://drive.google.com/file/d/' + id + '/view'),
    getBlob: () => fetchDriveFileBlob_(id, title, mimeType)
  };
}

function normalizeHeader_(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function buildControlPartKey_(controlNumber, partNumber) {
  const control = String(controlNumber || '').trim().toUpperCase();
  const part = String(partNumber || '').trim().toUpperCase();
  return control && part ? control + '|' + part : '';
}

function fetchDriveFileBlob_(fileId, name, mimeType) {
  const response = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) + '?alt=media',
    {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    }
  );
  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error('Drive media download ' + status + ': ' + response.getContentText().slice(0, 300));
  }
  return response.getBlob().setName(name).setContentType(mimeType || response.getBlob().getContentType());
}

function moveToProcessed_(file, processedFolderId) {
  Drive.Files.patch(
    {},
    file.getId(),
    {
      addParents: processedFolderId,
      removeParents: CONFIG.photoFolderId
    }
  );
}

function escapeDriveQuery_(value) {
  return String(value || '').replace(/'/g, '\\\'');
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(CONFIG.spreadsheetId);
}

function getMasterSheet_() {
  const ss = getSpreadsheet_();
  return ss.getSheetByName(CONFIG.masterSheetName) || ss.insertSheet(CONFIG.masterSheetName);
}

function normalizeOcrLines_(text) {
  return String(text || '')
    .replace(/\r/g, '\n')
    .split('\n')
    .map(line => line.replace(/\s+/g, ' ').trim())
    .filter(Boolean);
}

function findCustomer_(lines) {
  const known = ['GB MANUFACTURING', 'HILL MANUFACTURING', 'MANCOR INDIANA', 'MEC', 'T & W STAMPING', 'T&W STAMPING'];
  const upperLines = lines.map(line => line.toUpperCase());
  for (const customer of known) {
    if (upperLines.some(line => line.includes(customer))) return customer.replace('T&W', 'T & W');
  }
  const candidate = lines.find(line => /^[A-Z0-9 &.-]{3,}$/.test(line) && /MANUFACTURING|STAMPING|MANCOR|MEC|INDUSTRIES/.test(line.toUpperCase()));
  return candidate || '';
}

function findPartNumber_(lines, tokens, controlNumber) {
  const joined = lines.join(' ');
  const labeled = joined.match(/PART\s*NUMBER:?\s*([A-Z0-9-]{5,})/i);
  if (labeled) return labeled[1].toUpperCase();
  const ignored = new Set(['ECOAT', 'POWDER', 'MASK', 'OPERATION', 'NUMBER', 'QUANTITY', 'PARTS', 'DATE', String(controlNumber || '')]);
  return (tokens || []).find(token => {
    const value = String(token).toUpperCase();
    return !ignored.has(value) && value !== String(controlNumber || '') && /[0-9]/.test(value) && value.length >= 5;
  }) || '';
}

function findControlNumber_(lines, tokens) {
  const fiveDigit = (tokens || []).filter(token => /^\d{5}$/.test(token));
  if (fiveDigit.length) return fiveDigit[fiveDigit.length - 1];
  const joined = lines.join(' ');
  const match = joined.match(/\b5\d{4}\b/);
  return match ? match[0] : '';
}

function findQuantity_(lines) {
  const joined = lines.join(' ');
  const labeled = joined.match(/(?:QTY|QUANTITY)\D{0,8}(\d{1,6})/i);
  if (labeled) return labeled[1];
  const ecoatPowderMatch = joined.match(/\bECOAT\s+(\d{1,6})\s+POWDER\s+(\d{1,6})\b/i);
  if (ecoatPowderMatch) return ecoatPowderMatch[2];
  const powderIndex = lines.findIndex(line => /^POWDER$/i.test(line) || /POWDER\s+\d{1,6}/i.test(line));
  if (powderIndex >= 0) {
    const windowText = lines.slice(powderIndex, powderIndex + 4).join(' ');
    const match = windowText.match(/\b\d{1,6}\b/);
    if (match) return match[0];
  }
  return '';
}

function findPowder_(text) {
  const code = (text.match(/\b(?:YN|PCS)\d{3,6}\b/i) || [''])[0].toUpperCase();
  let color = '';
  if (/MED(?:IUM)?\s*GL(?:OSS)?\s*BL(?:ACK|K)/i.test(text)) color = 'MED GL BLACK';
  if (!color && /GL\s*BLACK/i.test(text)) color = 'GL BLACK';
  if (code && color) return code + ' ' + color;
  return code || color;
}

function normalizePowderCode_(value) {
  const text = String(value || '').trim().toUpperCase();
  const match = text.match(/\b(?:YN|PCS)\d+\b/);
  return match ? match[0] : text;
}

function findFilmSpec_(text) {
  const match = text.match(/\b\d(?:\.\d)?\s*-\s*\d(?:\.\d)?\s*MIL\b/i);
  return match ? match[0].toUpperCase() : '';
}

function findJobLot_(lines) {
  const joined = lines.join(' ');
  const po = joined.match(/\bPO\s*\d+\b/i);
  if (po) return po[0].replace(/\s+/g, '');
  const dateLot = joined.match(/\b0[1-9]\d{6}\b/);
  return dateLot ? dateLot[0] : '';
}

function findToolingText_(lines) {
  return lines
    .filter(line => /TOOL|PC\/LB|PC\/HR|POWERCRON|BANGER|LANDICE|KICK PLATE|PREHANG|SPACE/i.test(line))
    .join('\n');
}

function findSpecialHandling_(lines) {
  return lines
    .filter(line => /PACK OUT|PROTECT|FOAM|PREHANG|SEPARATE/i.test(line))
    .join('; ');
}

function findDescription_(lines, partNumber) {
  if (!partNumber) return '';
  const idx = lines.findIndex(line => line.includes(partNumber));
  const nearby = lines.slice(Math.max(0, idx), idx + 4).find(line => /\([A-Z/]+\)|BRKT|HEAT|SUPPORT|FOOT|STEP|SHIELD/i.test(line));
  return nearby || '';
}
