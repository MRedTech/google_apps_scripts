// ==================================================
// SECURE ENTRY VISITOR - SENSORY (GOOGLE APPS SCRIPT)
// Worker -> GAS mirror for Google Sheet + Google Drive
//
// Worker contract:
// - POST action="SYNC"         : upsert CHECKED_IN / CHECKED_OUT by record ID
// - POST action="DELETE_DRIVE" : trash Drive files during Worker retention cleanup
// - POST action="CLEANUP_AGE"  : optional manual Sheet/Drive retention cleanup
//
// Sheet columns:
// A CHECK IN | B NAME | C MYKAD/PASSPORT | D REG.NUM | E CONTACT
// F VISITOR PASS NUMBER | G UNIT NUMBER | H CHECK OUT | I STATUS
// J PHOTO LINK | K RECORD ID | L SYNC VERSION
// ==================================================

const SHEET_NAME = "VISITOR";

// These can be overridden using Apps Script > Project Settings > Script Properties:
// SPREADSHEET_ID, DRIVE_FOLDER_ID, SYNC_TOKEN
const SPREADSHEET_ID_FALLBACK = "1qQTFba0FlZWDxG9yUt-Ofzs16AwjXInKyPFee_BZJHs";
const DRIVE_FOLDER_ID_FALLBACK = "1HLaxdhQ2E4UM4PZaeJHAx9DIU0gFD3Gu";
const SYNC_TOKEN_FALLBACK = "se_sync_4YUsinaKn9no5wgRt9lcFfVoek6jlG4SkrRHLw84X2o";
const NOTICE_FOLDER_ID_FALLBACK = "1c7w3sDcie1szqxDrbzqwHUN5S-KKW6FE";
const PMO_EMAIL_FALLBACK = "edreborn86@gmail.com";
const NOTICE_SITE_NAME_FALLBACK = "SENSORY RESIDENCE";

const TIME_ZONE = "Asia/Kuala_Lumpur";
const RETENTION_DAYS = 90;

const COL = Object.freeze({
  CHECK_IN: 1,
  NAME: 2,
  MYKAD_PASSPORT: 3,
  REGNUM: 4,
  CONTACT: 5,
  VISITOR_PASS: 6,
  UNIT_NUMBER: 7,
  CHECK_OUT: 8,
  STATUS: 9,
  PHOTO_LINK: 10,
  RECORD_ID: 11,
  SYNC_VERSION: 12
});

const HEADERS = [[
  "CHECK IN",
  "NAME",
  "MYKAD/PASSPORT",
  "REG.NUM",
  "CONTACT",
  "VISITOR PASS NUMBER",
  "UNIT NUMBER",
  "CHECK OUT",
  "STATUS",
  "PHOTO LINK",
  "RECORD ID",
  "SYNC VERSION"
]];

/** =========================
 * Basic helpers
 * ========================= */
function output_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function toText_(value) {
  return value == null ? "" : String(value).trim();
}

function toUpper_(value) {
  return toText_(value).toUpperCase();
}

function normalizeText_(value) {
  const text = toText_(value);
  if (!text) return "";

  return text
    .normalize("NFKD")
    .replace(/[^\w\s\/\-\&\(\)]/g, "")
    .replace(/_/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function escapeFormulaText_(value) {
  return toText_(value).replace(/"/g, '""');
}

function formatPhone_(value) {
  let phone = toText_(value).replace(/[^0-9]/g, "");
  if (phone && !phone.startsWith("0")) phone = "0" + phone;
  return phone ? "'" + phone : "";
}

function positiveInt_(value, fallback) {
  const parsed = parseInt(toText_(value), 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

function normalizeStatus_(value, checkOutTime) {
  const status = toUpper_(value).replace(/[^A-Z0-9]/g, "");
  if (status === "CHECKEDOUT" || toText_(checkOutTime)) return "CHECKED_OUT";
  return "CHECKED_IN";
}

function formatIsoDateTime_(value) {
  const raw = toText_(value);
  if (!raw) return "";

  const date = new Date(raw);
  if (isNaN(date.getTime())) return raw;

  return Utilities.formatDate(date, TIME_ZONE, "d/M/yyyy HH:mm:ss");
}

function parseDmyDateTime_(value) {
  const raw = toText_(value);
  const match = raw.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
  );
  if (!match) return null;

  const day = parseInt(match[1], 10);
  const month = parseInt(match[2], 10) - 1;
  const year = parseInt(match[3], 10);
  const hour = parseInt(match[4] || "0", 10);
  const minute = parseInt(match[5] || "0", 10);
  const second = parseInt(match[6] || "0", 10);

  const date = new Date(year, month, day, hour, minute, second);
  return isNaN(date.getTime()) ? null : date;
}

function isHttpUrl_(value) {
  return /^https?:\/\//i.test(toText_(value));
}

function extractDriveFileId_(url) {
  const text = toText_(url);
  if (!text) return "";

  let match = text.match(/[?&]id=([A-Za-z0-9_-]{10,})/);
  if (match && match[1]) return match[1];

  match = text.match(/\/d\/([A-Za-z0-9_-]{10,})/);
  return match && match[1] ? match[1] : "";
}

function driveViewUrl_(fileId) {
  const id = toText_(fileId);
  return id ? "https://drive.google.com/uc?export=view&id=" + id : "";
}

/** =========================
 * Configuration helpers
 * ========================= */
function scriptProperty_(name, fallback) {
  const value = PropertiesService.getScriptProperties().getProperty(name);
  return toText_(value) || toText_(fallback);
}

function spreadsheetId_() {
  const id = scriptProperty_("SPREADSHEET_ID", SPREADSHEET_ID_FALLBACK);
  if (!id || id === "PASTE_VISITOR_SPREADSHEET_ID") {
    throw new Error("SPREADSHEET_ID is not configured.");
  }
  return id;
}

function driveFolderId_() {
  const id = scriptProperty_("DRIVE_FOLDER_ID", DRIVE_FOLDER_ID_FALLBACK);
  if (!id || id === "PASTE_VISITOR_DRIVE_FOLDER_ID") {
    throw new Error("DRIVE_FOLDER_ID is not configured.");
  }
  return id;
}

function syncToken_() {
  const token = scriptProperty_("SYNC_TOKEN", SYNC_TOKEN_FALLBACK);
  if (!token) throw new Error("SYNC_TOKEN is not configured.");
  return token;
}

function noticeFolderId_() {
  const id = scriptProperty_("NOTICE_FOLDER_ID", NOTICE_FOLDER_ID_FALLBACK);
  if (!id) throw new Error("NOTICE_FOLDER_ID is not configured.");
  return id;
}

function pmoEmail_() {
  const email = scriptProperty_("PMO_EMAIL", PMO_EMAIL_FALLBACK);
  if (!email) throw new Error("PMO_EMAIL is not configured.");
  return email;
}

function noticeSiteName_() {
  return scriptProperty_("NOTICE_SITE_NAME", NOTICE_SITE_NAME_FALLBACK);
}

function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  return active || SpreadsheetApp.openById(spreadsheetId_());
}

function getSheet_() {
  const sheet = getSpreadsheet_().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Sheet not found: " + SHEET_NAME);
  return sheet;
}

function getDriveFolder_() {
  return DriveApp.getFolderById(driveFolderId_());
}

function getNoticeFolder_() {
  return DriveApp.getFolderById(noticeFolderId_());
}

/** =========================
 * One-time Sheet setup
 * Run setupVisitorSheet() once before Web App deployment.
 * ========================= */
function setupVisitorSheet() {
  const spreadsheet = getSpreadsheet_();
  spreadsheet.setSpreadsheetTimeZone(TIME_ZONE);

  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);

  sheet.getRange(1, 1, 1, HEADERS[0].length).setValues(HEADERS);
  sheet.setFrozenRows(1);

  const header = sheet.getRange(1, 1, 1, HEADERS[0].length);
  header.setFontWeight("bold");
  header.setHorizontalAlignment("center");

  // Preserve leading zeroes and exact identifiers.
  sheet.getRange("C:C").setNumberFormat("@");
  sheet.getRange("D:D").setNumberFormat("@");
  sheet.getRange("E:E").setNumberFormat("@");
  sheet.getRange("F:F").setNumberFormat("@");
  sheet.getRange("G:G").setNumberFormat("@");
  sheet.getRange("K:K").setNumberFormat("@");
  sheet.getRange("L:L").setNumberFormat("0");

  sheet.setColumnWidth(COL.CHECK_IN, 145);
  sheet.setColumnWidth(COL.NAME, 220);
  sheet.setColumnWidth(COL.MYKAD_PASSPORT, 145);
  sheet.setColumnWidth(COL.REGNUM, 110);
  sheet.setColumnWidth(COL.CONTACT, 120);
  sheet.setColumnWidth(COL.VISITOR_PASS, 130);
  sheet.setColumnWidth(COL.UNIT_NUMBER, 120);
  sheet.setColumnWidth(COL.CHECK_OUT, 145);
  sheet.setColumnWidth(COL.STATUS, 115);
  sheet.setColumnWidth(COL.PHOTO_LINK, 220);

  // Internal columns used for reliable upsert/version control.
  try { sheet.hideColumns(COL.RECORD_ID, 2); } catch (_) {}

  return {
    success: true,
    sheetName: SHEET_NAME,
    columns: HEADERS[0].length
  };
}

/** =========================
 * Record lookup + row helpers
 * ========================= */
function findRecordRow_(sheet, recordId) {
  const id = toText_(recordId);
  if (!id) return 0;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const finder = sheet
    .getRange(2, COL.RECORD_ID, lastRow - 1, 1)
    .createTextFinder(id)
    .matchEntireCell(true);

  const cell = finder.findNext();
  return cell ? cell.getRow() : 0;
}

function existingPhotoInfo_(sheet, rowNumber) {
  if (!rowNumber) return { driveFileId: "", driveUrl: "" };

  const driveUrl = toText_(
    sheet.getRange(rowNumber, COL.PHOTO_LINK).getDisplayValue()
  );

  return {
    driveFileId: extractDriveFileId_(driveUrl),
    driveUrl: isHttpUrl_(driveUrl) ? driveUrl : ""
  };
}

function makeNameCell_(name, photoUrl) {
  const cleanName = normalizeText_(name);
  const url = toText_(photoUrl);

  if (!url) return cleanName;

  return '=HYPERLINK("' +
    escapeFormulaText_(url) +
    '","' +
    escapeFormulaText_(cleanName) +
    '")';
}

function buildRowValues_(data, photoUrl, syncVersion) {
  const checkIn = formatIsoDateTime_(data.checkInTime || data.createdAt);
  const checkOut = formatIsoDateTime_(data.checkOutTime);
  const status = normalizeStatus_(data.status, data.checkOutTime);

  return [[
    checkIn,                                      // A CHECK IN
    makeNameCell_(data.namePassport, photoUrl),   // B NAME
    normalizeText_(data.mykadPassport),           // C MYKAD/PASSPORT
    normalizeText_(data.regnum),                  // D REG.NUM
    formatPhone_(data.contact),                   // E CONTACT
    normalizeText_(data.visitorPassNumber),       // F VISITOR PASS
    normalizeText_(data.unitNumber),              // G UNIT NUMBER
    checkOut,                                     // H CHECK OUT
    status,                                       // I STATUS
    toText_(photoUrl),                            // J PHOTO LINK
    toText_(data.id),                             // K RECORD ID
    syncVersion                                   // L SYNC VERSION
  ]];
}

/** =========================
 * Image upload
 * ========================= */
function safeFilenamePart_(value) {
  return normalizeText_(value)
    .replace(/[^A-Z0-9\- ]/g, "")
    .replace(/\s+/g, "_")
    .slice(0, 60);
}

function uploadWorkerImage_(data) {
  const imageViewUrl = toText_(data.imageViewUrl);
  if (!imageViewUrl) return { driveFileId: "", driveUrl: "" };

  const response = UrlFetchApp.fetch(imageViewUrl, {
    method: "get",
    followRedirects: true,
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error("Image fetch failed (" + code + ")");
  }

  const namePart =
    safeFilenamePart_(data.namePassport) ||
    safeFilenamePart_(data.regnum) ||
    "VISITOR";

  const passPart = safeFilenamePart_(data.visitorPassNumber);
  const filename =
    (passPart ? passPart + "_" : "") +
    namePart +
    "_" +
    Date.now() +
    ".jpg";

  const blob = response.getBlob().setName(filename);
  const file = getDriveFolder_().createFile(blob);
  const fileId = file.getId();

  return {
    driveFileId: fileId,
    driveUrl: driveViewUrl_(fileId)
  };
}

/** =========================
 * Retention cleanup
 * - Sheet is ordered by CHECK IN because new records append.
 * - Removes the oldest contiguous block older than RETENTION_DAYS.
 * ========================= */
function cleanupByAge_(sheet, retentionDays) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const cutoff = new Date(
    Date.now() - positiveInt_(retentionDays, RETENTION_DAYS) * 86400000
  );

  const timestamps = sheet
    .getRange(2, COL.CHECK_IN, lastRow - 1, 1)
    .getDisplayValues();

  let rowsToDelete = 0;

  for (let index = 0; index < timestamps.length; index++) {
    const date = parseDmyDateTime_(timestamps[index][0]);
    if (date && date < cutoff) rowsToDelete++;
    else break;
  }

  if (rowsToDelete < 1) return 0;

  const photoLinks = sheet
    .getRange(2, COL.PHOTO_LINK, rowsToDelete, 1)
    .getDisplayValues();

  for (let index = 0; index < photoLinks.length; index++) {
    const fileId = extractDriveFileId_(photoLinks[index][0]);
    if (!fileId) continue;

    try {
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch (_) {}
  }

  sheet.deleteRows(2, rowsToDelete);
  return rowsToDelete;
}

/** =========================
 * SYNC upsert
 * - New CHECK IN: append one row and upload photo if supplied.
 * - CHECK OUT: update that same row by Record ID.
 * - Older/equal syncVersion: return success without overwriting newer state.
 * ========================= */
function handleSync_(data) {
  const recordId = toText_(data.id);
  if (!recordId) {
    return output_({
      success: false,
      error: true,
      retryable: false,
      message: "Record ID is required."
    });
  }

  const incomingVersion = positiveInt_(data.syncVersion, 1);
  const lock = LockService.getScriptLock();

  let lockTaken = false;
  let newlyUploadedFileId = "";
  let rowCommitted = false;

  try {
    lock.waitLock(30000);
    lockTaken = true;

    const sheet = getSheet_();
    const rowNumber = findRecordRow_(sheet, recordId);

    if (rowNumber) {
      const storedVersion = positiveInt_(
        sheet.getRange(rowNumber, COL.SYNC_VERSION).getValue(),
        1
      );

      const existingPhoto = existingPhotoInfo_(sheet, rowNumber);

      // Idempotency and stale-version protection.
      if (incomingVersion <= storedVersion) {
        return output_({
          success: true,
          duplicate: true,
          stale: incomingVersion < storedVersion,
          recordId: recordId,
          syncVersion: storedVersion,
          driveFileId: existingPhoto.driveFileId,
          driveUrl: existingPhoto.driveUrl
        });
      }
    }

    let photoInfo = rowNumber
      ? existingPhotoInfo_(sheet, rowNumber)
      : { driveFileId: "", driveUrl: "" };

    // Upload only when this record does not already have a Drive photo.
    if (!photoInfo.driveUrl && toText_(data.imageViewUrl)) {
      photoInfo = uploadWorkerImage_(data);
      newlyUploadedFileId = photoInfo.driveFileId;
    }

    const rowValues = buildRowValues_(
      data,
      photoInfo.driveUrl,
      incomingVersion
    );

    if (rowNumber) {
      sheet
        .getRange(rowNumber, 1, 1, HEADERS[0].length)
        .setValues(rowValues);
    } else {
      const newRow = sheet.getLastRow() + 1;
      sheet
        .getRange(newRow, 1, 1, HEADERS[0].length)
        .setValues(rowValues);
    }

    rowCommitted = true;

    let cleanupWarning = "";
    let deletedRows = 0;

    try {
      deletedRows = cleanupByAge_(sheet, RETENTION_DAYS);
    } catch (cleanupError) {
      cleanupWarning = "Record saved; cleanup warning: " +
        (cleanupError.message || String(cleanupError));
    }

    return output_({
      success: true,
      recordId: recordId,
      syncVersion: incomingVersion,
      status: normalizeStatus_(data.status, data.checkOutTime),
      driveFileId: photoInfo.driveFileId,
      driveUrl: photoInfo.driveUrl,
      deletedRows: deletedRows,
      warning: cleanupWarning
    });
  } catch (error) {
    // Prevent orphaned Drive images if the Sheet write failed.
    if (newlyUploadedFileId && !rowCommitted) {
      try {
        DriveApp.getFileById(newlyUploadedFileId).setTrashed(true);
      } catch (_) {}
    }

    return output_({
      success: false,
      error: true,
      retryable: true,
      message: error.message || String(error)
    });
  } finally {
    if (lockTaken) {
      try { lock.releaseLock(); } catch (_) {}
    }
  }
}


/** =========================
 * OVERNIGHT PARKING CHARGE NOTICE
 * - Runs only for totalChargeCents > 0.
 * - Creates a PDF in the configured Drive folder and emails it to PMO.
 * - Guard house does not receive payment.
 * ========================= */
function moneyFromCents_(value) {
  const cents = Math.max(0, Math.floor(Number(value) || 0));
  return "RM" + (cents / 100).toFixed(2);
}

function durationFromSeconds_(value) {
  let total = Math.max(0, Math.floor(Number(value) || 0));
  const hours = Math.floor(total / 3600);
  total %= 3600;
  const minutes = Math.floor(total / 60);
  const seconds = total % 60;
  return String(hours).padStart(2, "0") + ":" +
    String(minutes).padStart(2, "0") + ":" +
    String(seconds).padStart(2, "0");
}

function safeNoticeFilename_(value) {
  return toUpper_(value)
    .replace(/[^A-Z0-9_-]/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 80);
}

function styleParagraph_(paragraph, options) {
  options = options || {};
  if (options.align) paragraph.setAlignment(options.align);
  if (options.spacingBefore != null) paragraph.setSpacingBefore(options.spacingBefore);
  if (options.spacingAfter != null) paragraph.setSpacingAfter(options.spacingAfter);
  if (options.lineSpacing != null) paragraph.setLineSpacing(options.lineSpacing);

  const text = paragraph.editAsText();
  if (options.bold != null) text.setBold(!!options.bold);
  if (options.size) text.setFontSize(options.size);
  if (options.color) text.setForegroundColor(options.color);
  if (options.fontFamily) text.setFontFamily(options.fontFamily);
  return paragraph;
}

function prepareNoticeCell_(cell, backgroundColor, padding) {
  const inset = padding == null ? 8 : padding;
  cell.clear();
  cell.setBackgroundColor(backgroundColor || '#FFFFFF');
  cell.setPaddingTop(inset);
  cell.setPaddingBottom(inset);
  cell.setPaddingLeft(inset);
  cell.setPaddingRight(inset);
  cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
  return cell;
}

function appendNoticeField_(cell, label, value, options) {
  options = options || {};
  prepareNoticeCell_(
    cell,
    options.backgroundColor || '#F7FBF8',
    options.padding == null ? 8 : options.padding
  );

  styleParagraph_(cell.appendParagraph(toUpper_(label)), {
    bold: true,
    size: options.labelSize || 8,
    color: options.labelColor || '#5E7465',
    fontFamily: options.fontFamily || 'Arial',
    spacingAfter: 2
  });

  styleParagraph_(cell.appendParagraph(toText_(value) || '-'), {
    bold: options.valueBold !== false,
    size: options.valueSize || 10,
    color: options.valueColor || '#173C28',
    fontFamily: options.fontFamily || 'Arial',
    spacingAfter: 0
  });

  return cell;
}

function noticeUnitsVisited_(data) {
  const units = [];
  const seen = {};

  function addUnit_(value) {
    const unit = toUpper_(value);
    const key = unit.replace(/[^A-Z0-9]/g, '');
    if (!key || seen[key]) return;
    seen[key] = true;
    units.push(unit);
  }

  const supplied = Array.isArray(data && data.unitsVisited)
    ? data.unitsVisited
    : [];

  for (let index = 0; index < supplied.length; index++) {
    addUnit_(supplied[index]);
  }

  if (!units.length) {
    const activities = Array.isArray(data && data.activityLog)
      ? data.activityLog
      : [];

    for (let index = activities.length - 1; index >= 0; index--) {
      addUnit_(activities[index] && activities[index].unitNumber);
    }
  }

  if (!units.length) addUnit_(data && data.unitNumber);
  return units;
}

function isMultiUnitNotice_(data) {
  return noticeUnitsVisited_(data).length > 1;
}

function appendModernNoticeSummary_(body, data, regnum) {
  const summaryTable = body.appendTable([
    ['', ''],
    ['', ''],
    ['', '']
  ]);

  summaryTable.setBorderWidth(0);

  const unitsVisited = noticeUnitsVisited_(data);
  const multiUnit = unitsVisited.length > 1;
  const unitLabel = multiUnit ? 'UNITS VISITED' : 'UNIT NUMBER';
  const unitValue = unitsVisited.length
    ? unitsVisited.join(', ')
    : (toUpper_(data.unitNumber) || '-');

  const fields = [
    ['VISITOR NAME', toUpper_(data.visitorName) || '-'],
    ['SESSION STARTED', formatIsoDateTime_(data.sessionStartedAt) || '-'],
    ['VEHICLE REG. NO.', regnum || '-'],
    ['FINAL CHECK OUT', formatIsoDateTime_(data.finalCheckOutAt) || '-'],
    [unitLabel, unitValue],
    ['TOTAL PARKING DURATION', durationFromSeconds_(data.totalDurationSeconds)]
  ];

  for (let index = 0; index < fields.length; index++) {
    const rowIndex = Math.floor(index / 2);
    const columnIndex = index % 2;
    const cell = summaryTable.getCell(rowIndex, columnIndex);
    cell.setWidth(250);
    appendNoticeField_(cell, fields[index][0], fields[index][1], {
      backgroundColor: rowIndex % 2 === 0 ? '#F3F9F4' : '#FFFFFF',
      padding: 8,
      labelSize: 8,
      valueSize: fields[index][0] === 'UNITS VISITED' ? 8 : 10
    });
  }

  return summaryTable;
}

function appendChargeSummary_(body, data) {
  const chargeTable = body.appendTable([['', '']]);
  chargeTable.setBorderWidth(0);

  const detailsCell = chargeTable.getCell(0, 0);
  prepareNoticeCell_(detailsCell, '#F2F8F3', 10);
  detailsCell.setWidth(345);

  styleParagraph_(detailsCell.appendParagraph('CHARGE SUMMARY'), {
    bold: true,
    size: 11,
    color: '#176B39',
    fontFamily: 'Montserrat',
    spacingAfter: 5
  });

  const detailsTable = detailsCell.appendTable([
    ['FIRST 24 HOURS', 'FREE'],
    ['ADDITIONAL 24-HOUR PERIOD(S)', String(Math.max(0, Number(data.chargeBlocks) || 0))],
    ['RATE', moneyFromCents_(data.rateCents) + ' PER ADDITIONAL STARTED 24-HOUR PERIOD']
  ]);

  detailsTable.setBorderColor('#DCE9DF');
  detailsTable.setBorderWidth(0.5);

  for (let r = 0; r < detailsTable.getNumRows(); r++) {
    const row = detailsTable.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setBackgroundColor('#F8FBF9');
      cell.setPaddingTop(5);
      cell.setPaddingBottom(5);
      cell.setPaddingLeft(5);
      cell.setPaddingRight(5);
      cell.setWidth(c === 0 ? 185 : 150);

      const paragraph = cell.getChild(0).asParagraph();
      paragraph.setAlignment(
        c === 0
          ? DocumentApp.HorizontalAlignment.LEFT
          : DocumentApp.HorizontalAlignment.RIGHT
      );
      paragraph.setSpacingBefore(0).setSpacingAfter(0);
      paragraph.editAsText()
        .setFontFamily('Arial')
        .setFontSize(r === 2 ? 7 : 8)
        .setBold(c === 1)
        .setForegroundColor(c === 0 ? '#5E7465' : '#173C28');
    }
  }

  const totalCell = chargeTable.getCell(0, 1);
  prepareNoticeCell_(totalCell, '#E8F4EA', 12);
  totalCell.setWidth(155);

  styleParagraph_(totalCell.appendParagraph('TOTAL AMOUNT DUE'), {
    align: DocumentApp.HorizontalAlignment.CENTER,
    bold: true,
    size: 9,
    color: '#176B39',
    fontFamily: 'Montserrat',
    spacingAfter: 8
  });

  styleParagraph_(
    totalCell.appendParagraph(moneyFromCents_(data.totalChargeCents)),
    {
      align: DocumentApp.HorizontalAlignment.CENTER,
      bold: true,
      size: 24,
      color: '#176B39',
      fontFamily: 'Montserrat',
      spacingAfter: 0
    }
  );

  return chargeTable;
}

function appendChargeAllocation_(body, data) {
  const unitsVisited = noticeUnitsVisited_(data);
  const allocation = Array.isArray(data && data.chargeAllocation)
    ? data.chargeAllocation
    : [];

  // Keep ordinary one-unit notices concise.
  if (unitsVisited.length < 2 || !allocation.length) return null;

  styleParagraph_(body.appendParagraph('CHARGE ALLOCATION'), {
    bold: true,
    size: 10,
    color: '#176B39',
    fontFamily: 'Montserrat',
    spacingAfter: 5
  });

  const rows = [[
    'CHARGE PERIOD',
    'UNIT NUMBER',
    'AMOUNT'
  ]];

  for (let index = 0; index < allocation.length; index++) {
    const item = allocation[index] || {};
    const periodNumber = Math.max(1, Number(item.periodNumber) || (index + 1));
    rows.push([
      'ADDITIONAL 24-HOUR PERIOD ' + periodNumber,
      toUpper_(item.unitNumber) || '-',
      moneyFromCents_(
        item.amountCents == null ? data.rateCents : item.amountCents
      )
    ]);
  }

  const table = body.appendTable(rows);
  table.setBorderColor('#DCE8DF');
  table.setBorderWidth(0.5);

  const widths = [285, 135, 80];

  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setWidth(widths[c]);
      cell.setPaddingTop(r === 0 ? 5 : 4);
      cell.setPaddingBottom(r === 0 ? 5 : 4);
      cell.setPaddingLeft(5);
      cell.setPaddingRight(5);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      cell.setBackgroundColor(
        r === 0
          ? '#E6F1E8'
          : (r % 2 === 0 ? '#F8FBF9' : '#FFFFFF')
      );

      const paragraph = cell.getChild(0).asParagraph();
      paragraph
        .setAlignment(
          c === 0
            ? DocumentApp.HorizontalAlignment.LEFT
            : DocumentApp.HorizontalAlignment.CENTER
        )
        .setSpacingBefore(0)
        .setSpacingAfter(0);

      paragraph.editAsText()
        .setFontFamily('Arial')
        .setFontSize(r === 0 ? 7 : 7.5)
        .setBold(r === 0 || c === 2)
        .setForegroundColor(r === 0 ? '#176B39' : '#243F2E');
    }
  }

  return table;
}

function appendModernActivityTable_(body, activities) {
  const activityRows = [[
    'ENTRY',
    'PASS NO.',
    'UNIT NUMBER',
    'CHECKED IN',
    'CHECKED OUT',
    'DURATION'
  ]];

  for (let index = 0; index < activities.length; index++) {
    const item = activities[index] || {};
    activityRows.push([
      'ENTRY ' + (activities.length - index),
      toUpper_(item.visitorPassNumber) || '-',
      toUpper_(item.unitNumber) || '-',
      formatIsoDateTime_(item.checkInTime) || '-',
      formatIsoDateTime_(item.checkOutTime) || '-',
      durationFromSeconds_(item.durationSeconds)
    ]);
  }

  if (!activities.length) {
    activityRows.push(['-', '-', '-', '-', '-', '-']);
  }

  const table = body.appendTable(activityRows);
  table.setBorderColor('#DCE8DF');
  table.setBorderWidth(0.5);

  const widths = [50, 50, 70, 130, 130, 70];

  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setWidth(widths[c]);
      cell.setPaddingTop(r === 0 ? 5 : 4);
      cell.setPaddingBottom(r === 0 ? 5 : 4);
      cell.setPaddingLeft(2);
      cell.setPaddingRight(2);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      cell.setBackgroundColor(
        r === 0
          ? '#E6F1E8'
          : (r % 2 === 0 ? '#F8FBF9' : '#FFFFFF')
      );

      const paragraph = cell.getChild(0).asParagraph();
      paragraph
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setSpacingBefore(0)
        .setSpacingAfter(0);

      paragraph.editAsText()
        .setFontFamily('Arial')
        .setFontSize(r === 0 ? 6.5 : 6)
        .setBold(r === 0)
        .setForegroundColor(r === 0 ? '#176B39' : '#243F2E');
    }
  }

  return table;
}

function appendFixedNoticeFooter_(doc) {
  const footer = doc.addFooter();
  footer.clear();
  footer.appendHorizontalRule();

  styleParagraph_(
    footer.appendParagraph('PAYMENT IS NOT COLLECTED AT THE GUARD HOUSE.'),
    {
      align: DocumentApp.HorizontalAlignment.CENTER,
      bold: true,
      size: 8,
      color: '#176B39',
      fontFamily: 'Arial',
      spacingBefore: 4,
      spacingAfter: 2
    }
  );

  styleParagraph_(
    footer.appendParagraph(
      'This notice is generated from the Secure Entry recorded check-in, check-out and accumulated parking activity.'
    ),
    {
      align: DocumentApp.HorizontalAlignment.CENTER,
      size: 7,
      color: '#718579',
      fontFamily: 'Arial',
      spacingAfter: 3
    }
  );

  styleParagraph_(footer.appendParagraph('POWERED BY MRED TECH'), {
    align: DocumentApp.HorizontalAlignment.CENTER,
    bold: true,
    size: 7,
    color: '#95A39A',
    fontFamily: 'Montserrat',
    spacingAfter: 0
  });

  return footer;
}

function findExistingNoticePdf_(folder, filename) {
  const files = folder.getFilesByName(filename);
  return files.hasNext() ? files.next() : null;
}

function createParkingNoticePdf_(data) {
  const noticeNumber = toUpper_(data.noticeNumber);
  const version = positiveInt_(data.noticeVersion, 1);
  const regnum = toUpper_(data.vehicleRegNumber);
  const filename = safeNoticeFilename_(
    'OVERNIGHT_PARKING_NOTICE_' + noticeNumber + '_' + regnum + '_V' + version
  ) + '.pdf';

  const folder = getNoticeFolder_();
  const existing = findExistingNoticePdf_(folder, filename);
  if (existing) {
    return {
      file: existing,
      filename: filename,
      created: false
    };
  }

  const doc = DocumentApp.create('TEMP_' + filename.replace(/\.pdf$/i, ''));
  const body = doc.getBody();
  body.clear();

  // Reserve enough space so the real Google Docs footer stays at the bottom
  // and never overlaps the activity log.
  body.setMarginTop(30);
  body.setMarginBottom(82);
  body.setMarginLeft(36);
  body.setMarginRight(36);

  const headingFont = 'Montserrat';

  // Main title first, followed by the site name. Both use the same font family.
  styleParagraph_(body.appendParagraph('OVERNIGHT PARKING CHARGE NOTICE'), {
    align: DocumentApp.HorizontalAlignment.CENTER,
    bold: true,
    size: 15,
    color: '#176B39',
    fontFamily: headingFont,
    spacingAfter: 5
  });

  styleParagraph_(
    body.appendParagraph(toUpper_(data.siteName || noticeSiteName_())),
    {
      align: DocumentApp.HorizontalAlignment.CENTER,
      bold: true,
      size: 10,
      color: '#176B39',
      fontFamily: headingFont,
      spacingAfter: 5
    }
  );

  styleParagraph_(body.appendParagraph('NOTICE NO.: ' + noticeNumber), {
    align: DocumentApp.HorizontalAlignment.CENTER,
    bold: true,
    size: 8,
    color: '#6A7F70',
    fontFamily: 'Arial',
    spacingAfter: 10
  });

  // Visitor Pass is intentionally omitted here because it is already shown
  // for every entry in the Parking Activity Log.
  appendModernNoticeSummary_(body, data, regnum);

  body.appendParagraph('').setSpacingAfter(2);
  appendChargeSummary_(body, data);

  if (isMultiUnitNotice_(data)) {
    body.appendParagraph('').setSpacingAfter(2);
    appendChargeAllocation_(body, data);
  }

  body.appendParagraph('').setSpacingAfter(2);
  styleParagraph_(
    body.appendParagraph('PARKING ACTIVITY LOG - LATEST ACTIVITY FIRST'),
    {
      bold: true,
      size: 10,
      color: '#176B39',
      fontFamily: headingFont,
      spacingAfter: 5
    }
  );

  const activities = Array.isArray(data.activityLog) ? data.activityLog : [];
  appendModernActivityTable_(body, activities);

  // A real document footer remains at the bottom even when the activity log
  // contains only one or two records. It repeats automatically on extra pages.
  appendFixedNoticeFooter_(doc);

  doc.saveAndClose();

  const tempFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = tempFile.getAs(MimeType.PDF).setName(filename);
  const pdfFile = folder.createFile(pdfBlob);

  try { tempFile.setTrashed(true); } catch (_) {}

  return {
    file: pdfFile,
    filename: filename,
    created: true
  };
}

function parkingNoticeEmailBody_(data) {
  const unitsVisited = noticeUnitsVisited_(data);
  const multiUnit = unitsVisited.length > 1;
  const lines = [
    "OVERNIGHT PARKING CHARGE NOTICE",
    "",
    "NOTICE NO.: " + toUpper_(data.noticeNumber),
    "VEHICLE REG. NO.: " + toUpper_(data.vehicleRegNumber),
    (multiUnit ? "UNITS VISITED: " : "UNIT NUMBER: ") +
      (unitsVisited.join(", ") || toUpper_(data.unitNumber) || "-"),
    "VISITOR PASS NO.: " + toUpper_(data.visitorPassNumber),
    "FINAL CHECK OUT: " + (formatIsoDateTime_(data.finalCheckOutAt) || "-"),
    "TOTAL PARKING DURATION: " + durationFromSeconds_(data.totalDurationSeconds),
    "TOTAL AMOUNT DUE: " + moneyFromCents_(data.totalChargeCents)
  ];

  const allocation = Array.isArray(data && data.chargeAllocation)
    ? data.chargeAllocation
    : [];

  if (multiUnit && allocation.length) {
    lines.push("", "CHARGE ALLOCATION:");
    for (let index = 0; index < allocation.length; index++) {
      const item = allocation[index] || {};
      const periodNumber = Math.max(1, Number(item.periodNumber) || (index + 1));
      lines.push(
        "ADDITIONAL 24-HOUR PERIOD " + periodNumber +
        " - " + (toUpper_(item.unitNumber) || "-") +
        " - " + moneyFromCents_(
          item.amountCents == null ? data.rateCents : item.amountCents
        )
      );
    }
  }

  lines.push(
    "",
    "The attached PDF contains the complete parking activity log.",
    "Payment is not collected at the guard house.",
    "",
    "Secure Entry - Sensory Residence"
  );

  return lines.join("\n");
}

function handleParkingNotice_(data) {
  const totalChargeCents = Math.max(0, Math.floor(Number(data.totalChargeCents) || 0));
  if (totalChargeCents <= 0) {
    return output_({
      success: false,
      error: true,
      retryable: false,
      message: "Parking Notice is not required when the total charge is RM0.00."
    });
  }

  const noticeId = toText_(data.noticeId);
  const noticeNumber = toUpper_(data.noticeNumber);
  if (!noticeId || !noticeNumber) {
    return output_({
      success: false,
      error: true,
      retryable: false,
      message: "Notice ID and Notice Number are required."
    });
  }

  const lock = LockService.getScriptLock();
  let lockTaken = false;

  try {
    lock.waitLock(30000);
    lockTaken = true;

    const result = createParkingNoticePdf_(data);
    const pdfFile = result.file;

    let metadata = {};
    try {
      metadata = JSON.parse(toText_(pdfFile.getDescription()) || "{}");
    } catch (_) {
      metadata = {};
    }

    if (
      metadata.noticeId === noticeId &&
      Number(metadata.noticeVersion || 0) === positiveInt_(data.noticeVersion, 1) &&
      metadata.emailStatus === "SENT"
    ) {
      return output_({
        success: true,
        duplicate: true,
        noticeNumber: noticeNumber,
        pdfFileId: pdfFile.getId(),
        pdfUrl: "https://drive.google.com/file/d/" + pdfFile.getId() + "/view",
        emailTo: metadata.emailTo || pmoEmail_(),
        emailSentAt: metadata.emailSentAt || ""
      });
    }

    const emailTo = pmoEmail_();
    const subject =
      "OVERNIGHT PARKING NOTICE - " +
      toUpper_(data.vehicleRegNumber) +
      " - " +
      moneyFromCents_(totalChargeCents);

    MailApp.sendEmail({
      to: emailTo,
      subject: subject,
      body: parkingNoticeEmailBody_(data),
      attachments: [pdfFile.getBlob().setName(result.filename)],
      name: "Secure Entry - Sensory Residence"
    });

    const emailSentAt = new Date().toISOString();
    pdfFile.setDescription(JSON.stringify({
      noticeId: noticeId,
      noticeNumber: noticeNumber,
      noticeVersion: positiveInt_(data.noticeVersion, 1),
      emailStatus: "SENT",
      emailTo: emailTo,
      emailSentAt: emailSentAt
    }));

    const previousPdfId = toText_(data.existingPdfFileId);
    if (previousPdfId && previousPdfId !== pdfFile.getId()) {
      try { DriveApp.getFileById(previousPdfId).setTrashed(true); } catch (_) {}
    }

    return output_({
      success: true,
      noticeNumber: noticeNumber,
      pdfFileId: pdfFile.getId(),
      pdfUrl: "https://drive.google.com/file/d/" + pdfFile.getId() + "/view",
      emailTo: emailTo,
      emailSentAt: emailSentAt
    });
  } catch (error) {
    return output_({
      success: false,
      error: true,
      retryable: true,
      message: error.message || String(error)
    });
  } finally {
    if (lockTaken) {
      try { lock.releaseLock(); } catch (_) {}
    }
  }
}

/** =========================
 * DELETE_DRIVE action
 * ========================= */
function handleDeleteDrive_(data) {
  const ids = Array.isArray(data.fileIds) ? data.fileIds : [];
  let deletedCount = 0;

  for (let index = 0; index < ids.length; index++) {
    const fileId = toText_(ids[index]);
    if (!fileId) continue;

    try {
      DriveApp.getFileById(fileId).setTrashed(true);
      deletedCount++;
    } catch (_) {}
  }

  return output_({
    success: true,
    action: "DELETE_DRIVE",
    deletedCount: deletedCount
  });
}

/** =========================
 * Web App entry points
 * ========================= */
function doPost(e) {
  let data = {};

  try {
    data = JSON.parse(
      e && e.postData && e.postData.contents
        ? e.postData.contents
        : "{}"
    );
  } catch (error) {
    return output_({
      success: false,
      error: true,
      retryable: false,
      message: "Invalid JSON: " + error.message
    });
  }

  try {
    if (toText_(data.token) !== syncToken_()) {
      return output_({
        success: false,
        error: true,
        retryable: false,
        message: "Unauthorized"
      });
    }

    const action = toUpper_(data.action || "SYNC");

    if (action === "SYNC") {
      return handleSync_(data);
    }

    if (action === "PARKING_NOTICE") {
      return handleParkingNotice_(data);
    }

    if (action === "DELETE_DRIVE") {
      return handleDeleteDrive_(data);
    }

    if (action === "CLEANUP_AGE") {
      const lock = LockService.getScriptLock();
      let lockTaken = false;

      try {
        lock.waitLock(30000);
        lockTaken = true;

        const deletedRows = cleanupByAge_(getSheet_(), RETENTION_DAYS);
        return output_({
          success: true,
          action: "CLEANUP_AGE",
          deletedRows: deletedRows
        });
      } finally {
        if (lockTaken) {
          try { lock.releaseLock(); } catch (_) {}
        }
      }
    }

    return output_({
      success: false,
      error: true,
      retryable: false,
      message: "Unsupported action."
    });
  } catch (error) {
    return output_({
      success: false,
      error: true,
      retryable: true,
      message: error.message || String(error)
    });
  }
}

function doGet() {
  return output_({
    ok: true,
    service: "secure-entry-visitor-gas",
    sheetName: SHEET_NAME,
    schemaVersion: 2,
    parkingNotice: true,
    noticeEmail: pmoEmail_()
  });
}
