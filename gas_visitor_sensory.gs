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
    schemaVersion: 1
  });
}
