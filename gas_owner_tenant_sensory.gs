// ==================================================
// SECURE ENTRY - OWNER | TENANT - SENSORY
// GOOGLE APPS SCRIPT - SYNC MIRROR
// ==================================================
//
// Source of truth:
//   Frontend -> Cloudflare Worker -> D1 / R2
//
// GAS purpose only:
//   1. Receive background SYNC from Worker
//   2. Fetch NEW ID photo from protected Worker /image URL
//   3. Save photo to Google Drive
//   4. Append clean registration row to Google Sheet
//   5. Return Drive file ID + URL to Worker
//   6. Support DELETE_DRIVE from Worker retention cleanup
//   7. Keep Sheet mirror to 90 days
//
// IMPORTANT:
// - Search remains on Cloudflare D1. No GAS search/cache code is used here.
// - UNIT NUMBER and CATEGORY remain separate columns.
// - Do NOT rebuild legacy "OWNER ( UNIT )" category format.
// ==================================================

const SHEET_NAME = "OWNER TENANT";

// Replace these two values before deployment.
const SPREADSHEET_ID = "1rrDeFJAT33bOEoGOBSmcdpAj6QJnIqVyNytwd0JQL7E";
const DRIVE_FOLDER_ID = "18cm-Bzxeud0rO0mf7rvZnz1Z9tJXDsPy";

// SYNC_TOKEN must be configured in Apps Script Script Properties and must be
// EXACTLY the same as the secret saved in Cloudflare Worker owner-tenant-sensory.

const RETENTION_DAYS = 90;
const TIMEZONE = "Asia/Kuala_Lumpur";

// Sheet columns:
// A TIMESTAMP
// B NAME
// C MYKAD/PASSPORT
// D REG.NUM
// E CONTACT
// F UNIT NUMBER
// G CATEGORY
// H REASON
// I REASON OTHER
// J TOWER
// K PHOTO LINK
// L RECORD ID
const HEADERS = [
  "TIMESTAMP",
  "NAME",
  "MYKAD/PASSPORT",
  "REG.NUM",
  "CONTACT",
  "UNIT NUMBER",
  "CATEGORY",
  "REASON",
  "REASON OTHER",
  "TOWER",
  "PHOTO LINK",
  "RECORD ID"
];

const COL = {
  TIMESTAMP: 1,
  NAME: 2,
  ID: 3,
  REG: 4,
  CONTACT: 5,
  UNIT: 6,
  CATEGORY: 7,
  REASON: 8,
  REASON_OTHER: 9,
  TOWER: 10,
  PHOTO: 11,
  RECORD_ID: 12
};

// Cache idempotency window.
// Worker normally retries within this period.
const SYNC_DONE_TTL_SEC = 21600; // 6 hours


// ==================================================
// OUTPUT / BASIC HELPERS
// ==================================================
function output_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function toText_(value) {
  return value == null ? "" : String(value).trim();
}

function normalizeText_(value) {
  const s = toText_(value);
  if (!s) return "";

  return s
    .normalize("NFKD")
    .replace(/[^\w\s\/\-\&\(\)@']/g, "")
    .replace(/_/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function normalizeCategory_(value) {
  const v = normalizeText_(value);
  if (v === "OWNER" || v === "TENANT") return v;
  return "";
}

function formatPhone_(value) {
  let digits = toText_(value).replace(/[^0-9]/g, "");
  if (digits && !digits.startsWith("0")) digits = "0" + digits;
  return digits ? ("'" + digits) : "";
}

function formatTimestamp_(createdAt) {
  let dt = null;

  if (createdAt) {
    const parsed = new Date(createdAt);
    if (!isNaN(parsed.getTime())) dt = parsed;
  }

  if (!dt) dt = new Date();

  return Utilities.formatDate(dt, TIMEZONE, "d/M/yyyy HH:mm:ss");
}

function escapeForFormula_(value) {
  return toText_(value).replace(/"/g, '""');
}

function isHttpUrl_(value) {
  return /^https?:\/\//i.test(toText_(value));
}

function extractDriveFileId_(url) {
  const s = toText_(url);
  if (!s) return "";

  let m = s.match(/[?&]id=([\w-]+)/i);
  if (m && m[1]) return m[1];

  m = s.match(/\/d\/([\w-]+)/i);
  return m && m[1] ? m[1] : "";
}


// ==================================================
// SHEET HELPERS
// ==================================================
function getSheet_() {
  const spreadsheetId = toText_(SPREADSHEET_ID);

  if (!spreadsheetId || spreadsheetId.indexOf("PASTE_") === 0) {
    throw new Error("SPREADSHEET_ID has not been configured.");
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("Sheet not found: " + SHEET_NAME);
  }

  return sheet;
}

function ensureHeaders_(sheet) {
  const width = HEADERS.length;

  if (sheet.getMaxColumns() < width) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      width - sheet.getMaxColumns()
    );
  }

  const current = sheet.getRange(1, 1, 1, width).getDisplayValues()[0];
  const hasAnyHeader = current.some(v => toText_(v) !== "");

  if (!hasAnyHeader) {
    sheet.getRange(1, 1, 1, width).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }

  for (let i = 0; i < width; i++) {
    if (toText_(current[i]).toUpperCase() !== HEADERS[i]) {
      throw new Error(
        "Sheet header mismatch at column " +
        String.fromCharCode(65 + i) +
        ". Expected: " + HEADERS[i] +
        " | Found: " + (current[i] || "(blank)")
      );
    }
  }
}

function findRecordRowById_(sheet, recordId) {
  const id = toText_(recordId);
  const lastRow = sheet.getLastRow();
  if (!id || lastRow <= 1) return 0;

  const match = sheet
    .getRange(2, COL.RECORD_ID, lastRow - 1, 1)
    .createTextFinder(id)
    .matchEntireCell(true)
    .findNext();

  return match ? match.getRow() : 0;
}


// ==================================================
// DRIVE HELPERS
// ==================================================
function getDriveFolder_() {
  const folderId = toText_(DRIVE_FOLDER_ID);

  if (!folderId || folderId.indexOf("PASTE_") === 0) {
    throw new Error("DRIVE_FOLDER_ID has not been configured.");
  }

  return DriveApp.getFolderById(folderId);
}

function fetchAndStoreImage_(imageViewUrl, data, recordId) {
  const url = toText_(imageViewUrl);
  if (!url) {
    return { driveFileId: "", driveUrl: "", createdNew: false };
  }

  if (!isHttpUrl_(url)) {
    throw new Error("Invalid imageViewUrl.");
  }

  const safeName = toText_(data && data.namePassport).replace(/[^\w-]/g, "_") || "PHOTO";
  const safeRecordId = toText_(recordId).replace(/[^\w-]/g, "_");
  const filename = safeName + "_" + safeRecordId + ".jpg";
  const folder = getDriveFolder_();
  const existingFiles = folder.getFilesByName(filename);

  while (existingFiles.hasNext()) {
    const existingFile = existingFiles.next();
    if (existingFile.isTrashed()) continue;

    const existingFileId = existingFile.getId();
    return {
      driveFileId: existingFileId,
      driveUrl: "https://drive.google.com/uc?export=view&id=" + existingFileId,
      createdNew: false
    };
  }

  const response = UrlFetchApp.fetch(url, {
    method: "get",
    followRedirects: true,
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();

  if (status < 200 || status >= 300) {
    throw new Error("Image fetch failed (" + status + ")");
  }

  const blob = response.getBlob().setName(filename);
  const file = folder.createFile(blob);

  const driveFileId = file.getId();
  const driveUrl =
    "https://drive.google.com/uc?export=view&id=" + driveFileId;

  return {
    driveFileId: driveFileId,
    driveUrl: driveUrl,
    createdNew: true
  };
}


// ==================================================
// 90-DAY SHEET RETENTION
// - Registration mirror only.
// - Owner directory is in D1 unit_owners and is NOT handled by GAS.
// ==================================================
function parseSheetTimestamp_(value) {
  const s = toText_(value);
  if (!s) return null;

  const m = s.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})/
  );

  if (!m) return null;

  const dt = new Date(
    Number(m[3]),
    Number(m[2]) - 1,
    Number(m[1]),
    Number(m[4]),
    Number(m[5]),
    Number(m[6])
  );

  return isNaN(dt.getTime()) ? null : dt;
}

function cleanupByAge_(sheet, retentionDays) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  const cutoff = new Date(
    Date.now() - Number(retentionDays || RETENTION_DAYS) * 86400000
  );

  // Rows are appended chronologically, so expired rows form one block at the top.
  const firstTimestamp = sheet
    .getRange(2, COL.TIMESTAMP)
    .getDisplayValue();

  const firstDate = parseSheetTimestamp_(firstTimestamp);

  if (!firstDate || firstDate >= cutoff) {
    return false;
  }

  const rowCount = lastRow - 1;
  const timestamps = sheet
    .getRange(2, COL.TIMESTAMP, rowCount, 1)
    .getDisplayValues();

  let expiredRowCount = 0;

  for (let i = 0; i < rowCount; i++) {
    const dt = parseSheetTimestamp_(timestamps[i][0]);

    if (dt && dt < cutoff) {
      expiredRowCount++;
    } else {
      break;
    }
  }

  if (expiredRowCount <= 0) return false;

  const photoValues = sheet
    .getRange(2, COL.PHOTO, expiredRowCount, 1)
    .getValues();

  const nameFormulas = sheet
    .getRange(2, COL.NAME, expiredRowCount, 1)
    .getFormulas();

  const safeRows = [];

  for (let i = 0; i < expiredRowCount; i++) {
    const photoValue = toText_(photoValues[i][0]);
    const formula = toText_(nameFormulas[i][0]);
    const hyperlinkMatch = formula.match(/HYPERLINK\(\s*"([^"]+)"/i);

    let fileId = extractDriveFileId_(photoValue);

    if (!fileId && hyperlinkMatch && hyperlinkMatch[1]) {
      fileId = extractDriveFileId_(hyperlinkMatch[1]);
    }

    const hasPhotoReference = !!photoValue || !!hyperlinkMatch;

    // Rows with no Drive/photo reference can be removed immediately.
    if (!hasPhotoReference) {
      safeRows.push(i + 2);
      continue;
    }

    // Fail safe: keep the Sheet row if a photo reference exists but the
    // Drive file ID cannot be resolved.
    if (!fileId) {
      continue;
    }

    try {
      const file = DriveApp.getFileById(fileId);

      if (!file.isTrashed()) {
        file.setTrashed(true);
      }

      // Only remove the Sheet row after Drive confirms the file is trashed.
      if (file.isTrashed()) {
        safeRows.push(i + 2);
      }
    } catch (err) {
      // Keep the Sheet row so a later cleanup run can retry.
    }
  }

  if (!safeRows.length) return false;

  // Delete confirmed-safe Sheet rows from bottom to top in contiguous groups
  // so row-number shifts cannot cause the wrong row to be removed.
  let groupEnd = safeRows[safeRows.length - 1];
  let groupStart = groupEnd;

  for (let i = safeRows.length - 2; i >= -1; i--) {
    const row = i >= 0 ? safeRows[i] : 0;

    if (row && row === groupStart - 1) {
      groupStart = row;
      continue;
    }

    sheet.deleteRows(groupStart, groupEnd - groupStart + 1);

    if (row) {
      groupStart = row;
      groupEnd = row;
    }
  }

  return true;
}


// ==================================================
// MANUAL SETUP HELPER
// Run once after filling SPREADSHEET_ID.
// It only prepares/verifies the header row.
// ==================================================
function setupOwnerTenantSheet() {
  const sheet = getSheet_();
  ensureHeaders_(sheet);

  return {
    ok: true,
    sheet: SHEET_NAME,
    columns: HEADERS.length
  };
}


// ==================================================
// OPTIONAL HEALTH CHECK
// Open the deployed Web App URL in browser.
// ==================================================
function doGet() {
  return output_({
    ok: true,
    service: "SECURE ENTRY OWNER | TENANT - SENSORY GAS",
    sheet: SHEET_NAME,
    retentionDays: RETENTION_DAYS
  });
}


// ==================================================
// doPost - Worker -> GAS
//
// SYNC payload:
// {
//   token,
//   action: "SYNC",
//   id,
//   createdAt,
//   clientTxnId,
//   deviceId,
//   namePassport,
//   mykadPassport,
//   regnum,
//   contact,
//   unitNumber,
//   category,
//   reason,
//   reasonOther,
//   tower,
//   imageViewUrl
// }
//
// DELETE_DRIVE payload:
// {
//   token,
//   action: "DELETE_DRIVE",
//   fileIds: [...]
// }
// ==================================================
function doPost(e) {
  let data = {};
  let newDriveFileId = "";
  let newDriveUrl = "";
  let createdDriveFileId = "";
  let rowCommitted = false;
  let lock = null;
  let lockTaken = false;

  try {
    try {
      data = JSON.parse(
        e && e.postData && e.postData.contents
          ? e.postData.contents
          : "{}"
      );
    } catch (jsonErr) {
      return output_({
        success: false,
        error: true,
        retryable: false,
        message: "Invalid JSON: " + jsonErr.message
      });
    }

    const syncToken = toText_(
      PropertiesService.getScriptProperties().getProperty("SYNC_TOKEN")
    );

    if (!syncToken || toText_(data.token) !== syncToken) {
      return output_({
        success: false,
        error: true,
        retryable: false,
        message: "Unauthorized"
      });
    }

    const action = toText_(data.action || "SYNC").toUpperCase();

    // ==================================================
    // DELETE_DRIVE
    // Used by Cloudflare Worker retention cleanup.
    // ==================================================
    if (action === "DELETE_DRIVE") {
      const ids = Array.isArray(data.fileIds) ? data.fileIds : [];
      const confirmedIds = [];
      const failedIds = [];

      for (let i = 0; i < ids.length; i++) {
        const fileId = toText_(ids[i]);
        if (!fileId) continue;

        try {
          const file = DriveApp.getFileById(fileId);
          if (!file.isTrashed()) {
            file.setTrashed(true);
          }

          if (file.isTrashed()) {
            confirmedIds.push(fileId);
          } else {
            failedIds.push(fileId);
          }
        } catch (err) {
          failedIds.push(fileId);
        }
      }

      return output_({
        success: true,
        action: "DELETE_DRIVE",
        deletedCount: confirmedIds.length,
        confirmedIds: confirmedIds,
        failedIds: failedIds
      });
    }

    // ==================================================
    // CLEANUP_AGE
    // Optional manual/on-demand action.
    // ==================================================
    if (action === "CLEANUP_AGE") {
      lock = LockService.getScriptLock();

      try {
        lock.waitLock(15000);
        lockTaken = true;
      } catch (lockErr) {
        return output_({
          success: false,
          error: true,
          retryable: true,
          message: "Server busy: cleanup lock timeout"
        });
      }

      const sheet = getSheet_();
      ensureHeaders_(sheet);

      const cleaned = cleanupByAge_(sheet, RETENTION_DAYS);

      return output_({
        success: true,
        action: "CLEANUP_AGE",
        cleaned: !!cleaned
      });
    }

    if (action !== "SYNC") {
      return output_({
        success: false,
        error: true,
        retryable: false,
        message: "Unsupported action: " + action
      });
    }

    // ==================================================
    // Basic clean-field validation.
    // Worker already validates these, but GAS verifies again.
    // ==================================================
    const recordId = toText_(data.id);
    const namePassport = normalizeText_(data.namePassport);
    const mykadPassport = normalizeText_(data.mykadPassport);
    const regnum = normalizeText_(data.regnum);
    const contact = formatPhone_(data.contact);
    const unitNumber = normalizeText_(data.unitNumber);
    const category = normalizeCategory_(data.category);
    const reason = normalizeText_(data.reason);
    const reasonOther =
      reason === "OTHER"
        ? normalizeText_(data.reasonOther)
        : "";
    const tower = normalizeText_(data.tower);

    // DEFAULTER mode is an audit-only registration. The Worker intentionally
    // sends NAME and MYKAD/PASSPORT blank for these records, while NORMAL
    // registrations continue to require both identity fields. Old NORMAL
    // records that used REASON = DEFAULTER still contain identity fields and
    // therefore continue through the original NORMAL validation path.
    const isDefaulterAudit =
      reason === "DEFAULTER" && !namePassport && !mykadPassport;

    if (!recordId) {
      return output_({
        success: false,
        error: true,
        retryable: false,
        message: "Missing record id."
      });
    }

    // Shared fields remain mandatory for NORMAL and DEFAULTER.
    if (!regnum || !contact || !unitNumber) {
      return output_({
        success: false,
        error: true,
        retryable: false,
        message: "Missing required registration field."
      });
    }

    // NORMAL keeps the existing identity requirement unchanged. DEFAULTER
    // audit rows are allowed to mirror to Sheet with blank NAME / MYKAD.
    if (!isDefaulterAudit && (!namePassport || !mykadPassport)) {
      return output_({
        success: false,
        error: true,
        retryable: false,
        message: "Missing required registration field."
      });
    }

    if (category !== "OWNER" && category !== "TENANT") {
      return output_({
        success: false,
        error: true,
        retryable: false,
        message: "Invalid category."
      });
    }

    if (!reason) {
      return output_({
        success: false,
        error: true,
        retryable: false,
        message: "Reason is required."
      });
    }

    if (reason === "OTHER" && !reasonOther) {
      return output_({
        success: false,
        error: true,
        retryable: false,
        message: "Reason Other is required."
      });
    }

    if (tower !== "TOWER A" && tower !== "TOWER B") {
      return output_({
        success: false,
        error: true,
        retryable: false,
        message: "Tower must be TOWER A or TOWER B."
      });
    }

    // ==================================================
    // Idempotency
    // Prevent duplicate Sheet rows during normal Worker retries.
    // ==================================================
    const cache = CacheService.getScriptCache();
    const doneKey = "SYNCED_OWNER_TENANT_" + recordId;
    const doneRaw = cache.get(doneKey);

    if (doneRaw) {
      let done = {};
      try {
        done = JSON.parse(doneRaw);
      } catch (ignore) {}

      return output_({
        success: true,
        duplicate: true,
        driveFileId: toText_(done.driveFileId),
        driveUrl: toText_(done.driveUrl)
      });
    }

    // ==================================================
    // Critical section: durable duplicate check + Drive + Sheet commit.
    // ==================================================
    lock = LockService.getScriptLock();

    try {
      lock.waitLock(15000);
      lockTaken = true;
    } catch (lockErr) {
      throw new Error("Server busy: sync lock timeout");
    }

    const sheet = getSheet_();
    ensureHeaders_(sheet);

    const existingRow = findRecordRowById_(sheet, recordId);
    if (existingRow) {
      const existingDriveUrl = toText_(
        sheet.getRange(existingRow, COL.PHOTO).getValue()
      );

      return output_({
        success: true,
        duplicate: true,
        driveFileId: extractDriveFileId_(existingDriveUrl),
        driveUrl: existingDriveUrl
      });
    }

    // If no imageViewUrl, this is a reused-photo registration and PHOTO LINK
    // remains blank. Existing active deterministic files are reused safely.
    const imageResult = fetchAndStoreImage_(
      data.imageViewUrl,
      {
        namePassport: namePassport,
        regnum: regnum
      },
      recordId
    );

    newDriveFileId = imageResult.driveFileId;
    newDriveUrl = imageResult.driveUrl;
    if (imageResult.createdNew) createdDriveFileId = newDriveFileId;

    const hasNewPhoto = !!newDriveUrl;

    const nameCell = hasNewPhoto
      ? '=HYPERLINK("' +
        escapeForFormula_(newDriveUrl) +
        '","' +
        escapeForFormula_(namePassport) +
        '")'
      : namePassport;

    const rowValues = [[
      formatTimestamp_(data.createdAt), // A TIMESTAMP
      nameCell,                         // B NAME
      mykadPassport,                    // C MYKAD/PASSPORT
      regnum,                           // D REG.NUM
      contact,                          // E CONTACT
      unitNumber,                       // F UNIT NUMBER
      category,                         // G CATEGORY
      reason,                           // H REASON
      reasonOther,                      // I REASON OTHER
      tower,                            // J TOWER
      hasNewPhoto ? newDriveUrl : "",   // K PHOTO LINK
      recordId                          // L RECORD ID
    ]];

    const newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, HEADERS.length).setValues(rowValues);
    rowCommitted = true;

    let cleaned = false;
    let warning = "";

    try {
      cleaned = cleanupByAge_(sheet, RETENTION_DAYS);
    } catch (cleanupErr) {
      // The row is already safely written.
      // Do not return failure because Worker retry could create a duplicate row.
      warning =
        "Sheet row saved; retention cleanup warning: " +
        (cleanupErr.message || String(cleanupErr));
    }

    // Mark idempotency only after the Sheet row exists.
    try {
      cache.put(
        doneKey,
        JSON.stringify({
          driveFileId: newDriveFileId,
          driveUrl: newDriveUrl
        }),
        SYNC_DONE_TTL_SEC
      );
    } catch (cacheErr) {
      // Best effort only.
    }

    return output_({
      success: true,
      cleaned: !!cleaned,
      driveFileId: newDriveFileId,
      driveUrl: newDriveUrl,
      warning: warning
    });

  } catch (err) {
    // If Drive upload succeeded but Sheet append failed,
    // trash the new file so there is no orphan Drive image.
    if (createdDriveFileId && !rowCommitted) {
      try {
        DriveApp.getFileById(createdDriveFileId).setTrashed(true);
      } catch (cleanupErr) {}
    }

    return output_({
      success: false,
      error: true,
      retryable: true,
      message: err && err.message ? err.message : String(err)
    });

  } finally {
    if (lockTaken && lock) {
      try {
        lock.releaseLock();
      } catch (ignore) {}
    }
  }
}
