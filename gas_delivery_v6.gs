// ==================================================
// SECURE ENTRY - SENSORY (BACKEND - SEARCH + SYNC MIRROR)
// ==================================================
// doPost: SYNC from Worker
//   - NEW (NO_RECORD / EXPIRED): imageViewUrl -> Bearer-auth fetch image -> Drive
//   - Append row
//   - ✅ Incremental cache update (avoid rebuild penalty)
//
// doGet : TURBO search
//   - Compact cache index: KEY -> anyRow(latest) + actRow(proof photo)
//   - ACTIVE (FOUND) = actRow exists AND proof photo link exists AND within retention days
//   - EXPIRED = anyRow exists BUT no proof photo (or proof outside retention window)
// ==================================================

const SHEET_NAME = "SENSORY";
// If this Apps Script is standalone (not bound to the Sheet), set SPREADSHEET_ID.
// Optional override: set Script Property "SPREADSHEET_ID".
const SPREADSHEET_ID_FALLBACK = "1lfThC9DkZnCF0NWW0wJ9FP2W0DeTy8qpTPgD1dKil-I";
const DRIVE_FOLDER_ID = "1lrjbyVWGcBCEQE5vc08qPty14Yn5-HaI";

// Cache keys (compact string - avoid size limit)
const KEY_REG  = "IDX_REG_V2";
const KEY_ID   = "IDX_ID_V2";
const KEY_META = "IDX_META_V2";
const CACHE_TTL = 3600; // seconds
const RETENTION_DAYS = 90; // Auto-delete data + Drive photo after 90 days
const SYNC_META_SHEET_NAME = "SYNC META"; // hidden helper sheet for durable idempotency metadata

const REBUILD_COOLDOWN_MS = 30000; // prevent rebuild storms (30s)

// ==================================================
// Chunked Cache Helpers (avoid CacheService value size limit)
// - Stores large index strings in multiple cache keys:
//   <base>_N = number of chunks
//   <base>_0.. <base>_(N-1) = chunk payloads
// ==================================================
const CACHE_CHUNK_SIZE = 80000; // chars (safe under CacheService per-key limits)

function cacheGetChunked_(cache, baseKey) {
  const nStr = cache.get(baseKey + "_N");
  if (nStr === null) return null; // cache miss
  const n = parseInt(nStr || "0", 10) || 0;
  if (n <= 0) return "";
  let out = "";
  for (let i = 0; i < n; i++) {
    const part = cache.get(baseKey + "_" + i);
    if (part === null) return null; // incomplete -> treat as miss
    out += part;
  }
  return out;
}

function cachePutChunked_(cache, baseKey, value, ttlSec) {
  value = value || "";
  const ttl = ttlSec || CACHE_TTL;

  const oldNStr = cache.get(baseKey + "_N");
  const oldN = oldNStr === null ? 0 : (parseInt(oldNStr || "0", 10) || 0);

  const n = Math.max(1, Math.ceil(value.length / CACHE_CHUNK_SIZE));

  // Write chunks
  for (let i = 0; i < n; i++) {
    const start = i * CACHE_CHUNK_SIZE;
    const chunk = value.substring(start, start + CACHE_CHUNK_SIZE);
    cache.put(baseKey + "_" + i, chunk, ttl);
  }

  // Remove any leftover old chunks
  if (oldN > n) {
    for (let i = n; i < oldN; i++) {
      cache.remove(baseKey + "_" + i);
    }
  }

  cache.put(baseKey + "_N", String(n), ttl);
}

function cacheRemoveChunked_(cache, baseKey) {
  const nStr = cache.get(baseKey + "_N");
  if (nStr !== null) {
    const n = parseInt(nStr || "0", 10) || 0;
    for (let i = 0; i < n; i++) {
      cache.remove(baseKey + "_" + i);
    }
  }
  cache.remove(baseKey + "_N");
  cache.remove(baseKey);
}

// ==================================================
// Cache meta signature
// - lastRow alone is NOT enough when sheet is full (rolling delete keeps lastRow constant)
// - we include last timestamp cell (col A) as a cheap change detector
// ==================================================
function metaSig_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return "0|";
  // Column A stores timestamp text
  const lastTs = sheet.getRange(lastRow, 1).getDisplayValue();
  return String(lastRow) + "|" + String(lastTs || "");
}

// ==================================================
// Helpers
// ==================================================
function output_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function toText_(v) {
  return (v == null ? "" : String(v)).trim();
}

function getSyncToken_() {
  return toText_(PropertiesService.getScriptProperties().getProperty("SYNC_TOKEN"));
}

function getImageViewToken_() {
  return toText_(PropertiesService.getScriptProperties().getProperty("IMAGE_VIEW_TOKEN"));
}

function ensureSyncMetaSheet_(mainSheet) {
  const ss = mainSheet.getParent();
  let metaSheet = ss.getSheetByName(SYNC_META_SHEET_NAME);
  if (!metaSheet) {
    metaSheet = ss.insertSheet(SYNC_META_SHEET_NAME);
  }

  const expectedHeaders = [["TIMESTAMP", "SYNC ID", "DRIVE FILE ID", "DRIVE URL"]];
  const currentHeaders = metaSheet.getRange(1, 1, 1, 4).getDisplayValues()[0].map(v => toText_(v).toUpperCase());

  const emptyHeader = currentHeaders.every(v => !v);
  const isOldOrder =
    currentHeaders[0] === "SYNC ID" &&
    currentHeaders[1] === "DRIVE FILE ID" &&
    currentHeaders[2] === "DRIVE URL" &&
    currentHeaders[3] === "TIMESTAMP";

  const isNewOrder =
    currentHeaders[0] === "TIMESTAMP" &&
    currentHeaders[1] === "SYNC ID" &&
    currentHeaders[2] === "DRIVE FILE ID" &&
    currentHeaders[3] === "DRIVE URL";

  if (emptyHeader) {
    metaSheet.getRange(1, 1, 1, 4).setValues(expectedHeaders);
  } else if (isOldOrder) {
    // One-time reorder of existing SYNC META data:
    // old = SYNC ID | DRIVE FILE ID | DRIVE URL | TIMESTAMP
    // new = TIMESTAMP | SYNC ID | DRIVE FILE ID | DRIVE URL
    const lastMetaRow = metaSheet.getLastRow();
    const oldRows = lastMetaRow > 1
      ? metaSheet.getRange(2, 1, lastMetaRow - 1, 4).getDisplayValues()
      : [];

    metaSheet.getRange(1, 1, 1, 4).setValues(expectedHeaders);

    if (oldRows.length) {
      const reordered = oldRows.map(r => [
        toText_(r[3]),
        toText_(r[0]),
        toText_(r[1]),
        toText_(r[2])
      ]);
      metaSheet.getRange(2, 1, reordered.length, 4).setValues(reordered);
    }
  } else if (!isNewOrder) {
    throw new Error("SYNC META header mismatch");
  }

  // One-time migration from the earlier J/K implementation.
  // Only remove J/K when their headers match exactly, so unrelated columns are never touched.
  if (mainSheet.getMaxColumns() >= 11) {
    const jHeader = toText_(mainSheet.getRange(1, 10).getDisplayValue()).toUpperCase();
    const kHeader = toText_(mainSheet.getRange(1, 11).getDisplayValue()).toUpperCase();

    if (jHeader === "SYNC ID" && kHeader === "DRIVE FILE ID") {
      const lastRow = mainSheet.getLastRow();
      if (lastRow > 1) {
        const legacy = mainSheet.getRange(2, 1, lastRow - 1, 11).getDisplayValues();
        const metaLastRow = metaSheet.getLastRow();
        const existingIds = new Set();
        if (metaLastRow > 1) {
          metaSheet.getRange(2, 2, metaLastRow - 1, 1).getDisplayValues().forEach(r => {
            const v = toText_(r[0]);
            if (v) existingIds.add(v);
          });
        }

        const rowsToAdd = [];
        legacy.forEach(r => {
          const syncId = toText_(r[9]);
          if (!syncId || existingIds.has(syncId)) return;
          const driveFileId = toText_(r[10]);
          const driveUrl = toText_(r[8]);
          const ts = toText_(r[0]);
          rowsToAdd.push([ts, syncId, driveFileId, driveUrl]);
          existingIds.add(syncId);
        });

        if (rowsToAdd.length) {
          metaSheet.getRange(metaSheet.getLastRow() + 1, 1, rowsToAdd.length, 4).setValues(rowsToAdd);
        }
      }

      mainSheet.deleteColumns(10, 2);
    }
  }

  try { metaSheet.hideSheet(); } catch (hideErr) {}
  return metaSheet;
}

function findSyncRecordById_(metaSheet, reqId) {
  const id = toText_(reqId);
  const lastRow = metaSheet.getLastRow();
  if (!id || lastRow <= 1) return null;

  const match = metaSheet
    .getRange(2, 2, lastRow - 1, 1)
    .createTextFinder(id)
    .matchEntireCell(true)
    .findNext();

  if (!match) return null;

  const row = match.getRow();
  const vals = metaSheet.getRange(row, 1, 1, 4).getDisplayValues()[0];
  const driveFileId = toText_(vals[2]);
  let driveUrl = toText_(vals[3]);
  if (!driveUrl && driveFileId) {
    driveUrl = "https://drive.google.com/uc?export=view&id=" + driveFileId;
  }

  return {
    row: row,
    driveFileId: driveFileId,
    driveUrl: driveUrl
  };
}

function appendSyncMeta_(metaSheet, reqId, driveFileId, driveUrl, timestampText) {
  const row = metaSheet.getLastRow() + 1;
  metaSheet.getRange(row, 1, 1, 4).setValues([[
    toText_(timestampText),
    toText_(reqId),
    toText_(driveFileId),
    toText_(driveUrl)
  ]]);
  return row;
}

function cleanupSyncMetaByAge_(metaSheet, retentionDays) {
  const lastRow = metaSheet.getLastRow();
  if (lastRow <= 1) return false;

  const cutoff = new Date(Date.now() - (Number(retentionDays || 90) * 86400000));
  const vals = metaSheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
  const rowsToDelete = [];

  for (let i = 0; i < vals.length; i++) {
    const dt = parseTimestampDMY_(vals[i][0]);
    const driveFileId = toText_(vals[i][2]);

    // RET-01: old metadata without a Drive file can expire normally.
    // If a Drive file id exists, keep this row as a recovery anchor until
    // DELETE_DRIVE confirms the physical Drive deletion.
    if (dt && dt < cutoff && !driveFileId) rowsToDelete.push(i + 2);
  }

  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    metaSheet.deleteRow(rowsToDelete[i]);
  }

  return rowsToDelete.length > 0;
}

function deleteSyncMetaById_(metaSheet, reqId) {
  const id = toText_(reqId);
  const lastRow = metaSheet.getLastRow();
  if (!id || lastRow <= 1) return false;

  const match = metaSheet
    .getRange(2, 2, lastRow - 1, 1)
    .createTextFinder(id)
    .matchEntireCell(true)
    .findNext();

  if (!match) return false;
  metaSheet.deleteRow(match.getRow());
  return true;
}

function escapeForFormula_(s) {
  // For formula strings like =HYPERLINK("url","text")
  // Double quotes must be escaped by doubling them.
  return toText_(s).replace(/"/g, '""');
}


function normKey_(s) {
  return toText_(s).toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function normalizeText(str) {
  str = toText_(str);
  if (!str) return "";
  return str
    .normalize("NFKD")
    .replace(/[^\w\s\/\-\&\(\)]/g, "")
    .replace(/_/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function formatPhone(contact) {
  let phone = toText_(contact);
  if (phone && !phone.startsWith("0")) phone = "0" + phone;
  if (phone) phone = "'" + phone;
  return phone.toUpperCase();
}

function formatDateTimeDMY(dateObj) {
  const d = dateObj.getDate();
  const m = dateObj.getMonth() + 1;
  const y = dateObj.getFullYear();
  const h = String(dateObj.getHours()).padStart(2, "0");
  const min = String(dateObj.getMinutes()).padStart(2, "0");
  const s = String(dateObj.getSeconds()).padStart(2, "0");
  return d + "/" + m + "/" + y + " " + h + ":" + min + ":" + s;
}

// Canonical registration timestamp from Worker.
// Worker sends createdAt as an ISO timestamp. Use that original event time for
// Sheet + SYNC META retention; fall back to current time only when missing/invalid.
function resolveRegistrationDate_(createdAtRaw) {
  const raw = toText_(createdAtRaw);
  if (!raw) return new Date();

  const dt = new Date(raw);
  return isNaN(dt.getTime()) ? new Date() : dt;
}

function parseDMYDate_(dmyStr) {
  const s = toText_(dmyStr);
  const m = s.match(/^\s*(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}):(\d{2}))?/);
  if (!m) return null;
  const dd = parseInt(m[1], 10);
  const mm = parseInt(m[2], 10);
  const yy = parseInt(m[3], 10);
  const hh = parseInt(m[4] || "0", 10);
  const mi = parseInt(m[5] || "0", 10);
  const ss = parseInt(m[6] || "0", 10);
  if (!dd || !mm || !yy) return null;
  const dt = new Date(yy, mm - 1, dd, hh, mi, ss);
  return isNaN(dt.getTime()) ? null : dt;
}

function isExpiredByAge_(dmyStr, days) {
  const dt = parseDMYDate_(dmyStr);
  if (!dt) return false;
  const ageMs = Date.now() - dt.getTime();
  const maxMs = (days || 0) * 24 * 60 * 60 * 1000;
  return maxMs > 0 && ageMs > maxMs;
}


function getSheet_() {
  const propsId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  const id = toText_(propsId) || SPREADSHEET_ID_FALLBACK;
  // Prefer openById for Web App (standalone) reliability; fallback to active spreadsheet if available.
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Sheet not found: " + SHEET_NAME);
  return sheet;
}

function isHttpUrl_(s) {
  return /^https?:\/\//i.test(toText_(s));
}

function extractDriveFileId_(url) {
  // Expected: https://drive.google.com/uc?export=view&id=<FILEID>
  const m = toText_(url).match(/[?&]id=([\w-]+)/i);
  return m && m[1] ? m[1] : "";
}

// ==================================================
// ==================================================
// Age retention cleanup (RETENTION_DAYS)
// - Deletes oldest rows older than RETENTION_DAYS
// - Also trashes related Drive photo (col I or HYPERLINK in col B)
// - Designed to be cheap on submit: only runs full scan if oldest row is already older than cutoff
// ==================================================
function parseTimestampDMY_(tsText) {
  tsText = toText_(tsText);
  if (!tsText) return null;
  // Expected format: D/M/YYYY HH:MM:SS
  const m = tsText.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})/);
  if (!m) return null;
  const d = parseInt(m[1], 10);
  const mo = parseInt(m[2], 10) - 1;
  const y = parseInt(m[3], 10);
  const h = parseInt(m[4], 10);
  const mi = parseInt(m[5], 10);
  const s = parseInt(m[6], 10);
  const dt = new Date(y, mo, d, h, mi, s);
  return isNaN(dt.getTime()) ? null : dt;
}

const RETENTION_FULL_SCAN_INTERVAL_MS = 60 * 60 * 1000; // NEW-RET-04: full out-of-order sweep at least hourly
const RETENTION_LAST_FULL_SCAN_PROP = "RETENTION_LAST_FULL_SCAN_MS";

function deleteSheetRowsDescending_(sheet, rowNumbers) {
  if (!rowNumbers || !rowNumbers.length) return;

  // Delete lower rows first so earlier row numbers stay valid.
  const rows = rowNumbers.slice().sort((a, b) => b - a);
  let blockHigh = rows[0];
  let blockLow = rows[0];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row === blockLow - 1) {
      blockLow = row;
      continue;
    }

    sheet.deleteRows(blockLow, blockHigh - blockLow + 1);
    blockHigh = row;
    blockLow = row;
  }

  sheet.deleteRows(blockLow, blockHigh - blockLow + 1);
}

function cleanupByAge_(sheet, retentionDays, forceFullScan) {
  const result = {
    cleaned: false,
    rowsDeleted: 0,
    driveDeleteAttempted: 0,
    driveDeleteConfirmed: 0,
    driveDeleteFailed: 0
  };

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return result;

  const cutoff = new Date(Date.now() - (Number(retentionDays || 120) * 86400000));
  const nowMs = Date.now();
  const props = PropertiesService.getScriptProperties();

  // NEW-RET-04:
  // Rows are normally chronological, but a delayed sync can append an older
  // createdAt below newer rows. We therefore cannot assume all expired rows
  // form one contiguous block at the top of SENSORY.
  //
  // Keep the normal submit path light:
  // - if row 2 is already expired (or its timestamp is invalid), scan now;
  // - otherwise do a full order-independent sweep at least once per hour;
  // - CLEANUP_AGE can force an immediate full sweep.
  const firstTsText = sheet.getRange(2, 1).getDisplayValue();
  const firstDt = parseTimestampDMY_(firstTsText);
  const lastFullScanMs = parseInt(props.getProperty(RETENTION_LAST_FULL_SCAN_PROP) || "0", 10) || 0;
  const fullScanDue = (nowMs - lastFullScanMs) >= RETENTION_FULL_SCAN_INTERVAL_MS;
  const shouldFullScan = !!forceFullScan || !firstDt || firstDt < cutoff || fullScanDue;

  if (!shouldFullScan) return result;

  const numRows = lastRow - 1;
  const tsVals = sheet.getRange(2, 1, numRows, 1).getDisplayValues();
  const expiredIndexes = [];
  let invalidTimestampCount = 0;

  for (let i = 0; i < numRows; i++) {
    const dt = parseTimestampDMY_(tsVals[i][0]);
    if (!dt) {
      invalidTimestampCount++;
      continue;
    }
    if (dt < cutoff) expiredIndexes.push(i);
  }

  if (invalidTimestampCount > 0) {
    console.warn(
      "NEW-RET-04 retention sweep skipped " + invalidTimestampCount +
      " row(s) with invalid timestamp"
    );
  }

  if (!expiredIndexes.length) {
    // Successful order-independent sweep; throttle the next full scan.
    props.setProperty(RETENTION_LAST_FULL_SCAN_PROP, String(nowMs));
    return result;
  }

  // Read proof-photo columns only when at least one row is actually expired.
  const photoLinksI = sheet.getRange(2, 9, numRows, 1).getValues();
  const nameFormulasB = sheet.getRange(2, 2, numRows, 1).getFormulas();
  const rowsToDelete = [];

  // RET-02 remains intact:
  // - Sheet rows still expire at the retention boundary.
  // - Drive deletion failures are counted + logged instead of being swallowed.
  // - SYNC META remains the durable recovery anchor (RET-01) until Worker/GAS
  //   DELETE_DRIVE later confirms the physical Drive deletion.
  for (let x = 0; x < expiredIndexes.length; x++) {
    const i = expiredIndexes[x];
    const sheetRow = i + 2;
    rowsToDelete.push(sheetRow);

    let url = toText_(photoLinksI[i][0]);
    let fileId = extractDriveFileId_(url);

    if (!fileId) {
      const f = toText_(nameFormulasB[i][0]);
      const mm = f.match(/HYPERLINK\(\s*"([^"]+)"/i);
      if (mm && mm[1]) fileId = extractDriveFileId_(mm[1]);
    }

    if (fileId) {
      result.driveDeleteAttempted++;
      try {
        const file = DriveApp.getFileById(fileId);
        if (!file.isTrashed()) file.setTrashed(true);
        result.driveDeleteConfirmed++;
      } catch (err) {
        result.driveDeleteFailed++;
        console.warn(
          "RET-02 Drive cleanup pending for file " + fileId + ": " +
          toText_(err && err.message ? err.message : err).slice(0, 200)
        );
      }
    }
  }

  // Delete every expired row regardless of where it sits in the Sheet.
  // Descending contiguous blocks avoid row-number shift errors and reduce API calls.
  deleteSheetRowsDescending_(sheet, rowsToDelete);

  // Record the sweep only after row deletion succeeds.
  props.setProperty(RETENTION_LAST_FULL_SCAN_PROP, String(nowMs));

  result.cleaned = true;
  result.rowsDeleted = rowsToDelete.length;
  return result;
}

// NOTE:
// - MAX ROW strategy removed.
// - Expiry is determined by retention window (RETENTION_DAYS) + presence of proof photo link.

// ==================================================
// Compact cache line helpers
// Line format: KEY|anyRow(base36)|actRow(base36)\n
// Stored string is prefixed with '\n' so we can search safely with '\nKEY|'
// ==================================================
function normalizeCompact_(s) {
  s = s || "";
  if (!s) return "";
  return s.startsWith("\n") ? s : ("\n" + s);
}

function upsertCompactLine_(compactStr, key, anyRow, actRow) {
  const line = key + "|" + Number(anyRow || 0).toString(36) + "|" + Number(actRow || 0).toString(36) + "\n";

  compactStr = compactStr || "";
  if (!compactStr) return "\n" + line;

  compactStr = normalizeCompact_(compactStr);
  if (!compactStr.endsWith("\n")) compactStr += "\n";

  const needle = "\n" + key + "|";
  const pos = compactStr.indexOf(needle);
  if (pos === -1) return compactStr + line;

  const start = pos + 1; // actual line start (skip leading \n)
  const end = compactStr.indexOf("\n", start);
  if (end === -1) return compactStr.substring(0, start) + line;

  return compactStr.substring(0, start) + line + compactStr.substring(end + 1);
}

function resolveCompact_(target, compactStr) {
  compactStr = compactStr || "";
  if (!compactStr) return { status: "NO_RECORD", anyRow: 0, actRow: 0 };

  compactStr = normalizeCompact_(compactStr);

  const needle = "\n" + target + "|";
  const pos = compactStr.indexOf(needle);
  if (pos === -1) return { status: "NO_RECORD", anyRow: 0, actRow: 0 };

  const start = pos + 1;
  const end = compactStr.indexOf("\n", start);
  const line = (end === -1) ? compactStr.substring(start) : compactStr.substring(start, end);

  const parts = line.split("|");
  const anyRow = parseInt(parts[1] || "0", 36) || 0;
  const actRow = parseInt(parts[2] || "0", 36) || 0;

  if (!anyRow) return { status: "NO_RECORD", anyRow: 0, actRow: 0 };
  if (actRow) return { status: "ACTIVE", anyRow: anyRow, actRow: actRow };
  return { status: "EXPIRED", anyRow: anyRow, actRow: 0 };
}

function getIndex_() {
  const cache = CacheService.getScriptCache();
  const lock = LockService.getScriptLock();

  const sheet = getSheet_();
  const currentMeta = metaSig_(sheet);

  const regStr = cacheGetChunked_(cache, KEY_REG);
  const idStr = cacheGetChunked_(cache, KEY_ID);
  const metaStr = cache.get(KEY_META);
  if (regStr !== null && idStr !== null && metaStr === currentMeta) {
    return { reg: regStr || "", id: idStr || "", meta: metaStr || "" };
  }

  lock.waitLock(15000);
  try {
    const reg2 = cacheGetChunked_(cache, KEY_REG);
    const id2 = cacheGetChunked_(cache, KEY_ID);
    const meta2 = cache.get(KEY_META);
    if (reg2 !== null && id2 !== null && meta2 === currentMeta) {
      return { reg: reg2 || "", id: id2 || "", meta: meta2 || "" };
    }

    const idx = buildIndexCache_();
    cachePutChunked_(cache, KEY_REG, idx.reg || "", CACHE_TTL);
    cachePutChunked_(cache, KEY_ID, idx.id || "", CACHE_TTL);
    cache.put(KEY_META, currentMeta, CACHE_TTL);

    return { reg: idx.reg || "", id: idx.id || "", meta: currentMeta };
  } finally {
    lock.releaseLock();
  }
}
// ==================================================
// Force rebuild index cache (used when mismatch detected by Worker)
// - Clears cache keys and rebuilds compact index.
// - Uses script properties cooldown to avoid rebuild storms.
// ==================================================
function forceRebuildIndexCache_(sheet) {
  sheet = sheet || getSheet_();
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();

  // Fast cooldown check (no lock)
  const last = parseInt(props.getProperty("LAST_INDEX_REBUILD_MS") || "0", 10) || 0;
  if (now - last < REBUILD_COOLDOWN_MS) return { rebuilt: false, cooldown: true };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const last2 = parseInt(props.getProperty("LAST_INDEX_REBUILD_MS") || "0", 10) || 0;
    if (now - last2 < REBUILD_COOLDOWN_MS) return { rebuilt: false, cooldown: true };

    props.setProperty("LAST_INDEX_REBUILD_MS", String(now));

    const cache = CacheService.getScriptCache();
    cacheRemoveChunked_(cache, KEY_REG);
    cacheRemoveChunked_(cache, KEY_ID);
    cache.remove(KEY_META);

    const idx = buildIndexCache_();
    const metaNow = metaSig_(sheet);
    cachePutChunked_(cache, KEY_REG, idx.reg || "", CACHE_TTL);
    cachePutChunked_(cache, KEY_ID, idx.id || "", CACHE_TTL);
    cache.put(KEY_META, metaNow, CACHE_TTL);

    return { rebuilt: true, cooldown: false };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}



// ==================================================
// TURBO INDEX BUILDER (COMPACT CACHE)
// ACTIVE (FOUND) proof = Column I has photo link (http)
// ==================================================
function buildIndexCache_() {
  const sheet = getSheet_();

  const startRow = 2;
  const lastRow = sheet.getLastRow();
  const numRows = Math.max(0, lastRow - 1);
  if (numRows <= 0) return { reg: "", id: "" };

  // Single read (C..I) for speed: C=ID, D=REG, I=PHOTO
  // Range width 7 columns: 3..9
  const vals = sheet.getRange(startRow, 3, numRows, 7).getValues();

  const byReg = Object.create(null);
  const byId = Object.create(null);

  // Bottom-up: newest row wins for anyRow, and newest proof row wins for actRow
  for (let i = numRows - 1; i >= 0; i--) {
    const rowNum = startRow + i;

    const idKey = normKey_(vals[i][0]); // col C
    const regKey = normKey_(vals[i][1]); // col D
    const photo = toText_(vals[i][6]); // col I (index 6)
    const isActive = isHttpUrl_(photo);

    if (regKey) {
      if (!byReg[regKey]) byReg[regKey] = [rowNum, 0];
      if (isActive && byReg[regKey][1] === 0) byReg[regKey][1] = rowNum;
    }

    if (idKey) {
      if (!byId[idKey]) byId[idKey] = [rowNum, 0];
      if (isActive && byId[idKey][1] === 0) byId[idKey][1] = rowNum;
    }
  }

  function toCompactStr_(mapObj) {
    const keys = Object.keys(mapObj);
    if (!keys.length) return "";
    let out = "";
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      const anyRow = mapObj[k][0] || 0;
      const actRow = mapObj[k][1] || 0;
      out += k + "|" + anyRow.toString(36) + "|" + actRow.toString(36) + "\n";
    }
    return out ? "\n" + out : "";
  }

  return {
    reg: toCompactStr_(byReg),
    id: toCompactStr_(byId),
  };
}


// ==================================================
// doPost (SYNC dari Worker)
// Expect:
// { token, namePassport, mykadPassport, regnum, contact, remark, unitNumber, tower, reason, reasonOther, imageViewUrl }
// ==================================================
function doPost(e) {
  let data = {};
  let driveFileId = "";
  let photoUrl = "";
  let rowCommitted = false;
  let lock = null;
  let lockTaken = false;

  try {
    try {
      data = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : "{}");
    } catch (err) {
      return output_({ success: false, error: true, retryable: false, message: "Invalid JSON: " + err.message });
    }

    const syncToken = getSyncToken_();
    if (!syncToken) {
      return output_({ success: false, error: true, retryable: false, message: "SYNC_TOKEN is not configured" });
    }

    if (toText_(data.token) !== syncToken) {
      return output_({ success: false, error: true, retryable: false, message: "Unauthorized" });
    }

    const action = toText_(data.action).toUpperCase();

    // =========================
    // Action: DELETE_DRIVE (from Worker scheduled cleanup)
    // RET-01 payload (preferred):
    // { token, action:"DELETE_DRIVE", items:[{syncId, driveFileId, requireDrive}] }
    // Legacy { fileIds:[...] } remains supported.
    // =========================
    if (action === "DELETE_DRIVE") {
      lock = LockService.getScriptLock();
      try {
        lock.waitLock(30000);
        lockTaken = true;
      } catch (lockErr) {
        return output_({
          success: false,
          error: true,
          retryable: true,
          message: "Server busy: Drive cleanup lock timeout"
        });
      }

      const sheet = getSheet_();
      const syncMetaSheet = ensureSyncMetaSheet_(sheet);

      let items = [];
      if (data.items && Array.isArray(data.items)) {
        items = data.items;
      } else if (data.fileIds && Array.isArray(data.fileIds)) {
        items = data.fileIds.map(fileId => ({
          syncId: "",
          driveFileId: toText_(fileId),
          requireDrive: true
        }));
      }

      const confirmedIds = [];
      const failedIds = [];
      const failedDetails = [];

      for (let i = 0; i < items.length; i++) {
        const item = items[i] || {};
        const syncId = toText_(item.syncId);
        let driveFileId = toText_(item.driveFileId);
        const requireDrive = item.requireDrive === true;
        const responseId = syncId || driveFileId;

        let meta = null;
        if (syncId) {
          meta = findSyncRecordById_(syncMetaSheet, syncId);
          if (!driveFileId && meta) driveFileId = toText_(meta.driveFileId);
        }

        if (!responseId) continue;

        if (!driveFileId) {
          if (!requireDrive) {
            if (syncId) deleteSyncMetaById_(syncMetaSheet, syncId);
            confirmedIds.push(responseId);
          } else {
            failedIds.push(responseId);
            failedDetails.push({ id: responseId, reason: "DRIVE_FILE_ID_UNRESOLVED" });
          }
          continue;
        }

        try {
          const file = DriveApp.getFileById(driveFileId);
          if (!file.isTrashed()) file.setTrashed(true);

          // Remove durable metadata only after Drive deletion is confirmed.
          if (syncId) deleteSyncMetaById_(syncMetaSheet, syncId);
          confirmedIds.push(responseId);
        } catch (err) {
          failedIds.push(responseId);
          failedDetails.push({
            id: responseId,
            reason: "DRIVE_DELETE_FAILED",
            message: toText_(err && err.message ? err.message : err).slice(0, 200)
          });
        }
      }

      return output_({
        success: true,
        action: "DELETE_DRIVE",
        deletedCount: confirmedIds.length,
        confirmedIds: confirmedIds,
        failedIds: failedIds,
        failedDetails: failedDetails
      });
    }

    // =========================
    // Action: CLEANUP_AGE (on-demand retention cleanup)
    // Payload: { token, action:"CLEANUP_AGE" }
    // =========================
    if (action === "CLEANUP_AGE") {
      lock = LockService.getScriptLock();
      try {
        lock.waitLock(15000);
        lockTaken = true;
      } catch (lockErr) {
        return output_({ success: false, error: true, retryable: true, message: "Server busy: cleanup lock timeout" });
      }

      const sheet = getSheet_();
      const syncMetaSheet = ensureSyncMetaSheet_(sheet);
      const cleanupResult = cleanupByAge_(sheet, RETENTION_DAYS, true);
      const didMeta = cleanupSyncMetaByAge_(syncMetaSheet, RETENTION_DAYS);
      if (cleanupResult.cleaned) {
        const cache2 = CacheService.getScriptCache();
        cacheRemoveChunked_(cache2, KEY_REG);
        cacheRemoveChunked_(cache2, KEY_ID);
        cache2.remove(KEY_META);
      }
      return output_({
        success: true,
        action: "CLEANUP_AGE",
        cleaned: !!cleanupResult.cleaned,
        rowsDeleted: cleanupResult.rowsDeleted,
        driveDeleteAttempted: cleanupResult.driveDeleteAttempted,
        driveDeleteConfirmed: cleanupResult.driveDeleteConfirmed,
        driveDeleteFailed: cleanupResult.driveDeleteFailed,
        driveCleanupPending: cleanupResult.driveDeleteFailed > 0,
        syncMetaCleaned: !!didMeta
      });
    }

    const cache = CacheService.getScriptCache();

    // =========================
    // Durable idempotency
    // - Worker id is required for SYNC.
    // - Cache remains only a fast helper.
    // - Hidden SYNC META sheet is the durable source of truth.
    // =========================
    const reqId = toText_(data.id);
    if (!reqId) {
      return output_({ success: false, error: true, retryable: false, message: "Missing sync id" });
    }

    const doneKey = "SYNCED_ID_" + reqId;
    const doneRaw = cache.get(doneKey);
    if (doneRaw) {
      let doneObj = {};
      try { doneObj = JSON.parse(doneRaw || "{}"); } catch (e1) { doneObj = {}; }
      return output_({
        success: true,
        duplicate: true,
        driveFileId: toText_(doneObj.driveFileId),
        driveUrl: toText_(doneObj.driveUrl)
      });
    }

    // Acquire the lock BEFORE durable duplicate check and Drive upload.
    // This prevents two concurrent retries from both creating a Drive file / Sheet row.
    lock = LockService.getScriptLock();
    try {
      lock.waitLock(30000);
      lockTaken = true;
    } catch (lockErr) {
      throw new Error("Server busy: sync lock timeout");
    }

    const sheet = getSheet_();
    const syncMetaSheet = ensureSyncMetaSheet_(sheet);

    const existing = findSyncRecordById_(syncMetaSheet, reqId);
    if (existing) {
      try {
        cache.put(doneKey, JSON.stringify({
          driveFileId: existing.driveFileId,
          driveUrl: existing.driveUrl
        }), 21600);
      } catch (e2) {}

      return output_({
        success: true,
        duplicate: true,
        driveFileId: existing.driveFileId,
        driveUrl: existing.driveUrl
      });
    }

    // =========================
    // Photo strategy
    // - If imageViewUrl exists, it MUST be fetched successfully, otherwise FAIL (Worker will retry).
    // - Drive upload happens inside the same lock as the durable duplicate check.
    // =========================
    const imageViewUrl = toText_(data.imageViewUrl);

    if (imageViewUrl) {
      // NEW-SEC-04:
      // IMAGE_VIEW_TOKEN lives in Apps Script Script Properties and is sent
      // only in the Authorization header. This also remains compatible during
      // rollout with Worker v10 URLs that still contain the old query token.
      const imageViewToken = getImageViewToken_();
      if (!imageViewToken) {
        throw new Error("IMAGE_VIEW_TOKEN is not configured");
      }

      const resp = UrlFetchApp.fetch(imageViewUrl, {
        method: "get",
        headers: {
          Authorization: "Bearer " + imageViewToken
        },
        followRedirects: true,
        muteHttpExceptions: true,
      });

      const code = resp.getResponseCode();
      if (!(code >= 200 && code < 300)) {
        throw new Error("Image fetch failed (" + code + ")");
      }

      const driveFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const filename =
        (normalizeText(data.namePassport) || normalizeText(data.regnum) || "PHOTO") +
        "_" +
        Date.now() +
        ".jpg";

      const blob = resp.getBlob().setName(filename);
      const file = driveFolder.createFile(blob);
      driveFileId = file.getId();
      photoUrl = "https://drive.google.com/uc?export=view&id=" + driveFileId;
    }

    // NEW proof only when photoUrl exists
    const shouldCreateHyperlink = !!photoUrl;

    // Name (hyperlink only if NEW proof row created)
    const nameText = normalizeText(data.namePassport);
    const safeName = escapeForFormula_(nameText);
    const safePhotoUrl = escapeForFormula_(photoUrl);

    const nameCellValue = shouldCreateHyperlink
      ? '=HYPERLINK("' + safePhotoUrl + '","' + safeName + '")'
      : nameText;

    // Category (remark)
    let remarkValue = normalizeText(data.remark);
    const unitNumberValue = normalizeText(data.unitNumber);
    if ((remarkValue === "OWNER" || remarkValue === "TENANT") && unitNumberValue) {
      remarkValue = remarkValue + " ( " + unitNumberValue + " )";
    }

    // Reason
    let reasonValue = "";
    const mainReason = toText_(data.reason).toUpperCase();
    const reasonOther = toText_(data.reasonOther);
    if (mainReason === "OTHER" && reasonOther) {
      reasonValue = "OTHER ( " + reasonOther.toUpperCase() + " )";
    } else if (mainReason) {
      reasonValue = mainReason;
    }

    const registrationDate = resolveRegistrationDate_(data.createdAt);
    const photoLinkColI = shouldCreateHyperlink ? photoUrl : "";

    const timestampText = formatDateTimeDMY(registrationDate);
    const rowValues = [[
      timestampText,                     // A TIMESTAMP
      nameCellValue,                     // B NAME (formula OR text)
      normalizeText(data.mykadPassport), // C MYKAD/PASSPORT
      normalizeText(data.regnum),        // D REGNUM
      formatPhone(data.contact),         // E CONTACT
      remarkValue,                       // F CATEGORY
      normalizeText(data.tower),         // G TOWER
      reasonValue,                       // H REASON
      photoLinkColI                      // I PHOTO LINK (proof only)
    ]];

    const newRow = sheet.getLastRow() + 1;
    let mainRowWritten = false;
    try {
      sheet.getRange(newRow, 1, 1, 9).setValues(rowValues);
      mainRowWritten = true;

      // Durable idempotency metadata is kept outside the visible SENSORY table.
      appendSyncMeta_(syncMetaSheet, reqId, driveFileId, photoUrl, timestampText);
      rowCommitted = true;
    } catch (commitErr) {
      // If metadata commit fails, remove the just-written main row so a retry cannot create a duplicate.
      if (mainRowWritten) {
        try { sheet.deleteRow(newRow); } catch (rollbackErr) {}
      }
      throw commitErr;
    }

    let deleted = false;
    let warning = "";

    try {
      // =========================
      // Auto delete (rolling)
      // =========================
      const cleanupResult = cleanupByAge_(sheet, RETENTION_DAYS);
      deleted = !!cleanupResult.cleaned;
      cleanupSyncMetaByAge_(syncMetaSheet, RETENTION_DAYS);

      if (cleanupResult.driveDeleteFailed > 0) {
        warning = "Drive cleanup pending: " + cleanupResult.driveDeleteFailed + " file(s)";
      }

      // =========================
      // Cache handling
      // =========================
      if (deleted) {
        // Required by Edos: clear cache when autoDelete happens
        cacheRemoveChunked_(cache, KEY_REG);
        cacheRemoveChunked_(cache, KEY_ID);
        cache.remove(KEY_META);

        // Rebuild now so next search stays fast
        const idx = buildIndexCache_();
        const metaNow = metaSig_(sheet);
        cachePutChunked_(cache, KEY_REG, idx.reg || "", CACHE_TTL);
        cachePutChunked_(cache, KEY_ID, idx.id || "", CACHE_TTL);
        cache.put(KEY_META, metaNow, CACHE_TTL);
      } else {
        // Incremental update (no row shifting)
        let regStr = cacheGetChunked_(cache, KEY_REG);
        let idStr = cacheGetChunked_(cache, KEY_ID);

        // If cache missing, build once
        if (regStr === null || idStr === null) {
          const idx2 = buildIndexCache_();
          regStr = idx2.reg || "";
          idStr = idx2.id || "";
        }

        const regKey = normKey_(data.regnum);
        const idKey = normKey_(data.mykadPassport);

        // Keep old actRow unless this is a NEW proof row
        if (regKey) {
          const old = resolveCompact_(regKey, regStr);
          const actRow = shouldCreateHyperlink ? newRow : (old.actRow || 0);
          regStr = upsertCompactLine_(regStr, regKey, newRow, actRow);
        }

        if (idKey) {
          const old2 = resolveCompact_(idKey, idStr);
          const actRow2 = shouldCreateHyperlink ? newRow : (old2.actRow || 0);
          idStr = upsertCompactLine_(idStr, idKey, newRow, actRow2);
        }

        const metaNow2 = metaSig_(sheet);
        cachePutChunked_(cache, KEY_REG, regStr || "", CACHE_TTL);
        cachePutChunked_(cache, KEY_ID, idStr || "", CACHE_TTL);
        cache.put(KEY_META, metaNow2, CACHE_TTL);
      }
    } catch (cacheErr) {
      // Row is already safely written. Do not fail sync and cause duplicate retry.
      const cacheWarning = "Sheet row saved; cache cleanup/update warning: " + cacheErr.message;
      warning = warning ? (warning + " | " + cacheWarning) : cacheWarning;
    }

    // Fast helper only; durable authority is the hidden SYNC META sheet.
    try {
      cache.put(doneKey, JSON.stringify({
        driveFileId: driveFileId,
        driveUrl: photoUrl
      }), 21600);
    } catch (e3) {}

    return output_({
      success: true,
      deleted: !!deleted,
      driveFileId: driveFileId,
      driveUrl: photoUrl,
      warning: warning
    });
  } catch (err) {
    // Avoid orphaned Drive image if image uploaded but row not written.
    if (driveFileId && !rowCommitted) {
      try { DriveApp.getFileById(driveFileId).setTrashed(true); } catch (e4) {}
    }
    return output_({ success: false, error: true, retryable: true, message: err.message || String(err) });
  } finally {
    if (lockTaken && lock) {
      try { lock.releaseLock(); } catch (e5) {}
    }
  }
}

// ==================================================
// doGet (Search)
// Response format used by frontend:
// - { exist:false }
// - { exist:true, hasHyperlink:false, data:{} }  // EXPIRED
// - { exist:true, hasHyperlink:true, data:{... , photoLink:"..."} } // ACTIVE
//
// IMPORTANT concept:
// - data fields should come from LATEST row (anyRow)
// - photoLink should come from PROOF row (actRow)
// ==================================================
function doGet(e) {
  try {
    const p = (e && e.parameter) ? e.parameter : {};
    const value = toText_(p.value);
    const field = toText_(p.field);

    if (!value) return output_({ exist: false });

    const target = normKey_(value);

    // If Worker detects mismatch, it retries with forceRebuild=1 to refresh cache safely.
    let sheet = null;
    if (toText_(p.forceRebuild) === "1") {
      sheet = getSheet_();
      forceRebuildIndexCache_(sheet);
    }

    const idx = getIndex_();

    let result;
    if (field) {
      const f = normKey_(field);
      if (f === "REGNUM") result = resolveCompact_(target, idx.reg);
      else if (f === "MYKADPASSPORT") result = resolveCompact_(target, idx.id);
      else result = { status: "NO_RECORD", anyRow: 0, actRow: 0 };
    } else {
      result = resolveCompact_(target, idx.reg);
      if (result.status === "NO_RECORD") result = resolveCompact_(target, idx.id);
    }

    if (result.status === "NO_RECORD") return output_({ exist: false });

// If index says EXPIRED (no photo proof), return expired fast (no extra reads)
if (result.status === "EXPIRED") {
  return output_({ exist: true, hasHyperlink: false, data: {} });
}

    sheet = sheet || getSheet_();

// Latest details
const anyRow = result.anyRow;
const full = sheet.getRange(anyRow, 1, 1, 9).getValues()[0];

// Proof photo link
let photoLink = "";
const actRow = result.actRow;

// Fast path: if latest row itself has proof, reuse the same read (no extra getRange)
if (actRow && actRow === anyRow) {
  const maybe = toText_(full[8]);
  if (isHttpUrl_(maybe)) photoLink = maybe;
} else if (actRow) {
  const proofRow = sheet.getRange(actRow, 9, 1, 1).getValues()[0][0];
  const proofLink = toText_(proofRow);
  if (isHttpUrl_(proofLink)) photoLink = proofLink;
}

// Proof timestamp (same row as photoLink proof)
let proofTimestamp = "";
if (actRow && actRow === anyRow) {
  proofTimestamp = toText_(full[0]);
} else if (actRow) {
  const tsCell = sheet.getRange(actRow, 1, 1, 1).getValues()[0][0];
  proofTimestamp = toText_(tsCell);
}

// Enforce expiry by retention window (even if photoLink still exists in old rows)
if (!photoLink || isExpiredByAge_(proofTimestamp, RETENTION_DAYS)) {
  return output_({ exist: true, hasHyperlink: false, data: {} });
}

    return output_({
      exist: true,
      hasHyperlink: true,
      proofTimestamp: proofTimestamp,
      data: {
        namePassport: full[1] || "",
        mykadPassport: full[2] || "",
        regnum: full[3] || "",
        contact: full[4] || "",
        remark: full[5] || "",
        photoLink: photoLink
      }
    });
  } catch (err) {
    return output_({ error: true, message: err.message });
  }
}

// ==================================================
// Optional: manual cache reset (run once if needed)
// ==================================================
function resetIndexCache() {
  const cache = CacheService.getScriptCache();
  cacheRemoveChunked_(cache, KEY_REG);
  cacheRemoveChunked_(cache, KEY_ID);
  cache.remove(KEY_META);
  try { PropertiesService.getScriptProperties().deleteProperty("LAST_INDEX_REBUILD_MS"); } catch (e) {}
  return { ok: true };
}
