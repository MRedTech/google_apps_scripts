// ==================================================
// SECURE ENTRY - SENSORY (BACKEND - SEARCH + SYNC MIRROR)
// ==================================================
// doPost: SYNC from Worker
//   - NEW (NO_RECORD / EXPIRED): imageViewUrl -> fetch image -> Drive
//   - Append row
//   - ✅ Incremental cache update (avoid rebuild penalty)
//
// doGet : TURBO search
//   - Compact cache index: KEY -> anyRow(latest) + actRow(proof photo)
//   - ACTIVE (FOUND) = actRow exists AND proof photo link exists AND within retention days
//   - EXPIRED = anyRow exists BUT no proof photo (or proof outside retention window)
// ==================================================

const SHEET_NAME = "MREDTECH";
// If this Apps Script is standalone (not bound to the Sheet), set SPREADSHEET_ID.
// Optional override: set Script Property "SPREADSHEET_ID".
const SPREADSHEET_ID_FALLBACK = "1vD4CLiI1lPI65I2wvDkYlUw9L_hoe1gIPNpuCj7YrKg";
const DRIVE_FOLDER_ID = "1lrjbyVWGcBCEQE5vc08qPty14Yn5-HaI";
const SYNC_TOKEN = "SE_SYNC_9f3c1a7b2d4e6f8091ab3cd5ef678901R";

// Cache keys (compact string - avoid size limit)
const KEY_REG  = "IDX_REG_V2";
const KEY_ID   = "IDX_ID_V2";
const KEY_META = "IDX_META_V2";
const CACHE_TTL = 3600; // seconds
const RETENTION_DAYS = 1;

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

function cleanupByAge_(sheet, retentionDays) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;
  const cutoff = new Date(Date.now() - (Number(retentionDays || 120) * 86400000));

  // Quick check: if the oldest row is newer than cutoff, skip
  const firstTsText = sheet.getRange(2, 1).getDisplayValue();
  const firstDt = parseTimestampDMY_(firstTsText);
  if (!firstDt || firstDt >= cutoff) return false;

  // Read timestamps from top until we reach cutoff (oldest block is contiguous)
  const numRows = lastRow - 1;
  const tsVals = sheet.getRange(2, 1, numRows, 1).getDisplayValues();
  let rowsToDelete = 0;
  for (let i = 0; i < numRows; i++) {
    const dt = parseTimestampDMY_(tsVals[i][0]);
    if (dt && dt < cutoff) rowsToDelete++;
    else break;
  }
  if (rowsToDelete <= 0) return false;

  // Delete Drive photos for the rows being deleted
  const photoLinksI = sheet.getRange(2, 9, rowsToDelete, 1).getValues();
  const nameFormulasB = sheet.getRange(2, 2, rowsToDelete, 1).getFormulas();

  for (let i = 0; i < rowsToDelete; i++) {
    let url = toText_(photoLinksI[i][0]);
    let fileId = extractDriveFileId_(url);

    if (!fileId) {
      const f = toText_(nameFormulasB[i][0]);
      const mm = f.match(/HYPERLINK\(\s*"([^"]+)"/i);
      if (mm && mm[1]) fileId = extractDriveFileId_(mm[1]);
    }

    if (fileId) {
      try { DriveApp.getFileById(fileId).setTrashed(true); } catch (err) {}
    }
  }

  sheet.deleteRows(2, rowsToDelete);
  return true;
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

  const regStr = cache.get(KEY_REG);
  const idStr = cache.get(KEY_ID);
  const metaStr = cache.get(KEY_META);
  if (regStr !== null && idStr !== null && metaStr === currentMeta) {
    return { reg: regStr || "", id: idStr || "", meta: metaStr || "" };
  }

  lock.waitLock(15000);
  try {
    const reg2 = cache.get(KEY_REG);
    const id2 = cache.get(KEY_ID);
    const meta2 = cache.get(KEY_META);
    if (reg2 !== null && id2 !== null && meta2 === currentMeta) {
      return { reg: reg2 || "", id: id2 || "", meta: meta2 || "" };
    }

    const idx = buildIndexCache_();
    cache.put(KEY_REG, idx.reg || "", CACHE_TTL);
    cache.put(KEY_ID, idx.id || "", CACHE_TTL);
    cache.put(KEY_META, currentMeta, CACHE_TTL);

    return { reg: idx.reg || "", id: idx.id || "", meta: currentMeta };
  } finally {
    lock.releaseLock();
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
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = getSheet_();

    let data = {};
    try {
      data = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : "{}");
    } catch (err) {
      return output_({ success: false, error: true, message: "Invalid JSON: " + err.message });
    }

    if (toText_(data.token) !== SYNC_TOKEN) {
      return output_({ success: false, error: true, message: "Unauthorized" });
    }

    // =========================
    // Action: DELETE_DRIVE (from Worker scheduled cleanup)
    // Payload: { token, action:"DELETE_DRIVE", fileIds:[...] }
    // =========================
    const action = toText_(data.action).toUpperCase();
    if (action === "DELETE_DRIVE") {
      const ids = (data.fileIds && Array.isArray(data.fileIds)) ? data.fileIds : [];
      let deletedCount = 0;

      for (let i = 0; i < ids.length; i++) {
        const fileId = toText_(ids[i]);
        if (!fileId) continue;
        try {
          DriveApp.getFileById(fileId).setTrashed(true);
          deletedCount++;
        } catch (err) {
          // ignore per-file errors
        }
      }

      return output_({ success: true, action: "DELETE_DRIVE", deletedCount: deletedCount });
    }

    // =========================
    // Action: CLEANUP_AGE (on-demand retention cleanup)
    // Payload: { token, action:"CLEANUP_AGE" }
    // =========================
    if (action === "CLEANUP_AGE") {
      const did = cleanupByAge_(sheet, RETENTION_DAYS);
      if (did) {
        const cache2 = CacheService.getScriptCache();
        cache2.remove(KEY_REG);
        cache2.remove(KEY_ID);
        cache2.remove(KEY_META);
      }
      return output_({ success: true, action: "CLEANUP_AGE", cleaned: !!did });
    }

    const cache = CacheService.getScriptCache();

    // =========================
    // Idempotency (prevent duplicate rows on Worker retry)
    // =========================
    const reqId = toText_(data.id);
    const doneKey = reqId ? ("SYNCED_ID_" + reqId) : "";
    if (doneKey && cache.get(doneKey)) {
      return output_({ success: true, duplicate: true });
    }

    // =========================
    // Photo strategy
    // - If imageViewUrl exists, it MUST be fetched successfully, otherwise FAIL (Worker will retry)
    // - We only create proof (hyperlink + col I) when NEW photo is stored
    // =========================
    let photoUrl = "";
    let driveFileId = "";
    const imageViewUrl = toText_(data.imageViewUrl);

    if (imageViewUrl) {
      const resp = UrlFetchApp.fetch(imageViewUrl, {
        method: "get",
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

    const now = new Date();
    const photoLinkColI = shouldCreateHyperlink ? photoUrl : "";

    // =========================
    // Write row (single batch write for speed)
    // =========================
    const newRow = sheet.getLastRow() + 1;

    sheet.getRange(newRow, 1, 1, 9).setValues([[
      formatDateTimeDMY(now),            // A TIMESTAMP
      nameCellValue,                     // B NAME (formula OR text)
      normalizeText(data.mykadPassport), // C MYKAD/PASSPORT
      normalizeText(data.regnum),        // D REGNUM
      formatPhone(data.contact),         // E CONTACT
      remarkValue,                       // F CATEGORY
      normalizeText(data.tower),         // G TOWER
      reasonValue,                       // H REASON
      photoLinkColI                      // I PHOTO LINK (proof only)
    ]]);

    // =========================
    // Auto delete (rolling)
    // =========================
    const ageDeleted = cleanupByAge_(sheet, RETENTION_DAYS);
    // No MAX ROW rolling deletion: Google Sheet is not limited by row count.
    const deleted = ageDeleted;


    // =========================
    // Cache handling
    // =========================
    if (deleted) {
      // Required by Edos: clear cache when autoDelete happens
      cache.remove(KEY_REG);
      cache.remove(KEY_ID);
      cache.remove(KEY_META);

      // Rebuild now so next search stays fast
      const idx = buildIndexCache_();
      const metaNow = metaSig_(sheet);
      cache.put(KEY_REG, idx.reg || "", CACHE_TTL);
      cache.put(KEY_ID, idx.id || "", CACHE_TTL);
      cache.put(KEY_META, metaNow, CACHE_TTL);
    } else {
      // Incremental update (no row shifting)
      let regStr = cache.get(KEY_REG);
      let idStr = cache.get(KEY_ID);

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
      cache.put(KEY_REG, regStr || "", CACHE_TTL);
      cache.put(KEY_ID, idStr || "", CACHE_TTL);
      cache.put(KEY_META, metaNow2, CACHE_TTL);
    }

    // Mark idempotency key last (only when everything is OK)
    if (doneKey) cache.put(doneKey, "1", 21600); // 6 hours

    return output_({ success: true, deleted: !!deleted, driveFileId: driveFileId, driveUrl: photoUrl });
  } catch (err) {
    return output_({ success: false, error: true, message: err.message });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
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

const sheet = getSheet_();

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
  cache.remove(KEY_REG);
  cache.remove(KEY_ID);
  cache.remove(KEY_META);
  return { ok: true };
}
