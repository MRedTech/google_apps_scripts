// ==================================================
// SECURE ENTRY - SENSORY (BACKEND - SEARCH + SYNC MIRROR)
// ==================================================
// doPost: SYNC from Worker
//   - NEW (NO_RECORD / EXPIRED): imageViewUrl -> fetch image -> Drive
//   - FOUND: existingPhotoLink -> reuse (no fetch)
//   - Append row
//   - Auto-delete oldest rows beyond MAX_ROW (delete Drive file)
//   - ✅ Incremental cache update (avoid rebuild penalty)
//
// doGet : TURBO search
//   - Compact cache index: KEY -> anyRow(latest) + actRow(proof photo)
//   - ACTIVE (FOUND) = actRow exists within last MAX_ROW
//   - EXPIRED = anyRow exists BUT actRow=0 (proof photo pushed out/deleted)
// ==================================================

const SHEET_NAME = "SENSORY";
const DRIVE_FOLDER_ID = "1lrjbyVWGcBCEQE5vc08qPty14Yn5-HaI";
const MAX_ROW = 6000;
const SYNC_TOKEN = "SE_SYNC_9f3c1a7b2d4e6f8091ab3cd5ef678901R";

// Cache keys (compact string - avoid size limit)
const KEY_REG  = "IDX_REG_V2";
const KEY_ID   = "IDX_ID_V2";
const KEY_META = "IDX_META_V2";
const CACHE_TTL = 3600; // seconds

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

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
// Auto delete jika > MAX_ROW baris, dan padam gambar di Drive
// ==================================================
function autoDeleteRows_(sheet, maxRows) {
  const totalRows = sheet.getLastRow();
  if (totalRows <= maxRows + 1) return false;

  const rowsToDelete = totalRows - maxRows - 1; // keep header row
  if (rowsToDelete <= 0) return false;

  // Column I photo link (preferred source)
  // Column B hyperlink formula (fallback) - just in case column I was cleared but hyperlink still exists
  const photoLinksI = sheet.getRange(2, 9, rowsToDelete, 1).getValues();
  const nameFormulasB = sheet.getRange(2, 2, rowsToDelete, 1).getFormulas();

  for (let i = 0; i < rowsToDelete; i++) {
    let url = toText_(photoLinksI[i][0]);
    let fileId = extractDriveFileId_(url);

    if (!fileId) {
      const f = toText_(nameFormulasB[i][0]);
      // Extract the first URL inside HYPERLINK("<url>", ...)
      const m = f.match(/HYPERLINK\(\s*"([^"]+)"/i);
      if (m && m[1]) fileId = extractDriveFileId_(m[1]);
    }

    if (fileId) {
      try {
        DriveApp.getFileById(fileId).setTrashed(true);
      } catch (err) {
        // ignore
      }
    }
  }

  sheet.deleteRows(2, rowsToDelete);
  return true;
}

// ==================================================
// Compact cache line helpers
// Line format: KEY|anyRow(base36)|actRow(base36)\n
// ==================================================
function upsertCompactLine_(compactStr, key, anyRow, actRow) {
  const k = key + "|";
  const line = key + "|" + Number(anyRow || 0).toString(36) + "|" + Number(actRow || 0).toString(36) + "\n";

  compactStr = compactStr || "";
  const pos = compactStr.indexOf(k);
  if (pos === -1) return compactStr + line;

  const end = compactStr.indexOf("\n", pos);
  if (end === -1) return compactStr.substring(0, pos) + line;
  return compactStr.substring(0, pos) + line + compactStr.substring(end + 1);
}

function resolveCompact_(target, compactStr) {
  compactStr = compactStr || "";
  if (!compactStr) return { status: "NO_RECORD", anyRow: 0, actRow: 0 };

  const needle = target + "|";
  const pos = compactStr.indexOf(needle);
  if (pos === -1) return { status: "NO_RECORD", anyRow: 0, actRow: 0 };

  const end = compactStr.indexOf("\n", pos);
  const line = (end === -1) ? compactStr.substring(pos) : compactStr.substring(pos, end);
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
  const numRows = Math.max(0, Math.min(MAX_ROW, lastRow - 1));
  if (numRows <= 0) return { reg: "", id: "" };

  // Only read what we need (fast)
  const colC = sheet.getRange(startRow, 3, numRows, 1).getValues(); // MYKAD/PASSPORT
  const colD = sheet.getRange(startRow, 4, numRows, 1).getValues(); // REGNUM
  const colI = sheet.getRange(startRow, 9, numRows, 1).getValues(); // PHOTO LINK

  const byReg = {};
  const byId = {};

  // Bottom-up: newest row wins for anyRow, and newest proof row wins for actRow
  for (let i = numRows - 1; i >= 0; i--) {
    const rowNum = startRow + i;

    const regKey = normKey_(colD[i][0]);
    const idKey = normKey_(colC[i][0]);

    const photo = toText_(colI[i][0]);
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
    return out;
  }

  return {
    reg: toCompactStr_(byReg),
    id: toCompactStr_(byId),
  };
}

// ==================================================
// doPost (SYNC dari Worker)
// Expect:
// { token, namePassport, mykadPassport, regnum, contact, remark, unitNumber, tower, reason, reasonOther, imageViewUrl, existingPhotoLink }
// ==================================================
function doPost(e) {
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

    // ===== Photo strategy =====
    // Priority:
    // 1) If imageViewUrl exists & fetch success -> upload to Drive (NEW)
    // 2) Else if existingPhotoLink exists -> reuse proof photo (FOUND)
    let photoUrl = "";

    const imageViewUrl = toText_(data.imageViewUrl);
    const existingPhotoLink = toText_(data.existingPhotoLink);

    if (imageViewUrl) {
      const resp = UrlFetchApp.fetch(imageViewUrl, {
        method: "get",
        followRedirects: true,
        muteHttpExceptions: true,
      });

      const code = resp.getResponseCode();
      if (code >= 200 && code < 300) {
        const driveFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
        const filename =
          (normalizeText(data.namePassport) || normalizeText(data.regnum) || "PHOTO") +
          "_" +
          Date.now() +
          ".jpg";

        const blob = resp.getBlob().setName(filename);
        const file = driveFolder.createFile(blob);
        photoUrl = "https://drive.google.com/uc?export=view&id=" + file.getId();
      }
    }

    if (!photoUrl && isHttpUrl_(existingPhotoLink)) {
      photoUrl = existingPhotoLink; // reuse (FOUND)
    }

    // ===== Name (hyperlink only if NEW photo was stored) =====
    // IMPORTANT concept:
    // - Record FOUND: we DO NOT create a new hyperlink proof row.
    //   We only reuse existingPhotoLink for backend mirror, while the new row can keep name as plain text.
    // - Proof row exists only when photoUrl is NEW (captured) -> column I has photo -> actRow stays alive.
    const nameText = normalizeText(data.namePassport);
    const shouldCreateHyperlink = !!imageViewUrl && !!photoUrl; // NEW photo only

    const nameCellValue = shouldCreateHyperlink
      ? '=HYPERLINK("' + photoUrl + '","' + nameText + '")'
      : nameText;

    // ===== Category (remark) =====
    let remarkValue = normalizeText(data.remark);
    const unitNumberValue = normalizeText(data.unitNumber);
    if ((remarkValue === "OWNER" || remarkValue === "TENANT") && unitNumberValue) {
      remarkValue = remarkValue + " ( " + unitNumberValue + " )";
    }

    // ===== Reason =====
    let reasonValue = "";
    const mainReason = toText_(data.reason).toUpperCase();
    const reasonOther = toText_(data.reasonOther);
    if (mainReason === "OTHER" && reasonOther) {
      reasonValue = "OTHER ( " + reasonOther.toUpperCase() + " )";
    } else if (mainReason) {
      reasonValue = mainReason;
    }

    const now = new Date();

    // ===== Append row =====
    // Column I photo link:
    // - NEW: store photoUrl
    // - FOUND: leave blank (so it won't refresh proof window)
    const photoLinkColI = shouldCreateHyperlink ? photoUrl : "";

    sheet.appendRow([
      formatDateTimeDMY(now),            // A TIMESTAMP
      "",                                // B NAME (set after)
      normalizeText(data.mykadPassport), // C MYKAD/PASSPORT
      normalizeText(data.regnum),        // D REGNUM
      formatPhone(data.contact),         // E CONTACT
      remarkValue,                       // F CATEGORY
      normalizeText(data.tower),         // G TOWER
      reasonValue,                       // H REASON
      photoLinkColI                      // I PHOTO LINK (proof only)
    ]);

    const lastRow = sheet.getLastRow();

    // Column B: set formula OR plain text
    if (shouldCreateHyperlink && typeof nameCellValue === "string" && nameCellValue.indexOf("=HYPERLINK") === 0) {
      sheet.getRange(lastRow, 2).setFormula(nameCellValue);
    } else {
      sheet.getRange(lastRow, 2).setValue(nameCellValue);
    }

    // ===== Auto delete (rolling) =====
    const deleted = autoDeleteRows_(sheet, MAX_ROW);

    // ===== Cache handling =====
    // Rule:
    // 1) If deleted=true (rows shifted) -> clear cache keys, then rebuild immediately (keep search fast)
    // 2) If deleted=false -> keep cache, incremental update
    const cache = CacheService.getScriptCache();

    if (deleted) {
      // ✅ Required by Edos: clear cache when autoDelete happens
      cache.remove(KEY_REG);
      cache.remove(KEY_ID);
      cache.remove(KEY_META);

      // Rebuild now so next search stays fast (avoid first-search penalty)
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
        const idx = buildIndexCache_();
        regStr = idx.reg || "";
        idStr = idx.id || "";
      }

      const regKey = normKey_(data.regnum);
      const idKey = normKey_(data.mykadPassport);

      // Keep old actRow unless this is a NEW proof row
      if (regKey) {
        const old = resolveCompact_(regKey, regStr);
        const actRow = shouldCreateHyperlink ? lastRow : (old.actRow || 0);
        regStr = upsertCompactLine_(regStr, regKey, lastRow, actRow);
      }

      if (idKey) {
        const old2 = resolveCompact_(idKey, idStr);
        const actRow2 = shouldCreateHyperlink ? lastRow : (old2.actRow || 0);
        idStr = upsertCompactLine_(idStr, idKey, lastRow, actRow2);
      }

      const metaNow = metaSig_(sheet);
      cache.put(KEY_REG, regStr || "", CACHE_TTL);
      cache.put(KEY_ID, idStr || "", CACHE_TTL);
      cache.put(KEY_META, metaNow, CACHE_TTL);
    }

    return output_({ success: true, deleted: !!deleted });
  } catch (err) {
    return output_({ success: false, error: true, message: err.message });
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
    if (actRow) {
      const proofRow = sheet.getRange(actRow, 9, 1, 1).getValues()[0][0];
      const proofLink = toText_(proofRow);
      if (isHttpUrl_(proofLink)) photoLink = proofLink;
    }

    return output_({
      exist: true,
      hasHyperlink: true,
      data: {
        namePassport: full[1] || "",
        mykadPassport: full[2] || "",
        regnum: full[3] || "",
        contact: full[4] || "",
        remark: full[5] || "",
        tower: full[6] || "",
        reason: full[7] || "",
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
