const SEARCH_RECORD_CONFIG = {
  sheetName: 'SEARCH RECORD',
  inputCell: 'C4',
  statusCell: 'C5',
  detectCell: 'C6',
  totalMatchCell: 'C7',
  historyTitleRow: 10,
  historyHeaderRow: 11,
  historyStartRow: 12,
  helperStartCol: 2,
  historyColumnCount: 9,
  maxDisplayRows: 300
};

const SEARCH_INPUT_PLACEHOLDER = 'MYKAD / PASSPORT / REG. NUM.';

function setupSearchRecordSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SEARCH_RECORD_CONFIG.sheetName);

  if (!sh) {
    sh = ss.insertSheet(SEARCH_RECORD_CONFIG.sheetName);
  }

  sh.clear();
  sh.clearFormats();
  sh.clearNotes();

  try {
    const filter = sh.getFilter();
    if (filter) filter.remove();
  } catch (err) {}

  if (sh.getMaxRows() < 400) {
    sh.insertRowsAfter(sh.getMaxRows(), 400 - sh.getMaxRows());
  }
  if (sh.getMaxColumns() < 15) {
    sh.insertColumnsAfter(sh.getMaxColumns(), 15 - sh.getMaxColumns());
  }

  sh.getRange('A1:O400').breakApart();
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(4);

  sh.setColumnWidth(1, 20);
  sh.setColumnWidth(2, 170);
  sh.setColumnWidth(3, 210);
  sh.setColumnWidth(4, 180);
  sh.setColumnWidth(5, 180);
  sh.setColumnWidth(6, 180);
  sh.setColumnWidth(7, 180);
  sh.setColumnWidth(8, 150);
  sh.setColumnWidth(9, 130);
  sh.setColumnWidth(10, 140);

  sh.getRange('B1:J1').merge();
  sh.getRange('B2:J2').merge();

  sh.getRange('C4:F4').merge();
  sh.getRange('C5:F5').merge();
  sh.getRange('C6:F6').merge();
  sh.getRange('C7:F7').merge();
  sh.getRange('B10:J10').merge();

  sh.getRange('B1').setValue('SECURE ENTRY | SEARCH RECORD SENSORY');
  sh.getRange('B2').setValue('REGISTER.ACCESS.SECURE');

  sh.getRange('B4').setValue('SEARCH INPUT');
  sh.getRange('B5').setValue('STATUS');
  sh.getRange('B6').setValue('AUTO DETECT');
  sh.getRange('B7').setValue('TOTAL MATCH');

  sh.getRange('B10').setValue('MATCHING ENTRY RECORDS');

  sh.getRange(
    SEARCH_RECORD_CONFIG.historyHeaderRow,
    SEARCH_RECORD_CONFIG.helperStartCol,
    1,
    9
  ).setValues([[
    'TIMESTAMP',
    'NAME',
    'MYKAD / PASSPORT',
    'REGISTRATION NUMBER',
    'CONTACT',
    'CATEGORY',
    'TOWER',
    'REASON',
    'PHOTO LINK'
  ]]);

  styleSearchRecordSheet_(sh);
  applySearchInputPlaceholder_(sh);
  sh.getRange(SEARCH_RECORD_CONFIG.inputCell).activate();
}

function styleSearchRecordSheet_(sh) {
  // Title
  sh.getRange('B1:J1')
    .setBackground('#e8f5e9')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setFontSize(18)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  // Subtitle
  sh.getRange('B2:J2')
    .setBackground('#ffffff')
    .setFontColor('#558b2f')
    .setFontStyle('italic')
    .setFontWeight('normal')
    .setFontSize(10)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  // Left labels
  sh.getRange('B4:B7')
    .setBackground('#ffffff')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  // Input / status / detect / total cells
  sh.getRange('C4:F7')
    .setBackground('#f1f8e9')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#a5d6a7', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Section title
  sh.getRange('B10:J10')
    .setBackground('#f1f8e9')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#dcedc8', SpreadsheetApp.BorderStyle.SOLID);

  // Table header
  sh.getRange(
    SEARCH_RECORD_CONFIG.historyHeaderRow,
    SEARCH_RECORD_CONFIG.helperStartCol,
    1,
    9
  )
    .setBackground('#e8f5e9')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#c8e6c9', SpreadsheetApp.BorderStyle.SOLID);

  // General font
  sh.getRange('B1:J250').setFontFamily('Arial');

  // Row heights similar to dashboard feel
  sh.setRowHeight(1, 32);
  sh.setRowHeight(2, 22);

  for (let r = 4; r <= 7; r++) {
    sh.setRowHeight(r, 24);
  }

  sh.setRowHeight(10, 24);
  sh.setRowHeight(11, 24);
}

function runSearchRecord() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(DASHBOARD_CONFIG.sourceSheetName);
  const searchSheet = ss.getSheetByName(SEARCH_RECORD_CONFIG.sheetName);

  if (!rawSheet) throw new Error('Source sheet "SENSORY" was not found.');
  if (!searchSheet) throw new Error('Sheet "SEARCH RECORD" was not found. Please run setupSearchRecordSheet() first.');

  const rawInput = safeString_(searchSheet.getRange(SEARCH_RECORD_CONFIG.inputCell).getDisplayValue());

  clearSearchRecordResultsOnly_(searchSheet);

  if (!rawInput || isSearchInputPlaceholder_(rawInput)) {
    return;
  }

  const detectedField = detectSearchFieldSearchRecord_(rawInput);
  const normalizedInput = normalizeSearchKey_(rawInput);

  searchSheet.getRange(SEARCH_RECORD_CONFIG.detectCell).setValue(detectedField);

  const result = searchRecordsFromRawSheet_(rawSheet, detectedField, normalizedInput);
  renderSearchRecord_(searchSheet, result, detectedField, normalizedInput);
}

function renderSearchRecord_(sheet, result, detectedField, normalizedInput) {
  sheet.getRange(SEARCH_RECORD_CONFIG.totalMatchCell).setValue(result.matches.length);

  if (!result.matches.length) {
    sheet.getRange(SEARCH_RECORD_CONFIG.statusCell).setValue('NO RECORD FOUND');
    colorSearchStatusCell_(sheet.getRange(SEARCH_RECORD_CONFIG.statusCell), 'NO RECORD FOUND');
    return;
  }

  if (!result.isMatch) {
    sheet.getRange(SEARCH_RECORD_CONFIG.statusCell).setValue('MISMATCH DETECTED');
    colorSearchStatusCell_(sheet.getRange(SEARCH_RECORD_CONFIG.statusCell), 'MISMATCH DETECTED');
    return;
  }

  const primary = result.primary;
  const status = primary.hasPhotoEvidence ? 'RECORD FOUND' : 'RECORD EXPIRED';

  sheet.getRange(SEARCH_RECORD_CONFIG.statusCell).setValue(status);
  colorSearchStatusCell_(sheet.getRange(SEARCH_RECORD_CONFIG.statusCell), status);

  const displayMatches = result.matches.slice(0, SEARCH_RECORD_CONFIG.maxDisplayRows);

  const values = displayMatches.map(function (item) {
    return [
      normalizeSearchTimestamp_(item.timestamp),
      toUpperSearch_(item.namePassport),
      toUpperSearch_(item.mykadPassport),
      toUpperSearch_(item.regnum),
      formatContactSearch_(item.contact),
      toUpperSearch_(item.remark),
      toUpperSearch_(item.tower),
      toUpperSearch_(item.reason),
      item.effectivePhotoLink
        ? '=HYPERLINK("' + escapeFormulaTextSearch_(item.effectivePhotoLink) + '","OPEN PHOTO")'
        : ''
    ];
  });

  if (values.length) {
    const startRow = SEARCH_RECORD_CONFIG.historyStartRow;
    const startCol = SEARCH_RECORD_CONFIG.helperStartCol;

    sheet.getRange(startRow, startCol, values.length, 9).setValues(values);

    sheet.getRange(startRow, startCol, values.length, 1)
      .setNumberFormat('dd/MM/yyyy HH:mm:ss')
      .setHorizontalAlignment('left');

    sheet.getRange(startRow, startCol + 4, values.length, 1)
      .setNumberFormat('@')
      .setHorizontalAlignment('left');

    applyHistoryRowColors_(sheet, displayMatches);

    try {
      const filter = sheet.getFilter();
      if (filter) filter.remove();
    } catch (err) {}

    sheet.getRange(
      SEARCH_RECORD_CONFIG.historyHeaderRow,
      SEARCH_RECORD_CONFIG.helperStartCol,
      values.length + 1,
      9
    ).createFilter();
  }
}

function applyHistoryRowColors_(sheet, matches) {
  if (!matches.length) return;

  const colors = matches.map(function (item) {
    const bg = item.hasDirectPhotoLink
      ? '#E8F5E9'
      : item.hasPhotoEvidence
      ? '#F1F8E9'
      : '#FFF8E1';

    return [bg, bg, bg, bg, bg, bg, bg, bg, bg];
  });

  sheet.getRange(
    SEARCH_RECORD_CONFIG.historyStartRow,
    SEARCH_RECORD_CONFIG.helperStartCol,
    colors.length,
    9
  ).setBackgrounds(colors);
}

function searchRecordsFromRawSheet_(rawSheet, detectedField, normalizedInput) {
  const lastRow = rawSheet.getLastRow();
  if (lastRow < 2) {
    return { matches: [], primary: null, isMatch: false };
  }

  const values = rawSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const formulas = rawSheet.getRange(2, 1, lastRow - 1, 9).getFormulas();

  const targetIndex = detectedField === 'REGNUM' ? 3 : 2;
  const matches = [];

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    const formulaRow = formulas[i];
    const rawValue = row[targetIndex];
    const candidate = normalizeSearchKey_(rawValue);

    if (!candidate || candidate !== normalizedInput) continue;

    const directPhotoLink = safeString_(row[8]);
    const hyperlinkFromName = extractUrlFromHyperlinkFormulaSearch_(formulaRow[1]);
    const effectivePhotoLink = directPhotoLink || hyperlinkFromName || '';

    matches.push({
      rowNumber: i + 2,
      timestamp: row[0],
      namePassport: row[1],
      mykadPassport: row[2],
      regnum: row[3],
      contact: row[4],
      remark: row[5],
      tower: row[6],
      reason: row[7],
      directPhotoLink: directPhotoLink,
      hyperlinkFromName: hyperlinkFromName,
      effectivePhotoLink: effectivePhotoLink,
      hasDirectPhotoLink: !!directPhotoLink,
      hasPhotoEvidence: !!effectivePhotoLink
    });
  }

  if (!matches.length) {
    return { matches: [], primary: null, isMatch: false };
  }

  const primary =
    matches.find(function (item) { return item.hasDirectPhotoLink; }) ||
    matches.find(function (item) { return item.hasPhotoEvidence; }) ||
    matches[0];

  const returnedRegNorm = normalizeSearchKey_(primary.regnum);
  const returnedIdNorm = normalizeSearchKey_(primary.mykadPassport);

  let isMatch = true;
  if (detectedField === 'REGNUM') {
    isMatch = returnedRegNorm ? returnedRegNorm === normalizedInput : true;
  } else if (detectedField === 'MYKADPASSPORT') {
    isMatch = returnedIdNorm ? returnedIdNorm === normalizedInput : true;
  }

  return {
    matches: matches,
    primary: primary,
    isMatch: isMatch
  };
}

function clearSearchRecord() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SEARCH_RECORD_CONFIG.sheetName);
  if (!sheet) return;

  sheet.getRange(SEARCH_RECORD_CONFIG.inputCell).clearContent();
  clearSearchRecordResultsOnly_(sheet);
  applySearchInputPlaceholder_(sheet);
}

function clearSearchRecordResultsOnly_(sheet) {
  try {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
  } catch (err) {}

  sheet.getRangeList([
    SEARCH_RECORD_CONFIG.statusCell,
    SEARCH_RECORD_CONFIG.detectCell,
    SEARCH_RECORD_CONFIG.totalMatchCell
  ]).clearContent();

  sheet.getRange('C5:F5').setBackground('#F7FFF9');

  const maxRowsToClear = Math.max(sheet.getMaxRows() - SEARCH_RECORD_CONFIG.historyStartRow + 1, 1);
  sheet.getRange(
    SEARCH_RECORD_CONFIG.historyStartRow,
    SEARCH_RECORD_CONFIG.helperStartCol,
    maxRowsToClear,
    9
  ).clearContent().setBackground(null).setBorder(false, false, false, false, false, false);
}

function normalizeSearchKey_(value) {
  return safeUpper_(value).replace(/[^A-Z0-9]/g, '');
}

function extractUrlFromHyperlinkFormulaSearch_(formula) {
  const text = safeString_(formula);
  if (!text) return '';

  const match = text.match(/^=HYPERLINK\("([^"]+)"/i);
  return match ? match[1].trim() : '';
}

function toDirectImageUrlSearch_(url) {
  const raw = safeString_(url);
  if (!raw) return '';

  const fileId = extractDriveFileIdSearch_(raw);
  if (fileId) {
    return 'https://drive.google.com/uc?export=view&id=' + fileId;
  }

  if (/^https?:\/\//i.test(raw)) return raw;
  return '';
}

function extractDriveFileIdSearch_(url) {
  const text = safeString_(url);
  if (!text) return '';

  let match = text.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];

  match = text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match) return match[1];

  return '';
}

function colorSearchStatusCell_(range, status) {
  const text = safeUpper_(status);

  if (text === 'RECORD FOUND') {
    range.setBackground('#D9EAD3').setFontColor('#176B39').setFontWeight('bold').setHorizontalAlignment('left');
  } else if (text === 'RECORD EXPIRED') {
    range.setBackground('#FFF2CC').setFontColor('#7F6000').setFontWeight('bold').setHorizontalAlignment('left');
  } else if (text === 'NO RECORD FOUND') {
    range.setBackground('#F4CCCC').setFontColor('#990000').setFontWeight('bold').setHorizontalAlignment('left');
  } else {
    range.setBackground('#D9EAF7').setFontColor('#1C4587').setFontWeight('bold').setHorizontalAlignment('left');
  }
}

function toUpperSearch_(value) {
  return safeUpper_(value);
}

function escapeFormulaTextSearch_(text) {
  return String(text || '').replace(/"/g, '""');
}

function formatContactSearch_(value) {
  let text = String(value || '').replace(/[^0-9]/g, '');

  if (!text) return '';

  if (text.charAt(0) !== '0') {
    text = '0' + text;
  }

  return text;
}

function detectSearchFieldSearchRecord_(raw) {
  const s = String(raw || '').trim().toUpperCase();
  const digitsOnly = s.replace(/[^0-9]/g, '');
  const alnumOnly = s.replace(/[^A-Z0-9]/g, '');

  if (digitsOnly.length === 12 && !/[A-Z]/.test(alnumOnly)) return 'MYKADPASSPORT';
  if (/^\d{6}-\d{2}-\d{4}$/.test(s)) return 'MYKADPASSPORT';
  if (/^[A-Z]{1,3}\d{6,}[A-Z]{3}$/.test(alnumOnly)) return 'MYKADPASSPORT';
  if (/^[A-Z]{1,3}\d{6,}$/.test(alnumOnly)) return 'MYKADPASSPORT';

  if (/^[A-Z]{1,3}\d{4}[A-Z]?$/.test(alnumOnly)) return 'REGNUM';

  return /[A-Z]/.test(alnumOnly) ? 'REGNUM' : 'MYKADPASSPORT';
}

function onSelectionChange(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== SEARCH_RECORD_CONFIG.sheetName) return;

    const inputRange = sheet.getRange(SEARCH_RECORD_CONFIG.inputCell);
    const selectedA1 = e.range.getA1Notation();
    const inputValue = String(inputRange.getDisplayValue() || '').trim();

    // Bila user pilih/tap C4, placeholder terus hilang
    if (selectedA1 === SEARCH_RECORD_CONFIG.inputCell) {
      if (isSearchInputPlaceholder_(inputValue)) {
        inputRange.clearContent();
      }

      inputRange
        .setFontColor('#000000')
        .setFontStyle('normal')
        .setHorizontalAlignment('left');

      return;
    }

    // Bila user keluar dari C4 dan cell kosong, placeholder muncul balik
    if (!inputValue) {
      applySearchInputPlaceholder_(sheet);
    }

  } catch (err) {
    Logger.log('onSelectionChange error: ' + err);
  }
}

function isSearchInputPlaceholder_(value) {
  return safeUpper_(value) === SEARCH_INPUT_PLACEHOLDER;
}

function applySearchInputPlaceholder_(sheet) {
  const range = sheet.getRange(SEARCH_RECORD_CONFIG.inputCell);
  const current = safeString_(range.getDisplayValue());

  if (!current) {
    range
      .setValue(SEARCH_INPUT_PLACEHOLDER)
      .setFontColor('#9AA0A6')
      .setFontStyle('italic')
      .setHorizontalAlignment('left');
  }
}

function clearSearchInputPlaceholder_(sheet) {
  const range = sheet.getRange(SEARCH_RECORD_CONFIG.inputCell);
  const current = safeString_(range.getDisplayValue());

  if (isSearchInputPlaceholder_(current)) {
    range.clearContent();
  }

  range
    .setFontColor('#000000')
    .setFontStyle('normal')
    .setHorizontalAlignment('left');
}

function setSearchInputActiveStyle_(sheet) {
  sheet.getRange(SEARCH_RECORD_CONFIG.inputCell)
    .setFontColor('#000000')
    .setFontStyle('normal')
    .setHorizontalAlignment('left');
}

function normalizeSearchTimestamp_(value) {
  if (typeof normalizeDateValue_ === 'function') {
    const parsed = normalizeDateValue_(value);
    if (parsed) return parsed;
  }
  return value;
}

function stripSearchPlaceholder_(value) {
  const text = String(value || '');
  const placeholder = String(SEARCH_INPUT_PLACEHOLDER || '');

  if (!text) return '';
  if (!placeholder) return text;

  if (safeUpper_(text) === safeUpper_(placeholder)) {
    return '';
  }

  const escaped = placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return text.replace(new RegExp('^\\s*' + escaped + '\\s*', 'i'), '');
}
