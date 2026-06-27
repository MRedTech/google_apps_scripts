const SEARCH_RECORD_CONFIG = {
  sheetName: 'SEARCH RECORD',
  inputCell: 'C4',
  statusCell: 'C5',
  detectCell: 'C6',
  totalMatchCell: 'C7',
  actionSearchCell: 'F4',
  actionClearCell: 'F5',
  actionSearchLabelRange: 'G4:H4',
  actionClearLabelRange: 'G5:H5',
  activeAfterActionCell: 'B1',
  historyTitleRow: 10,
  historyHeaderRow: 11,
  historyStartRow: 12,
  helperStartCol: 2,
  historyLeftStartCol: 2,
  historyLeftColumnCount: 4,
  historyRightStartCol: 7,
  historyRightColumnCount: 5,
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
  sh.setColumnWidth(6, 42);
  sh.setColumnWidth(7, 170);
  sh.setColumnWidth(8, 180);
  sh.setColumnWidth(9, 150);
  sh.setColumnWidth(10, 130);
  sh.setColumnWidth(11, 140);

  sh.getRange('B1:K1').merge();
  sh.getRange('B2:K2').merge();
  sh.getRange('C4:E4').merge();
  sh.getRange('C5:E5').merge();
  sh.getRange('C6:E6').merge();
  sh.getRange('C7:E7').merge();
  sh.getRange('G4:H4').merge();
  sh.getRange('G5:H5').merge();
  sh.getRange('B10:K10').merge();

  sh.getRange('B1').setValue('SECURE ENTRY | SEARCH RECORD ' + getSearchSiteLabel_());
  sh.getRange('B2').setValue('REGISTER.ACCESS.SECURE');

  sh.getRange('B4').setValue('SEARCH INPUT');
  sh.getRange('B5').setValue('STATUS');
  sh.getRange('B6').setValue('AUTO DETECT');
  sh.getRange('B7').setValue('TOTAL MATCH');

  sh.getRange('B10').setValue('MATCHING ENTRY RECORDS');
  sh.getRange('G4').setValue('SEARCH RECORD');
  sh.getRange('G5').setValue('CLEAR SEARCH');

  sh.getRange(SEARCH_RECORD_CONFIG.historyHeaderRow, SEARCH_RECORD_CONFIG.historyLeftStartCol, 1, SEARCH_RECORD_CONFIG.historyLeftColumnCount)
    .setValues([[
      'TIMESTAMP',
      'NAME',
      'MYKAD / PASSPORT',
      'REGISTRATION NUMBER'
    ]]);

  sh.getRange(SEARCH_RECORD_CONFIG.historyHeaderRow, SEARCH_RECORD_CONFIG.historyRightStartCol, 1, SEARCH_RECORD_CONFIG.historyRightColumnCount)
    .setValues([[
      'CONTACT',
      'CATEGORY',
      'TOWER',
      'REASON',
      'PHOTO LINK'
    ]]);

  sh.getRange('F11')
    .clearContent()
    .clearDataValidations()
    .setBackground('#ffffff')
    .setBorder(false, false, false, false, false, false);

  styleSearchRecordSheet_(sh);
  setupSearchRecordActionCheckboxes_(sh);
  sh.getRange(SEARCH_RECORD_CONFIG.inputCell).setNumberFormat('@');
  applySearchInputPlaceholder_(sh);
  sh.getRange(SEARCH_RECORD_CONFIG.inputCell).activate();
}

function styleSearchRecordSheet_(sh) {
  sh.getRange('B1:K1')
    .setBackground('#e8f5e9')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setFontSize(18)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sh.getRange('B2:K2')
    .setBackground('#ffffff')
    .setFontColor('#558b2f')
    .setFontStyle('italic')
    .setFontWeight('normal')
    .setFontSize(10)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sh.getRange('B4:B7')
    .setBackground('#ffffff')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sh.getRange('C4:E7')
    .setBackground('#f1f8e9')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#a5d6a7', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sh.getRange('B10:K10')
    .setBackground('#f1f8e9')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#dcedc8', SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(SEARCH_RECORD_CONFIG.historyHeaderRow, SEARCH_RECORD_CONFIG.historyLeftStartCol, 1, SEARCH_RECORD_CONFIG.historyLeftColumnCount)
    .setBackground('#e8f5e9')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#c8e6c9', SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(SEARCH_RECORD_CONFIG.historyHeaderRow, SEARCH_RECORD_CONFIG.historyRightStartCol, 1, SEARCH_RECORD_CONFIG.historyRightColumnCount)
    .setBackground('#e8f5e9')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#c8e6c9', SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange('F11')
    .clearContent()
    .clearDataValidations()
    .setBackground('#ffffff')
    .setBorder(false, false, false, false, false, false);

  sh.getRange('B1:K250').setFontFamily('Arial');
  sh.setRowHeight(1, 32);
  sh.setRowHeight(2, 22);

  for (let r = 4; r <= 7; r++) {
    sh.setRowHeight(r, 24);
  }

  sh.setRowHeight(10, 24);
  sh.setRowHeight(11, 24);
  styleSearchRecordActionArea_(sh);
}


function styleSearchRecordActionArea_(sh) {
  sh.getRange('F4:F5')
    .setBackground('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(false, false, false, false, false, false);

  sh.getRange('G4:H5')
    .setBackground('#ffffff')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
}

function setupSearchRecordActionCheckboxes_(sheet) {
  restoreSearchRecordActionCheckboxes_(sheet);
}

function restoreSearchRecordActionCheckboxes_(sheet) {
  const checkboxCells = [
    SEARCH_RECORD_CONFIG.actionSearchCell,
    SEARCH_RECORD_CONFIG.actionClearCell
  ];

  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();

  checkboxCells.forEach(function (a1) {
    const range = sheet.getRange(a1);
    range.clearContent();
    range.clearDataValidations();
    range.setDataValidation(checkboxRule);
    range.setValue(false);
    range
      .setBackground('#ffffff')
      .setFontColor('#000000')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(false, false, false, false, false, false);
  });

  sheet.getRange('G4').setValue('SEARCH RECORD');

  sheet.getRange(SEARCH_RECORD_CONFIG.actionSearchLabelRange)
    .setBackground('#ffffff')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange('G5').setValue('CLEAR SEARCH');

  sheet.getRange(SEARCH_RECORD_CONFIG.actionClearLabelRange)
    .setBackground('#ffffff')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
}

function setSearchRecordActionBusy_(sheet, action) {
  const searchBox = sheet.getRange(SEARCH_RECORD_CONFIG.actionSearchCell);
  const clearBox = sheet.getRange(SEARCH_RECORD_CONFIG.actionClearCell);

  if (action === 'SEARCH') {
    searchBox.setValue(true).setBackground('#D9EAD3');
    clearBox.clearDataValidations().clearContent().setBackground('#E0E0E0').setBorder(false, false, false, false, false, false);
  } else {
    clearBox.setValue(true).setBackground('#D9EAD3');
    searchBox.clearDataValidations().clearContent().setBackground('#E0E0E0').setBorder(false, false, false, false, false, false);
  }

  sheet.getRange(SEARCH_RECORD_CONFIG.actionSearchLabelRange)
    .setBackground(action === 'SEARCH' ? '#E8F5E9' : '#F5F5F5')
    .setFontColor(action === 'SEARCH' ? '#1b5e20' : '#777777');

  sheet.getRange(SEARCH_RECORD_CONFIG.actionClearLabelRange)
    .setBackground(action === 'CLEAR' ? '#E8F5E9' : '#F5F5F5')
    .setFontColor(action === 'CLEAR' ? '#1b5e20' : '#777777');
}

function handleSearchRecordCheckboxEdit_(e) {
  if (!e || !e.range) return false;

  const range = e.range;
  const sheet = range.getSheet();
  if (!sheet || sheet.getName() !== SEARCH_RECORD_CONFIG.sheetName) return false;

  const editedCell = range.getA1Notation();
  const isSearchAction = editedCell === SEARCH_RECORD_CONFIG.actionSearchCell;
  const isClearAction = editedCell === SEARCH_RECORD_CONFIG.actionClearCell;

  if (!isSearchAction && !isClearAction) return false;
  if (String(e.value || '').toUpperCase() !== 'TRUE') return true;

  const lock = LockService.getDocumentLock();
  let lockAcquired = false;

  try {
    lock.waitLock(10000);
    lockAcquired = true;

    sheet.getRange(SEARCH_RECORD_CONFIG.activeAfterActionCell).activate();
    SpreadsheetApp.flush();

    setSearchRecordActionBusy_(sheet, isSearchAction ? 'SEARCH' : 'CLEAR');
    SpreadsheetApp.flush();

    if (isSearchAction) {
      runSearchRecord();
    } else {
      clearSearchRecord();
    }
  } catch (err) {
    sheet.getRange(SEARCH_RECORD_CONFIG.statusCell).setValue('ACTION ERROR');
    colorSearchStatusCell_(sheet.getRange(SEARCH_RECORD_CONFIG.statusCell), 'ACTION ERROR');
    throw err;
  } finally {
    restoreSearchRecordActionCheckboxes_(sheet);
    sheet.getRange(SEARCH_RECORD_CONFIG.activeAfterActionCell).activate();
    SpreadsheetApp.flush();

    if (lockAcquired) {
      lock.releaseLock();
    }
  }

  return true;
}

function runSearchRecord() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheetName = getSearchSourceSheetName_();
  const rawSheet = ss.getSheetByName(sourceSheetName);
  const searchSheet = ss.getSheetByName(SEARCH_RECORD_CONFIG.sheetName);

  if (!rawSheet) throw new Error('Source sheet "' + sourceSheetName + '" was not found.');
  if (!searchSheet) throw new Error('Sheet "SEARCH RECORD" was not found. Please run setupSearchRecordSheet() first.');

  const rawInputOriginal = safeString_(searchSheet.getRange(SEARCH_RECORD_CONFIG.inputCell).getDisplayValue());
  const rawInput = stripSearchPlaceholder_(rawInputOriginal).trim();

  clearSearchRecordResultsOnly_(searchSheet);

  if (!rawInput) {
    return;
  }

  const detectedField = detectSearchFieldSearchRecord_(rawInput);
  const normalizedInput = normalizeSearchKey_(rawInput);

  searchSheet.getRange(SEARCH_RECORD_CONFIG.detectCell).setValue(detectedField);

  const result = searchRecordsFromRawSheet_(rawSheet, detectedField, normalizedInput);
  renderSearchRecord_(searchSheet, result);
}

function renderSearchRecord_(sheet, result) {
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
    const leftValues = values.map(function (row) { return row.slice(0, 4); });
    const rightValues = values.map(function (row) { return row.slice(4); });

    sheet.getRange(startRow, SEARCH_RECORD_CONFIG.historyLeftStartCol, values.length, SEARCH_RECORD_CONFIG.historyLeftColumnCount)
      .setValues(leftValues);

    sheet.getRange(startRow, SEARCH_RECORD_CONFIG.historyRightStartCol, values.length, SEARCH_RECORD_CONFIG.historyRightColumnCount)
      .setValues(rightValues);

    sheet.getRange(startRow, SEARCH_RECORD_CONFIG.historyLeftStartCol, values.length, 1)
      .setNumberFormat('dd/MM/yyyy HH:mm:ss')
      .setHorizontalAlignment('left');

    sheet.getRange(startRow, SEARCH_RECORD_CONFIG.historyLeftStartCol + 2, values.length, 1)
      .setNumberFormat('@')
      .setHorizontalAlignment('left');

    sheet.getRange(startRow, SEARCH_RECORD_CONFIG.historyRightStartCol, values.length, 1)
      .setNumberFormat('@')
      .setHorizontalAlignment('left');

    applyHistoryRowColors_(sheet, displayMatches);
  }
}

function applyHistoryRowColors_(sheet, matches) {
  if (!matches.length) return;

  const leftColors = [];
  const rightColors = [];

  matches.forEach(function (item) {
    const bg = item.hasDirectPhotoLink
      ? '#E8F5E9'
      : item.hasPhotoEvidence
      ? '#F1F8E9'
      : '#FFF8E1';

    leftColors.push([bg, bg, bg, bg]);
    rightColors.push([bg, bg, bg, bg, bg]);
  });

  sheet.getRange(
    SEARCH_RECORD_CONFIG.historyStartRow,
    SEARCH_RECORD_CONFIG.historyLeftStartCol,
    leftColors.length,
    SEARCH_RECORD_CONFIG.historyLeftColumnCount
  ).setBackgrounds(leftColors);

  sheet.getRange(
    SEARCH_RECORD_CONFIG.historyStartRow,
    SEARCH_RECORD_CONFIG.historyRightStartCol,
    rightColors.length,
    SEARCH_RECORD_CONFIG.historyRightColumnCount
  ).setBackgrounds(rightColors);
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
  sheet.getRange(SEARCH_RECORD_CONFIG.inputCell).setNumberFormat('@');
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

  sheet.getRange('C5:E5').setBackground('#F7FFF9');

  const maxRowsToClear = Math.max(sheet.getMaxRows() - SEARCH_RECORD_CONFIG.historyStartRow + 1, 1);
  sheet.getRange(
    SEARCH_RECORD_CONFIG.historyStartRow,
    SEARCH_RECORD_CONFIG.historyLeftStartCol,
    maxRowsToClear,
    SEARCH_RECORD_CONFIG.historyLeftColumnCount
  ).clearContent().setBackground(null).setBorder(false, false, false, false, false, false);

  sheet.getRange(
    SEARCH_RECORD_CONFIG.historyStartRow,
    SEARCH_RECORD_CONFIG.historyRightStartCol,
    maxRowsToClear,
    SEARCH_RECORD_CONFIG.historyRightColumnCount
  ).clearContent().setBackground(null).setBorder(false, false, false, false, false, false);

  sheet.getRange(SEARCH_RECORD_CONFIG.historyStartRow, 6, maxRowsToClear, 1)
    .clearContent()
    .clearDataValidations()
    .setBackground(null)
    .setBorder(false, false, false, false, false, false);
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

function isSearchInputPlaceholder_(value) {
  return safeUpper_(value) === SEARCH_INPUT_PLACEHOLDER;
}

function applySearchInputPlaceholder_(sheet) {
  const range = sheet.getRange(SEARCH_RECORD_CONFIG.inputCell);
  const current = safeString_(range.getDisplayValue());

  range.setNumberFormat('@');

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

  range.setNumberFormat('@');

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
    .setNumberFormat('@')
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

function getSearchSourceSheetName_() {
  if (typeof DASHBOARD_CONFIG !== 'undefined' && DASHBOARD_CONFIG.sourceSheetName) {
    return String(DASHBOARD_CONFIG.sourceSheetName).trim();
  }
  return 'SENSORY';
}

function getSearchSiteLabel_() {
  return safeUpper_(getSearchSourceSheetName_());
}


function handleSearchRecordEdit(e) {
  if (handleSearchRecordCheckboxEdit_(e)) return;
}

/*
ONEDIT INTEGRATION NOTE:
Use handleSearchRecordEdit(e) as an installable On edit trigger for the SEARCH RECORD checkbox.

If this Apps Script project already has an onEdit(e) function and you prefer to combine triggers, add this line at the top of that existing onEdit(e):

  if (handleSearchRecordCheckboxEdit_(e)) return;

Do not create a second simple onEdit(e) function in the same project.
*/
