const SEARCH_RECORD_CONFIG = {
  sheetName: 'SEARCH RECORD',
  inputCell: 'C4',
  statusCell: 'C5',
  detectCell: 'C6',
  totalMatchCell: 'C7',
  actionClearCell: 'F7',
  actionClearLabelRange: 'G7:H7',
  activeAfterActionCell: 'B1',
  editableRanges: ['C4:E4', 'F7'],
  protectionDescription: 'SECURE ENTRY SEARCH RECORD LOCK',
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
  sh.getRange('A1:O400').clearDataValidations();
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(4);

  sh.setColumnWidth(1, 20);
  sh.setColumnWidth(2, 170);
  sh.setColumnWidth(3, 210);
  sh.setColumnWidth(4, 180);
  sh.setColumnWidth(5, 180);
  sh.setColumnWidth(6, 24); // dedicated narrow checkbox / spacer column, same concept as dashboard
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
  sh.getRange('G7:H7').merge();
  sh.getRange('B10:K10').merge();

  clearSearchRecordHeaderValidation_(sh);

  sh.getRange('B1').setValue('SECURE ENTRY | SEARCH RECORD ' + getSearchSiteLabel_());
  sh.getRange('B2').setValue('REGISTER.ACCESS.SECURE');

  sh.getRange('B4').setValue('SEARCH INPUT');
  sh.getRange('B5').setValue('STATUS');
  sh.getRange('B6').setValue('AUTO DETECT');
  sh.getRange('B7').setValue('TOTAL MATCH');

  sh.getRange('B10').setValue('MATCHING ENTRY RECORDS');
  sh.getRange('G7').setValue('CLEAR SEARCH');

  sh.getRange(SEARCH_RECORD_CONFIG.historyHeaderRow, SEARCH_RECORD_CONFIG.historyLeftStartCol, 1, SEARCH_RECORD_CONFIG.historyLeftColumnCount)
    .setValues([[
      'TIMESTAMP',
      'NAME',
      'MYKAD / PASSPORT',
      'REGISTRATION NUMBER'
    ]]);

  sh.getRange('F11:G11')
    .merge()
    .setValue('CONTACT');

  sh.getRange(SEARCH_RECORD_CONFIG.historyHeaderRow, 8, 1, 4)
    .setValues([[
      'CATEGORY',
      'TOWER',
      'REASON',
      'PHOTO LINK'
    ]]);

  styleSearchRecordSheet_(sh);
  setupSearchRecordActionCheckboxes_(sh);
  sh.getRange(SEARCH_RECORD_CONFIG.inputCell).setNumberFormat('@');
  applySearchInputPlaceholder_(sh);
  applySearchRecordSheetProtection_(sh);
  sh.getRange(SEARCH_RECORD_CONFIG.inputCell).activate();
}

function clearSearchRecordHeaderValidation_(sheet) {
  // Keep header/title area free from any old checkbox/data-validation rules.
  // This prevents Google Sheets on tablet from showing validation errors when B1 is activated.
  sheet.getRangeList([
    'B1:K1',
    'B2:K2',
    'B10:K10'
  ]).clearDataValidations();
}


function applySearchRecordSheetProtection_(sheet) {
  // Lock the SEARCH RECORD sheet and leave only the agreed user input cells editable:
  // C4:E4 = SEARCH INPUT, F7 = CLEAR SEARCH checkbox.
  const editableRanges = SEARCH_RECORD_CONFIG.editableRanges.map(function (a1) {
    return sheet.getRange(a1);
  });

  // Remove old protections only on this SEARCH RECORD sheet before applying the current rule.
  // This prevents outdated protected ranges from blocking C4:E4 or F7.
  try {
    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (protection) {
      protection.remove();
    });
  } catch (err) {}

  try {
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (protection) {
      protection.remove();
    });
  } catch (err) {}

  const protection = sheet.protect().setDescription(SEARCH_RECORD_CONFIG.protectionDescription);
  protection.setWarningOnly(false);
  protection.setUnprotectedRanges(editableRanges);

  // Keep the script runner as the sheet-protection editor, then remove other editors
  // from this protection only. Spreadsheet sharing itself is not changed.
  try {
    const me = Session.getEffectiveUser();
    protection.addEditor(me);

    protection.getEditors().forEach(function (editor) {
      if (editor.getEmail && editor.getEmail() !== me.getEmail()) {
        protection.removeEditor(editor);
      }
    });

    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  } catch (err) {
    Logger.log('Search Record protection editor cleanup skipped: ' + err);
  }
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
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#dcedc8', SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange('B11:K11')
    .setBackground('#e8f5e9')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#c8e6c9', SpreadsheetApp.BorderStyle.SOLID);

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
  sh.getRange('F4:H6')
    .breakApart()
    .clearContent()
    .clearDataValidations()
    .clearFormat()
    .setBackground('#ffffff')
    .setBorder(false, false, false, false, false, false);

  sh.getRange('F7')
    .setBackground('#ffffff')
    .setFontColor('#9AA0A6')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle')
    .setBorder(false, false, false, false, false, false);

  sh.getRange('G7:H7')
    .setBackground('#ffffff')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBorder(false, false, false, false, false, false);
}

function setupSearchRecordActionCheckboxes_(sheet) {
  // SEARCH RECORD checkbox has been removed.
  // Search runs when the user edits C4 and presses Enter.
  // Only CLEAR SEARCH remains as a dashboard-style checkbox at F7.
  const actionArea = sheet.getRange('F4:H7');
  actionArea
    .breakApart()
    .clearContent()
    .clearDataValidations()
    .clearFormat()
    .setBackground('#ffffff')
    .setBorder(false, false, false, false, false, false);

  setupSearchRecordClearCheckboxCell_(sheet);

  sheet.getRange(SEARCH_RECORD_CONFIG.actionClearLabelRange)
    .breakApart()
    .merge()
    .setValue('CLEAR SEARCH')
    .setBackground('#ffffff')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBorder(false, false, false, false, false, false);
}

function setupSearchRecordClearCheckboxCell_(sheet) {
  const range = sheet.getRange(SEARCH_RECORD_CONFIG.actionClearCell);

  // Static checkbox: create it only during setup, then later only setValue(false) to untick.
  range
    .breakApart()
    .clearContent()
    .clearDataValidations()
    .clearFormat()
    .insertCheckboxes()
    .setValue(false)
    .setBackground('#ffffff')
    .setFontColor('#9AA0A6')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle')
    .setBorder(false, false, false, false, false, false);
}

function isSearchRecordClearCheckboxHealthy_(range) {
  const rule = range.getDataValidation();
  const hasCheckboxRule = !!rule &&
    rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX;

  return hasCheckboxRule && typeof range.getValue() === 'boolean';
}

function setSearchRecordActiveCell_(sheet) {
  clearSearchRecordHeaderValidation_(sheet);
  const target = sheet.getRange(SEARCH_RECORD_CONFIG.activeAfterActionCell);

  try {
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  } catch (err) {}

  try {
    sheet.setActiveSelection(SEARCH_RECORD_CONFIG.activeAfterActionCell);
  } catch (err) {}

  try {
    target.activate();
  } catch (err) {}

  SpreadsheetApp.flush();
}


function resetSearchRecordActionCheckboxes_(sheet) {
  const checkbox = sheet.getRange(SEARCH_RECORD_CONFIG.actionClearCell);

  // Keep the checkbox static. Do not rebuild/remove validation here.
  // Just untick it after clear action is completed.
  checkbox
    .setValue(false)
    .setBackground('#ffffff')
    .setFontColor('#9AA0A6')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle')
    .setBorder(false, false, false, false, false, false);

  sheet.getRange(SEARCH_RECORD_CONFIG.actionClearLabelRange)
    .setValue('CLEAR SEARCH')
    .setBackground('#ffffff')
    .setFontColor('#1b5e20')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBorder(false, false, false, false, false, false);

  sheet.getRange('F4:H7')
    .setBackground('#ffffff')
    .setBorder(false, false, false, false, false, false);

  SpreadsheetApp.flush();
}

function handleSearchRecordSheetEdit_(e) {
  try {
    if (!e || !e.range) return false;

    const range = e.range;
    const sheet = range.getSheet();
    if (!sheet || sheet.getName() !== SEARCH_RECORD_CONFIG.sheetName) return false;

    const editedCell = range.getA1Notation();
    const isInputEdit = editedCell === SEARCH_RECORD_CONFIG.inputCell;
    const isClearAction = editedCell === SEARCH_RECORD_CONFIG.actionClearCell;

    if (!isInputEdit && !isClearAction) return false;

    const lock = LockService.getDocumentLock();
    let lockAcquired = false;

    try {
      lock.waitLock(30000);
      lockAcquired = true;

      // Move active selection away first, before search/clear starts.
      setSearchRecordActiveCell_(sheet);

      if (isInputEdit) {
        setSearchInputActiveStyle_(sheet);

        const rawInputOriginal = safeString_(sheet.getRange(SEARCH_RECORD_CONFIG.inputCell).getDisplayValue());
        const rawInput = stripSearchPlaceholder_(rawInputOriginal).trim();

        if (!rawInput) {
          clearSearchRecordResultsOnly_(sheet);
          applySearchInputPlaceholder_(sheet);
        } else {
          runSearchRecord();
        }

        SpreadsheetApp.flush();
        setSearchRecordActiveCell_(sheet);
        return true;
      }

      if (isClearAction) {
        const checkboxValue = range.getValue();
        const eventValue = String(e.value || '').toUpperCase();
        const isChecked = checkboxValue === true || eventValue === 'TRUE';

        if (!isChecked) {
          // Ignore untick/manual false edits. The checkbox stays static.
          setSearchRecordActiveCell_(sheet);
          return true;
        }

        clearSearchRecord();
        SpreadsheetApp.flush();

        resetSearchRecordActionCheckboxes_(sheet);
        setSearchRecordActiveCell_(sheet);
        return true;
      }

    } catch (err) {
      try {
        sheet.getRange(SEARCH_RECORD_CONFIG.statusCell).setValue('ACTION ERROR');
        colorSearchStatusCell_(sheet.getRange(SEARCH_RECORD_CONFIG.statusCell), 'ACTION ERROR');
        SpreadsheetApp.getActiveSpreadsheet().toast(
          'Search Record action failed: ' + err.message,
          'Secure Entry Search Record',
          8
        );
      } catch (noticeErr) {
        Logger.log('Search Record edit error notice failed: ' + noticeErr);
      }
      throw err;

    } finally {
      if (isClearAction) {
        resetSearchRecordActionCheckboxes_(sheet);
      }
      setSearchRecordActiveCell_(sheet);

      if (lockAcquired) {
        lock.releaseLock();
      }
    }

    return true;

  } catch (outerErr) {
    Logger.log('handleSearchRecordSheetEdit_ error: ' + outerErr);
    throw outerErr;
  }
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
  const timestampDisplayValues = normalizeSearchTimestampsForDisplay_(
    displayMatches.map(function (item) { return item.timestamp; })
  );

  const values = displayMatches.map(function (item, index) {
    return [
      timestampDisplayValues[index],
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

    sheet.getRange(startRow, SEARCH_RECORD_CONFIG.historyLeftStartCol, values.length, 1)
      .setNumberFormat('@');

    sheet.getRange(startRow, SEARCH_RECORD_CONFIG.historyLeftStartCol, values.length, SEARCH_RECORD_CONFIG.historyLeftColumnCount)
      .setValues(leftValues);

    sheet.getRange(startRow, SEARCH_RECORD_CONFIG.historyRightStartCol, values.length, SEARCH_RECORD_CONFIG.historyRightColumnCount)
      .setValues(rightValues);

    // Result alignment:
    // All result columns = left, except MYKAD / PASSPORT = center.
    sheet.getRange(startRow, SEARCH_RECORD_CONFIG.historyLeftStartCol, values.length, SEARCH_RECORD_CONFIG.historyLeftColumnCount)
      .setHorizontalAlignment('left');

    sheet.getRange(startRow, SEARCH_RECORD_CONFIG.historyRightStartCol, values.length, SEARCH_RECORD_CONFIG.historyRightColumnCount)
      .setHorizontalAlignment('left');

    sheet.getRange(startRow, SEARCH_RECORD_CONFIG.historyLeftStartCol, values.length, 1)
      .setNumberFormat('@');

    sheet.getRange(startRow, SEARCH_RECORD_CONFIG.historyLeftStartCol + 2, values.length, 1)
      .setNumberFormat('@')
      .setHorizontalAlignment('center');

    sheet.getRange(startRow, SEARCH_RECORD_CONFIG.historyLeftStartCol + 3, values.length, 1)
      .setNumberFormat('@');

    sheet.getRange(startRow, SEARCH_RECORD_CONFIG.historyRightStartCol, values.length, 1)
      .setNumberFormat('@');

    applyHistoryRowColors_(sheet, displayMatches);
  }
}

function applyHistoryRowColors_(sheet, matches) {
  if (!matches.length) return;

  const leftColors = [];
  const spacerColors = [];
  const rightColors = [];

  matches.forEach(function (item, index) {
    const bg = index % 2 === 0 ? '#FFFFFF' : '#F3F7F3';
    leftColors.push([bg, bg, bg, bg]);
    spacerColors.push([bg]);
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
    6,
    spacerColors.length,
    1
  ).setBackgrounds(spacerColors);

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

  const dataRange = rawSheet.getRange(2, 1, lastRow - 1, 9);
  const values = dataRange.getValues();
  const displayValues = dataRange.getDisplayValues();
  const formulas = dataRange.getFormulas();

  const targetIndex = detectedField === 'REGNUM' ? 3 : 2;
  const matches = [];

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    const displayRow = displayValues[i];
    const formulaRow = formulas[i];
    const rawValue = row[targetIndex];
    const candidate = normalizeSearchKey_(rawValue);

    if (!candidate || candidate !== normalizedInput) continue;

    const directPhotoLink = safeString_(row[8]);
    const hyperlinkFromName = extractUrlFromHyperlinkFormulaSearch_(formulaRow[1]);
    const effectivePhotoLink = directPhotoLink || hyperlinkFromName || '';

    matches.push({
      rowNumber: i + 2,
      timestamp: displayRow[0] || row[0],
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
  // Keep Secure Entry timestamps in Malaysia display format (DD/MM/YYYY HH:mm:ss).
  // Do not pass text dates into generic parsers because ambiguous dates like 12/05/2026
  // may be interpreted as MM/DD/YYYY and displayed wrongly as 05/12/2026.
  const tz = Session.getScriptTimeZone() || 'Asia/Kuala_Lumpur';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, tz, 'dd/MM/yyyy HH:mm:ss');
  }

  const text = safeString_(value).trim();
  if (!text) return '';

  const match = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (match) {
    const day = String(match[1]).padStart(2, '0');
    const month = String(match[2]).padStart(2, '0');
    const year = match[3];
    const hour = String(match[4] || '00').padStart(2, '0');
    const minute = String(match[5] || '00').padStart(2, '0');
    const second = String(match[6] || '00').padStart(2, '0');
    return day + '/' + month + '/' + year + ' ' + hour + ':' + minute + ':' + second;
  }

  return text;
}

function normalizeSearchTimestampsForDisplay_(timestamps) {
  const now = new Date();
  const maxAllowedDate = new Date(now.getTime() + (24 * 60 * 60 * 1000));
  let previousDate = maxAllowedDate;

  return timestamps.map(function (value) {
    const normalizedText = normalizeSearchTimestamp_(value);
    const options = getSearchTimestampOptions_(normalizedText);

    if (!options.length) {
      return normalizedText;
    }

    const validInOrder = options
      .filter(function (item) {
        return item.date.getTime() <= previousDate.getTime() &&
          item.date.getTime() <= maxAllowedDate.getTime();
      })
      .sort(function (a, b) { return b.date.getTime() - a.date.getTime(); });

    const validNotFuture = options
      .filter(function (item) {
        return item.date.getTime() <= maxAllowedDate.getTime();
      })
      .sort(function (a, b) { return b.date.getTime() - a.date.getTime(); });

    const chosen = validInOrder[0] || validNotFuture[0] || options[0];
    previousDate = chosen.date;

    return chosen.text;
  });
}

function getSearchTimestampOptions_(text) {
  const match = safeString_(text).trim().match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (!match) return [];

  const first = parseInt(match[1], 10);
  const second = parseInt(match[2], 10);
  const year = parseInt(match[3], 10);
  const hour = parseInt(match[4] || '0', 10);
  const minute = parseInt(match[5] || '0', 10);
  const sec = parseInt(match[6] || '0', 10);

  const options = [];
  const ddmm = makeSearchTimestampOption_(first, second, year, hour, minute, sec);
  if (ddmm) options.push(ddmm);

  if (first <= 12 && second <= 12 && first !== second) {
    const mmddCorrected = makeSearchTimestampOption_(second, first, year, hour, minute, sec);
    if (mmddCorrected) {
      const duplicate = options.some(function (item) {
        return item.text === mmddCorrected.text;
      });
      if (!duplicate) options.push(mmddCorrected);
    }
  }

  return options;
}

function makeSearchTimestampOption_(day, month, year, hour, minute, sec) {
  const date = new Date(year, month - 1, day, hour, minute, sec);

  if (date.getFullYear() !== year ||
      date.getMonth() !== month - 1 ||
      date.getDate() !== day ||
      date.getHours() !== hour ||
      date.getMinutes() !== minute ||
      date.getSeconds() !== sec) {
    return null;
  }

  return {
    date: date,
    text: String(day).padStart(2, '0') + '/' +
      String(month).padStart(2, '0') + '/' +
      String(year) + ' ' +
      String(hour).padStart(2, '0') + ':' +
      String(minute).padStart(2, '0') + ':' +
      String(sec).padStart(2, '0')
  };
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
  if (handleSearchRecordSheetEdit_(e)) return;
}

/*
ONEDIT INTEGRATION NOTE:
Use handleSearchRecordEdit(e) as an installable On edit trigger for the SEARCH RECORD sheet.

Current flow:
- Edit C4 and press Enter to run search.
- Tick F7 to clear the search input and results.

If this Apps Script project already has an onEdit(e) function and you prefer to combine triggers, add this line at the top of that existing onEdit(e):

  if (handleSearchRecordSheetEdit_(e)) return;

Do not create a second simple onEdit(e) function in the same project.
*/
