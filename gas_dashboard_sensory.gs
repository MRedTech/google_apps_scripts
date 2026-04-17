const DASHBOARD_CONFIG = {
  sourceSheetName: 'SENSORY',
  dashboardSheetName: 'DASHBOARD',
  timezone: 'Asia/Kuala_Lumpur',
  headerTitle: 'SECURE ENTRY DASHBOARD | SENSORY',
  headerSubtitle: 'REGISTER.ACCESS.SECURE',
  filterStartCell: 'B3',
  filterEndCell: 'D3',
  helperStartColumn: 14, // N
  helperStartRow: 75
};

const CATEGORY_GROUPS = {
  'GRAB FOOD': 'FOOD DELIVERY',
  'FOOD PANDA': 'FOOD DELIVERY',
  'SHOPEE FOOD': 'FOOD DELIVERY',
  'MCD FOOD DELIVERY': 'FOOD DELIVERY',
  'RESTAURANT DELIVERY': 'FOOD DELIVERY',
  'DOMINOS DELIVERY': 'FOOD DELIVERY',
  'DHL': 'PARCEL / COURIER',
  'FLASH EXPRESS': 'PARCEL / COURIER',
  'GDEX': 'PARCEL / COURIER',
  'J&T EXPRESS': 'PARCEL / COURIER',
  'NINJA VAN': 'PARCEL / COURIER',
  'POS LAJU': 'PARCEL / COURIER',
  'SHOPEE EXPRESS': 'PARCEL / COURIER',
  'LAZADA': 'PARCEL / COURIER',
  'LALAMOVE': 'PARCEL / COURIER',
  'ABX EXPRESS': 'PARCEL / COURIER',
  'CITY LINK': 'PARCEL / COURIER',
  'FEDEX': 'PARCEL / COURIER',
  'SKYNET': 'PARCEL / COURIER',
  'E-HAILING': 'E-HAILING / DROP OFF / PICK UP',
  'DROP OFF / PICK UP': 'E-HAILING / DROP OFF / PICK UP',
  'OWNER': 'OWNER / TENANT',
  'TENANT': 'OWNER / TENANT',
  'OTHER': 'OTHERS'
};

const ACTIVE_GROUPS = [
  'FOOD DELIVERY',
  'PARCEL / COURIER',
  'E-HAILING / DROP OFF / PICK UP'
];

const TOWER_GROUP_ORDER = ACTIVE_GROUPS.slice();

function onOpen() {
  hideSensitiveSheets_();
}

function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(DASHBOARD_CONFIG.sourceSheetName);
  if (!sourceSheet) throw new Error("Source sheet 'SENSORY' was not found.");

  const dashboardSheet = getOrCreateSheet_(ss, DASHBOARD_CONFIG.dashboardSheetName);
  initializeDashboardLayout_(dashboardSheet, sourceSheet);
  refreshDashboard();
  applyDashboardProtection_();
  applySensoryProtection_();
}

function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(DASHBOARD_CONFIG.sourceSheetName);
  const dashboardSheet = ss.getSheetByName(DASHBOARD_CONFIG.dashboardSheetName);

  if (!sourceSheet) throw new Error("Source sheet 'SENSORY' was not found.");
  if (!dashboardSheet) throw new Error("Dashboard sheet 'DASHBOARD' was not found. Please run buildDashboard() first.");

  const allRows = getSourceRows_(sourceSheet);
  if (!allRows.length) {
    writeEmptyDashboardState_(dashboardSheet);
    return;
  }

  const dateFilter = getDashboardDateFilter_(dashboardSheet, allRows);
  const filteredRows = filterRowsByDate_(allRows, dateFilter.startKey, dateFilter.endKey);
  const stats = buildStatistics_(filteredRows, allRows);

  writeDashboardValues_(dashboardSheet, stats, dateFilter);
  writeHelperTables_(dashboardSheet, stats);
  writeTowerGroupCategoryVisible_(dashboardSheet, stats.towerDistribution);
  buildCharts_(dashboardSheet, stats);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function initializeDashboardLayout_(sheet, sourceSheet) {
  removeAllCharts_(sheet);
  sheet.clear();
  sheet.getRange('A1:L95').breakApart();
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(4);

  if (sheet.getMaxRows() < 140) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 140 - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < 45) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 45 - sheet.getMaxColumns());
  }

  sheet.getRange('A1:L95')
    .setBackground('#ffffff')
    .setBorder(false, false, false, false, false, false);

  sheet.setColumnWidths(1, 12, 105);
  sheet.setColumnWidth(1, 55);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 20);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 20);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 20);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 20);
  sheet.setColumnWidth(10, 120);
  sheet.setColumnWidth(11, 20);
  sheet.setColumnWidth(12, 135);

  sheet.getRange('A1:L1').merge();
  sheet.getRange('A1')
    .setValue(DASHBOARD_CONFIG.headerTitle)
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setBackground('#e8f5e9')
    .setFontColor('#1b5e20');

  sheet.getRange('A2:F2').merge();
  sheet.getRange('A2:F2')
    .setValue(DASHBOARD_CONFIG.headerSubtitle)
    .setFontStyle('italic')
    .setFontSize(10)
    .setFontColor('#558b2f')
    .setHorizontalAlignment('left');

  sheet.getRange('A3').setValue('FILTER START DATE').setFontWeight('bold').setFontColor('#1b5e20');
  sheet.getRange('C3').setValue('FILTER END DATE').setFontWeight('bold').setFontColor('#1b5e20');
  sheet.getRange('J2').setValue('LAST UPDATED').setFontWeight('bold').setFontColor('#1b5e20');
  sheet.getRange('K2:L2').merge();
  sheet.getRange('K2:L2').setHorizontalAlignment('left');

  const firstDate = getEarliestDateFromSource_(sourceSheet);
  const lastDate = getLatestDateFromSource_(sourceSheet);

  sheet.getRange(DASHBOARD_CONFIG.filterStartCell).setValue(firstDate || new Date()).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(DASHBOARD_CONFIG.filterEndCell).setValue(lastDate || new Date()).setNumberFormat('dd/MM/yyyy');

  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  sheet.getRange(DASHBOARD_CONFIG.filterStartCell).setDataValidation(dateRule);
  sheet.getRange(DASHBOARD_CONFIG.filterEndCell).setDataValidation(dateRule);

  styleFilterCell_(sheet.getRange(DASHBOARD_CONFIG.filterStartCell));
  styleFilterCell_(sheet.getRange(DASHBOARD_CONFIG.filterEndCell));

  createCardFrames_(sheet);
  createSectionFrames_(sheet);

  for (let r = 5; r <= 95; r++) {
    sheet.setRowHeight(r, 24);
  }

  sheet.hideColumns(DASHBOARD_CONFIG.helperStartColumn, 27);
}

function styleFilterCell_(range) {
  range
    .setBackground('#f1f8e9')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#a5d6a7', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function createCardFrames_(sheet) {
  const cards = [
    { title: 'TOTAL REGISTRATION', titleRange: 'A5:B5', valueRange: 'A6:B6', noteRange: 'A7:B7' },
    { title: 'PEAK HOUR', titleRange: 'C5:D5', valueRange: 'C6:D6', noteRange: 'C7:D7' },
    { title: 'RECORD FOUND RATE', titleRange: 'E5:F5', valueRange: 'E6:F6', noteRange: 'E7:F7' },
    { title: 'NEW RECORD RATE', titleRange: 'G5:H5', valueRange: 'G6:H6', noteRange: 'G7:H7' }
  ];

  cards.forEach(card => {
    sheet.getRange(card.titleRange).merge();
    sheet.getRange(card.valueRange).merge();
    sheet.getRange(card.noteRange).merge();

    sheet.getRange(card.titleRange)
      .setValue(card.title)
      .setBackground('#f7fbf7')
      .setFontSize(9)
      .setFontWeight('bold')
      .setFontColor('#2e7d32')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, false, true, true, true, '#c8e6c9', SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(card.valueRange)
      .setBackground('#f7fbf7')
      .setFontSize(18)
      .setFontWeight('bold')
      .setFontColor('#1b5e20')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(false, true, false, true, true, true, '#c8e6c9', SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(card.noteRange)
      .setBackground('#f7fbf7')
      .setFontSize(9)
      .setFontColor('#558b2f')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(false, true, true, true, true, true, '#c8e6c9', SpreadsheetApp.BorderStyle.SOLID);
  });
}

function createSectionFrames_(sheet) {
  const sections = [
    { title: 'REGISTRATION BY GROUP CATEGORY & TOWER', range: 'I5:L15', titleCell: 'I5' },
    { title: 'HOURLY REGISTRATION TREND', range: 'A17:F33', titleCell: 'A17' },
    { title: 'MONTHLY REGISTRATION TREND', range: 'G17:L33', titleCell: 'G17' },
    { title: 'CATEGORY BREAKDOWN', range: 'A35:F53', titleCell: 'A35' },
    { title: 'GROUP CATEGORY BREAKDOWN', range: 'G35:L53', titleCell: 'G35' },
    { title: 'TOWER COMPARISON BY GROUP CATEGORY', range: 'A55:F72', titleCell: 'A55' },
    { title: 'RECORD STATUS', range: 'G55:L72', titleCell: 'G55' }
  ];

  sections.forEach(section => {
    sheet.getRange(section.range)
      .setBorder(true, true, true, true, true, true, '#dcedc8', SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(section.titleCell)
      .setValue(section.title)
      .setFontWeight('bold')
      .setFontColor('#1b5e20')
      .setBackground('#f1f8e9');
  });
}

function getSourceRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const tsDisplay = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  const otherValues = sheet.getRange(2, 2, lastRow - 1, 8).getValues();
  const rows = [];

  for (let i = 0; i < lastRow - 1; i++) {
    const timestamp = normalizeDateValue_(tsDisplay[i][0]);
    if (!(timestamp instanceof Date) || isNaN(timestamp.getTime())) continue;

    const row = otherValues[i];
    rows.push({
      rowIndex: i + 2,
      timestamp,
      name: safeUpper_(row[0]),
      idNo: safeString_(row[1]),
      regNum: safeUpper_(row[2]),
      contact: safeString_(row[3]),
      category: normalizeCategory_(row[4]),
      tower: normalizeTower_(row[5]),
      reason: safeString_(row[6]),
      photoLink: safeString_(row[7])
    });
  }

  return rows;
}

function normalizeDateValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return new Date(
      value.getFullYear(),
      value.getMonth(),
      value.getDate(),
      value.getHours(),
      value.getMinutes(),
      value.getSeconds()
    );
  }

  if (typeof value === 'string' && value.trim()) {
    const text = value.trim();
    const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (match) {
      const day = parseInt(match[1], 10);
      const month = parseInt(match[2], 10) - 1;
      const year = parseInt(match[3], 10);
      const hour = parseInt(match[4] || '0', 10);
      const minute = parseInt(match[5] || '0', 10);
      const second = parseInt(match[6] || '0', 10);

      const parsed = new Date(year, month, day, hour, minute, second);
      if (
        parsed.getFullYear() === year &&
        parsed.getMonth() === month &&
        parsed.getDate() === day &&
        parsed.getHours() === hour &&
        parsed.getMinutes() === minute &&
        parsed.getSeconds() === second
      ) {
        return parsed;
      }
    }
  }

  return null;
}

function normalizeTower_(value) {
  const tower = safeUpper_(value);
  if (tower === 'TOWER A') return 'Tower A';
  if (tower === 'TOWER B') return 'Tower B';
  if (tower === 'TOWER A & B') return 'Tower A & B';
  return safeString_(value);
}

function normalizeCategory_(value) {
  let text = safeUpper_(value);
  if (!text) return '';

  text = text
    .replace(/\s+/g, ' ')
    .replace(/\s*\/\s*/g, ' / ')
    .replace(/\s*-\s*/g, '-')
    .trim();

  if (text === 'E HAILING' || text === 'EHAILING' || text === 'E-HAILING') {
    return 'E-HAILING';
  }

  if (
    text === 'DROP OFF/PICK UP' ||
    text === 'DROP OFF / PICK UP' ||
    text === 'DROP-OFF / PICK-UP' ||
    text === 'DROP OFF / PICK-UP' ||
    text === 'DROP-OFF / PICK UP' ||
    text === 'DROPOFF/PICKUP' ||
    text === 'DROP OFF' ||
    text === 'DROP-OFF' ||
    text === 'PICK UP' ||
    text === 'PICK-UP'
  ) {
    return 'DROP OFF / PICK UP';
  }

  if (text === 'DOMINOS' || text === "DOMINO'S DELIVERY" || text === 'DOMINOS DELIVERY') {
    return 'DOMINOS DELIVERY';
  }

  if (text === 'ABX' || text === 'ABX EXPRESS') {
    return 'ABX EXPRESS';
  }

  if (text === 'CITYLINK' || text === 'CITY LINK') {
    return 'CITY LINK';
  }

  if (text === 'FED EX' || text === 'FEDEX') {
    return 'FEDEX';
  }

  if (text === 'SKY NET' || text === 'SKYNET') {
    return 'SKYNET';
  }

  if (text === 'OTHERS') {
    return 'OTHER';
  }

  const knownCategories = [
    'GRAB FOOD',
    'FOOD PANDA',
    'SHOPEE FOOD',
    'MCD FOOD DELIVERY',
    'RESTAURANT DELIVERY',
    'DOMINOS DELIVERY',
    'DHL',
    'FLASH EXPRESS',
    'GDEX',
    'J&T EXPRESS',
    'NINJA VAN',
    'POS LAJU',
    'SHOPEE EXPRESS',
    'LAZADA',
    'LALAMOVE',
    'ABX EXPRESS',
    'CITY LINK',
    'FEDEX',
    'SKYNET',
    'E-HAILING',
    'DROP OFF / PICK UP',
    'OWNER',
    'TENANT',
    'OTHER'
  ];

  for (let i = 0; i < knownCategories.length; i++) {
    const category = knownCategories[i];
    if (text === category || text.startsWith(category + ' ') || text.startsWith(category + '(')) {
      return category;
    }
  }

  return text;
}

function parseDashboardFilterDate_(sheet, a1Notation) {
  const text = safeString_(sheet.getRange(a1Notation).getDisplayValue());
  return normalizeDateValue_(text);
}

function getDashboardDateFilter_(sheet, allRows) {
  const earliest = toDateOnly_(allRows.reduce((min, row) => row.timestamp < min ? row.timestamp : min, allRows[0].timestamp));
  const latest = toDateOnly_(allRows.reduce((max, row) => row.timestamp > max ? row.timestamp : max, allRows[0].timestamp));

  let start = parseDashboardFilterDate_(sheet, DASHBOARD_CONFIG.filterStartCell) || earliest;
  let end = parseDashboardFilterDate_(sheet, DASHBOARD_CONFIG.filterEndCell) || latest;

  start = toDateOnly_(start);
  end = toDateOnly_(end);

  if (dateKey_(start) > dateKey_(end)) {
    const temp = start;
    start = end;
    end = temp;
  }

  sheet.getRange(DASHBOARD_CONFIG.filterStartCell).setValue(start).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(DASHBOARD_CONFIG.filterEndCell).setValue(end).setNumberFormat('dd/MM/yyyy');

  return {
    start,
    end,
    startKey: dateKey_(start),
    endKey: dateKey_(end),
    startLabel: displayDate_(start),
    endLabel: displayDate_(end)
  };
}

function filterRowsByDate_(rows, startKey, endKey) {
  return rows.filter(row => {
    const rowKey = dateKey_(row.timestamp);
    return rowKey >= startKey && rowKey <= endKey;
  });
}

function buildStatistics_(filteredRows, allRows) {
  const statsRows = filteredRows.filter(row => ACTIVE_GROUPS.includes(CATEGORY_GROUPS[row.category] || ''));
  const statsAllRows = allRows.filter(row => ACTIVE_GROUPS.includes(CATEGORY_GROUPS[row.category] || ''));

  const hourlyTrend = aggregatePeakHour_(statsRows);
  const monthly = aggregateMonthlyGroupLast3Months_(statsAllRows);
  const category = aggregateCountByLabel_(statsRows, row => row.category || 'UNKNOWN');
  const groupCategory = aggregateCountByLabel_(statsRows, row => CATEGORY_GROUPS[row.category] || '');
  const towerDistribution = buildTowerGroupBreakdown_(statsRows);
  const recordStatus = buildRecordStatus_(statsRows);

  const peakHourTop = hourlyTrend
    .slice()
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .slice(0, 1);

  const totalFiltered = statsRows.length;
  const existingCount = recordStatus.find(r => r[0] === 'EXISTING RECORD')?.[1] || 0;
  const newCount = recordStatus.find(r => r[0] === 'NEW RECORD')?.[1] || 0;
  const foundRate = totalFiltered ? `${Math.round((existingCount / totalFiltered) * 100)}%` : '0%';
  const newRate = totalFiltered ? `${Math.round((newCount / totalFiltered) * 100)}%` : '0%';
  const topPeakHour = peakHourTop.length ? peakHourTop[0][0] : '-';

  return {
    totalFiltered,
    peakHour: topPeakHour,
    foundRate,
    newRate,
    hourlyTrend,
    monthly,
    category,
    groupCategory,
    towerDistribution,
    recordStatus
  };
}

function aggregateMonthlyGroupLast3Months_(rows) {
  const now = new Date();
  const currentMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const monthNames = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];

  const buckets = [];
  const bucketMap = new Map();

  for (let i = 2; i >= 0; i--) {
    const d = new Date(currentMonth.getFullYear(), currentMonth.getMonth() - i, 1);
    const key = d.getFullYear() * 100 + (d.getMonth() + 1);
    const label = monthNames[d.getMonth()] + ' ' + d.getFullYear();
    const bucket = [label, 0, 0, 0];
    buckets.push(bucket);
    bucketMap.set(key, bucket);
  }

  rows.forEach(row => {
    const d = row.timestamp;
    const key = d.getFullYear() * 100 + (d.getMonth() + 1);
    const group = CATEGORY_GROUPS[row.category] || '';
    if (!bucketMap.has(key)) return;

    const bucket = bucketMap.get(key);
    if (group === 'FOOD DELIVERY') {
      bucket[1] += 1;
    } else if (group === 'E-HAILING / DROP OFF / PICK UP') {
      bucket[2] += 1;
    } else if (group === 'PARCEL / COURIER') {
      bucket[3] += 1;
    }
  });

  return buckets;
}

function aggregateCountByLabel_(rows, labelFn) {
  const map = new Map();
  rows.forEach(row => {
    const label = labelFn(row);
    map.set(label, (map.get(label) || 0) + 1);
  });
  return Array.from(map.entries()).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
}

function aggregatePeakHour_(rows) {
  const map = new Map();
  rows.forEach(row => {
    const label = Utilities.formatDate(row.timestamp, DASHBOARD_CONFIG.timezone, 'HH:00');
    map.set(label, (map.get(label) || 0) + 1);
  });
  return Array.from(map.entries()).sort((a, b) => a[0].localeCompare(b[0]));
}

function buildTowerGroupBreakdown_(rows) {
  const result = Object.fromEntries(TOWER_GROUP_ORDER.map(group => [group, { 'Tower A': 0, 'Tower B': 0 }]));

  rows.forEach(row => {
    const group = CATEGORY_GROUPS[row.category] || 'OTHERS';
    if (row.tower === 'Tower A' || row.tower === 'Tower A & B') {
      result[group]['Tower A'] += 1;
    }
    if (row.tower === 'Tower B' || row.tower === 'Tower A & B') {
      result[group]['Tower B'] += 1;
    }
  });

  return [
    ['GROUP CATEGORY', 'Tower A', 'Tower B'],
    ...TOWER_GROUP_ORDER.map(group => [group, result[group]['Tower A'], result[group]['Tower B']])
  ];
}

function buildRecordStatus_(rows) {
  let newRecordCount = 0;
  let existingRecordCount = 0;

  rows.forEach(row => {
    if (row.photoLink) newRecordCount += 1;
    else existingRecordCount += 1;
  });

  return [
    ['NEW RECORD', newRecordCount],
    ['EXISTING RECORD', existingRecordCount]
  ];
}

function addAnnotationRows_(rows) {
  return rows.map(row => [row[0], row[1], String(row[1])]);
}

function writeDashboardValues_(sheet, stats, dateFilter) {
  sheet.getRange('K2').setValue(Utilities.formatDate(new Date(), DASHBOARD_CONFIG.timezone, 'dd/MM/yyyy HH:mm:ss'));

  const values = {
    'A6:B6': stats.totalFiltered,
    'C6:D6': stats.peakHour || '-',
    'E6:F6': stats.foundRate,
    'G6:H6': stats.newRate
  };

  const notes = {
    'A7:B7': `${dateFilter.startLabel} - ${dateFilter.endLabel}`,
    'C7:D7': 'Highest peak hour within the selected filter',
    'E7:F7': 'Existing records / total within filter',
    'G7:H7': 'New records / total within filter'
  };

  Object.keys(values).forEach(rangeA1 => sheet.getRange(rangeA1).setValue(values[rangeA1]));
  Object.keys(notes).forEach(rangeA1 => sheet.getRange(rangeA1).setValue(notes[rangeA1]));
}

function writeTowerGroupCategoryVisible_(sheet, towerDistribution) {
  const startRow = 5;
  const startCol = 9; // Column I
  const totalCols = 4;
  const clearRows = 12;

  sheet.getRange(startRow, startCol, clearRows, totalCols)
    .clearContent()
    .clearFormat()
    .breakApart();

  sheet.getRange(startRow, startCol, 1, totalCols)
    .merge()
    .setValue('REGISTRATION BY GROUP CATEGORY & TOWER')
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontColor('#1b5e20')
    .setBackground('#f1f8e9')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange(startRow + 1, startCol, 1, 2)
    .merge()
    .setValue('GROUP CATEGORY')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange(startRow + 1, startCol + 2)
    .setValue('A')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange(startRow + 1, startCol + 3)
    .setValue('B')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  let r = startRow + 2;
  const rows = towerDistribution.slice(1);

  rows.forEach(item => {
    const category = item[0];
    const towerA = item[1] || 0;
    const towerB = item[2] || 0;

    sheet.getRange(r, startCol, 1, 2)
      .merge()
      .setValue(category)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setFontSize(10);

    sheet.getRange(r, startCol + 2)
      .setValue(towerA)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontSize(10);

    sheet.getRange(r, startCol + 3)
      .setValue(towerB)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontSize(10);

    r++;
  });

  const totalA = rows.reduce((sum, item) => sum + (item[1] || 0), 0);
  const totalB = rows.reduce((sum, item) => sum + (item[2] || 0), 0);
  const grandTotal = totalA + totalB;

  sheet.getRange(r, startCol, 1, 2)
    .merge()
    .setValue('TOTAL TOWER')
    .setFontWeight('bold')
    .setBackground('#f7fbf7')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange(r, startCol + 2)
    .setValue(totalA)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange(r, startCol + 3)
    .setValue(totalB)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  r++;

  sheet.getRange(r, startCol, 1, 3)
    .merge()
    .setValue('GRAND TOTAL')
    .setFontWeight('bold')
    .setBackground('#f1f8e9')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange(r, startCol + 3)
    .setValue(grandTotal)
    .setFontWeight('bold')
    .setFontColor('#1b5e20')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  r++;

  sheet.getRange(r, startCol, 1, totalCols)
    .merge()
    .setValue('Note: Grand Total under Tower Summary may be higher than Total Registration because records tagged as "Tower A & B" are counted under both towers for distribution analysis.')
    .setFontSize(8)
    .setFontStyle('italic')
    .setFontColor('#616161')
    .setWrap(true)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.setRowHeight(r, 42);
  r++;

  const totalRows = r - startRow;

  sheet.getRange(startRow, startCol, totalRows, totalCols)
    .setBorder(true, true, true, true, true, true, '#dcedc8', SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle');

  sheet.setColumnWidth(startCol, 170);
  sheet.setColumnWidth(startCol + 1, 135);
  sheet.setColumnWidth(startCol + 2, 60);
  sheet.setColumnWidth(startCol + 3, 60);
}

function writeHelperTables_(sheet, stats) {
  const startCol = DASHBOARD_CONFIG.helperStartColumn;
  const startRow = DASHBOARD_CONFIG.helperStartRow;

  sheet.getRange(startRow, startCol, 120, 24).clearContent().clearFormat();

  const tables = [
    { title: 'HOURLY', headers: ['HOUR', 'TOTAL'], rows: stats.hourlyTrend, colOffset: 0 },
    { title: 'MONTHLY', headers: ['MONTH', 'FOOD DELIVERY', 'E-HAILING', 'PARCEL'], rows: stats.monthly, colOffset: 3, width: 4 },
    { title: 'CATEGORY', headers: ['CATEGORY', 'TOTAL', 'ANNOTATION'], rows: addAnnotationRows_(stats.category), colOffset: 7, width: 3 },
    { title: 'GROUP CATEGORY', headers: ['GROUP', 'TOTAL'], rows: stats.groupCategory, colOffset: 11 },
    { title: 'TOWER DISTRIBUTION', headers: ['GROUP CATEGORY', 'Tower A', 'Tower B'], rows: stats.towerDistribution.slice(1), colOffset: 14, width: 3 },
    { title: 'RECORD STATUS', headers: ['STATUS', 'TOTAL'], rows: stats.recordStatus, colOffset: 18 }
  ];

  tables.forEach(table => {
    const col = startCol + table.colOffset;
    const width = table.width || 2;

    sheet.getRange(startRow, col).setValue(table.title).setFontWeight('bold');
    sheet.getRange(startRow + 1, col, 1, width)
      .setValues([table.headers])
      .setFontWeight('bold')
      .setBackground('#e8f5e9');

    if (table.rows && table.rows.length) {
      sheet.getRange(startRow + 2, col, table.rows.length, width).setValues(table.rows);
      sheet.getRange(startRow + 2, col, table.rows.length, 1).setNumberFormat('@');

      if (table.title === 'CATEGORY') {
        sheet.getRange(startRow + 2, col + 1, table.rows.length, 1).setNumberFormat('0');
        sheet.getRange(startRow + 2, col + 2, table.rows.length, 1).setNumberFormat('@');
      } else if (width > 1) {
        sheet.getRange(startRow + 2, col + 1, table.rows.length, width - 1).setNumberFormat('0');
      }
    }
  });
}

function buildCharts_(sheet, stats) {
  removeAllCharts_(sheet);

  const startRow = DASHBOARD_CONFIG.helperStartRow;
  const startCol = DASHBOARD_CONFIG.helperStartColumn;

  const hourlyRange = sheet.getRange(startRow + 1, startCol, Math.max(stats.hourlyTrend.length + 1, 2), 2);
  const monthlyRange = sheet.getRange(startRow + 1, startCol + 3, Math.max(stats.monthly.length + 1, 2), 4);
  const categoryRange = sheet.getRange(startRow + 1, startCol + 7, Math.max(stats.category.length + 1, 2), 3);
  const groupRange = sheet.getRange(startRow + 1, startCol + 11, Math.max(stats.groupCategory.length + 1, 2), 2);
  const comparisonRange = sheet.getRange(startRow + 1, startCol + 14, stats.towerDistribution.length, 3);
  const recordRange = sheet.getRange(startRow + 1, startCol + 18, Math.max(stats.recordStatus.length + 1, 2), 2);

  const maxHourly = Math.max(...stats.hourlyTrend.map(r => r[1]), 0);
  const maxMonthly = Math.max(...stats.monthly.flatMap(r => [r[1], r[2], r[3]]), 0);
  const maxCategory = Math.max(...stats.category.map(r => r[1]), 0);
  const maxTower = Math.max(...stats.towerDistribution.slice(1).flatMap(r => [r[1], r[2]]), 0);

  const hourlyStep = getAutoStep_(maxHourly);
  const monthlyStep = getAutoStep_(maxMonthly);
  const categoryStep = getAutoStep_(maxCategory);
  const towerStep = getAutoStep_(maxTower);

  const hourlyMax = roundUpToStep_(maxHourly, hourlyStep);
  const monthlyMax = roundUpToStep_(maxMonthly, monthlyStep);
  const categoryMax = roundUpToStep_(maxCategory, categoryStep);
  const towerMax = roundUpToStep_(maxTower, towerStep);

  const charts = [
    sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(hourlyRange)
      .setOption('title', 'Hourly Registration Trend')
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { slantedText: true, slantedTextAngle: 45 })
      .setOption('vAxis', {
        textStyle: { fontSize: 10 },
        viewWindow: { min: 0, max: hourlyMax },
        gridlines: { count: Math.floor(hourlyMax / hourlyStep) + 1 },
        minorGridlines: { count: 0 }
      })
      .setPosition(18, 1, 0, 0)
      .setNumHeaders(1)
      .build(),

    sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(monthlyRange)
      .setOption('title', 'Monthly Registration Trend by Group Category')
      .setOption('legend', { position: 'top', textStyle: { fontSize: 10 } })
      .setOption('isStacked', false)
      .setOption('series', {
        0: { color: '#DB4437' },
        1: { color: '#4285F4' },
        2: { color: '#F4B400' }
      })
      .setOption('vAxis', {
        textStyle: { fontSize: 10 },
        viewWindow: { min: 0, max: monthlyMax },
        gridlines: { count: Math.floor(monthlyMax / monthlyStep) + 1 },
        minorGridlines: { count: 0 }
      })
      .setPosition(18, 7, 0, 0)
      .setNumHeaders(1)
      .build(),

    sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(categoryRange)
      .setOption('title', 'Category Breakdown')
      .setOption('legend', { position: 'none' })
      .setOption('annotations', { alwaysOutside: true, textStyle: { fontSize: 10, color: '#1b5e20', auraColor: 'none' } })
      .setOption('hAxis', {
        viewWindow: { min: 0, max: categoryMax },
        gridlines: { count: Math.floor(categoryMax / categoryStep) + 1 },
        minorGridlines: { count: 0 }
      })
      .setPosition(36, 1, 0, 0)
      .setNumHeaders(1)
      .build(),

    sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(groupRange)
      .setOption('title', 'Group Category Breakdown')
      .setOption('pieHole', 0.45)
      .setOption('legend', { position: 'labeled', textStyle: { color: '#1b5e20' } })
      .setOption('pieSliceTextStyle', { color: '#1b5e20' })
      .setOption('sliceVisibilityThreshold', 0)
      .setPosition(36, 7, 0, 0)
      .setNumHeaders(1)
      .build(),

    sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(comparisonRange)
      .setOption('title', 'Tower Comparison by Group Category')
      .setOption('legend', { position: 'top' })
      .setOption('isStacked', false)
      .setOption('series', {
        0: { color: '#4285F4' },
        1: { color: '#0F9D58' }
      })
      .setOption('hAxis', {
        textStyle: { fontSize: 10 },
        viewWindow: { min: 0, max: towerMax },
        gridlines: { count: Math.floor(towerMax / towerStep) + 1 },
        minorGridlines: { count: 0 }
      })
      .setOption('vAxis', {
        textStyle: { fontSize: 10 }
      })
      .setPosition(56, 1, 0, 0)
      .setNumHeaders(1)
      .build(),

    sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(recordRange)
      .setOption('title', 'Record Status')
      .setOption('pieHole', 0.45)
      .setOption('legend', { position: 'labeled', textStyle: { color: '#1b5e20' } })
      .setOption('pieSliceTextStyle', { color: '#1b5e20' })
      .setOption('sliceVisibilityThreshold', 0)
      .setPosition(56, 7, 0, 0)
      .setNumHeaders(1)
      .build()
  ];

  charts.forEach(chart => sheet.insertChart(chart));
}

function writeEmptyDashboardState_(sheet) {
  removeAllCharts_(sheet);
  sheet.getRange('A6:B6').setValue(0);
  sheet.getRange('C6:D6').setValue('-');
  sheet.getRange('E6:F6').setValue('0%');
  sheet.getRange('G6:H6').setValue('0%');
  sheet.getRange('K2').setValue('NO DATA');
  sheet.getRange('I5:L15').clearContent();
}

function getEarliestDateFromSource_(sheet) {
  const rows = getSourceRows_(sheet);
  if (!rows.length) return null;
  return toDateOnly_(rows.reduce((min, row) => row.timestamp < min ? row.timestamp : min, rows[0].timestamp));
}

function getLatestDateFromSource_(sheet) {
  const rows = getSourceRows_(sheet);
  if (!rows.length) return null;
  return toDateOnly_(rows.reduce((max, row) => row.timestamp > max ? row.timestamp : max, rows[0].timestamp));
}

function dateKey_(date) {
  const d = normalizeDateValue_(date);
  if (!d) return 0;
  return d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate();
}

function toDateOnly_(date) {
  const d = normalizeDateValue_(date);
  if (!d) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function displayDate_(date) {
  const d = normalizeDateValue_(date);
  if (!d) return '';
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return day + '/' + month + '/' + year;
}

function roundUpToStep_(value, step) {
  if (!value || value < 0) return step;
  return Math.ceil(value / step) * step;
}

function getAutoStep_(maxValue) {
  if (maxValue <= 50) return 5;
  if (maxValue <= 100) return 10;
  if (maxValue <= 250) return 25;
  if (maxValue <= 500) return 50;
  if (maxValue <= 1000) return 100;
  return 200;
}

function safeString_(value) {
  return value == null ? '' : String(value).trim();
}

function safeUpper_(value) {
  return safeString_(value).toUpperCase();
}

function removeAllCharts_(sheet) {
  const charts = sheet.getCharts();
  charts.forEach(chart => sheet.removeChart(chart));
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== DASHBOARD_CONFIG.dashboardSheetName) return;

    const a1 = e.range.getA1Notation();
    const isStart = a1 === DASHBOARD_CONFIG.filterStartCell;
    const isEnd = a1 === DASHBOARD_CONFIG.filterEndCell;

    if (!isStart && !isEnd) return;

    refreshDashboard();

    styleFilterCell_(sheet.getRange(DASHBOARD_CONFIG.filterStartCell));
    styleFilterCell_(sheet.getRange(DASHBOARD_CONFIG.filterEndCell));

    sheet.getRange(DASHBOARD_CONFIG.filterStartCell).setNumberFormat('dd/MM/yyyy');
    sheet.getRange(DASHBOARD_CONFIG.filterEndCell).setNumberFormat('dd/MM/yyyy');

  } catch (err) {
    Logger.log('onEdit error: ' + err);
  }
}

function applyDashboardProtection_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_CONFIG.dashboardSheetName);
  if (!sheet) throw new Error("Dashboard sheet was not found.");

  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  protections.forEach(p => {
    if (p.canEdit()) p.remove();
  });

  const protection = sheet.protect().setDescription('Lock dashboard except B3 and D3');

  protection.setUnprotectedRanges([
    sheet.getRange(DASHBOARD_CONFIG.filterStartCell),
    sheet.getRange(DASHBOARD_CONFIG.filterEndCell)
  ]);

  const me = Session.getEffectiveUser();
  protection.addEditor(me);

  const otherEditors = protection.getEditors().filter(user => user.getEmail() !== me.getEmail());
  if (otherEditors.length) {
    protection.removeEditors(otherEditors);
  }

  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}

function applySensoryProtection_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_CONFIG.sourceSheetName);
  if (!sheet) throw new Error("Source sheet 'SENSORY' was not found.");

  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  protections.forEach(p => {
    if (p.canEdit()) p.remove();
  });

  const protection = sheet.protect().setDescription('Lock full SENSORY sheet');

  const me = Session.getEffectiveUser();
  protection.addEditor(me);

  const otherEditors = protection.getEditors().filter(user => user.getEmail() !== me.getEmail());
  if (otherEditors.length) {
    protection.removeEditors(otherEditors);
  }

  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}
