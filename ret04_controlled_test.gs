// ==================================================
// TEMPORARY CONTROLLED TEST — NEW-RET-04
// Safe scope:
// - Creates temporary sheet "RET04 TEST"
// - Does NOT write/delete any row in SENSORY
// - Uses no Drive file IDs / URLs
// - Restores RETENTION_LAST_FULL_SCAN_MS after the test
// - Deletes the temporary sheet automatically
// Delete this function after PASS.
// ==================================================
function test_NEW_RET_04_OUT_OF_ORDER_() {
  const TEST_SHEET_NAME = "RET04 TEST";
  const mainSheet = getSheet_();
  const ss = mainSheet.getParent();
  const props = PropertiesService.getScriptProperties();

  const productionRowsBefore = mainSheet.getLastRow();
  const oldFullScanProp = props.getProperty(RETENTION_LAST_FULL_SCAN_PROP);

  let testSheet = ss.getSheetByName(TEST_SHEET_NAME);
  if (testSheet) {
    ss.deleteSheet(testSheet);
  }

  testSheet = ss.insertSheet(TEST_SHEET_NAME);

  try {
    // Same A:I shape as SENSORY, but test data only.
    testSheet.getRange(1, 1, 1, 9).setValues([[
      "TIMESTAMP",
      "NAME",
      "MYKAD/PASSPORT",
      "REG.NUM",
      "CONTACT",
      "CATEGORY",
      "TOWER",
      "REMARK",
      "PHOTO LINK"
    ]]);

    const now = new Date();

    function daysAgo_(days) {
      return formatDateTimeDMY(new Date(now.getTime() - (days * 86400000)));
    }

    // Intentionally OUT OF ORDER:
    // 10 days  = keep
    // 120 days = expired (middle)
    // 20 days  = keep
    // 100 days = expired (middle)
    // 5 days   = keep
    const rows = [
      [daysAgo_(10),  "RET04 KEEP A", "", "RET04A", "", "OTHER", "TOWER A", "", ""],
      [daysAgo_(120), "RET04 EXPIRE A", "", "RET04X1", "", "OTHER", "TOWER A", "", ""],
      [daysAgo_(20),  "RET04 KEEP B", "", "RET04B", "", "OTHER", "TOWER A", "", ""],
      [daysAgo_(100), "RET04 EXPIRE B", "", "RET04X2", "", "OTHER", "TOWER A", "", ""],
      [daysAgo_(5),   "RET04 KEEP C", "", "RET04C", "", "OTHER", "TOWER A", "", ""]
    ];

    testSheet.getRange(2, 1, rows.length, 9).setValues(rows);

    // Force full order-independent sweep.
    const cleanup = cleanupByAge_(testSheet, 90, true);

    const remainingRows = testSheet.getLastRow() > 1
      ? testSheet.getRange(2, 1, testSheet.getLastRow() - 1, 9).getDisplayValues()
      : [];

    const remainingNames = remainingRows.map(r => String(r[1] || "").trim());
    const expectedNames = ["RET04 KEEP A", "RET04 KEEP B", "RET04 KEEP C"];

    const productionRowsAfter = mainSheet.getLastRow();

    const pass =
      cleanup &&
      cleanup.rowsDeleted === 2 &&
      cleanup.driveDeleteAttempted === 0 &&
      cleanup.driveDeleteConfirmed === 0 &&
      cleanup.driveDeleteFailed === 0 &&
      remainingNames.length === 3 &&
      expectedNames.every(name => remainingNames.indexOf(name) !== -1) &&
      remainingNames.indexOf("RET04 EXPIRE A") === -1 &&
      remainingNames.indexOf("RET04 EXPIRE B") === -1 &&
      productionRowsBefore === productionRowsAfter;

    const result = {
      PASSED: pass,
      rowsDeleted: cleanup ? cleanup.rowsDeleted : null,
      remainingNames: remainingNames,
      driveDeleteAttempted: cleanup ? cleanup.driveDeleteAttempted : null,
      productionRowsBefore: productionRowsBefore,
      productionRowsAfter: productionRowsAfter,
      productionSheetTouched: productionRowsBefore !== productionRowsAfter
    };

    console.log("NEW-RET-04 CONTROLLED TEST RESULT:");
    console.log(JSON.stringify(result, null, 2));

    if (!pass) {
      throw new Error("NEW-RET-04 controlled test FAILED: " + JSON.stringify(result));
    }

    return result;

  } finally {
    // Restore production retention throttle state exactly as it was.
    if (oldFullScanProp === null) {
      props.deleteProperty(RETENTION_LAST_FULL_SCAN_PROP);
    } else {
      props.setProperty(RETENTION_LAST_FULL_SCAN_PROP, oldFullScanProp);
    }

    const cleanupSheet = ss.getSheetByName(TEST_SHEET_NAME);
    if (cleanupSheet) {
      ss.deleteSheet(cleanupSheet);
    }
  }
}
