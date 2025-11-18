// === CONFIGURATION ===
const TMS_ID = "YOUR_TMS_SOURCE_ID_PLACEHOLDER"; 
const TW_For_View_ID = "YOUR_TW_FOR_VIEW_ID_PLACEHOLDER";
const SHEET_NAME = "TW RSV Plans";

const SOURCE_COLS = 35;      // A:AI
const FORMULA_START_COL = 36; // AJ
const FORMULA_END_COL = 46;   // AT
const FORMULA_COLS = FORMULA_END_COL - FORMULA_START_COL + 1;
const DATA_START_ROW = 3;    // Start pasting data from Row 3
const CHUNK_SIZE = 3000;     // Number of rows to paste per chunk


// =============== SECTION 1: Import Data + Restore Formulas ===============
function import_RSV_Plans_Final() {
  const startTime = new Date();
  const TMS = SpreadsheetApp.openById(TMS_ID);
  const TW_For_View = SpreadsheetApp.openById(TW_For_View_ID);
  const RAW = TMS.getSheetByName(SHEET_NAME);
  const NEW = TW_For_View.getSheetByName(SHEET_NAME);

  try {
    const lastRow = RAW.getLastRow();

    // Step 1. Clear old data
    NEW.getRange(DATA_START_ROW, 1, NEW.getMaxRows() - DATA_START_ROW + 1, FORMULA_END_COL).clearContent();
    Logger.log(`Cleared A:AT (from row ${DATA_START_ROW} onwards)`);

    // Step 2. Chunked data copy (A:AI)
    let startRow = 2;
    let totalRows = 0;
    while (startRow <= lastRow) {
      const numRows = Math.min(CHUNK_SIZE, lastRow - startRow + 1);
      const data = RAW.getRange(startRow, 1, numRows, SOURCE_COLS).getValues();
      NEW.getRange(startRow + 1, 1, numRows, SOURCE_COLS).setValues(data); // +1 -> Start pasting from Row 3
      startRow += numRows;
      totalRows += numRows;
    }

    // Step 3. Restore formulas
    restoreFormulasFromRow2(NEW);

    const duration = ((new Date()) - startTime) / 1000;
    Logger.log(`✅ Import completed, total ${totalRows} rows, took ${duration.toFixed(1)} seconds.`);

    // Step 4. Create subsequent trigger
    const trigger = ScriptApp.newTrigger("standardizeData")
      .timeBased()
      .after(10 * 1000) // Execute after 10 seconds
      .create();
    Logger.log(`⏱ Trigger for standardizeData created (ID: ${trigger.getUniqueId()}), executing in 10 seconds.`);

  } catch (err) {
    Logger.log(`❌ import_RSV_Plans_Final failed: ${err.message}`);
  }
}


// =============== SECTION 2: Data Standardization ===============
function standardizeData() {
  // Prevent duplicate triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "standardizeData") {
      ScriptApp.deleteTrigger(t);
    }
  });

  try {
    const ss = SpreadsheetApp.openById(TW_For_View_ID);
    const sheet = ss.getSheetByName("TW RSV Plans");
    const lastRow = sheet.getLastRow();
    const START_ROW = 3;
    const K_COL = 11;  // Column K

    Logger.log(`Starting check for Column K date format, range: Row ${START_ROW} to ${lastRow}`);

    // Part 1: Column K Date Formatting
    const kRange = sheet.getRange(START_ROW, K_COL, lastRow - START_ROW + 1, 1);
    const kValues = kRange.getValues();
    let formattedCount = 0;
    for (let i = 0; i < kValues.length; i++) {
      const value = kValues[i][0];
      const row = START_ROW + i;
      if (value instanceof Date) {
        sheet.getRange(row, K_COL).setNumberFormat("m/d");
        formattedCount++;
      }
    }
    Logger.log(`✅ Column K completed: ${formattedCount} dates set to m/d format.`);

    // Part 2: Paste AJ-AT as values
    const VALUE_START_ROW = 3;
    const AJ_COL = 36;
    const AT_COL = 46;
    const numCols = AT_COL - AJ_COL + 1;

    const valueRange = sheet.getRange(VALUE_START_ROW, AJ_COL, lastRow - VALUE_START_ROW + 1, numCols);
    const values = valueRange.getValues();
    valueRange.setValues(values);
    Logger.log(`✅ Columns AJ-AT pasted as values.`);

    Logger.log("standardizeData processing fully completed!");

  } catch (e) {
    Logger.log(`❌ standardizeData failed: ${e.message}`);
  }
}


// =============== UTILITY FUNCTION: Restore Formulas ===============
function restoreFormulasFromRow2(sheet) {
  try {
    const formulaTexts = sheet.getRange(2, FORMULA_START_COL, 1, FORMULA_COLS).getValues();
    const formulas = formulaTexts.map(row =>
      row.map(cell => {
        if (!cell) return "";
        const text = cell.toString().trim();
        return text.startsWith("'=") ? text.slice(1) : text;
      })
    );
    sheet.getRange(3, FORMULA_START_COL, 1, FORMULA_COLS).setFormulas(formulas);
    Logger.log(`🪄 Restored ${FORMULA_COLS} formulas from Row 2 to Row 3 (AJ:AT).`);
  } catch (e) {
    Logger.log(`⚠️ restoreFormulasFromRow2 failed: ${e.message}`);
  }
}