// === 設定區 ===
const TMS_ID = "1jhP5lZeiNGJq7tumkcKZQDRyjK4GQTwueD2rLcAlvgg";
const TW_For_View_ID = "1sDlWkzxDsg69lYHYHiDEMLN7xarvtzrqsFLRiFqhbgk";
const SHEET_NAME = "TW RSV Plans";

const SOURCE_COLS = 35;      // A:AI
const FORMULA_START_COL = 36; // AJ
const FORMULA_END_COL = 46;   // AT
const FORMULA_COLS = FORMULA_END_COL - FORMULA_START_COL + 1;
const DATA_START_ROW = 3;    // 從第 3 列開始貼資料
const CHUNK_SIZE = 3000;     // 分塊貼值筆數


// =============== 第一段：匯入資料 + 還原公式 ===============
function import_RSV_Plans_Final() {
  const startTime = new Date();
  const TMS = SpreadsheetApp.openById(TMS_ID);
  const TW_For_View = SpreadsheetApp.openById(TW_For_View_ID);
  const RAW = TMS.getSheetByName(SHEET_NAME);
  const NEW = TW_For_View.getSheetByName(SHEET_NAME);

  try {
    const lastRow = RAW.getLastRow();

    // Step 1. 清空舊資料
    NEW.getRange(DATA_START_ROW, 1, NEW.getMaxRows() - DATA_START_ROW + 1, FORMULA_END_COL).clearContent();
    Logger.log(`已清空 A:AT (第 ${DATA_START_ROW} 列以後)`);

    // Step 2. 分塊複製資料 (A:AI)
    let startRow = 2;
    let totalRows = 0;
    while (startRow <= lastRow) {
      const numRows = Math.min(CHUNK_SIZE, lastRow - startRow + 1);
      const data = RAW.getRange(startRow, 1, numRows, SOURCE_COLS).getValues();
      NEW.getRange(startRow + 1, 1, numRows, SOURCE_COLS).setValues(data); // +1 → 從第 3 列開始貼
      startRow += numRows;
      totalRows += numRows;
    }

    // Step 3. 還原公式
    restoreFormulasFromRow2(NEW);

    const duration = ((new Date()) - startTime) / 1000;
    Logger.log(`✅ 匯入完成，共 ${totalRows} 筆資料，耗時 ${duration.toFixed(1)} 秒。`);

    // Step 4. 建立接續 trigger
    const trigger = ScriptApp.newTrigger("standardizeData")
      .timeBased()
      .after(10 * 1000) // 10 秒後執行
      .create();
    Logger.log(`⏱ 已建立 standardizeData 的觸發器（ID: ${trigger.getUniqueId()}），10 秒後執行。`);

  } catch (err) {
    Logger.log(`❌ import_RSV_Plans_Final 執行失敗：${err.message}`);
  }
}


// =============== 第二段：標準化格式 ===============
function standardizeData() {
  // 防重複觸發器
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
    const K_COL = 11;  // K 欄

    Logger.log(`開始檢查 K 欄的日期格式，範圍：第 ${START_ROW} 到 ${lastRow} 列`);

    // Part 1: K 欄日期格式
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
    Logger.log(`✅ K 欄完成：共 ${formattedCount} 個日期設定為 m/d 格式。`);

    // Part 2: AJ～AT 欄貼成值
    const VALUE_START_ROW = 3;
    const AJ_COL = 36;
    const AT_COL = 46;
    const numCols = AT_COL - AJ_COL + 1;

    const valueRange = sheet.getRange(VALUE_START_ROW, AJ_COL, lastRow - VALUE_START_ROW + 1, numCols);
    const values = valueRange.getValues();
    valueRange.setValues(values);
    Logger.log(`✅ AJ～AT 欄資料已貼成值。`);

    Logger.log("standardizeData 全部處理完成！");

  } catch (e) {
    Logger.log(`❌ standardizeData 執行失敗：${e.message}`);
  }
}


// =============== 工具函式：還原公式 ===============
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
    Logger.log(`🪄 已從第 2 列貼回 ${FORMULA_COLS} 欄公式到第 3 列 (AJ:AT)。`);
  } catch (e) {
    Logger.log(`⚠️ restoreFormulasFromRow2 失敗：${e.message}`);
  }
}
