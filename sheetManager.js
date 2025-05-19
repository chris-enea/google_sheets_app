/**
 * @OnlyCurrentDoc
 */

// --- SHARED CONSTANTS ---
const MASTER_SHEET_NAME = "Master Item List";
const SPEC_FFE_COLUMN_HEADER = "SPEC/FFE";

// Master list column names used by one or both processes
const MASTER_ROOM_COL_NAME = "ROOM";
const MASTER_TYPE_COL_NAME = "TYPE";
const MASTER_ITEM_COL_NAME = "ITEM"; // Source for Allowances ITEM column
const MASTER_QTY_COL_NAME = "QUANTITY";
const MASTER_LOW_UNIT_COST_COL_NAME = "LOW";
const MASTER_LOW_TOTAL_HEADER = "LOW TOTAL";
const MASTER_HIGH_UNIT_COST_COL_NAME = "HIGH";
const MASTER_HIGH_TOTAL_HEADER = "HIGH TOTAL";

// Add other master column names here if they become shared or are useful as constants

// Default column width for new/blank columns in target sheets
const DEFAULT_TARGET_COL_WIDTH = 100;

// --- Budget Sheet Constants ---
const BUDGET_SHEET_NAME = "Budget";
const BUDGET_TYPE_COL_HEADER = "TYPE";
const BUDGET_TOTAL_LOW_COL_HEADER = "TOTAL LOW";
const BUDGET_TOTAL_HIGH_COL_HEADER = "TOTAL HIGH";
const BUDGET_HEADERS = [
  "CATEGORIES", "TYPE", "SET ALLOWANCE", "LOW",
  "TOTAL LOW", "HIGH", "TOTAL HIGH", "NOTES"
];
const DEFAULT_BUDGET_COL_WIDTH = 100;

/**
 * Main orchestrator function called by the UI menu.
 * It processes items from "Master Item List" and splits them into "FFE" and "Allowances" sheets.
 */
function splitItemsByFFE() { // Renaming this might be good later, but UI calls this.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);

  if (!masterSheet) {
    SpreadsheetApp.getUi().alert(`Sheet "${MASTER_SHEET_NAME}" not found.`);
    return;
  }

  const masterDataRange = masterSheet.getDataRange();
  const allMasterValues = masterDataRange.getValues();
  const allMasterFormulas = masterDataRange.getFormulas();

  if (allMasterValues.length === 0) {
    SpreadsheetApp.getUi().alert(`Sheet "${MASTER_SHEET_NAME}" is empty.`);
    return;
  }

  const masterHeaders = allMasterValues[0];
  const specFfeColumnIndex = masterHeaders.indexOf(SPEC_FFE_COLUMN_HEADER);

  if (specFfeColumnIndex === -1) {
    SpreadsheetApp.getUi().alert(`Column "${SPEC_FFE_COLUMN_HEADER}" not found in sheet "${MASTER_SHEET_NAME}".`);
    return;
  }

  // Pre-fetch all master content formatting arrays once
  let masterContentDetails = {};
  if (allMasterValues.length > 1) { // If there are data rows beyond the header
    const masterContentDataRange = masterSheet.getRange(2, 1, allMasterValues.length - 1, masterHeaders.length);
    masterContentDetails = {
      textStyles: masterContentDataRange.getTextStyles(),
      backgrounds: masterContentDataRange.getBackgrounds(),
      fontColors: masterContentDataRange.getFontColors(),
      fontWeights: masterContentDataRange.getFontWeights(),
      fontStyles: masterContentDataRange.getFontStyles(),
      fontLines: masterContentDataRange.getFontLines(),
      horizontalAlignments: masterContentDataRange.getHorizontalAlignments(),
      verticalAlignments: masterContentDataRange.getVerticalAlignments(),
      wrapStrategies: masterContentDataRange.getWrapStrategies(),
      numberFormats: masterContentDataRange.getNumberFormats()
    };
  }

  // --- FFE Processing Configuration ---
  const ffeSheetName = "FFE";
  const ffeTargetHeaders = masterHeaders.filter((header, index) => index !== specFfeColumnIndex);
  
  const ffeColumnMapping = masterHeaders.map((header, masterColIdx) => {
    if (masterColIdx === specFfeColumnIndex) return null; // Skip SPEC/FFE column
    return {
      targetHeaderName: header,
      sourceMasterColumnName: header,
      isFormulaColumn: (header === MASTER_LOW_TOTAL_HEADER || header === MASTER_HIGH_TOTAL_HEADER),
      masterColumnIndex: masterColIdx // Store for direct access
    };
  }).filter(Boolean); // Remove null entry for SPEC/FFE

  _processAndCopyItemsInternal(
    ss, masterSheet, allMasterValues, allMasterFormulas, masterHeaders, specFfeColumnIndex, masterContentDetails,
    "FFE", ffeSheetName, ffeTargetHeaders, ffeColumnMapping,
    true // deleteSpecFfeColumnInTarget (true for FFE as header is copied then col deleted)
  );

  // --- SPEC (Allowances) Processing Configuration ---
  const allowancesSheetName = "Allowances";
  const allowancesTargetHeaders = [
    "CATEGORIES", "TYPE", "ITEM", "SET ALLOWANCE",
    "QUANTITY", "LOW", "TOTAL LOW", "HIGH", "TOTAL HIGH", "NOTES"
  ];
  
  const allowancesColumnMapping = [
    { targetHeaderName: "CATEGORIES", isBlank: true, defaultFormatSourceMasterColName: masterHeaders[0] || "" }, // Use first master col for style
    { targetHeaderName: "TYPE", sourceMasterColumnName: MASTER_TYPE_COL_NAME },
    { targetHeaderName: "ITEM", sourceMasterColumnName: MASTER_ITEM_COL_NAME },
    { targetHeaderName: "SET ALLOWANCE", isBlank: true, defaultFormatSourceMasterColName: masterHeaders[0] || "" }, // Use first master col for style
    { targetHeaderName: "QUANTITY", sourceMasterColumnName: MASTER_QTY_COL_NAME },
    { targetHeaderName: "LOW", sourceMasterColumnName: MASTER_LOW_UNIT_COST_COL_NAME },
    { targetHeaderName: "TOTAL LOW", sourceMasterColumnName: MASTER_LOW_TOTAL_HEADER, isFormulaColumn: true },
    { targetHeaderName: "HIGH", sourceMasterColumnName: MASTER_HIGH_UNIT_COST_COL_NAME },
    { targetHeaderName: "TOTAL HIGH", sourceMasterColumnName: MASTER_HIGH_TOTAL_HEADER, isFormulaColumn: true },
    { targetHeaderName: "NOTES", isBlank: true }
  ];
  
  // Add masterColIndex to allowancesColumnMapping and validate required columns
  const tempMasterHeaderIndices = {};
  masterHeaders.forEach((h, i) => tempMasterHeaderIndices[h] = i);

  for (const mapping of allowancesColumnMapping) {
    if (mapping.sourceMasterColumnName) {
      mapping.masterColumnIndex = tempMasterHeaderIndices[mapping.sourceMasterColumnName];
      if (mapping.masterColumnIndex === undefined) {
        SpreadsheetApp.getUi().alert(`Configuration error for Allowances: Master column "${mapping.sourceMasterColumnName}" not found.`);
        return; // Stop if config is bad
      }
    }
    if (mapping.isBlank && mapping.defaultFormatSourceMasterColName) {
        mapping.defaultFormatMasterColIndex = tempMasterHeaderIndices[mapping.defaultFormatSourceMasterColName];
         if (mapping.defaultFormatMasterColIndex === undefined && masterHeaders.length > 0) {
           // Fallback to actual first column if named one isn't found
            mapping.defaultFormatMasterColIndex = 0;
         } else if (masterHeaders.length === 0) {
            mapping.defaultFormatMasterColIndex = -1; // No columns to get style from
         }
    }
  }

  _processAndCopyItemsInternal(
    ss, masterSheet, allMasterValues, allMasterFormulas, masterHeaders, specFfeColumnIndex, masterContentDetails,
    "SPEC", allowancesSheetName, allowancesTargetHeaders, allowancesColumnMapping,
    false // deleteSpecFfeColumnInTarget (false for Allowances, it has its own headers)
  );
}


/**
 * Internal helper function to process and copy items based on provided configuration.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} masterSheet The master data sheet.
 * @param {Array<Array<String|Number>>} allMasterValues All values from masterSheet.
 * @param {Array<Array<String>>} allMasterFormulas All formulas from masterSheet.
 * @param {Array<String>} masterHeaders The header row of masterSheet.
 * @param {Number} specFfeColumnIndex The 0-based index of the 'SPEC/FFE' column in masterSheet.
 * @param {Object} masterContentDetails Object containing pre-fetched formatting arrays from masterSheet.
 * @param {String} filterValue The value to filter by in the 'SPEC/FFE' column (e.g., "FFE", "SPEC").
 * @param {String} targetSheetName The name of the sheet to create/clear and copy data to.
 * @param {Array<String>} targetHeadersArray The header row for the target sheet.
 * @param {Array<Object>} columnMappingConfig Configuration for mapping master columns to target columns.
 * @param {Boolean} copyMasterHeaderAndDelCol For FFE: true to copy master header then delete SPEC/FFE. For Allowances: false.
 */
function _processAndCopyItemsInternal(
    ss, masterSheet, allMasterValues, allMasterFormulas, masterHeaders, specFfeColumnIndex, masterContentDetails,
    filterValue, targetSheetName, targetHeadersArray, columnMappingConfig, copyMasterHeaderAndDelCol) {

  let targetSheet = ss.getSheetByName(targetSheetName);
  if (targetSheet) {
    targetSheet.clear();
  } else {
    targetSheet = ss.insertSheet(targetSheetName);
  }

  // Set/Copy Headers
  if (copyMasterHeaderAndDelCol) { // FFE case
    if (masterSheet.getLastColumn() > 0 && masterSheet.getLastRow() > 0) {
      const masterHeaderRange = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn());
      masterHeaderRange.copyTo(targetSheet.getRange(1, 1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      targetSheet.deleteColumn(specFfeColumnIndex + 1);
    } else {
      SpreadsheetApp.getUi().alert(`Master sheet has no columns or rows to copy for ${targetSheetName} header.`);
      return;
    }
  } else { // Allowances case (and any future similar cases)
    const targetHeaderRange = targetSheet.getRange(1, 1, 1, targetHeadersArray.length);
    targetHeaderRange.setValues([targetHeadersArray]);
    targetHeaderRange.setFontWeight("bold");
  }

  const processedRows = [];
  const rowFormatDetails = {
    textStyles: [], backgrounds: [], fontColors: [], fontWeights: [],
    fontStyles: [], fontLines: [], horizontalAlignments: [],
    verticalAlignments: [], wrapStrategies: [], numberFormats: []
  };
  const originalMasterRowNumbers = [];

  const hasMasterContent = masterContentDetails.textStyles && masterContentDetails.textStyles.length > 0;

  for (let i = 1; i < allMasterValues.length; i++) { // Iterate master data rows (skip header)
    const masterRowValues = allMasterValues[i];
    const masterRowFormulas = allMasterFormulas[i];
    const formatSourceRowIndex = i - 1; // 0-indexed for masterContentDetails arrays

    if (masterRowValues[specFfeColumnIndex] && masterRowValues[specFfeColumnIndex].toString().trim().toUpperCase() === filterValue) {
      const singleTargetRow = [];
      const singleRowFormat_TextStyles = [];
      const singleRowFormat_Backgrounds = [];
      const singleRowFormat_FontColors = [];
      const singleRowFormat_FontWeights = [];
      const singleRowFormat_FontStyles = [];
      const singleRowFormat_FontLines = [];
      const singleRowFormat_HorizontalAlignments = [];
      const singleRowFormat_VerticalAlignments = [];
      const singleRowFormat_WrapStrategies = [];
      const singleRowFormat_NumberFormats = [];

      for (const mapping of columnMappingConfig) {
        let cellValue = "";
        let srcMasterColIdx = -1;

        if (mapping.isBlank) {
          cellValue = "";
          srcMasterColIdx = mapping.defaultFormatMasterColIndex !== undefined ? mapping.defaultFormatMasterColIndex : (masterHeaders.length > 0 ? 0 : -1);
        } else {
          srcMasterColIdx = mapping.masterColumnIndex;
          if (srcMasterColIdx === undefined || srcMasterColIdx < 0 || srcMasterColIdx >= masterHeaders.length) {
             SpreadsheetApp.getUi().alert(`Error processing ${targetSheetName}: Misconfiguration for column '${mapping.targetHeaderName}', invalid master column index.`);
             cellValue = "ERROR"; // Or skip row / handle error
          } else {
            if (mapping.isFormulaColumn && masterRowFormulas[srcMasterColIdx]) {
              cellValue = masterRowFormulas[srcMasterColIdx];
            } else {
              cellValue = masterRowValues[srcMasterColIdx];
            }
          }
        }
        singleTargetRow.push(cellValue);

        // Collect formatting for this cell
        if (hasMasterContent && srcMasterColIdx !== -1) {
            singleRowFormat_TextStyles.push(masterContentDetails.textStyles[formatSourceRowIndex][srcMasterColIdx]);
            singleRowFormat_Backgrounds.push(masterContentDetails.backgrounds[formatSourceRowIndex][srcMasterColIdx]);
            singleRowFormat_FontColors.push(masterContentDetails.fontColors[formatSourceRowIndex][srcMasterColIdx]);
            singleRowFormat_FontWeights.push(masterContentDetails.fontWeights[formatSourceRowIndex][srcMasterColIdx]);
            singleRowFormat_FontStyles.push(masterContentDetails.fontStyles[formatSourceRowIndex][srcMasterColIdx]);
            singleRowFormat_FontLines.push(masterContentDetails.fontLines[formatSourceRowIndex][srcMasterColIdx]);
            singleRowFormat_HorizontalAlignments.push(masterContentDetails.horizontalAlignments[formatSourceRowIndex][srcMasterColIdx]);
            singleRowFormat_VerticalAlignments.push(masterContentDetails.verticalAlignments[formatSourceRowIndex][srcMasterColIdx]);
            singleRowFormat_WrapStrategies.push(masterContentDetails.wrapStrategies[formatSourceRowIndex][srcMasterColIdx]);
            singleRowFormat_NumberFormats.push(masterContentDetails.numberFormats[formatSourceRowIndex][srcMasterColIdx]);
        } else { // Push nulls or defaults if no master content or invalid srcMasterColIdx for formatting
            singleRowFormat_TextStyles.push(null); singleRowFormat_Backgrounds.push(null);
            singleRowFormat_FontColors.push(null); singleRowFormat_FontWeights.push(null);
            singleRowFormat_FontStyles.push(null); singleRowFormat_FontLines.push(null);
            singleRowFormat_HorizontalAlignments.push(null); singleRowFormat_VerticalAlignments.push(null);
            singleRowFormat_WrapStrategies.push(null); singleRowFormat_NumberFormats.push(null);
        }
      }
      processedRows.push(singleTargetRow);
      originalMasterRowNumbers.push(i + 1); // 1-indexed master row number

      if (hasMasterContent) {
        rowFormatDetails.textStyles.push(singleRowFormat_TextStyles);
        rowFormatDetails.backgrounds.push(singleRowFormat_Backgrounds);
        rowFormatDetails.fontColors.push(singleRowFormat_FontColors);
        rowFormatDetails.fontWeights.push(singleRowFormat_FontWeights);
        rowFormatDetails.fontStyles.push(singleRowFormat_FontStyles);
        rowFormatDetails.fontLines.push(singleRowFormat_FontLines);
        rowFormatDetails.horizontalAlignments.push(singleRowFormat_HorizontalAlignments);
        rowFormatDetails.verticalAlignments.push(singleRowFormat_VerticalAlignments);
        rowFormatDetails.wrapStrategies.push(singleRowFormat_WrapStrategies);
        rowFormatDetails.numberFormats.push(singleRowFormat_NumberFormats);
      }
    }
  }

  if (processedRows.length === 0) {
    SpreadsheetApp.getUi().alert(`No data rows with "${filterValue}" in column "${SPEC_FFE_COLUMN_HEADER}" found for sheet "${targetSheetName}".`);
    // If FFE processing yields no rows, we might not want to proceed to SPEC, or handle it gracefully.
    // For now, it will just alert and the orchestrator will call for SPEC anyway.
    return; 
  }

  const numDataRows = processedRows.length;
  const numDataCols = processedRows[0].length;
  const targetDataRangeToWrite = targetSheet.getRange(2, 1, numDataRows, numDataCols);
  
  targetDataRangeToWrite.setValues(processedRows);

  if (hasMasterContent && rowFormatDetails.textStyles.length > 0) {
    targetDataRangeToWrite.setTextStyles(rowFormatDetails.textStyles);
    targetDataRangeToWrite.setBackgrounds(rowFormatDetails.backgrounds);
    targetDataRangeToWrite.setFontColors(rowFormatDetails.fontColors);
    targetDataRangeToWrite.setFontWeights(rowFormatDetails.fontWeights);
    targetDataRangeToWrite.setFontStyles(rowFormatDetails.fontStyles);
    targetDataRangeToWrite.setFontLines(rowFormatDetails.fontLines);
    targetDataRangeToWrite.setHorizontalAlignments(rowFormatDetails.horizontalAlignments);
    targetDataRangeToWrite.setVerticalAlignments(rowFormatDetails.verticalAlignments);
    targetDataRangeToWrite.setWrapStrategies(rowFormatDetails.wrapStrategies);
    targetDataRangeToWrite.setNumberFormats(rowFormatDetails.numberFormats);
  }

  // --- Column Widths ---
  if (copyMasterHeaderAndDelCol) { // FFE case
    let ffeSheetColIdx = 1;
    for (let masterSheetColIdx = 0; masterSheetColIdx < masterHeaders.length; masterSheetColIdx++) {
      if (masterSheetColIdx !== specFfeColumnIndex) {
        targetSheet.setColumnWidth(ffeSheetColIdx, masterSheet.getColumnWidth(masterSheetColIdx + 1));
        ffeSheetColIdx++;
      }
    }
  } else { // Allowances (and similar non-header-copying cases)
      for (let col = 0; col < targetHeadersArray.length; col++) {
          const mapping = columnMappingConfig[col];
          let widthToSet = DEFAULT_TARGET_COL_WIDTH;
          if (mapping && mapping.sourceMasterColumnName && mapping.masterColumnIndex !== undefined && mapping.masterColumnIndex >= 0) {
              try { // It's possible masterColumnIndex is for a column that doesn't exist if config is bad.
                widthToSet = masterSheet.getColumnWidth(mapping.masterColumnIndex + 1);
              } catch (e) {
                // keep default width
              }
          } else if (mapping && mapping.isBlank) {
              widthToSet = DEFAULT_TARGET_COL_WIDTH; // Default for blank columns
          }
          targetSheet.setColumnWidth(col + 1, widthToSet);
      }
  }
  
  // --- Row Heights ---
  for (let i = 0; i < numDataRows; i++) {
    targetSheet.setRowHeight(i + 2, masterSheet.getRowHeight(originalMasterRowNumbers[i]));
  }

  SpreadsheetApp.getUi().alert(`"${filterValue}" items processed and copied to sheet "${targetSheetName}" successfully.`);
} 

/**
 * Summarizes SPEC items from Master Item List and updates the Budget sheet.
 * Totals "TOTAL LOW" and "TOTAL HIGH" for each unique ITEM marked as SPEC.
 * Updates existing items in the Budget sheet or adds new ones.
 * Preserves CATEGORIES, SET ALLOWANCE, and NOTES columns in the Budget sheet.
 */
function updateBudgetFromSpecItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // --- 1. Get Master Sheet Data ---
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    ui.alert(`Sheet "${MASTER_SHEET_NAME}" not found.`);
    return;
  }

  const masterDataRange = masterSheet.getDataRange();
  const allMasterValues = masterDataRange.getValues(); // Includes headers

  if (allMasterValues.length <= 1) { // Only headers or empty
    ui.alert(`Sheet "${MASTER_SHEET_NAME}" has no data to process.`);
    return;
  }

  const masterHeaders = allMasterValues[0];
  const specFfeColIdx = masterHeaders.indexOf(SPEC_FFE_COLUMN_HEADER);
  const masterTypeColIdx = masterHeaders.indexOf(MASTER_TYPE_COL_NAME);
  const masterLowTotalColIdx = masterHeaders.indexOf(MASTER_LOW_TOTAL_HEADER);
  const masterHighTotalColIdx = masterHeaders.indexOf(MASTER_HIGH_TOTAL_HEADER);

  if (specFfeColIdx === -1) {
    ui.alert(`Column "${SPEC_FFE_COLUMN_HEADER}" not found in "${MASTER_SHEET_NAME}".`);
    return;
  }
  if (masterTypeColIdx === -1) {
    ui.alert(`Column "${MASTER_TYPE_COL_NAME}" not found in "${MASTER_SHEET_NAME}".`);
    return;
  }
  if (masterLowTotalColIdx === -1) {
    ui.alert(`Column "${MASTER_LOW_TOTAL_HEADER}" not found in "${MASTER_SHEET_NAME}".`);
    return;
  }
  if (masterHighTotalColIdx === -1) {
    ui.alert(`Column "${MASTER_HIGH_TOTAL_HEADER}" not found in "${MASTER_SHEET_NAME}".`);
    return;
  }

  // --- 2. Aggregate SPEC Data from Master Sheet ---
  const specTotals = {}; // E.g., {"Type A": { low: 100, high: 200 }}

  for (let i = 1; i < allMasterValues.length; i++) {
    const masterRow = allMasterValues[i];
    if (masterRow[specFfeColIdx] && masterRow[specFfeColIdx].toString().trim().toUpperCase() === "SPEC") {
      const typeName = masterRow[masterTypeColIdx] ? masterRow[masterTypeColIdx].toString().trim() : null;
      if (!typeName) continue; // Skip if type name is blank

      const lowTotal = parseFloat(masterRow[masterLowTotalColIdx]) || 0;
      const highTotal = parseFloat(masterRow[masterHighTotalColIdx]) || 0;

      if (!specTotals[typeName]) {
        specTotals[typeName] = { low: 0, high: 0 };
      }
      specTotals[typeName].low += lowTotal;
      specTotals[typeName].high += highTotal;
    }
  }

  if (Object.keys(specTotals).length === 0) {
    ui.alert(`No "SPEC" items found in "${MASTER_SHEET_NAME}" to update the Budget with.`);
    return;
  }

  // --- 3. Prepare and Update Budget Sheet ---
  let budgetSheet = ss.getSheetByName(BUDGET_SHEET_NAME);

  if (!budgetSheet) {
    budgetSheet = ss.insertSheet(BUDGET_SHEET_NAME);
    const headerRange = budgetSheet.getRange(1, 1, 1, BUDGET_HEADERS.length);
    headerRange.setValues([BUDGET_HEADERS]).setFontWeight("bold");
    for (let i = 0; i < BUDGET_HEADERS.length; i++) {
        budgetSheet.setColumnWidth(i + 1, DEFAULT_BUDGET_COL_WIDTH);
    }
    // Adjust specific column widths if needed, e.g. for TYPE or NOTES
    const typeColBudgetIdx = BUDGET_HEADERS.indexOf(BUDGET_TYPE_COL_HEADER);
    if (typeColBudgetIdx !== -1) budgetSheet.setColumnWidth(typeColBudgetIdx + 1, 200); // Example width
    const notesColBudgetIdx = BUDGET_HEADERS.indexOf("NOTES");
    if (notesColBudgetIdx !== -1) budgetSheet.setColumnWidth(notesColBudgetIdx + 1, 300); // Example width
  }

  const budgetDataRange = budgetSheet.getDataRange();
  const budgetValues = budgetDataRange.getValues();
  let budgetSheetHeaders;

  if (budgetValues.length === 0) { // Sheet exists but is completely empty
    budgetSheet.getRange(1, 1, 1, BUDGET_HEADERS.length).setValues([BUDGET_HEADERS]).setFontWeight("bold");
     for (let i = 0; i < BUDGET_HEADERS.length; i++) {
        budgetSheet.setColumnWidth(i + 1, DEFAULT_BUDGET_COL_WIDTH);
    }
    budgetSheetHeaders = BUDGET_HEADERS;
  } else {
    budgetSheetHeaders = budgetValues[0];
  }
  
  const budgetTypeColIdx = budgetSheetHeaders.indexOf(BUDGET_TYPE_COL_HEADER);
  const budgetTotalLowColIdx = budgetSheetHeaders.indexOf(BUDGET_TOTAL_LOW_COL_HEADER);
  const budgetTotalHighColIdx = budgetSheetHeaders.indexOf(BUDGET_TOTAL_HIGH_COL_HEADER);

  if (budgetTypeColIdx === -1) {
    ui.alert(`Column "${BUDGET_TYPE_COL_HEADER}" not found in "${BUDGET_SHEET_NAME}".`);
    return;
  }
  if (budgetTotalLowColIdx === -1) {
    ui.alert(`Column "${BUDGET_TOTAL_LOW_COL_HEADER}" not found in "${BUDGET_SHEET_NAME}".`);
    return;
  }
  if (budgetTotalHighColIdx === -1) {
    ui.alert(`Column "${BUDGET_TOTAL_HIGH_COL_HEADER}" not found in "${BUDGET_SHEET_NAME}".`);
    return;
  }

  const typeRowMap = {}; // Map type name to 1-based row index in Budget sheet. Renamed from itemRowMap
  for (let r = 1; r < budgetValues.length; r++) { // Start from 1 to skip header
    const typeNameInBudget = budgetValues[r][budgetTypeColIdx];
    if (typeNameInBudget) {
      typeRowMap[typeNameInBudget.toString().trim()] = r + 1; // r+1 for 1-based index
    }
  }

  const rowsToAdd = [];
  let updatedCount = 0;
  let addedCount = 0;

  for (const typeName in specTotals) {
    const data = specTotals[typeName];
    if (typeRowMap[typeName]) {
      const rowToUpdate = typeRowMap[typeName];
      budgetSheet.getRange(rowToUpdate, budgetTotalLowColIdx + 1).setValue(data.low.toFixed(2));
      budgetSheet.getRange(rowToUpdate, budgetTotalHighColIdx + 1).setValue(data.high.toFixed(2));
      updatedCount++;
    } else {
      const newRow = new Array(budgetSheetHeaders.length).fill("");
      newRow[budgetTypeColIdx] = typeName;
      newRow[budgetTotalLowColIdx] = data.low.toFixed(2);
      newRow[budgetTotalHighColIdx] = data.high.toFixed(2);
      rowsToAdd.push(newRow);
      addedCount++;
    }
  }

  if (rowsToAdd.length > 0) {
    budgetSheet.getRange(budgetSheet.getLastRow() + 1, 1, rowsToAdd.length, budgetSheetHeaders.length)
               .setValues(rowsToAdd);
  }

  ui.alert(`Budget sheet updated: ${updatedCount} item(s) updated, ${addedCount} item(s) added.`);
} 