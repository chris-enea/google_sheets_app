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

  // --- Copy FFE data to Pricing Sheet ---
  const pricingSheetName = "Pricing";
  const ffeSourceColumnsForPricing = ["ROOM", "TYPE", "ITEM", "QUANTITY", "LOW TOTAL", "HIGH TOTAL"];
  const pricingTargetColumns = ["Room", "Item Type", "Item Name", "Quantity", "Budget Low", "Budget High"];
  
  // Check if ffeSheetName is valid (it's defined earlier in FFE Processing Configuration)
  if (ffeSheetName) {
    _copyFFEDataToPriceSheet(ss, ffeSheetName, pricingSheetName, ffeSourceColumnsForPricing, pricingTargetColumns);
  } else {
    SpreadsheetApp.getUi().alert("FFE sheet name was not defined. Cannot copy to Pricing sheet.");
    // This case should ideally not be reached if FFE processing ran.
  }

  // --- SPEC Processing Configuration ---
  const allowancesSheetName = "SPEC";
  const allowancesTargetHeaders = [
    "CATEGORIES", "TYPE", "ITEM", "ACTUAL PRICE",
    "QUANTITY", "LOW", "TOTAL LOW", "HIGH", "TOTAL HIGH", "NOTES"
  ];
  
  const allowancesColumnMapping = [
    { targetHeaderName: "CATEGORIES", isBlank: true, defaultFormatSourceMasterColName: masterHeaders[0] || "" }, // Use first master col for style
    { targetHeaderName: "TYPE", sourceMasterColumnName: MASTER_TYPE_COL_NAME },
    { targetHeaderName: "ITEM", sourceMasterColumnName: MASTER_ITEM_COL_NAME },
    { targetHeaderName: "ACTUAL PRICE", isBlank: true, defaultFormatSourceMasterColName: masterHeaders[0] || "" }, // Use first master col for style
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
    const targetHeaderValueRange = targetSheet.getRange(1, 1, 1, targetHeadersArray.length);
    targetHeaderValueRange.setValues([targetHeadersArray]); // Set all header values first.

    // Then apply formatting cell by cell
    for (let i = 0; i < targetHeadersArray.length; i++) {
      const targetCell = targetSheet.getRange(1, i + 1);
      const headerText = targetHeadersArray[i]; // Primarily for logging or potential direct matching if needed
      const currentMapping = columnMappingConfig[i];
      
      let sourceFormatMasterColZeroBasedIndex = -1;

      // Determine source column for formatting using a cascade of priorities:
      // 1. From explicitly mapped sourceMasterColumnName (via currentMapping.masterColumnIndex)
      if (currentMapping.masterColumnIndex !== undefined && currentMapping.masterColumnIndex >= 0) {
        sourceFormatMasterColZeroBasedIndex = currentMapping.masterColumnIndex;
      } else {
        // 2. Direct match of target header text to a master header text
        const directMatchInMaster = masterHeaders.indexOf(headerText);
        if (directMatchInMaster !== -1) {
          sourceFormatMasterColZeroBasedIndex = directMatchInMaster;
        } 
        // 3. From defaultFormatSourceMasterColName in mapping (for blank/new columns, via currentMapping.defaultFormatMasterColIndex)
        else if (currentMapping.defaultFormatMasterColIndex !== undefined && currentMapping.defaultFormatMasterColIndex >= 0) {
          sourceFormatMasterColZeroBasedIndex = currentMapping.defaultFormatMasterColIndex;
        } 
        // 4. Absolute fallback: style from the first column of master headers, if master has headers
        else if (masterHeaders.length > 0) {
          sourceFormatMasterColZeroBasedIndex = 0; 
        }
      }

      // Apply format if a valid source column index was found and master sheet is usable
      if (sourceFormatMasterColZeroBasedIndex !== -1 && 
          masterSheet.getLastRow() > 0 && 
          (masterSheet.getLastColumn() - 1) >= sourceFormatMasterColZeroBasedIndex) { // Ensure source column index is within bounds
        try {
          const sourceCellToCopyFormatFrom = masterSheet.getRange(1, sourceFormatMasterColZeroBasedIndex + 1);
          sourceCellToCopyFormatFrom.copyTo(targetCell, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
          // Value is already set by targetHeaderValueRange.setValues() above. PASTE_FORMAT should not clear it.
        } catch (e) {
          Logger.log(`Error copying format for Allowances header '${headerText}' from master column index ${sourceFormatMasterColZeroBasedIndex}: ${e.toString()}`);
          targetCell.setFontWeight("bold"); // Fallback on error during copy
        }
      } else {
        // Fallback if no source format could be identified, or master sheet is empty/too small
        targetCell.setFontWeight("bold"); 
      }
    }
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

      // Determine the target row number for A1 notation formulas
      // processedRows contains data rows. Header is row 1. So, first data row is 2.
      // Length of processedRows is current number of data rows already added.
      // So, the next row to be added will be at index processedRows.length, making its sheet row number processedRows.length + 2.
      const targetSheetRowNum = processedRows.length + 2;

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
              // Existing formula logic
              let baseFormula = masterRowFormulas[srcMasterColIdx];

              if (targetSheetName === "SPEC") {
                if (mapping.targetHeaderName === "TOTAL LOW") {
                  const qtyColLetter = getColumnLetter(targetHeadersArray.indexOf(MASTER_QTY_COL_NAME));
                  const lowColLetter = getColumnLetter(targetHeadersArray.indexOf(MASTER_LOW_UNIT_COST_COL_NAME));
                  if (qtyColLetter && lowColLetter) {
                    baseFormula = `=${lowColLetter}${targetSheetRowNum}*${qtyColLetter}${targetSheetRowNum}`;
                  } else {
                    Logger.log(`Could not find columns for SPEC 'TOTAL LOW' formula. QTY header: ${MASTER_QTY_COL_NAME} (index: ${targetHeadersArray.indexOf(MASTER_QTY_COL_NAME)}), LOW header: ${MASTER_LOW_UNIT_COST_COL_NAME} (index: ${targetHeadersArray.indexOf(MASTER_LOW_UNIT_COST_COL_NAME)})`);
                    // Fallback to master formula or error - original behavior if columns not found
                  }
                } else if (mapping.targetHeaderName === "TOTAL HIGH") {
                  const qtyColLetter = getColumnLetter(targetHeadersArray.indexOf(MASTER_QTY_COL_NAME));
                  const highColLetter = getColumnLetter(targetHeadersArray.indexOf(MASTER_HIGH_UNIT_COST_COL_NAME));
                  if (qtyColLetter && highColLetter) {
                    baseFormula = `=${highColLetter}${targetSheetRowNum}*${qtyColLetter}${targetSheetRowNum}`;
                  } else {
                    Logger.log(`Could not find columns for SPEC 'TOTAL HIGH' formula. QTY header: ${MASTER_QTY_COL_NAME} (index: ${targetHeadersArray.indexOf(MASTER_QTY_COL_NAME)}), HIGH header: ${MASTER_HIGH_UNIT_COST_COL_NAME} (index: ${targetHeadersArray.indexOf(MASTER_HIGH_UNIT_COST_COL_NAME)})`);
                    // Fallback to master formula or error - original behavior if columns not found
                  }
                }
              }
              cellValue = baseFormula;
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
              try {
                // Ensure masterColumnIndex is within bounds of actual masterSheet columns before calling getColumnWidth
                if ((mapping.masterColumnIndex + 1) > masterSheet.getLastColumn()) {
                    Logger.log(`WARNING: Master column index ${mapping.masterColumnIndex + 1} for target header '${mapping.targetHeaderName}' (target col ${col + 1}) is out of bounds for masterSheet (last column: ${masterSheet.getLastColumn()}). Using default width.`);
                    // widthToSet remains DEFAULT_TARGET_COL_WIDTH, which is already set
                } else {
                    widthToSet = masterSheet.getColumnWidth(mapping.masterColumnIndex + 1);
                }
              } catch (e) {
                Logger.log(`Error getting column width for target header '${mapping.targetHeaderName}' (target col ${col + 1}) from master col ${mapping.masterColumnIndex + 1}: ${e.toString()}. Using default width.`);
                // widthToSet remains DEFAULT_TARGET_COL_WIDTH, which is already set
              }
          } else if (mapping && mapping.isBlank) {
              // widthToSet remains DEFAULT_TARGET_COL_WIDTH for blank columns, already set
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

// Helper function to convert 0-indexed column to A1 letter
function getColumnLetter(colIndex) {
  let temp, letter = '';
  while (colIndex >= 0) {
    temp = colIndex % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    colIndex = Math.floor(colIndex / 26) - 1;
  }
  return letter;
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

// --- Function to copy specified columns from FFE to Price Sheet ---
/**
 * Copies specified columns from the FFE sheet to the Price sheet.
 * Assumes row-by-row correspondence after headers.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {string} ffeSheetName The name of the source FFE sheet.
 * @param {string} pricingSheetName The name of the target Pricing sheet.
 * @param {Array<string>} ffeSourceHeaders Array of header names of columns to copy from FFE sheet.
 * @param {Array<string>} pricingTargetHeaders Array of header names for the Pricing sheet (in corresponding order to ffeSourceHeaders).
 */
function _copyFFEDataToPriceSheet(ss, ffeSheetName, pricingSheetName, ffeSourceHeaders, pricingTargetHeaders) {
  const ui = SpreadsheetApp.getUi();

  const ffeSheet = ss.getSheetByName(ffeSheetName);
  if (!ffeSheet) {
    ui.alert(`Sheet "${ffeSheetName}" not found. Cannot copy data to "${pricingSheetName}".`);
    return;
  }

  const ffeDataRange = ffeSheet.getDataRange();
  const ffeValues = ffeDataRange.getValues();

  if (ffeValues.length <= 1) { // Only headers or empty
    ui.alert(`Sheet "${ffeSheetName}" has no data to copy to "${pricingSheetName}".`);
    return; // Stop if FFE has no data, as there's nothing to copy.
  }

  const ffeHeaderRow = ffeValues[0].map(String);
  const ffeSourceIndices = ffeSourceHeaders.map(header => ffeHeaderRow.indexOf(header));

  const missingSourceHeadersInFFE = [];
  ffeSourceHeaders.forEach((header, index) => {
    if (ffeSourceIndices[index] === -1) {
      missingSourceHeadersInFFE.push(header);
    }
  });

  if (missingSourceHeadersInFFE.length > 0) {
    ui.alert(`The following source columns were not found in the "${ffeSheetName}" sheet: ${missingSourceHeadersInFFE.join(', ')}. Cannot copy data to "${pricingSheetName}".`);
    return;
  }

  const ffeDataRows = [];
  if (ffeValues.length > 1) {
    for (let i = 1; i < ffeValues.length; i++) {
      const ffeRow = ffeValues[i];
      const pricingRowValues = ffeSourceIndices.map(sourceIndex => ffeRow[sourceIndex]);
      ffeDataRows.push(pricingRowValues);
    }
  }
  
  if (ffeDataRows.length === 0) { // Should be caught by ffeValues.length <=1, but as a safeguard.
      ui.alert(`Sheet "${ffeSheetName}" has no data rows to copy to "${pricingSheetName}".`);
      return;
  }

  let pricingSheet = ss.getSheetByName(pricingSheetName);
  let targetColumnIndicesInPricingSheet = []; // To store 1-based column numbers
  let writeHeaders = false;

  if (!pricingSheet) {
    pricingSheet = ss.insertSheet(pricingSheetName);
    writeHeaders = true;
    // The target columns are simply 1 to N in this new sheet
    targetColumnIndicesInPricingSheet = pricingTargetHeaders.map((_, i) => i + 1);
    Logger.log(`Created sheet "${pricingSheetName}". Headers will be written.`);
  } else {
    const actualPricingSheetHeaders = pricingSheet.getRange(1, 1, 1, pricingSheet.getLastColumn()).getValues()[0].map(String);
    const missingTargetHeadersInPricing = [];
    
    // Populate targetColumnIndicesInPricingSheet based on presence of pricingTargetHeaders
    for (const targetHeader of pricingTargetHeaders) {
      const index = actualPricingSheetHeaders.indexOf(targetHeader);
      if (index === -1) {
        missingTargetHeadersInPricing.push(targetHeader);
      } else {
        targetColumnIndicesInPricingSheet.push(index + 1); // 1-based column index
      }
    }

    if (missingTargetHeadersInPricing.length > 0) {
      ui.alert(`The "${pricingSheetName}" sheet is missing the following required columns: "${missingTargetHeadersInPricing.join(', ')}". Please add these columns or ensure they are named correctly in the first row.`);
      return;
    }
    
    // Check for existing data in target columns from row 2 downwards
    if (pricingSheet.getLastRow() > 1) {
      for (const colIndexOneBased of targetColumnIndicesInPricingSheet) {
        const columnDataRange = pricingSheet.getRange(2, colIndexOneBased, pricingSheet.getLastRow() - 1, 1);
        const columnValues = columnDataRange.getValues();
        // Check if any cell in the column (from row 2) has content
        if (columnValues.flat().some(cell => cell !== "" && cell !== null)) {
          ui.alert(`The "${pricingSheetName}" sheet already contains data in one or more target columns (e.g., ${pricingTargetHeaders.join('/')}) starting from row 2. Please clear this data manually if you wish to proceed.`);
          return;
        }
      }
    }
    Logger.log(`Found required headers in "${pricingSheetName}". Target columns are clear from row 2.`);
  }

  // Write headers if the sheet was newly created
  if (writeHeaders) {
    pricingSheet.getRange(1, 1, 1, pricingTargetHeaders.length).setValues([pricingTargetHeaders]).setFontWeight("bold");
    Logger.log(`Wrote headers to new sheet "${pricingSheetName}".`);
  }
  
  // Write FFE data to the identified target columns in Pricing sheet
  // This check ensures we have a valid mapping of target columns before attempting to write
  if (ffeDataRows.length > 0 && targetColumnIndicesInPricingSheet.length === pricingTargetHeaders.length) {
    for (let rowIndex = 0; rowIndex < ffeDataRows.length; rowIndex++) {
      const rowData = ffeDataRows[rowIndex]; // This is an array of values from FFE for one row
      for (let colDataIndex = 0; colDataIndex < rowData.length; colDataIndex++) {
        const targetSheetCol = targetColumnIndicesInPricingSheet[colDataIndex];
        pricingSheet.getRange(rowIndex + 2, targetSheetCol).setValue(rowData[colDataIndex]);
      }
    }
    Logger.log(`Copied ${ffeDataRows.length} data rows to specific columns in "${pricingSheetName}".`);
  } else if (ffeDataRows.length > 0) {
      Logger.log(`Mismatch between number of target columns identified (${targetColumnIndicesInPricingSheet.length}) and number of source columns mapped (${pricingTargetHeaders.length}). Data not written to "${pricingSheetName}".`);
      // This case implies an issue with populating targetColumnIndicesInPricingSheet correctly.
  }

  // Auto-resize the target columns in the pricing sheet
  if (targetColumnIndicesInPricingSheet.length > 0) {
    for (const colIndexOneBased of targetColumnIndicesInPricingSheet) {
      pricingSheet.autoResizeColumn(colIndexOneBased);
    }
  }
} 