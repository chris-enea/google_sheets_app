/**
 * Budget Management Module for Google Apps Script Project
 * 
 * This file contains all the functions related to handling budget data and visualization
 * including fetching budget data and processing it.
 * 
 * The core functionality includes:
 * - Retrieving and processing budget data from the Budget sheet
 * - Calculating budget summaries and room-by-room breakdowns
 * - Displaying budget views in the main dashboard
 */

/**
 * Fetches and processes budget data from the Budget sheet
 * 
 * This function:
 * 1. Retrieves data from the Budget sheet
 * 2. Processes it into a structured format with room-by-room breakdowns
 * 3. Calculates budget totals, spent amounts, and remaining balances
 * 
 * @param {string} sheetId - The ID of the spreadsheet to get budget data from
 * @return {Object} Processed budget data with summary and rooms information
 */
function getBudgetData(sheetId) {
    try {
      // Use the provided sheetId if available, otherwise fall back to active spreadsheet
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const budgetSheet = ss.getSheetByName('Master Items List');
      
      if (!budgetSheet) {
        Logger.log('Budget sheet not found. Trying to use Items sheet instead.');
        // Try to use the Items sheet as a fallback
        const itemsSheet = ss.getSheetByName('Master Items List');
        if (!itemsSheet) {
          throw new Error('Neither Budget nor Items sheet found');
        }
        
        // Get data from Items sheet
        const dataRange = itemsSheet.getDataRange();
        const values = dataRange.getValues();
        
        Logger.log("Using Items sheet with " + values.length + " rows");
        
        // Process data from Items sheet instead
        return processItemsSheetForBudget(values);
      }
      
      // Get all data from the Budget sheet
      const dataRange = budgetSheet.getDataRange();
      const values = dataRange.getValues();
      
      Logger.log("Total rows in Budget sheet: " + values.length);
      
      // Skip header row and process data rows
      const headerRow = values[0];
      const dataRows = values.slice(1);
      
      Logger.log("Number of data rows: " + dataRows.length);
      
      // Create index map for columns
      const colIndexes = {};
      headerRow.forEach((header, index) => {
        colIndexes[header.toString().trim()] = index;
      });
      
      Logger.log("Column indexes: " + JSON.stringify(colIndexes));
      
      // Required columns
      const requiredColumns = ['ROOM', 'TYPE', 'ITEM', 'QUANTITY', 'LOW BUDGET', 'LOW BUDGET TOTAL', 'HIGH BUDGET', 'HIGH BUDGET TOTAL'];
      
      // Verify all required columns exist
      requiredColumns.forEach(column => {
        if (colIndexes[column] === undefined) {
          throw new Error(`Required column "${column}" not found in Budget sheet`);
        }
      });
      
      // Process room data
      const roomData = {};
      let totalLowBudget = 0;
      let totalHighBudget = 0;
      
      dataRows.forEach((row, index) => {
        // Skip empty rows
        if (!row[colIndexes['ROOM']] || !row[colIndexes['ITEM']]) {
          Logger.log("Skipping empty row at index: " + (index + 2));
          return;
        }
        
        const roomName = row[colIndexes['ROOM']].toString().trim();
        const itemType = row[colIndexes['TYPE']].toString().trim();
        const itemName = row[colIndexes['ITEM']].toString().trim();
        const quantity = parseFloat(row[colIndexes['QUANTITY']]) || 0;
        const lowPrice = parseFloat(row[colIndexes['LOW BUDGET']]) || 0;
        const lowTotal = parseFloat(row[colIndexes['LOW BUDGET TOTAL']]) || 0;
        const highPrice = parseFloat(row[colIndexes['HIGH BUDGET']]) || 0;
        const highTotal = parseFloat(row[colIndexes['HIGH BUDGET TOTAL']]) || 0;
        
        Logger.log("Processing row " + (index + 2) + ": Room=" + roomName + ", Item=" + itemName);
        
        // Initialize room if not exists
        if (!roomData[roomName]) {
          roomData[roomName] = {
            name: roomName,
            lowBudget: 0,
            highBudget: 0,
            items: []
          };
          Logger.log("Created new room: " + roomName);
        }
        
        // Add item to room
        roomData[roomName].items.push({
          item: itemName,
          type: itemType,
          quantity: quantity,
          low: lowPrice,
          lowTotal: lowTotal,
          high: highPrice,
          highTotal: highTotal
        });
        
        // Add to room totals
        roomData[roomName].lowBudget += lowTotal;
        roomData[roomName].highBudget += highTotal;
        
        // Add to overall totals
        totalLowBudget += lowTotal;
        totalHighBudget += highTotal;
      });
      
      Logger.log("Found rooms: " + Object.keys(roomData).join(", "));
      
      // Convert rooms object to array and sort by high budget amount (highest first)
      const roomsArray = Object.values(roomData).sort((a, b) => b.highBudget - a.highBudget);
      
      Logger.log("Final number of rooms: " + roomsArray.length);
      
      // Create final result object
      const result = {
        summary: {
          totalLowBudget: totalLowBudget,
          totalHighBudget: totalHighBudget
        },
        rooms: roomsArray
      };
      
      Logger.log("Budget data processed: " + JSON.stringify(result.summary));
      return result;
      
    } catch (error) {
      Logger.log('Error in getBudgetData: ' + error.toString());
      throw new Error('Failed to process budget data: ' + error.message);
    }
  } 

/**
 * Processes data from the Items sheet to create budget summary data
 * when the dedicated Budget sheet is not available.
 * 
 * @param {Array} values - 2D array of values from the Items sheet
 * @return {Object} Processed budget data with summary and rooms information
 */
function processItemsSheetForBudget(values) {
  try {
    if (!values || values.length < 2) {
      throw new Error('Items sheet is empty or has insufficient data');
    }
    
    // Get the header row
    const headerRow = values[0];
    const dataRows = values.slice(1);
    
    // Map column indices
    const colIndexes = {};
    headerRow.forEach((header, index) => {
      colIndexes[header.toString().trim()] = index;
    });
    
    // Check if required columns exist
    const roomIndex = colIndexes['ROOM'] || -1;
    const itemIndex = colIndexes['ITEM'] || -1;
    const typeIndex = colIndexes['TYPE'] || -1;
    const quantityIndex = colIndexes['QUANTITY'] || -1;
    
    // Budget columns
    const lowBudgetIndex = colIndexes['LOW BUDGET'] || -1;
    const highBudgetIndex = colIndexes['HIGH BUDGET'] || -1;
    
    // Budget total columns
    const lowBudgetTotalIndex = colIndexes['LOW BUDGET TOTAL'] || -1;
    const highBudgetTotalIndex = colIndexes['HIGH BUDGET TOTAL'] || -1;
    
    if (roomIndex === -1 || itemIndex === -1 || lowBudgetIndex === -1 || highBudgetIndex === -1) {
      throw new Error('Required columns missing in Items sheet: ROOM, ITEM, LOW BUDGET, and HIGH BUDGET columns are needed.');
    }
    
    // Process the data by room
    const roomData = {};
    let totalLowBudget = 0;
    let totalHighBudget = 0;
    
    dataRows.forEach((row, index) => {
      // Skip empty rows
      if (!row[roomIndex] || !row[itemIndex]) {
        return;
      }
      
      const roomName = row[roomIndex].toString().trim();
      const itemName = row[itemIndex].toString().trim();
      const itemType = typeIndex !== -1 ? row[typeIndex].toString().trim() : '';
      const quantity = quantityIndex !== -1 ? (parseInt(row[quantityIndex]) || 1) : 1;
      const lowBudget = parseFloat(row[lowBudgetIndex]) || 0;
      const highBudget = parseFloat(row[highBudgetIndex]) || 0;
      
      // Calculate totals - use budget total columns if available, otherwise calculate
      let lowTotal, highTotal;
      
      if (lowBudgetTotalIndex !== -1 && row[lowBudgetTotalIndex] !== null && !isNaN(row[lowBudgetTotalIndex])) {
        lowTotal = parseFloat(row[lowBudgetTotalIndex]) || 0;
      } else {
        lowTotal = lowBudget * quantity;
      }
      
      if (highBudgetTotalIndex !== -1 && row[highBudgetTotalIndex] !== null && !isNaN(row[highBudgetTotalIndex])) {
        highTotal = parseFloat(row[highBudgetTotalIndex]) || 0;
      } else {
        highTotal = highBudget * quantity;
      }
      
      // Initialize room if not exists
      if (!roomData[roomName]) {
        roomData[roomName] = {
          name: roomName,
          lowBudget: 0,
          highBudget: 0,
          items: []
        };
      }
      
      // Add item to room
      roomData[roomName].items.push({
        item: itemName,
        type: itemType,
        quantity: quantity,
        low: lowBudget,
        lowTotal: lowTotal,
        high: highBudget,
        highTotal: highTotal
      });
      
      // Add to room totals
      roomData[roomName].lowBudget += lowTotal;
      roomData[roomName].highBudget += highTotal;
      
      // Add to overall totals
      totalLowBudget += lowTotal;
      totalHighBudget += highTotal;
    });
    
    // Convert rooms object to array and sort by high budget amount
    const roomsArray = Object.values(roomData).sort((a, b) => b.highBudget - a.highBudget);
    
    // Create final result object
    const result = {
      summary: {
        totalLowBudget: totalLowBudget,
        totalHighBudget: totalHighBudget
      },
      rooms: roomsArray
    };
    
    Logger.log("Budget data processed from Items sheet: " + JSON.stringify(result.summary));
    return result;
    
  } catch (error) {
    Logger.log('Error in processItemsSheetForBudget: ' + error.toString());
    throw new Error('Failed to process budget data from Items sheet: ' + error.message);
  }
} 