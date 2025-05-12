/**
 * Item Management Module
 * Contains all functionality related to managing project items and rooms.
 */

/**
 * Fetches all rooms from the Data sheet.
 * 
 * @return {Object} Object containing rooms data and success status
 */
function getRooms() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const dataSheet = ss.getSheetByName("Data");
      
      if (!dataSheet) {
        return {
          success: false,
          error: "Data sheet not found in the spreadsheet"
        };
      }
      
      // Get the range containing rooms in column A
      const dataRange = dataSheet.getRange("A:A");
      const values = dataRange.getValues();
      
      // Find the data table - look for header "Room" in column A
      let headerRowIndex = -1;
      for (let i = 0; i < values.length; i++) {
        if (values[i][0] === "Rooms") {
          headerRowIndex = i;
          break;
        }
      }
      
      if (headerRowIndex === -1) {
        return {
          success: false,
          error: "Room header not found in column A of Data sheet"
        };
      }
      
      // Extract rooms (skipping the header row)
      const rooms = [];
      for (let i = headerRowIndex + 1; i < values.length; i++) {
        const roomName = values[i][0];
        // Stop if we hit an empty cell
        if (!roomName) {
          break;
        }
        rooms.push(roomName);
      }
      
      Logger.log(`Found ${rooms.length} rooms in the Data sheet`);
      return {
        success: true,
        rooms: rooms,
        headerRowIndex: headerRowIndex
      };
      
    } catch (error) {
      Logger.log("Error in getRooms: " + error.message);
      return {
        success: false,
        error: "Error retrieving rooms: " + error.message
      };
    }
  }
  
  /**
   * Adds a new room to the Data sheet.
   * 
   * @param {string} roomName - The name of the room to add
   * @return {Object} Object containing success status
   */
  function addRoom(roomName) {
    try {
      if (!roomName || roomName.trim() === "") {
        return {
          success: false,
          error: "Room name cannot be empty"
        };
      }
      
      // Convert room name to uppercase
      const uppercaseRoomName = roomName.trim().toUpperCase();
      
      // Get current rooms to check for duplicates and find insertion point
      const roomsResult = getRooms();
      if (!roomsResult.success) {
        return roomsResult; // Forward the error
      }
      
      const existingRooms = roomsResult.rooms;
      const headerRowIndex = roomsResult.headerRowIndex;
      
      // Check for duplicate (case insensitive)
      if (existingRooms.some(room => room.toUpperCase() === uppercaseRoomName)) {
        return {
          success: false,
          error: `Room "${uppercaseRoomName}" already exists`
        };
      }
      
      // Get the Data sheet
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const dataSheet = ss.getSheetByName("Data");
      
      // Calculate the row to insert at (header row + existing rooms + 1)
      const insertRowIndex = headerRowIndex + 1 + existingRooms.length;
      
      // Insert the new room (in uppercase)
      dataSheet.getRange(insertRowIndex + 1, 1).setValue(uppercaseRoomName);
      
      Logger.log(`Added new room "${uppercaseRoomName}" at row ${insertRowIndex + 1}`);
      return {
        success: true,
        roomName: uppercaseRoomName
      };
      
    } catch (error) {
      Logger.log("Error in addRoom: " + error.message);
      return {
        success: false,
        error: "Error adding room: " + error.message
      };
    }
  }
  
  /**
   * Shows the item manager dialog.
   */
  function showItemManager() {
    const html = HtmlService.createTemplateFromFile('ItemManager')
      .evaluate()
      .setWidth(1500)
      .setHeight(1000)
      .setTitle('Item Manager');
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Item Manager');
  }
  
  /**
   * CORE UTILITY FUNCTIONS
   * These functions provide shared functionality for both the dialog and dashboard interfaces
   */
  
  /**
   * Core function to retrieve selected rooms from the temporary sheet.
   * Used by both dashboard and dialog interfaces.
   * 
   * @param {string} sheetId - Optional spreadsheet ID. If not provided, uses active spreadsheet.
   * @return {Array} Array of selected room names
   */
  function getSelectedRoomsCore(sheetId) {
    try {
      // Use provided sheetId if available, otherwise fallback to active spreadsheet
      const ss = sheetId 
        ? SpreadsheetApp.openById(sheetId) 
        : SpreadsheetApp.getActiveSpreadsheet();
      
      const tempSheet = ss.getSheetByName("_TempSelectedRooms");
      let selectedRooms = [];
      
      if (tempSheet) {
        // Get rooms from the temporary sheet (skip header row)
        const lastRow = tempSheet.getLastRow();
        if (lastRow > 1) {
          const tempRoomsRange = tempSheet.getRange(2, 1, lastRow - 1, 1);
          const tempRoomsValues = tempRoomsRange.getValues();
          
          // Extract room names
          selectedRooms = tempRoomsValues.map(row => row[0]).filter(room => room);
          Logger.log(`Core: Found ${selectedRooms.length} selected rooms in temp sheet`);
        }
      }
      
      return selectedRooms;
    } catch (error) {
      Logger.log("Error in getSelectedRoomsCore: " + error.message);
      return [];
    }
  }
  
  /**
   * Core function to save selected rooms to the temporary sheet.
   * Used by both dashboard and dialog interfaces.
   * 
   * @param {Array} selectedRooms - Array of room names to save
   * @param {string} sheetId - Optional spreadsheet ID. If not provided, uses active spreadsheet.
   * @return {Boolean} Success status
   */
  function saveSelectedRoomsCore(selectedRooms, sheetId) {
    try {
      if (!selectedRooms || !Array.isArray(selectedRooms)) {
        Logger.log("Invalid rooms data passed to saveSelectedRoomsCore");
        return false;
      }
      
      // Use provided sheetId if available, otherwise fallback to active spreadsheet
      const ss = sheetId 
        ? SpreadsheetApp.openById(sheetId) 
        : SpreadsheetApp.getActiveSpreadsheet();
      
      // Create or get the temporary sheet
      let tempSheet = ss.getSheetByName("_TempSelectedRooms");
      if (tempSheet) {
        // Clear existing content if sheet exists
        tempSheet.clear();
      } else {
        // Create the sheet if it doesn't exist
        tempSheet = ss.insertSheet("_TempSelectedRooms");
        // Hide the sheet as it's for temporary storage only
        tempSheet.hideSheet();
      }
      
      // Add header row
      tempSheet.getRange(1, 1).setValue("Selected Rooms");
      
      // Add the selected rooms
      selectedRooms.forEach((room, index) => {
        tempSheet.getRange(index + 2, 1).setValue(room);
      });
      
      Logger.log(`Core: Saved ${selectedRooms.length} selected rooms to temporary sheet`);
      return true;
    } catch (error) {
      Logger.log("Error in saveSelectedRoomsCore: " + error.message);
      return false;
    }
  }
  
  /**
   * Core function to prepare item data for display and editing.
   * Used by both dashboard and dialog interfaces.
   * 
   * @param {Array} selectedRooms - Optional array of room names to create placeholder items for
   * @param {string} sheetId - Optional spreadsheet ID. If not provided, uses active spreadsheet.
   * @return {Object} Prepared item data
   */
  function prepareItemDataCore(selectedRooms = null, sheetId = null) {
    try {
      // If no selected rooms were provided, get them from the temp sheet
      if (!selectedRooms || !Array.isArray(selectedRooms) || selectedRooms.length === 0) {
        selectedRooms = getSelectedRoomsCore(sheetId);
      }
      
      // Get raw items data from Items sheet, passing along the sheetId
      const itemsData = getItemsData(selectedRooms, sheetId);
      
      // Add any additional processing needed by both interfaces here
      
      return {
        success: true,
        ...itemsData
      };
    } catch (error) {
      Logger.log("Error in prepareItemDataCore: " + error.message);
      return {
        success: false,
        error: "Error preparing item data: " + error.message
      };
    }
  }
  
  /**
   * Shows the item selector dialog for adding items to selected rooms.
   * 
   * @param {Array} selectedRooms - Array of room names selected by the user
   */
  function showItemSelector(selectedRooms) {
    try {
      // Preload items data
      const preloadedData = preloadItemData();
      
      // Create HTML template and pass the preloaded data
      const template = HtmlService.createTemplateFromFile('ItemSelector');
      template.preloadedAvailableItems = preloadedData.availableItems;
      template.preloadedCombinedItems = preloadedData.combinedItems;
      
      if (selectedRooms && Array.isArray(selectedRooms)) {
        template.selectedRoomsJson = JSON.stringify(selectedRooms);
      }
      
      const html = template
        .evaluate()
        .setWidth(1500)
        .setHeight(1000)
        .setTitle('Add Items to Rooms');
      
      SpreadsheetApp.getUi().showModalDialog(html, 'Add Items to Rooms');
    } catch (error) {
      Logger.log("Error in showItemSelector: " + error.message);
      SpreadsheetApp.getUi().alert("Error showing item selector: " + error.message);
    }
  }
  
  /**
   * Gets available items from the Data sheet with caching
   * 
   * @return {Object} Result object with success status and available items array
   */
  function getAvailableItems() {
    try {
      // Check cache first
      const cache = CacheService.getScriptCache();
      const cachedItems = cache.get('availableItems');
      
      if (cachedItems) {
        const items = JSON.parse(cachedItems);
        Logger.log(`Retrieved ${items.length} available items from cache`);
        return {
          success: true,
          availableItems: items
        };
      }
      
      Logger.log("No cached data found, fetching available items from Data sheet");
      
      // Get the active spreadsheet
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const dataSheet = ss.getSheetByName("Data");
      
      if (!dataSheet) {
        const error = "Data sheet not found";
        Logger.log(error);
        return {
          success: false,
          error: error
        };
      }
      
      // Get header row first to find the column indices
      const headerRange = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn());
      const headerValues = headerRange.getValues()[0];
      
      // Find column indices
      const typeColIndex = headerValues.indexOf("Item-Type");
      const itemColIndex = headerValues.indexOf("Item-Name");
      
      if (typeColIndex === -1 || itemColIndex === -1) {
        const error = "Required columns 'Item-Type' and/or 'Item-Name' not found in Data sheet";
        Logger.log(error);
        return {
          success: false,
          error: error
        };
      }
      
      // Get all data at once to minimize calls
      const dataRange = dataSheet.getDataRange();
      const values = dataRange.getValues();
      
      if (values.length <= 1) {
        Logger.log("Data sheet is empty or contains only headers");
        return {
          success: true,
          availableItems: []
        };
      }
      
      // Process data rows (skip header row)
      const availableItems = [];
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        
        // Skip empty rows
        if (!row[typeColIndex] && !row[itemColIndex]) {
          continue;
        }
        
        // Only add if the item name is present
        if (row[itemColIndex]) {
          availableItems.push({
            type: row[typeColIndex] || "",
            item: row[itemColIndex] || ""
          });
        }
      }
      
      // Cache the result for 10 minutes (600 seconds)
      cache.put('availableItems', JSON.stringify(availableItems), 600);
      
      Logger.log(`Processed ${availableItems.length} available items from Data sheet and cached the result`);
      
      return {
        success: true,
        availableItems: availableItems
      };
      
    } catch (e) {
      const errorMsg = `Error getting available items: ${e.toString()}`;
      Logger.log(errorMsg);
      return {
        success: false,
        error: errorMsg
      };
    }
  }
  
  /**
   * Saves the current item selections to a temporary sheet.
   * This allows the selections to be preserved when navigating back and forth.
   * 
   * @param {Object} roomItems - Object with room names as keys and arrays of items as values
   * @return {Object} Object containing success status
   */
  function saveItemSelections(roomItems) {
    try {
      if (!roomItems || typeof roomItems !== 'object') {
        return {
          success: false,
          error: "Invalid room items data"
        };
      }
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      // Create or get the temporary selection sheet
      let tempSelectionSheet = ss.getSheetByName("_TempItemSelections");
      if (tempSelectionSheet) {
        // Clear existing content if sheet exists
        tempSelectionSheet.clear();
      } else {
        // Create the sheet if it doesn't exist
        tempSelectionSheet = ss.insertSheet("_TempItemSelections");
        // Hide the sheet as it's for temporary storage only
        tempSelectionSheet.hideSheet();
      }
      
      // Store the raw JSON data of roomItems for easier retrieval
      let jsonData = JSON.stringify(roomItems);
      tempSelectionSheet.getRange(1, 1).setValue("ROOM_ITEMS_JSON");
      tempSelectionSheet.getRange(1, 2).setValue(jsonData);
      
      Logger.log(`Saved item selections to temporary sheet: ${jsonData}`);
      return {
        success: true
      };
      
    } catch (error) {
      Logger.log("Error in saveItemSelections: " + error.message);
      return {
        success: false,
        error: "Error saving item selections: " + error.message
      };
    }
  }
  
  /**
   * Gets the selected rooms from the temporary sheet and any previously saved item selections.
   * 
   * @param {string} sheetId - Optional spreadsheet ID. If not provided, uses active spreadsheet.
   * @return {Object} Object containing the list of selected rooms and any saved item selections
   */
  function getSelectedRooms(sheetId) {
    try {
      // Use the core function to get the rooms
      const selectedRooms = getSelectedRoomsCore(sheetId);
      
      // Check for any previously saved item selections
      // Use provided sheetId if available, otherwise fallback to active spreadsheet
      const ss = sheetId 
        ? SpreadsheetApp.openById(sheetId) 
        : SpreadsheetApp.getActiveSpreadsheet();
      
      const tempSelectionSheet = ss.getSheetByName("_TempItemSelections");
      let savedSelections = {};
      
      if (tempSelectionSheet) {
        try {
          const jsonCell = tempSelectionSheet.getRange(1, 2);
          const jsonValue = jsonCell.getValue();
          
          if (jsonValue) {
            savedSelections = JSON.parse(jsonValue);
            Logger.log(`Retrieved saved selections: ${jsonValue}`);
          }
        } catch (parseError) {
          Logger.log(`Error parsing saved selections: ${parseError.message}`);
          // Continue with empty savedSelections if there's an error
        }
      }
      
      return {
        success: true,
        selectedRooms: selectedRooms,
        savedSelections: savedSelections
      };
      
    } catch (error) {
      Logger.log("Error in getSelectedRooms: " + error.message);
      return {
        success: false,
        error: "Error retrieving selected rooms: " + error.message
      };
    }
  }
  
  /**
   * Saves items for each room to the spreadsheet.
   * 
   * @param {Object} roomItems - Object with room names as keys and arrays of items as values
   * @return {Object} Object containing success status
   */
  function saveRoomItems(roomItems) {
    try {
      if (!roomItems || typeof roomItems !== 'object') {
        return {
          success: false,
          error: "Invalid room items data"
        };
      }
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      // Create or get the RoomItems sheet
      let roomItemsSheet = ss.getSheetByName("RoomItems");
      if (!roomItemsSheet) {
        // Create the sheet if it doesn't exist
        roomItemsSheet = ss.insertSheet("RoomItems");
        
        // Add headers
        roomItemsSheet.getRange(1, 1, 1, 2).setValues([["Room", "Item"]]);
        roomItemsSheet.getRange(1, 1, 1, 2).setFontWeight("bold");
      } else {
        // Clear existing room-item mappings if sheet exists (except the header)
        const lastRow = Math.max(roomItemsSheet.getLastRow(), 1);
        if (lastRow > 1) {
          roomItemsSheet.getRange(2, 1, lastRow - 1, 2).clearContent();
        }
      }
      
      // Prepare data to write - flatten the room-items structure
      const flatData = [];
      let totalItems = 0;
      
      Object.keys(roomItems).forEach(room => {
        const items = roomItems[room];
        if (Array.isArray(items) && items.length > 0) {
          items.forEach(item => {
            // Check if item is an object with 'item' property or a simple string
            const itemValue = typeof item === 'object' && item !== null ? item.item : item;
            
            if (itemValue && typeof itemValue === 'string' && itemValue.trim() !== '') {
              flatData.push([room, itemValue.trim()]);
              totalItems++;
            }
          });
        }
      });
      
      // Write the data to the sheet
      if (flatData.length > 0) {
        roomItemsSheet.getRange(2, 1, flatData.length, 2).setValues(flatData);
      }
      
      // Keep the temporary sheet for reuse instead of deleting it
      
      Logger.log(`Saved ${totalItems} items for ${Object.keys(roomItems).length} rooms`);
      return {
        success: true,
        itemCount: totalItems,
        roomCount: Object.keys(roomItems).length
      };
      
    } catch (error) {
      Logger.log("Error in saveRoomItems: " + error.message);
      return {
        success: false,
        error: "Error saving room items: " + error.message
      };
    }
  }
  
  /**
   * Saves items data to the Items sheet
   * Optimized for batch operations
   * 
   * @param {Array} items - Array of item objects
   * @return {Object} Result object with success status and count
   */
  function saveItemsToSheet(items) {
    try {
      // Validate input
      if (!items || !Array.isArray(items) || items.length === 0) {
        Logger.log("No valid items to save");
        return { success: false, error: "No valid items to save" };
      }
      
      Logger.log(`Saving ${items.length} items to the sheet`);
      
      // Get reference to spreadsheet - using one call to SpreadsheetApp
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      // Get or create the Items sheet
      let sheet = ss.getSheetByName("Items");
      if (!sheet) {
        Logger.log("Creating new Items sheet");
        sheet = ss.insertSheet("Items");
      }
      
      // Clear existing content but keep the sheet
      sheet.clear();
      
      // Set up headers
      const headers = [
        "Room", 
        "Type", 
        "Item", 
        "Quantity", 
        "Low Budget", 
        "Low Budget Total", 
        "High Budget", 
        "High Budget Total"
      ];
      
      // Prepare all data rows at once to minimize service calls
      const rows = [headers];
      
      // Process and validate each item
      items.forEach(item => {
        // Ensure quantity is at least 1
        const quantity = Math.max(1, parseInt(item.quantity) || 1);
        
        // Handle null/undefined/empty budget values
        const lowBudget = (item.lowBudget !== null && item.lowBudget !== undefined && item.lowBudget !== '') 
          ? parseFloat(item.lowBudget) || 0 
          : '';
        
        const highBudget = (item.highBudget !== null && item.highBudget !== undefined && item.highBudget !== '') 
          ? parseFloat(item.highBudget) || 0
          : '';
        
        // Calculate totals only if budget values exist
        const lowBudgetTotal = lowBudget !== '' ? lowBudget * quantity : '';
        const highBudgetTotal = highBudget !== '' ? highBudget * quantity : '';
        
        // Add the row
        rows.push([
          item.room || '',
          item.type || '',
          item.item || '',
          quantity,
          lowBudget,
          lowBudgetTotal,
          highBudget,
          highBudgetTotal
      ]);
      });
      
      // Write all data in a single operation
      const range = sheet.getRange(1, 1, rows.length, headers.length);
      range.setValues(rows);
        
        // Format number columns
      if (rows.length > 1) {
        // Format quantity column as whole numbers
        sheet.getRange(2, 4, rows.length - 1, 1).setNumberFormat('0');
        
        // Format budget columns as currency
        const budgetCols = [5, 6, 7, 8]; // Columns E, F, G, H
        budgetCols.forEach(col => {
          sheet.getRange(2, col, rows.length - 1, 1).setNumberFormat('$#,##0.00');
        });
      }
      
      Logger.log(`Successfully saved ${items.length} items to Items sheet`);
      return {
        success: true,
        count: items.length
      };
      
    } catch (e) {
      const errorMsg = `Error saving items to sheet: ${e.toString()}`;
      Logger.log(errorMsg);
      return {
        success: false,
        error: errorMsg
      };
    }
  }
  
  /**
   * Shows the item update dialog for editing existing items.
   * If selectedRooms are provided, will prepare the form for adding new items to those rooms.
   * 
   * @param {Array} selectedRooms - Optional array of room names selected by the user
   */
  function showItemUpdate(selectedRooms) {
    try {
      // Preload items data
      const preloadedData = preloadItemData();
      
      // Create HTML template and pass the preloaded data
      const template = HtmlService.createTemplateFromFile('ItemUpdate');
      template.preloadedAvailableItems = preloadedData.availableItems;
      template.preloadedCombinedItems = preloadedData.combinedItems;
      
      // Always set selectedRoomsJson (use empty array if not provided)
      template.selectedRoomsJson = JSON.stringify(selectedRooms && Array.isArray(selectedRooms) ? selectedRooms : []);
      
      const html = template
        .evaluate()
        .setWidth(1500)
        .setHeight(1000)
        .setTitle('Update Items');
      
      SpreadsheetApp.getUi().showModalDialog(html, 'Update Items');
    } catch (error) {
      Logger.log("Error in showItemUpdate: " + error.message);
      SpreadsheetApp.getUi().alert("Error showing item update dialog: " + error.message);
    }
      }
      
  /**
   * Helper function to ensure rooms are saved to the temp sheet.
   * This is important so getSelectedRooms and getItemsData can access them.
   * 
   * @param {Array} rooms - Array of room names to save
   */
  function saveRoomsToTempSheet(rooms) {
    // Use the core function
    saveSelectedRoomsCore(rooms);
  }
  
  /**
   * Gets items data from the Items sheet
   * 
   * @param {Array} selectedRooms - Optional array of room names to filter by
   * @param {string} sheetId - Optional spreadsheet ID. If not provided, uses active spreadsheet.
   * @return {Object} Object containing items data or error information
   */
  function getItemsData(selectedRooms, sheetId) {
    try {
      Logger.log("Getting items data from sheet");
      
      // Use provided sheetId if available, otherwise fallback to active spreadsheet
      const ss = sheetId 
        ? SpreadsheetApp.openById(sheetId) 
        : SpreadsheetApp.getActiveSpreadsheet();
      
      const itemsSheet = ss.getSheetByName("Items");
      
      if (!itemsSheet) {
        Logger.log("Items sheet not found, creating one");
        // Create items sheet if it doesn't exist
        const newSheet = ss.insertSheet("Items");
        const headers = ["Room", "Type", "Item", "Quantity", "Low Budget", "Low Budget Total", "High Budget", "High Budget Total"];
        newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        newSheet.setFrozenRows(1);
        
        return {
          success: true,
          items: [],
          itemsByRoom: {}
        };
      }
      
      // Get data range all at once to minimize API calls
      const dataRange = itemsSheet.getDataRange();
      const values = dataRange.getValues();
      
      if (values.length <= 1) {
        Logger.log("Items sheet is empty or contains only headers");
        return {
          success: true,
          items: [],
          itemsByRoom: {}
        };
      }
      
      // Extract headers
      const headers = values[0];
      const roomIndex = headers.indexOf("Room");
      const typeIndex = headers.indexOf("Type");
      const itemIndex = headers.indexOf("Item");
      const quantityIndex = headers.indexOf("Quantity");
      const lowBudgetIndex = headers.indexOf("Low Budget");
      const lowBudgetTotalIndex = headers.indexOf("Low Budget Total");
      const highBudgetIndex = headers.indexOf("High Budget");
      const highBudgetTotalIndex = headers.indexOf("High Budget Total");
      
      // Check required columns
      if (roomIndex === -1 || itemIndex === -1) {
        const error = "Required columns missing in Items sheet";
        Logger.log(error);
        return {
          success: false,
          error: error
        };
      }
      
      // Process data
      const items = [];
      const itemsByRoom = {};
      let totalLowBudget = 0;
      let totalHighBudget = 0;
      
      // Start from row 1 (skip headers)
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        const room = row[roomIndex] || "";
        
        // Skip completely empty rows
        if (!room && !row[itemIndex]) {
          continue;
        }
        
        // Extract and validate item data
        const quantity = parseInt(row[quantityIndex]) || 1;
        let lowBudget = parseFloat(row[lowBudgetIndex]) || null;
        let highBudget = parseFloat(row[highBudgetIndex]) || null;
        
        // Calculate budget totals
        const lowBudgetTotal = lowBudget ? lowBudget * quantity : null;
        const highBudgetTotal = highBudget ? highBudget * quantity : null;
        
        // Accumulate totals
        if (lowBudgetTotal) totalLowBudget += lowBudgetTotal;
        if (highBudgetTotal) totalHighBudget += highBudgetTotal;
        
        const item = {
          room: room,
          type: row[typeIndex] || "",
          item: row[itemIndex] || "",
          quantity: quantity,
          lowBudget: lowBudget,
          lowBudgetTotal: lowBudgetTotal,
          highBudget: highBudget,
          highBudgetTotal: highBudgetTotal
        };
        
        items.push(item);
        
        // Organize items by room
        if (!itemsByRoom[room]) {
          itemsByRoom[room] = [];
        }
        itemsByRoom[room].push(item);
      }
      
      Logger.log(`Processed ${items.length} items from sheet`);
      
      return {
        success: true,
        items: items,
        itemsByRoom: itemsByRoom,
        totalLowBudget: totalLowBudget,
        totalHighBudget: totalHighBudget
      };
      
    } catch (e) {
      const errorMsg = `Error getting items data: ${e.toString()}`;
      Logger.log(errorMsg);
      return {
        success: false,
        error: errorMsg
      };
    }
  }
  
  /**
   * Get room names from the Data sheet.
   * This is a helper function used by both getRoomsForDashboard and other functions.
   * 
   * @param {string} sheetId - Optional spreadsheet ID. If not provided, uses active spreadsheet.
   * @return {Array} Array of room names
   */
  function getRoomNamesFromSheet(sheetId = null) {
    try {
      // Use provided sheetId if available, otherwise fallback to active spreadsheet
      const ss = sheetId 
        ? SpreadsheetApp.openById(sheetId) 
        : SpreadsheetApp.getActiveSpreadsheet();
      
      const dataSheet = ss.getSheetByName("Data");
      
      if (!dataSheet) {
        Logger.log("Data sheet not found");
        return [];
      }
      
      // Get the range containing rooms in column A
      const dataRange = dataSheet.getRange("A:A");
      const values = dataRange.getValues();
      
      // Find the data table - look for header "Rooms" in column A
      let headerRowIndex = -1;
      for (let i = 0; i < values.length; i++) {
        if (values[i][0] === "Rooms") {
          headerRowIndex = i;
          break;
        }
      }
      
      if (headerRowIndex === -1) {
        Logger.log("Rooms header not found in column A of Data sheet");
        return [];
      }
      
      // Extract rooms (skipping the header row)
      const rooms = [];
      for (let i = headerRowIndex + 1; i < values.length; i++) {
        const roomName = values[i][0];
        // Stop if we hit an empty cell
        if (!roomName) {
          break;
        }
        rooms.push(roomName);
      }
      
      return rooms;
    } catch (error) {
      Logger.log("Error in getRoomNamesFromSheet:", error);
      return [];
    }
  }
  
  /**
   * Gets rooms for display in the dashboard
   * Similar to getRooms but formatted for dashboard use
   * 
   * @param {string} sheetId - Optional spreadsheet ID. If not provided, uses active spreadsheet.
   * @return {Object} Object containing rooms data and success status
   */
  function getRoomsForDashboard(sheetId = null) {
    try {
      // Get all room names
      const rooms = getRoomNamesFromSheet(sheetId);
      
      // Get the currently selected rooms from temp sheet using core function
      const selectedRooms = getSelectedRoomsCore(sheetId);
      
      return {
        success: true,
        rooms: rooms,
        selectedRooms: selectedRooms
      };
    } catch (e) {
      Logger.log("Error in getRoomsForDashboard:", e);
      return {
        success: false,
        error: e.toString()
      };
    }
  }
  
  /**
   * Saves selected rooms without immediately showing ItemUpdate dialog
   * This is useful for the dashboard integration
   * 
   * @param {Array} selectedRooms - The array of selected room names
   * @param {string} sheetId - Optional spreadsheet ID. If not provided, uses active spreadsheet.
   * @return {Object} - Result object with success flag
   */
  function saveSelectedRoomsOnly(selectedRooms, sheetId = null) {
    try {
      if (!selectedRooms || !Array.isArray(selectedRooms)) {
        throw new Error("Selected rooms must be an array");
      }
      
      // Save to user properties for backward compatibility
      const userProps = PropertiesService.getUserProperties();
      userProps.setProperty("selectedRooms", JSON.stringify(selectedRooms));
      
      // Use the core function to save rooms to the temp sheet
      const saved = saveSelectedRoomsCore(selectedRooms, sheetId);
      
      if (!saved) {
        return {
          success: false,
          error: "Failed to save selected rooms"
        };
      }
      
      // Return success
      return {
        success: true,
        selectedRooms: selectedRooms
      };
    } catch (e) {
      Logger.log("Error saving selected rooms:", e);
      return {
        success: false,
        error: e.toString()
      };
    }
  }
  
  /**
   * Preloads item data to improve loading performance
   * Returns preloaded data that can be passed directly to HTML templates
   * 
   * @return {Object} Object containing items data for HTML templates
   */
  function preloadItemData() {
    try {
      // Get available items for autocomplete
      const availableItemsResult = getAvailableItems();
      
      // Get combined items using the dedicated function
      const combinedItemsResult = getCombinedItems();
      
      // Create a template data object with the preloaded data
      const templateData = {
        availableItems: availableItemsResult.success ? JSON.stringify(availableItemsResult.items || []) : "[]",
        combinedItems: combinedItemsResult.success ? JSON.stringify(combinedItemsResult.items || []) : "[]"
      };
      
      Logger.log(`Preloaded ${JSON.parse(templateData.availableItems).length} available items and ${JSON.parse(templateData.combinedItems).length} combined items`);
      
      return templateData;
    } catch (error) {
      Logger.log("Error in preloadItemData: " + error.message);
      return {
        availableItems: "[]",
        combinedItems: "[]"
      };
    }
  }
  
  /**
   * ITEM UPDATE CORE FUNCTIONS
   * These core functions provide shared functionality for both the dashboard and dialog interfaces
   */
  
  /**
   * Core function to validate item data before saving.
   * Used by both dashboard and dialog interfaces.
   * 
   * @param {Array} items - Array of item objects to validate
   * @return {Object} Validation result with success flag and any error messages
   */
  function validateItemsDataCore(items) {
    try {
      Logger.log(`Starting validation of ${items ? items.length : 0} items`);
      
      if (!items || !Array.isArray(items) || items.length === 0) {
        Logger.log("No items to validate or invalid input format");
        return {
          success: false,
          error: "No items to validate"
        };
      }
      
      // Check for required fields and validate data types
      let invalidItems = [];
      let validatedItems = [];
      
      items.forEach((item, index) => {
        // Create a new object to avoid mutating the original
        const validatedItem = {}; 
        
        // Required fields validation
        if (!item.room || !item.item || item.item.trim() === "") {
          invalidItems.push({
            index: index,
            item: item,
            reason: "Missing room or item name"
          });
          Logger.log(`Invalid item at index ${index}: Missing room or item name`);
          return; // Skip this item
        }
        
        // Copy all fields and ensure they have appropriate types
        validatedItem.room = String(item.room).trim();
        validatedItem.item = String(item.item).trim();
        validatedItem.type = item.type ? String(item.type).trim() : "";
        
        // Ensure quantity is at least 1 and is a number
        validatedItem.quantity = item.quantity !== undefined && item.quantity !== null ? 
          Math.max(1, parseInt(item.quantity) || 1) : 1;
        
        // Handle budget values - ensure they're proper numbers or null
        if (item.lowBudget !== undefined && item.lowBudget !== null && item.lowBudget !== "" && !isNaN(parseFloat(item.lowBudget))) {
          validatedItem.lowBudget = parseFloat(item.lowBudget);
        } else {
          validatedItem.lowBudget = null;
        }
        
        if (item.highBudget !== undefined && item.highBudget !== null && item.highBudget !== "" && !isNaN(parseFloat(item.highBudget))) {
          validatedItem.highBudget = parseFloat(item.highBudget);
        } else {
          validatedItem.highBudget = null;
        }
        
        // Calculate budget totals
        validatedItem.lowBudgetTotal = validatedItem.lowBudget !== null ? 
          validatedItem.lowBudget * validatedItem.quantity : null;
        
        validatedItem.highBudgetTotal = validatedItem.highBudget !== null ? 
          validatedItem.highBudget * validatedItem.quantity : null;
        
        // Add the validated item to our results
        validatedItems.push(validatedItem);
      });
      
      if (invalidItems.length > 0) {
        Logger.log(`Validation found ${invalidItems.length} invalid items`);
        return {
          success: false,
          error: `${invalidItems.length} invalid items found`,
          invalidItems: invalidItems
        };
      }
      
      Logger.log(`Validation successful, processed ${validatedItems.length} items`);
      return {
        success: true,
        items: validatedItems
      };
      
    } catch (error) {
      Logger.log("Error in validateItemsDataCore: " + error.message);
      return {
        success: false,
        error: "Error validating items: " + error.message
      };
    }
  }
  
  /**
   * Core function to calculate room budget totals from items.
   * Used by both dashboard and dialog interfaces.
   * 
   * @param {Array} items - Array of item objects for a room
   * @return {Object} Budget totals for the room
   */
  function calculateRoomTotalsCore(items) {
    try {
      if (!items || !Array.isArray(items)) {
        return {
          lowTotal: 0,
          highTotal: 0
        };
      }
      
      let lowTotal = 0;
      let highTotal = 0;
      
      items.forEach(item => {
        if (item.lowBudgetTotal !== null) {
          lowTotal += parseFloat(item.lowBudgetTotal);
        }
        
        if (item.highBudgetTotal !== null) {
          highTotal += parseFloat(item.highBudgetTotal);
        }
      });
      
      return {
        lowTotal: lowTotal,
        highTotal: highTotal
      };
    } catch (error) {
      Logger.log("Error in calculateRoomTotalsCore: " + error.message);
      return {
        lowTotal: 0,
        highTotal: 0,
        error: error.message
      };
    }
  }
  
  /**
   * Core function to calculate project-wide budget totals.
   * Used by both dashboard and dialog interfaces.
   * 
   * @param {Array} allItems - Array of all item objects
   * @return {Object} Budget totals for the entire project
   */
  function calculateProjectTotalsCore(allItems) {
    try {
      if (!allItems || !Array.isArray(allItems)) {
        return {
          success: false,
          error: "No items provided"
        };
      }
      
      // Group items by room
      const itemsByRoom = {};
      allItems.forEach(item => {
        if (!itemsByRoom[item.room]) {
          itemsByRoom[item.room] = [];
        }
        itemsByRoom[item.room].push(item);
      });
      
      // Calculate totals for each room
      const roomTotals = {};
      let projectLowTotal = 0;
      let projectHighTotal = 0;
      
      Object.keys(itemsByRoom).forEach(room => {
        const totals = calculateRoomTotalsCore(itemsByRoom[room]);
        roomTotals[room] = totals;
        
        projectLowTotal += totals.lowTotal;
        projectHighTotal += totals.highTotal;
      });
      
      return {
        success: true,
        projectTotals: {
          lowTotal: projectLowTotal,
          highTotal: projectHighTotal
        },
        roomTotals: roomTotals
      };
      
    } catch (error) {
      Logger.log("Error in calculateProjectTotalsCore: " + error.message);
      return {
        success: false,
        error: "Error calculating project totals: " + error.message
      };
    }
  }
  
  /**
   * Core function to save items to spreadsheet.
   * Performs validation and processes items before saving.
   * Used by both dashboard and dialog interfaces.
   * 
   * @param {Array} items - Array of item objects to save
   * @return {Object} Result of the save operation
   */
  function saveItemsCore(items) {
    try {
      Logger.log(`saveItemsCore called with ${items ? items.length : 0} items`);
      
      // Validate items first
      const validationResult = validateItemsDataCore(items);
      if (!validationResult.success) {
        Logger.log(`Validation failed: ${validationResult.error}`);
        return validationResult; // Return validation errors
      }
      
      // Use validated items for saving
      const validatedItems = validationResult.items;
      Logger.log(`Validation passed, proceeding to save ${validatedItems.length} items`);
      
      // Save to spreadsheet using the existing function
      const saveResult = saveItemsToSheet(validatedItems);
      
      if (!saveResult.success) {
        Logger.log(`Failed to save items: ${saveResult.error}`);
      } else {
        Logger.log(`Successfully saved ${saveResult.count} items to sheet`);
      }
      
      return saveResult;
      
    } catch (error) {
      Logger.log("Error in saveItemsCore: " + error.message);
      return {
        success: false,
        error: "Error saving items: " + error.message
      };
    }
  }
  
  /**
   * Core function to prepare item data for UI display.
   * This combines functionality from getItemsData with additional processing.
   * Used by both dashboard and dialog interfaces.
   * 
   * @param {Array} selectedRooms - Array of selected room names
   * @param {string} sheetId - Optional spreadsheet ID. If not provided, uses active spreadsheet.
   * @return {Object} Prepared item data ready for UI display
   */
  function prepareItemsForUICore(selectedRooms, sheetId = null) {
    try {
      // Get base item data
      const itemsData = getItemsData(selectedRooms, sheetId);
      
      if (!itemsData.success) {
        return itemsData; // Return error information
      }
      
      // Get available items for autocomplete
      const availableItemsResult = getAvailableItems();
      
      // Add additional data needed by UI
      const preparedData = {
        success: true,
        selectedRooms: selectedRooms || itemsData.selectedRooms || [],
        items: itemsData.items || [],
        itemsByRoom: itemsData.itemsByRoom || {},
        availableItems: availableItemsResult.success ? availableItemsResult.availableItems || [] : []
      };
      
      // Generate properly formatted combined items for autocomplete
      preparedData.combinedItems = generateCombinedItems(preparedData.availableItems);
      
      // Calculate totals for each room
      const roomTotals = {};
      Object.keys(preparedData.itemsByRoom).forEach(room => {
        roomTotals[room] = calculateRoomTotalsCore(preparedData.itemsByRoom[room]);
      });
      
      preparedData.roomTotals = roomTotals;
      
      return preparedData;
    } catch (error) {
      Logger.log("Error in prepareItemsForUICore: " + error.message);
      return {
        success: false,
        error: "Error preparing item data: " + error.message
      };
    }
  }
  
  /**
   * Get item data prepared for dashboard UI.
   * Wrapper around prepareItemsForUICore for dashboard use.
   * 
   * @param {string} sheetId - Optional spreadsheet ID. If not provided, uses active spreadsheet.
   * @return {Object} Item data prepared for dashboard display
   */
  function getItemUpdateContentForDashboard(sheetId = null) {
    try {
      // Get selected rooms
      const selectedRooms = getSelectedRoomsCore(sheetId);
      
      if (!selectedRooms || selectedRooms.length === 0) {
        return {
          success: false,
          error: "No rooms selected. Please select rooms first."
        };
      }
      
      // Use the core function to prepare items, passing along the sheetId
      return prepareItemsForUICore(selectedRooms, sheetId);
      
    } catch (error) {
      Logger.log("Error in getItemUpdateContentForDashboard: " + error.message);
      return {
        success: false,
        error: "Error preparing dashboard content: " + error.message
      };
    }
  }
  
  /**
   * Saves items data from dashboard or dialog UI.
   * Wrapper around saveItemsCore for consistent saving from any UI.
   * 
   * @param {Array} items - Array of item objects with room, type, item, quantity, lowBudget, etc.
   * @return {Object} Object containing success status
   */
  function saveItemsFromUI(items) {
    try {
      // Use the core function to save items
      const result = saveItemsCore(items);
      
      if (!result.success) {
        return result; // Forward the error
      }
      
      return {
        success: true,
        message: `Successfully saved ${result.count} items.`,
        itemCount: result.count
      };
    } catch (error) {
      Logger.log("Error in saveItemsFromUI: " + error.message);
      return {
        success: false,
        error: "Error saving items: " + error.message
      };
    }
  }
  
  // Function alias for backward compatibility
  function saveItemsFromDashboard(items) {
    return saveItemsFromUI(items);
  }
  
  /**
   * Generates formatted combined items with type information for autocomplete
   * This improves on the existing implementation to properly format type-item combinations
   * 
   * @param {Array} availableItems - Array of item objects with type and item properties
   * @return {Array} Array of formatted item strings for autocomplete
   */
  function generateCombinedItems(availableItems) {
    try {
      if (!availableItems || !Array.isArray(availableItems)) {
        Logger.log("Invalid input to generateCombinedItems");
        return [];
      }
      
      const combinedItems = [];
      const uniqueItems = new Set();
      
      availableItems.forEach(itemObj => {
        if (itemObj && typeof itemObj === 'object' && itemObj.item) {
          const type = itemObj.type || "";
          const item = itemObj.item.toString().trim();
          
          if (item) {
            // Create a combined string with type info if available
            const combined = type ? `${type} - ${item}` : item;
            
            // Only add if not already in the set (avoid duplicates)
            if (!uniqueItems.has(combined)) {
              uniqueItems.add(combined);
              combinedItems.push(combined);
            }
          }
        }
      });
      
      Logger.log(`Generated ${combinedItems.length} combined items for autocomplete`);
      return combinedItems;
    } catch (e) {
      Logger.log(`Error generating combined items: ${e.toString()}`);
      return [];
    }
  }
  
  /**
   * Gets the combined items for autocomplete from available items
   * 
   * @return {Object} Object with success flag and combined items array
   */
  function getCombinedItems() {
    try {
      // Get available items using the optimized getAvailableItems function
      const availableItemsResult = getAvailableItems();
      
      if (!availableItemsResult.success) {
        Logger.log("Failed to get available items for combined items");
        return { 
          success: false, 
          message: "Failed to retrieve available items", 
          items: [] 
        };
      }
      
      // Use the improved generator function to create combined items
      const combinedItems = generateCombinedItems(availableItemsResult.items);
      
      return {
        success: true,
        message: `Retrieved ${combinedItems.length} combined items`,
        items: combinedItems
      };
    } catch (e) {
      Logger.log(`Error in getCombinedItems: ${e.toString()}`);
      return {
        success: false,
        message: `Error retrieving combined items: ${e.toString()}`,
        items: []
      };
    }
  }
  
  /**
   * Get project summary data for the dashboard
   * Returns budget totals, room count, item count, and project name
   * @param {string} sheetId - Spreadsheet ID to use for getting the data
   * @return {Object} Summary data for the project
   */
  function getProjectSummary(sheetId) {
    try {
      if (!sheetId) {
        return {
          success: false,
          error: "Error: Sheet ID not provided"
        };
      }
      
      // Get spreadsheet by ID and name
      const ss = SpreadsheetApp.openById(sheetId);
      const projectName = ss.getName();
      
      // Initialize summary data
      let summaryData = {
        success: true,
        projectName: projectName,
        totalLowBudget: 0,
        totalHighBudget: 0,
        roomCount: 0,
        itemCount: 0
      };
      
      // Get selected rooms
      let selectedRoomsResult = getSelectedRooms(sheetId);
      if (selectedRoomsResult.success && selectedRoomsResult.selectedRooms) {
        summaryData.roomCount = selectedRoomsResult.selectedRooms.length;
      }
      
      // Get items data
      let itemsSheet = ss.getSheetByName("Items");
      if (itemsSheet) {
        const data = itemsSheet.getDataRange().getValues();
        
        // Skip header row
        if (data.length > 1) {
          summaryData.itemCount = data.length - 1;
          
          // Calculate totals
          // Assuming column format: Room, Type, Item, Quantity, Low Budget, Low Budget Total, High Budget, High Budget Total
          let lowBudgetTotal = 0;
          let highBudgetTotal = 0;
          
          for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const lowBudgetTotalCell = row[5]; // Column F: Low Budget Total
            const highBudgetTotalCell = row[7]; // Column H: High Budget Total
            
            // Add to totals if the values are numbers
            if (typeof lowBudgetTotalCell === 'number') {
              lowBudgetTotal += lowBudgetTotalCell;
            }
            
            if (typeof highBudgetTotalCell === 'number') {
              highBudgetTotal += highBudgetTotalCell;
            }
          }
          
          summaryData.totalLowBudget = lowBudgetTotal;
          summaryData.totalHighBudget = highBudgetTotal;
        }
      }
      
      // Log success
      console.log('Project summary loaded successfully');
      return summaryData;
      
    } catch (error) {
      console.error('Error getting project summary: ' + error);
      return {
        success: false,
        error: 'Failed to load project summary: ' + error.toString()
      };
    }
  } 