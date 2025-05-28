/**
 * Item Management Module
 * Contains all functionality related to managing project items and rooms.
 */

  
  /**
   * Fetches master room data from the Data sheet.
   * 
   * @param {string} sheetId - Optional spreadsheet ID. If not provided, uses DATA_SHEET_ID.
   * @return {Object} Object containing rooms data, header index, and success status
   */
  function getMasterRoomData(sheetId = null) {
    try {
      const id = sheetId || PropertiesService.getScriptProperties().getProperty('DATA_SHEET_ID');
      const ss = SpreadsheetApp.openById(id);
      const dataSheet = ss.getSheetByName("Data");

      if (!dataSheet) {
        return {
          success: false,
          rooms: [],
          headerRowIndex: -1,
          error: "Data sheet not found in the spreadsheet"
        };
      }

      const dataRange = dataSheet.getRange("A:A");
      const values = dataRange.getValues();
      
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
          rooms: [],
          headerRowIndex: -1,
          error: "Rooms header not found in column A of Data sheet"
        };
      }

      const rooms = [];
      for (let i = headerRowIndex + 1; i < values.length; i++) {
        const roomName = values[i][0];
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
      Logger.log("Error in getMasterRoomData: " + error.message);
      return {
        success: false,
        rooms: [],
        headerRowIndex: -1,
        error: "Error retrieving master room data: " + error.message
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
      const roomsResult = getMasterRoomData();
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
      const ss = SpreadsheetApp.openById(ScriptProperties.getProperty('DATA_SHEET_ID'));
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
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
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
      
      // Get raw items data from Master Items List sheet, passing along the sheetId
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
      
      // Get the master data spreadsheet
      const sheetId = PropertiesService.getScriptProperties().getProperty('DATA_SHEET_ID');
      const dataSheet = SpreadsheetApp.openById(sheetId).getSheetByName("Data");
      
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
   * @param {string} sheetId - Optional spreadsheet ID. If not provided, uses active spreadsheet.
   * @return {Object} Object containing success status
   */
  function saveItemSelections(roomItems, sheetId = null) {
    try {
      if (!roomItems || typeof roomItems !== 'object') {
        return {
          success: false,
          error: "Invalid room items data"
        };
      }
      
      // Use provided sheetId if available, otherwise fallback to active spreadsheet
      const ss = sheetId 
        ? SpreadsheetApp.openById(sheetId) 
        : SpreadsheetApp.getActiveSpreadsheet();
      
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
      
      // Get room-category selections
      const roomTypeSelectionsResult = getRoomTypeSelectionsCore(sheetId); // Pass sheetId if necessary
      const roomCategories = roomTypeSelectionsResult.success ? roomTypeSelectionsResult.roomTypes : {};

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
        roomCategories: roomCategories, // Added roomCategories
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
      const ss = sheetId ? SpreadsheetApp.openById(sheetId) : SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('Master Item List');
      
      if (!sheet) {
        Logger.log('Master Item List sheet not found in getItemsData');
        return {
          success: false,
          error: "Sheet 'Master Item List' not found."
        };
      }
      
      const dataRange = sheet.getDataRange();
      const allData = dataRange.getValues();
      
      if (allData.length < 2) {
        Logger.log('Master Item List sheet is empty or has only headers.');
        return {
          success: true,
          items: [],
          itemsByRoom: {},
          selectedRooms: selectedRooms || [] 
        };
      }
      
      const headers = allData[0].map(header => header.toString().trim());
      // Define expected headers for robust mapping
      const expectedHeaders = ["Room", "Type", "Item", "Quantity", "Low Budget", "High Budget", "SPEC/FFE"]; 
      // "ID" header is intentionally NOT included as per plan to use row numbers.
      
      // Create a map of header names to their column indices
      const headerMap = {};
      headers.forEach((header, index) => {
        headerMap[header] = index;
      });
      
      // Validate that all expected headers are present
      for (const eh of expectedHeaders) {
        if (headerMap[eh] === undefined) {
          const errorMsg = `Missing expected header '${eh}' in 'Master Item List'. Please add it and try again.`;
          Logger.log(errorMsg);
          return {
            success: false,
            error: errorMsg
          };
        }
      }
      
      const items = [];
      const itemsByRoom = {};
      
      // Start from row 1 (data starts at index 1, after headers at index 0)
      // The actual sheet row number is `i + 1` because allData is 0-indexed array of rows.
      // And if header is row 1, then data starts at row 2. So `allData[1]` is sheet row 2.
      // Therefore, the rowNumber for `allData[i]` (where i > 0) is `i + 1`.
      for (let i = 1; i < allData.length; i++) {
        const row = allData[i];
        const currentSheetRowNumber = i + 1; // 1-indexed physical row number in the sheet
        
        // Basic validation: ensure the row has enough columns for expected headers
        if (row.length < expectedHeaders.length) {
          // Logger.log(`Row ${currentSheetRowNumber} has insufficient columns. Skipping.`);
          // continue; // Skip malformed rows
        }

        const roomName = row[headerMap["Room"]] ? row[headerMap["Room"]].toString().trim() : 'Unassigned';
        
        // If selectedRooms are provided, only process items for those rooms
        if (selectedRooms && selectedRooms.length > 0 && !selectedRooms.includes(roomName)) {
          continue;
        }
        
        const item = {
          // id: row[headerMap["ID"]] ? row[headerMap["ID"]].toString().trim() : `temp_${Date.now()}_${i}`, // Using temp ID for now, will be replaced by server on save if needed
          // Temporary client-side ID if no actual ID column exists. Will be replaced by rowNumber strategy.
          // For new items added on client, they will get a 'new_...' id.
          // For existing items, their true identifier will be 'rowNumber'.
          id: `row_${currentSheetRowNumber}`, // Temporary unique ID for client-side processing until full refactor
          rowNumber: currentSheetRowNumber, // Key addition: 1-indexed physical row number
          room: roomName,
          type: row[headerMap["Type"]] ? row[headerMap["Type"]].toString().trim() : "",
          item: row[headerMap["Item"]] ? row[headerMap["Item"]].toString().trim() : "",
          quantity: parseInt(row[headerMap["Quantity"]]) || 1,
          lowBudget: parseFloat(row[headerMap["Low Budget"]]) || null,
          highBudget: parseFloat(row[headerMap["High Budget"]]) || null,
          specFfe: row[headerMap["SPEC/FFE"]] ? row[headerMap["SPEC/FFE"]].toString().trim() : "",
          // Calculate totals - these might be better calculated on client or based on need
          lowBudgetTotal: null,
          highBudgetTotal: null
        };
        
        // Basic validation: ensure item name is present
        if (!item.item) {
          // Logger.log(`Row ${currentSheetRowNumber} is missing an item name. Skipping.`);
          // continue;
        }
        
        // Calculate total budgets if individual budgets and quantity are present
        if (item.quantity && item.lowBudget !== null) {
          item.lowBudgetTotal = item.quantity * item.lowBudget;
        }
        if (item.quantity && item.highBudget !== null) {
          item.highBudgetTotal = item.quantity * item.highBudget;
        }
        
        items.push(item);
        
        if (!itemsByRoom[item.room]) {
          itemsByRoom[item.room] = [];
        }
        itemsByRoom[item.room].push(item);
      }
      
      return {
        success: true,
        items: items,
        itemsByRoom: itemsByRoom,
        selectedRooms: selectedRooms || Object.keys(itemsByRoom) // If no selectedRooms passed, return all rooms found
      };
      
    } catch (e) {
      Logger.log('Error in getItemsData: ' + e.toString() + ' Stack: ' + e.stack);
      return {
        success: false,
        error: "An error occurred while fetching item data: " + e.message
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
        : SpreadsheetApp.openById(ScriptProperties.getProperty('DATA_SHEET_ID'));
      
      const dataSheet = ss.getSheetByName("Data");

      console.log(dataSheet);
      
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
      let dataSheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
        
      const masterRoomResult = getMasterRoomData(dataSheetId); // Changed from getRoomNamesFromSheet
      let rooms = [];
      if (masterRoomResult.success) {
        rooms = masterRoomResult.rooms;
      } else {
        // Log the error or handle it as needed if master rooms can't be fetched
        Logger.log("Error fetching master rooms for dashboard: " + (masterRoomResult.error || 'Unknown error'));
      }
      
      // Get the currently selected rooms from temp sheet using core function
      const selectedRooms = getSelectedRoomsCore(sheetId);

      console.log(selectedRooms);
      
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
  function saveItemsToMasterList(itemsToSave, sheetId = null) {
    let backupSheetName = null;
    const processedItemsWithRowNumbers = []; // To store items with their final row numbers

    try {
      Logger.log(`Starting ROW-NUMBER BASED save of ${itemsToSave ? itemsToSave.length : 0} items. Received sheetId: ${sheetId}`);

      const ss = sheetId ? SpreadsheetApp.openById(sheetId) : SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        Logger.log(`Spreadsheet object is null. sheetId provided was: ${sheetId}`);
        return { success: false, error: "Spreadsheet not found (object was null)." };
      }

      const masterSheetName = "Master Item List";
      const masterSheet = ss.getSheetByName(masterSheetName);
      if (!masterSheet) {
        Logger.log(`Sheet "${masterSheetName}" not found in spreadsheet ID: ${ss.getId()}.`);
        return { success: false, error: `Sheet "${masterSheetName}" not found.` };
      }

      // --- 1. Safety Backup --- (Keep existing backup logic)
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
      backupSheetName = `${masterSheetName}_Backup_${timestamp}`;
      try {
        const existingBackup = ss.getSheetByName(backupSheetName);
        if (existingBackup) ss.deleteSheet(existingBackup);
        masterSheet.copyTo(ss).setName(backupSheetName).hideSheet();
        Logger.log(`Successfully created backup: ${backupSheetName}`);
      } catch (e) {
        Logger.log(`Error creating backup sheet: ${e.message}`);
        backupSheetName = null;
      }

      // --- 2. Initial Validation & Preparation of Incoming Items ---
      if (!itemsToSave || !Array.isArray(itemsToSave)) itemsToSave = [];

      const itemsToUpdateInPlace = [];
      const itemsToAppend = [];
      const clientSentInvalidItems = [];

      itemsToSave.forEach((item, index) => {
        if (!item || typeof item !== 'object') {
          clientSentInvalidItems.push({ index, item, reason: "Item is not an object" });
          return;
        }
        if (!item.room || String(item.room).trim() === "" || !item.item || String(item.item).trim() === "") {
          clientSentInvalidItems.push({ index, item, reason: "Missing room or item name" });
          return;
        }

        // Common validation for all items
        const validatedItem = {
          room: String(item.room).trim().toUpperCase(),
          type: item.type ? String(item.type).trim().toUpperCase() : "",
          item: String(item.item).trim().toUpperCase(),
          quantity: item.quantity !== undefined && item.quantity !== null ? Math.max(1, parseInt(item.quantity) || 1) : 1,
          lowBudget: (item.lowBudget !== undefined && item.lowBudget !== null && String(item.lowBudget).trim() !== "" && !isNaN(parseFloat(item.lowBudget))) ? parseFloat(item.lowBudget) : null,
          highBudget: (item.highBudget !== undefined && item.highBudget !== null && String(item.highBudget).trim() !== "" && !isNaN(parseFloat(item.highBudget))) ? parseFloat(item.highBudget) : null,
          specFfe: item.specFfe ? String(item.specFfe).trim().toUpperCase() : "",
          originalTemporaryId: item.id && String(item.id).startsWith('new_') ? item.id : null // Capture client's temporary ID
        };
        validatedItem.lowBudgetTotal = validatedItem.lowBudget !== null ? validatedItem.lowBudget * validatedItem.quantity : null;
        validatedItem.highBudgetTotal = validatedItem.highBudget !== null ? validatedItem.highBudget * validatedItem.quantity : null;

        if (item.rowNumber && Number.isInteger(item.rowNumber) && item.rowNumber > 0) {
          validatedItem.rowNumber = item.rowNumber;
          itemsToUpdateInPlace.push(validatedItem);
        } else {
          itemsToAppend.push(validatedItem); // Will get rowNumber after append
        }
      });

      if (clientSentInvalidItems.length > 0) {
        Logger.log(`Validation found ${clientSentInvalidItems.length} invalid items from client. Aborting save.`);
        return { success: false, error: `${clientSentInvalidItems.length} invalid items received. Save aborted.`, invalidItems: clientSentInvalidItems, backupSheetName };
      }
      Logger.log(`Separated items: ${itemsToUpdateInPlace.length} to update, ${itemsToAppend.length} to append.`);

      // --- 3. Get Sheet Headers and Data ---
      const headerRowValues = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
      const headerMap = {};
      headerRowValues.forEach((header, index) => headerMap[header.toString().trim()] = index);
      
      const expectedHeaderOrder = ["ROOM", "TYPE", "ITEM", "QUANTITY", "LOW BUDGET", "HIGH BUDGET", "LOW BUDGET TOTAL", "HIGH BUDGET TOTAL", "SPEC/FFE"];
      // Validate headerMap contains all expectedHeaderOrder keys, crucial for writing data correctly.
      for (const h of expectedHeaderOrder) {
        if (headerMap[h] === undefined) {
          const errorMsg = `Master Item List is missing critical header: '${h}'. Cannot proceed.`;
          Logger.log(errorMsg);
          return { success: false, error: errorMsg, backupSheetName };
        }
      }

      let allSheetData = [];
      const lastRow = masterSheet.getLastRow();
      const lastCol = masterSheet.getLastColumn();
      if (lastRow > 1) { // Only read if there's data beyond headers
        allSheetData = masterSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      }

      // --- 4. Process In-Place Updates ---
      itemsToUpdateInPlace.forEach(item => {
        const rowIndexInSheetData = item.rowNumber - 2; // -1 for 0-indexed, -1 because allSheetData excludes header
        if (rowIndexInSheetData >= 0 && rowIndexInSheetData < allSheetData.length) {
          const sheetRowArray = allSheetData[rowIndexInSheetData];
          // Update the sheetRowArray based on item properties and headerMap
          sheetRowArray[headerMap["ROOM"]] = item.room;
          sheetRowArray[headerMap["TYPE"]] = item.type;
          sheetRowArray[headerMap["ITEM"]] = item.item;
          sheetRowArray[headerMap["QUANTITY"]] = item.quantity;
          sheetRowArray[headerMap["LOW BUDGET"]] = item.lowBudget;
          sheetRowArray[headerMap["HIGH BUDGET"]] = item.highBudget;
          sheetRowArray[headerMap["LOW BUDGET TOTAL"]] = item.lowBudgetTotal;
          sheetRowArray[headerMap["HIGH BUDGET TOTAL"]] = item.highBudgetTotal;
          sheetRowArray[headerMap["SPEC/FFE"]] = item.specFfe;
          // Add to processed list, rowNumber is known
          processedItemsWithRowNumbers.push(item);
        } else {
          Logger.log(`Warning: Item designated for update at row ${item.rowNumber} is out of sheet bounds (max data row ${lastRow -1}). Will attempt to append instead.`);
          item.rowNumber = null; // Clear invalid row number
          itemsToAppend.push(item); // Re-route to append
        }
      });
      Logger.log(`${itemsToUpdateInPlace.length - itemsToAppend.filter(i => i.rowNumber === null).length} items prepared for in-place update in memory.`);

      // --- 5. Write Updated Data (if any changes made to existing rows) ---
      if (itemsToUpdateInPlace.length > 0 && allSheetData.length > 0) {
        masterSheet.getRange(2, 1, allSheetData.length, lastCol).setValues(allSheetData);
        Logger.log('Successfully wrote back updated existing rows data.');
      }

      // --- 6. Process Appends ---
      if (itemsToAppend.length > 0) {
        const newRowsDataArray = itemsToAppend.map(item => {
          const newRow = new Array(expectedHeaderOrder.length).fill(null);
          newRow[headerMap["ROOM"]] = item.room;
          newRow[headerMap["TYPE"]] = item.type;
          newRow[headerMap["ITEM"]] = item.item;
          newRow[headerMap["QUANTITY"]] = item.quantity;
          newRow[headerMap["LOW BUDGET"]] = item.lowBudget;
          newRow[headerMap["HIGH BUDGET"]] = item.highBudget;
          newRow[headerMap["LOW BUDGET TOTAL"]] = item.lowBudgetTotal;
          newRow[headerMap["HIGH BUDGET TOTAL"]] = item.highBudgetTotal;
          newRow[headerMap["SPEC/FFE"]] = item.specFfe;
          return newRow;
        });

        const appendStartRow = masterSheet.getLastRow() + 1;
        masterSheet.getRange(appendStartRow, 1, newRowsDataArray.length, expectedHeaderOrder.length).setValues(newRowsDataArray);
        Logger.log(`Appended ${newRowsDataArray.length} new rows starting at row ${appendStartRow}.`);

        // Assign rowNumbers to appended items and add to processed list
        itemsToAppend.forEach((item, index) => {
          item.rowNumber = appendStartRow + index;
          processedItemsWithRowNumbers.push(item);
        });
      }
      
      // --- 7. Apply SPEC/FFE Data Validation to all data rows ---
      const finalLastDataRow = masterSheet.getLastRow();
      if (finalLastDataRow > 1) { // If there are any data rows (header is row 1)
          const specFfeColIndex = headerMap["SPEC/FFE"];
          if (specFfeColIndex !== undefined) {
              const specFfeSheetCol = specFfeColIndex + 1; // 1-indexed column
              const numDataRows = finalLastDataRow - 1;
              const specFfeRange = masterSheet.getRange(2, specFfeSheetCol, numDataRows, 1);
              const rule = SpreadsheetApp.newDataValidation()
                                       .requireValueInList(['SPEC', 'FFE', ''], true) // Allow blank
                                       .setAllowInvalid(false)
                                       .build();
              specFfeRange.setDataValidation(rule);
              Logger.log(`Applied SPEC/FFE dropdown validation to column ${specFfeSheetCol}, for ${numDataRows} data rows.`);
          } else {
              Logger.log("SPEC/FFE header not found, skipping data validation for it.");
          }
      } else {
          Logger.log("No data rows found after save, skipping SPEC/FFE validation.");
      }

      // --- 8. Sort processedItemsWithRowNumbers by their final rowNumber for consistent return ---
      processedItemsWithRowNumbers.sort((a, b) => (a.rowNumber || 0) - (b.rowNumber || 0));
      
      Logger.log(`Save successful using row numbers. Processed: ${processedItemsWithRowNumbers.length} items.`);
      return {
        success: true,
        items: processedItemsWithRowNumbers, // Return all items with their final row numbers
        count: processedItemsWithRowNumbers.length,
        backupSheetName: backupSheetName
      };
      
    } catch (error) {
      Logger.log(`Critical Error in saveItemsToMasterList (Row Number Based): ${error.message} Stack: ${error.stack}`);
      return {
        success: false,
        error: `Error saving items: ${error.message}`,
        backupSheetName: backupSheetName
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
            const combined = type ? `${type} : ${item}` : item;
            
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
  
  /**
   * Fetches all types from the Data sheet.
   * 
   * @return {Object} Object containing types data and success status
   */
  function getTypes() {
    try {
      const ss = SpreadsheetApp.openById(ScriptProperties.getProperty('DATA_SHEET_ID'));
      const dataSheet = ss.getSheetByName("Data");
      
      if (!dataSheet) {
        return {
          success: false,
          error: "Data sheet not found in the spreadsheet"
        };
      }
      
      // Find the column containing Types
      const headerRow = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
      const typeColIndex = headerRow.indexOf("Type");
      
      if (typeColIndex === -1) {
        return {
          success: false,
          error: "Type header not found in Data sheet"
        };
      }
      
      // Get the range containing types in the Type column
      const dataRange = dataSheet.getRange(2, typeColIndex + 1, dataSheet.getLastRow() - 1, 1);
      const values = dataRange.getValues();
      
      // Extract unique types (skip empty cells)
      const typesSet = new Set();
      values.forEach(row => {
        if (row[0] && row[0].trim() !== "") {
          typesSet.add(row[0]);
        }
      });
      
          // Convert to array and sort alphabetically
    const types = Array.from(typesSet).sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
    
    Logger.log(`Found ${types.length} types in the Data sheet (sorted alphabetically)`);
      return {
        success: true,
        types: types
      };
      
    } catch (error) {
      Logger.log("Error in getTypes: " + error.message);
      return {
        success: false,
        error: "Error retrieving types: " + error.message
      };
    }
  }
  
  /**
   * Get room-type selections from the temporary sheet.
   * 
   * @param {string} sheetId - Optional spreadsheet ID
   * @return {Object} Object containing room-type selections
   */
  function getRoomTypeSelectionsCore(sheetId) {
    try {
      // Use provided sheetId if available, otherwise fallback to active spreadsheet
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      const tempSheet = ss.getSheetByName("_TempRoomTypes");
      let roomTypes = {};
      
      if (tempSheet) {
        // Get room-type selections from the temporary sheet (skip header row)
        const lastRow = tempSheet.getLastRow();
        if (lastRow > 1) {
          const dataRange = tempSheet.getRange(2, 1, lastRow - 1, 2);
          const dataValues = dataRange.getValues();
          
          // Organize room-type selections
          dataValues.forEach(row => {
            const room = row[0];
            const type = row[1];
            
            if (room && type) {
              if (!roomTypes[room]) {
                roomTypes[room] = [];
              }
              roomTypes[room].push(type);
            }
          });
          
          Logger.log(`Found room-type selections for ${Object.keys(roomTypes).length} rooms`);
        }
      }
      
      return {
        success: true,
        roomTypes: roomTypes
      };
    } catch (error) {
      Logger.log("Error in getRoomTypeSelectionsCore: " + error.message);
      return {
        success: false,
        error: "Error retrieving room-type selections: " + error.message
      };
    }
  }
  
  /**
   * Saves room-type selections to the temporary sheet.
   * 
   * @param {Object} roomTypes - Object with room names as keys and arrays of types as values
   * @param {string} sheetId - Optional spreadsheet ID
   * @return {Object} Success status
   */
  function saveRoomTypeSelections(roomTypes, sheetId) {
    try {
      if (!roomTypes || typeof roomTypes !== 'object') {
        return {
          success: false,
          error: "Invalid room-type data: Input was null or not an object."
        };
      }
      // Use provided sheetId if available, otherwise fallback to active spreadsheet
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      // Create or get the temporary sheet
      let tempSheet = ss.getSheetByName("_TempRoomTypes");
      if (tempSheet) {
        // Clear existing content if sheet exists
        tempSheet.clear();
      } else {
        // Create the sheet if it doesn't exist
        tempSheet = ss.insertSheet("_TempRoomTypes");
        // Hide the sheet as it's for temporary storage only
        tempSheet.hideSheet();
      }
      // Add header row
      tempSheet.getRange(1, 1, 1, 2).setValues([["Room", "Type"]]);
      // Prepare the data to write
      const flatData = [];
      Object.keys(roomTypes).forEach(room => {
        const types = roomTypes[room];
        if (Array.isArray(types) && types.length > 0) {
          types.forEach(type => {
            if (type && String(type).trim() !== "") { // Ensure type is not null/empty
              flatData.push([room, String(type).trim()]);
            }
          });
        }
      });
      // Check if there is any valid data to save
      if (flatData.length === 0) {
        Logger.log("No valid room-type selections to save after processing.");
        // It might be preferable to still clear the sheet and return success,
        // or return an error/specific status if no data means an issue.
        // For now, let's clear and return success as it reflects the (empty) state.
        // tempSheet.getRange(2, 1, Math.max(1, tempSheet.getLastRow() -1), 2).clearContent(); // Clear if needed
        return {
          success: true, // Or false, depending on desired behavior for empty valid save
          message: "No room-type selections to save.",
          count: 0
        };
      }
      // Write the data to the sheet
      if (flatData.length > 0) {
        tempSheet.getRange(2, 1, flatData.length, 2).setValues(flatData);
      }
      Logger.log(`Saved ${flatData.length} room-type selections to temporary sheet`);
      return {
        success: true,
        count: flatData.length
      };
    } catch (error) {
      Logger.log("Error in saveRoomTypeSelections: " + error.message);
      return {
        success: false,
        error: "Error saving room-type selections: " + error.message
      };
    }
  }
  
  /**
   * Retrieves room and type data for the Project_Details_ sidebar.
   * 
   * @return {Object} Data for rendering room categories section
   */
  function getRoomCategoriesData() {
    try {
      // Get selected rooms
      const selectedRooms = getSelectedRoomsCore();
      
      // Get available types
      const typesResult = getTypes();
      const types = typesResult.success ? typesResult.types : [];
      
      // Get existing room-type selections
      const roomTypeSelectionsResult = getRoomTypeSelectionsCore();
      const roomTypes = roomTypeSelectionsResult.success ? roomTypeSelectionsResult.roomTypes : {};
      
      return {
        success: true,
        selectedRooms: selectedRooms,
        availableTypes: types,
        roomTypes: roomTypes
      };
    } catch (error) {
      Logger.log("Error in getRoomCategoriesData: " + error.message);
      return {
        success: false,
        error: "Error getting room categories data: " + error.message
      };
    }
  }
  
  /**
   * Updates a single room-category assignment in the _TempRoomTypes sheet.
   * This function now reads and writes data in a flat row-based format,
   * consistent with getRoomTypeSelectionsCore and saveRoomTypeSelections.
   *
   * @param {string} sheetId The ID of the spreadsheet.
   * @param {string} roomName The name of the room.
   * @param {string} categoryName The name of the category.
   * @param {boolean} isSelected True if the category should be assigned, false to unassign.
   * @return {Object} Object with success status and optional error message.
   */
  function updateRoomCategoryAssignment(sheetId, roomName, categoryName, isSelected) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let tempSheet = ss.getSheetByName("_TempRoomTypes");

      if (!tempSheet) {
        tempSheet = ss.insertSheet("_TempRoomTypes");
        tempSheet.hideSheet();
        tempSheet.getRange(1, 1, 1, 2).setValues([["Room", "Type"]]); // Add header
      }

      // 1. Read existing flat data and reconstruct roomTypes object
      const roomTypes = {};
      const lastRow = tempSheet.getLastRow();
      if (lastRow > 1) { // Check if there's data beyond the header
        const dataRange = tempSheet.getRange(2, 1, lastRow - 1, 2);
        const values = dataRange.getValues();
        values.forEach(row => {
          const rName = row[0];
          const cName = row[1];
          if (rName && cName) {
            if (!roomTypes[rName]) {
              roomTypes[rName] = [];
            }
            if (!roomTypes[rName].includes(cName)) { // Ensure no duplicates if sheet had them
                roomTypes[rName].push(cName);
            }
          }
        });
      }

      // 2. Modify the in-memory roomTypes object
      if (!roomTypes[roomName]) {
        roomTypes[roomName] = [];
      }
      const categoryIndex = roomTypes[roomName].indexOf(categoryName);

      if (isSelected) {
        if (categoryIndex === -1) {
          roomTypes[roomName].push(categoryName);
        }
      } else {
        if (categoryIndex > -1) {
          roomTypes[roomName].splice(categoryIndex, 1);
        }
      }
      
      // If a room has no categories after modification, remove the room key
      if (roomTypes[roomName] && roomTypes[roomName].length === 0) {
          delete roomTypes[roomName];
      }

      // 3. Clear existing data rows (below header)
      if (lastRow > 1) {
        tempSheet.getRange(2, 1, lastRow - 1, 2).clearContent();
      }

      // 4. Write the updated roomTypes object back as flat rows
      const flatData = [];
      Object.keys(roomTypes).forEach(rName => {
        const categories = roomTypes[rName];
        if (Array.isArray(categories)) {
          categories.forEach(cName => {
            flatData.push([rName, cName]);
          });
        }
      });

      if (flatData.length > 0) {
        tempSheet.getRange(2, 1, flatData.length, 2).setValues(flatData);
      }
      
      Logger.log(`Updated room-category assignment for Room: ${roomName}, Category: ${categoryName}, Selected: ${isSelected}. Wrote ${flatData.length} rows.`);
      return { success: true };

    } catch (error) {
      Logger.log("Error in updateRoomCategoryAssignment: " + error.message + " (Room: " + roomName + ", Category: " + categoryName + ", Selected: " + isSelected + ")");
      return { success: false, error: "Error updating room category assignment: " + error.message };
    }
  }

  /**
   * Get types for a specific room.
   * Used for filtering available items by type.
   * 
   * @param {string} room - Room name
   * @return {Array} Array of type names for the room
   */
  function getTypesForRoom(room) {
    try {
      const roomTypesResult = getRoomTypeSelectionsCore();
      
      if (!roomTypesResult.success || !roomTypesResult.roomTypes[room]) {
        return [];
      }
      
      return roomTypesResult.roomTypes[room];
    } catch (error) {
      Logger.log("Error in getTypesForRoom: " + error.message);
      return [];
    }
  }
  
  /**
   * Retrieves item data for the item selection interface based on selected room categories.
   * This function prepares category-filtered items for the item selection interface.
   *
   * @param {string} sheetId - Optional spreadsheet ID. If not provided, uses active spreadsheet.
   * @param {Array} targetRooms - Optional array of specific rooms to retrieve data for
   * @return {Object} Object containing item data organized by room and category
   */
  function getItemSelectionData(sheetId = null, targetRooms = null) {
    try {
      // Use provided sheetId if available, otherwise fallback to active spreadsheet
      let ss;
      if (sheetId) {
        ss = SpreadsheetApp.openById(sheetId);
      } else {
        ss = SpreadsheetApp.getActiveSpreadsheet();
      }

      // Get selected rooms if not provided
      let selectedRooms = targetRooms;
      if (!selectedRooms || !Array.isArray(selectedRooms) || selectedRooms.length === 0) {
        selectedRooms = getSelectedRoomsCore(sheetId);
      }

      // If no rooms are selected, return empty result
      if (selectedRooms.length === 0) {
        return {
          success: false,
          error: "No rooms selected. Please select rooms first."
        };
      }

      // Get the room-type mappings
      const roomCategoriesData = getRoomTypeSelectionsCore(sheetId);
      const roomTypesMap = roomCategoriesData.success ? roomCategoriesData.roomTypes : {};
      
      // Get all available types/categories
      const typesData = getTypes();
      const availableTypes = typesData.success ? typesData.types : [];
      
      // Get all available items from the Data sheet
      const dataSheet = ss.getSheetByName("Data");
      if (!dataSheet) {
        return {
          success: false,
          error: "Data sheet not found in the spreadsheet"
        };
      }
      
      // Get data from the Data sheet
      const dataRange = dataSheet.getDataRange();
      const values = dataRange.getValues();
      
      // Headers are in the first row
      const headerRow = values[0];
      let itemNameColIndex = -1;
      let itemTypeColIndex = -1;
      
      // Find the column indices for "Item-Name" and "Item-Type"
      for (let j = 0; j < headerRow.length; j++) {
        if (headerRow[j] === "Item-Name") {
          itemNameColIndex = j;
        } else if (headerRow[j] === "Item-Type") {
          itemTypeColIndex = j;
        }
      }
      
      if (itemNameColIndex === -1 || itemTypeColIndex === -1) {
        return {
          success: false,
          error: "Required columns 'Item-Name' and 'Item-Type' not found in the Data sheet"
        };
      }
      
      // Extract items data (skipping the header row)
      const allItemsWithTypes = [];
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        const itemName = row[itemNameColIndex];
        const itemType = row[itemTypeColIndex];
        
        // Stop if we hit an empty item name
        if (!itemName) {
          break;
        }
        
        allItemsWithTypes.push({
          type: itemType || '',
          item: itemName
        });
      }
      
      // Organize items by type
      const itemsByType = {};
      allItemsWithTypes.forEach(item => {
        const type = item.type.trim().toUpperCase() || 'UNCATEGORIZED';
        if (!itemsByType[type]) {
          itemsByType[type] = [];
        }
        itemsByType[type].push(item);
      });
      
      // For each selected room, find the assigned types and the corresponding items
      const itemsByRoom = {};
      const combinedItems = [];
      const allItems = [];
      
      selectedRooms.forEach(room => {
        // Get types assigned to this room
        const roomTypes = roomTypesMap[room] || [];
        
        // Initialize items array for this room
        itemsByRoom[room] = [];
        
        // For each type assigned to the room, get items of that type
        roomTypes.forEach(type => {
          let itemsOfType = itemsByType[type.toUpperCase()] || [];
          // Sort items alphabetically by item name (case-insensitive)
          itemsOfType = itemsOfType.slice().sort((a, b) =>
            a.item.toLowerCase().localeCompare(b.item.toLowerCase())
          );

          // Add each item to the room's items
          itemsOfType.forEach(item => {
            const itemForRoom = {
              room: room,
              type: item.type.toUpperCase(),
              item: item.item.toUpperCase(),
              quantity: 1,
              isSelected: false
            };
            
            itemsByRoom[room].push(itemForRoom);
            allItems.push(itemForRoom);
            
            // Add to combined items for autocomplete
            const combinedItem = `${item.type} : ${item.item}`;
            if (!combinedItems.includes(combinedItem)) {
              combinedItems.push(combinedItem);
            }
          });
        });
      });
      
      Logger.log(`Prepared item selection data with ${allItems.length} items across ${selectedRooms.length} rooms`);
      
      return {
        success: true,
        items: allItems,
        itemsByRoom: itemsByRoom,
        selectedRooms: selectedRooms,
        combinedItems: combinedItems,
        availableItems: allItemsWithTypes
      };
      
    } catch (error) {
      Logger.log("Error in getItemSelectionData: " + error.message);
      return {
        success: false,
        error: "Error retrieving item selection data: " + error.message
      };
    }
  }
  
  // HELPER FUNCTION FOR TEMPORARY ITEM DATA
  /**
   * Gets or creates the temporary sheet for storing incomplete item data.
   * @return {Sheet | null} The temporary data sheet or null if an error occurs.
   * @private
   */
  function _getTempItemDataSheet() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetName = "_TempItemData";
      let sheet = ss.getSheetByName(sheetName);

      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.hideSheet();
        // Set up headers if needed, e.g., ['Timestamp', 'ItemDataJSON']
        sheet.getRange(1, 1, 1, 2).setValues([['Timestamp', 'ItemDataJSON']]);
        Logger.log(`Created and hid temporary sheet: ${sheetName}`);
      }
      return sheet;
    } catch (e) {
      Logger.log(`Error in _getTempItemDataSheet: ${e.message}`);
      return null;
    }
  }

  /**
   * Saves partially entered item data to a temporary sheet.
   * For simplicity, this version overwrites any existing temporary data.
   * A more advanced version could handle multiple drafts or user-specific drafts.
   * @param {Object} itemData The item data object to save.
   * @return {Object} Object with success status and optional error message.
   */
  function saveTemporaryItemData(itemData) {
    try {
      if (!itemData || Object.keys(itemData).length === 0) {
        return { success: true, message: "No data to save." }; // Not an error, just nothing to do
      }

      const sheet = _getTempItemDataSheet();
      if (!sheet) {
        return { success: false, error: "Could not access temporary data sheet." };
      }

      // Clear previous data (assuming one draft slot for now, e.g., row 2)
      // If sheet has more than 1 row (header), clear row 2. Otherwise, it's empty or just header.
      if (sheet.getLastRow() > 1) {
          sheet.getRange(2, 1, 1, sheet.getLastColumn()).clearContent();
      }
      
      const jsonData = JSON.stringify(itemData);
      sheet.getRange(2, 1).setValue(new Date()); // Timestamp
      sheet.getRange(2, 2).setValue(jsonData);   // Item Data as JSON

      Logger.log("Saved temporary item data.");
      return { success: true };
    } catch (e) {
      Logger.log(`Error in saveTemporaryItemData: ${e.toString()}`);
      return { success: false, error: e.toString() };
    }
  }

  /**
   * Loads temporarily stored item data.
   * @return {Object} Object with success status, data (if found), and optional error message.
   */
  function loadTemporaryItemData() {
    try {
      const sheet = _getTempItemDataSheet();
      if (!sheet) {
        // If the sheet doesn't exist, it means no data was ever saved.
        return { success: true, data: null, message: "Temporary data sheet not found." };
      }

      // Assuming data is in row 2, column 2
      if (sheet.getLastRow() < 2) {
        return { success: true, data: null, message: "No temporary data found." }; // No data rows
      }

      const jsonData = sheet.getRange(2, 2).getValue();

      if (!jsonData) {
        return { success: true, data: null, message: "No temporary item data found." };
      }

      const itemData = JSON.parse(jsonData);
      Logger.log("Loaded temporary item data.");
      return { success: true, data: itemData };
    } catch (e) {
      Logger.log(`Error in loadTemporaryItemData: ${e.toString()}`);
      // If there's an error (e.g., parsing), it's better to return null data
      // than to break the client-side.
      return { success: false, data: null, error: e.toString() };
    }
  }

  /**
   * Clears any temporarily stored item data.
   * @return {Object} Object with success status and optional error message.
   */
  function clearTemporaryItemData() {
    try {
      const sheet = _getTempItemDataSheet(); // This will get or create it.
      if (!sheet) {
        // If sheet couldn't be accessed even to clear, log it but don't necessarily fail hard client-side.
        Logger.log("Could not access temporary data sheet to clear, but proceeding as if cleared.");
        return { success: true, message: "Temp sheet not accessible, assumed clear." };
      }
      
      // Clear data from row 2 (timestamp and JSON)
      // Check if there's anything to clear beyond the header
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, 1, sheet.getLastColumn()).clearContent(); // Clear the second row
        Logger.log("Cleared temporary item data.");
      } else {
        Logger.log("No temporary item data to clear.");
      }
      return { success: true };
    } catch (e) {
      Logger.log(`Error in clearTemporaryItemData: ${e.toString()}`);
      return { success: false, error: e.toString() };
    }
  }