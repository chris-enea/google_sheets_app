// Wrapper function for default dashboard view
function showProjectDashboardDefault() {
  // Pass null or your default tab ID, e.g., 'homeView' or 'dashboardHome'
  // The receiving template needs to handle a null or undefined initialTabToOpen gracefully.
  showProjectDashboard(null); 
}

// Wrapper function for Room Manager view
function showProjectDashboardRoomManager() {
  showProjectDashboard('rooms'); // MODIFIED: Pass the actual data-tab ID 'rooms'
}

  // New wrapper function for Room Categories view
  function showProjectDashboardRoomCategories() {
    showProjectDashboard('roomCategories'); 
  }

  // New wrapper function for Items view
  function showProjectDashboardItems() {
    showProjectDashboard('items'); 
  }

/**
 * Shows the Project Details modal, optionally opening to a specific tab.
 * @param {string} [initialTabId] - Optional ID of the tab to open initially.
 */
function showProjectDashboard(initialTabId) {
  var dataSheetId = PropertiesService.getScriptProperties().getProperty('DATA_SHEET_ID');
  var template = HtmlService.createTemplateFromFile('Project_Details_');
  template.dataSheetId = dataSheetId; // Pass to template
  
  if (initialTabId) {
    template.initialTabToOpen = initialTabId;
  } else {
    template.initialTabToOpen = null; // Ensure it's explicitly null for the scriptlet
  }
  Logger.log('[ui.js] showProjectDashboard - Setting initialTabToOpen to: ' + template.initialTabToOpen);

  var htmlOutput = template.evaluate()
    .setWidth(1500)
    .setHeight(1000);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Project Details');
}

/**
 * Gets dashboard data using script properties
 */
function getDashboardData(folderId, projectName) {
    try {
      
      if (!folderId) {
        return { 
          files: null, 
          projectName,
          error: null 
        };
      }
  
      try {
        const parentFolder = DriveApp.getFolderById(folderId);
        const subfolders = parentFolder.getFolders();
  
        const data = {};
        while (subfolders.hasNext()) {
          const folder = subfolders.next();
          const files = folder.getFiles();
          const fileList = [];
  
          while (files.hasNext()) {
            const file = files.next();
            fileList.push({
              name: file.getName(),
              url: file.getUrl(),
              icon: getFileIcon(file.getMimeType())
            });
          }
  
          if (fileList.length > 0) {
            data[folder.getName()] = fileList;
          }
        }
  
        return {
          projectName,
          files: Object.keys(data).length > 0 ? data : null,
          error: null
        };
      } catch (folderError) {
        Logger.log(`Error accessing folder: ${folderError.message}`);
        
        // Check for permissions issues
        if (folderError.message.includes('PERMISSION_DENIED')) {
          return { 
            files: null, 
            projectName,
            error: "Permission denied. This script doesn't have access to the folder you specified. Please check that the folder ID is correct and that you've granted the necessary permissions."
          };
        } else if (folderError.message.includes('not found')) {
          return { 
            files: null, 
            projectName,
            error: "Folder not found. Please check that the folder ID is correct."
          };
        } else {
          return { 
            files: null, 
            projectName,
            error: `Error accessing folder: ${folderError.message}`
          };
        }
      }
    } catch (error) {
      Logger.log(`Error in getDashboardData: ${error.message}`);
      return { 
        files: null, 
        projectName: getProjectName(),
        error: `Error loading dashboard data: ${error.message}` 
      };
    }
  }

  /**
 * Gathers data for selected items and opens the sidebar.
 */
function openEmailSidebar() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

    if (!sheet) {
      throw new Error(`Sheet named "${CONFIG.SHEET_NAME}" not found.`);
    }

    const headerRowValues = sheet.getRange(CONFIG.HEADER_ROWS, 1, 1, sheet.getLastColumn()).getValues()[0];

    function getColumnIndex(columnName, headerArray) {
      const index = headerArray.indexOf(columnName);
      if (index === -1) {
        throw new Error(`Column "${columnName}" not found in header row. Please check sheet headers and CONFIG settings.`);
      }
      return index; // Returns 0-based index
    }

    const checkboxColIdx = getColumnIndex(CONFIG.CHECKBOX_COL_NAME, headerRowValues);
    const descriptionColIdx = getColumnIndex(CONFIG.DESCRIPTION_COL_NAME, headerRowValues);
    // SKU, Manufacturer, and Dimensions might not always be present, handle gracefully if name is not in CONFIG or not found
    const skuNumberColIdx = CONFIG.SKU_NUMBER_COL_NAME ? getColumnIndex(CONFIG.SKU_NUMBER_COL_NAME, headerRowValues) : -1;
    const manufacturerColIdx = CONFIG.MANUFACTURER_COL_NAME ? getColumnIndex(CONFIG.MANUFACTURER_COL_NAME, headerRowValues) : -1;
    const dimensionsColIdx = CONFIG.DIMENSIONS_COL_NAME ? getColumnIndex(CONFIG.DIMENSIONS_COL_NAME, headerRowValues) : -1;


    const dataRange = sheet.getDataRange();
    const allValues = dataRange.getValues();

    let itemsToEmail = [];
    let rowsToUpdate = [];

    for (let i = CONFIG.HEADER_ROWS; i < allValues.length; i++) {
      const row = allValues[i];
      // Ensure checkbox column index is valid and row has enough columns
      if (checkboxColIdx !== -1 && row.length > checkboxColIdx && row[checkboxColIdx] === true) {
        const description = sanitizeInput(row[descriptionColIdx] || 'No Description');
        
        const partNumber = skuNumberColIdx !== -1 && row.length > skuNumberColIdx ? sanitizeInput(row[skuNumberColIdx] || '') : '';
        const manufacturer = manufacturerColIdx !== -1 && row.length > manufacturerColIdx ? sanitizeInput(row[manufacturerColIdx] || '') : '';
        const dimensions = dimensionsColIdx !== -1 && row.length > dimensionsColIdx ? sanitizeInput(row[dimensionsColIdx] || '') : '';

        itemsToEmail.push({
          rowNum: i + 1, // Keep 1-based for display or other logic if needed
          description: description,
          partNumber: partNumber,
          manufacturer: manufacturer,
          dimensions: dimensions
        });
        rowsToUpdate.push(i + 1);
      }
    }

    if (itemsToEmail.length === 0) {
      ss.toast("No items selected. Please check the boxes first.");
      return;
    }

    if (itemsToEmail.length > CONFIG.MAX_ITEMS_PER_EMAIL) {
      throw new Error(`Too many items selected. Maximum allowed is ${CONFIG.MAX_ITEMS_PER_EMAIL}`);
    }

    // Generate Email Content
    const projectName = getProjectNameFromProperties(); // Get project name from properties
    const defaultSubject = `${CONFIG.EMAIL_SUBJECT_PREFIX} - ${projectName} - ${CONFIG.YOUR_COMPANY_NAME}`;
    const { htmlBody, plainBody } = generateEmailBodies(itemsToEmail);

    let userPrimaryEmail = Session.getEffectiveUser().getEmail();
    let userAliases = [];
    try {
      userAliases = GmailApp.getAliases();
    } catch (e) {
      Logger.log("Could not fetch Gmail aliases: " + e.message);
      // User might not have Gmail enabled or script lacks permission, proceed with primary only
    }
    // Ensure primary email is in the list if not already included by getAliases()
    if (userPrimaryEmail && userAliases.indexOf(userPrimaryEmail) === -1) {
      userAliases.unshift(userPrimaryEmail); // Add to the beginning
    }

    // Create and show sidebar
    const htmlTemplate = HtmlService.createTemplateFromFile('sidebarHTML');
    htmlTemplate.defaultEmail = CONFIG.DEFAULT_VENDOR_EMAIL;
    htmlTemplate.defaultSubject = defaultSubject;
    htmlTemplate.htmlBodyContent = htmlBody;
    htmlTemplate.plainBodyContent = plainBody;
    htmlTemplate.rowsToUpdateJson = JSON.stringify(rowsToUpdate);
    htmlTemplate.userPrimaryEmail = userPrimaryEmail;
    htmlTemplate.userAliases = JSON.stringify(userAliases); // Pass as JSON string
    htmlTemplate.preferredSenderAlias = CONFIG.SENDER_ALIAS_EMAIL;

    const html = htmlTemplate.evaluate()
      .setTitle('Email Price Requests')
      .setWidth(600);
    ui.showSidebar(html);
  } catch (error) {
    Logger.log(`Error in openEmailSidebar: ${error.message}\nStack: ${error.stack}`);
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
  }
}