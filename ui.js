/**
 * Adds custom menus to the spreadsheet UI.
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
      
    // Client Dashboard menu
    ui.createMenu('Setup Sheet')
      // Calls showProjectDashboard to open the default view (e.g., Home)
      // Pass null or a default tab ID if your Project_Details_ expects one for the main view
      .addItem('Open Dashboard', 'showProjectDashboardDefault') 
      // Calls showProjectDashboard with 'roomManagerView' to open the Room Manager tab
      .addItem('1. Select Rooms', 'showProjectDashboardRoomManager') 
      // New menu item
      .addItem('2. Select Room Categories', 'showProjectDashboardRoomCategories')
      // New menu item
      .addItem('3. Select Items', 'showProjectDashboardItems')
      .addToUi();
    
      ui.createMenu('Sheet Manager')
      .addItem('Split Items by SPEC/FFE', 'splitItemsByFFE')
      .addToUi();
      
  } catch (error) {
    Logger.log(`Error in onOpen: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Configuration Error: ${error.message}`);
  }
}

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
   * Gets an appropriate icon name based on file mime type
   * @param {string} mimeType - The mime type of the file
   * @return {string} Name of the Material Icon to use
   */
  function getFileIcon(mimeType) {
    const icons = {
      'application/vnd.google-apps.document': 'description',
      'application/vnd.google-apps.spreadsheet': 'table_chart',
      'application/vnd.google-apps.presentation': 'slideshow',
      'application/pdf': 'picture_as_pdf',
      'application/vnd.google-apps.folder': 'folder',
      'image/jpeg': 'image',
      'image/png': 'image',
      'image/gif': 'gif',
      'application/zip': 'archive',
      'application/vnd.google-apps.form': 'assignment',
      'application/vnd.google-apps.drawing': 'brush',
      'text/plain': 'text_snippet',
      'text/html': 'code',
      'text/css': 'code',
      'text/javascript': 'code',
      'default': 'insert_drive_file'
    };
    return icons[mimeType] || icons['default'];
  }
  
  /**
   * Shows the folder picker dialog.
   */
  function showFolderPicker() {
    // Create HTML dialog from the FolderPicker.html file
    var html = HtmlService.createHtmlOutputFromFile('FolderPicker')
        .setWidth(600)
        .setHeight(400)
        .setTitle('Select Google Drive Folder');
    
    // Show as a modal dialog
    SpreadsheetApp.getUi().showModalDialog(html, 'Select Folder');
  }