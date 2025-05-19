/**
 * Include an HTML file in the template
 * This enables modular components in HTML templates
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSettings() {
  var props = PropertiesService.getScriptProperties();
  return {
    asanaToken: props.getProperty('ASANA_TOKEN') || '',
    sheetId: props.getProperty('SHEET_ID') || '',
    projectColor: props.getProperty('PROJECT_COLOR') || '#26717D'
  };
}

function saveSettings(asanaToken, sheetId, projectColor) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('ASANA_TOKEN', asanaToken);
  props.setProperty('SHEET_ID', sheetId);
  props.setProperty('PROJECT_COLOR', projectColor || '#26717D');
  return true;
}

/**
 * Fetches tasks from multiple Asana projects for Gantt chart display.
 * Only includes tasks with both start_on and due_on.
 * Groups by project and section, combining all into a single task list.
 * @return {Object} Object containing success status and tasks array
 */
function getAsanaTasksForGantt() {
  try {
    const asanaToken = PropertiesService.getScriptProperties().getProperty('ASANA_TOKEN');
    
    if (!asanaToken) {
      return {
        success: false,
        error: "Asana token not configured. Please set up your Asana integration in the Settings."
      };
    }
    
    // Get projects from Google Sheet instead of hardcoding
    const projectsResult = getProjects();
    if (!projectsResult.success) {
      return {
        success: false,
        error: "Failed to fetch project information: " + projectsResult.error
      };
    }
    
    // Filter projects that have an Asana Project ID
    const projects = projectsResult.projects.filter(project => project.asanaProjectId);
    
    // Log all projects to check which ones have Asana Project IDs
    Logger.log(`Found ${projects.length} projects with Asana Project IDs:`);
    projects.forEach(p => Logger.log(`Project: ${p.name}, ID: ${p.asanaProjectId}`));
    
    if (projects.length === 0) {
      return {
        success: false,
        error: "No projects with Asana Project IDs found in your Google Sheet."
      };
    }
    
    const resultTasks = [];
    
    for (const project of projects) {
      try {
        // Get the project name and ID
        const projectName = project.name;
        const projectId = project.asanaProjectId ? project.asanaProjectId.replace(/["']/g, '').trim() : '';
        const projectColor = project.projectColor || '#26717D'; // Use project color if available
        
        // Log details about this specific project
        Logger.log(`Processing project: ${projectName}, cleaned ID: "${projectId}"`);
        
        // Skip if project ID is missing
        if (!projectId) continue;
        
        // Fetch tasks with start_on and due_on fields for this project
        const ganttTaskOptFields = 'gid,name,completed,start_on,due_on,memberships.section.name,permalink_url';
        const ganttTaskQueryParams = { opt_fields: ganttTaskOptFields };
        Logger.log(`Fetching Gantt tasks for project: ${projectName} (${projectId}) with params: ${JSON.stringify(ganttTaskQueryParams)}`);
        
        const taskFetchResult = fetchAsanaTasksForProjectId(projectId, asanaToken, ganttTaskQueryParams);

        if (!taskFetchResult.success) {
          Logger.log(`Failed to fetch tasks for project ${projectName}. Error: ${taskFetchResult.error}`);
          continue; // Skip this project but continue with others
        }
        const projectTasks = taskFetchResult.tasks || [];
        
        // Log the number of tasks fetched and how many have start/due dates
        const tasksWithDates = projectTasks.filter(t => t.start_on && t.due_on).length;
        Logger.log(`Project ${projectName}: Retrieved ${projectTasks.length} tasks, ${tasksWithDates} have start/due dates`);
        
        // Process tasks for this project
        projectTasks.forEach(task => {
          if (task.start_on && task.due_on) {
            let sectionName = "Uncategorized";
            
            if (task.memberships && task.memberships.length > 0) {
              for (const membership of task.memberships) {
                // Get section name
                if (membership.section && membership.section.name) {
                  sectionName = membership.section.name;
                  break; // Once we found a section, we can stop looking
                }
              }
            }
            
            // Format the task name to include the project name
            const formattedName = `[${projectName}] ${task.name}`;
            
            resultTasks.push({
              id: task.gid,
              name: formattedName,
              project: projectName, // Use project name
              projectId: projectId, // Keep the ID for reference
              projectColor: projectColor, // Include project color
              section: sectionName,
              start_on: task.start_on,
              due_on: task.due_on,
              completed: task.completed,
              url: task.permalink_url || `https://app.asana.com/0/${projectId}/${task.gid}`
            });
          }
        });
      } catch (projectError) {
        // Log the error but continue processing other projects
        Logger.log(`Error fetching project ${project.name}: ${projectError.message}`);
      }
    }
    
    // Return the combined tasks from all projects
    return {
      success: true,
      tasks: resultTasks
    };
  } catch (error) {
    Logger.log("Error in getAsanaTasksForGantt: " + error.message);
    return {
      success: false,
      error: "Error fetching Asana tasks for Gantt: " + error.message
    };
  }
}

/**
 * Get all projects from the Google Sheet
 * @return {Object} Object containing success status and projects array
 */
function getProjects() {
  try {
    const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    
    if (!sheetId) {
      return {
        success: false,
        error: "Google Sheet ID not configured. Please set up your Sheet ID in the Settings."
      };
    }
    
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName('Projects');
    
    if (!sheet) {
      return {
        success: false,
        error: "Projects sheet not found. Please create a sheet named 'Projects' in your Google Sheet."
      };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Remove header row
    
    // Find column indices with the specified column names
    const nameIndex = headers.indexOf('Project_name');
    const clientNameIndex = headers.indexOf('Client_name');
    const clientEmailIndex = headers.indexOf('Client_email');
    const clientAddressIndex = headers.indexOf('Client_address');
    const statusIndex = headers.indexOf('Status');
    const asanaProjectIdIndex = headers.indexOf('AsanaProjectID');
    const sheetIdIndex = headers.indexOf('SheetID');
    const folderIdIndex = headers.indexOf('FolderID');
    const projectColorIndex = headers.indexOf('ProjectColor');
    const architectIndex = headers.indexOf('Architect');
    const architectEmailIndex = headers.indexOf('Architect_email');
    const contractorIndex = headers.indexOf('Contractor');
    const contractorEmailIndex = headers.indexOf('Contractor_email');
    
    if (nameIndex === -1) {
      return {
        success: false,
        error: "Required column 'Project_name' not found in the Projects sheet."
      };
    }
    
    const projects = data.map((row, index) => {
      return {
        id: index + 1, // Use row index as ID
        name: row[nameIndex] || '',
        client: clientNameIndex !== -1 ? row[clientNameIndex] || '' : '',
        clientEmail: clientEmailIndex !== -1 ? row[clientEmailIndex] || '' : '',
        clientAddress: clientAddressIndex !== -1 ? row[clientAddressIndex] || '' : '',
        status: statusIndex !== -1 ? row[statusIndex] || 'Not Started' : 'Not Started',
        asanaProjectId: asanaProjectIdIndex !== -1 ? String(row[asanaProjectIdIndex] || '').replace(/["']/g, '').trim() : '',
        sheetId: sheetIdIndex !== -1 ? row[sheetIdIndex] || '' : '',
        folderId: folderIdIndex !== -1 ? row[folderIdIndex] || '' : '',
        projectColor: projectColorIndex !== -1 ? row[projectColorIndex] || '#26717D' : '#26717D',
        architect: architectIndex !== -1 ? row[architectIndex] || '' : '',
        architectEmail: architectEmailIndex !== -1 ? row[architectEmailIndex] || '' : '',
        contractor: contractorIndex !== -1 ? row[contractorIndex] || '' : '',
        contractorEmail: contractorEmailIndex !== -1 ? row[contractorEmailIndex] || '' : ''
      };
    });
    
    return {
      success: true,
      projects: projects
    };
  } catch (error) {
    Logger.log("Error in getProjects: " + error.message);
    return {
      success: false,
      error: "Error fetching projects: " + error.message
    };
  }
}

/**
 * Ensures that the 'Projects' sheet has all required standard columns.
 * If a column is missing, it's added to the sheet header and the headers array.
 * @param {Sheet} sheet - The Google Apps Script Sheet object for 'Projects'.
 * @param {Array<string>} headers - The current array of header names from the sheet.
 * @return {Array<string>} The updated array of header names, including any added columns.
 * @private
 */
function _ensureProjectSheetColumns(sheet, headers) {
  const standardColumns = [
    'ProjectColor', 'Architect', 'Architect_email', 'Contractor', 'Contractor_email'
  ];

  let currentLastHeaderColumn = headers.length; // 0-based index, so length is the 1-based index of next new column

  standardColumns.forEach(columnName => {
    if (headers.indexOf(columnName) === -1) {
      currentLastHeaderColumn++; // Increment first to get 1-based column index for new column
      sheet.getRange(1, currentLastHeaderColumn).setValue(columnName).setFontWeight('bold');
      headers.push(columnName); // Add to the headers array being managed
    }
  });
  return headers;
}

/**
 * Add a new project to the Google Sheet
 * @param {Object} project - Project object with name, client, status, etc.
 * @return {Object} Object containing success status and newly added project
 */
function addProject(project) {
  try {
    const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    
    if (!sheetId) {
      return {
        success: false,
        error: "Google Sheet ID not configured. Please set up your Sheet ID in the Settings."
      };
    }
    
    const ss = SpreadsheetApp.openById(sheetId);
    let sheet = ss.getSheetByName('Projects');
    
    // Create Projects sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('Projects');
      // Add headers with only the specified column names
      sheet.appendRow(['Project_name', 'Client_name', 'Client_email', 'Client_address', 'Status', 'AsanaProjectID', 'SheetID', 'FolderID', 'ProjectColor', 'Architect', 'Architect_email', 'Contractor', 'Contractor_email']);
      sheet.getRange(1, 1, 1, 13).setFontWeight('bold');
    }
    
    // Get headers to ensure correct column order
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    headers = _ensureProjectSheetColumns(sheet, headers); // Use helper to add missing columns and update headers

    // Define/Re-define indices AFTER headers array might have been modified by the helper
    const nameIndex = headers.indexOf('Project_name');
    const clientNameIndex = headers.indexOf('Client_name');
    const clientEmailIndex = headers.indexOf('Client_email');
    const clientAddressIndex = headers.indexOf('Client_address');
    const statusIndex = headers.indexOf('Status');
    const asanaProjectIdIndex = headers.indexOf('AsanaProjectID');
    const sheetIdIndex = headers.indexOf('SheetID');
    const folderIdIndex = headers.indexOf('FolderID');
    const projectColorIndex = headers.indexOf('ProjectColor');
    const architectIndex = headers.indexOf('Architect');
    const architectEmailIndex = headers.indexOf('Architect_email');
    const contractorIndex = headers.indexOf('Contractor');
    const contractorEmailIndex = headers.indexOf('Contractor_email');
    
    // Create a new row with data in the correct column positions
    const newRow = [];
    for (let i = 0; i < headers.length; i++) {
      if (i === nameIndex) newRow[i] = project.name;
      else if (i === clientNameIndex) newRow[i] = project.client || '';
      else if (i === clientEmailIndex) newRow[i] = project.clientEmail || '';
      else if (i === clientAddressIndex) newRow[i] = project.clientAddress || '';
      else if (i === statusIndex) newRow[i] = project.status || 'Not Started';
      else if (i === asanaProjectIdIndex) newRow[i] = project.asanaProjectId ? `"${project.asanaProjectId}"` : '';
      else if (i === sheetIdIndex) newRow[i] = project.sheetId || '';
      else if (i === folderIdIndex) newRow[i] = project.folderId || '';
      else if (i === projectColorIndex || headers[i] === 'ProjectColor') newRow[i] = project.projectColor || '#26717D';
      else if (i === architectIndex || headers[i] === 'Architect') newRow[i] = project.architect || '';
      else if (i === architectEmailIndex || headers[i] === 'Architect_email') newRow[i] = project.architectEmail || '';
      else if (i === contractorIndex || headers[i] === 'Contractor') newRow[i] = project.contractor || '';
      else if (i === contractorEmailIndex || headers[i] === 'Contractor_email') newRow[i] = project.contractorEmail || '';
      else newRow[i] = ''; // Fill any other columns with empty string
    }
    
    // Append the new row
    sheet.appendRow(newRow);
    
    // Return success with the newly added project
    return {
      success: true,
      project: {
        id: sheet.getLastRow() - 1, // Row index as ID
        name: project.name,
        client: project.client || '',
        clientEmail: project.clientEmail || '',
        clientAddress: project.clientAddress || '',
        status: project.status || 'Not Started',
        asanaProjectId: project.asanaProjectId || '',
        sheetId: project.sheetId || '',
        folderId: project.folderId || '',
        projectColor: project.projectColor || '#26717D',
        architect: project.architect || '',
        architectEmail: project.architectEmail || '',
        contractor: project.contractor || '',
        contractorEmail: project.contractorEmail || ''
      }
    };
  } catch (error) {
    Logger.log("Error in addProject: " + error.message);
    return {
      success: false,
      error: "Error adding project: " + error.message
    };
  }
}

/**
 * Get a specific project by ID from the Google Sheet
 * @param {number} projectId - The project ID to fetch
 * @return {Object} Object containing success status and project data
 */
function getProjectById(projectId) {
  try {
    const result = getProjects();
    
    if (!result.success) {
      return result; // Return the error from getProjects
    }
    
    // Find the project with the matching ID
    const project = result.projects.find(p => p.id === parseInt(projectId, 10));
    
    if (!project) {
      return {
        success: false,
        error: `Project with ID ${projectId} not found.`
      };
    }
    
    return {
      success: true,
      project: project
    };
  } catch (error) {
    Logger.log("Error in getProjectById: " + error.message);
    return {
      success: false,
      error: "Error fetching project: " + error.message
    };
  }
}

/**
 * Update an existing project in the Google Sheet
 * @param {number} id - The row index of the project to update (1-based, 2 is the first data row)
 * @param {Object} project - Project object with updated data
 * @return {Object} Object containing success status and updated project
 */
function updateProject(id, project) {
  try {
    const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    
    if (!sheetId) {
      return {
        success: false,
        error: "Google Sheet ID not configured. Please set up your Sheet ID in the Settings."
      };
    }
    
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName('Projects');
    
    if (!sheet) {
      return {
        success: false,
        error: "Projects sheet not found. Please create a sheet named 'Projects' in your Google Sheet."
      };
    }
    
    // Convert to row number (add 1 for header row)
    const rowNumber = parseInt(id, 10) + 1;
    
    // Verify row exists
    if (rowNumber <= 1 || rowNumber > sheet.getLastRow()) {
      return {
        success: false,
        error: `Project with ID ${id} not found.`
      };
    }
    
    // Get headers to ensure correct column order
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    headers = _ensureProjectSheetColumns(sheet, headers); // Use helper to add missing columns and update headers
    
    // Define/Re-define indices AFTER headers array might have been modified by the helper
    const nameIndex = headers.indexOf('Project_name');
    const clientNameIndex = headers.indexOf('Client_name');
    const clientEmailIndex = headers.indexOf('Client_email');
    const clientAddressIndex = headers.indexOf('Client_address');
    const statusIndex = headers.indexOf('Status');
    const asanaProjectIdIndex = headers.indexOf('AsanaProjectID');
    const sheetIdIndex = headers.indexOf('SheetID');
    const folderIdIndex = headers.indexOf('FolderID');
    const projectColorIndex = headers.indexOf('ProjectColor');
    const architectIndex = headers.indexOf('Architect');
    const architectEmailIndex = headers.indexOf('Architect_email');
    const contractorIndex = headers.indexOf('Contractor');
    const contractorEmailIndex = headers.indexOf('Contractor_email');
    
    // Update the cells in the row
    if (nameIndex !== -1) sheet.getRange(rowNumber, nameIndex + 1).setValue(project.name || '');
    if (clientNameIndex !== -1) sheet.getRange(rowNumber, clientNameIndex + 1).setValue(project.client || '');
    if (clientEmailIndex !== -1) sheet.getRange(rowNumber, clientEmailIndex + 1).setValue(project.clientEmail || '');
    if (clientAddressIndex !== -1) sheet.getRange(rowNumber, clientAddressIndex + 1).setValue(project.clientAddress || '');
    if (statusIndex !== -1) sheet.getRange(rowNumber, statusIndex + 1).setValue(project.status || 'Not Started');
    if (asanaProjectIdIndex !== -1) sheet.getRange(rowNumber, asanaProjectIdIndex + 1).setValue(project.asanaProjectId ? `"${project.asanaProjectId}"` : '');
    if (sheetIdIndex !== -1) sheet.getRange(rowNumber, sheetIdIndex + 1).setValue(project.sheetId || '');
    if (folderIdIndex !== -1) sheet.getRange(rowNumber, folderIdIndex + 1).setValue(project.folderId || '');
    if (projectColorIndex !== -1) sheet.getRange(rowNumber, projectColorIndex + 1).setValue(project.projectColor || '#26717D');
    if (architectIndex !== -1) sheet.getRange(rowNumber, architectIndex + 1).setValue(project.architect || '');
    if (architectEmailIndex !== -1) sheet.getRange(rowNumber, architectEmailIndex + 1).setValue(project.architectEmail || '');
    if (contractorIndex !== -1) sheet.getRange(rowNumber, contractorIndex + 1).setValue(project.contractor || '');
    if (contractorEmailIndex !== -1) sheet.getRange(rowNumber, contractorEmailIndex + 1).setValue(project.contractorEmail || '');
    
    // Return success with the updated project
    return {
      success: true,
      project: {
        id: id, // Keep the same ID
        name: project.name,
        client: project.client || '',
        clientEmail: project.clientEmail || '',
        clientAddress: project.clientAddress || '',
        status: project.status || 'Not Started',
        asanaProjectId: project.asanaProjectId || '',
        sheetId: project.sheetId || '',
        folderId: project.folderId || '',
        projectColor: project.projectColor || '#26717D',
        architect: project.architect || '',
        architectEmail: project.architectEmail || '',
        contractor: project.contractor || '',
        contractorEmail: project.contractorEmail || ''
      }
    };
  } catch (error) {
    Logger.log("Error in updateProject: " + error.message);
    return {
      success: false,
      error: "Error updating project: " + error.message
    };
  }
}

/**
 * Fetches tasks from Asana for a specific project.
 * @param {string} projectId - The Asana project ID
 * @return {Object} Object containing success status, tasks grouped by section, and section order
 */
function getAsanaTasksForProject(projectId) {
  try {
    // If no project ID is provided, return error
    if (!projectId) {
      return {
        success: false,
        error: "No Asana project ID provided"
      };
    }
    
    // Clean the projectId by removing any quotes
    projectId = String(projectId).replace(/["']/g, '').trim();
    
    // Get Asana token from properties
    const asanaToken = PropertiesService.getScriptProperties().getProperty('ASANA_TOKEN');
    
    if (!asanaToken) {
      return {
        success: false,
        error: "Asana token not configured. Please set up your Asana integration in the Settings."
      };
    }
    
    Logger.log("Asana token: " + (asanaToken ? "present" : "missing"));
    Logger.log("Asana project ID: " + projectId);
    
    // Fetch tasks from Asana - Only fetch incomplete tasks with completed=false parameter
    const taskOptFields = 'name,completed,due_on,notes,assignee.name,memberships.section.name,permalink_url'; // Added permalink_url for consistency with Gantt
    const taskQueryParams = { opt_fields: taskOptFields, completed: 'false' }; // Ensure completed is a string 'false' for query param
    
    // Log the request URL for debugging (without token)
    Logger.log(`Requesting Asana tasks for project ${projectId} with params: ${JSON.stringify(taskQueryParams)}`);

    const tasksResult = fetchAsanaTasksForProjectId(projectId, asanaToken, taskQueryParams); // Direct function call
    
    Logger.log("Asana tasks response success: " + tasksResult.success);
    if (!tasksResult.success) {
      Logger.log("Error fetching Asana tasks: " + tasksResult.error);
      return { success: false, error: "Failed to fetch tasks from Asana: " + tasksResult.error };
    }
    
    // Parse the response
    const tasks = tasksResult.tasks || []; // tasksResult.tasks directly contains the array
    
    Logger.log("Received " + tasks.length + " open tasks from Asana");
    
    // Fetch sections for this project
    const sectionsResult = fetchAsanaSectionsForProjectId(projectId, asanaToken); // Direct function call
    
    Logger.log("Asana sections response success: " + sectionsResult.success);
    
    // Create a map of section GIDs to section names
    const sectionMap = {}; // This map is not strictly needed if we rely on sectionOrder from API
    const sectionOrder = [];
    
    if (sectionsResult.success && sectionsResult.sections) {
      sectionsResult.sections.forEach(section => {
        // sectionMap[section.gid] = section.name; // sectionMap might still be useful if tasks don't have section.name directly
        if (section.name) { // Ensure section name exists
          sectionOrder.push(section.name);
        }
      });
      
      Logger.log("Received " + sectionsResult.sections.length + " sections from Asana");
      Logger.log("Section order: " + sectionOrder.join(", "));
    } else {
      Logger.log("Failed to fetch sections, using fallback grouping. Error: " + (sectionsResult.error || 'Unknown error'));
    }
    
    // Group tasks by section
    const tasksBySection = {};
    let totalTasks = 0;
    
    tasks.forEach(task => {
      // Skip completed tasks (just in case some completed tasks are still returned)
      if (task.completed) return;
      
      totalTasks++;
      
      let sectionName = "Uncategorized";
      
      // Try to find the section this task belongs to
      if (task.memberships && task.memberships.length > 0) {
        for (const membership of task.memberships) {
          if (membership.section && membership.section.name) {
            sectionName = membership.section.name;
            break;
          }
        }
      }
      
      // Initialize section array if not exists
      if (!tasksBySection[sectionName]) {
        tasksBySection[sectionName] = [];
      }
      
      // Add task to its section
      tasksBySection[sectionName].push({
        name: task.name,
        completed: false, // All tasks should be incomplete due to `completed=false` query param
        dueDate: task.due_on,
        notes: task.notes,
        assignee: task.assignee ? task.assignee.name : null,
        url: task.permalink_url || `https://app.asana.com/0/${projectId}/${task.gid}` // Added permalink_url
      });
    });
    
    Logger.log(`Total open tasks: ${totalTasks}`);
    
    // Log task count per section
    Object.keys(tasksBySection).forEach(section => {
      Logger.log(`Section "${section}": ${tasksBySection[section].length} tasks`);
    });
    
    return {
      success: true,
      tasks: tasksBySection,
      sectionOrder: sectionOrder
    };
    
  } catch (error) {
    Logger.log("Error in getAsanaTasksForProject: " + error.message);
    return {
      success: false,
      error: "Error fetching Asana tasks: " + error.message
    };
  }
}

/**
 * Opens a Google Sheet and returns its URL.
 * This can be used to provide a direct link to the Sheet.
 * 
 * @param {string} sheetId - The ID of the Google Sheet
 * @return {Object} Object containing success status and URL
 */
function openProjectSheet(sheetId) {
  try {
    if (!sheetId) {
      return {
        success: false,
        error: "No Sheet ID provided"
      };
    }
    
    // Try to open the spreadsheet to verify it exists and is accessible
    try {
      const ss = SpreadsheetApp.openById(sheetId);
      const sheetName = ss.getName();
      
      // Construct the URL to the spreadsheet
      const url = `https://docs.google.com/spreadsheets/d/${sheetId}/edit`;
      
      return {
        success: true,
        url: url,
        name: sheetName
      };
    } catch (e) {
      Logger.log("Error opening spreadsheet: " + e.message);
      return {
        success: false,
        error: "Could not open spreadsheet. Make sure the Sheet ID is correct and you have access to it."
      };
    }
  } catch (error) {
    Logger.log("Error in openProjectSheet: " + error.message);
    return {
      success: false,
      error: "Error opening Google Sheet: " + error.message
    };
  }
}

/**
 * Loads Project_Details_.html content and passes project data to it
 * @param {string} projectId - The ID of the project to load details for
 * @return {string} The HTML content for the project details
 */
function getProjectDetailsContent(projectId) {
  try {
    // Fetch the project data
    const projectResult = getProjectById(projectId);
    
    if (!projectResult.success) {
      throw new Error("Failed to load project data: " + projectResult.error);
    }
    
    // Create a template from the Project_Details_.html file
    const template = HtmlService.createTemplateFromFile('Project_Details_');
    
    // Pass project data to the template
    template.project = projectResult.project;
    
    // Evaluate and return the HTML content
    return template.evaluate().getContent();
  } catch (error) {
    Logger.log("Error in getProjectDetailsContent: " + error.message);
    return "<div class='error-message'>Error loading project details: " + error.message + "</div>";
  }
}

/**
 * Makes a copy of the Master Item List Template, renames it to Master Item List,
 * and returns the URL to open the new sheet
 * 
 * @param {string} sheetId - The ID of the spreadsheet
 * @return {Object} Object containing success status and URL to the new sheet
 */
function openMasterItemListTemplate() {
  try {
    // Open the spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = spreadsheet.getSheetByName("Master Item List Template");
    
    if (!templateSheet) {
      return {
        success: false,
        error: "Master Item List Template sheet not found"
      };
    }
    
    // Check if Master Item List already exists
    let masterSheet = spreadsheet.getSheetByName("Master Item List");
    
    // If it exists, return an error instead of deleting
    if (masterSheet) {
      return {
        success: false,
        error: 'A sheet named "Master Item List" already exists. Please delete or rename it before creating a new one.'
      };
    }
    
    // Create a copy of the template sheet
    masterSheet = templateSheet.copyTo(spreadsheet);
    masterSheet.setName("Master Item List");
    
    // Activate the new sheet so it opens first
    masterSheet.activate();
    
    // Create a URL that will open this specific sheet
    const url = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/edit#gid=${masterSheet.getSheetId()}`;
    
    // Return the URL to the client
    return {
      success: true,
      url: url
    };
  } catch (error) {
    Logger.log("Error in openMasterItemListTemplate: " + error.message);
    return {
      success: false,
      error: "Error creating Master Item List sheet: " + error.message
    };
  }
}

/**
 * Sets the active sheet to "Master Item List" for the user.
 * Returns success/error.
 */
function setActiveMasterItemListSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Master Item List");
    if (!sheet) {
      return { success: false, error: 'Sheet "Master Item List" not found.' };
    }
    ss.setActiveSheet(sheet);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Creates a copy of Master Item List Template, then saves selected items to the copy.
 * This keeps the original template clean while providing the filled-in Master Item List.
 * 
 * @param {Array} items - Selected items to save
 * @return {Object} Object containing success status and result information
 */
function saveSelectedItemsWithTemplateCopy(items) {
  try {
    // Validate input
    if (!items || !Array.isArray(items) || items.length === 0) {
      return {
        success: false,
        error: "No items provided to save"
      };
    }
    
    // Step 1: Get the active spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Step 2: Find the template sheet
    const templateSheet = spreadsheet.getSheetByName("Master Item List Template");
    if (!templateSheet) {
      return {
        success: false,
        error: 'Template sheet "Master Item List Template" not found'
      };
    }
    
    // Step 3: Check if Master Item List already exists
    let masterSheet = spreadsheet.getSheetByName("Master Item List");
    
    // If it exists, return an error
    if (masterSheet) {
      return {
        success: false,
        error: 'A sheet named "Master Item List" already exists. Please delete or rename it before creating a new one.'
      };
    }
    
    // Step 4: Create a copy of the template
    masterSheet = templateSheet.copyTo(spreadsheet);
    masterSheet.setName("Master Item List");
    
    // Step 5: Save items to the new copy
    const sheet = masterSheet;
    const existingDataRange = sheet.getDataRange();
    const existingData = existingDataRange.getValues();
    
    // Get header row to find columns
    const headers = existingData[0];
    const roomCol = headers.indexOf("ROOM");
    const typeCol = headers.indexOf("TYPE");
    const itemCol = headers.indexOf("ITEM");
    const qtyCol = headers.indexOf("QUANTITY");
    const lowCol = headers.indexOf("LOW");
    const lowTotalCol = headers.indexOf("LOW TOTAL");
    const highCol = headers.indexOf("HIGH");
    const highTotalCol = headers.indexOf("HIGH TOTAL");
    const specFfeCol = headers.indexOf("SPEC/FFE");
    
    // Check that all required columns exist
    if (roomCol === -1 || typeCol === -1 || itemCol === -1 || qtyCol === -1 || specFfeCol === -1) {
      return {
        success: false,
        error: "Required columns missing in Master Item List"
      };
    }
    
    // Create array of new rows to add
    const newRows = [];
    
    // Process each selected item
    items.forEach((item, idx) => {
      const newRow = new Array(headers.length).fill("");
      const rowNum = 2 + idx; // Data starts at row 2
      newRow[roomCol] = (item.room || '').toUpperCase();
      newRow[typeCol] = (item.type || '').toUpperCase();
      newRow[itemCol] = (item.item || '').toUpperCase();
      newRow[qtyCol] = Math.max(1, parseInt(item.quantity) || 1);
      // LOW and HIGH as '0', LOW TOTAL and HIGH TOTAL as formulas
      if (lowCol !== -1) newRow[lowCol] = '0';
      if (lowTotalCol !== -1) newRow[lowTotalCol] = `=E${rowNum}*D${rowNum}`;
      if (highCol !== -1) newRow[highCol] = '0';
      if (highTotalCol !== -1) newRow[highTotalCol] = `=G${rowNum}*D${rowNum}`;
      if (specFfeCol !== -1) newRow[specFfeCol] = '';
      newRows.push(newRow);
    });
    
    // If we have new rows to add
    if (newRows.length > 0) {
      // Get the row where we should start adding data (after header)
      const startRow = 2; // 1-based index, row 2 is first row after header
      // Get the range to insert data
      const insertRange = sheet.getRange(startRow, 1, newRows.length, headers.length);
      // Insert data
      insertRange.setValues(newRows);
      // Create a SPEC/FFE dropdown validation for the last column
      if (specFfeCol !== -1) {
        const specFfeRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['SPEC', 'FFE'], true)
          .build();
        sheet.getRange(startRow, specFfeCol + 1, newRows.length, 1).setDataValidation(specFfeRule);
      }
      // Format quantity column as whole numbers
      sheet.getRange(startRow, qtyCol + 1, newRows.length, 1).setNumberFormat('0');
      // Format LOW, LOW TOTAL, HIGH, HIGH TOTAL columns as integers (no decimals)
      const intCols = [lowCol, lowTotalCol, highCol, highTotalCol].filter(idx => idx !== -1);
      intCols.forEach(col => {
        sheet.getRange(startRow, col + 1, newRows.length, 1).setNumberFormat('0');
      });
    }
    
    // Step 6: Activate the new sheet
    masterSheet.activate();
    
    // Return success
    return {
      success: true,
      itemCount: items.length,
      message: `Successfully created Master Item List with ${items.length} items`
    };
    
  } catch (error) {
    Logger.log("Error in saveSelectedItemsWithTemplateCopy: " + error.message);
    return {
      success: false,
      error: "Error saving items: " + error.message
    };
  }
}

