/**
 * Include an HTML file in the template
 * This enables modular components in HTML templates
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  const page = e.parameter.page || 'dashboard';
  const id = e.parameter.id || null;
  
  let template;
  
  switch (page.toLowerCase()) {
    case 'project':
      template = HtmlService.createTemplateFromFile('Project_Details');
      template.projectId = id;
      break;
    case 'gantt':
      template = HtmlService.createTemplateFromFile('Dashboard');
      template.mode = 'gantt';
      template.projectId = id;
      break;
    default:
      template = HtmlService.createTemplateFromFile('Dashboard');
      template.mode = 'dashboard';
  }
  
  // Return the evaluated template as an HtmlOutput object
  const output = template.evaluate()
    .setTitle('Project Management Dashboard')
    .setFaviconUrl('https://app.asana.com/favicon.ico');
  
  return output;
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
        const tasksUrl = `https://app.asana.com/api/1.0/projects/${projectId}/tasks?opt_fields=gid,name,completed,start_on,due_on,memberships.section.name,permalink_url`;
        const options = {
          method: 'get',
          headers: {
            'Authorization': 'Bearer ' + asanaToken
          },
          muteHttpExceptions: true
        };
        
        Logger.log(`Fetching tasks for project: ${projectName} (${projectId})`);
        const response = UrlFetchApp.fetch(tasksUrl, options);
        const responseCode = response.getResponseCode();
        
        if (responseCode !== 200) {
          Logger.log(`Failed to fetch tasks for project ${projectName} (HTTP ${responseCode})`);
          continue; // Skip this project but continue with others
        }
        
        const data = JSON.parse(response.getContentText());
        const projectTasks = data.data || [];
        
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
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
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
    
    // Add ProjectColor column if it doesn't exist
    if (projectColorIndex === -1) {
      const lastCol = headers.length + 1;
      sheet.getRange(1, lastCol).setValue('ProjectColor');
      sheet.getRange(1, lastCol).setFontWeight('bold');
      headers.push('ProjectColor');
    }
    
    // Add Architect column if it doesn't exist
    if (architectIndex === -1) {
      const lastCol = headers.length + 1;
      sheet.getRange(1, lastCol).setValue('Architect');
      sheet.getRange(1, lastCol).setFontWeight('bold');
      headers.push('Architect');
    }
    
    // Add Architect_email column if it doesn't exist
    if (architectEmailIndex === -1) {
      const lastCol = headers.length + 1;
      sheet.getRange(1, lastCol).setValue('Architect_email');
      sheet.getRange(1, lastCol).setFontWeight('bold');
      headers.push('Architect_email');
    }
    
    // Add Contractor column if it doesn't exist
    if (contractorIndex === -1) {
      const lastCol = headers.length + 1;
      sheet.getRange(1, lastCol).setValue('Contractor');
      sheet.getRange(1, lastCol).setFontWeight('bold');
      headers.push('Contractor');
    }
    
    // Add Contractor_email column if it doesn't exist
    if (contractorEmailIndex === -1) {
      const lastCol = headers.length + 1;
      sheet.getRange(1, lastCol).setValue('Contractor_email');
      sheet.getRange(1, lastCol).setFontWeight('bold');
      headers.push('Contractor_email');
    }
    
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
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
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
    
    // Add missing columns if they don't exist (same code as in addProject)
    // Add ProjectColor column if it doesn't exist
    if (projectColorIndex === -1) {
      const lastCol = headers.length + 1;
      sheet.getRange(1, lastCol).setValue('ProjectColor');
      sheet.getRange(1, lastCol).setFontWeight('bold');
      headers.push('ProjectColor');
    }
    
    // Add Architect column if it doesn't exist
    if (architectIndex === -1) {
      const lastCol = headers.length + 1;
      sheet.getRange(1, lastCol).setValue('Architect');
      sheet.getRange(1, lastCol).setFontWeight('bold');
      headers.push('Architect');
    }
    
    // Add Architect_email column if it doesn't exist
    if (architectEmailIndex === -1) {
      const lastCol = headers.length + 1;
      sheet.getRange(1, lastCol).setValue('Architect_email');
      sheet.getRange(1, lastCol).setFontWeight('bold');
      headers.push('Architect_email');
    }
    
    // Add Contractor column if it doesn't exist
    if (contractorIndex === -1) {
      const lastCol = headers.length + 1;
      sheet.getRange(1, lastCol).setValue('Contractor');
      sheet.getRange(1, lastCol).setFontWeight('bold');
      headers.push('Contractor');
    }
    
    // Add Contractor_email column if it doesn't exist
    if (contractorEmailIndex === -1) {
      const lastCol = headers.length + 1;
      sheet.getRange(1, lastCol).setValue('Contractor_email');
      sheet.getRange(1, lastCol).setFontWeight('bold');
      headers.push('Contractor_email');
    }
    
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
    const tasksUrl = `https://app.asana.com/api/1.0/projects/${projectId}/tasks?opt_fields=name,completed,due_on,notes,assignee.name,memberships.section.name&completed=false`;
    
    // Log the request URL for debugging (without token)
    Logger.log("Requesting Asana tasks from: " + tasksUrl);
    
    // Set up request options with authorization header
    const options = {
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + asanaToken
      },
      muteHttpExceptions: true
    };
    
    // Make the API request
    const response = UrlFetchApp.fetch(tasksUrl, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    
    Logger.log("Asana tasks response code: " + responseCode);
    Logger.log("Asana tasks response length: " + responseBody.length + " characters");
    
    // Check for successful response
    if (responseCode !== 200) {
      Logger.log("Error fetching Asana tasks: " + responseBody);
      return {
        success: false,
        error: "Failed to fetch tasks from Asana (HTTP " + responseCode + ")"
      };
    }
    
    // Parse the response
    const data = JSON.parse(responseBody);
    const tasks = data.data || [];
    
    Logger.log("Received " + tasks.length + " open tasks from Asana");
    
    // Fetch sections for this project
    const sectionsUrl = `https://app.asana.com/api/1.0/projects/${projectId}/sections`;
    const sectionsResponse = UrlFetchApp.fetch(sectionsUrl, options);
    const sectionsResponseCode = sectionsResponse.getResponseCode();
    const sectionsData = JSON.parse(sectionsResponse.getContentText());
    
    Logger.log("Asana sections response code: " + sectionsResponseCode);
    
    // Create a map of section GIDs to section names
    const sectionMap = {};
    const sectionOrder = [];
    
    if (sectionsResponseCode === 200 && sectionsData.data) {
      sectionsData.data.forEach(section => {
        sectionMap[section.gid] = section.name;
        sectionOrder.push(section.name);
      });
      
      Logger.log("Received " + sectionsData.data.length + " sections from Asana");
      Logger.log("Section order: " + sectionOrder.join(", "));
    } else {
      Logger.log("Failed to fetch sections, using fallback grouping");
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
        completed: false, // All tasks should be incomplete
        dueDate: task.due_on,
        notes: task.notes,
        assignee: task.assignee ? task.assignee.name : null,
        url: `https://app.asana.com/0/${projectId}/${task.gid}`
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