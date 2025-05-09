/**
 * Asana Integration Module
 * Contains all functionality related to the Asana task management integration.
 */

/**
 * Fetches tasks from Asana for the configured project.
 * Retrieves tasks, sections, and organizes them for display in the dashboard.
 * 
 * @return {Object} Object containing success status, tasks grouped by section, and section order
 */
function getAsanaTasks() {
    try {
      // Get Asana credentials (try both property sources for compatibility)
      const asanaToken = PropertiesService.getScriptProperties().getProperty('ASANA_TOKEN');
      const asanaProjectId = PropertiesService.getDocumentProperties().getProperty('asanaProjectId');
      
      Logger.log("Asana token: " + (asanaToken ? "present" : "missing"));
      Logger.log("Asana project ID: " + (asanaProjectId ? asanaProjectId : "missing"));
      
      // Check if credentials are available
      if (!asanaToken || !asanaProjectId) {
        return {
          success: false,
          error: "Asana credentials not configured. Please set up your Asana integration in the Settings."
        };
      }
      
      // Fetch tasks from Asana
      const tasksUrl = `https://app.asana.com/api/1.0/projects/${asanaProjectId}/tasks?opt_fields=name,completed,due_on,notes,assignee.name,memberships.section.name`;
      
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
      
      Logger.log("Received " + tasks.length + " tasks from Asana");
      
      // Log a few tasks for debugging
      if (tasks.length > 0) {
        Logger.log("First few tasks:");
        for (let i = 0; i < Math.min(3, tasks.length); i++) {
          Logger.log(JSON.stringify(tasks[i]));
        }
      }
      
      // Fetch sections for this project
      const sectionsUrl = `https://app.asana.com/api/1.0/projects/${asanaProjectId}/sections`;
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
      let completedTasks = 0;
      
      tasks.forEach(task => {
        totalTasks++;
        if (task.completed) completedTasks++;
        
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
          completed: task.completed,
          dueDate: task.due_on,
          notes: task.notes,
          assignee: task.assignee ? task.assignee.name : null,
          url: `https://app.asana.com/0/${asanaProjectId}/${task.gid}`
        });
      });
      
      Logger.log(`Total tasks: ${totalTasks}, Completed: ${completedTasks}`);
      
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
      Logger.log("Error in getAsanaTasks: " + error.message);
      return {
        success: false,
        error: "Error fetching Asana tasks: " + error.message
      };
    }
  }
  
  /**
   * Fetches sections from Asana for the configured project.
   * Used to populate dropdowns in the task creation form.
   * 
   * @return {Object} Object containing success status and sections array
   */
  function getAsanaSections() {
    try {
      // Get Asana credentials
      const asanaToken = PropertiesService.getDocumentProperties().getProperty('asanaToken');
      const asanaProjectId = PropertiesService.getDocumentProperties().getProperty('asanaProjectId');
      
      // Check if credentials are available
      if (!asanaToken || !asanaProjectId) {
        return {
          success: false,
          error: "Asana credentials not configured"
        };
      }
      
      // Set up request options
      const options = {
        method: 'get',
        headers: {
          'Authorization': 'Bearer ' + asanaToken
        },
        muteHttpExceptions: true
      };
      
      // Fetch sections for the project
      const sectionsUrl = `https://app.asana.com/api/1.0/projects/${asanaProjectId}/sections`;
      const response = UrlFetchApp.fetch(sectionsUrl, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode !== 200) {
        Logger.log("Error fetching Asana sections: " + response.getContentText());
        return {
          success: false,
          error: "Failed to fetch sections from Asana (HTTP " + responseCode + ")"
        };
      }
      
      // Parse response and extract sections
      const data = JSON.parse(response.getContentText());
      const sections = data.data || [];
      
      // Transform to simpler format for the frontend
      const simplifiedSections = sections.map(section => ({
        gid: section.gid,
        name: section.name
      }));
      
      Logger.log("Fetched " + simplifiedSections.length + " sections from Asana");
      
      return {
        success: true,
        sections: simplifiedSections
      };
      
    } catch (error) {
      Logger.log("Error in getAsanaSections: " + error.message);
      return {
        success: false,
        error: "Error fetching Asana sections: " + error.message
      };
    }
  }
  
  /**
   * Validates Asana credentials by making a test API call.
   * 
   * @param {string} token - The Asana Personal Access Token
   * @param {string} projectId - The Asana Project ID
   * @return {Object} Object containing validation status and error message if any
   */
  function validateAsanaCredentials(token, projectId) {
    try {
      if (!token || !projectId) {
        return {
          valid: false,
          error: "Asana token and project ID are required"
        };
      }
      
      // Set up request options with authorization header
      const options = {
        method: 'get',
        headers: {
          'Authorization': 'Bearer ' + token
        },
        muteHttpExceptions: true
      };
      
      // Try to fetch project details to validate the credentials
      const url = `https://app.asana.com/api/1.0/projects/${projectId}`;
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      
      Logger.log("Asana validation response code: " + responseCode);
      
      if (responseCode !== 200) {
        const errorBody = response.getContentText();
        Logger.log("Asana validation error: " + errorBody);
        
        if (responseCode === 401) {
          return {
            valid: false,
            error: "Invalid Asana token. Please check your Personal Access Token."
          };
        } else if (responseCode === 404) {
          return {
            valid: false,
            error: "Asana project not found. Please check your Project ID."
          };
        } else {
          return {
            valid: false,
            error: "Error validating Asana credentials (HTTP " + responseCode + ")"
          };
        }
      }
      
      // Credentials are valid
      const data = JSON.parse(response.getContentText());
      return {
        valid: true,
        projectName: data.data ? data.data.name : "Unknown Project"
      };
      
    } catch (error) {
      Logger.log("Error in validateAsanaCredentials: " + error.message);
      return {
        valid: false,
        error: "Error validating Asana credentials: " + error.message
      };
    }
  }
  
  /**
   * Creates a new task in Asana.
   * 
   * @param {Object} taskData - Object containing task details
   * @param {string} taskData.name - The name/title of the task
   * @param {string} taskData.notes - Task description or notes (optional)
   * @param {string} taskData.dueDate - Due date in YYYY-MM-DD format (optional)
   * @param {string} taskData.assignee - Email of the assignee (optional)
   * @param {string} taskData.section - Name of the section (optional)
   * @return {Object} Object containing success status and task details or error message
   */
  function createAsanaTask(taskData) {
    try {
      // Get Asana credentials
      const asanaToken = PropertiesService.getScriptProperties().getProperty('ASANA_TOKEN');
      const asanaProjectId = PropertiesService.getDocumentProperties().getProperty('asanaProjectId');
      
      if (!asanaToken || !asanaProjectId) {
        return {
          success: false,
          error: "Asana credentials not configured"
        };
      }
      
      if (!taskData || !taskData.name) {
        return {
          success: false,
          error: "Task name is required"
        };
      }
      
      // Prepare task data for Asana API
      const payload = {
        data: {
          name: taskData.name,
          projects: [asanaProjectId]
        }
      };
      
      // Add optional fields if provided
      if (taskData.notes) payload.data.notes = taskData.notes;
      if (taskData.dueDate) payload.data.due_on = taskData.dueDate;
      
      // Set up request options
      const options = {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'Authorization': 'Bearer ' + asanaToken
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      
      // Create the task
      const url = 'https://app.asana.com/api/1.0/tasks';
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode !== 201) {
        Logger.log("Error creating Asana task: " + response.getContentText());
        return {
          success: false,
          error: "Failed to create task (HTTP " + responseCode + ")"
        };
      }
      
      // Task created successfully
      const data = JSON.parse(response.getContentText());
      const newTask = data.data;
      
      // Handle section assignment if provided
      if (taskData.section && newTask.gid) {
        assignTaskToSection(newTask.gid, taskData.section, asanaToken, asanaProjectId);
      }
      
      // Handle assignee if provided (by email)
      if (taskData.assignee && newTask.gid) {
        assignTaskToUser(newTask.gid, taskData.assignee, asanaToken);
      }
      
      return {
        success: true,
        task: {
          id: newTask.gid,
          name: newTask.name,
          url: `https://app.asana.com/0/${asanaProjectId}/${newTask.gid}`
        }
      };
      
    } catch (error) {
      Logger.log("Error in createAsanaTask: " + error.message);
      return {
        success: false,
        error: "Error creating Asana task: " + error.message
      };
    }
  }
  
  /**
   * Helper function to assign a task to a specific section.
   * 
   * @private
   * @param {string} taskGid - The Asana task GID
   * @param {string} sectionName - The name of the section
   * @param {string} token - The Asana token
   * @param {string} projectId - The Asana project ID
   */
  function assignTaskToSection(taskGid, sectionName, token, projectId) {
    try {
      // First, get all sections to find the right one
      const sectionsUrl = `https://app.asana.com/api/1.0/projects/${projectId}/sections`;
      const options = {
        method: 'get',
        headers: {
          'Authorization': 'Bearer ' + token
        },
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(sectionsUrl, options);
      if (response.getResponseCode() !== 200) {
        Logger.log("Failed to fetch sections for task assignment");
        return;
      }
      
      const sections = JSON.parse(response.getContentText()).data;
      const targetSection = sections.find(section => section.name === sectionName);
      
      if (!targetSection) {
        Logger.log(`Section "${sectionName}" not found`);
        return;
      }
      
      // Assign the task to the section
      const addToSectionUrl = `https://app.asana.com/api/1.0/sections/${targetSection.gid}/addTask`;
      const addOptions = {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'Authorization': 'Bearer ' + token
        },
        payload: JSON.stringify({
          data: { task: taskGid }
        }),
        muteHttpExceptions: true
      };
      
      const addResponse = UrlFetchApp.fetch(addToSectionUrl, addOptions);
      Logger.log("Section assignment response: " + addResponse.getResponseCode());
      
    } catch (error) {
      Logger.log("Error assigning task to section: " + error.message);
    }
  }
  
  /**
   * Helper function to assign a task to a user by email.
   * 
   * @private
   * @param {string} taskGid - The Asana task GID
   * @param {string} userEmail - The email of the user to assign
   * @param {string} token - The Asana token
   */
  function assignTaskToUser(taskGid, userEmail, token) {
    try {
      // Find user by email
      const usersUrl = `https://app.asana.com/api/1.0/users?opt_fields=email`;
      const options = {
        method: 'get',
        headers: {
          'Authorization': 'Bearer ' + token
        },
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(usersUrl, options);
      if (response.getResponseCode() !== 200) {
        Logger.log("Failed to fetch users for task assignment");
        return;
      }
      
      const users = JSON.parse(response.getContentText()).data;
      const targetUser = users.find(user => user.email === userEmail);
      
      if (!targetUser) {
        Logger.log(`User with email "${userEmail}" not found`);
        return;
      }
      
      // Assign the task to the user
      const updateTaskUrl = `https://app.asana.com/api/1.0/tasks/${taskGid}`;
      const updateOptions = {
        method: 'put',
        contentType: 'application/json',
        headers: {
          'Authorization': 'Bearer ' + token
        },
        payload: JSON.stringify({
          data: { assignee: targetUser.gid }
        }),
        muteHttpExceptions: true
      };
      
      const updateResponse = UrlFetchApp.fetch(updateTaskUrl, updateOptions);
      Logger.log("User assignment response: " + updateResponse.getResponseCode());
      
    } catch (error) {
      Logger.log("Error assigning task to user: " + error.message);
    }
  }
  
  /**
   * Fetches tasks from Asana for the configured project, for Gantt chart display.
   * Only includes tasks with both start_on and due_on.
   * Groups by section and returns id, name, section, start_on, due_on, completed, url.
   * @return {Object} Object containing success status and tasks array
   */
  function getAsanaTasksForGantt() {
    try {
      const asanaToken = PropertiesService.getScriptProperties().getProperty('ASANA_TOKEN');
      const asanaProjectId = PropertiesService.getDocumentProperties().getProperty('asanaProjectId');
      if (!asanaToken || !asanaProjectId) {
        return {
          success: false,
          error: "Asana credentials not configured. Please set up your Asana integration in the Settings."
        };
      }
      // Fetch tasks with start_on and due_on fields
      const tasksUrl = `https://app.asana.com/api/1.0/projects/${asanaProjectId}/tasks?opt_fields=gid,name,completed,start_on,due_on,memberships.section.name,permalink_url`;
      const options = {
        method: 'get',
        headers: {
          'Authorization': 'Bearer ' + asanaToken
        },
        muteHttpExceptions: true
      };
      Logger.log(tasksUrl);
      const response = UrlFetchApp.fetch(tasksUrl, options);
      const responseCode = response.getResponseCode();
      if (responseCode !== 200) {
        return {
          success: false,
          error: "Failed to fetch tasks from Asana (HTTP " + responseCode + ")"
        };
      }
      const data = JSON.parse(response.getContentText());
      const tasks = data.data || [];
      // Group by section, but return a flat array for Gantt
      const resultTasks = [];
      tasks.forEach(task => {
        if (task.start_on && task.due_on) {
          let sectionName = "Uncategorized";
          if (task.memberships && task.memberships.length > 0) {
            for (const membership of task.memberships) {
              if (membership.section && membership.section.name) {
                sectionName = membership.section.name;
                break;
              }
            }
          }
          resultTasks.push({
            id: task.gid,
            name: task.name,
            section: sectionName,
            start_on: task.start_on,
            due_on: task.due_on,
            completed: task.completed,
            url: task.permalink_url || `https://app.asana.com/0/${asanaProjectId}/${task.gid}`
          });
        }
      });
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