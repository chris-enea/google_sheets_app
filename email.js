/**
 * Helper function to generate both HTML and plain text email bodies.
 */
function generateEmailBodies(itemsToEmail, customMessage = '') {
    try {
      // Default message if none provided
      const defaultMessage = `Dear Vendor,\n\nWe would like to request price quotes and current availability for the following items:\n\n`;
      const defaultHtmlMessage = `
        <p>Dear Vendor,</p>
        <p>We would like to request price quotes and current availability for the following items:</p>
      `;
  
      // Use custom message if provided, otherwise use default
      const plainIntro = customMessage || defaultMessage;
      const htmlIntro = customMessage ? 
        `<p>${customMessage.replace(/\n/g, '</p><p>')}</p>` : 
        defaultHtmlMessage;
  
      let plainBody = plainIntro;
      let htmlBody = htmlIntro + `
        <table border="1" style="border-collapse: collapse; width: 100%; font-size: 10pt; margin-bottom: 15px;">
          <thead style="background-color: #f2f2f2;">
            <tr>
              <th style="padding: 5px; text-align: left;">Description</th>
              ${CONFIG.SKU_NUMBER_COL_NAME ? '<th style="padding: 5px; text-align: left;">SKU</th>' : ''}
              ${CONFIG.MANUFACTURER_COL_NAME ? '<th style="padding: 5px; text-align: left;">Manufacturer</th>' : ''}
            </tr>
          </thead>
          <tbody>
      `;
  
      itemsToEmail.forEach(item => {
        // Ensure we have valid item data
        if (!item) return;
        
        const description = sanitizeInput(item.description || 'No Description');
        const type = item.type ? ' - ' + sanitizeInput(item.type) : '';
        const quantity = item.quantity ? ' (Qty: ' + sanitizeInput(item.quantity) + ')' : '';
        
        // Only include dimensions in description, manufacturer is in its own column
        const dimensions = item.dimensions ? '\nDimensions: ' + sanitizeInput(item.dimensions) : '';
        
        const fullDescription = [description, type, quantity].filter(Boolean).join('');
        
        // Only add to email if we have at least a description
        if (description !== 'No Description') {
          plainBody += `- ${fullDescription}${dimensions}\n\n`;
        htmlBody += `
            <tr>
                <td style="padding: 5px;">
                  ${fullDescription}
                  ${item.dimensions ? '<br>Dimensions: ' + item.dimensions : ''}
                </td>
                ${CONFIG.SKU_NUMBER_COL_NAME ? '<td style="padding: 5px;">' + (item.partNumber || '') + '</td>' : ''}
                ${CONFIG.MANUFACTURER_COL_NAME ? '<td style="padding: 5px;">' + (item.manufacturer || '') + '</td>' : ''}
            </tr>
        `;
        }
      });
  
      const defaultClosing = `\nPlease provide the unit price and expected lead time for each item listed above.\n\nThank you,\n${CONFIG.YOUR_COMPANY_NAME}`;
      const defaultHtmlClosing = `
          </tbody>
        </table>
        <p>Please provide the unit price and expected lead time for each item listed above.</p>
        <p>Thank you,<br>${CONFIG.YOUR_COMPANY_NAME}</p>
      `;
  
      plainBody += defaultClosing;
      htmlBody += defaultHtmlClosing;
  
      return { htmlBody, plainBody };
    } catch (error) {
      Logger.log(`Error in generateEmailBodies: ${error.message}\nStack: ${error.stack}`);
      throw new Error('Failed to generate email content');
    }
  }
  
  /**
   * Sends the email directly using data submitted from the sidebar.
   * @param {object} emailDetails - Object containing recipient, subject, htmlBody, plainBody, rowsToUpdateJson, fromEmail.
   * @return {object} An object with success status and a message string.
   */
  function sendEmailDirectly(emailDetails) {
    try {
      // Input validation (basic)
      if (!emailDetails || typeof emailDetails !== 'object') {
        return { success: false, message: "Invalid email details format." };
      }
      const requiredFields = ['recipient', 'subject', 'htmlBody', 'rowsToUpdateJson', 'fromEmail'];
      for (const field of requiredFields) {
        if (!emailDetails[field]) {
          return { success: false, message: `Missing required field: ${field}` };
        }
      }
      if (!isValidEmail(emailDetails.recipient)) {
        return { success: false, message: "Invalid recipient email address format." };
      }
      if (!isValidEmail(emailDetails.fromEmail)) {
        // This should be caught by client-side, but good to double check
        return { success: false, message: "Invalid 'From' email address format." };
      }
  
      const projectName = getProjectNameFromProperties();
      const finalSubject = `${CONFIG.EMAIL_SUBJECT_PREFIX} - ${projectName} - ${CONFIG.YOUR_COMPANY_NAME}`;
  
      GmailApp.sendEmail(
        emailDetails.recipient,
        finalSubject, // Use the server-generated subject for consistency
        emailDetails.plainBody || "See HTML body.",
        {
          htmlBody: emailDetails.htmlBody,
          name: CONFIG.YOUR_COMPANY_NAME,
          from: emailDetails.fromEmail // Use the 'fromEmail' passed from the client
        }
      );
      Logger.log(`Email sent to ${emailDetails.recipient} from ${emailDetails.fromEmail} with subject: ${finalSubject}`);
  
      // Update sheet
      const rowsToUpdate = JSON.parse(emailDetails.rowsToUpdateJson);
      if (rowsToUpdate.length > 0) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
        
        if (!sheet) {
          // Log error but don't let sheet update failure prevent success message for email send
          Logger.log(`Sheet ${CONFIG.SHEET_NAME} not found during update phase after sending email.`);
        } else {
          const headerRowValues = sheet.getRange(CONFIG.HEADER_ROWS, 1, 1, sheet.getLastColumn()).getValues()[0];
          function getColumnIndex(columnName, headerArray) {
            const index = headerArray.indexOf(columnName);
            if (index === -1) {
              Logger.log(`Column "${columnName}" not found in header row. This column won't be updated.`);
              return -1;
            }
            return index;
          }
          const statusColIdx = CONFIG.STATUS_COL_NAME ? getColumnIndex(CONFIG.STATUS_COL_NAME, headerRowValues) : -1;
          const checkboxColIdx = getColumnIndex(CONFIG.CHECKBOX_COL_NAME, headerRowValues);
          if (checkboxColIdx === -1) {
              Logger.log(`Critical column "${CONFIG.CHECKBOX_COL_NAME}" not found. Cannot uncheck boxes.`);
          } else {
              const now = new Date();
              const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
              for (const rowNum of rowsToUpdate) {
                try {
                  if (statusColIdx !== -1) {
                    sheet.getRange(rowNum, statusColIdx + 1).setValue(`Emailed ${timestamp} (Sent from: ${emailDetails.fromEmail})`);
                  }
                  sheet.getRange(rowNum, checkboxColIdx + 1).setValue(false);
                } catch (e) {
                  Logger.log(`Error updating row ${rowNum} after sending email: ${e.message}`);
                }
              }
          }
        }
      }
      return { success: true, message: `Email successfully sent from ${emailDetails.fromEmail}!` };
    } catch (error) {
      Logger.log(`Error in sendEmailDirectly: ${error.message}\nStack: ${error.stack}`);
      return { success: false, message: `Failed to send email: ${error.message}` };
    }
  }
  
  /**
   * Gets all email vendors and their associated items from the sheet.
   */
  function getEmailVendors() {
    try {
      console.log("Starting getEmailVendors function");
      validateConfig();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
      
      if (!sheet) {
        console.error(`Sheet ${CONFIG.SHEET_NAME} not found`);
        throw new Error(`Sheet ${CONFIG.SHEET_NAME} not found`);
      }
  
      console.log("Getting data range from sheet");
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      
      // Log headers for debugging
      console.log("Headers found: " + JSON.stringify(headers));
      console.log("Looking for Manufacturer column with name: " + CONFIG.MANUFACTURER_COL_NAME);
      
      // Find column indices
      const vendorColIndex = headers.indexOf(CONFIG.VENDOR_COL_NAME);
      const propertyColIndex = headers.indexOf(CONFIG.PROPERTY_COL_NAME);
      const roomColIndex = headers.indexOf(CONFIG.ROOM_COL_NAME);
      const descriptionColIndex = headers.indexOf(CONFIG.DESCRIPTION_COL_NAME);
      const typeColIndex = headers.indexOf(CONFIG.TYPE_COL_NAME);
      const quantityColIndex = headers.indexOf(CONFIG.QUANTITY_COL_NAME);
      const manufacturerColIndex = headers.indexOf(CONFIG.MANUFACTURER_COL_NAME);
      const skuNumberColIndex = headers.indexOf(CONFIG.SKU_NUMBER_COL_NAME);
  
      // Log column indices for debugging
      console.log("Column indices:");
      console.log("Vendor: " + vendorColIndex);
      console.log("Property: " + propertyColIndex);
      console.log("Room: " + roomColIndex);
      console.log("Description: " + descriptionColIndex);
      console.log("Type: " + typeColIndex);
      console.log("Quantity: " + quantityColIndex);
      console.log("Manufacturer: " + manufacturerColIndex);
      console.log("SKU: " + skuNumberColIndex);
  
      if (vendorColIndex === -1) {
        console.error("Required column (Vendor) not found in sheet");
        throw new Error("Required column (Vendor) not found in sheet");
      }
  
      // Group items by vendor
      const vendors = {};
      console.log("Starting to process rows. Total rows: " + data.length);
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const vendorName = row[vendorColIndex];
  
        if (!vendorName) {
          console.log("Skipping row " + i + " - no vendor name");
          continue;
        }
  
        if (!vendors[vendorName]) {
          console.log("Creating new vendor entry for: " + vendorName);
          vendors[vendorName] = {
            name: vendorName,
            items: []
          };
        }
  
        // Log row data for debugging
        console.log("Processing row " + i + " for vendor " + vendorName);
        console.log("Manufacturer value: " + (manufacturerColIndex !== -1 ? row[manufacturerColIndex] : 'not found'));
  
        vendors[vendorName].items.push({
          property: propertyColIndex !== -1 ? row[propertyColIndex] : '',
          room: roomColIndex !== -1 ? row[roomColIndex] : '',
          description: descriptionColIndex !== -1 ? row[descriptionColIndex] : '',
          type: typeColIndex !== -1 ? row[typeColIndex] : '',
          quantity: quantityColIndex !== -1 ? row[quantityColIndex] : '',
          manufacturer: manufacturerColIndex !== -1 ? row[manufacturerColIndex] : '',
          partNumber: skuNumberColIndex !== -1 ? row[skuNumberColIndex] : ''
        });
      }
  
      console.log("Finished processing all rows");
      return {
        vendors: Object.values(vendors)
      };
    } catch (error) {
      console.error(`Error in getEmailVendors: ${error.message}\nStack: ${error.stack}`);
      throw new Error(`Failed to get vendors: ${error.message}`);
    }
  }
  
  /**
   * Creates a draft email for a specific vendor.
   */
  function createVendorEmailDraft(vendorName, emailAddress, customMessage = '') {
    try {
      validateConfig();
      const vendors = getEmailVendors().vendors;
      const vendor = vendors.find(v => v.name === vendorName);
      
      if (!vendor) {
        throw new Error(`Vendor ${vendorName} not found`);
      }
  
      if (!isValidEmail(emailAddress)) {
        throw new Error("Invalid email address format");
      }
  
      const { htmlBody, plainBody } = generateEmailBodies(vendor.items, customMessage);
      
      // Get the sheet name for the subject
      const projectName = getProjectNameFromProperties();
      
      // Create a draft email
      GmailApp.createDraft(
        emailAddress,
        `${CONFIG.EMAIL_SUBJECT_PREFIX} - ${projectName} - ${CONFIG.YOUR_COMPANY_NAME}`,
        plainBody,
        {
          htmlBody: htmlBody,
          name: CONFIG.YOUR_COMPANY_NAME
        }
      );
  
      return "Draft email created successfully for " + vendorName;
    } catch (error) {
      Logger.log(`Error in createVendorEmailDraft: ${error.message}\nStack: ${error.stack}`);
      throw new Error(`Failed to create draft email: ${error.message}`);
    }
  }
  
  /**
   * Creates a draft email and returns its URL
   */
  function createAndOpenDraft(emailDetails) {
    try {
      // Input validation
      if (!emailDetails || typeof emailDetails !== 'object') {
        // Return error object for consistency with how client handles it
        return { success: false, message: "Invalid email details format" };
      }
  
      // Add fromEmail to required fields check
      const requiredFields = ['recipient', 'subject', 'htmlBody', 'rowsToUpdateJson', 'fromEmail'];
      for (const field of requiredFields) {
        if (!emailDetails[field]) {
          return { success: false, message: `Missing required field: ${field}` };
        }
      }
  
      if (!isValidEmail(emailDetails.recipient)) {
        return { success: false, message: "Invalid recipient email address format" };
      }
      if (!isValidEmail(emailDetails.fromEmail)) {
         // This should be caught by client-side, but good to double check
        return { success: false, message: "Invalid 'From' email address format." };
      }
  
      const projectName = getProjectNameFromProperties();
      const finalSubject = `${CONFIG.EMAIL_SUBJECT_PREFIX} - ${projectName} - ${CONFIG.YOUR_COMPANY_NAME}`;
      let draft;
  
      draft = GmailApp.createDraft(
        emailDetails.recipient,
        finalSubject, // Use server-generated subject
        emailDetails.plainBody || emailDetails.htmlBody,
        {
          htmlBody: emailDetails.htmlBody,
          name: CONFIG.YOUR_COMPANY_NAME,
          from: emailDetails.fromEmail // Use the 'fromEmail' passed from the client
        }
      );
      Logger.log(`Draft created for ${emailDetails.recipient} from ${emailDetails.fromEmail} with subject ${finalSubject}`);
  
      // Update sheet
      const rowsToUpdate = JSON.parse(emailDetails.rowsToUpdateJson);
      if (rowsToUpdate.length > 0) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
        
        if (!sheet) {
          throw new Error(`Sheet ${CONFIG.SHEET_NAME} not found during update phase`);
        }

        const headerRowValues = sheet.getRange(CONFIG.HEADER_ROWS, 1, 1, sheet.getLastColumn()).getValues()[0];
        function getColumnIndex(columnName, headerArray) {
          const index = headerArray.indexOf(columnName);
          if (index === -1) {
            Logger.log(`Column "${columnName}" not found in header row. This column won't be updated.`);
            return -1;
          }
          return index;
        }

        const statusColIdx = CONFIG.STATUS_COL_NAME ? getColumnIndex(CONFIG.STATUS_COL_NAME, headerRowValues) : -1;
        const checkboxColIdx = getColumnIndex(CONFIG.CHECKBOX_COL_NAME, headerRowValues);
        if (checkboxColIdx === -1) {
            Logger.log(`Critical column "${CONFIG.CHECKBOX_COL_NAME}" not found. Cannot uncheck boxes.`);
        } else {
            const now = new Date();
            const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
            for (const rowNum of rowsToUpdate) {
              try {
                if (statusColIdx !== -1) {
                  sheet.getRange(rowNum, statusColIdx + 1).setValue(`Draft created ${timestamp} (From: ${emailDetails.fromEmail})`);
                }
                sheet.getRange(rowNum, checkboxColIdx + 1).setValue(false);
              } catch (e) {
                Logger.log(`Error updating row ${rowNum} after sending email: ${e.message}`);
              }
            }
        }
      }
  
      // Get the draft URL
      const draftId = draft.getId();
      const draftUrl = `https://mail.google.com/mail/u/0/#drafts/${draftId}`;
  
      return {
        success: true,
        url: draftUrl,
        message: `Draft for ${emailDetails.recipient} created successfully from ${emailDetails.fromEmail}!` // Updated message
      };
    } catch (error) {
      Logger.log(`Error in createAndOpenDraft: ${error.message}\nStack: ${error.stack}`);
      return {
        success: false,
        message: error.message
      };
    }
  }