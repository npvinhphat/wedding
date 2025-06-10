/*** @OnlyCurrentDoc */

// Global variables
const SHEET_NAME = "RSVP Responses";
const EMAILS = ["npvinhphat@gmail.com", "maitnp93@gmail.com"];
const SENDER_NAME = "Wedding RSVP System";

/**
 * Process the POST request when a form is submitted
 */
function doPost(e) {
  // Process the form submission with post method
  var result = processForm(e.parameter);
  // Return JSON result
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Process the GET request (for testing)
 */
function doGet(e) {
  return ContentService.createTextOutput("The RSVP system is working. Please submit the form from the wedding website.");
}

/**
 * Process the form data and store it in a spreadsheet
 * @param {Object} formData - The form data submitted
 * @return {Object} - Result of the operation
 */
function processForm(formData) {
  try {
    // Get active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        "Name",
        "Attending",
        "Location",
        "Guest Count",
        "Bringing Kids",
        "Kids Count",
        "Message",
        "User ID",
        "Submission Count",
        "Last Update"
      ]);
    }
    
    // Check if the user already submitted
    var userId = formData.userId || "";
    var existingRow = findUserRow(sheet, userId);
    
    // Current date
    var timestamp = new Date().toISOString();
    
    // Prepare data
    var rowData = [
      formData.name || "",
      formData.attending || "No",
      formData.location || "",
      formData.guestCount || "0",
      formData.bringingKids || "No",
      formData.kidsCount || "0",
      formData.message || "",
      userId,
      "1",  // Submission count (will be incremented for existing users)
      timestamp
    ];
    
    // If user has already submitted, update their row
    if (existingRow > 0) {
      // Get submission count from the 9th column (Submission Count)
      var submissionCount = parseInt(sheet.getRange(existingRow, 9).getValue()) || 0;
      submissionCount++;
      rowData[9] = submissionCount.toString();
      
      // Update row
      sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
      
      // Send email notification about the update
      sendUpdateEmail(formData, submissionCount);
      
      return {
        result: "success",
        message: "RSVP updated successfully",
        isUpdate: true
      };
    } 
    // Otherwise, add a new row
    else {
      sheet.appendRow(rowData);
      
      // Send email notification about the new submission
      sendNewSubmissionEmail(formData);
      
      return {
        result: "success",
        message: "RSVP submitted successfully",
        isUpdate: false
      };
    }
  } catch (error) {
    return {
      result: "error",
      message: error.toString()
    };
  }
}

/**
 * Find if a user already exists in the spreadsheet by userId
 * @param {Sheet} sheet - The spreadsheet sheet
 * @param {string} userId - The userId to look for
 * @return {number} - Row number if found, 0 if not found
 */
function findUserRow(sheet, userId) {
  if (!userId) return 0;
  
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  // Start from 1 to skip header row
  for (var i = 1; i < values.length; i++) {
    // Column structure: Name[0], Attending[1], Location[2], Guest Count[3], 
    // Bringing Kids[4], Kids Count[5], Message[6], User ID[7], Submission Count[8], Last Update[9]
    // Check column index 7 (8th column) for userId
    if (values[i][7] === userId) {
      return i + 1; // +1 because sheet rows are 1-indexed
    }
  }
  
  return 0;
}

/**
 * Send email notification about new submissions
 * @param {Object} formData - The form data submitted
 */
function sendNewSubmissionEmail(formData) {
  try {
    var subject = "New Wedding RSVP Submission: " + formData.name;
    
    var body = "A new RSVP has been submitted:\n\n" +
      "Name: " + formData.name + "\n" +
      "Attending: " + formData.attending + "\n";
      
    if (formData.attending === "Yes") {
      body += "Location: " + formData.location + "\n" +
        "Guest Count: " + formData.guestCount + "\n" +
        "Bringing Kids: " + formData.bringingKids + "\n";
        
      if (formData.bringingKids === "Yes") {
        body += "Kids Count: " + formData.kidsCount + "\n";
      }
    }
    
    body += "Message: " + formData.message + "\n";
    body += "User ID: " + formData.userId + "\n";
    body += "Submitted At: " + new Date().toLocaleString() + "\n\n";
    body += "View all responses: https://docs.google.com/spreadsheets/d/1XDAKnZhTZ9Z0aa6lC_WC9SSsmd98vY91fhjWOuuk2HE/edit?gid=1755333359#gid=1755333359";
    
    if (EMAILS && EMAILS.length > 0) {
      // Send to all email recipients
      EMAILS.forEach(email => {
        GmailApp.sendEmail(email, subject, body, {
          name: SENDER_NAME
        });
      });
    }
  } catch (e) {
    console.error("Error sending email: " + e.toString());
  }
}

/**
 * Send email notification about updated submissions
 * @param {Object} formData - The form data submitted
 * @param {number} submissionCount - Number of times user has submitted
 */
function sendUpdateEmail(formData, submissionCount) {
  try {
    var subject = "Updated Wedding RSVP: " + formData.name;
    
    var body = "An RSVP has been updated (submission #" + submissionCount + "):\n\n" +
      "Name: " + formData.name + "\n" +
      "Attending: " + formData.attending + "\n";
      
    if (formData.attending === "Yes") {
      body += "Location: " + formData.location + "\n" +
        "Guest Count: " + formData.guestCount + "\n" +
        "Bringing Kids: " + formData.bringingKids + "\n";
        
      if (formData.bringingKids === "Yes") {
        body += "Kids Count: " + formData.kidsCount + "\n";
      }
    }
    
    body += "Message: " + formData.message + "\n";
    body += "User ID: " + formData.userId + "\n";
    body += "Updated At: " + new Date().toLocaleString() + "\n\n";
    body += "View all responses: https://docs.google.com/spreadsheets/d/1XDAKnZhTZ9Z0aa6lC_WC9SSsmd98vY91fhjWOuuk2HE/edit?gid=1755333359#gid=1755333359";
    
    if (EMAILS && EMAILS.length > 0) {
      // Send to all email recipients
      EMAILS.forEach(email => {
        GmailApp.sendEmail(email, subject, body, {
          name: SENDER_NAME
        });
      });
    }
  } catch (e) {
    console.error("Error sending email: " + e.toString());
  }
}
