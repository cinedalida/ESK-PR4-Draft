// ============================================================================
// TYPEFORM TO GOOGLE SHEETS AUTOMATION SYSTEM
// ============================================================================

// Configuration Constants
const CONFIG = {
  TYPEFORM_API_BASE: "https://api.typeform.com",
  TYPEFORM_TOKEN:
    PropertiesService.getScriptProperties().getProperty("TYPEFORM_TOKEN"),
  MEGA_SHEET_NAME: "Sprint Survey Master Database",
  LOG_SHEET_NAME: "System Logs",
  FOLDER_NAME: "EOS Survey Data Sheets",
  MAX_RETRIES: 3,
  BATCH_SIZE: 100,
  // Pre-configured survey information
  KNOWN_SURVEYS: [
    {
      id: "<TYPEFORM_FORM_ID_1>",
      name: "Intro to Machine Learning",
      url: "https://form.typeform.com/to/<TYPEFORM_FORM_ID_1>",
    },
    {
      id: "<TYPEFORM_FORM_ID_2>",
      name: "Data Engineering 101",
      url: "https://form.typeform.com/to/<TYPEFORM_FORM_ID_2>",
    },
    {
      id: "<TYPEFORM_FORM_ID_3>",
      name: "Natural Language Processing 1",
      url: "https://form.typeform.com/to/<TYPEFORM_FORM_ID_3>",
    },
  ],
};

// ============================================================================
// MAIN ORCHESTRATION FUNCTIONS
// ============================================================================

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================
// Initialize the automation system with pre-configured settings
function initializeSystem() {
  try {
    Logger.log("Initializing Typeform Automation System...");

    // Set the API token if not already set
    const currentToken =
      PropertiesService.getScriptProperties().getProperty("TYPEFORM_TOKEN");
    if (!currentToken) {
      PropertiesService.getScriptProperties().setProperty(
        "TYPEFORM_TOKEN",
        "YOUR_TYPEFORM_API_TOKEN"
      );
      Logger.log("API token configured automatically");
    }

    // Create mega database sheet if it doesn't exist
    const megaSheet = getMegaSheet();
    if (!megaSheet) {
      createMegaSheet();
    }

    // Create log sheet if it doesn't exist
    const logSheet = getLogSheet();
    if (!logSheet) {
      createLogSheet();
    }

    // Create folder for data sheets
    createDataFolder();

    // Set up custom menu
    createCustomMenu();

    // Log known surveys
    logActivity(
      "INFO",
      `System initialized with ${CONFIG.KNOWN_SURVEYS.length} pre-configured surveys`
    );
    CONFIG.KNOWN_SURVEYS.forEach((survey) => {
      logActivity("INFO", `Survey: ${survey.name} (ID: ${survey.id})`);
    });

    Logger.log("System initialization complete!");
    return { success: true, message: "System initialized successfully" };
  } catch (error) {
    Logger.log("Initialization error: " + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Main function to process all surveys
 */
function processAllSurveys() {
  try {
    logActivity("INFO", "Starting bulk survey processing");

    const forms = getTypeforms();
    const results = [];

    for (const form of forms) {
      const result = processSingleSurvey(form.id, form.title);
      results.push(result);

      // Brief pause between requests to respect API limits
      Utilities.sleep(1000);
    }

    updateMegaSheetStatus();
    logActivity("INFO", `Completed processing ${results.length} surveys`);

    return results;
  } catch (error) {
    logActivity("ERROR", `Bulk processing failed: ${error.toString()}`);
    throw error;
  }
}

/**
 * Process a single survey by ID with enhanced name detection
 */
function processSingleSurvey(formId, sprintName = null) {
  let retryCount = 0;

  // Try to get survey name from known surveys if not provided
  if (!sprintName) {
    const knownSurvey = CONFIG.KNOWN_SURVEYS.find(
      (survey) => survey.id === formId
    );
    if (knownSurvey) {
      sprintName = knownSurvey.name;
      logActivity(
        "INFO",
        `Using pre-configured name: ${sprintName} for form ${formId}`
      );
    }
  }

  while (retryCount < CONFIG.MAX_RETRIES) {
    try {
      logActivity(
        "INFO",
        `Processing survey ${formId} (attempt ${retryCount + 1})`
      );

      // Update status to "In Progress"
      updateSurveyStatus(formId, "In Progress");

      // Get survey metadata
      const formData = getTypeformData(formId);
      const surveyTitle = sprintName || formData.title || `Survey_${formId}`;

      // Fetch responses
      const responses = getTypeformResponses(formId);

      if (responses.length === 0) {
        logActivity("WARNING", `No responses found for survey ${formId}`);
        updateSurveyStatus(formId, "Complete");
        return { formId, success: true, message: "No responses to process" };
      }

      // Process and clean data
      const cleanedData = processResponseData(responses, formData.fields);

      // Create or update individual sheet
      const sheetUrl = createOrUpdateDataSheet(
        surveyTitle,
        cleanedData,
        formId
      );

      // Update mega database
      updateMegaDatabase(surveyTitle, formId, sheetUrl, responses.length);

      // Update status to "Complete"
      updateSurveyStatus(formId, "Complete");

      logActivity("SUCCESS", `Successfully processed survey ${formId}`);
      return { formId, success: true, sheetUrl };
    } catch (error) {
      retryCount++;
      logActivity(
        "ERROR",
        `Attempt ${retryCount} failed for survey ${formId}: ${error.toString()}`
      );

      if (retryCount >= CONFIG.MAX_RETRIES) {
        updateSurveyStatus(formId, "Failed");
        throw new Error(
          `Failed to process survey ${formId} after ${
            CONFIG.MAX_RETRIES
          } attempts: ${error.toString()}`
        );
      }

      // Exponential backoff
      Utilities.sleep(Math.pow(2, retryCount) * 1000);
    }
  }
}

// ============================================================================
// TYPEFORM API INTEGRATION
// ============================================================================

/**
 * Get list of all Typeforms
 */
function getTypeforms() {
  const url = `${CONFIG.TYPEFORM_API_BASE}/forms`;
  const options = {
    method: "GET",
    headers: {
      Authorization: `Bearer ${CONFIG.TYPEFORM_TOKEN}`,
      "Content-Type": "application/json",
    },
  };

  const response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() !== 200) {
    throw new Error(
      `API Error: ${response.getResponseCode()} - ${response.getContentText()}`
    );
  }

  const data = JSON.parse(response.getContentText());
  return data.items || [];
}

/**
 * Get Typeform metadata
 */
function getTypeformData(formId) {
  const url = `${CONFIG.TYPEFORM_API_BASE}/forms/${formId}`;
  const options = {
    method: "GET",
    headers: {
      Authorization: `Bearer ${CONFIG.TYPEFORM_TOKEN}`,
      "Content-Type": "application/json",
    },
  };

  const response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() !== 200) {
    throw new Error(
      `API Error: ${response.getResponseCode()} - ${response.getContentText()}`
    );
  }

  return JSON.parse(response.getContentText());
}

/**
 * Get Typeform responses
 */
function getTypeformResponses(formId, pageSize = CONFIG.BATCH_SIZE) {
  let allResponses = [];
  let before = null;

  do {
    const url = `${CONFIG.TYPEFORM_API_BASE}/forms/${formId}/responses`;
    const params = [`page_size=${pageSize}`];

    if (before) {
      params.push(`before=${before}`);
    }

    const fullUrl = `${url}?${params.join("&")}`;

    const options = {
      method: "GET",
      headers: {
        Authorization: `Bearer ${CONFIG.TYPEFORM_TOKEN}`,
        "Content-Type": "application/json",
      },
    };

    const response = UrlFetchApp.fetch(fullUrl, options);

    if (response.getResponseCode() !== 200) {
      throw new Error(
        `API Error: ${response.getResponseCode()} - ${response.getContentText()}`
      );
    }

    const data = JSON.parse(response.getContentText());
    const responses = data.items || [];

    allResponses = allResponses.concat(responses);

    // Set up for next page if there are more responses
    before =
      responses.length > 0 ? responses[responses.length - 1].token : null;

    // Break if we got fewer responses than requested (last page)
    if (responses.length < pageSize) {
      break;
    }

    // Brief pause between API calls
    Utilities.sleep(500);
  } while (before);

  return allResponses;
}

// ============================================================================
// DATA PROCESSING FUNCTIONS
// ============================================================================

/**
 * Process and clean response data
 */
function processResponseData(responses, fields) {
  if (!responses || responses.length === 0) {
    return [];
  }

  // Create field mapping for easier access
  const fieldMap = {};
  if (fields) {
    fields.forEach((field) => {
      fieldMap[field.id] = {
        title: field.title || field.id,
        type: field.type,
      };
    });
  }

  const processedData = [];

  responses.forEach((response, index) => {
    try {
      const row = {
        "Response ID": response.token,
        "Submitted At": formatDate(response.submitted_at),
        "Response Index": index + 1,
      };

      // Process answers
      if (response.answers) {
        response.answers.forEach((answer) => {
          const fieldInfo = fieldMap[answer.field.id] || {
            title: answer.field.id,
          };
          const fieldTitle = cleanFieldTitle(fieldInfo.title);

          // Extract answer value based on type
          row[fieldTitle] = extractAnswerValue(answer, fieldInfo.type);
        });
      }

      processedData.push(row);
    } catch (error) {
      logActivity(
        "ERROR",
        `Error processing response ${index}: ${error.toString()}`
      );
      // Continue processing other responses
    }
  });

  return processedData;
}

/**
 * Extract answer value from response
 */
function extractAnswerValue(answer, fieldType) {
  if (!answer) return "";

  try {
    // Handle different answer types
    if (answer.text !== undefined) return cleanText(answer.text);
    if (answer.email !== undefined) return answer.email;
    if (answer.url !== undefined) return answer.url;
    if (answer.number !== undefined) return answer.number;
    if (answer.boolean !== undefined) return answer.boolean;
    if (answer.date !== undefined) return formatDate(answer.date);

    // Handle choice-based answers
    if (answer.choice) {
      if (answer.choice.label) return cleanText(answer.choice.label);
      if (answer.choice.other) return cleanText(answer.choice.other);
    }

    // Handle multiple choices
    if (answer.choices) {
      const labels = answer.choices.map(
        (choice) => choice.label || choice.other || "Unknown"
      );
      return labels.join(", ");
    }

    // Handle file uploads
    if (answer.file_url) return answer.file_url;

    // Handle payment
    if (answer.payment)
      return `${answer.payment.amount} ${answer.payment.currency}`;

    return "";
  } catch (error) {
    logActivity(
      "WARNING",
      `Error extracting answer value: ${error.toString()}`
    );
    return "";
  }
}

/**
 * Clean and standardize text
 */
function cleanText(text) {
  if (!text) return "";

  return text
    .toString()
    .trim()
    .replace(/\s+/g, " ")
    .replace(/[\r\n]+/g, " | ");
}

/**
 * Clean field titles for column headers
 */
function cleanFieldTitle(title) {
  if (!title) return "Unknown Field";

  return title
    .toString()
    .trim()
    .replace(/[^\w\s-]/g, "")
    .replace(/\s+/g, " ")
    .substring(0, 100); // Limit length
}

/**
 * Format date consistently
 */
function formatDate(dateString) {
  if (!dateString) return "";

  try {
    const date = new Date(dateString);
    return date.toLocaleString();
  } catch (error) {
    return dateString;
  }
}

// ============================================================================
// GOOGLE SHEETS MANAGEMENT
// ============================================================================

/**
 * Get or create mega database sheet
 */
function getMegaSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(CONFIG.MEGA_SHEET_NAME);
}

/**
 * Create mega database sheet
 */
function createMegaSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.insertSheet(CONFIG.MEGA_SHEET_NAME);

  // Set up headers
  const headers = [
    "Sprint Name",
    "Form ID",
    "Source Sheet Link",
    "Response Count",
    "Date Created",
    "Last Updated",
    "Processing Status",
    "Action",
  ];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#4285F4");
  headerRange.setFontColor("white");

  // Format columns
  sheet.setColumnWidth(1, 200); // Sprint Name
  sheet.setColumnWidth(2, 150); // Form ID
  sheet.setColumnWidth(3, 300); // Source Sheet Link
  sheet.setColumnWidth(4, 120); // Response Count
  sheet.setColumnWidth(5, 150); // Date Created
  sheet.setColumnWidth(6, 150); // Last Updated
  sheet.setColumnWidth(7, 120); // Processing Status
  sheet.setColumnWidth(8, 100); // Action

  return sheet;
}

/**
 * Get or create log sheet
 */
function getLogSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(CONFIG.LOG_SHEET_NAME);
}

/**
 * Create log sheet
 */
function createLogSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.insertSheet(CONFIG.LOG_SHEET_NAME);

  const headers = ["Timestamp", "Level", "Message"];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#34A853");
  headerRange.setFontColor("white");

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 500);

  return sheet;
}

/**
 * Create or update individual data sheet
 */
function createOrUpdateDataSheet(sprintName, data, formId) {
  try {
    const folder = getOrCreateDataFolder();
    const fileName = `${sprintName}_Data`;

    // Check if spreadsheet already exists
    const files = folder.getFilesByName(fileName);
    let spreadsheet;

    if (files.hasNext()) {
      // Update existing spreadsheet
      const file = files.next();
      spreadsheet = SpreadsheetApp.openById(file.getId());
      const sheet = spreadsheet.getActiveSheet();
      sheet.clear();
    } else {
      // Create new spreadsheet
      spreadsheet = SpreadsheetApp.create(fileName);
      const file = DriveApp.getFileById(spreadsheet.getId());
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    }

    const sheet = spreadsheet.getActiveSheet();
    sheet.setName("Survey Data");

    if (data && data.length > 0) {
      // Get all unique column headers
      const allHeaders = new Set();
      data.forEach((row) => {
        Object.keys(row).forEach((key) => allHeaders.add(key));
      });

      const headers = Array.from(allHeaders);

      // Create header row
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      sheet.getRange(1, 1, 1, headers.length).setBackground("#FF9900");
      sheet.getRange(1, 1, 1, headers.length).setFontColor("white");

      // Add data rows
      const rows = data.map((row) =>
        headers.map((header) => row[header] || "")
      );

      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      }

      // Auto-resize columns
      sheet.autoResizeColumns(1, headers.length);
    }

    // Add metadata sheet
    addMetadataSheet(spreadsheet, formId, data.length);

    return spreadsheet.getUrl();
  } catch (error) {
    logActivity(
      "ERROR",
      `Error creating data sheet for ${sprintName}: ${error.toString()}`
    );
    throw error;
  }
}

/**
 * Add metadata sheet to data spreadsheet
 */
function addMetadataSheet(spreadsheet, formId, responseCount) {
  let metaSheet = spreadsheet.getSheetByName("Metadata");

  if (!metaSheet) {
    metaSheet = spreadsheet.insertSheet("Metadata");
  } else {
    metaSheet.clear();
  }

  const metadata = [
    ["Property", "Value"],
    ["Form ID", formId],
    ["Response Count", responseCount],
    ["Last Updated", new Date().toLocaleString()],
    ["Processed By", "Typeform Automation System"],
  ];

  metaSheet.getRange(1, 1, metadata.length, 2).setValues(metadata);
  metaSheet.getRange(1, 1, 1, 2).setFontWeight("bold");
  metaSheet.autoResizeColumns(1, 2);
}

/**
 * Get or create data folder
 */
function getOrCreateDataFolder() {
  const folders = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(CONFIG.FOLDER_NAME);
  }
}

/**
 * Create data folder (for initialization)
 */
function createDataFolder() {
  return getOrCreateDataFolder();
}

/**
 * Update mega database with survey information
 */
function updateMegaDatabase(sprintName, formId, sheetUrl, responseCount) {
  const sheet = getMegaSheet();
  if (!sheet) throw new Error("Mega database sheet not found");

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find existing row or create new one
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === formId) {
      // Form ID column
      rowIndex = i;
      break;
    }
  }

  const now = new Date().toLocaleString();
  const rowData = [
    sprintName,
    formId,
    sheetUrl,
    responseCount,
    rowIndex === -1 ? now : data[rowIndex][4], // Keep original creation date
    now,
    "Complete",
    "Update",
  ];

  if (rowIndex === -1) {
    // Add new row
    sheet.appendRow(rowData);
    rowIndex = sheet.getLastRow();
  } else {
    // Update existing row
    sheet.getRange(rowIndex + 1, 1, 1, rowData.length).setValues([rowData]);
  }

  // Format the new/updated row
  formatMegaSheetRow(sheet, rowIndex + 1);
}

/**
 * Format mega sheet row
 */
function formatMegaSheetRow(sheet, rowNum) {
  const range = sheet.getRange(rowNum, 1, 1, 8);

  // Status color coding
  const statusCell = sheet.getRange(rowNum, 7);
  const status = statusCell.getValue();

  switch (status) {
    case "Complete":
      statusCell.setBackground("#D9EAD3");
      break;
    case "In Progress":
      statusCell.setBackground("#FFF2CC");
      break;
    case "Failed":
      statusCell.setBackground("#F4CCCC");
      break;
  }

  // Make sheet link clickable
  const linkCell = sheet.getRange(rowNum, 3);
  const url = linkCell.getValue();
  if (url && url.includes("spreadsheet")) {
    linkCell.setFormula(`=HYPERLINK("${url}", "Open Sheet")`);
  }
}

/**
 * Update survey processing status
 */
function updateSurveyStatus(formId, status) {
  const sheet = getMegaSheet();
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === formId) {
      sheet.getRange(i + 1, 7).setValue(status);
      sheet.getRange(i + 1, 6).setValue(new Date().toLocaleString());
      formatMegaSheetRow(sheet, i + 1);
      break;
    }
  }
}

/**
 * Update mega sheet status summary
 */
function updateMegaSheetStatus() {
  // This function can be expanded to add summary statistics
  logActivity("INFO", "Mega sheet status updated");
}

// ============================================================================
// LOGGING AND ERROR HANDLING
// ============================================================================

/**
 * Log activity to the log sheet
 */
function logActivity(level, message) {
  try {
    const sheet = getLogSheet();
    if (sheet) {
      const timestamp = new Date().toLocaleString();
      sheet.appendRow([timestamp, level, message]);

      // Keep only last 1000 entries
      if (sheet.getLastRow() > 1001) {
        sheet.deleteRow(2);
      }
    }

    // Also log to Apps Script console
    Logger.log(`[${level}] ${message}`);
  } catch (error) {
    Logger.log("Logging error: " + error.toString());
  }
}

// ============================================================================
// UI AND MENU FUNCTIONS
// ============================================================================

/**
 * Create custom menu with your specific options
 */
function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🔄 Typeform Automation")
    .addItem("🚀 Initialize System", "initializeSystem")
    .addSeparator()
    .addItem("📊 Process All Surveys", "processAllSurveys")
    .addItem("🎯 Process Your 3 Surveys", "processYourSurveys")
    .addItem("🔍 Process Single Survey", "showSingleSurveyDialog")
    .addSeparator()
    .addItem("⚡ Quick Setup & Test", "quickSetupYourSurveys")
    .addSeparator()
    .addItem("📋 View System Logs", "showLogsSheet")
    .addItem("🔧 System Settings", "showSettingsDialog")
    .addItem("ℹ️ About", "showAboutDialog")
    .addToUi();
}

/**
 * Show single survey processing dialog with your survey suggestions
 */
function showSingleSurveyDialog() {
  const ui = SpreadsheetApp.getUi();

  const suggestions = CONFIG.KNOWN_SURVEYS.map(
    (s) => `${s.name}: ${s.id}`
  ).join("\n• ");

  const result = ui.prompt(
    "Process Single Survey",
    `Enter Typeform ID or URL:\n\n💡 Your surveys:\n• ${suggestions}\n\nEnter ID or full URL:`,
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const input = result.getResponseText().trim();
    const formId = extractFormId(input);

    if (formId) {
      try {
        processSingleSurvey(formId);
        ui.alert(
          "Success",
          `Survey ${formId} processed successfully!`,
          ui.ButtonSet.OK
        );
      } catch (error) {
        ui.alert(
          "Error",
          `Failed to process survey: ${error.toString()}`,
          ui.ButtonSet.OK
        );
      }
    } else {
      ui.alert(
        "Invalid Input",
        "Please provide a valid Typeform ID or URL.",
        ui.ButtonSet.OK
      );
    }
  }
}

/**
 * Extract form ID from URL or return as-is if already an ID
 */
function extractFormId(input) {
  if (!input) return null;

  // If it's already a form ID (alphanumeric string)
  if (/^[a-zA-Z0-9]+$/.test(input)) {
    return input;
  }

  // Extract from Typeform URL
  const urlMatch = input.match(/typeform\.com\/to\/([a-zA-Z0-9]+)/);
  if (urlMatch) {
    return urlMatch[1];
  }

  return null;
}

/**
 * Show logs sheet
 */
function showLogsSheet() {
  const sheet = getLogSheet();
  if (sheet) {
    SpreadsheetApp.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert(
      "Error",
      "Log sheet not found. Please initialize the system first.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Show settings dialog with current token pre-populated
 */
function showSettingsDialog() {
  const ui = SpreadsheetApp.getUi();
  const currentToken =
    PropertiesService.getScriptProperties().getProperty("TYPEFORM_TOKEN") || "";

  // Show current token status
  let message = "Current API Token Status:\n";
  if (currentToken) {
    message += `✅ Token configured: ${currentToken.substring(
      0,
      10
    )}...${currentToken.substring(currentToken.length - 4)}\n\n`;
    message += "Enter new token to update, or click Cancel to keep current:";
  } else {
    message += "❌ No token configured\n\n";
    message += "Please enter your Typeform API Token:";
  }

  const result = ui.prompt("System Settings", message, ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() === ui.Button.OK) {
    const token = result.getResponseText().trim();
    if (token) {
      PropertiesService.getScriptProperties().setProperty(
        "TYPEFORM_TOKEN",
        token
      );
      ui.alert("Success", "API token updated successfully!", ui.ButtonSet.OK);
    }
  }
}

/**
 * Show about dialog with your specific survey information
 */
function showAboutDialog() {
  const ui = SpreadsheetApp.getUi();

  const message = `
Typeform to Google Sheets Automation System
Version 1.0

🎯 Your Pre-configured Surveys:
• Intro to Machine Learning (ID: <TYPEFORM_FORM_ID_1>)
• Data Engineering 101 (ID: <TYPEFORM_FORM_ID_1>)  
• Natural Language Processing 1 (ID: <TYPEFORM_FORM_ID_3>)

✨ Features:
• Automated data retrieval from Typeform
• Data cleaning and standardization
• Centralized tracking dashboard
• Individual data sheets for each survey
• Comprehensive logging and error handling

🚀 Quick Start:
1. System is pre-configured with your API token
2. Use 📊 Process All Surveys to get all 3 (pre-made) surveys
3. Or use 🔍 Process Single Survey with IDs above
4. Check Master Database for results and links

🔧 Your API Token: <YOUR_TYPEFORM_API_TOKEN>

For support or questions, check System Logs for detailed information.
  `;

  ui.alert("About Typeform Automation", message, ui.ButtonSet.OK);
}

// ============================================================================
// TRIGGER MANAGEMENT
// ============================================================================

/**
 * Create scheduled triggers
 */
function createScheduledTriggers() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === "scheduledProcessing") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new daily trigger
  ScriptApp.newTrigger("scheduledProcessing")
    .timeBased()
    .everyDays(1)
    .atHour(9) // 9 AM
    .create();

  logActivity("INFO", "Scheduled triggers created");
}

/**
 * Scheduled processing function
 */
function scheduledProcessing() {
  try {
    logActivity("INFO", "Starting scheduled processing");
    processAllSurveys();
    logActivity("SUCCESS", "Scheduled processing completed");
  } catch (error) {
    logActivity("ERROR", `Scheduled processing failed: ${error.toString()}`);
  }
}

// ============================================================================
// QUICK SETUP AND TESTING FUNCTIONS
// ============================================================================

/**
 * Quick setup function for your specific surveys
 */
function quickSetupYourSurveys() {
  try {
    Logger.log("🚀 Running quick setup for your surveys...");

    // Initialize system
    initializeSystem();

    // Test API connection
    const apiTest = testApiConnection();
    if (!apiTest.success) {
      throw new Error("API connection failed: " + apiTest.error);
    }

    Logger.log(`✅ API connected successfully. Found ${apiTest.count} forms.`);

    // Process your 3 specific surveys
    const results = [];
    for (const survey of CONFIG.KNOWN_SURVEYS) {
      try {
        Logger.log(`📊 Processing: ${survey.name}...`);
        const result = processSingleSurvey(survey.id, survey.name);
        results.push({ ...result, name: survey.name });
        Logger.log(`✅ Completed: ${survey.name}`);
      } catch (error) {
        Logger.log(`❌ Failed: ${survey.name} - ${error.toString()}`);
        results.push({
          formId: survey.id,
          success: false,
          name: survey.name,
          error: error.toString(),
        });
      }
    }

    // Summary
    const successful = results.filter((r) => r.success).length;
    Logger.log(
      `🎉 Quick setup complete! ${successful}/${CONFIG.KNOWN_SURVEYS.length} surveys processed successfully.`
    );

    return results;
  } catch (error) {
    Logger.log("❌ Quick setup failed: " + error.toString());
    throw error;
  }
}

/**
 * Process just your known surveys (alternative to processAllSurveys)
 */
function processYourSurveys() {
  try {
    logActivity("INFO", "Processing your pre-configured surveys");

    const results = [];

    for (const survey of CONFIG.KNOWN_SURVEYS) {
      const result = processSingleSurvey(survey.id, survey.name);
      results.push(result);

      // Brief pause between requests
      Utilities.sleep(1000);
    }

    updateMegaSheetStatus();
    logActivity("INFO", `Completed processing ${results.length} surveys`);

    return results;
  } catch (error) {
    logActivity("ERROR", `Processing failed: ${error.toString()}`);
    throw error;
  }
}

/**
 * Test API connection
 */
function testApiConnection() {
  try {
    const forms = getTypeforms();
    logActivity(
      "SUCCESS",
      `API connection successful. Found ${forms.length} forms.`
    );
    return { success: true, count: forms.length };
  } catch (error) {
    logActivity("ERROR", `API connection failed: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * Manual trigger for testing
 */
function runManualTest() {
  try {
    // Test API connection
    const apiTest = testApiConnection();
    if (!apiTest.success) {
      throw new Error("API connection failed: " + apiTest.error);
    }

    // Get first form for testing
    const forms = getTypeforms();
    if (forms.length === 0) {
      throw new Error("No forms found");
    }

    // Process first form
    const result = processSingleSurvey(forms[0].id, forms[0].title);

    logActivity("SUCCESS", "Manual test completed successfully");
    return result;
  } catch (error) {
    logActivity("ERROR", `Manual test failed: ${error.toString()}`);
    throw error;
  }
}

// ============================================================================
// INSTALLATION AND SETUP
// ============================================================================

/**
 * One-time setup function
 */
function onInstall() {
  initializeSystem();
  createCustomMenu();
}

/**
 * On spreadsheet open
 */
function onOpen() {
  createCustomMenu();
}
