/**
 * Email-to-Spreadsheet Logger
 * 
 * Automatically log Gmail messages to a Google Sheet based on filter rules.
 * 
 * Setup: Click Extensions → Apps Script, paste this code, save, then run setupTemplate().
 */

// Configuration Constants
const SHEET_NAMES = {
  SETTINGS: 'Settings',
  RULES: 'Rules',
  CATEGORIES: 'Categories',
  LOG: 'Log'
};

const LABEL_NAME = 'EmailLogger/Processed';

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📧 Email Logger')
    .addItem('⚙️ Setup Template', 'setupTemplate')
    .addSeparator()
    .addItem('▶️ Run Logger Now', 'runLogger')
    .addItem('⏸️ Stop Auto-Logging', 'deleteTriggers')
    .addSeparator()
    .addItem('📋 Check Status', 'showStatus')
    .addToUi();
}

/**
 * One-click setup: Creates all sheets with headers, formatting, and sample data
 * Idempotent: safe to run multiple times
 */
function setupTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Create or get sheets
  createSettingsSheet(ss);
  createRulesSheet(ss);
  createCategoriesSheet(ss);
  createLogSheet(ss);
  
  // Set up time-driven trigger
  setupTrigger();
  
  ui.alert('✅ Setup Complete', 
    'Template created!\n\n' +
    'Next steps:\n' +
    '1. Configure your Gmail search rules in the "Rules" tab\n' +
    '2. Set your auto-categorization keywords in "Categories"\n' +
    '3. Run "▶️ Run Logger Now" to test, or wait for auto-run\n\n' +
    'The logger will run automatically every 15 minutes.', 
    ui.ButtonSet.OK);
}

/**
 * Creates Settings sheet with configuration options
 */
function createSettingsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  
  if (sheet) {
    // Sheet exists, skip
    return;
  }
  
  sheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
  
  // Headers
  sheet.getRange('A1:B1').setValues([['Setting', 'Value']]);
  sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  
  // Settings data
  const settings = [
    ['Polling Interval (minutes)', '15'],
    ['Max Emails Per Run', '100'],
    ['Log Body Content', 'No'],
    ['Body Max Characters', '500'],
    ['Label for Processed Emails', LABEL_NAME],
    ['Last Run Timestamp', 'Never']
  ];
  
  sheet.getRange(2, 1, settings.length, 2).setValues(settings);
  
  // Formatting
  sheet.getRange('A:A').setFontWeight('bold');
  sheet.autoResizeColumns(1, 2);
  sheet.setColumnWidth(1, 30);
  sheet.setColumnWidth(2, 30);
  
  // Add note
  sheet.getRange('A1').setNote('Configure how the email logger behaves');
}

/**
 * Creates Rules sheet for defining Gmail search queries
 */
function createRulesSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.RULES);
  
  if (sheet) {
    return;
  }
  
  sheet = ss.insertSheet(SHEET_NAMES.RULES);
  
  // Headers
  const headers = ['Enabled', 'Rule Name', 'Gmail Search Query', 'Description'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  
  // Sample rules
  const sampleRules = [
    ['Yes', 'Stripe Receipts', 'from:stripe.com subject:receipt', 'Payment receipts from Stripe'],
    ['Yes', 'Client Emails', 'from:client@example.com OR from:another@client.com', 'Emails from important clients'],
    ['No', 'Newsletters', 'label:Newsletter', 'Weekly newsletters (disabled by default)']
  ];
  
  sheet.getRange(2, 1, sampleRules.length, sampleRules[0].length).setValues(sampleRules);
  
  // Data validation for Enabled column
  const enabledValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'])
    .build();
  sheet.getRange('A2:A').setDataValidation(enabledValidation);
  
  // Formatting
  sheet.autoResizeColumns(1, headers.length);
  sheet.setColumnWidth(1, 10);
  sheet.setColumnWidth(2, 20);
  sheet.setColumnWidth(3, 40);
  sheet.setColumnWidth(4, 30);
  
  // Conditional formatting for Enabled column
  const yesRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Yes')
    .setBackground('#d9ead3')
    .setRanges([sheet.getRange('A2:A')])
    .build();
  const noRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('No')
    .setBackground('#fce5cd')
    .setRanges([sheet.getRange('A2:A')])
    .build();
  sheet.setConditionalFormatRules([yesRule, noRule]);
}

/**
 * Creates Categories sheet for auto-categorization
 */
function createCategoriesSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.CATEGORIES);
  
  if (sheet) {
    return;
  }
  
  sheet = ss.insertSheet(SHEET_NAMES.CATEGORIES);
  
  // Headers
  const headers = ['Category Name', 'Keywords (comma-separated)', 'Priority'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  
  // Sample categories
  const sampleCategories = [
    ['Receipts', 'receipt, invoice, payment, paid, order confirmation', '1'],
    ['Client Communication', 'client, project, deliverable, milestone', '2'],
    ['Support', 'support, help, ticket, issue, bug', '3'],
    ['Newsletter', 'newsletter, digest, weekly, monthly', '4']
  ];
  
  sheet.getRange(2, 1, sampleCategories.length, sampleCategories[0].length).setValues(sampleCategories);
  
  // Formatting
  sheet.autoResizeColumns(1, headers.length);
  sheet.setColumnWidth(1, 20);
  sheet.setColumnWidth(2, 40);
  sheet.setColumnWidth(3, 10);
  
  // Note
  sheet.getRange('A1').setNote('Emails are checked against these categories in priority order. First match wins.');
}

/**
 * Creates Log sheet for storing email records
 */
function createLogSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.LOG);
  
  if (sheet) {
    return;
  }
  
  sheet = ss.insertSheet(SHEET_NAMES.LOG);
  
  // Headers
  const headers = [
    'Logged At',
    'Rule Name',
    'From',
    'To',
    'Subject',
    'Date Sent',
    'Category',
    'Snippet',
    'Labels',
    'Message ID',
    'Thread ID',
    'Gmail Link'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  
  // Formatting
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  
  // Set column widths
  const widths = [18, 20, 25, 25, 40, 18, 15, 50, 20, 25, 25, 50];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
}

/**
 * Sets up time-driven trigger based on settings
 */
function setupTrigger() {
  // Delete existing triggers first
  deleteTriggers();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  
  if (!settingsSheet) {
    return;
  }
  
  const intervalCell = settingsSheet.getRange('B2').getValue();
  const intervalMinutes = parseInt(intervalCell) || 15;
  
  // Create new trigger
  ScriptApp.newTrigger('runLogger')
    .timeBased()
    .everyMinutes(intervalMinutes)
    .create();
}

/**
 * Deletes all time-driven triggers for this script
 */
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runLogger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Main logger function - searches Gmail and logs new messages
 */
function runLogger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  const rulesSheet = ss.getSheetByName(SHEET_NAMES.RULES);
  const categoriesSheet = ss.getSheetByName(SHEET_NAMES.CATEGORIES);
  const logSheet = ss.getSheetByName(SHEET_NAMES.LOG);
  
  if (!settingsSheet || !rulesSheet || !logSheet) {
    Logger.log('Required sheets not found. Run setupTemplate() first.');
    return;
  }
  
  // Get settings
  const maxEmails = parseInt(settingsSheet.getRange('B3').getValue()) || 100;
  const logBody = settingsSheet.getRange('B4').getValue().toString().toLowerCase() === 'yes';
  const bodyMaxChars = parseInt(settingsSheet.getRange('B5').getValue()) || 500;
  
  // Get categories for auto-categorization
  const categories = getCategories(categoriesSheet);
  
  // Get enabled rules
  const rules = getEnabledRules(rulesSheet);
  
  if (rules.length === 0) {
    Logger.log('No enabled rules found.');
    return;
  }
  
  // Ensure label exists for tracking processed emails
  let processedLabel;
  try {
    processedLabel = GmailApp.getUserLabelByName(LABEL_NAME);
    if (!processedLabel) {
      processedLabel = GmailApp.createLabel(LABEL_NAME);
    }
  } catch (e) {
    Logger.log('Could not create/access label: ' + e);
    return;
  }
  
  let totalLogged = 0;
  const now = new Date();
  
  // Process each rule
  rules.forEach(rule => {
    try {
      const threads = GmailApp.search(rule.query, 0, maxEmails);
      
      threads.forEach(thread => {
        const messages = thread.getMessages();
        
        messages.forEach(message => {
          // Check if already processed
          const labels = message.getLabelIds();
          if (labels.includes(processedLabel.getId())) {
            return;
          }
          
          // Log the message
          const logRow = createLogRow(message, rule, categories, logBody, bodyMaxChars);
          logSheet.appendRow(logRow);
          
          // Mark as processed
          message.addLabel(processedLabel);
          totalLogged++;
        });
      });
      
      Logger.log(`Processed rule "${rule.name}": ${threads.length} threads found`);
      
    } catch (e) {
      Logger.log(`Error processing rule "${rule.name}": ${e}`);
    }
  });
  
  // Update last run timestamp
  settingsSheet.getRange('B7').setValue(now.toISOString());
  
  Logger.log(`Email Logger completed. Total emails logged: ${totalLogged}`);
}

/**
 * Extracts enabled rules from Rules sheet
 */
function getEnabledRules(sheet) {
  const data = sheet.getDataRange().getValues();
  const rules = [];
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === 'Yes' && row[2]) { // Enabled and has query
      rules.push({
        enabled: row[0],
        name: row[1],
        query: row[2],
        description: row[3]
      });
    }
  }
  
  return rules;
}

/**
 * Extracts categories from Categories sheet
 */
function getCategories(sheet) {
  if (!sheet) {
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const categories = [];
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] && row[1]) {
      categories.push({
        name: row[0],
        keywords: row[1].toString().split(',').map(k => k.trim().toLowerCase()),
        priority: parseInt(row[2]) || 999
      });
    }
  }
  
  // Sort by priority
  categories.sort((a, b) => a.priority - b.priority);
  
  return categories;
}

/**
 * Categorizes email based on subject and snippet
 */
function categorizeEmail(subject, snippet, categories) {
  const text = (subject + ' ' + snippet).toLowerCase();
  
  for (const category of categories) {
    for (const keyword of category.keywords) {
      if (text.includes(keyword)) {
        return category.name;
      }
    }
  }
  
  return 'Uncategorized';
}

/**
 * Creates a log row from a Gmail message
 */
function createLogRow(message, rule, categories, logBody, bodyMaxChars) {
  const subject = message.getSubject();
  const snippet = message.getPlainBody().substring(0, logBody ? bodyMaxChars : 100);
  const category = categorizeEmail(subject, snippet, categories);
  
  // Build Gmail link
  const messageId = message.getId();
  const threadId = message.getThread().getId();
  const gmailLink = `https://mail.google.com/mail/u/0/#inbox/${threadId}`;
  
  // Get labels as comma-separated string
  const labels = message.getLabels().map(l => l.getName()).join(', ');
  
  return [
    new Date().toISOString(),      // Logged At
    rule.name,                      // Rule Name
    message.getFrom(),              // From
    message.getTo(),                // To
    subject,                        // Subject
    message.getDate().toISOString(), // Date Sent
    category,                       // Category
    snippet,                        // Snippet
    labels,                         // Labels
    messageId,                      // Message ID
    threadId,                       // Thread ID
    gmailLink                       // Gmail Link
  ];
}

/**
 * Shows current status dialog
 */
function showStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  
  if (!settingsSheet) {
    SpreadsheetApp.getUi().alert('Setup Required', 'Please run "⚙️ Setup Template" first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const lastRun = settingsSheet.getRange('B7').getValue();
  const logSheet = ss.getSheetByName(SHEET_NAMES.LOG);
  const logCount = logSheet ? logSheet.getLastRow() - 1 : 0;
  
  const triggers = ScriptApp.getProjectTriggers();
  const triggerCount = triggers.filter(t => t.getHandlerFunction() === 'runLogger').length;
  
  const status = triggerCount > 0 ? '✅ Active' : '⏸️ Paused';
  
  SpreadsheetApp.getUi().alert(
    '📋 Email Logger Status',
    `${status}\n\n` +
    `Last Run: ${lastRun}\n` +
    `Emails Logged: ${logCount}\n` +
    `Auto-Run Triggers: ${triggerCount}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
