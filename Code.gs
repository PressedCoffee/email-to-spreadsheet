/**
 * Email-to-Spreadsheet Logger v2.0
 * 
 * Inbox capture layer for the Gmail Automation Suite.
 * Logs Gmail messages with enriched metadata for relationship intelligence.
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
    '3. Add your email addresses to Settings (Internal Emails)\n' +
    '4. Run "▶️ Run Logger Now" to test, or wait for auto-run\n\n' +
    'The logger will run automatically every 15 minutes.', 
    ui.ButtonSet.OK);
}

/**
 * Creates Settings sheet with configuration options
 */
function createSettingsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  
  if (sheet) {
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
    ['Last Run Timestamp', 'Never'],
    ['', ''],
    ['Internal Emails (comma-separated)', ''],
    ['Client ID Prefix', 'C'],
    ['Auto-assign Client IDs', 'Yes']
  ];
  
  sheet.getRange(2, 1, settings.length, 2).setValues(settings);
  
  // Formatting
  sheet.getRange('A:A').setFontWeight('bold');
  sheet.autoResizeColumns(1, 2);
  sheet.setColumnWidth(1, 35);
  sheet.setColumnWidth(2, 50);
  
  // Add notes
  sheet.getRange('A8').setNote('Your email addresses (used to determine direction: inbound vs outbound)');
  sheet.getRange('A9').setNote('Prefix for auto-generated Client IDs (e.g., C001, C002)');
  sheet.getRange('A10').setNote('Automatically assign Client IDs to new email addresses');
};

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
    ['Yes', 'All Inbox', 'in:inbox', 'All inbox messages'],
    ['Yes', 'All Sent', 'in:sent', 'All sent messages'],
    ['No', 'Stripe Receipts', 'from:stripe.com subject:receipt', 'Payment receipts from Stripe'],
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
  sheet.setColumnWidth(3, 50);
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
  sheet.setColumnWidth(2, 45);
  sheet.setColumnWidth(3, 10);
  
  // Note
  sheet.getRange('A1').setNote('Emails are checked against these categories in priority order. First match wins.');
}

/**
 * Creates Log sheet for storing email records with enriched metadata
 */
function createLogSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.LOG);
  
  if (sheet) {
    return;
  }
  
  sheet = ss.insertSheet(SHEET_NAMES.LOG);
  
  // Headers - enriched metadata
  const headers = [
    'Logged At',           // A - When we logged it
    'Rule Name',           // B - Which rule matched
    'From',                // C - Raw From header
    'To',                  // D - Raw To header
    'Subject',             // E - Email subject
    'Date Sent',           // F - When email was sent
    'Category',            // G - Auto-categorized
    'Snippet',             // H - Body preview
    'Labels',              // I - Gmail labels
    'Message ID',          // J - Unique message ID
    'Thread ID',           // K - Gmail thread ID
    'Gmail Link',          // L - Direct link
    // Enriched metadata
    'Primary Contact Email',  // M - Extracted primary contact
    'Direction',              // N - Inbound / Outbound / Internal
    'Domain',                 // O - Email domain
    'Participants',           // P - All participants (comma-separated)
    'Participants Count',     // Q - Number of participants
    'Is Internal',            // R - Yes/No
    'Thread Message Count',   // S - Messages in thread
    'Thread Start Date',      // T - First message date
    'Last Message In Thread', // U - Last message date
    'Last Sender',            // V - Who sent last message
    'Waiting On',             // W - Who we're waiting for reply from
    'Client ID'               // X - Auto-assigned or matched client ID
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  
  // Formatting
  sheet.setFrozenRows(1);
  
  // Set column widths
  const widths = [18, 15, 25, 25, 40, 18, 15, 40, 20, 25, 25, 45, 
                  25, 12, 20, 50, 15, 12, 18, 18, 18, 25, 25, 12];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
}

/**
 * Sets up time-driven trigger based on settings
 */
function setupTrigger() {
  deleteTriggers();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  
  if (!settingsSheet) {
    return;
  }
  
  const intervalCell = settingsSheet.getRange('B2').getValue();
  const intervalMinutes = parseInt(intervalCell) || 15;
  
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

// ============================================================================
// HELPER FUNCTIONS FOR EMAIL PARSING AND ENRICHMENT
// ============================================================================

/**
 * Extracts the primary email address from a From/To header
 * @param {string} header - Raw email header (e.g., "John Doe <john@example.com>")
 * @returns {string} - Clean email address
 */
function extractEmailAddress(header) {
  if (!header) return '';
  
  // Handle multiple recipients (To field)
  const emails = header.match(/[\w.-]+@[\w.-]+\.\w+/g);
  if (emails && emails.length > 0) {
    return emails[0].toLowerCase();
  }
  return '';
}

/**
 * Extracts all email addresses from a header
 * @param {string} header - Raw email header
 * @returns {string[]} - Array of email addresses
 */
function extractAllEmailAddresses(header) {
  if (!header) return [];
  const emails = header.match(/[\w.-]+@[\w.-]+\.\w+/g);
  return emails ? emails.map(e => e.toLowerCase()) : [];
}

/**
 * Extracts domain from an email address
 * @param {string} email - Email address
 * @returns {string} - Domain part
 */
function extractDomain(email) {
  if (!email || !email.includes('@')) return '';
  return email.split('@')[1].toLowerCase();
}

/**
 * Determines the direction of an email (Inbound, Outbound, Internal)
 * @param {string} from - From email address
 * @param {string[]} to - To email addresses
 * @param {string[]} internalEmails - List of internal email addresses
 * @returns {string} - 'Inbound', 'Outbound', or 'Internal'
 */
function determineDirection(from, to, internalEmails) {
  const fromLower = from.toLowerCase();
  const toLower = to.map(e => e.toLowerCase());
  const internalLower = internalEmails.map(e => e.toLowerCase());
  
  const fromIsInternal = internalLower.includes(fromLower);
  const toHasInternal = toLower.some(e => internalLower.includes(e));
  
  if (fromIsInternal && toHasInternal) {
    return 'Internal';
  } else if (fromIsInternal) {
    return 'Outbound';
  } else if (toHasInternal) {
    return 'Inbound';
  }
  
  // Default: if we can't determine, check if we're in the To list
  return 'Inbound';
}

/**
 * Checks if an email is internal (to/from our domain)
 * @param {string} from - From email address
 * @param {string[]} to - To email addresses
 * @param {string[]} internalEmails - List of internal email addresses
 * @returns {boolean}
 */
function isInternalEmail(from, to, internalEmails) {
  const fromLower = from.toLowerCase();
  const toLower = to.map(e => e.toLowerCase());
  const internalLower = internalEmails.map(e => e.toLowerCase());
  
  return internalLower.includes(fromLower) && toLower.every(e => internalLower.includes(e));
}

/**
 * Extracts primary contact email from a message
 * For inbound: the sender
 * For outbound: the first recipient
 * @param {string} direction - 'Inbound', 'Outbound', or 'Internal'
 * @param {string} from - From email address
 * @param {string[]} to - To email addresses
 * @param {string[]} internalEmails - List of internal email addresses
 * @returns {string} - Primary contact email
 */
function extractPrimaryContact(direction, from, to, internalEmails) {
  const internalLower = internalEmails.map(e => e.toLowerCase());
  
  if (direction === 'Inbound') {
    return from.toLowerCase();
  } else if (direction === 'Outbound') {
    // Return first non-internal recipient
    for (const email of to) {
      if (!internalLower.includes(email.toLowerCase())) {
        return email.toLowerCase();
      }
    }
    // Fallback to first To
    return to.length > 0 ? to[0].toLowerCase() : '';
  } else {
    // Internal - return first non-internal participant or empty
    for (const email of to) {
      if (!internalLower.includes(email.toLowerCase())) {
        return email.toLowerCase();
      }
    }
    return from.toLowerCase();
  }
}

/**
 * Gets all unique participants in a message
 * @param {string} from - From email address
 * @param {string[]} to - To email addresses
 * @param {string} cc - CC header (optional)
 * @returns {string[]} - Unique participants
 */
function getParticipants(from, to, cc) {
  const participants = new Set();
  
  if (from) {
    const fromEmails = extractAllEmailAddresses(from);
    fromEmails.forEach(e => participants.add(e.toLowerCase()));
  }
  
  if (to) {
    to.forEach(e => participants.add(e.toLowerCase()));
  }
  
  if (cc) {
    const ccEmails = extractAllEmailAddresses(cc);
    ccEmails.forEach(e => participants.add(e.toLowerCase()));
  }
  
  return Array.from(participants);
}

/**
 * Builds a thread summary by fetching all messages in the thread
 * @param {GmailThread} thread - Gmail thread object
 * @returns {Object} - Thread summary
 */
function buildThreadSummary(thread) {
  const messages = thread.getMessages();
  const messageCount = messages.length;
  
  if (messageCount === 0) {
    return {
      messageCount: 0,
      startDate: null,
      lastDate: null,
      lastSender: '',
      firstSender: ''
    };
  }
  
  // Sort by date
  messages.sort((a, b) => a.getDate().getTime() - b.getDate().getTime());
  
  const startDate = messages[0].getDate();
  const lastDate = messages[messages.length - 1].getDate();
  const lastSender = extractEmailAddress(messages[messages.length - 1].getFrom());
  const firstSender = extractEmailAddress(messages[0].getFrom());
  
  return {
    messageCount,
    startDate,
    lastDate,
    lastSender,
    firstSender
  };
}

/**
 * Computes "Waiting On" - who we're waiting for a reply from
 * @param {string} direction - Direction of current message
 * @param {string} lastSender - Who sent the last message in thread
 * @param {string} primaryContact - Primary contact email
 * @param {string[]} internalEmails - List of internal email addresses
 * @returns {string} - 'Us', 'Them', or 'Unknown'
 */
function computeWaitingOn(direction, lastSender, primaryContact, internalEmails) {
  const internalLower = internalEmails.map(e => e.toLowerCase());
  const lastSenderLower = lastSender.toLowerCase();
  
  if (internalLower.includes(lastSenderLower)) {
    // Last message was from us
    return 'Them';
  } else {
    // Last message was from them
    return 'Us';
  }
}

/**
 * Gets or assigns a Client ID for an email address
 * @param {string} email - Email address to look up
 * @param {Spreadsheet} ss - Spreadsheet object
 * @returns {string} - Client ID (existing or newly assigned)
 */
function getOrAssignClientId(email, ss) {
  // Client ID assignment is handled by client-reminder app
  // For now, return empty string - this will be populated by the sync
  return '';
}

/**
 * Gets internal email addresses from settings
 * @param {Spreadsheet} ss - Spreadsheet object
 * @returns {string[]} - Array of internal email addresses
 */
function getInternalEmails(ss) {
  const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!settingsSheet) return [];
  
  const internalEmailsStr = settingsSheet.getRange('B8').getValue();
  if (!internalEmailsStr) return [];
  
  return internalEmailsStr.toString()
    .split(',')
    .map(e => e.trim().toLowerCase())
    .filter(e => e.length > 0);
}

// ============================================================================
// MAIN LOGGER FUNCTION
// ============================================================================

/**
 * Main logger function - searches Gmail and logs new messages with enriched metadata
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
  const internalEmails = getInternalEmails(ss);
  
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
        // Check if thread already processed
        const threadLabels = thread.getLabels();
        const hasProcessedLabel = threadLabels.some(label => label.getName() === LABEL_NAME);
        if (hasProcessedLabel) {
          return;
        }
        
        // Build thread summary
        const threadSummary = buildThreadSummary(thread);
        
        const messages = thread.getMessages();
        messages.forEach(message => {
          // Log the message with enriched metadata
          const logRow = createEnrichedLogRow(message, rule, categories, logBody, bodyMaxChars, internalEmails, threadSummary, ss);
          logSheet.appendRow(logRow);
          totalLogged++;
        });
        
        // Mark entire thread as processed
        thread.addLabel(processedLabel);
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
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === 'Yes' && row[2]) {
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
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const categories = [];
  
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
 * Creates an enriched log row from a Gmail message
 */
function createEnrichedLogRow(message, rule, categories, logBody, bodyMaxChars, internalEmails, threadSummary, ss) {
  const subject = message.getSubject();
  const fromRaw = message.getFrom();
  const toRaw = message.getTo();
  const ccRaw = message.getCc() || '';
  const snippet = message.getPlainBody().substring(0, logBody ? bodyMaxChars : 100);
  const category = categorizeEmail(subject, snippet, categories);
  
  // Extract email addresses
  const fromEmail = extractEmailAddress(fromRaw);
  const toEmails = extractAllEmailAddresses(toRaw);
  const allParticipants = getParticipants(fromRaw, toEmails, ccRaw);
  
  // Determine direction
  const direction = determineDirection(fromEmail, toEmails, internalEmails);
  
  // Extract primary contact
  const primaryContact = extractPrimaryContact(direction, fromEmail, toEmails, internalEmails);
  
  // Extract domain
  const domain = extractDomain(primaryContact);
  
  // Is internal?
  const isInternal = isInternalEmail(fromEmail, toEmails, internalEmails) ? 'Yes' : 'No';
  
  // Waiting on
  const waitingOn = computeWaitingOn(direction, threadSummary.lastSender, primaryContact, internalEmails);
  
  // Build Gmail link
  const messageId = message.getId();
  const threadId = message.getThread().getId();
  const gmailLink = `https://mail.google.com/mail/u/0/#inbox/${threadId}`;
  
  // Get labels
  const labels = message.getThread().getLabels().map(l => l.getName()).join(', ');
  
  // Client ID (will be populated by client-reminder sync)
  const clientId = '';
  
  return [
    new Date().toISOString(),        // A - Logged At
    rule.name,                         // B - Rule Name
    fromRaw,                           // C - From (raw)
    toRaw,                             // D - To (raw)
    subject,                           // E - Subject
    message.getDate().toISOString(),   // F - Date Sent
    category,                          // G - Category
    snippet,                           // H - Snippet
    labels,                            // I - Labels
    messageId,                         // J - Message ID
    threadId,                          // K - Thread ID
    gmailLink,                         // L - Gmail Link
    // Enriched metadata
    primaryContact,                    // M - Primary Contact Email
    direction,                         // N - Direction
    domain,                            // O - Domain
    allParticipants.join(', '),        // P - Participants
    allParticipants.length,            // Q - Participants Count
    isInternal,                        // R - Is Internal
    threadSummary.messageCount,        // S - Thread Message Count
    threadSummary.startDate ? threadSummary.startDate.toISOString() : '',  // T - Thread Start Date
    threadSummary.lastDate ? threadSummary.lastDate.toISOString() : '',    // U - Last Message In Thread
    threadSummary.lastSender,          // V - Last Sender
    waitingOn,                         // W - Waiting On
    clientId                           // X - Client ID
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