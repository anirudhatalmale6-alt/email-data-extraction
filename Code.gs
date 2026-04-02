// ============================================================
// Email Data Extraction to Google Sheet
// Automated system for tip.top.towing.reviews@gmail.com
// ============================================================

// ---- CONFIGURATION ----
const CONFIG = {
  OPENAI_API_KEY: 'YOUR_OPENAI_API_KEY_HERE',  // Replace with your actual key
  SPREADSHEET_NAME: 'Customer Data - Master Sheet',
  DATA_SHEET_NAME: 'Customer Data',
  LOG_SHEET_NAME: 'Processing Log',
  PROCESSED_LABEL: 'Processed',
  OPENAI_MODEL: 'gpt-4o-mini',
  MAX_EMAILS_PER_RUN: 50
};

// ---- MAIN ENTRY POINT ----

/**
 * Main function - runs on a time-driven trigger (e.g., every hour).
 * Scans inbox for unprocessed emails, extracts customer data, appends to sheet.
 */
function processNewEmails() {
  const ss = getOrCreateSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.DATA_SHEET_NAME);
  const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);

  // Get or create the "Processed" label
  let label = GmailApp.getUserLabelByName(CONFIG.PROCESSED_LABEL);
  if (!label) {
    label = GmailApp.createLabel(CONFIG.PROCESSED_LABEL);
  }

  // Search for unprocessed emails (not labeled as Processed)
  const query = '-label:' + CONFIG.PROCESSED_LABEL + ' in:inbox';
  const threads = GmailApp.search(query, 0, CONFIG.MAX_EMAILS_PER_RUN);

  if (threads.length === 0) {
    logEntry(logSheet, 'No new emails to process.');
    return;
  }

  let totalExtracted = 0;
  let totalErrors = 0;

  for (const thread of threads) {
    const messages = thread.getMessages();

    for (const message of messages) {
      try {
        const extractedRecords = processMessage(message);

        if (extractedRecords && extractedRecords.length > 0) {
          for (const record of extractedRecords) {
            if (appendIfNotDuplicate(dataSheet, record)) {
              totalExtracted++;
            }
          }
        }

        logEntry(logSheet, 'Processed: ' + message.getSubject() + ' | Records: ' + (extractedRecords ? extractedRecords.length : 0));
      } catch (e) {
        totalErrors++;
        logEntry(logSheet, 'ERROR processing "' + message.getSubject() + '": ' + e.message);
      }
    }

    // Mark thread as processed
    thread.addLabel(label);
  }

  // Update dashboard
  updateDashboard(ss, totalExtracted, totalErrors, threads.length);

  logEntry(logSheet, '--- Run complete. Emails: ' + threads.length + ', Records extracted: ' + totalExtracted + ', Errors: ' + totalErrors + ' ---');
}

// ---- EMAIL PROCESSING ----

/**
 * Process a single email message - handles both inline content and attachments.
 */
function processMessage(message) {
  let allRecords = [];

  // 1. Process the email body (inline tables / text)
  const body = message.getBody(); // HTML body
  const plainBody = message.getPlainBody();
  const subject = message.getSubject();
  const emailDate = message.getDate();

  if (body && body.trim().length > 0) {
    const bodyRecords = extractFromText(plainBody || body, subject, emailDate);
    if (bodyRecords && bodyRecords.length > 0) {
      allRecords = allRecords.concat(bodyRecords);
    }
  }

  // 2. Process attachments (Excel files)
  const attachments = message.getAttachments();
  Logger.log('Found ' + attachments.length + ' attachment(s)');
  for (const attachment of attachments) {
    const name = attachment.getName().toLowerCase();
    Logger.log('Attachment: ' + attachment.getName() + ' | Type: ' + attachment.getContentType() + ' | Size: ' + attachment.getSize());
    if (name.endsWith('.xlsx') || name.endsWith('.xls') || name.endsWith('.csv')) {
      try {
        const attachmentRecords = extractFromAttachment(attachment, subject, emailDate);
        Logger.log('Extracted ' + (attachmentRecords ? attachmentRecords.length : 0) + ' records from attachment');
        if (attachmentRecords && attachmentRecords.length > 0) {
          allRecords = allRecords.concat(attachmentRecords);
        }
      } catch (e) {
        Logger.log('Error processing attachment ' + name + ': ' + e.message);
      }
    } else {
      Logger.log('Skipping attachment (not xlsx/xls/csv): ' + name);
    }
  }

  return allRecords;
}

/**
 * Extract customer data from email body text using OpenAI.
 */
function extractFromText(text, subject, emailDate) {
  // Clean up the text - remove excessive whitespace and HTML artifacts
  let cleanText = text.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();

  // Limit text length to avoid token limits
  if (cleanText.length > 8000) {
    cleanText = cleanText.substring(0, 8000);
  }

  if (cleanText.length < 20) return [];

  const prompt = `You are a data extraction assistant. Extract customer/contact records from the following email content.

This is a SWOOP review notification email. For each record, extract EXACTLY these fields:
- customer_name: The value of "Pickup Contact" (this is the customer's name)
- phone_number: The value of "Pickup Number" (this is the customer's phone number)
- date_of_business: Use the email date: ${emailDate.toISOString().split('T')[0]}

Email Subject: ${subject}
Email Date: ${emailDate.toISOString().split('T')[0]}

Email Content:
${cleanText}

Return a JSON array of objects. Each object must have exactly these keys: customer_name, phone_number, date_of_business.
- Format phone numbers consistently (include country code if present)
- Format dates as YYYY-MM-DD
- If a field is not found, use "N/A"
- Only return the JSON array, nothing else.
- If no customer data is found, return an empty array: []`;

  const response = callOpenAI(prompt);
  return parseOpenAIResponse(response, emailDate);
}

/**
 * Extract customer data from an Excel/CSV attachment using OpenAI.
 */
function extractFromAttachment(attachment, subject, emailDate) {
  const name = attachment.getName().toLowerCase();
  let textContent = '';

  if (name.endsWith('.csv')) {
    // CSV - read directly
    textContent = attachment.getDataAsString();
  } else {
    // XLSX/XLS - convert via Google Drive
    textContent = convertExcelToText(attachment);
  }

  if (!textContent || textContent.trim().length < 20) return [];

  // Limit text length
  if (textContent.length > 10000) {
    textContent = textContent.substring(0, 10000);
  }

  const prompt = `You are a data extraction assistant. Extract ALL customer/contact records from the following spreadsheet data.

For each row in the spreadsheet, extract EXACTLY these fields:
- customer_name: The value from the "Contact" column (this is the customer's name)
- phone_number: The value from the phone number column (look for columns containing phone numbers)
- date_of_business: The value from the "Survey Sent" column (this is the date of business). If not found, use: ${emailDate.toISOString().split('T')[0]}

Source file: ${attachment.getName()}

Spreadsheet Data:
${textContent}

Return a JSON array of objects. Each object must have exactly these keys: customer_name, phone_number, date_of_business.
- Format phone numbers consistently (include country code if present)
- Format dates as YYYY-MM-DD
- If a field is not found, use "N/A"
- Only return the JSON array, nothing else.
- If no customer data is found, return an empty array: []`;

  const response = callOpenAI(prompt);
  return parseOpenAIResponse(response, emailDate);
}

/**
 * Convert an Excel attachment to text by uploading to Google Drive and reading as a Sheet.
 */
function convertExcelToText(attachment) {
  let tempFileId = null;

  try {
    // Method 1: Use DriveApp to create the file, then convert
    const blob = attachment.copyBlob();
    blob.setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

    // Upload as Excel file first
    const tempFile = DriveApp.createFile(blob);
    tempFileId = tempFile.getId();
    Logger.log('Uploaded temp Excel file to Drive: ' + tempFileId);

    // Convert to Google Sheets using Drive API v2
    const convertedFile = Drive.Files.copy(
      { title: 'temp_converted_' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS },
      tempFileId
    );

    // Delete the original Excel temp file
    DriveApp.getFileById(tempFileId).setTrashed(true);
    tempFileId = convertedFile.id;

    Logger.log('Converted to Google Sheet: ' + convertedFile.id);

    // Open as spreadsheet and read all data
    const spreadsheet = SpreadsheetApp.openById(convertedFile.id);
    const sheets = spreadsheet.getSheets();
    let allText = '';

    for (const sheet of sheets) {
      const data = sheet.getDataRange().getValues();
      Logger.log('Sheet "' + sheet.getName() + '" has ' + data.length + ' rows');
      for (const row of data) {
        allText += row.join('\t') + '\n';
      }
      allText += '\n';
    }

    Logger.log('Converted Excel to text: ' + allText.length + ' chars');
    return allText;
  } catch (e) {
    Logger.log('Excel conversion error (method 1): ' + e.message);

    // Method 2: Try direct insert with convert flag
    try {
      const blob2 = attachment.copyBlob();
      const file = Drive.Files.insert(
        { title: 'temp_direct_' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS },
        blob2,
        { convert: true }
      );

      if (tempFileId) {
        try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch(x) {}
      }
      tempFileId = file.id;

      const spreadsheet = SpreadsheetApp.openById(file.id);
      const sheets = spreadsheet.getSheets();
      let allText = '';
      for (const sheet of sheets) {
        const data = sheet.getDataRange().getValues();
        for (const row of data) {
          allText += row.join('\t') + '\n';
        }
      }
      Logger.log('Method 2 succeeded: ' + allText.length + ' chars');
      return allText;
    } catch (e2) {
      Logger.log('Excel conversion error (method 2): ' + e2.message);
      return '';
    }
  } finally {
    // Clean up temp file
    if (tempFileId) {
      try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch(e) {}
    }
  }
}

// ---- OPENAI INTEGRATION ----

/**
 * Call OpenAI API to extract structured data.
 */
function callOpenAI(prompt) {
  const url = 'https://api.openai.com/v1/chat/completions';

  const payload = {
    model: CONFIG.OPENAI_MODEL,
    messages: [
      {
        role: 'system',
        content: 'You are a precise data extraction assistant. Always respond with valid JSON arrays only. No markdown, no explanations.'
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    temperature: 0.1,
    max_tokens: 4000
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.OPENAI_API_KEY
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    throw new Error('OpenAI API error (' + responseCode + '): ' + response.getContentText());
  }

  const json = JSON.parse(response.getContentText());
  return json.choices[0].message.content;
}

/**
 * Parse OpenAI response into structured records.
 */
function parseOpenAIResponse(responseText, fallbackDate) {
  try {
    // Clean up response - remove markdown code blocks if present
    let cleaned = responseText.trim();
    if (cleaned.startsWith('```')) {
      cleaned = cleaned.replace(/```json?\n?/g, '').replace(/```/g, '').trim();
    }

    const records = JSON.parse(cleaned);

    if (!Array.isArray(records)) return [];

    // Validate and clean each record
    return records.map(r => ({
      customer_name: (r.customer_name || 'N/A').toString().trim(),
      phone_number: (r.phone_number || 'N/A').toString().trim(),
      date_of_business: (r.date_of_business || fallbackDate.toISOString().split('T')[0]).toString().trim()
    })).filter(r => r.customer_name !== 'N/A' || r.phone_number !== 'N/A');

  } catch (e) {
    Logger.log('Failed to parse OpenAI response: ' + e.message + '\nResponse: ' + responseText);
    return [];
  }
}

// ---- SPREADSHEET OPERATIONS ----

/**
 * Get or create the master spreadsheet with proper headers.
 */
function getOrCreateSpreadsheet() {
  // Search for existing spreadsheet
  const files = DriveApp.getFilesByName(CONFIG.SPREADSHEET_NAME);

  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }

  // Create new spreadsheet
  const ss = SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);

  // Setup Data sheet
  const dataSheet = ss.getActiveSheet();
  dataSheet.setName(CONFIG.DATA_SHEET_NAME);

  // Headers
  const headers = ['Customer Name', 'Phone Number', 'Date of Business', 'Source Email Subject', 'Date Processed'];
  dataSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  dataSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF');
  dataSheet.setFrozenRows(1);

  // Set column widths
  dataSheet.setColumnWidth(1, 200); // Name
  dataSheet.setColumnWidth(2, 160); // Phone
  dataSheet.setColumnWidth(3, 140); // Date of Business
  dataSheet.setColumnWidth(4, 250); // Source
  dataSheet.setColumnWidth(5, 140); // Date Processed

  // Setup Dashboard sheet
  const dashSheet = ss.insertSheet('Dashboard');
  dashSheet.getRange('A1').setValue('Email Data Extraction - Dashboard').setFontSize(16).setFontWeight('bold');
  dashSheet.getRange('A3').setValue('Last Run:');
  dashSheet.getRange('A4').setValue('Total Records:');
  dashSheet.getRange('A5').setValue('Last Run - Records Added:');
  dashSheet.getRange('A6').setValue('Last Run - Errors:');
  dashSheet.getRange('A7').setValue('Last Run - Emails Processed:');
  dashSheet.getRange('B3:B7').setFontWeight('bold');
  dashSheet.setColumnWidth(1, 220);
  dashSheet.setColumnWidth(2, 200);

  // Setup Log sheet
  const logSheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
  logSheet.getRange(1, 1, 1, 2).setValues([['Timestamp', 'Message']]);
  logSheet.getRange(1, 1, 1, 2)
    .setFontWeight('bold')
    .setBackground('#34A853')
    .setFontColor('#FFFFFF');
  logSheet.setFrozenRows(1);
  logSheet.setColumnWidth(1, 180);
  logSheet.setColumnWidth(2, 600);

  // Move Dashboard to first position
  ss.setActiveSheet(dashSheet);
  ss.moveActiveSheet(1);

  return ss;
}

/**
 * Append a record to the data sheet if it's not a duplicate.
 * Deduplication based on customer_name + phone_number.
 */
function appendIfNotDuplicate(dataSheet, record) {
  const lastRow = dataSheet.getLastRow();

  if (lastRow > 1) {
    // Get existing data for dedup check
    const existingData = dataSheet.getRange(2, 1, lastRow - 1, 2).getValues();

    const nameNormalized = record.customer_name.toLowerCase().trim();
    const phoneNormalized = normalizePhone(record.phone_number);

    for (const row of existingData) {
      const existingName = row[0].toString().toLowerCase().trim();
      const existingPhone = normalizePhone(row[1].toString());

      // Match on name + phone (both must match to be a duplicate)
      if (existingName === nameNormalized && existingPhone === phoneNormalized) {
        return false; // Duplicate found
      }
    }
  }

  // Append new record
  const newRow = [
    record.customer_name,
    record.phone_number,
    record.date_of_business,
    record.source || '',
    new Date().toISOString().split('T')[0]
  ];

  dataSheet.appendRow(newRow);
  return true;
}

/**
 * Normalize phone number for comparison (strip non-digits).
 */
function normalizePhone(phone) {
  return phone.replace(/[^0-9]/g, '');
}

/**
 * Update the dashboard with run statistics.
 */
function updateDashboard(ss, recordsAdded, errors, emailsProcessed) {
  const dashSheet = ss.getSheetByName('Dashboard');
  if (!dashSheet) return;

  const dataSheet = ss.getSheetByName(CONFIG.DATA_SHEET_NAME);
  const totalRecords = Math.max(0, dataSheet.getLastRow() - 1);

  dashSheet.getRange('B3').setValue(new Date().toLocaleString());
  dashSheet.getRange('B4').setValue(totalRecords);
  dashSheet.getRange('B5').setValue(recordsAdded);
  dashSheet.getRange('B6').setValue(errors);
  dashSheet.getRange('B7').setValue(emailsProcessed);
}

/**
 * Add an entry to the processing log.
 */
function logEntry(logSheet, message) {
  logSheet.appendRow([new Date().toLocaleString(), message]);
}

// ---- TRIGGER MANAGEMENT ----

/**
 * Setup: Run this once to create the hourly trigger.
 * Go to Apps Script Editor > Run > setupTrigger
 */
function setupTrigger() {
  // Remove any existing triggers first
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processNewEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // Create hourly trigger
  ScriptApp.newTrigger('processNewEmails')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('Hourly trigger created successfully!');
}

/**
 * Setup: Run this once to create a daily trigger instead (if preferred).
 */
function setupDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processNewEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  ScriptApp.newTrigger('processNewEmails')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('Daily trigger created (8 AM) successfully!');
}

/**
 * Utility: Run this manually to process emails on demand.
 */
function runManually() {
  processNewEmails();
  Logger.log('Manual run complete. Check your spreadsheet!');
}

/**
 * Utility: Run this to test the OpenAI connection.
 */
function testOpenAI() {
  const testPrompt = 'Return this exact JSON: [{"customer_name":"Test User","phone_number":"555-1234","date_of_business":"2026-01-01"}]';
  const result = callOpenAI(testPrompt);
  Logger.log('OpenAI Response: ' + result);

  const parsed = parseOpenAIResponse(result, new Date());
  Logger.log('Parsed: ' + JSON.stringify(parsed));

  if (parsed.length > 0) {
    Logger.log('SUCCESS - OpenAI connection is working!');
  } else {
    Logger.log('WARNING - Could not parse response. Check API key.');
  }
}
